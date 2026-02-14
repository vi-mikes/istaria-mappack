/*
MapPackSyncTool (manifest-driven sync)
High-level behavior
- User selects an Istaria installation folder (must contain istaria.exe).
- Tool syncs remote resources_override\mappack\resources content into: <selected>\resources_override\mappack\resources
- Remote source is described by a JSON manifest (mappack_manifest.json) containing SHA-256 hashes.
- For each file:
	- If local exists and SHA-256 matches manifest: skip.
	- Otherwise download to a temp file, hash while downloading, verify SHA-256,
	  then replace destination.
Safety invariants
1) Manifest MUST be downloaded and parsed successfully before any delete occurs.
2) Downloads are verified against manifest SHA-256 before replacing local files.
Notes
- UI remains responsive: sync runs on a worker thread; UI updates use PostMessage.
- This file is intentionally kept as a single translation unit for easy building.
*/

//////////////////////////////////////////////////////
// Comment this out to suppress some local testing messages and/or presets
// #define DEBUG_MESSAGE
//////////////////////////////////////////////////////

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
#include <winver.h>
#pragma comment(lib, "Version.lib")
#include <wincrypt.h>
#include <softpub.h>
#include <wintrust.h>
#include <tlhelp32.h>
#include <cwchar>
#include <cwctype>
#include <bcrypt.h>
#pragma comment(lib, "bcrypt.lib")
#include <shlobj.h>
#include <shellapi.h>
#include <richedit.h>
#include <commctrl.h>
#include <commdlg.h>
#include "resource.h"   // IDI_WINDOWSPROJECT1, IDI_SMALL
#include <algorithm>
#include <filesystem>
#include <fstream>
#include <ctime>
#include <process.h>
#include <string>
#include <vector>
#include <unordered_set>
#include <atomic>
#include <cstring>     // memcpy
#include <cstdint>
#include <functional>  // std::function
#include <memory>      // std::unique_ptr
#include <winhttp.h>

// ------------------------------------------------------------
// Version helper: read FileVersion from VERSIONINFO resource
// ------------------------------------------------------------
#include <vector>

static std::wstring GetExeFileVersionString()
{
	wchar_t path[MAX_PATH] = {};
	if (!GetModuleFileNameW(nullptr, path, MAX_PATH))
		return L"unknown";

	DWORD handle = 0;
	DWORD size = GetFileVersionInfoSizeW(path, &handle);
	if (size == 0)
		return L"unknown";

	std::vector<BYTE> buffer(size);
	if (!GetFileVersionInfoW(path, handle, size, buffer.data()))
		return L"unknown";

	void* verData = nullptr;
	UINT len = 0;

	if (VerQueryValueW(buffer.data(),
		L"\\StringFileInfo\\040904b0\\FileVersion",
		&verData, &len) && len > 0)
	{
		return std::wstring(static_cast<wchar_t*>(verData));
	}

	return L"unknown";
}
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "wintrust.lib")
#pragma comment(lib, "crypt32.lib")
#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Comdlg32.lib")
#pragma comment(lib, "Shell32.lib")
namespace fs = std::filesystem;

// --------------------------------------------------
// Status helper (standardized error propagation)
// --------------------------------------------------
struct Status {
	bool ok = true;
	DWORD win32 = 0;
	HRESULT hr = S_OK;      // optional
	std::wstring msg;       // human-readable

	static Status Ok() { return Status{}; }

	static Status FailMsg(std::wstring m) {
		Status s; s.ok = false; s.msg = std::move(m); return s;
	}
	static Status FailWin32(std::wstring m, DWORD e) {
		Status s; s.ok = false; s.win32 = e; s.msg = std::move(m); return s;
	}
	static Status FailLastErr(std::wstring m) {
		return FailWin32(std::move(m), GetLastError());
	}
	static Status FailHr(std::wstring m, HRESULT h) {
		Status s; s.ok = false; s.hr = h; s.msg = std::move(m); return s;
	}
};

static inline std::wstring StatusToWString(const Status& st) {
	if (st.ok) return L"";
	std::wstring m = st.msg.empty() ? L"Error" : st.msg;
	if (st.win32 != 0) {
		m += L" (win32=" + std::to_wstring(st.win32) + L")";
	}
	else if (FAILED(st.hr)) {
		m += L" (hr=" + std::to_wstring((long long)st.hr) + L")";
	}
	return m;
}

static inline bool SetErrFromStatus(std::wstring* outErrW, const Status& st) {
	if (!outErrW) return st.ok;
	*outErrW = StatusToWString(st);
	return st.ok;
}


// --------------------------------------------------
// App identity (window title)
// --------------------------------------------------
#define MAP_PACK_SYNC_TOOL_NAME    L"MapPack Sync Tool"
static inline std::wstring GetDisplayVersion()
{
	return L"v" + std::wstring(GetExeFileVersionString().c_str());
}

static const std::wstring kWindowTitle =
std::wstring(MAP_PACK_SYNC_TOOL_NAME) + L" " + GetDisplayVersion();

namespace AppConstants
{
	static constexpr long kManifestConnectTimeoutSec = 15L;
	static constexpr long kManifestTimeoutSec = 120L;
	static constexpr long kFileConnectTimeoutMs = 15000L;
	static constexpr long kFileTimeoutMs = 0L;
	// We are defining this twice because WinHttp expects wide
	static constexpr const char* kUserAgent = "MapPackSyncTool by Cegaiel";
	static constexpr const wchar_t* kUserAgentW = L"MapPackSyncTool by Cegaiel";
}
struct WinHttpHandleDeleter { void operator()(HINTERNET h) const noexcept { if (h) WinHttpCloseHandle(h); } };
using WinHttpHandle = std::unique_ptr<void, WinHttpHandleDeleter>;

// --------------------------------------------------
// Global Remote config (manifest + root)
// --------------------------------------------------
static constexpr const char* kRemoteHost = "https://istaria-mappack.s3.us-west-2.amazonaws.com";
static constexpr const char* kRemoteRootPath = "/resources_override/";
static constexpr const char* kManifestPath = "/mappack_manifest.json";
static constexpr const char* kManifestOldPath = "/mappack_manifest_old.json";
static constexpr const wchar_t* kUpdateExeUrl = L"https://istaria-mappack.s3.us-west-2.amazonaws.com/MapPackSyncTool.exe";
static constexpr const wchar_t* kUpdateVersionUrl = L"https://istaria-mappack.s3.us-west-2.amazonaws.com/version.txt";

// --------------------------------------------------
// Global UI layout (initial size + margins)
// --------------------------------------------------
static const int MAIN_WINDOW_WIDTH = 825;  // initial window width
static const int MAIN_WINDOW_HEIGHT = 800;   // initial window height
// RichEdit margins (client area)
static const int OUTPUT_MARGIN_LEFT = 10;
static const int OUTPUT_MARGIN_TOP = 95;
static const int OUTPUT_MARGIN_RIGHT = 10;
static const int OUTPUT_MARGIN_BOTTOM = 10;
// Minimum client size (prevents resizing too small)
static const int MIN_CLIENT_W = 825;
static const int MIN_CLIENT_H = MAIN_WINDOW_HEIGHT;

// --------------------------------------------------
// Small RAII helpers (Win32 resources)
// --------------------------------------------------
struct unique_handle
{
	HANDLE h = nullptr;
	unique_handle() = default;
	explicit unique_handle(HANDLE handle) : h(handle) {}
	unique_handle(const unique_handle&) = delete;
	unique_handle& operator=(const unique_handle&) = delete;
	unique_handle(unique_handle&& o) noexcept : h(o.h) { o.h = nullptr; }
	unique_handle& operator=(unique_handle&& o) noexcept
	{
		if (this != &o) { reset(); h = o.h; o.h = nullptr; }
		return *this;
	}
	~unique_handle() { reset(); }
	void reset(HANDLE handle = nullptr)
	{
		if (h && h != INVALID_HANDLE_VALUE) CloseHandle(h);
		h = handle;
	}
	HANDLE get() const { return h; }
	HANDLE release() { HANDLE tmp = h; h = nullptr; return tmp; }
	explicit operator bool() const { return (h && h != INVALID_HANDLE_VALUE); }
};
struct unique_hfont
{
	HFONT h = nullptr;
	unique_hfont() = default;
	explicit unique_hfont(HFONT font) : h(font) {}
	unique_hfont(const unique_hfont&) = delete;
	unique_hfont& operator=(const unique_hfont&) = delete;
	unique_hfont(unique_hfont&& o) noexcept : h(o.h) { o.h = nullptr; }
	unique_hfont& operator=(unique_hfont&& o) noexcept
	{
		if (this != &o) { reset(); h = o.h; o.h = nullptr; }
		return *this;
	}
	~unique_hfont() { reset(); }
	void reset(HFONT font = nullptr)
	{
		if (h) DeleteObject(h);
		h = font;
	}
	HFONT get() const { return h; }
	HFONT release() { HFONT tmp = h; h = nullptr; return tmp; }
	explicit operator bool() const { return h != nullptr; }
};
struct unique_heap_wstr
{
	wchar_t* p = nullptr;
	unique_heap_wstr() = default;
	explicit unique_heap_wstr(wchar_t* ptr) : p(ptr) {}
	unique_heap_wstr(const unique_heap_wstr&) = delete;
	unique_heap_wstr& operator=(const unique_heap_wstr&) = delete;
	unique_heap_wstr(unique_heap_wstr&& o) noexcept : p(o.p) { o.p = nullptr; }
	unique_heap_wstr& operator=(unique_heap_wstr&& o) noexcept
	{
		if (this != &o) { reset(); p = o.p; o.p = nullptr; }
		return *this;
	}
	~unique_heap_wstr() { reset(); }
	void reset(wchar_t* ptr = nullptr)
	{
		if (p) HeapFree(GetProcessHeap(), 0, p);
		p = ptr;
	}
	wchar_t* get() const { return p; }
	wchar_t* release() { wchar_t* tmp = p; p = nullptr; return tmp; }
	explicit operator bool() const { return p != nullptr; }
};

// --------------------------------------------------
// Main window state (UI handles + worker-thread coordination)
// --------------------------------------------------
struct AppState
{
	HWND hMainWnd = nullptr;
	HWND hFolderLabel = nullptr;
	HWND hBrowseBtn = nullptr;
	HWND hRunButton = nullptr;
	HWND hCancelBtn = nullptr;
	HWND hDeleteBtn = nullptr;
	HWND hCopyLogBtn = nullptr;
	HWND hSaveLogBtn = nullptr;
	HWND hCheckUpdatesBtn = nullptr;
	HWND hHelpBtn = nullptr;
	HWND hFolderEdit = nullptr;
	HWND hOutput = nullptr;
	HWND hProgress = nullptr;
	HWND hProgressText = nullptr;
	HWND hTooltip = nullptr;
	HFONT hFontUI = nullptr;
	HFONT hFontMono = nullptr;
	HANDLE hWorkerThread = nullptr;
	std::atomic_bool isRunning{ false };
	std::atomic_bool cancelRequested{ false };
	bool pendingExitAfterWorker = false;
	// Track whether we currently have marquee enabled (so we can cleanly toggle style)
	bool progressMarqueeOn = false;
	// Track last known progress range/pos (for cancel freeze)
	int progressTotal = 100;
	int progressPos = 0;
	bool progressFrozenOnCancel = false;
	HANDLE hUpdateThread = nullptr;
	std::atomic_bool isUpdateRunning{ false };
	bool logActionsArmed = false;
};
static AppState* g_state = nullptr;
static HANDLE g_hSingleInstanceMutex = nullptr;
// Remember the last directory used by the Save Log dialog.
static std::wstring g_lastSaveDir;

// Forward declarations used before their definitions.
static fs::path GetThisExePath();
static void TrimInPlace(std::wstring& s);
static void StripSurroundingQuotes(std::wstring& s);
static std::wstring Utf8ToWide(const std::string& s);

// --------------------------------------------------
// Settings persistence (portable INI next to EXE)
// - Remembers the last selected Istaria base folder.
// --------------------------------------------------
static const wchar_t* kSettingsIniName = L"MapPackSyncTool.ini";
static const wchar_t* kIniSectionSettings = L"Settings";
static const wchar_t* kIniKeyLastFolder = L"LastFolder";

static std::wstring GetSettingsIniPath()
{
	// Portable: store settings INI next to the EXE.
	wchar_t exePathW[MAX_PATH]{};
	DWORD n = GetModuleFileNameW(nullptr, exePathW, (DWORD)_countof(exePathW));
	if (n == 0 || n >= _countof(exePathW))
		return std::wstring(kSettingsIniName); // fallback: current working dir

	fs::path exePath(exePathW);
	fs::path iniPath = exePath.parent_path() / kSettingsIniName;
	return iniPath.wstring();
}


static std::wstring IniReadLastFolder()
{
	wchar_t buf[4096]{};
	const std::wstring iniPath = GetSettingsIniPath();
	// If file/key doesn't exist, this returns empty string.
	GetPrivateProfileStringW(kIniSectionSettings, kIniKeyLastFolder, L"", buf, (DWORD)_countof(buf), iniPath.c_str());
	std::wstring out(buf);
	TrimInPlace(out);
	StripSurroundingQuotes(out);
	return out;
}

static void IniWriteLastFolder(const std::wstring& folder)
{
	std::wstring v = folder;
	TrimInPlace(v);
	StripSurroundingQuotes(v);
	if (v.empty())
		return;
	const std::wstring iniPath = GetSettingsIniPath();
	// This creates the INI if it doesn't exist yet.
	WritePrivateProfileStringW(kIniSectionSettings, kIniKeyLastFolder, v.c_str(), iniPath.c_str());
}


struct CancelToken
{
	std::atomic_bool* flag = nullptr;
	bool IsCanceled() const noexcept
	{
		return flag && flag->load(std::memory_order_relaxed);
	}
};
static void ComputeWindowSizeFromClientStyle(DWORD style, DWORD ex, int clientW, int clientH, int& outWinW, int& outWinH)
{
	RECT r{ 0, 0, clientW, clientH };
	AdjustWindowRectEx(&r, style, FALSE, ex);
	outWinW = (r.right - r.left);
	outWinH = (r.bottom - r.top);
}
static void LayoutMainWindow(HWND hwnd, AppState* st)
{
	if (!st) return;
	RECT rc{};
	GetClientRect(hwnd, &rc);
	const int cw = (rc.right - rc.left);
	const int ch = (rc.bottom - rc.top);
	const int m = 10;
	const int rowY = 12;
	const int ctrlH = 22;
	const int gap = 8;
	const int labelW = 170;
	// Buttons anchored to the right
	const int btnW = 92;
	const int btnH = ctrlH;
	int right = cw - m;
	int deleteX = right - btnW;
	int cancelX = deleteX - gap - btnW;
	int runX = cancelX - gap - btnW;
	int browseX = runX - gap - btnW;
	int labelX = m;
	int labelY = 15;
	int editX = labelX + labelW + gap;
	int editY = rowY;
	int editW = browseX - gap - editX;
	if (editW < 50) editW = 50;
	// Progress bar + status text
	int progY = rowY + ctrlH + 10;
	int progW = cw - (m * 2);
	if (progW < 10) progW = 10;

	// Check-for-updates + Copy/Save buttons live on the progress-text row (right side)
	const int updateW = 130;
	const int helpW = 22;
	int progTextY = progY + 18;
	int helpX = right - helpW;
	int saveLogX = helpX - gap - btnW;
	int copyLogX = saveLogX - gap - btnW;
	int updateX = copyLogX - gap - updateW;
	int progTextW = updateX - gap - m;
	if (progTextW < 50) progTextW = 50;

	// Output box fills remaining space
	int outY = progY + 44;
	int outW = cw - (m * 2);
	int outH = ch - outY - m;
	if (outW < 10) outW = 10;
	if (outH < 10) outH = 10;
	// Batch repositioning to reduce redraw/flicker
	HDWP hdwp = BeginDeferWindowPos(13);
	if (!hdwp) {
		if (st->hFolderLabel) MoveWindow(st->hFolderLabel, labelX, labelY, labelW, 20, TRUE);
		if (st->hFolderEdit)  MoveWindow(st->hFolderEdit, editX, editY, editW, ctrlH, TRUE);
		if (st->hBrowseBtn)   MoveWindow(st->hBrowseBtn, browseX, rowY, btnW, btnH, TRUE);
		if (st->hRunButton)   MoveWindow(st->hRunButton, runX, rowY, btnW, btnH, TRUE);
		if (st->hCancelBtn)   MoveWindow(st->hCancelBtn, cancelX, rowY, btnW, btnH, TRUE);
		if (st->hDeleteBtn)   MoveWindow(st->hDeleteBtn, deleteX, rowY, btnW, btnH, TRUE);
		if (st->hProgress)     MoveWindow(st->hProgress, m, progY, progW, 14, TRUE);
		if (st->hProgressText) MoveWindow(st->hProgressText, m, progTextY, progTextW, 22, TRUE);
		if (st->hCheckUpdatesBtn) MoveWindow(st->hCheckUpdatesBtn, updateX, progTextY, updateW, btnH, TRUE);
		if (st->hCopyLogBtn)   MoveWindow(st->hCopyLogBtn, copyLogX, progTextY, btnW, btnH, TRUE);
		if (st->hSaveLogBtn)   MoveWindow(st->hSaveLogBtn, saveLogX, progTextY, btnW, btnH, TRUE);
		if (st->hHelpBtn)         MoveWindow(st->hHelpBtn, helpX, progTextY, helpW, btnH, TRUE);
		if (st->hOutput)       MoveWindow(st->hOutput, m, outY, outW, outH, TRUE);
		return;
	}
	auto defer = [&](HWND h, int x, int y, int w, int hgt) {
		if (!h) return;
		hdwp = DeferWindowPos(hdwp, h, nullptr, x, y, w, hgt, SWP_NOZORDER | SWP_NOACTIVATE);
	};
	defer(st->hFolderLabel, labelX, labelY, labelW, 20);
	defer(st->hFolderEdit, editX, editY, editW, ctrlH);
	defer(st->hBrowseBtn, browseX, rowY, btnW, btnH);
	defer(st->hRunButton, runX, rowY, btnW, btnH);
	defer(st->hCancelBtn, cancelX, rowY, btnW, btnH);
	defer(st->hDeleteBtn, deleteX, rowY, btnW, btnH);
	defer(st->hProgress, m, progY, progW, 14);
	defer(st->hProgressText, m, progTextY, progTextW, 22);
	defer(st->hCheckUpdatesBtn, updateX, progTextY, updateW, btnH);
	defer(st->hCopyLogBtn, copyLogX, progTextY, btnW, btnH);
	defer(st->hSaveLogBtn, saveLogX, progTextY, btnW, btnH);
	defer(st->hHelpBtn, helpX, progTextY, helpW, btnH);
	defer(st->hOutput, m, outY, outW, outH);
	EndDeferWindowPos(hdwp);
}
// Custom message for log appends (main thread only)
static constexpr UINT WM_APP_LOG = WM_APP + 1;
// Progress bar messages (main thread only)
static constexpr UINT WM_APP_PROGRESS_MARQ_ON = WM_APP + 10;
static constexpr UINT WM_APP_PROGRESS_MARQ_OFF = WM_APP + 11;
static constexpr UINT WM_APP_PROGRESS_INIT = WM_APP + 12; // wParam = total files
static constexpr UINT WM_APP_PROGRESS_SET = WM_APP + 13;  // wParam = current (0..total)
static constexpr UINT WM_APP_PROGRESS_TEXT = WM_APP + 14; // lParam = wchar_t* (HeapAlloc), UI frees
// Worker completion message (main thread only)
static constexpr UINT WM_APP_WORKER_DONE = WM_APP + 20;


// Unified UI event message (main thread only): lParam = UiEvent* (new), UI frees
static constexpr UINT WM_APP_UI_EVENT = WM_APP + 2;

enum class UiEventKind : int
{
	LogAppendW = 1,
	ProgressMarqueeOn,
	ProgressMarqueeOff,
	ProgressInit,       // u1 = total
	ProgressSet,        // u1 = pos
	ProgressTextW,      // text = progress label
	WorkerDone,
	UpdateResultPtr     // ptr = UpdateResult*
};

struct UiEvent
{
	UiEventKind kind{};
	size_t u1 = 0;
	size_t u2 = 0;
	void* ptr = nullptr;
	std::wstring text;
};

static bool PostUiEvent(UiEvent* ev)
{
	if (!ev) return false;
	if (!g_state || !g_state->hMainWnd)
	{
		delete ev;
		return false;
	}
	if (PostMessageW(g_state->hMainWnd, WM_APP_UI_EVENT, 0, (LPARAM)ev))
		return true;
	delete ev;
	return false;
}
// --------------------------------------------------
// PROGRESS BAR (FIXED: dynamic marquee style toggle)
// --------------------------------------------------
static void SetProgressMarquee(AppState* st, bool on)
{
	if (!st || !st->hProgress) return;
	LONG_PTR style = GetWindowLongPtrW(st->hProgress, GWL_STYLE);
	if (on)
	{
		if (!st->progressMarqueeOn)
		{
			// Add PBS_MARQUEE style only when needed
			style |= PBS_MARQUEE;
			SetWindowLongPtrW(st->hProgress, GWL_STYLE, style);
			SetWindowPos(st->hProgress, nullptr, 0, 0, 0, 0,
				SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
			// Turn marquee on
			SendMessageW(st->hProgress, PBM_SETMARQUEE, TRUE, 30);
			st->progressMarqueeOn = true;
		}
	}
	else
	{
		if (st->progressMarqueeOn)
		{
			// Turn marquee off first
			SendMessageW(st->hProgress, PBM_SETMARQUEE, FALSE, 0);
			// Remove PBS_MARQUEE so it draws as a normal smooth bar (prevents  segmented bars  look)
			style &= ~PBS_MARQUEE;
			SetWindowLongPtrW(st->hProgress, GWL_STYLE, style);
			SetWindowPos(st->hProgress, nullptr, 0, 0, 0, 0,
				SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED);
			// Reset to a clean deterministic state
			SendMessageW(st->hProgress, PBM_SETRANGE32, 0, 1);
			SendMessageW(st->hProgress, PBM_SETPOS, 0, 0);
			st->progressMarqueeOn = false;
		}
	}
}
static void FreezeProgressOnCancel(AppState* st)
{
	if (!st || !st->hProgress) return;
	// Once frozen, ignore any later progress/marquee messages.
	st->progressFrozenOnCancel = true;
	// Ensure marquee is off and show as fully filled.
	SetProgressMarquee(st, false);
	int total = st->progressTotal;
	if (total <= 0) total = 100;
	SendMessageW(st->hProgress, PBM_SETRANGE32, 0, total);
	SendMessageW(st->hProgress, PBM_SETPOS, total, 0);
}

// --------------------------------------------------
// UI output helpers (main thread)
// --------------------------------------------------
static std::wstring GetOutputTextW()
{
	if (!g_state || !g_state->hOutput) return L"";
	const int len = GetWindowTextLengthW(g_state->hOutput);
	if (len <= 0) return L"";

	std::wstring buf;
	buf.resize((size_t)len);
	GetWindowTextW(g_state->hOutput, buf.data(), len + 1);
	return buf;
}

static void UpdateLogActionButtonsEnabled()
{
	if (!g_state) return;
	if (g_state->isRunning.load())
	{
		if (g_state->hCopyLogBtn) EnableWindow(g_state->hCopyLogBtn, FALSE);
		if (g_state->hSaveLogBtn) EnableWindow(g_state->hSaveLogBtn, FALSE);
		return;
	}
	if (g_state->hCopyLogBtn) EnableWindow(g_state->hCopyLogBtn, TRUE);
	if (g_state->hSaveLogBtn) EnableWindow(g_state->hSaveLogBtn, g_state->logActionsArmed ? TRUE : FALSE);
}

static void UpdateCheckUpdatesButtonEnabled()
{
	if (!g_state || !g_state->hCheckUpdatesBtn) return;
	const bool enable = !g_state->isRunning.load() && !g_state->isUpdateRunning.load();
	EnableWindow(g_state->hCheckUpdatesBtn, enable ? TRUE : FALSE);
}

static void UpdateHelpButtonEnabled()
{
	if (!g_state || !g_state->hHelpBtn) return;
	const bool enable = g_state->logActionsArmed && !g_state->isRunning.load();
	EnableWindow(g_state->hHelpBtn, enable ? TRUE : FALSE);
}


// --------------------------------------------------
// Unified UI toggle when a worker (sync/remove) is running.
// Requirement: disable every button except Cancel Sync while running.
// --------------------------------------------------
static void SetUiForWorkerRunning(AppState* st, bool running)
{
	if (!st) return;

	// Disable everything except Cancel while running.
	if (st->hBrowseBtn) EnableWindow(st->hBrowseBtn, running ? FALSE : TRUE);
	if (st->hRunButton) EnableWindow(st->hRunButton, running ? FALSE : TRUE);
	if (st->hDeleteBtn) EnableWindow(st->hDeleteBtn, running ? FALSE : TRUE);
	if (st->hFolderEdit) EnableWindow(st->hFolderEdit, running ? FALSE : TRUE);
	if (st->hHelpBtn && running) EnableWindow(st->hHelpBtn, FALSE);

	// These buttons have additional logic (log empty / update thread), but they must be disabled while running.
	// We hard-disable them here when running; when idle we defer to the existing helper logic.
	if (running)
	{
		if (st->hCopyLogBtn) EnableWindow(st->hCopyLogBtn, FALSE);
		if (st->hSaveLogBtn) EnableWindow(st->hSaveLogBtn, FALSE);
		if (st->hCheckUpdatesBtn) EnableWindow(st->hCheckUpdatesBtn, FALSE);
	}
	else
	{
		UpdateLogActionButtonsEnabled();
		UpdateCheckUpdatesButtonEnabled();
		UpdateHelpButtonEnabled();
	}

	// Cancel is the inverse (only enabled while running).
	if (st->hCancelBtn) EnableWindow(st->hCancelBtn, running ? TRUE : FALSE);
}

static void AppendToOutputW(const wchar_t* text)
{
	if (!g_state || !g_state->hOutput || !text) return;
	const LRESULT len = SendMessageW(g_state->hOutput, WM_GETTEXTLENGTH, 0, 0);
	SendMessageW(g_state->hOutput, EM_SETSEL, (WPARAM)len, (LPARAM)len);
	SendMessageW(g_state->hOutput, EM_REPLACESEL, FALSE, (LPARAM)text);
	SendMessageW(g_state->hOutput, EM_SCROLLCARET, 0, 0);
	UpdateLogActionButtonsEnabled();
}

static void ClearOutput()
{
	if (g_state && g_state->hOutput)
		SetWindowTextW(g_state->hOutput, L"");
	UpdateLogActionButtonsEnabled();
}

// Loads MapPackSyncTool.txt (next to the EXE) into the RichEdit output box.
// - If clearFirst is true, the output is cleared before loading.
// - If showErrorBox is true, errors are shown in a MessageBox; otherwise failures are silent.
static bool LoadHelpTextIntoOutput(bool clearFirst, bool showErrorBox)
{
	if (!g_state || !g_state->hOutput)
		return false;

	if (clearFirst)
		ClearOutput();

	fs::path exePath = GetThisExePath();
	fs::path txtPath = exePath.empty() ? fs::path(L"MapPackSyncTool.txt") : (exePath.parent_path() / L"MapPackSyncTool.txt");

	std::ifstream in(txtPath, std::ios::binary);
	if (!in)
	{
		if (showErrorBox)
		{
			std::wstring msg = L"Could not open:\r\n\r\n" + txtPath.wstring();
			MessageBoxW(g_state->hMainWnd, msg.c_str(), L"MapPack Sync Tool", MB_OK | MB_ICONERROR);
		}
		return false;
	}

	std::string bytes;
	in.seekg(0, std::ios::end);
	std::streamoff sz = in.tellg();
	if (sz < 0) sz = 0;
	bytes.resize((size_t)sz);
	in.seekg(0, std::ios::beg);
	if (!bytes.empty())
		in.read(bytes.data(), (std::streamsize)bytes.size());

	std::wstring textW;
	const unsigned char* b = (const unsigned char*)bytes.data();
	const size_t n = bytes.size();
	if (n >= 2 && b[0] == 0xFF && b[1] == 0xFE)
	{
		// UTF-16 LE BOM
		const size_t wcharCount = (n - 2) / 2;
		textW.assign((const wchar_t*)(b + 2), (const wchar_t*)(b + 2) + wcharCount);
	}
	else
	{
		size_t start = 0;
		if (n >= 3 && b[0] == 0xEF && b[1] == 0xBB && b[2] == 0xBF)
			start = 3; // UTF-8 BOM
		textW = Utf8ToWide(bytes.substr(start));
	}

	SetWindowTextW(g_state->hOutput, textW.c_str());
	// Scroll to top
	SendMessageW(g_state->hOutput, EM_SETSEL, 0, 0);
	SendMessageW(g_state->hOutput, EM_SCROLLCARET, 0, 0);
	SendMessageW(g_state->hOutput, WM_VSCROLL, SB_TOP, 0);
	UpdateLogActionButtonsEnabled();
	return true;
}

static bool CopyOutputToClipboard()
{
	if (!g_state || !g_state->hMainWnd) return false;
	std::wstring textW = GetOutputTextW();
	if (textW.empty())
	{
		MessageBoxW(g_state->hMainWnd, L"Log is empty. Nothing to Copy!", L"Copy Log", MB_OK | MB_ICONINFORMATION);
		return false;
	}
	if (!OpenClipboard(g_state->hMainWnd))
	{
		MessageBoxW(g_state->hMainWnd, L"Failed to open the clipboard.", L"Copy Log", MB_OK | MB_ICONERROR);
		return false;
	}
	EmptyClipboard();
	const size_t bytes = (textW.size() + 1) * sizeof(wchar_t);
	HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, bytes);
	if (!hMem)
	{
		CloseClipboard();
		MessageBoxW(g_state->hMainWnd, L"Failed to allocate clipboard memory.", L"Copy Log", MB_OK | MB_ICONERROR);
		return false;
	}
	void* pMem = GlobalLock(hMem);
	if (!pMem)
	{
		GlobalFree(hMem);
		CloseClipboard();
		MessageBoxW(g_state->hMainWnd, L"Failed to lock clipboard memory.", L"Copy Log", MB_OK | MB_ICONERROR);
		return false;
	}
	memcpy(pMem, textW.c_str(), bytes);
	GlobalUnlock(hMem);
	if (!SetClipboardData(CF_UNICODETEXT, hMem))
	{
		GlobalFree(hMem);
		CloseClipboard();
		MessageBoxW(g_state->hMainWnd, L"Failed to set clipboard data.", L"Copy Log", MB_OK | MB_ICONERROR);
		return false;
	}
	// clipboard owns hMem now
	CloseClipboard();
	MessageBoxW(g_state->hMainWnd, L"Log has been copied to clipboard", L"Copy Log", MB_OK | MB_ICONINFORMATION);
	return true;
}
static void SaveOutputToFile()
{
	if (!g_state || !g_state->hMainWnd) return;
	std::wstring textW = GetOutputTextW();
	if (textW.empty())
	{
		MessageBoxW(g_state->hMainWnd, L"Log is empty. Nothing to Save!", L"Save Log", MB_OK | MB_ICONINFORMATION);
		return;
	}

	wchar_t filePath[MAX_PATH]{};
	// Default filename includes Unix epoch time (seconds).
	const time_t epoch = time(nullptr);
	std::wstring defaultName = L"MapPackSyncTool_" + std::to_wstring((long long)epoch) + L"_Log.txt";
	wcsncpy_s(filePath, defaultName.c_str(), _TRUNCATE);

	OPENFILENAMEW ofn{};
	ofn.lStructSize = sizeof(ofn);
	ofn.hwndOwner = g_state->hMainWnd;
	ofn.lpstrFile = filePath;
	ofn.nMaxFile = (DWORD)_countof(filePath);
	ofn.lpstrFilter = L"Text Files (*.txt)\0*.txt\0All Files (*.*)\0*.*\0";
	ofn.nFilterIndex = 1;
	ofn.Flags = OFN_OVERWRITEPROMPT | OFN_PATHMUSTEXIST | OFN_NOCHANGEDIR;
	ofn.lpstrDefExt = L"txt";
	ofn.lpstrTitle = L"Save Log As";
	// Initial directory behavior:
	//  - If we have a remembered directory from a previous save, start there.
	//  - Otherwise, default to: <FolderEdit>\resources_override\mappack (if it exists).
	std::wstring initialDir;
	if (!g_lastSaveDir.empty())
	{
		ofn.lpstrInitialDir = g_lastSaveDir.c_str();
	}
	else if (g_state->hFolderEdit)
	{
		wchar_t base[MAX_PATH]{};
		GetWindowTextW(g_state->hFolderEdit, base, (int)_countof(base));
		if (base[0] != L'\0')
		{
			try
			{
				fs::path p(base);
				p /= L"resources_override";
				p /= L"mappack";
				if (fs::exists(p) && fs::is_directory(p))
				{
					initialDir = p.wstring();
					ofn.lpstrInitialDir = initialDir.c_str();
				}
			}
			catch (...) {}
		}
	}

	if (!GetSaveFileNameW(&ofn))
		return;

	// Remember the directory for next time.
	try
	{
		g_lastSaveDir = fs::path(ofn.lpstrFile).parent_path().wstring();
	}
	catch (...) {}

	std::ofstream f(ofn.lpstrFile, std::ios::binary);


	if (!f)
	{
		MessageBoxW(g_state->hMainWnd, L"Failed to open file for writing.", L"Save Log", MB_OK | MB_ICONERROR);
		return;
	}

	// Write UTF-16LE with BOM so Notepad opens it reliably.
	const unsigned char bom[2] = { 0xFF, 0xFE };
	f.write((const char*)bom, 2);
	f.write((const char*)textW.data(), (std::streamsize)textW.size() * (std::streamsize)sizeof(wchar_t));
}
static HFONT CreatePointFont(HWND hwndRef, int pointSize, const wchar_t* faceName, bool bold = false)
{
	HDC hdc = GetDC(hwndRef);
	int logPixelsY = GetDeviceCaps(hdc, LOGPIXELSY);
	ReleaseDC(hwndRef, hdc);
	int height = -MulDiv(pointSize, logPixelsY, 72);
	return CreateFontW(
		height, 0, 0, 0,
		bold ? FW_SEMIBOLD : FW_NORMAL,
		FALSE, FALSE, FALSE,
		DEFAULT_CHARSET, OUT_DEFAULT_PRECIS,
		CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY,
		DEFAULT_PITCH | FF_DONTCARE,
		faceName ? faceName : L"Segoe UI");
}
static void SetControlFont(HWND h, HFONT f)
{
	if (!h || !f) return;
	SendMessageW(h, WM_SETFONT, (WPARAM)f, TRUE);
}
static void AddTooltip(HWND hwndTip, HWND hwndTarget, const wchar_t* tipText)
{
	if (!hwndTip || !hwndTarget || !tipText) return;
	TOOLINFOW ti{};
	ti.cbSize = sizeof(ti);
	ti.uFlags = TTF_IDISHWND | TTF_SUBCLASS;
	ti.hwnd = GetParent(hwndTarget);
	ti.uId = (UINT_PTR)hwndTarget;
	ti.lpszText = (LPWSTR)tipText;
	SendMessageW(hwndTip, TTM_ADDTOOLW, 0, (LPARAM)&ti);
}
static std::wstring Utf8ToWide(const std::string& s)
{
	if (s.empty()) return {};
	int needed = MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), nullptr, 0);
	if (needed <= 0) return {};
	std::wstring out((size_t)needed, L'\0');
	MultiByteToWideChar(CP_UTF8, 0, s.c_str(), (int)s.size(), out.data(), needed);
	return out;
}
static std::string WideToUtf8(const std::wstring& ws)
{
	if (ws.empty()) return {};
	int needed = WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), nullptr, 0, nullptr, nullptr);
	if (needed <= 0) return {};
	std::string out((size_t)needed, '\0');
	WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), out.data(), needed, nullptr, nullptr);
	return out;
}

static inline bool SetErrFromStatus(std::string* outErrU8, const Status& st) {
	if (!outErrU8) return st.ok;
	std::wstring w = StatusToWString(st);
	*outErrU8 = w.empty() ? std::string() : WideToUtf8(w);
	return st.ok;
}


static std::string PathToUtf8(const fs::path& p)
{
	return WideToUtf8(p.wstring());
}
static wchar_t* DupWideForPostUtf8(const std::string& s)
{
	std::wstring ws = Utf8ToWide(s);
	size_t bytes = (ws.size() + 1) * sizeof(wchar_t);
	wchar_t* buf = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, bytes);
	if (!buf) return nullptr;
	memcpy(buf, ws.c_str(), ws.size() * sizeof(wchar_t));
	buf[ws.size()] = L'\0';
	return buf;
}
static wchar_t* DupWideForPostW(const std::wstring& ws)
{
	size_t bytes = (ws.size() + 1) * sizeof(wchar_t);
	wchar_t* buf = (wchar_t*)HeapAlloc(GetProcessHeap(), 0, bytes);
	if (!buf) return nullptr;
	memcpy(buf, ws.c_str(), ws.size() * sizeof(wchar_t));
	buf[ws.size()] = L'\0';
	return buf;
}
static bool PostHeapMessageW(UINT msg, wchar_t* payload)
{
	if (!payload) return false;
	if (!g_state || !g_state->hMainWnd)
	{
		HeapFree(GetProcessHeap(), 0, payload);
		return false;
	}
	if (PostMessageW(g_state->hMainWnd, msg, 0, (LPARAM)payload))
		return true;
	HeapFree(GetProcessHeap(), 0, payload);
	return false;
}
// Thread-safe log: posts to main thread
static void Log(const std::string& textUtf8)
{
	// Thread-safe log: post a unified UI event carrying wide text.
	wchar_t* w = DupWideForPostUtf8(textUtf8);
	if (!w) return;
	UiEvent* ev = new UiEvent();
	ev->kind = UiEventKind::LogAppendW;
	ev->text.assign(w);
	HeapFree(GetProcessHeap(), 0, w);
	PostUiEvent(ev);
}
static void LogSeparator(char ch = '_', size_t count = 150)
{
	//Usage Examples:
	//LogSeparator();                  // default
	//LogSeparator('-');               // dashed line
	//LogSeparator('=', 80);           // shorter

	std::string line(count, ch);
	Log(line + "\r\n");
}
static void PostProgressTextW(const std::wstring& textWide)
{
	UiEvent* ev = new UiEvent();
	ev->kind = UiEventKind::ProgressTextW;
	ev->text = textWide;
	PostUiEvent(ev);
}
static void PostProgressMarqueeOn()
{
	UiEvent* ev = new UiEvent();
	ev->kind = UiEventKind::ProgressMarqueeOn;
	PostUiEvent(ev);
}
static void PostProgressMarqueeOff()
{
	UiEvent* ev = new UiEvent();
	ev->kind = UiEventKind::ProgressMarqueeOff;
	PostUiEvent(ev);
}
static void PostProgressInit(size_t total)
{
	UiEvent* ev = new UiEvent();
	ev->kind = UiEventKind::ProgressInit;
	ev->u1 = total;
	PostUiEvent(ev);
}
static void PostProgressSet(size_t pos)
{
	UiEvent* ev = new UiEvent();
	ev->kind = UiEventKind::ProgressSet;
	ev->u1 = pos;
	PostUiEvent(ev);
}
static bool CheckAndHandleCancel(const CancelToken& cancel, const char* logLine)
{
	if (!cancel.IsCanceled()) return false;
	if (logLine) Log(logLine);
	PostProgressTextW(L"Canceled.");
	return true;
}
static bool IsCanceledNoNotify(const CancelToken& cancel)
{
	return cancel.IsCanceled();
}
// --------------------------------------------------
// Scope-exit helper
// --------------------------------------------------
struct ScopeExit
{
	std::function<void()> fn;
	~ScopeExit() { if (fn) fn(); }
};

// --------------------------------------------------
// RAII wrappers for common Win32/Crypto handle types
// --------------------------------------------------

struct BcryptAlgHandle
{
	BCRYPT_ALG_HANDLE h = nullptr;
	BcryptAlgHandle() = default;
	explicit BcryptAlgHandle(BCRYPT_ALG_HANDLE handle) : h(handle) {}
	~BcryptAlgHandle() { if (h) { BCryptCloseAlgorithmProvider(h, 0); h = nullptr; } }
	BcryptAlgHandle(const BcryptAlgHandle&) = delete;
	BcryptAlgHandle& operator=(const BcryptAlgHandle&) = delete;
	BcryptAlgHandle(BcryptAlgHandle&& o) noexcept : h(o.h) { o.h = nullptr; }
	BcryptAlgHandle& operator=(BcryptAlgHandle&& o) noexcept { if (this != &o) { if (h) BCryptCloseAlgorithmProvider(h, 0); h = o.h; o.h = nullptr; } return *this; }
	operator BCRYPT_ALG_HANDLE() const { return h; }
	bool valid() const { return h != nullptr; }
};

struct BcryptHashHandle
{
	BCRYPT_HASH_HANDLE h = nullptr;
	BcryptHashHandle() = default;
	explicit BcryptHashHandle(BCRYPT_HASH_HANDLE handle) : h(handle) {}
	~BcryptHashHandle() { if (h) { BCryptDestroyHash(h); h = nullptr; } }
	BcryptHashHandle(const BcryptHashHandle&) = delete;
	BcryptHashHandle& operator=(const BcryptHashHandle&) = delete;
	BcryptHashHandle(BcryptHashHandle&& o) noexcept : h(o.h) { o.h = nullptr; }
	BcryptHashHandle& operator=(BcryptHashHandle&& o) noexcept { if (this != &o) { if (h) BCryptDestroyHash(h); h = o.h; o.h = nullptr; } return *this; }
	operator BCRYPT_HASH_HANDLE() const { return h; }
	bool valid() const { return h != nullptr; }
};

struct CryptMsgHandle
{
	HCRYPTMSG h = nullptr;
	CryptMsgHandle() = default;
	explicit CryptMsgHandle(HCRYPTMSG handle) : h(handle) {}
	~CryptMsgHandle() { if (h) { CryptMsgClose(h); h = nullptr; } }
	CryptMsgHandle(const CryptMsgHandle&) = delete;
	CryptMsgHandle& operator=(const CryptMsgHandle&) = delete;
	CryptMsgHandle(CryptMsgHandle&& o) noexcept : h(o.h) { o.h = nullptr; }
	CryptMsgHandle& operator=(CryptMsgHandle&& o) noexcept { if (this != &o) { if (h) CryptMsgClose(h); h = o.h; o.h = nullptr; } return *this; }
	operator HCRYPTMSG() const { return h; }
	bool valid() const { return h != nullptr; }
};

struct CertStoreHandle
{
	HCERTSTORE h = nullptr;
	CertStoreHandle() = default;
	explicit CertStoreHandle(HCERTSTORE handle) : h(handle) {}
	~CertStoreHandle() { if (h) { CertCloseStore(h, 0); h = nullptr; } }
	CertStoreHandle(const CertStoreHandle&) = delete;
	CertStoreHandle& operator=(const CertStoreHandle&) = delete;
	CertStoreHandle(CertStoreHandle&& o) noexcept : h(o.h) { o.h = nullptr; }
	CertStoreHandle& operator=(CertStoreHandle&& o) noexcept { if (this != &o) { if (h) CertCloseStore(h, 0); h = o.h; o.h = nullptr; } return *this; }
	operator HCERTSTORE() const { return h; }
	bool valid() const { return h != nullptr; }
};

struct CertContextHandle
{
	PCCERT_CONTEXT p = nullptr;
	CertContextHandle() = default;
	explicit CertContextHandle(PCCERT_CONTEXT ctx) : p(ctx) {}
	~CertContextHandle() { if (p) { CertFreeCertificateContext(p); p = nullptr; } }
	CertContextHandle(const CertContextHandle&) = delete;
	CertContextHandle& operator=(const CertContextHandle&) = delete;
	CertContextHandle(CertContextHandle&& o) noexcept : p(o.p) { o.p = nullptr; }
	CertContextHandle& operator=(CertContextHandle&& o) noexcept { if (this != &o) { if (p) CertFreeCertificateContext(p); p = o.p; o.p = nullptr; } return *this; }
	operator PCCERT_CONTEXT() const { return p; }
	bool valid() const { return p != nullptr; }
};

// --------------------------------------------------
// String cleanup helpers (for edit box paths)
// --------------------------------------------------
static void TrimInPlace(std::wstring& s)
{
	auto isws = [](wchar_t c) { return c == L' ' || c == L'	' || c == L'\r' || c == L'\n'; };
	size_t b = 0;
	while (b < s.size() && isws(s[b])) ++b;
	size_t e = s.size();
	while (e > b && isws(s[e - 1])) --e;
	s = s.substr(b, e - b);
}
static void StripSurroundingQuotes(std::wstring& s)
{
	if (s.size() >= 2 && s.front() == L'"' && s.back() == L'"')
		s = s.substr(1, s.size() - 2);
}
// --------------------------------------------------
// Preflight validation
// --------------------------------------------------
struct PreflightResult
{
	bool ok = false;
	fs::path localBase;
	fs::path localSyncRoot;
	std::vector<std::string> errors;
};
struct SyncConfig
{
	std::string remoteHost;
	std::string remoteRootPath;   // NOTE: must end with '/' because JoinUrl/manifest normalization assume it.
	std::string manifestUrl;
	fs::path localBase;
	fs::path localSyncRoot;
};
static PreflightResult ValidateFolderSelection(const std::wstring& folderWs)
{
	PreflightResult r{};
	std::error_code ec;
	if (folderWs.empty())
	{
		r.errors.push_back("ERROR: Istaria Base Game Folder not selected. You need to choose a valid Istaria game folder to sync.\r\n");
		return r;
	}
	if (!fs::exists(folderWs, ec) || !fs::is_directory(folderWs, ec))
	{
		r.errors.push_back("ERROR: Selected folder '" + WideToUtf8(folderWs) + "' does not exist. You need to choose a valid Istaria game folder to sync.\r\n");
		if (ec) r.errors.push_back("fs::exists/is_directory error: " + std::string(ec.message()) + "\r\n");
		return r;
	}
	r.localBase = folderWs;
	r.localSyncRoot = r.localBase / "resources_override" / "mappack";
	fs::path istariaExe = r.localBase / "istaria.exe";
	ec.clear();
	if (!fs::exists(istariaExe, ec))
	{
		r.errors.push_back("ERROR: Selected folder does not contain istaria.exe. You need to choose a valid Istaria game folder to sync.\r\n");
		if (ec) r.errors.push_back("fs::exists error: " + std::string(ec.message()) + "\r\n");
		return r;
	}
	ec.clear();
	if (!fs::exists(r.localSyncRoot, ec))
	{
		ec.clear();
		fs::create_directories(r.localSyncRoot, ec);
		if (ec)
		{
			r.errors.push_back("ERROR: Failed to create resources_override folder\r\n");
			r.errors.push_back("Folder:   " + PathToUtf8(r.localBase) + "\r\n");
			r.errors.push_back("Target:   " + PathToUtf8(r.localSyncRoot) + "\r\n");
			r.errors.push_back("create_directories error: " + std::string(ec.message()) + "\r\n");
			return r;
		}
	}
	else
	{
		ec.clear();
		if (!fs::is_directory(r.localSyncRoot, ec) || ec)
		{
			r.errors.push_back("ERROR: resources_override exists but is not a directory\r\n");
			r.errors.push_back("Folder:   " + PathToUtf8(r.localBase) + "\r\n");
			r.errors.push_back("Path:     " + PathToUtf8(r.localSyncRoot) + "\r\n");
			if (ec) r.errors.push_back("is_directory error: " + std::string(ec.message()) + "\r\n");
			return r;
		}
	}
	r.ok = true;
	return r;
}
// --------------------------------------------------
// SHA-256 (Windows CNG / BCrypt)
// --------------------------------------------------
static bool Sha256FileHexLower(const fs::path& filePath, std::string& outHex)
{
	outHex.clear();

	BCRYPT_ALG_HANDLE rawAlg = nullptr;
	BCRYPT_HASH_HANDLE rawHash = nullptr;

	NTSTATUS st = BCryptOpenAlgorithmProvider(&rawAlg, BCRYPT_SHA256_ALGORITHM, nullptr, 0);
	if (st != 0) return false;
	BcryptAlgHandle hAlg(rawAlg);

	DWORD objLen = 0, cbData = 0, hashLen = 0;

	st = BCryptGetProperty(hAlg, BCRYPT_OBJECT_LENGTH, (PUCHAR)&objLen, sizeof(objLen), &cbData, 0);
	if (st != 0) return false;

	st = BCryptGetProperty(hAlg, BCRYPT_HASH_LENGTH, (PUCHAR)&hashLen, sizeof(hashLen), &cbData, 0);
	if (st != 0) return false;

	std::vector<UCHAR> obj(objLen);
	std::vector<UCHAR> hash(hashLen);

	st = BCryptCreateHash(hAlg, &rawHash, obj.data(), (ULONG)obj.size(), nullptr, 0, 0);
	if (st != 0) return false;
	BcryptHashHandle hHash(rawHash);

	std::ifstream f(filePath, std::ios::binary);
	if (!f) return false;

	std::vector<char> buf(64 * 1024);
	while (f)
	{
		f.read(buf.data(), (std::streamsize)buf.size());
		std::streamsize got = f.gcount();
		if (got <= 0) break;
		st = BCryptHashData(hHash, (PUCHAR)buf.data(), (ULONG)got, 0);
		if (st != 0) return false;
	}

	st = BCryptFinishHash(hHash, hash.data(), (ULONG)hash.size(), 0);
	if (st != 0) return false;

	static const char* hexd = "0123456789abcdef";
	outHex.resize(hash.size() * 2);
	for (size_t i = 0; i < hash.size(); ++i)
	{
		outHex[i * 2 + 0] = hexd[(hash[i] >> 4) & 0xF];
		outHex[i * 2 + 1] = hexd[hash[i] & 0xF];
	}
	return true;
}
static bool StartsWith(const std::string& s, const std::string& prefix)
{
	return s.size() >= prefix.size() && memcmp(s.data(), prefix.data(), prefix.size()) == 0;
}
static bool EqualIcaseAscii(const std::string& a, const std::string& b)
{
	if (a.size() != b.size()) return false;
	for (size_t i = 0; i < a.size(); ++i)
	{
		char ca = a[i], cb = b[i];
		if (ca >= 'A' && ca <= 'Z') ca = char(ca - 'A' + 'a');
		if (cb >= 'A' && cb <= 'Z') cb = char(cb - 'A' + 'a');
		if (ca != cb) return false;
	}
	return true;
}

// --------------------------------------------------
// Path normalization helpers
// --------------------------------------------------
// Converts '\' -> '/', strips leading '/' or '\', collapses repeated slashes,
// and resolves "." and ".." segments (never allowing ".." to escape above root).
static std::string NormalizePathGeneric(std::string_view in)
{
	std::string s;
	s.reserve(in.size());
	for (char c : in)
	{
		if (c == '\\') c = '/';
		s.push_back(c);
	}
	while (!s.empty() && (s.front() == '/')) s.erase(s.begin());

	// Collapse multiple '/'
	{
		std::string out;
		out.reserve(s.size());
		bool prevSlash = false;
		for (char c : s)
		{
			if (c == '/')
			{
				if (prevSlash) continue;
				prevSlash = true;
			}
			else
			{
				prevSlash = false;
			}
			out.push_back(c);
		}
		s.swap(out);
	}

	// Resolve dot segments
	std::vector<std::string> parts;
	parts.reserve(32);
	size_t i = 0;
	while (i <= s.size())
	{
		size_t j = s.find('/', i);
		if (j == std::string::npos) j = s.size();
		std::string seg = s.substr(i, j - i);
		if (!seg.empty() && seg != ".")
		{
			if (seg == "..")
			{
				if (!parts.empty()) parts.pop_back();
				// else: ignore attempts to escape above root
			}
			else
			{
				parts.push_back(std::move(seg));
			}
		}
		i = j + 1;
		if (j == s.size()) break;
	}
	std::string out;
	for (size_t k = 0; k < parts.size(); ++k)
	{
		if (k) out.push_back('/');
		out += parts[k];
	}
	return out;
}

// Normalizes a manifest "remotePath" into a stable relative path under our sync root.
// Strips leading '/', converts '\' -> '/', resolves "."/"..", and removes any leading
// "resources_override/mappack/", "mappack/", or "resources_override/" prefixes.
static std::string NormalizeManifestRel(std::string_view manifestRelPath)
{
	std::string rel = NormalizePathGeneric(manifestRelPath);

	const std::string p1 = "resources_override/mappack/";
	const std::string p2 = "mappack/";
	const std::string p3 = "resources_override/";
	if (StartsWith(rel, p1)) rel = rel.substr(p1.size());
	else if (StartsWith(rel, p2)) rel = rel.substr(p2.size());
	else if (StartsWith(rel, p3)) rel = rel.substr(p3.size());

	// One more pass in case stripping exposed "./" or "../"
	rel = NormalizePathGeneric(rel);
	return rel;
}

// Strict variant: rejects attempts to escape above root and disallows absolute/UNC/drive-qualified paths.
static bool NormalizeManifestRelStrict(std::string_view manifestRelPath, std::string& outRel, std::string* outErr)
{
	outRel.clear();
	if (outErr) outErr->clear();

	// Basic absolute/UNC/drive checks on the raw string (before normalization).
	// - Reject UNC like "\\" or "//"
	// - Reject drive-qualified like "C:\..." or "C:/..."
	// - Reject embedded NUL
	for (char c : manifestRelPath)
	{
		if (c == '\0')
		{
			if (outErr) *outErr = "path contains NUL";
			return false;
		}
	}
	if (manifestRelPath.size() >= 2)
	{
		const char a = manifestRelPath[0];
		const char b = manifestRelPath[1];
		if ((a == '\\' && b == '\\') || (a == '/' && b == '/'))
		{
			if (outErr) *outErr = "UNC paths not allowed";
			return false;
		}
	}
	if (manifestRelPath.size() >= 2 && ((manifestRelPath[1] == ':') && ((manifestRelPath[0] >= 'A' && manifestRelPath[0] <= 'Z') || (manifestRelPath[0] >= 'a' && manifestRelPath[0] <= 'z'))))
	{
		if (outErr) *outErr = "drive-qualified paths not allowed";
		return false;
	}

	// Normalize to generic '/' separators and resolve dot segments, but fail if ".." would escape above root.
	std::string s;
	s.reserve(manifestRelPath.size());
	for (char c : manifestRelPath)
		s.push_back(c == '\\' ? '/' : c);

	// Strip leading '/' (treat as relative for our purposes); still reject UNC was handled above.
	while (!s.empty() && s.front() == '/')
		s.erase(s.begin());

	std::vector<std::string> parts;
	parts.reserve(16);

	size_t i = 0;
	while (i <= s.size())
	{
		const size_t j = s.find('/', i);
		const size_t end = (j == std::string::npos) ? s.size() : j;
		std::string seg = s.substr(i, end - i);

		// Skip empty / "." segments
		if (!seg.empty() && seg != ".")
		{
			if (seg == "..")
			{
				if (parts.empty())
				{
					if (outErr) *outErr = "path attempts to escape root";
					return false;
				}
				parts.pop_back();
			}
			else
			{
				// Disallow ':' anywhere to avoid drive injection and ADS tricks.
				if (seg.find(':') != std::string::npos)
				{
					if (outErr) *outErr = "path contains ':'";
					return false;
				}
				parts.push_back(std::move(seg));
			}
		}

		if (j == std::string::npos) break;
		i = j + 1;
	}

	std::string rel;
	for (size_t k = 0; k < parts.size(); ++k)
	{
		if (k) rel.push_back('/');
		rel += parts[k];
	}

	// Remove known prefixes after normalization.
	const std::string p1 = "resources_override/mappack/";
	const std::string p2 = "resources_override/";
	const std::string p3 = "mappack/";
	if (StartsWith(rel, p1)) rel = rel.substr(p1.size());
	else if (StartsWith(rel, p2)) rel = rel.substr(p2.size());
	else if (StartsWith(rel, p3)) rel = rel.substr(p3.size());

	// Final safety: reject empty and any remaining dot segments.
	if (rel.empty())
	{
		if (outErr) *outErr = "path resolves to empty";
		return false;
	}
	if (rel.find("..") != std::string::npos)
	{
		// This is conservative; Normalize above should have removed legit dots; remaining ".." indicates something odd.
		if (outErr) *outErr = "path contains '..'";
		return false;
	}

	outRel = std::move(rel);
	return true;
}


static fs::path MakeDestPath(const fs::path& installRoot, std::string_view validatedRelPath)
{
	// validatedRelPath is produced by ValidateAndNormalizeManifest() and is guaranteed to be relative and safe.
	return installRoot / "resources_override" / "mappack" / fs::path(std::string(validatedRelPath));
}
// --------------------------------------------------
// Manifest parsing
// --------------------------------------------------
struct ManifestEntry
{
	std::string remotePath;  // normalized remote path (generic, '/' separators, no leading '/')
	std::string relPath;     // normalized relative path under resources_override/mappack/
	std::string sha256;      // expected SHA-256 (hex)
};

struct ManifestRawEntry
{
	std::string path;   // path string exactly as in manifest (UTF-8)
	std::string sha256; // sha256 exactly as in manifest (hex)
};


// -----------------------------
// Tiny JSON helpers (dependency-free)
// Goal: Parse just {"files":[{"path":"...","sha256":"..."}, ...]} robustly.
// - Does NOT assume key ordering
// - Ignores unknown fields (any JSON value type)
// - Supports \uXXXX escapes in strings
// - Includes depth limits for pathological inputs
// -----------------------------
static void SkipWs(const std::string& s, size_t& i)
{
	while (i < s.size())
	{
		const char c = s[i];
		if (c == ' ' || c == '\t' || c == '\r' || c == '\n') { ++i; continue; }
		break;
	}
}

static bool HexNibble(char c, uint32_t& out)
{
	if (c >= '0' && c <= '9') { out = uint32_t(c - '0'); return true; }
	if (c >= 'a' && c <= 'f') { out = uint32_t(c - 'a' + 10); return true; }
	if (c >= 'A' && c <= 'F') { out = uint32_t(c - 'A' + 10); return true; }
	return false;
}

static bool ReadHex4(const std::string& s, size_t& i, uint32_t& outCodePoint)
{
	outCodePoint = 0;
	for (int k = 0; k < 4; ++k)
	{
		if (i >= s.size()) return false;
		uint32_t n = 0;
		if (!HexNibble(s[i++], n)) return false;
		outCodePoint = (outCodePoint << 4) | n;
	}
	return true;
}

static void AppendUtf8(uint32_t cp, std::string& out)
{
	// Minimal UTF-8 encoding
	if (cp <= 0x7F)
	{
		out.push_back(char(cp));
	}
	else if (cp <= 0x7FF)
	{
		out.push_back(char(0xC0 | ((cp >> 6) & 0x1F)));
		out.push_back(char(0x80 | (cp & 0x3F)));
	}
	else if (cp <= 0xFFFF)
	{
		out.push_back(char(0xE0 | ((cp >> 12) & 0x0F)));
		out.push_back(char(0x80 | ((cp >> 6) & 0x3F)));
		out.push_back(char(0x80 | (cp & 0x3F)));
	}
	else
	{
		out.push_back(char(0xF0 | ((cp >> 18) & 0x07)));
		out.push_back(char(0x80 | ((cp >> 12) & 0x3F)));
		out.push_back(char(0x80 | ((cp >> 6) & 0x3F)));
		out.push_back(char(0x80 | (cp & 0x3F)));
	}
}

static bool ReadJsonString(const std::string& s, size_t& i, std::string& out, std::string* outErr)
{
	out.clear();
	if (i >= s.size() || s[i] != '"') { if (outErr) *outErr = "expected string"; return false; }
	++i;
	while (i < s.size())
	{
		const char c = s[i++];
		if (c == '"') return true;
		if (c == '\\')
		{
			if (i >= s.size()) { if (outErr) *outErr = "unterminated escape"; return false; }
			const char e = s[i++];
			switch (e)
			{
			case '"': out.push_back('"'); break;
			case '\\': out.push_back('\\'); break;
			case '/': out.push_back('/'); break;
			case 'b': out.push_back('\b'); break;
			case 'f': out.push_back('\f'); break;
			case 'n': out.push_back('\n'); break;
			case 'r': out.push_back('\r'); break;
			case 't': out.push_back('\t'); break;
			case 'u':
			{
				uint32_t cp = 0;
				if (!ReadHex4(s, i, cp)) { if (outErr) *outErr = "bad \\uXXXX escape"; return false; }
				// Surrogate pair handling
				if (cp >= 0xD800 && cp <= 0xDBFF)
				{
					// high surrogate; expect \uYYYY
					if (i + 2 <= s.size() && s[i] == '\\' && s[i + 1] == 'u')
					{
						i += 2;
						uint32_t low = 0;
						if (!ReadHex4(s, i, low)) { if (outErr) *outErr = "bad low surrogate"; return false; }
						if (low >= 0xDC00 && low <= 0xDFFF)
						{
							cp = 0x10000 + (((cp - 0xD800) << 10) | (low - 0xDC00));
						}
						else
						{
							if (outErr) *outErr = "invalid surrogate pair";
							return false;
						}
					}
					else
					{
						if (outErr) *outErr = "missing low surrogate";
						return false;
					}
				}
				AppendUtf8(cp, out);
				break;
			}
			default:
				if (outErr) *outErr = "unsupported escape sequence";
				return false;
			}
		}
		else
		{
			// Control chars are not valid in JSON strings
			if ((unsigned char)c < 0x20)
			{
				if (outErr) *outErr = "control character in string";
				return false;
			}
			out.push_back(c);
		}
	}
	if (outErr) *outErr = "unterminated string";
	return false;
}

static bool ReadJsonStringValue(const std::string& s, size_t& i, std::string& out)
{
	return ReadJsonString(s, i, out, nullptr);
}

static bool SkipJsonValue(const std::string& s, size_t& i, int depth, std::string* outErr);

static bool SkipJsonNumber(const std::string& s, size_t& i)
{
	// JSON number: -? (0|[1-9][0-9]*) (.[0-9]+)? ([eE][+-]?[0-9]+)?
	size_t start = i;
	if (i < s.size() && (s[i] == '-')) ++i;
	if (i >= s.size()) { i = start; return false; }
	if (s[i] == '0') { ++i; }
	else
	{
		if (s[i] < '1' || s[i] > '9') { i = start; return false; }
		while (i < s.size() && (s[i] >= '0' && s[i] <= '9')) ++i;
	}
	if (i < s.size() && s[i] == '.')
	{
		++i;
		if (i >= s.size() || s[i] < '0' || s[i] > '9') { i = start; return false; }
		while (i < s.size() && (s[i] >= '0' && s[i] <= '9')) ++i;
	}
	if (i < s.size() && (s[i] == 'e' || s[i] == 'E'))
	{
		++i;
		if (i < s.size() && (s[i] == '+' || s[i] == '-')) ++i;
		if (i >= s.size() || s[i] < '0' || s[i] > '9') { i = start; return false; }
		while (i < s.size() && (s[i] >= '0' && s[i] <= '9')) ++i;
	}
	return i > start;
}

static bool SkipJsonArray(const std::string& s, size_t& i, int depth, std::string* outErr)
{
	if (i >= s.size() || s[i] != '[') { if (outErr) *outErr = "expected '['"; return false; }
	++i;
	SkipWs(s, i);
	if (i < s.size() && s[i] == ']') { ++i; return true; }

	for (;;)
	{
		SkipWs(s, i);
		if (!SkipJsonValue(s, i, depth + 1, outErr)) return false;
		SkipWs(s, i);
		if (i >= s.size()) { if (outErr) *outErr = "unterminated array"; return false; }
		if (s[i] == ',') { ++i; continue; }
		if (s[i] == ']') { ++i; return true; }
		if (outErr) *outErr = "expected ',' or ']'";
		return false;
	}
}

static bool SkipJsonObject(const std::string& s, size_t& i, int depth, std::string* outErr)
{
	if (i >= s.size() || s[i] != '{') { if (outErr) *outErr = "expected '{'"; return false; }
	++i;
	SkipWs(s, i);
	if (i < s.size() && s[i] == '}') { ++i; return true; }

	for (;;)
	{
		SkipWs(s, i);
		std::string key;
		if (!ReadJsonString(s, i, key, outErr)) return false;
		SkipWs(s, i);
		if (i >= s.size() || s[i] != ':') { if (outErr) *outErr = "expected ':' after key"; return false; }
		++i;
		SkipWs(s, i);
		if (!SkipJsonValue(s, i, depth + 1, outErr)) return false;
		SkipWs(s, i);
		if (i >= s.size()) { if (outErr) *outErr = "unterminated object"; return false; }
		if (s[i] == ',') { ++i; continue; }
		if (s[i] == '}') { ++i; return true; }
		if (outErr) *outErr = "expected ',' or '}'";
		return false;
	}
}

static bool SkipJsonValue(const std::string& s, size_t& i, int depth, std::string* outErr)
{
	if (depth > 64) { if (outErr) *outErr = "JSON nesting too deep"; return false; }
	if (i >= s.size()) { if (outErr) *outErr = "unexpected end of JSON"; return false; }

	const char c = s[i];
	if (c == '"')
	{
		std::string tmp;
		return ReadJsonString(s, i, tmp, outErr);
	}
	if (c == '{') return SkipJsonObject(s, i, depth, outErr);
	if (c == '[') return SkipJsonArray(s, i, depth, outErr);

	// Literals
	if (c == 't')
	{
		if (s.compare(i, 4, "true") == 0) { i += 4; return true; }
		if (outErr) *outErr = "invalid literal";
		return false;
	}
	if (c == 'f')
	{
		if (s.compare(i, 5, "false") == 0) { i += 5; return true; }
		if (outErr) *outErr = "invalid literal";
		return false;
	}
	if (c == 'n')
	{
		if (s.compare(i, 4, "null") == 0) { i += 4; return true; }
		if (outErr) *outErr = "invalid literal";
		return false;
	}

	// Number
	if (c == '-' || (c >= '0' && c <= '9'))
	{
		if (SkipJsonNumber(s, i)) return true;
		if (outErr) *outErr = "invalid number";
		return false;
	}

	if (outErr) *outErr = "unexpected token";
	return false;
}

static bool HasDotDotSegment(const std::string& rawPath)
{
	// Detect ".." segments before normalization so we can reject rather than silently clamp.
	size_t i = 0;
	while (i <= rawPath.size())
	{
		size_t j = rawPath.find_first_of("/\\", i);
		if (j == std::string::npos) j = rawPath.size();
		const std::string seg = rawPath.substr(i, j - i);
		if (seg == "..") return true;
		i = j + 1;
	}
	return false;
}

static bool IsSafeManifestPath(const std::string& rawPath, std::string* outErr)
{
	if (rawPath.empty()) { if (outErr) *outErr = "empty path"; return false; }

	// Reject obvious absolute/UNC/drive paths
	if (rawPath.size() >= 2 && std::isalpha((unsigned char)rawPath[0]) && rawPath[1] == ':')
	{
		if (outErr) *outErr = "absolute drive path not allowed";
		return false;
	}
	if (rawPath.rfind("\\\\", 0) == 0 || rawPath.rfind("//", 0) == 0)
	{
		if (outErr) *outErr = "UNC path not allowed";
		return false;
	}

	// Reject traversal segments
	if (HasDotDotSegment(rawPath))
	{
		if (outErr) *outErr = "path traversal '..' not allowed";
		return false;
	}

	// Basic control character check
	for (unsigned char ch : rawPath)
	{
		if (ch < 0x20) { if (outErr) *outErr = "control character in path"; return false; }
	}
	return true;
}

static bool IsHex64(const std::string& s)
{
	if (s.size() != 64) return false;
	for (char c : s)
	{
		if (!((c >= '0' && c <= '9') ||
			(c >= 'a' && c <= 'f') ||
			(c >= 'A' && c <= 'F')))
			return false;
	}
	return true;
}

// New manifest format: top-level JSON object with "files": [ { "path": "...", "sha256": "..." }, ... ]

// Forward declarations (needed for Status wrappers)
struct ManifestRawEntry;
struct ManifestEntry;
static bool ParseManifestRaw(const std::string& jsonText, std::vector<ManifestRawEntry>& outFiles, std::string* outErr);

// Validate+normalize has a 4-arg core version (tracks duplicates) and a convenience 3-arg overload.
static bool ValidateAndNormalizeManifest(const std::vector<ManifestRawEntry>& rawFiles,
	std::vector<ManifestEntry>& outWorkList,
	std::unordered_set<std::string>& outManifestRelSet,
	std::string* outErr);

// Convenience overload used by Status wrapper and other call sites.
static bool ValidateAndNormalizeManifest(const std::vector<ManifestRawEntry>& rawFiles,
	std::vector<ManifestEntry>& outWorkList,
	std::string* outErr);

static bool ValidateAndNormalizeManifest(const std::vector<ManifestRawEntry>& rawFiles,
	std::vector<ManifestEntry>& outWorkList,
	std::string* outErr)
{
	std::unordered_set<std::string> dummySet;
	return ValidateAndNormalizeManifest(rawFiles, outWorkList, dummySet, outErr);
}
static Status ParseManifestRaw_Status(const std::string& jsonText, std::vector<ManifestRawEntry>& outEntries)
{
	std::string err;
	std::vector<ManifestRawEntry> tmp;
	if (!ParseManifestRaw(jsonText, tmp, &err)) {
		return Status::FailMsg(Utf8ToWide(err));
	}
	outEntries.swap(tmp);
	return Status::Ok();
}

static bool ParseManifestRaw(const std::string& jsonText, std::vector<ManifestRawEntry>& outFiles, std::string* outErr)
{
	outFiles.clear();
	if (outErr) outErr->clear();

	// Size guard (network input)
	static constexpr size_t kMaxManifestBytes = 50ull * 1024ull * 1024ull; // 50 MiB
	if (jsonText.size() > kMaxManifestBytes)
	{
		if (outErr) *outErr = "manifest too large";
		return false;
	}

	const std::string& s = jsonText;
	size_t i = 0;
	SkipWs(s, i);
	if (i >= s.size() || s[i] != '{')
	{
		if (outErr) *outErr = "expected top-level object";
		return false;
	}
	++i;

	bool foundFiles = false;

	for (;;)
	{
		SkipWs(s, i);
		if (i >= s.size()) { if (outErr) *outErr = "unterminated top-level object"; return false; }
		if (s[i] == '}') { ++i; break; }

		std::string key;
		if (!ReadJsonString(s, i, key, outErr)) return false;

		SkipWs(s, i);
		if (i >= s.size() || s[i] != ':') { if (outErr) *outErr = "expected ':' after key"; return false; }
		++i;
		SkipWs(s, i);

		if (key == "files")
		{
			foundFiles = true;
			if (i >= s.size() || s[i] != '[') { if (outErr) *outErr = "expected '[' for files"; return false; }
			++i;
			SkipWs(s, i);
			if (i < s.size() && s[i] == ']') { ++i; } // empty array allowed
			else
			{
				for (;;)
				{
					SkipWs(s, i);
					if (i >= s.size()) { if (outErr) *outErr = "unterminated files array"; return false; }
					if (s[i] != '{') { if (outErr) *outErr = "expected object in files array"; return false; }
					++i;

					std::string pathVal, hashVal;
					for (;;)
					{
						SkipWs(s, i);
						if (i >= s.size()) { if (outErr) *outErr = "unterminated file object"; return false; }
						if (s[i] == '}') { ++i; break; }

						std::string fkey;
						if (!ReadJsonString(s, i, fkey, outErr)) return false;
						SkipWs(s, i);
						if (i >= s.size() || s[i] != ':') { if (outErr) *outErr = "expected ':' in file object"; return false; }
						++i;
						SkipWs(s, i);

						if (fkey == "path")
						{
							if (!ReadJsonString(s, i, pathVal, outErr)) return false;
						}
						else if (fkey == "sha256" || fkey == "hash")
						{
							if (!ReadJsonString(s, i, hashVal, outErr)) return false;
						}
						else
						{
							// Skip any unknown field (string/number/object/array/bool/null)
							if (!SkipJsonValue(s, i, 0, outErr)) return false;
						}

						SkipWs(s, i);
						if (i >= s.size()) { if (outErr) *outErr = "unterminated file object"; return false; }
						if (s[i] == ',') { ++i; continue; }
						if (s[i] == '}') { ++i; break; }
						if (outErr) *outErr = "expected ',' or '}' in file object";
						return false;
					}

					if (!pathVal.empty() && !hashVal.empty())
					{
						// Validation is performed in ValidateAndNormalizeManifest().
						ManifestRawEntry e;
						e.path = pathVal; // normalized later by ValidateAndNormalizeManifest()
						e.sha256 = hashVal;
						outFiles.push_back(std::move(e));
					}

					SkipWs(s, i);
					if (i >= s.size()) { if (outErr) *outErr = "unterminated files array"; return false; }
					if (s[i] == ',') { ++i; continue; }
					if (s[i] == ']') { ++i; break; }
					if (outErr) *outErr = "expected ',' or ']' in files array";
					return false;
				}
			}
		}
		else
		{
			// Skip unknown top-level fields
			if (!SkipJsonValue(s, i, 0, outErr)) return false;
		}

		SkipWs(s, i);
		if (i >= s.size()) { if (outErr) *outErr = "unterminated top-level object"; return false; }
		if (s[i] == ',') { ++i; continue; }
		if (s[i] == '}') { ++i; break; }
		if (outErr) *outErr = "expected ',' or '}' in top-level object";
		return false;
	}

	if (!foundFiles)
	{
		if (outErr) *outErr = "missing 'files' key";
		return false;
	}
	// Empty files list is allowed but likely an error; keep it as failure to be safe.
	if (outFiles.empty())
	{
		if (outErr) *outErr = "'files' array is empty";
		return false;
	}
	return true;
}




static Status ValidateAndNormalizeManifest_Status(
	const std::vector<ManifestRawEntry>& rawEntries,
	std::vector<ManifestEntry>& outEntries)
{
	std::string err;
	std::vector<ManifestEntry> tmp;
	if (!ValidateAndNormalizeManifest(rawEntries, tmp, &err)) {
		return Status::FailMsg(Utf8ToWide(err));
	}
	outEntries.swap(tmp);
	return Status::Ok();
}

static bool ValidateAndNormalizeManifest(const std::vector<ManifestRawEntry>& rawFiles,
	std::vector<ManifestEntry>& outWorkList,
	std::unordered_set<std::string>& outManifestRelSet,
	std::string* outErr)
{
	outWorkList.clear();
	outManifestRelSet.clear();
	if (outErr) outErr->clear();

	if (rawFiles.empty())
	{
		if (outErr) *outErr = "'files' array is empty";
		return false;
	}

	outWorkList.reserve(rawFiles.size());
	outManifestRelSet.reserve(rawFiles.size() * 2);

	for (const auto& rf : rawFiles)
	{
		if (rf.path.empty())
		{
			if (outErr) *outErr = "file entry has empty path";
			return false;
		}
		if (!IsHex64(rf.sha256))
		{
			if (outErr) *outErr = "invalid sha256 for path: " + rf.path;
			return false;
		}

		ManifestEntry e;
		e.remotePath = NormalizePathGeneric(rf.path);

		std::string relErr;
		if (!NormalizeManifestRelStrict(e.remotePath, e.relPath, &relErr))
		{
			if (outErr) *outErr = "unsafe path: " + (relErr.empty() ? rf.path : relErr) + " (" + rf.path + ")";
			return false;
		}

		e.sha256 = rf.sha256;

		if (!outManifestRelSet.insert(e.relPath).second)
		{
			if (outErr) *outErr = "duplicate path in manifest: " + e.relPath;
			return false;
		}

		outWorkList.push_back(std::move(e));
	}

	std::sort(outWorkList.begin(), outWorkList.end(),
		[](const ManifestEntry& a, const ManifestEntry& b) { return a.relPath < b.relPath; });

	return true;
}


static bool CrackUrlWinHttp(const std::string& urlUtf8, std::wstring& outHost, std::wstring& outPath, INTERNET_PORT& outPort, bool& outSecure, std::string* outErr)
{
	outHost.clear();
	outPath.clear();
	outPort = 0;
	outSecure = false;

	std::wstring url = Utf8ToWide(urlUtf8);
	URL_COMPONENTS uc{};
	uc.dwStructSize = sizeof(uc);
	uc.dwHostNameLength = (DWORD)-1;
	uc.dwUrlPathLength = (DWORD)-1;
	uc.dwExtraInfoLength = (DWORD)-1;

	if (!WinHttpCrackUrl(url.c_str(), 0, 0, &uc))
	{
		if (outErr) *outErr = "WinHttpCrackUrl failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	if (uc.nScheme != INTERNET_SCHEME_HTTP && uc.nScheme != INTERNET_SCHEME_HTTPS)
	{
		if (outErr) *outErr = "Unsupported URL scheme (only http/https allowed)";
		return false;
	}

	outSecure = (uc.nScheme == INTERNET_SCHEME_HTTPS);
	outPort = uc.nPort;
	outHost.assign(uc.lpszHostName, uc.dwHostNameLength);

	// Path = url-path + extra-info (query/fragment). We keep query; callers may strip if desired.
	outPath.assign(uc.lpszUrlPath, uc.dwUrlPathLength);
	if (uc.dwExtraInfoLength > 0 && uc.lpszExtraInfo)
		outPath.append(uc.lpszExtraInfo, uc.dwExtraInfoLength);

	if (outPath.empty()) outPath = L"/";
	return true;
}

// --------------------------------------------------
// WinHTTP (no redirects; treat redirects as errors)
// --------------------------------------------------
struct WinHttpGetCtx
{
	WinHttpHandle session;
	WinHttpHandle connect;
	WinHttpHandle request;
};

static bool WinHttpOpenGet_NoRedirects(
	const std::string& urlUtf8,
	WinHttpGetCtx& out,
	const CancelToken& cancel,
	long connectTimeoutMs,
	long totalTimeoutMs,
	std::string* outErr,
	long* outHttp)
{
	if (outErr) outErr->clear();
	if (outHttp) *outHttp = 0;
	out = WinHttpGetCtx{};

	std::wstring host, path;
	INTERNET_PORT port = 0;
	bool secure = false;
	if (!CrackUrlWinHttp(urlUtf8, host, path, port, secure, outErr))
		return false;

	out.session = WinHttpHandle(WinHttpOpen(
		Utf8ToWide(AppConstants::kUserAgent).c_str(),
		WINHTTP_ACCESS_TYPE_DEFAULT_PROXY,
		WINHTTP_NO_PROXY_NAME,
		WINHTTP_NO_PROXY_BYPASS,
		0));
	if (!out.session)
	{
		if (outErr) *outErr = "WinHttpOpen failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	out.connect = WinHttpHandle(WinHttpConnect((HINTERNET)out.session.get(), host.c_str(), port, 0));
	if (!out.connect)
	{
		if (outErr) *outErr = "WinHttpConnect failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	DWORD flags = secure ? WINHTTP_FLAG_SECURE : 0;
	out.request = WinHttpHandle(WinHttpOpenRequest(
		(HINTERNET)out.connect.get(),
		L"GET",
		path.c_str(),
		nullptr,
		WINHTTP_NO_REFERER,
		WINHTTP_DEFAULT_ACCEPT_TYPES,
		flags));
	if (!out.request)
	{
		if (outErr) *outErr = "WinHttpOpenRequest failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	// Treat redirects as errors.
	DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_NEVER;
	(void)WinHttpSetOption((HINTERNET)out.request.get(), WINHTTP_OPTION_REDIRECT_POLICY, &redirectPolicy, sizeof(redirectPolicy));

	// Timeouts: resolve, connect, send, receive.
	const int resolveMs = (int)connectTimeoutMs;
	const int connectMs = (int)connectTimeoutMs;
	const int sendMs = (int)connectTimeoutMs;
	const int recvMs = (totalTimeoutMs <= 0) ? 0 : (int)totalTimeoutMs;
	(void)WinHttpSetTimeouts((HINTERNET)out.request.get(), resolveMs, connectMs, sendMs, recvMs);

	if (cancel.IsCanceled()) { if (outErr) *outErr = "Canceled"; return false; }

	if (!WinHttpSendRequest((HINTERNET)out.request.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0))
	{
		if (outErr) *outErr = "WinHttpSendRequest failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}
	if (!WinHttpReceiveResponse((HINTERNET)out.request.get(), nullptr))
	{
		if (outErr) *outErr = "WinHttpReceiveResponse failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	DWORD status = 0;
	DWORD statusSize = sizeof(status);
	(void)WinHttpQueryHeaders((HINTERNET)out.request.get(),
		WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER,
		WINHTTP_HEADER_NAME_BY_INDEX,
		&status,
		&statusSize,
		WINHTTP_NO_HEADER_INDEX);

	if (outHttp) *outHttp = (long)status;

	if (status >= 300 && status < 400)
	{
		if (outErr) *outErr = "HTTP redirect received; redirects are treated as errors";
		return false;
	}
	if (status < 200 || status >= 300)
	{
		if (outErr) *outErr = "HTTP status " + std::to_string(status);
		return false;
	}

	return true;
}

static bool WinHttpReadAllToString(
	HINTERNET hRequest,
	std::string& outBody,
	const CancelToken& cancel,
	size_t maxBytes,
	std::string* outErr)
{
	outBody.clear();
	if (outErr) outErr->clear();

	std::string body;
	body.reserve(64 * 1024);

	for (;;)
	{
		if (cancel.IsCanceled()) { if (outErr) *outErr = "Canceled"; return false; }

		DWORD avail = 0;
		if (!WinHttpQueryDataAvailable(hRequest, &avail))
		{
			if (outErr) *outErr = "WinHttpQueryDataAvailable failed (" + std::to_string(GetLastError()) + ")";
			return false;
		}
		if (avail == 0) break;

		// Enforce a hard cap for text downloads (manifest/version files).
		if (maxBytes > 0 && body.size() + (size_t)avail > maxBytes)
		{
			if (outErr) *outErr = "HTTP response too large";
			return false;
		}

		std::vector<char> buf(avail);
		DWORD read = 0;
		if (!WinHttpReadData(hRequest, buf.data(), avail, &read))
		{
			if (outErr) *outErr = "WinHttpReadData failed (" + std::to_string(GetLastError()) + ")";
			return false;
		}
		if (read == 0) break;

		body.append(buf.data(), buf.data() + read);
	}

	outBody.swap(body);
	return true;
}

static bool WinHttpGetToString_NoRedirects(
	const std::string& urlUtf8,
	std::string& outBody,
	const CancelToken& cancel,
	long connectTimeoutMs,
	long totalTimeoutMs,
	std::string* outErr,
	long* outHttp)
{
	WinHttpGetCtx ctx;
	if (!WinHttpOpenGet_NoRedirects(urlUtf8, ctx, cancel, connectTimeoutMs, totalTimeoutMs, outErr, outHttp))
		return false;

	// Text downloads should be small (manifest/version). Cap defensively.
	const size_t kMaxTextBytes = 4u * 1024u * 1024u; // 4 MB
	return WinHttpReadAllToString((HINTERNET)ctx.request.get(), outBody, cancel, kMaxTextBytes, outErr);
}


static Status WinHttpGetToString_NoRedirects_Status(
	const std::string& urlUtf8,
	std::string& outBody,
	const CancelToken& cancel,
	long connectTimeoutMs,
	long totalTimeoutMs,
	long* outHttp)
{
	std::string err;
	WinHttpGetCtx ctx;
	long http = 0;
	if (!WinHttpOpenGet_NoRedirects(urlUtf8, ctx, cancel, connectTimeoutMs, totalTimeoutMs, &err, &http)) {
		if (outHttp) *outHttp = http;
		return Status::FailMsg(Utf8ToWide(err));
	}

	// Text downloads should be small (manifest/version). Cap defensively.
	const size_t kMaxTextBytes = 4u * 1024u * 1024u; // 4 MB
	if (!WinHttpReadAllToString((HINTERNET)ctx.request.get(), outBody, cancel, kMaxTextBytes, &err)) {
		if (outHttp) *outHttp = http;
		return Status::FailMsg(Utf8ToWide(err));
	}

	if (outHttp) *outHttp = http;
	return Status::Ok();
}


static bool DownloadUrl(const std::string& url, std::string& out, const CancelToken& cancel, std::string* outErr = nullptr, long* outHttp = nullptr)
{
	const long connectMs = AppConstants::kManifestConnectTimeoutSec * 1000L;
	const long totalMs = AppConstants::kManifestTimeoutSec * 1000L;
	return WinHttpGetToString_NoRedirects(url, out, cancel, connectMs, totalMs, outErr, outHttp);
}
static std::string JoinUrl(const std::string& base, const std::string& path)
{
	if (base.empty()) return path;
	if (path.empty()) return base;
	const bool baseSlash = (!base.empty() && base.back() == '/');
	const bool pathSlash = (!path.empty() && path.front() == '/');
	if (baseSlash && pathSlash) return base + path.substr(1);
	if (!baseSlash && !pathSlash) return base + "/" + path;
	return base + path;
}
// --------------------------------------------------
// Download-to-file + hash-while-downloading (SHA-256)
// --------------------------------------------------
struct DlFileHashCtx
{
	FILE* f = nullptr;
	BCRYPT_ALG_HANDLE hAlg = nullptr;
	BCRYPT_HASH_HANDLE hHash = nullptr;
	std::vector<UCHAR> obj;
	std::vector<UCHAR> hash;
	bool ok = true;
};
static void DlCloseHash(DlFileHashCtx& ctx)
{
	if (ctx.hHash) { BCryptDestroyHash(ctx.hHash); ctx.hHash = nullptr; }
	if (ctx.hAlg) { BCryptCloseAlgorithmProvider(ctx.hAlg, 0); ctx.hAlg = nullptr; }
}
static bool DlInitSha256(DlFileHashCtx& ctx)
{
	NTSTATUS st = BCryptOpenAlgorithmProvider(&ctx.hAlg, BCRYPT_SHA256_ALGORITHM, nullptr, 0);
	if (st != 0) return false;
	DWORD objLen = 0, cbData = 0, hashLen = 0;
	st = BCryptGetProperty(ctx.hAlg, BCRYPT_OBJECT_LENGTH, (PUCHAR)&objLen, sizeof(objLen), &cbData, 0);
	if (st != 0) { DlCloseHash(ctx); return false; }
	st = BCryptGetProperty(ctx.hAlg, BCRYPT_HASH_LENGTH, (PUCHAR)&hashLen, sizeof(hashLen), &cbData, 0);
	if (st != 0) { DlCloseHash(ctx); return false; }
	ctx.obj.assign(objLen, 0);
	ctx.hash.assign(hashLen, 0);
	st = BCryptCreateHash(ctx.hAlg, &ctx.hHash, ctx.obj.data(), (ULONG)ctx.obj.size(), nullptr, 0, 0);
	if (st != 0) { DlCloseHash(ctx); return false; }
	return true;
}
static bool DlFinishSha256HexLower(DlFileHashCtx& ctx, std::string& outHexLower)
{
	outHexLower.clear();
	if (!ctx.hHash) return false;
	NTSTATUS st = BCryptFinishHash(ctx.hHash, ctx.hash.data(), (ULONG)ctx.hash.size(), 0);
	if (st != 0) return false;
	static const char* kHex = "0123456789abcdef";
	outHexLower.reserve(ctx.hash.size() * 2);
	for (UCHAR b : ctx.hash)
	{
		outHexLower.push_back(kHex[(b >> 4) & 0xF]);
		outHexLower.push_back(kHex[b & 0xF]);
	}
	return true;
}
static bool WinHttpDownloadToFileAndHash_NoRedirects(
	const std::string& urlUtf8,
	DlFileHashCtx& ctx,
	const CancelToken& cancel,
	long connectTimeoutMs,
	long totalTimeoutMs,
	std::string* outErr,
	long* outHttp)
{
	if (outErr) outErr->clear();
	if (outHttp) *outHttp = 0;

	WinHttpGetCtx http;
	if (!WinHttpOpenGet_NoRedirects(urlUtf8, http, cancel, connectTimeoutMs, totalTimeoutMs, outErr, outHttp))
		return false;

	for (;;)
	{
		if (cancel.IsCanceled()) { if (outErr) *outErr = "Canceled"; return false; }

		DWORD avail = 0;
		if (!WinHttpQueryDataAvailable((HINTERNET)http.request.get(), &avail))
		{
			if (outErr) *outErr = "WinHttpQueryDataAvailable failed (" + std::to_string(GetLastError()) + ")";
			return false;
		}
		if (avail == 0) break;

		std::vector<char> buf(avail);
		DWORD read = 0;
		if (!WinHttpReadData((HINTERNET)http.request.get(), buf.data(), avail, &read))
		{
			if (outErr) *outErr = "WinHttpReadData failed (" + std::to_string(GetLastError()) + ")";
			return false;
		}
		if (read == 0) break;

		if (ctx.f)
		{
			size_t wrote = fwrite(buf.data(), 1, read, ctx.f);
			if (wrote != read)
			{
				if (outErr) *outErr = "fwrite failed";
				return false;
			}
		}
		if (ctx.hHash)
		{
			NTSTATUS st = BCryptHashData(ctx.hHash, (PUCHAR)buf.data(), (ULONG)read, 0);
			if (st != 0)
			{
				if (outErr) *outErr = "BCryptHashData failed";
				return false;
			}
		}
	}

	return true;
}
static bool MoveReplace(const fs::path& from, const fs::path& to)
{
	std::wstring wFrom = from.wstring();
	std::wstring wTo = to.wstring();
	return MoveFileExW(wFrom.c_str(), wTo.c_str(), MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH) != 0;
}
static bool DownloadUrlToFileVerifySha256(
	const std::string& url,
	const fs::path& destFile,
	const std::string& expectedSha256HexLower,
	const CancelToken& cancel,
	std::string* outErr = nullptr,
	long* outHttp = nullptr)
{
	if (outErr) outErr->clear();
	if (outHttp) *outHttp = 0;
	std::error_code ec;
	fs::create_directories(destFile.parent_path(), ec);
	if (ec)
	{
		if (outErr) *outErr = "create_directories failed: " + ec.message();
		return false;
	}
	fs::path tmp = destFile;
	tmp += L".tmp";
	tmp += std::to_wstring(GetCurrentProcessId());
	tmp += L".";
	tmp += std::to_wstring(GetTickCount64());
	DlFileHashCtx ctx{};
	if (!DlInitSha256(ctx))
	{
		if (outErr) *outErr = "BCrypt SHA-256 init failed";
		return false;
	}
	_wfopen_s(&ctx.f, tmp.wstring().c_str(), L"wb");
	if (!ctx.f)
	{
		DlCloseHash(ctx);
		if (outErr) *outErr = "Failed to open temp file for writing";
		return false;
	}
	const long connectMs = AppConstants::kFileConnectTimeoutMs;
	const long totalMs = AppConstants::kFileTimeoutMs;
	std::string dlErr;
	long code = 0;
	bool ok = WinHttpDownloadToFileAndHash_NoRedirects(url, ctx, cancel, connectMs, totalMs, &dlErr, &code);
	if (outHttp) *outHttp = code;
	fflush(ctx.f);
	fclose(ctx.f);
	ctx.f = nullptr;
	if (!ok)
	{
		DlCloseHash(ctx);
		fs::remove(tmp, ec);
		if (outErr) *outErr = dlErr.empty() ? "Download failed" : dlErr;
		return false;
	}
	if (!ctx.ok)
	{
		DlCloseHash(ctx);
		fs::remove(tmp, ec);
		if (outErr) *outErr = "Write/hash failure during download";
		return false;
	}
	std::string gotHex;
	if (!DlFinishSha256HexLower(ctx, gotHex))
	{
		DlCloseHash(ctx);
		fs::remove(tmp, ec);
		if (outErr) *outErr = "BCryptFinishHash failed";
		return false;
	}
	DlCloseHash(ctx);
	if (!EqualIcaseAscii(gotHex, expectedSha256HexLower))
	{
		fs::remove(tmp, ec);
		if (outErr) *outErr = "SHA-256 mismatch after download";
		return false;
	}
	if (!MoveReplace(tmp, destFile))
	{
		fs::remove(tmp, ec);
		if (outErr) *outErr = "Failed to replace destination file";
		return false;
	}
	return true;
}
// --------------------------------------------------
// Helpers
// --------------------------------------------------
static bool is_space(char c) { return c == ' ' || c == '	' || c == '\r' || c == '\n'; }
static std::string ToLowerAscii(std::string s)
{
	for (char& c : s)
		if (c >= 'A' && c <= 'Z') c = char(c - 'A' + 'a');
	return s;
}
static std::string StripQueryAndFragment(std::string s)
{
	size_t q = s.find('?');
	size_t h = s.find('#');
	size_t cut = std::string::npos;
	if (q != std::string::npos) cut = q;
	if (h != std::string::npos) cut = (cut == std::string::npos) ? h : (cut < h ? cut : h);
	if (cut != std::string::npos) s.resize(cut);
	return s;
}
static int RemoveEmptyDirsBottomUp(const fs::path& root, bool removeRoot = false)
{
	std::error_code ec;
	int removedDirs = 0;
	if (!fs::exists(root, ec) || !fs::is_directory(root, ec))
		return 0;
	std::vector<fs::path> dirs;
	for (auto it = fs::recursive_directory_iterator(root, fs::directory_options::skip_permission_denied, ec);
		it != fs::recursive_directory_iterator(); it.increment(ec))
	{
		if (ec) { ec.clear(); continue; }
		if (it->is_directory(ec) && !ec) dirs.push_back(it->path());
		else ec.clear();
	}
	std::sort(dirs.begin(), dirs.end(), [](const fs::path& a, const fs::path& b)
		{
			return a.native().size() > b.native().size();
		});
	for (const auto& d : dirs)
	{
		if (fs::is_directory(d, ec) && fs::is_empty(d, ec))
		{
			if (fs::remove(d, ec))
			{
				Log("  REMOVED EMPTY DIR: " + PathToUtf8(d) + "\r\n");
				++removedDirs;
			}
			else
			{
				ec.clear();
			}
		}
		else
		{
			ec.clear();
		}
	}
	if (removeRoot)
	{
		if (fs::is_directory(root, ec) && fs::is_empty(root, ec))
		{
			if (fs::remove(root, ec))
			{
				Log("  REMOVED EMPTY DIR: " + PathToUtf8(root) + "\r\n");
				++removedDirs;
			}
		}
		else
		{
			ec.clear();
		}
	}
	return removedDirs;
}
// --------------------------------------------------
// Ensure prefs\ClientPrefs_Common.def has correct mapPath
// --------------------------------------------------
static void UpdateClientPrefsMapPath(const SyncConfig& cfg, const CancelToken& cancel, const std::string& desiredValue, const char* context)
{
	if (cancel.IsCanceled()) return;
	LogSeparator();
	try
	{
		const fs::path prefsFile = cfg.localBase / "prefs" / "ClientPrefs_Common.def";
		if (!fs::exists(prefsFile))
		{
			Log(std::string("\\prefs\\ClientPrefs_Common.def check: file not found: ") + PathToUtf8(prefsFile) + "\r\n");
			return;
		}
		std::ifstream in(prefsFile, std::ios::binary);
		if (!in)
		{
			Log(std::string("\\prefs\\ClientPrefs_Common.def check: failed to open for read: ") + PathToUtf8(prefsFile) + "\r\n");
			return;
		}
		std::string content((std::istreambuf_iterator<char>(in)), std::istreambuf_iterator<char>());
		in.close();
		const std::string needle = "string mapPath";
		size_t pos = content.find(needle);
		if (pos == std::string::npos)
		{
			Log(std::string("\\prefs\\ClientPrefs_Common.def check: 'string mapPath' not found in ") + PathToUtf8(prefsFile) + "\r\n");
			return;
		}
		size_t lineStart = content.rfind('\n', pos);
		if (lineStart == std::string::npos) lineStart = 0; else lineStart += 1;
		size_t lineEnd = content.find('\n', pos);
		if (lineEnd == std::string::npos) lineEnd = content.size();
		std::string line = content.substr(lineStart, lineEnd - lineStart);
		size_t q1 = line.find('"');
		if (q1 == std::string::npos)
		{
			Log(std::string("\\prefs\\ClientPrefs_Common.def check: mapPath line has no opening quote in ") + PathToUtf8(prefsFile) + "\r\n");
			return;
		}
		size_t q2 = line.find('"', q1 + 1);
		if (q2 == std::string::npos)
		{
			Log(std::string("\\prefs\\ClientPrefs_Common.def check: mapPath line has no closing quote in ") + PathToUtf8(prefsFile) + "\r\n");
			return;
		}
		std::string currentValue = line.substr(q1 + 1, q2 - (q1 + 1));
		if (currentValue == desiredValue)
		{
			Log(std::string("Checking 'string mapPath' in: \\prefs\\ClientPrefs_Common.def  (") + (context ? context : "") + "):\r\n  It's already correct -> " + desiredValue + "\r\n");
			return;
		}
		size_t firstNonWs = line.find_first_not_of(" 	");
		std::string indent = (firstNonWs == std::string::npos) ? "" : line.substr(0, firstNonWs);
		std::string newLine = indent + "string mapPath = \"" + desiredValue + "\"";
		if (!line.empty() && line.back() == '\r')
			newLine.push_back('\r');
		content.replace(lineStart, lineEnd - lineStart, newLine);
		fs::path tmp = prefsFile;
		tmp += L".tmp";
		std::ofstream out(tmp, std::ios::binary | std::ios::trunc);
		if (!out)
		{
			Log(std::string("\\prefs\\ClientPrefs_Common.def check: failed to open temp for write: ") + PathToUtf8(tmp) + "\r\n");
			return;
		}
		out.write(content.data(), (std::streamsize)content.size());
		out.close();
		std::error_code ec;
		fs::rename(tmp, prefsFile, ec);
		if (ec)
		{
			std::error_code ec2;
			fs::remove(prefsFile, ec2);
			fs::rename(tmp, prefsFile, ec2);
			if (ec2)
			{
				Log(std::string("\\prefs\\ClientPrefs_Common.def check: failed to replace file: ") + PathToUtf8(prefsFile) + "\r\n");
				std::error_code ec3;
				fs::remove(tmp, ec3);
				return;
			}
		}
		Log(std::string("Checking 'string mapPath' in: \\prefs\\ClientPrefs_Common.def  (") + (context ? context : "") + "):\r\n  It's Incorrect -> Updating...\r\n  Old: " + currentValue + "\r\n  New: " + desiredValue + "\r\n");
	}
	catch (...)
	{
		Log("\\prefs\\ClientPrefs_Common.def check: exception while updating ClientPrefs_Common.def\r\n");
	}
}

static void EnsureClientPrefsMapPath(const SyncConfig& cfg, const CancelToken& cancel)
{
	// Sync mode: point the game at the override mappack content.
	UpdateClientPrefsMapPath(cfg, cancel, "resources_override/mappack/resources/interface/maps", "Verify/Set to MapPack Install Path");
}

static void EnsureClientPrefsMapPath_Remove(const SyncConfig& cfg, const CancelToken& cancel)
{
	// Remove mode: restore mapPath to the built-in (non-override) mappack path.
	UpdateClientPrefsMapPath(cfg, cancel, "resources/mappack/resources/interface/maps", "Verify/Set to Normal/Vanilla Install Path");
}

// --------------------------------------------------
// Old-manifest cleanup (removes files created by older versions of the tool)
// Downloads mappack_manifest_old.json and deletes any listed files that still
// exist under: <Istaria>\resources_override\<path>
// --------------------------------------------------

static bool ParseManifestOldPaths(const std::string& jsonText, std::vector<std::string>& outPaths)
{
	outPaths.clear();
	const std::string& s = jsonText;
	size_t i = 0;

	const size_t posFiles = s.find("\"files\"");
	if (posFiles == std::string::npos) return false;

	i = posFiles + strlen("\"files\"");
	SkipWs(s, i);
	if (i >= s.size() || s[i] != ':') return false;
	++i;
	SkipWs(s, i);
	if (i >= s.size() || s[i] != '[') return false;
	++i;

	while (i < s.size())
	{
		SkipWs(s, i);
		if (i >= s.size()) return false;
		if (s[i] == ']') { ++i; break; }
		if (s[i] == ',') { ++i; continue; }
		if (s[i] != '{') return false;
		++i;

		std::string path;
		while (i < s.size())
		{
			SkipWs(s, i);
			if (i >= s.size()) return false;
			if (s[i] == '}') { ++i; break; }
			if (s[i] == ',') { ++i; continue; }

			std::string key;
			if (!ReadJsonStringValue(s, i, key)) return false;
			SkipWs(s, i);
			if (i >= s.size() || s[i] != ':') return false;
			++i;
			SkipWs(s, i);

			std::string val;
			if (!ReadJsonStringValue(s, i, val)) return false;

			if (key == "path") path = val;
		}

		if (!path.empty())
			outPaths.push_back(path);
	}

	return !outPaths.empty();
}

static void RemoveOldManifestListedFiles(const SyncConfig& cfg, const CancelToken& cancel)
{
	if (cancel.IsCanceled()) return;
	LogSeparator();
	Log("Downloading manifest for old MapPack 4.0 and earlier versions... ");

	const std::string url = JoinUrl(cfg.remoteHost, kManifestOldPath);
	std::string jsonText;
	std::string dlErr;
	long http = 0;
	if (!DownloadUrl(url, jsonText, cancel, &dlErr, &http))
	{
		Log("  (skipped) Could not download mappack_manifest_old.json (" + std::to_string(http) + "): " + dlErr + "\r\n");
		return;
	}

	std::vector<std::string> relPaths;
	if (!ParseManifestOldPaths(jsonText, relPaths))
	{
		Log("  (skipped) Could not parse mappack_manifest_old.json\r\n");
		return;
	}

	Log("Success!\r\n\r\n");
	LogSeparator();

	Log("Parsing and removing files from old MapPack 4.0...\r\n");
	const fs::path oldRoot = cfg.localBase / "resources_override";

	int deleted = 0;
	int failed = 0;
	for (const auto& rp : relPaths)
	{
		if (cancel.IsCanceled()) return;

		// Old files live directly under resources_override\resources\...
		std::error_code ec;
		fs::path local = oldRoot / fs::path(Utf8ToWide(rp));
		if (!fs::exists(local, ec) || ec) continue;
		ec.clear();
		if (!fs::is_regular_file(local, ec) || ec) continue;

		ec.clear();
		fs::remove(local, ec);
		if (!ec)
		{
			++deleted;
			Log("  DELETED (old): " + rp + "\r\n");
		}
		else
		{
			++failed;
			Log("ERROR deleting old file: " + rp + " (" + std::string(ec.message()) + ")\r\n");
		}
	}

	// Optional: remove now-empty dirs, but keep this scope tight to avoid touching anything else.
	const fs::path oldMapsRoot = oldRoot / "resources" / "interface" / "maps";
	const fs::path oldTexturesRoot = oldRoot / "resources" / "interface" / "themes" / "default" / "textures";

	if (deleted > 0 || failed > 0)
		Log("\r\n"); //Add blank line
	Log("MapPack 4.0 Deleted Files Summary:\r\n");
	Log("  Deletions: " + std::to_string(deleted) + "\r\n");
	Log("  Failed deletions: " + std::to_string(failed) + "\r\n");

	if (!cancel.IsCanceled())
	{
		LogSeparator();
		Log("Removing empty directories from MapPack 4.0 (maps/textures folders only)...\r\n");
		const int removedDirs = RemoveEmptyDirsBottomUp(oldMapsRoot);
		if (removedDirs == 0)
			Log("  No empty sub-directories found; Nothing  to delete.\r\n");
		else
		{
			Log("\r\nEmpty Subdirectories (old) Removal Summary:\r\n"); //zz
			Log("  Deletions: " + std::to_string(removedDirs) + "\r\n"); //zz
		}
	}
}

// which extensions to sync
static bool IsSyncedExt(const std::string& lowerExt)
{
	return (lowerExt == ".def" || lowerExt == ".png");
}
// --------------------------------------------------
// Folder picker
// --------------------------------------------------
static bool BrowseForFolder(HWND hwnd, std::wstring& outPath)
{
	BROWSEINFO bi{};
	bi.hwndOwner = hwnd;
	bi.lpszTitle = L"Select Istaria install folder (must contain istaria.exe)";
	bi.ulFlags = BIF_RETURNONLYFSDIRS | BIF_USENEWUI;
	PIDLIST_ABSOLUTE pidl = SHBrowseForFolder(&bi);
	if (!pidl) return false;
	std::vector<wchar_t> pathBuf(32768, L'\0');
	if (!SHGetPathFromIDListEx(pidl, pathBuf.data(), (UINT)pathBuf.size(), 0))
	{
		wchar_t path[MAX_PATH]{};
		if (!SHGetPathFromIDListW(pidl, path))
		{
			CoTaskMemFree(pidl);
			return false;
		}
		outPath = path;
		CoTaskMemFree(pidl);
		return true;
	}
	CoTaskMemFree(pidl);
	outPath = pathBuf.data();
	return true;
}
// --------------------------------------------------
// Manifest / Sync types
// --------------------------------------------------
struct SyncCounters
{
	size_t deleted = 0;
	size_t downloaded = 0;
	size_t updated = 0;
	size_t unchanged = 0;
	size_t failed = 0;
};
struct ManifestData
{
	std::vector<ManifestEntry> workList;
	std::unordered_set<std::string> manifestRelSet;
};
static std::string MakeFileUrlFromRemoteHost(const std::string& remotePath)
{
	// kRemoteHost has no trailing slash. remotePath in the manifest is expected to be relative,
	// but we defensively handle a leading '/' as well.
	std::string url = kRemoteHost;
	if (!url.empty() && url.back() != '/') url.push_back('/');
	if (!remotePath.empty() && remotePath.front() == '/')
		return url + remotePath.substr(1);
	return url + remotePath;
}
// --------------------------------------------------
// Manifest download + parsing
// --------------------------------------------------
static bool DownloadAndParseManifest(const SyncConfig& cfg, ManifestData& out, std::string& outErr, const CancelToken& cancel)
{
	if (cancel.IsCanceled()) { outErr = "canceled"; return false; }
	out = ManifestData{};
	outErr.clear();
	Log("Downloading MapPack 5.0 manifest... ");

	PostProgressMarqueeOn();
	PostProgressTextW(L"Downloading manifest...");
	std::string manifestText;
	std::string dlErr; long http = 0;
	if (!DownloadUrl(cfg.manifestUrl, manifestText, cancel, &dlErr, &http))
	{
		PostProgressMarqueeOff();
		outErr = "Manifest download failed (HTTP " + std::to_string(http) + "): " + dlErr;
		return false;
	}
	std::vector<ManifestRawEntry> rawFiles;
	std::string parseErr;
	if (!ParseManifestRaw(manifestText, rawFiles, &parseErr))
	{
		PostProgressMarqueeOff();
		outErr = "Manifest parse failed: " + (parseErr.empty() ? std::string("unknown error") : parseErr);
		return false;
	}
	PostProgressMarqueeOff();

	std::string valErr;
	if (!ValidateAndNormalizeManifest(rawFiles, out.workList, out.manifestRelSet, &valErr))
	{
		outErr = "Manifest validation failed: " + (valErr.empty() ? std::string("unknown error") : valErr);
		return false;
	}
	return true;
}
static void DeleteLocalFilesNotInManifest(const SyncConfig& cfg,
	const std::unordered_set<std::string>& manifestRelSet,
	const CancelToken& cancel)
{
	if (cancel.IsCanceled()) return;
	size_t failedDeletes = 0;
	size_t filesDeleted = 0;

	LogSeparator();
	Log("Clean-up MapPack 5.0: Searching local files that exist but are not in the manifest; Those need deleted...\r\n");
	std::vector<fs::path> localFiles;
	// Scan entire sync root (resources_override/mappack/)
	if (fs::exists(cfg.localSyncRoot))
	{
		for (auto& entry : fs::recursive_directory_iterator(cfg.localSyncRoot))
		{
			if (cancel.IsCanceled()) return;
			if (!entry.is_regular_file()) continue;
			localFiles.push_back(entry.path());
		}
	}
	else
	{
		Log("NOTE: sync root folder not found; nothing to delete.\r\n");
		Log("Expected: " + PathToUtf8(cfg.localSyncRoot) + "\r\n");
		return;
	}
	for (const auto& fullPath : localFiles)
	{
		if (CheckAndHandleCancel(cancel, "INFO: Canceling... stopping deletions.\r\n"))
			return;
		fs::path relPathFs = fs::relative(fullPath, cfg.localSyncRoot);
		std::string rel = NormalizeManifestRel(relPathFs.generic_string());
		if (manifestRelSet.find(rel) == manifestRelSet.end())
		{
			std::error_code ec;
			if (fs::remove(fullPath, ec))
			{
				Log("  DELETED: " + rel + "\r\n");
				++filesDeleted;
			}
			else
			{
				Log("  FAILED DELETE: " + rel + " (" + ec.message() + ")\r\n");
				++failedDeletes;
			}
		}
	}
	if (filesDeleted == 0 && failedDeletes == 0)
		Log("  No files found that needs to be deleted!\r\n");

	Log("\r\nFile Delete Summary:\r\n");
	Log("  Deletions:  " + std::to_string(filesDeleted) + "\r\n");
	Log("  Failed deletions:  " + std::to_string(failedDeletes) + "\r\n");
}
static void DownloadAndUpdateFiles(const SyncConfig& cfg, const ManifestData& md, SyncCounters& ioCounts, const CancelToken& cancel)
{
	if (cancel.IsCanceled()) return;
	Log("Parsing MapPack 5.0 manifest: Searching for any local files that are missing or has changed (Needs updated)...\r\n");
	bool anyChanged = false;
	PostProgressInit(md.workList.size());
	for (size_t i = 0; i < md.workList.size(); ++i)
	{
		if (CheckAndHandleCancel(cancel, "INFO: Canceled during parsing.\r\n"))
			break;
		std::string rel = md.workList[i].relPath;
		if (StartsWith(rel, "mappack/"))
			rel = rel.substr(strlen("mappack/"));
		const std::string& remotePath = md.workList[i].remotePath;
		const std::string& expectedHash = md.workList[i].sha256;
		ScopeExit progressGuard{ [i]() {
			PostProgressSet(i + 1);
		} };
		PostProgressTextW(
			L"File " + std::to_wstring(i + 1) + L"/" + std::to_wstring(md.workList.size()) +
			L": " + Utf8ToWide(rel));
		fs::path localFile = MakeDestPath(cfg.localBase, rel);
		if (fs::exists(localFile))
		{
			std::string localHash;
			if (!Sha256FileHexLower(localFile, localHash))
			{
				Log("FAILED HASH (local): " + rel + "\r\n");
				++ioCounts.failed;
				continue;
			}
			if (EqualIcaseAscii(localHash, expectedHash))
			{
				++ioCounts.unchanged;
				continue;
			}
		}
		std::string fileUrl = MakeFileUrlFromRemoteHost(remotePath);
		std::string dlErr2; long http2 = 0;
		bool existed = fs::exists(localFile);
		if (!DownloadUrlToFileVerifySha256(fileUrl, localFile, expectedHash, cancel, &dlErr2, &http2))
		{
			Log("FAILED DOWNLOAD: " + rel + " (HTTP " + std::to_string(http2) + ") " + dlErr2 + "\r\n");
			++ioCounts.failed;
			continue;
		}
		if (!existed)
		{
			Log("  DOWNLOADED: resources_override/mappack/" + rel + "\r\n");
			++ioCounts.downloaded; anyChanged = true;
		}
		else
		{
			Log("UPDATED: resources_override/mappack/" + rel + "\r\n");
			++ioCounts.updated; anyChanged = true;
		}
	}

	if (cancel.IsCanceled()) return;
	if (!anyChanged)
		Log("  No missing or changed files found. Your files are in sync with the manifest!\r\n");
	PostProgressTextW(L"Sync complete.");
}
static void LogSummaryAndCleanup(const SyncConfig& cfg, const SyncCounters& c, const CancelToken& cancel)
{
	if (cancel.IsCanceled()) return;

	Log("\r\n  Sync Summary:\r\n");
	Log("    Downloaded (missing):  " + std::to_string(c.downloaded) + "\r\n");
	Log("    Updated (different):  " + std::to_string(c.updated) + "\r\n");
	Log("    Unchanged (same):  " + std::to_string(c.unchanged) + "\r\n");
	Log("    Failed Downloads/Updates:  " + std::to_string(c.failed) + "\r\n");
}
static void RunSync(const SyncConfig& cfg, const CancelToken& cancel)
{
	ManifestData md;
	std::string err;
	if (!DownloadAndParseManifest(cfg, md, err, cancel))
	{
		Log("ERROR: " + err + "\r\n");
		Log("Aborting sync. No local deletes/cleanup will be performed.\r\n");
		PostProgressTextW(L"Aborted (manifest error).");
		return;
	}
	// Completes the earlier "Downloading manifest..." log line.
	Log("Success!\r\n");
	Log("  Manifest file count: " + std::to_string(md.workList.size()) + "\r\n");
	if (CheckAndHandleCancel(cancel, "INFO: Canceled after manifest.\r\n"))
		return;
	Log("\r\nSyncing MapPack 5.0 root folder:  " + PathToUtf8(cfg.localSyncRoot) + "\r\n");
	SyncCounters counts;
	LogSeparator();

	if (CheckAndHandleCancel(cancel, "INFO: Canceled before downloads.\r\n"))
		return;

	DownloadAndUpdateFiles(cfg, md, counts, cancel);
	if (CheckAndHandleCancel(cancel, "INFO: Canceled before deletions.\r\n"))
		return;

	LogSummaryAndCleanup(cfg, counts, cancel);

	DeleteLocalFilesNotInManifest(cfg, md.manifestRelSet, cancel);
	if (CheckAndHandleCancel(cancel, "INFO: Canceled before Log Summary & Cleanup.\r\n"))
		return;

	LogSeparator();
	Log("Clean-up/Remove MapPack 5.0 empty sub-directories...\r\n");
	const int removedDirs = RemoveEmptyDirsBottomUp(cfg.localSyncRoot, true);

	if (removedDirs == 0)
		Log("  No empty sub-directories found; Nothing  to delete.\r\n");
	else
		Log("\r\nEmpty Directory Removal Summary:\r\n");
	//	Log("  Directories removed: " + std::to_string(removedDirs) + "\r\n");

	if (IsCanceledNoNotify(cancel)) return;
	RemoveOldManifestListedFiles(cfg, cancel);
	if (IsCanceledNoNotify(cancel)) return;
	EnsureClientPrefsMapPath(cfg, cancel);
	if (IsCanceledNoNotify(cancel)) return;
	LogSeparator();
	Log("Sync complete.");
}


// --------------------------------------------------
// Remove (uninstall) MapPack files
// - Deletes files listed in mappack_manifest.json (under resources_override\mappack\...)
// - Deletes legacy files listed in mappack_manifest_old.json (under resources_override\...)
// - Restores prefs\ClientPrefs_Common.def mapPath to: resources/mappack/resources/interface/maps
// --------------------------------------------------
static void RemoveMapPackFiles(const SyncConfig& cfg, const CancelToken& cancel)
{
	ManifestData md;
	std::string err;
	if (!DownloadAndParseManifest(cfg, md, err, cancel))
	{
		Log("ERROR: " + err + "\r\n");
		Log("Aborting remove. No local deletes/cleanup will be performed.\r\n");
		PostProgressTextW(L"Aborted (manifest error).");
		return;
	}

	// Completes the earlier "Downloading MapPack 5.0 manifest..." log line.
	Log("Success!\r\n");
	Log("  Manifest file count: " + std::to_string(md.workList.size()) + "\r\n");
	if (CheckAndHandleCancel(cancel, "INFO: Canceled after manifest.\r\n"))
		return;

	LogSeparator();
	Log("Parsing and removing files from MapPack 5.0 manifest...\r\n");
	PostProgressInit(md.workList.size());

	size_t deleted = 0;
	size_t missing = 0;
	size_t failed = 0;

	for (size_t i = 0; i < md.workList.size(); ++i)
	{
		if (CheckAndHandleCancel(cancel, "INFO: Canceled during remove.\r\n"))
			break;

		const std::string& rel = md.workList[i].relPath;
		ScopeExit progressGuard{ [i]() { PostProgressSet(i + 1); } };
		PostProgressTextW(L"Removing " + std::to_wstring(i + 1) + L"/" + std::to_wstring(md.workList.size()) + L": " + Utf8ToWide(rel));

		fs::path localFile = MakeDestPath(cfg.localBase, rel);
		std::error_code ec;
		if (!fs::exists(localFile, ec) || ec)
		{
			++missing;
			continue;
		}
		ec.clear();
		if (fs::remove(localFile, ec) && !ec)
		{
			++deleted;
			Log("  DELETED: resources_override/mappack/" + rel + "\r\n");
		}
		else
		{
			++failed;
			Log("  FAILED DELETE: resources_override/mappack/" + rel + " (" + (ec ? ec.message() : std::string("unknown")) + ")\r\n");
		}
	}

	if (deleted > 0 || failed > 0)
		Log("\r\n"); //Add blank line

	Log("MapPack 5.0 Deleted Files Summary:\r\n");
	Log("  Deletions:  " + std::to_string(deleted) + "\r\n");
	Log("  File doesn't exist (already removed):  " + std::to_string(missing) + "\r\n");
	Log("  Failed deletions:  " + std::to_string(failed) + "\r\n");

	if (IsCanceledNoNotify(cancel)) return;
	LogSeparator();
	Log("Removing empty sub-directories from MapPack 5.0 (sync root)...\r\n");

	const int removedDirs = RemoveEmptyDirsBottomUp(cfg.localSyncRoot, true);
	if (removedDirs == 0)
		Log("  No empty sub-directories found; Nothing  to delete.\r\n");
	else
	{
		Log("\r\nEmpty Subdirectories (MapPack 5.0) Removal Summary:\r\n"); //zz
		Log("  Deletions: " + std::to_string(removedDirs) + "\r\n"); //zz
	}
	if (IsCanceledNoNotify(cancel)) return;
	RemoveOldManifestListedFiles(cfg, cancel);

	if (IsCanceledNoNotify(cancel)) return;
	EnsureClientPrefsMapPath_Remove(cfg, cancel);

	if (IsCanceledNoNotify(cancel)) return;
	LogSeparator();
	Log("All known versions of MapPack has been Removed/Uninstalled.\r\n");
	PostProgressTextW(L"Remove/Uninstall complete.");
}
static unsigned __stdcall WorkerThreadProc(void*)
{
	if (!g_state) return 0;
	std::wstring folderWs;
	int len = GetWindowTextLengthW(g_state->hFolderEdit);
	if (len > 0)
	{
		std::vector<wchar_t> buf((size_t)len + 1, L'\0');
		GetWindowTextW(g_state->hFolderEdit, buf.data(), (int)buf.size());
		folderWs.assign(buf.data());
	}
	TrimInPlace(folderWs);
	StripSurroundingQuotes(folderWs);
	PreflightResult pf = ValidateFolderSelection(folderWs);
	if (!pf.ok)
	{
		for (const auto& line : pf.errors)
			Log(line);
		g_state->isRunning.store(false);
		{ UiEvent* ev = new UiEvent(); ev->kind = UiEventKind::WorkerDone; PostUiEvent(ev); }
		return 0;
	}

	// Persist last used folder even when user manually types it (INI created on first write).
	if (!folderWs.empty())
		IniWriteLastFolder(folderWs);

	SyncConfig cfg;
	cfg.remoteHost = kRemoteHost;
	cfg.remoteRootPath = kRemoteRootPath;
	cfg.manifestUrl = JoinUrl(kRemoteHost, kManifestPath);
	cfg.localBase = pf.localBase;
	cfg.localSyncRoot = pf.localSyncRoot;
	CancelToken cancel{ &g_state->cancelRequested };
	RunSync(cfg, cancel);
	g_state->isRunning.store(false);
	{ UiEvent* ev = new UiEvent(); ev->kind = UiEventKind::WorkerDone; PostUiEvent(ev); }
	return 0;
}
static void StartSyncIfNotRunning()
{
	if (g_state->isRunning.exchange(true))
	{
		Log("INFO: Sync already running.");
		return;
	}
	g_state->cancelRequested.store(false);
	g_state->pendingExitAfterWorker = false;
	g_state->progressFrozenOnCancel = false;
	g_state->progressTotal = 100;
	g_state->progressPos = 0;
	SetUiForWorkerRunning(g_state, true);
	unsigned int tid = 0;
	unique_handle worker(reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, &WorkerThreadProc, nullptr, 0, &tid)));
	g_state->hWorkerThread = worker.release();
	if (!g_state->hWorkerThread)
	{
		g_state->isRunning.store(false);
		SetUiForWorkerRunning(g_state, false);
		Log("ERROR: Failed to start worker thread.");
	}
}

static unsigned __stdcall WorkerThreadProcRemove(void*)
{
	if (!g_state) return 0;
	std::wstring folderWs;
	int len = GetWindowTextLengthW(g_state->hFolderEdit);
	if (len > 0)
	{
		std::vector<wchar_t> buf((size_t)len + 1, L'\0');
		GetWindowTextW(g_state->hFolderEdit, buf.data(), (int)buf.size());
		folderWs.assign(buf.data());
	}
	TrimInPlace(folderWs);
	StripSurroundingQuotes(folderWs);

	PreflightResult pf = ValidateFolderSelection(folderWs);
	if (!pf.ok)
	{
		for (const auto& line : pf.errors)
			Log(line);
		g_state->isRunning.store(false);
		{ UiEvent* ev = new UiEvent(); ev->kind = UiEventKind::WorkerDone; PostUiEvent(ev); }
		return 0;
	}

	SyncConfig cfg;
	cfg.remoteHost = kRemoteHost;
	cfg.remoteRootPath = kRemoteRootPath;
	cfg.manifestUrl = JoinUrl(kRemoteHost, kManifestPath);
	cfg.localBase = pf.localBase;
	cfg.localSyncRoot = pf.localSyncRoot;

	CancelToken cancel{ &g_state->cancelRequested };
	RemoveMapPackFiles(cfg, cancel);

	g_state->isRunning.store(false);
	{ UiEvent* ev = new UiEvent(); ev->kind = UiEventKind::WorkerDone; PostUiEvent(ev); }
	return 0;
}

static void StartRemoveIfNotRunning()
{
	if (g_state->isRunning.exchange(true))
	{
		Log("INFO: Worker already running.");
		return;
	}
	g_state->cancelRequested.store(false);
	g_state->pendingExitAfterWorker = false;
	g_state->progressFrozenOnCancel = false;
	g_state->progressTotal = 100;
	g_state->progressPos = 0;
	SetUiForWorkerRunning(g_state, true);
	unsigned int tid = 0;
	unique_handle worker(reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, &WorkerThreadProcRemove, nullptr, 0, &tid)));
	g_state->hWorkerThread = worker.release();
	if (!g_state->hWorkerThread)
	{
		g_state->isRunning.store(false);
		SetUiForWorkerRunning(g_state, false);
		Log("ERROR: Failed to start remove worker thread.");
	}
}

// --------------------------------------------------
// Check if a process is running by executable name.
// Verify if istaria.exe is running. If true abort and give error.
// If we need to update ClientPrefs_Common.def then it would get undone when user exits Istaria.
// --------------------------------------------------
static bool IsProcessRunningByName(const wchar_t* exeName)
{
	if (!exeName || !*exeName) return false;
	HANDLE rawSnap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
	if (rawSnap == INVALID_HANDLE_VALUE) return false;
	unique_handle snap(rawSnap);
	PROCESSENTRY32W pe;
	pe.dwSize = sizeof(pe);
	if (Process32FirstW(snap.get(), &pe))
	{
		do
		{
			if (_wcsicmp(pe.szExeFile, exeName) == 0)
			{
				return true;
			}
		} while (Process32NextW(snap.get(), &pe));
	}
	return false;
}

// --------------------------------------------------
// Self-update (WinHTTP; no external libs; uses a small helper exe in the same folder)
// --------------------------------------------------
static fs::path GetThisExePath()
{
	wchar_t buf[MAX_PATH] = {};
	DWORD n = GetModuleFileNameW(nullptr, buf, MAX_PATH);
	if (n == 0 || n >= MAX_PATH) return {};
	return fs::path(buf);
}

static void TryDeleteUpdaterExeBestEffort(const fs::path& exeDir)
{
	// The updater cannot delete itself while running. Best-effort cleanup is performed
	// by the main app on startup after an update/relaunch.
	const fs::path candidates[] = {
		exeDir / L"MapPackSyncTool_Updater.exe",
		// Legacy name (in case an older build left it behind).
		exeDir / L"MapPackSyncTool_UpdateHelper.exe"
	};

	for (const auto& p : candidates)
	{
		std::error_code ec;
		if (!fs::exists(p, ec)) continue;

		// The updater may still be exiting; retry briefly.
		for (int i = 0; i < 15; ++i)
		{
			if (DeleteFileW(p.c_str()))
				break;
			Sleep(200); // ~3 seconds total
		}
	}
}


static bool AcquireSingleInstanceMutex()
{
	// Single-instance guard: prevents multiple copies from running (based on a named local mutex)
		// within the same user logon session.
	const wchar_t* kMutexName = L"Local\\MapPackSyncTool_SingleInstance";

	g_hSingleInstanceMutex = CreateMutexW(nullptr, TRUE, kMutexName);
	if (!g_hSingleInstanceMutex)
		return true; // fail open

	if (GetLastError() == ERROR_ALREADY_EXISTS)
		return false;

	return true;
}

static void ReleaseSingleInstanceMutex()
{
	if (g_hSingleInstanceMutex)
	{
		ReleaseMutex(g_hSingleInstanceMutex);
		CloseHandle(g_hSingleInstanceMutex);
		g_hSingleInstanceMutex = nullptr;
	}
}

static void ActivateExistingInstance()
{
	// Prefer class name; fall back to title.
	HWND hwnd = FindWindowW(L"DEF_SYNC_GUI", nullptr);
	if (!hwnd)
		hwnd = FindWindowW(nullptr, L"MapPack Sync Tool");

	if (!hwnd)
		return;

	if (IsIconic(hwnd))
		ShowWindow(hwnd, SW_RESTORE);
	else
		ShowWindow(hwnd, SW_SHOW);

	// Try to bring to foreground (may be blocked by foreground lock in some cases).
	SetForegroundWindow(hwnd);
}

static bool WinHttpDownloadUrlToFile(const wchar_t* url, const fs::path& outFile, std::wstring& outErr)
{
	outErr.clear();
	if (!url || !*url) { outErr = L"Empty URL"; return false; }

	URL_COMPONENTS uc{};
	uc.dwStructSize = sizeof(uc);
	uc.dwSchemeLength = (DWORD)-1;
	uc.dwHostNameLength = (DWORD)-1;
	uc.dwUrlPathLength = (DWORD)-1;
	uc.dwExtraInfoLength = (DWORD)-1;
	if (!WinHttpCrackUrl(url, 0, 0, &uc))
	{
		outErr = L"WinHttpCrackUrl failed";
		return false;
	}

	// WinHttpCrackUrl may return null pointers for some components; validate before constructing strings.
	if (!uc.lpszHostName || uc.dwHostNameLength == 0) { outErr = L"Invalid URL host"; return false; }
	if (!uc.lpszUrlPath || uc.dwUrlPathLength == 0) { outErr = L"Invalid URL path"; return false; }

	std::wstring host(uc.lpszHostName, uc.dwHostNameLength);
	std::wstring path;
	path.assign(uc.lpszUrlPath, uc.dwUrlPathLength);
	if (uc.lpszExtraInfo && uc.dwExtraInfoLength)
		path.append(uc.lpszExtraInfo, uc.dwExtraInfoLength);

	HINTERNET hSession = WinHttpOpen(AppConstants::kUserAgentW, WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
	if (!hSession) { outErr = L"WinHttpOpen failed"; return false; }
	ScopeExit closeSession{ [&]() { WinHttpCloseHandle(hSession); } };

	HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(), uc.nPort, 0);
	if (!hConnect) { outErr = L"WinHttpConnect failed"; return false; }
	ScopeExit closeConnect{ [&]() { WinHttpCloseHandle(hConnect); } };

	DWORD flags = 0;
	if (uc.nScheme == INTERNET_SCHEME_HTTPS) flags |= WINHTTP_FLAG_SECURE;
	HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", path.c_str(), nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
	if (!hRequest) { outErr = L"WinHttpOpenRequest failed"; return false; }
	ScopeExit closeRequest{ [&]() { WinHttpCloseHandle(hRequest); } };

	if (!WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0))
	{
		outErr = L"WinHttpSendRequest failed";
		return false;
	}
	if (!WinHttpReceiveResponse(hRequest, nullptr))
	{
		outErr = L"WinHttpReceiveResponse failed";
		return false;
	}

	// Check HTTP status
	DWORD status = 0; DWORD statusSize = sizeof(status);
	WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize, WINHTTP_NO_HEADER_INDEX);
	if (status < 200 || status >= 300)
	{
		wchar_t msg[128] = {};
		swprintf_s(msg, L"HTTP status %lu", (unsigned long)status);
		outErr = msg;
		return false;
	}

	// Stream to file
	HANDLE hFile = CreateFileW(outFile.c_str(), GENERIC_WRITE, FILE_SHARE_READ, nullptr, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, nullptr);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		outErr = L"CreateFile failed";
		return false;
	}
	ScopeExit closeFile{ [&]() { CloseHandle(hFile); } };

	std::vector<char> buf(64 * 1024);

	for (;;)
	{
		DWORD avail = 0;
		if (!WinHttpQueryDataAvailable(hRequest, &avail))
		{
			outErr = L"WinHttpQueryDataAvailable failed";
			return false;
		}
		if (avail == 0) break;
		DWORD toRead = min(avail, (DWORD)buf.size());
		DWORD read = 0;
		if (!WinHttpReadData(hRequest, buf.data(), toRead, &read))
		{
			outErr = L"WinHttpReadData failed";
			return false;
		}
		if (read == 0) break;
		DWORD wrote = 0;
		if (!WriteFile(hFile, buf.data(), read, &wrote, nullptr) || wrote != read)
		{
			outErr = L"WriteFile failed";
			return false;
		}
	}
	return true;
}


static bool WinHttpDownloadUrlToUtf8String(const wchar_t* url, std::string& out, std::wstring& outErr)
{
	out.clear();
	outErr.clear();
	if (!url || !*url) { outErr = L"Empty URL"; return false; }

	URL_COMPONENTS uc{};
	uc.dwStructSize = sizeof(uc);
	uc.dwSchemeLength = (DWORD)-1;
	uc.dwHostNameLength = (DWORD)-1;
	uc.dwUrlPathLength = (DWORD)-1;
	uc.dwExtraInfoLength = (DWORD)-1;
	if (!WinHttpCrackUrl(url, 0, 0, &uc)) { outErr = L"WinHttpCrackUrl failed"; return false; }
	if (!uc.lpszHostName || uc.dwHostNameLength == 0) { outErr = L"Invalid URL host"; return false; }
	if (!uc.lpszUrlPath || uc.dwUrlPathLength == 0) { outErr = L"Invalid URL path"; return false; }

	std::wstring host(uc.lpszHostName, uc.dwHostNameLength);
	std::wstring path;
	path.assign(uc.lpszUrlPath, uc.dwUrlPathLength);
	if (uc.lpszExtraInfo && uc.dwExtraInfoLength) path.append(uc.lpszExtraInfo, uc.dwExtraInfoLength);

	HINTERNET hSession = WinHttpOpen(AppConstants::kUserAgentW, WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0);
	if (!hSession) { outErr = L"WinHttpOpen failed"; return false; }
	ScopeExit closeSession{ [&]() { WinHttpCloseHandle(hSession); } };

	HINTERNET hConnect = WinHttpConnect(hSession, host.c_str(), uc.nPort, 0);
	if (!hConnect) { outErr = L"WinHttpConnect failed"; return false; }
	ScopeExit closeConnect{ [&]() { WinHttpCloseHandle(hConnect); } };

	DWORD flags = 0;
	if (uc.nScheme == INTERNET_SCHEME_HTTPS) flags |= WINHTTP_FLAG_SECURE;
	HINTERNET hRequest = WinHttpOpenRequest(hConnect, L"GET", path.c_str(), nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags);
	if (!hRequest) { outErr = L"WinHttpOpenRequest failed"; return false; }
	ScopeExit closeRequest{ [&]() { WinHttpCloseHandle(hRequest); } };

	if (!WinHttpSendRequest(hRequest, WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0)) { outErr = L"WinHttpSendRequest failed"; return false; }
	if (!WinHttpReceiveResponse(hRequest, nullptr)) { outErr = L"WinHttpReceiveResponse failed"; return false; }

	DWORD status = 0; DWORD statusSize = sizeof(status);
	WinHttpQueryHeaders(hRequest, WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize, WINHTTP_NO_HEADER_INDEX);
	if (status < 200 || status >= 300)
	{
		wchar_t msg[128] = {};
		swprintf_s(msg, L"HTTP status %lu", (unsigned long)status);
		outErr = msg;
		return false;
	}

	std::vector<char> buf(64 * 1024);
	for (;;)
	{
		DWORD avail = 0;
		if (!WinHttpQueryDataAvailable(hRequest, &avail)) { outErr = L"WinHttpQueryDataAvailable failed"; return false; }
		if (avail == 0) break;
		DWORD toRead = min(avail, (DWORD)buf.size());
		DWORD read = 0;
		if (!WinHttpReadData(hRequest, buf.data(), toRead, &read)) { outErr = L"WinHttpReadData failed"; return false; }
		if (read == 0) break;
		out.append(buf.data(), buf.data() + read);
	}
	return true;
}

static bool WinHttpDownloadUrlToWideString(const wchar_t* url, std::wstring& out, std::wstring& outErr)
{
	std::string bytes;
	if (!WinHttpDownloadUrlToUtf8String(url, bytes, outErr)) return false;
	out = Utf8ToWide(bytes);
	return true;
}
struct UpdateResult
{
	bool ok = false;
	bool different = false;
	std::wstring localVersion;
	std::wstring remoteVersion;
	std::string expectedSha256Lower; // from version.txt line 2
	std::wstring err;
	fs::path downloadedTemp;
};

static int RunUpdateHelperMode(int argc, wchar_t** argv)
{
	// args: --apply-update <pid> <downloaded_exe> <target_exe>
	if (argc < 5) return 2;
	DWORD pid = (DWORD)_wtoi(argv[2]);
	fs::path downloaded = argv[3];
	fs::path target = argv[4];

	HANDLE hProc = OpenProcess(SYNCHRONIZE, FALSE, pid);
	if (hProc)
	{
		WaitForSingleObject(hProc, 90 * 1000);
		CloseHandle(hProc);
	}

	if (!MoveFileExW(downloaded.c_str(), target.c_str(), MOVEFILE_REPLACE_EXISTING))
	{
		// Best-effort cleanup: do not leave the downloaded temp exe behind if applying the update fails.
		(void)DeleteFileW(downloaded.c_str());
		return 3;
	}

	// Ensure the updated exe will resolve side-by-side DLLs from its folder.
	SetDllDirectoryW(target.parent_path().c_str());

	ShellExecuteW(nullptr, L"open", target.c_str(), nullptr, target.parent_path().c_str(), SW_SHOWNORMAL);
	return 0;
}

static bool LaunchUpdateHelperAndExitCurrent(const fs::path& downloadedExe)
{
	fs::path curExe = GetThisExePath();
	if (curExe.empty()) return false;
	fs::path workDir = curExe.parent_path();
	fs::path helper = workDir / L"MapPackSyncTool_Updater.exe";
	CopyFileW(curExe.c_str(), helper.c_str(), FALSE);

	DWORD pid = GetCurrentProcessId();
	std::wstring cmd;
	cmd.reserve(1024);
	cmd.append(L"\"");
	cmd.append(helper.wstring());
	cmd.append(L"\" --apply-update ");
	cmd.append(std::to_wstring(pid));
	cmd.append(L" \"");
	cmd.append(downloadedExe.wstring());
	cmd.append(L"\" \"");
	cmd.append(curExe.wstring());
	cmd.append(L"\"");

	STARTUPINFOW si{}; si.cb = sizeof(si);
	PROCESS_INFORMATION pi{};
	std::vector<wchar_t> cmdBuf(cmd.begin(), cmd.end());
	cmdBuf.push_back(L'\0');
	BOOL ok = CreateProcessW(helper.c_str(), cmdBuf.data(), nullptr, nullptr, FALSE, 0, nullptr, workDir.c_str(), &si, &pi);
	if (!ok) return false;
	CloseHandle(pi.hThread);
	CloseHandle(pi.hProcess);

	PostMessageW(g_state->hMainWnd, WM_CLOSE, 0, 0);
	return true;
}


// --------------------------------------------------
// Update version.txt parsing + numeric version comparison
// version.txt format:
//   Line 1: numeric dotted version (e.g. 0.0.12)
//   Line 2: sha256=<64-hex>  (or just 64-hex)
// --------------------------------------------------
static bool ExtractExpectedSha256Lower(const std::wstring& lineIn, std::string& outLower)
{
	outLower.clear();
	std::wstring line = lineIn;
	TrimInPlace(line);
	if (line.empty()) return false;

	// Optional "sha256=" prefix (case-insensitive)
	const wchar_t* kPrefix = L"sha256=";
	if (line.size() >= 7)
	{
		bool match = true;
		for (size_t i = 0; i < 7; ++i)
		{
			wchar_t a = line[i];
			wchar_t b = kPrefix[i];
			if (a >= L'A' && a <= L'Z') a = (wchar_t)(a - L'A' + L'a');
			if (a != b) { match = false; break; }
		}
		if (match)
		{
			line = line.substr(7);
			TrimInPlace(line);
		}
	}

	if (line.size() != 64) return false;

	outLower.reserve(64);
	for (wchar_t wc : line)
	{
		char c = (wc <= 0x7F) ? (char)wc : '\0';
		bool ok = (c >= '0' && c <= '9') ||
			(c >= 'a' && c <= 'f') ||
			(c >= 'A' && c <= 'F');
		if (!ok) { outLower.clear(); return false; }
		if (c >= 'A' && c <= 'F') c = (char)(c - 'A' + 'a');
		outLower.push_back(c);
	}
	return true;
}

static bool ParseVersionTxt2Line(const std::wstring& txt, std::wstring& outVer, std::string& outShaLower, std::wstring& outErr)
{
	outVer.clear();
	outShaLower.clear();
	outErr.clear();

	// Split on '\n' into at least 2 lines.
	size_t p1 = txt.find(L'\n');
	if (p1 == std::wstring::npos)
	{
		outErr = L"version.txt must have 2 lines: version then sha256.";
		return false;
	}
	std::wstring line1 = txt.substr(0, p1);
	size_t start2 = p1 + 1;
	size_t p2 = txt.find(L'\n', start2);
	std::wstring line2 = (p2 == std::wstring::npos) ? txt.substr(start2) : txt.substr(start2, p2 - start2);

	TrimInPlace(line1);
	TrimInPlace(line2);
	if (!line1.empty() && line1.back() == L'\r') line1.pop_back();
	if (!line2.empty() && line2.back() == L'\r') line2.pop_back();

	if (line1.empty())
	{
		outErr = L"version.txt line 1 (version) is empty.";
		return false;
	}
	if (!ExtractExpectedSha256Lower(line2, outShaLower))
	{
		outErr = L"version.txt line 2 must be sha256=<64-hex> (or just 64-hex).";
		return false;
	}

	outVer = line1;
	return true;
}

static bool ParseNumericDottedVersion(const std::wstring& sIn, std::vector<uint32_t>& outParts)
{
	outParts.clear();
	std::wstring s = sIn;
	TrimInPlace(s);
	if (s.empty()) return false;

	// Strict: only digits and '.'; no leading/trailing '.'; no consecutive '.'
	if (s.front() == L'.' || s.back() == L'.') return false;

	uint64_t cur = 0;
	bool inPart = false;

	for (size_t i = 0; i < s.size(); ++i)
	{
		wchar_t c = s[i];
		if (c == L'.')
		{
			if (!inPart) return false; // consecutive dots or empty segment
			if (cur > 0xFFFFFFFFull) return false;
			outParts.push_back((uint32_t)cur);
			cur = 0;
			inPart = false;
			continue;
		}
		if (c >= L'0' && c <= L'9')
		{
			inPart = true;
			cur = cur * 10ull + (uint64_t)(c - L'0');
			if (cur > 0xFFFFFFFFull) return false;
			continue;
		}
		// Any other character is invalid (digits and dots only).
		return false;
	}
	if (!inPart) return false;
	outParts.push_back((uint32_t)cur);
	return true;
}

static int CompareNumericVersions(const std::vector<uint32_t>& a, const std::vector<uint32_t>& b)
{
	const size_t n = (a.size() > b.size()) ? a.size() : b.size();
	for (size_t i = 0; i < n; ++i)
	{
		uint32_t av = (i < a.size()) ? a[i] : 0;
		uint32_t bv = (i < b.size()) ? b[i] : 0;
		if (av < bv) return -1;
		if (av > bv) return 1;
	}
	return 0;
}


// ==========================================================
// Update security (production): Authenticode + pinned signer
// - Production mode: DEBUG_MESSAGE is commented out.
// - Non-production mode: DEBUG_MESSAGE enabled -> legacy SHA-256 path.
// ==========================================================
#ifndef DEBUG_MESSAGE
static const wchar_t* kAllowedSignerThumbprints[] =
{
	L"8788209B20FDFA15C95C40DCBFDC038B54CA11BB", // current signing cert
	// Add future renewal thumbprints here:
	// L"NEWTHUMBPRINTHERE",
};
static const size_t kAllowedSignerThumbprintsCount =
sizeof(kAllowedSignerThumbprints) / sizeof(kAllowedSignerThumbprints[0]);

static std::wstring BytesToHexUpper(const BYTE* data, DWORD cb)
{
	static const wchar_t* kHex = L"0123456789ABCDEF";
	std::wstring out;
	out.reserve((size_t)cb * 2);
	for (DWORD i = 0; i < cb; ++i)
	{
		BYTE b = data[i];
		out.push_back(kHex[(b >> 4) & 0xF]);
		out.push_back(kHex[b & 0xF]);
	}
	return out;
}

static bool IsAllowedSignerThumbprint(const std::wstring& thumbUpper)
{
	for (size_t i = 0; i < kAllowedSignerThumbprintsCount; ++i)
	{
		if (_wcsicmp(thumbUpper.c_str(), kAllowedSignerThumbprints[i]) == 0)
			return true;
	}
	return false;
}

static bool VerifyAuthenticodeSignatureCacheOnly(const wchar_t* filePath, std::wstring& outErr)
{
	outErr.clear();

	WINTRUST_FILE_INFO fi = {};
	fi.cbStruct = sizeof(fi);
	fi.pcwszFilePath = filePath;

	WINTRUST_DATA wtd = {};
	wtd.cbStruct = sizeof(wtd);
	wtd.dwUIChoice = WTD_UI_NONE;
	wtd.fdwRevocationChecks = WTD_REVOKE_WHOLECHAIN;
	wtd.dwUnionChoice = WTD_CHOICE_FILE;
	wtd.pFile = &fi;

	// Cache-only URL retrieval (as requested). This avoids network fetches for CRL/OCSP.
	wtd.dwProvFlags = WTD_CACHE_ONLY_URL_RETRIEVAL;

	wtd.dwStateAction = WTD_STATEACTION_VERIFY;

	GUID action = WINTRUST_ACTION_GENERIC_VERIFY_V2;
	LONG st = WinVerifyTrust(nullptr, &action, &wtd);

	wtd.dwStateAction = WTD_STATEACTION_CLOSE;
	(void)WinVerifyTrust(nullptr, &action, &wtd);

	if (st == ERROR_SUCCESS)
		return true;

	wchar_t buf[128] = {};
	_snwprintf_s(buf, _TRUNCATE, L"WinVerifyTrust failed (0x%08lX).", (unsigned long)st);
	outErr = buf;
	return false;
}

static bool GetSignerThumbprintSha1HexUpper(const wchar_t* filePath, std::wstring& outThumbUpper, std::wstring& outErr)
{
	outThumbUpper.clear();
	outErr.clear();

	CertStoreHandle hStore;
	CryptMsgHandle hMsg;

	DWORD dwEncoding = 0, dwContentType = 0, dwFormatType = 0;
	HCERTSTORE rawStore = nullptr;
	HCRYPTMSG rawMsg = nullptr;
	if (!CryptQueryObject(
		CERT_QUERY_OBJECT_FILE,
		filePath,
		CERT_QUERY_CONTENT_FLAG_PKCS7_SIGNED_EMBED,
		CERT_QUERY_FORMAT_FLAG_BINARY,
		0,
		&dwEncoding,
		&dwContentType,
		&dwFormatType,
		&rawStore,
		&rawMsg,
		nullptr))
	{
		outErr = L"CryptQueryObject failed.";
		return false;
	}

	// Wrap handles for automatic cleanup
	hStore = CertStoreHandle(rawStore);
	hMsg = CryptMsgHandle(rawMsg);

	DWORD cbSignerInfo = 0;
	if (!CryptMsgGetParam(hMsg, CMSG_SIGNER_INFO_PARAM, 0, nullptr, &cbSignerInfo) || cbSignerInfo == 0)
	{
		outErr = L"CryptMsgGetParam(CMSG_SIGNER_INFO_PARAM) failed.";
		return false;
	}

	std::vector<BYTE> signerInfoBuf(cbSignerInfo);
	if (!CryptMsgGetParam(hMsg, CMSG_SIGNER_INFO_PARAM, 0, signerInfoBuf.data(), &cbSignerInfo))
	{
		outErr = L"CryptMsgGetParam(CMSG_SIGNER_INFO_PARAM) read failed.";
		return false;
	}

	const CMSG_SIGNER_INFO* pSignerInfo = reinterpret_cast<const CMSG_SIGNER_INFO*>(signerInfoBuf.data());

	CERT_INFO certInfo = {};
	certInfo.Issuer = pSignerInfo->Issuer;
	certInfo.SerialNumber = pSignerInfo->SerialNumber;

	PCCERT_CONTEXT pCertContext = CertFindCertificateInStore(
		hStore,
		X509_ASN_ENCODING | PKCS_7_ASN_ENCODING,
		0,
		CERT_FIND_SUBJECT_CERT,
		&certInfo,
		nullptr);

	if (!pCertContext)
	{
		outErr = L"CertFindCertificateInStore failed.";
		return false;
	}

	DWORD cbHash = 0;
	if (!CertGetCertificateContextProperty(pCertContext, CERT_HASH_PROP_ID, nullptr, &cbHash) || cbHash == 0)
	{
		CertFreeCertificateContext(pCertContext);
		outErr = L"CertGetCertificateContextProperty(CERT_HASH_PROP_ID) failed.";
		return false;
	}

	std::vector<BYTE> hashBytes(cbHash);
	if (!CertGetCertificateContextProperty(pCertContext, CERT_HASH_PROP_ID, hashBytes.data(), &cbHash))
	{
		CertFreeCertificateContext(pCertContext);
		outErr = L"CertGetCertificateContextProperty(CERT_HASH_PROP_ID) read failed.";
		return false;
	}

	outThumbUpper = BytesToHexUpper(hashBytes.data(), cbHash);

	CertFreeCertificateContext(pCertContext);
	return true;
}

static bool VerifyDownloadedUpdateExe_Production(const wchar_t* filePath, std::wstring& outErr)
{
	outErr.clear();

	std::wstring sigErr;
	if (!VerifyAuthenticodeSignatureCacheOnly(filePath, sigErr))
	{
		outErr = L"Authenticode verification failed: " + sigErr;
		return false;
	}

	std::wstring thumbUpper;
	std::wstring thumbErr;
	if (!GetSignerThumbprintSha1HexUpper(filePath, thumbUpper, thumbErr))
	{
		outErr = L"Could not read signer certificate thumbprint: " + thumbErr;
		return false;
	}

	if (!IsAllowedSignerThumbprint(thumbUpper))
	{
		outErr = L"Signer certificate is not trusted.\r\n\r\nSigner thumbprint:\r\n" + thumbUpper;
		return false;
	}

	return true;
}
#endif // !DEBUG_MESSAGE

static unsigned __stdcall UpdateThreadProc(void* p)
{
	UpdateResult* res = (UpdateResult*)p;
	res->ok = false;
	res->different = false;
	res->err.clear();

	fs::path curExe = GetThisExePath();
	if (curExe.empty()) { res->err = L"Cannot locate current exe"; return 0; }

	// Update availability is decided by version.txt:
	//   line 1: numeric dotted version (e.g. 0.0.12)
	//   line 2: sha256=<64-hex> (or just 64-hex)
	res->localVersion = GetExeFileVersionString().c_str();
	TrimInPlace(res->localVersion);

	std::wstring versionTxt;
	std::wstring verErr;
	if (!WinHttpDownloadUrlToWideString(kUpdateVersionUrl, versionTxt, verErr))
	{
		res->err = L"Failed to download version.txt: " + verErr;
		return 0;
	}
	TrimInPlace(versionTxt);
	if (versionTxt.empty())
	{
		res->err = L"version.txt was empty";
		return 0;
	}

	std::wstring parseErr;
	if (!ParseVersionTxt2Line(versionTxt, res->remoteVersion, res->expectedSha256Lower, parseErr))
	{
		res->err = L"Invalid version.txt: " + parseErr;
		return 0;
	}

	// Strict numeric comparison (digits and dots only).
	std::vector<uint32_t> localParts, remoteParts;
	if (!ParseNumericDottedVersion(res->localVersion, localParts))
	{
		res->err = L"Local version is not numeric dotted: " + res->localVersion;
		return 0;
	}
	if (!ParseNumericDottedVersion(res->remoteVersion, remoteParts))
	{
		res->err = L"Remote version is not numeric dotted: " + res->remoteVersion;
		return 0;
	}

	const int cmp = CompareNumericVersions(remoteParts, localParts);
	// Update only when remote > local.
	if (cmp <= 0)
	{
		res->different = false;
		res->ok = true;
		return 0;
	}
	res->different = true;

	// New version available: download the updated exe to a temp file in the SAME directory as the exe so any relaunch uses the same dependency neighborhood.
	fs::path tempExe = curExe.parent_path() / L"MapPackSyncTool.exe.download";
	res->downloadedTemp = tempExe;

	std::wstring dlErr;
	if (!WinHttpDownloadUrlToFile(kUpdateExeUrl, tempExe, dlErr))
	{
		res->err = L"Download failed: " + dlErr;
		return 0;
	}

	// Verify downloaded update exe.
#ifndef DEBUG_MESSAGE
	// Production: STRICT Authenticode + pinned signer thumbprint allow-list (no hash fallback).
	std::wstring verifyErr;
	if (!VerifyDownloadedUpdateExe_Production(tempExe.c_str(), verifyErr))
	{
		(void)DeleteFileW(tempExe.c_str());
		res->err = L"Update rejected.\r\n\r\n" + verifyErr;
		return 0;
	}
#else
	// Non-production: legacy SHA-256 check (easy testing without a signature).
	std::string gotSha;
	if (!Sha256FileHexLower(tempExe, gotSha))
	{
		(void)DeleteFileW(tempExe.c_str());
		res->err = L"Failed to compute SHA-256 of downloaded update.";
		return 0;
	}
	if (_stricmp(gotSha.c_str(), res->expectedSha256Lower.c_str()) != 0)
	{
		(void)DeleteFileW(tempExe.c_str());
		res->err = std::wstring(L"Update SHA-256 mismatch; refusing to install.\r\n\r\nExpected SHA-256:\r\n") +
			std::wstring(res->expectedSha256Lower.begin(), res->expectedSha256Lower.end()) +
			std::wstring(L"\r\n\r\nCurrent SHA-256:\r\n") +
			std::wstring(gotSha.begin(), gotSha.end());
		return 0;
	}
#endif


	res->ok = true;
	return 0;
}

static void StartCheckForUpdates()
{
	AppState* st = g_state;
	if (!st) return;
	if (st->isUpdateRunning.exchange(true)) return;
	UpdateCheckUpdatesButtonEnabled();

	UpdateResult* res = new UpdateResult();
	uintptr_t th = _beginthreadex(nullptr, 0, UpdateThreadProc, res, 0, nullptr);
	if (!th)
	{
		st->isUpdateRunning.store(false);
		UpdateCheckUpdatesButtonEnabled();
		delete res;
		Log("ERROR: Failed to start update thread.\r\n");
		return;
	}
	st->hUpdateThread = (HANDLE)th;

	// Waiter thread: waits for the update thread then posts WM_APP+77 with UpdateResult* back to the UI thread.
	uintptr_t waiter = _beginthreadex(nullptr, 0, [](void* param)->unsigned {
		auto* pair = (std::pair<AppState*, UpdateResult*>*)param;
		WaitForSingleObject(pair->first->hUpdateThread, INFINITE);
		{ UiEvent* ev = new UiEvent(); ev->kind = UiEventKind::UpdateResultPtr; ev->ptr = pair->second; PostMessageW(pair->first->hMainWnd, WM_APP_UI_EVENT, 0, (LPARAM)ev); }
		delete pair;
		return 0;
		}, new std::pair<AppState*, UpdateResult*>(st, res), 0, nullptr);
	if (waiter) CloseHandle((HANDLE)waiter);
}
// --------------------------------------------------
// Window Proc
// --------------------------------------------------


// ---- WndProc message handlers (refactor for readability) ----
static LRESULT HandleWmSize(HWND hwnd, WPARAM wParam, LPARAM lParam)
{
	(void)wParam; (void)lParam;
	AppState* st2 = (AppState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
	if (st2 && st2->hOutput)
		LayoutMainWindow(hwnd, st2);
	return 0;
}

static LRESULT HandleWmGetMinMaxInfo(HWND hwnd, LPARAM lParam)
{
	MINMAXINFO* mmi = (MINMAXINFO*)lParam;
	DWORD style = (DWORD)GetWindowLongPtrW(hwnd, GWL_STYLE);
	DWORD ex = (DWORD)GetWindowLongPtrW(hwnd, GWL_EXSTYLE);
	RECT r{ 0, 0, MIN_CLIENT_W, MIN_CLIENT_H };
	AdjustWindowRectEx(&r, style, FALSE, ex);
	mmi->ptMinTrackSize.x = (r.right - r.left);
	mmi->ptMinTrackSize.y = (r.bottom - r.top);
	return 0;
}

static LRESULT HandleWmCommand(AppState* st, HWND hwnd, WPARAM wParam, LPARAM lParam)
{

	if ((HWND)lParam == st->hBrowseBtn)
	{
		std::wstring path;
		if (BrowseForFolder(hwnd, path))
		{
			SetWindowTextW(st->hFolderEdit, path.c_str());
			// Persist selection for next launch (INI created on first write).
			if (!path.empty())
				IniWriteLastFolder(path);
		}
	}
	else if ((HWND)lParam == st->hRunButton)
	{
		st->logActionsArmed = true;
		UpdateHelpButtonEnabled();
		ClearOutput();
		if (IsProcessRunningByName(L"istaria.exe"))
		{
			MessageBoxW(
				hwnd,
				L"Istaria is currently running.\r\n\r\n"
				L"Please close Istaria before running MapPack Sync Tool.\r\n"
				L"This tool should only be used when Istaria is not running.",
				L"MapPack Sync Tool",
				MB_ICONERROR | MB_OK
			);
			Log("ERROR: Aborted - istaria.exe is running. Exit the game before attempting to sync.\r\n");
			return 0;
		}
		StartSyncIfNotRunning();
	}
	else if ((HWND)lParam == st->hCancelBtn)
	{
		if (st && st->isRunning.load())
		{
			st->cancelRequested.store(true);
			EnableWindow(st->hCancelBtn, FALSE);
			PostProgressTextW(L"Cancel requested... finishing current transfer.");
			Log("INFO: Cancel sync requested.\r\n");
		}
	}
	else if ((HWND)lParam == st->hDeleteBtn)
	{
		st->logActionsArmed = true;
		UpdateHelpButtonEnabled();
		ClearOutput();
		if (IsProcessRunningByName(L"istaria.exe"))
		{
			MessageBoxW(
				hwnd,
				L"Istaria is currently running.\r\n\r\n"
				L"Please close Istaria before removing MapPack files.\r\n"
				L"This tool should only be used when Istaria is not running.",
				L"MapPack Sync Tool",
				MB_ICONERROR | MB_OK
			);
			Log("ERROR: Aborted remove - istaria.exe is running. Exit the game before attempting to remove.\r\n");
			return 0;
		}
		StartRemoveIfNotRunning();
	}
	else if ((HWND)lParam == st->hCopyLogBtn)
	{
		CopyOutputToClipboard();
	}
	else if ((HWND)lParam == st->hSaveLogBtn)
	{
		SaveOutputToFile();
	}
	else if ((HWND)lParam == st->hCheckUpdatesBtn)
	{
		StartCheckForUpdates();
	}
	else if ((HWND)lParam == st->hHelpBtn)
	{
		MessageBeep(MB_ICONASTERISK);
		const int r = MessageBoxW(
			hwnd,
			L"Clear Log and Load Help?",
			L"MapPack Sync Tool",
			MB_OKCANCEL | MB_ICONQUESTION
		);
		if (r == IDOK)
		{
			st->logActionsArmed = false;
			UpdateHelpButtonEnabled();
			// Clear then load MapPackSyncTool.txt into the log.
			LoadHelpTextIntoOutput(true, true);
		}
	}

	return DefWindowProc(hwnd, WM_COMMAND, wParam, lParam);
}

static LRESULT HandleWmAppUpdateResult(AppState* st, HWND hwnd, WPARAM wParam, LPARAM lParam)
{

	{
		std::unique_ptr<UpdateResult> res((UpdateResult*)lParam);
		if (st)
		{
			st->isUpdateRunning.store(false);
			UpdateCheckUpdatesButtonEnabled();
			if (st->hUpdateThread) { CloseHandle(st->hUpdateThread); st->hUpdateThread = nullptr; }
		}
		if (!res->ok)
		{
			std::wstring msg = L"Update check failed:\r\n\r\n" + res->err;
			MessageBoxW(hwnd, msg.c_str(), L"MapPack Sync Tool", MB_OK | MB_ICONERROR);
			if (!res->downloadedTemp.empty()) { DeleteFileW(res->downloadedTemp.c_str()); }
			return 0;
		}
		if (!res->different)
		{
			std::wstring msg = L"You are already have the latest version.\r\n\r\nCurrent version: v" + res->localVersion + L"\r\n\r\nAvailable Version: v" + res->remoteVersion;
			MessageBoxW(hwnd, msg.c_str(), L"MapPack Sync Tool", MB_OK | MB_ICONINFORMATION);
			if (!res->downloadedTemp.empty()) { DeleteFileW(res->downloadedTemp.c_str()); }
			return 0;
		}
		std::wstring prompt = L"New version of MapPack Sync Tool is available.\r\n\r\nCurrent version: v" + res->localVersion + L"\r\n\r\nAvailable version: v" + res->remoteVersion + L"\r\n\r\nClick OK to proceed with the update.";
		int r = MessageBoxW(hwnd, prompt.c_str(), L"Checking for Updates...", MB_OKCANCEL | MB_ICONINFORMATION);
		if (r != IDOK)
		{
			if (!res->downloadedTemp.empty()) { DeleteFileW(res->downloadedTemp.c_str()); }
			return 0;
		}
		// Launch helper (same directory) and exit so the helper can replace the running exe.
		if (!LaunchUpdateHelperAndExitCurrent(res->downloadedTemp))
		{
			MessageBoxW(hwnd, L"Failed to launch update helper.", L"MapPack Sync Tool", MB_OK | MB_ICONERROR);
			if (!res->downloadedTemp.empty()) { DeleteFileW(res->downloadedTemp.c_str()); }
		}
		return 0;
	}
}

static LRESULT HandleWmClose(AppState* st, HWND hwnd)
{

	{
		if (st && st->isRunning.load())
		{
			st->cancelRequested.store(true);
			st->pendingExitAfterWorker = true;
			if (st->hCancelBtn) EnableWindow(st->hCancelBtn, FALSE);
			FreezeProgressOnCancel(st);
			if (st->hDeleteBtn) EnableWindow(st->hDeleteBtn, TRUE);
			PostProgressTextW(L"Cancel requested... exiting when safe.");
			Log("INFO: Window close requested during sync; canceling.\r\n");
			return 0;
		}
		DestroyWindow(hwnd);
		return 0;
	}
}

static LRESULT HandleWmDestroy(AppState* st, HWND hwnd)
{

	if (st)
	{
		if (st->hTooltip) { DestroyWindow(st->hTooltip); st->hTooltip = nullptr; }
		if (st->hFontUI) { DeleteObject(st->hFontUI); st->hFontUI = nullptr; }
		if (st->hFontMono) { DeleteObject(st->hFontMono); st->hFontMono = nullptr; }
	}
	PostQuitMessage(0);

	return DefWindowProc(hwnd, WM_DESTROY, 0, 0);
}

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	AppState* st = (AppState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
	if (!st) st = g_state;
	switch (msg)
	{
	case WM_APP_UI_EVENT:
	{
		std::unique_ptr<UiEvent> ev((UiEvent*)lParam);
		if (!ev) return 0;
		AppState* st2 = (AppState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
		if (!st2) st2 = g_state;
		switch (ev->kind)
		{
		case UiEventKind::LogAppendW:
			AppendToOutputW(ev->text.c_str());
			break;
		case UiEventKind::ProgressMarqueeOn:
			if (st2)
			{
				if (st2->cancelRequested.load(std::memory_order_relaxed) || st2->progressFrozenOnCancel) { FreezeProgressOnCancel(st2); break; }
				SetProgressMarquee(st2, true);
			}
			break;
		case UiEventKind::ProgressMarqueeOff:
			if (st2)
			{
				if (st2->cancelRequested.load(std::memory_order_relaxed) || st2->progressFrozenOnCancel) { FreezeProgressOnCancel(st2); break; }
				SetProgressMarquee(st2, false);
			}
			break;
		case UiEventKind::ProgressInit:
			if (st2 && st2->hProgress)
			{
				int total = (int)ev->u1;
				if (total <= 0) total = 1;
				st2->progressTotal = total;
				st2->progressPos = 0;
				if (st2->cancelRequested.load(std::memory_order_relaxed) || st2->progressFrozenOnCancel) { FreezeProgressOnCancel(st2); break; }
				SetProgressMarquee(st2, false);
				SendMessageW(st2->hProgress, PBM_SETRANGE32, 0, total);
				SendMessageW(st2->hProgress, PBM_SETPOS, 0, 0);
			}
			break;
		case UiEventKind::ProgressSet:
			if (st2 && st2->hProgress)
			{
				int pos = (int)ev->u1;
				st2->progressPos = pos;
				if (st2->cancelRequested.load(std::memory_order_relaxed) || st2->progressFrozenOnCancel) { FreezeProgressOnCancel(st2); break; }
				SetProgressMarquee(st2, false);
				SendMessageW(st2->hProgress, PBM_SETPOS, (WPARAM)pos, 0);
			}
			break;
		case UiEventKind::ProgressTextW:
			if (st2 && st2->hProgressText) SetWindowTextW(st2->hProgressText, ev->text.c_str());
			break;
		case UiEventKind::WorkerDone:
			if (st2)
			{
				SetUiForWorkerRunning(st2, false);
				if (st2->hWorkerThread) { CloseHandle(st2->hWorkerThread); st2->hWorkerThread = nullptr; }
				if (st2->pendingExitAfterWorker) { DestroyWindow(hwnd); }
			}
			break;
		case UiEventKind::UpdateResultPtr:
			// Transfer ownership of UpdateResult* to the existing handler via lParam.
			return HandleWmAppUpdateResult(st2, hwnd, 0, (LPARAM)ev->ptr);
		default:
			break;
		}
		return 0;
	}
	case WM_APP_LOG:
	{
		unique_heap_wstr w((wchar_t*)lParam);
		if (w)
		{
			AppendToOutputW(w.get());
		}
		return 0;
	}
	case WM_APP_WORKER_DONE:
	{
		if (st)
		{
			SetUiForWorkerRunning(st, false);
			if (st->hWorkerThread)
			{
				CloseHandle(st->hWorkerThread);
				st->hWorkerThread = nullptr;
			}
			if (st->pendingExitAfterWorker)
			{
				DestroyWindow(hwnd);
				return 0;
			}
		}
		return 0;
	}
	case WM_APP_PROGRESS_MARQ_ON:
	{
		AppState* st2 = (AppState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
		if (!st2) return 0;
		if (st2->cancelRequested.load(std::memory_order_relaxed) || st2->progressFrozenOnCancel)
		{
			FreezeProgressOnCancel(st2);
			return 0;
		}
		SetProgressMarquee(st2, true);
		return 0;
	}
	case WM_APP_PROGRESS_MARQ_OFF:
	{
		AppState* st2 = (AppState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
		if (!st2) return 0;
		if (st2->cancelRequested.load(std::memory_order_relaxed) || st2->progressFrozenOnCancel)
		{
			FreezeProgressOnCancel(st2);
			return 0;
		}
		SetProgressMarquee(st2, false);
		return 0;
	}
	case WM_APP_PROGRESS_INIT:
	{
		AppState* st2 = (AppState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
		if (st2 && st2->hProgress)
		{
			int total = (int)wParam;
			if (total <= 0) total = 1;
			st2->progressTotal = total;
			st2->progressPos = 0;
			if (st2->cancelRequested.load(std::memory_order_relaxed) || st2->progressFrozenOnCancel)
			{
				FreezeProgressOnCancel(st2);
				return 0;
			}
			// Ensure marquee is fully off before switching to normal range/pos
			SetProgressMarquee(st2, false);
			SendMessageW(st2->hProgress, PBM_SETRANGE32, 0, total);
			SendMessageW(st2->hProgress, PBM_SETPOS, 0, 0);
		}
		return 0;
	}
	case WM_APP_PROGRESS_SET:
	{
		AppState* st2 = (AppState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
		if (st2 && st2->hProgress)
		{
			int pos = (int)wParam;
			st2->progressPos = pos;
			if (st2->cancelRequested.load(std::memory_order_relaxed) || st2->progressFrozenOnCancel)
			{
				FreezeProgressOnCancel(st2);
				return 0;
			}
			// If we somehow left marquee on, force it off
			SetProgressMarquee(st2, false);
			SendMessageW(st2->hProgress, PBM_SETPOS, (WPARAM)pos, 0);
		}
		return 0;
	}
	case WM_APP_PROGRESS_TEXT:
	{
		unique_heap_wstr w((wchar_t*)lParam);
		if (g_state && g_state->hProgressText)
			SetWindowTextW(g_state->hProgressText, w ? w.get() : L"");
		return 0;
	}
	case WM_SIZE:
		return HandleWmSize(hwnd, wParam, lParam);
	case WM_GETMINMAXINFO:
		return HandleWmGetMinMaxInfo(hwnd, lParam);
	case WM_COMMAND:
		return HandleWmCommand(st, hwnd, wParam, lParam);
	case WM_APP + 77:
		return HandleWmAppUpdateResult(st, hwnd, wParam, lParam);
	case WM_CLOSE:
		return HandleWmClose(st, hwnd);
	case WM_DESTROY:
		return HandleWmDestroy(st, hwnd);
	}
	return DefWindowProc(hwnd, msg, wParam, lParam);
}
// --------------------------------------------------
// WinMain
// --------------------------------------------------
int WINAPI wWinMain(_In_ HINSTANCE hInst, _In_opt_ HINSTANCE, _In_ PWSTR, _In_ int)
{
	static AppState state;
	g_state = &state;
	// Helper mode: apply an update after the main process exits.
	int argc = 0;
	wchar_t** argv = CommandLineToArgvW(GetCommandLineW(), &argc);
	if (argv && argc >= 2 && _wcsicmp(argv[1], L"--apply-update") == 0)
	{
		int rc = RunUpdateHelperMode(argc, argv);
		LocalFree(argv);
		return rc;
	}
	if (argv) LocalFree(argv);
	// Enforce single instance for normal GUI mode.
	if (!AcquireSingleInstanceMutex())
	{
		ActivateExistingInstance();
		return 0;
	}

	{
		fs::path curExe = GetThisExePath();
		if (!curExe.empty()) TryDeleteUpdaterExeBestEffort(curExe.parent_path());
	}
	// WinHTTP does not require global initialization.
	LoadLibraryW(L"Msftedit.dll");
	INITCOMMONCONTROLSEX icc{};
	icc.dwSize = sizeof(icc);
	icc.dwICC = ICC_WIN95_CLASSES | ICC_PROGRESS_CLASS;
	InitCommonControlsEx(&icc);
	WNDCLASSEX wc{};
	wc.cbSize = sizeof(wc);
	wc.lpfnWndProc = WndProc;
	wc.hInstance = hInst;
	wc.lpszClassName = L"DEF_SYNC_GUI";
	wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wc.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wc.hIcon = (HICON)LoadImageW(hInst, MAKEINTRESOURCEW(IDI_MAPPACKSYNCTOOL),
		IMAGE_ICON, 32, 32, LR_DEFAULTCOLOR);
	wc.hIconSm = (HICON)LoadImageW(hInst, MAKEINTRESOURCEW(IDI_SMALL),
		IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR);
	RegisterClassEx(&wc);
	DWORD exStyle = 0;
	DWORD style = (WS_OVERLAPPEDWINDOW & ~WS_MAXIMIZEBOX) | WS_VISIBLE | WS_CLIPCHILDREN;
	int winW = 0, winH = 0;
	ComputeWindowSizeFromClientStyle(style, exStyle, MAIN_WINDOW_WIDTH, MAIN_WINDOW_HEIGHT, winW, winH);
	int screenW = GetSystemMetrics(SM_CXSCREEN);
	int screenH = GetSystemMetrics(SM_CYSCREEN);
	int winX = (screenW > 0) ? ((screenW - winW) / 2) : 200;
	int winY = (screenH > 0) ? ((screenH - winH) / 2 - 20) : 200;
	g_state->hMainWnd = CreateWindowEx(
		exStyle, wc.lpszClassName,
		kWindowTitle.c_str(),
		style,
		winX, winY, winW, winH,
		nullptr, nullptr, hInst, nullptr);
	SetWindowLongPtrW(g_state->hMainWnd, GWLP_USERDATA, (LONG_PTR)g_state);
	SendMessageW(g_state->hMainWnd, WM_SETICON, ICON_BIG, (LPARAM)wc.hIcon);
	SendMessageW(g_state->hMainWnd, WM_SETICON, ICON_SMALL, (LPARAM)wc.hIconSm);
	ShowWindow(g_state->hMainWnd, SW_SHOW);
	UpdateWindow(g_state->hMainWnd);
	g_state->hFolderLabel = CreateWindowW(L"STATIC", L"Istaria Base Game Folder:", WS_CHILD | WS_VISIBLE,
		10, 15, 170, 20, g_state->hMainWnd, nullptr, hInst, nullptr);
	g_state->hFolderEdit = CreateWindowW(
		L"EDIT", L"",
		WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL,
		190, 12, 410, 22,
		g_state->hMainWnd, nullptr, hInst, nullptr);

	// Load last folder from portable INI (if present).
	{
		std::wstring last = IniReadLastFolder();
		if (!last.empty())
			SetWindowTextW(g_state->hFolderEdit, last.c_str());
	}

	// ===== DEBUG DEFAULT PATH =====
	//#if defined(DEBUG_MESSAGE) && defined(_DEBUG)
#if defined(DEBUG_MESSAGE)
	// Debug preset only if INI didn't already supply a value.
	if (GetWindowTextLengthW(g_state->hFolderEdit) == 0)
		SetWindowTextW(g_state->hFolderEdit, L"C:\\temp\\defs2");
#endif
	// ==============================

	g_state->hBrowseBtn = CreateWindowW(
		L"BUTTON", L"Browse...",
		WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
		620, 12, 80, 22,
		g_state->hMainWnd, nullptr, hInst, nullptr);
	g_state->hRunButton = CreateWindowW(
		L"BUTTON", L"Add / Sync",
		WS_CHILD | WS_VISIBLE | BS_DEFPUSHBUTTON,
		690, 12, 60, 22,
		g_state->hMainWnd, nullptr, hInst, nullptr);
	g_state->hCancelBtn = CreateWindowW(
		L"BUTTON", L"Cancel Action",
		WS_CHILD | WS_VISIBLE | WS_DISABLED | BS_PUSHBUTTON,
		755, 12, 60, 22,
		g_state->hMainWnd, nullptr, hInst, nullptr);
	g_state->hDeleteBtn = CreateWindowW(
		L"BUTTON", L"Remove",
		WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
		0, 0, 130, 22,
		g_state->hMainWnd, nullptr, hInst, nullptr);

	g_state->hCopyLogBtn = CreateWindowW(
		L"BUTTON", L"Copy Log",
		WS_CHILD | WS_VISIBLE | WS_DISABLED | BS_PUSHBUTTON,
		0, 0, 92, 22,
		g_state->hMainWnd, nullptr, hInst, nullptr);
	g_state->hSaveLogBtn = CreateWindowW(
		L"BUTTON", L"Save Log...",
		WS_CHILD | WS_VISIBLE | WS_DISABLED | BS_PUSHBUTTON,
		0, 0, 92, 22,
		g_state->hMainWnd, nullptr, hInst, nullptr);
	g_state->hCheckUpdatesBtn = CreateWindowW(
		L"BUTTON", L"Check for Updates",
		WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
		0, 0, 130, 22,
		g_state->hMainWnd, nullptr, hInst, nullptr);
	g_state->hHelpBtn = CreateWindowW(
		L"BUTTON", L"?",
		WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
		0, 0, 22, 22,
		g_state->hMainWnd, nullptr, hInst, nullptr);
	// --------------------------------------------------
	// PROGRESS BAR (FIXED: create WITHOUT PBS_MARQUEE)
	// --------------------------------------------------
	g_state->hProgress = CreateWindowExW(
		0, PROGRESS_CLASSW, nullptr,
		WS_CHILD | WS_VISIBLE | PBS_SMOOTH, // <-- no PBS_MARQUEE here
		10, 42, 800, 14,
		g_state->hMainWnd, (HMENU)2001, hInst, nullptr);
	SendMessageW(g_state->hProgress, PBM_SETRANGE32, 0, 1);
	SendMessageW(g_state->hProgress, PBM_SETPOS, 0, 0);
	g_state->hProgressText = CreateWindowExW(
		WS_EX_CLIENTEDGE, L"STATIC", L"Ready",
		WS_CHILD | WS_VISIBLE,
		10, 60, 800, 22,
		g_state->hMainWnd, (HMENU)2002, hInst, nullptr);
	g_state->hOutput = CreateWindowExW(
		0, L"RICHEDIT50W", L"", WS_CHILD | WS_VISIBLE | WS_BORDER |
		ES_MULTILINE | ES_AUTOVSCROLL | WS_VSCROLL | ES_READONLY,
		OUTPUT_MARGIN_LEFT, OUTPUT_MARGIN_TOP,
		MAIN_WINDOW_WIDTH - OUTPUT_MARGIN_LEFT - OUTPUT_MARGIN_RIGHT,
		MAIN_WINDOW_HEIGHT - OUTPUT_MARGIN_TOP - OUTPUT_MARGIN_BOTTOM,
		g_state->hMainWnd, nullptr, hInst, nullptr);
	UpdateLogActionButtonsEnabled();
	UpdateHelpButtonEnabled();
	// On startup, load MapPackSyncTool.txt (if present) into the log.
	LoadHelpTextIntoOutput(true, false);
	// --------------------------------------------------
	// Tooltips
	// --------------------------------------------------
	g_state->hTooltip = CreateWindowExW(
		WS_EX_TOPMOST,
		TOOLTIPS_CLASSW,
		nullptr,
		WS_POPUP | TTS_ALWAYSTIP | TTS_NOPREFIX,
		CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT, CW_USEDEFAULT,
		g_state->hMainWnd,
		nullptr,
		hInst,
		nullptr);
	if (g_state->hTooltip)
	{
		SetWindowPos(g_state->hTooltip, HWND_TOPMOST, 0, 0, 0, 0,
			SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
		SendMessageW(g_state->hTooltip, TTM_SETMAXTIPWIDTH, 0, 420);
		SendMessageW(g_state->hTooltip, TTM_SETDELAYTIME, TTDT_INITIAL, 500);
		// reduce hang-time a bit
		SendMessageW(g_state->hTooltip, TTM_SETDELAYTIME, TTDT_AUTOPOP, 4000);
		SendMessageW(g_state->hTooltip, TTM_ACTIVATE, TRUE, 0);
		AddTooltip(g_state->hTooltip, g_state->hBrowseBtn, L"Browse for your Istaria base install folder");
		AddTooltip(g_state->hTooltip, g_state->hRunButton, L"Download/Update/Sync/Install MapPack 5.0");
		AddTooltip(g_state->hTooltip, g_state->hCancelBtn, L"Cancel Sync.");
		AddTooltip(g_state->hTooltip, g_state->hDeleteBtn, L"Remove/Uninstall MapPack (New or Older versions)");
		AddTooltip(g_state->hTooltip, g_state->hFolderEdit, L"Path to your Istaria base install folder");
		AddTooltip(g_state->hTooltip, g_state->hCopyLogBtn, L"Copy Log to the clipboard");
		AddTooltip(g_state->hTooltip, g_state->hSaveLogBtn, L"Save Log to a .txt file");
		AddTooltip(g_state->hTooltip, g_state->hHelpBtn, L"Reload Help (Also displays upon startup)");
		AddTooltip(g_state->hTooltip, g_state->hCheckUpdatesBtn, L"Check for updates of MapPack Sync Tool");
	}
	LayoutMainWindow(g_state->hMainWnd, g_state);
	SendMessageW(g_state->hOutput, EM_EXLIMITTEXT, 0, (LPARAM)(8ULL * 1024ULL * 1024ULL));
	MSG msg;
	while (GetMessage(&msg, nullptr, 0, 0))
	{
		if (g_state && g_state->hTooltip)
			SendMessageW(g_state->hTooltip, TTM_RELAYEVENT, 0, (LPARAM)&msg);
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}
	// WinHTTP does not require global cleanup.
	ReleaseSingleInstanceMutex();
	return 0;
}
