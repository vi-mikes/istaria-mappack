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
3) "Not-in-manifest" deletions + empty-dir cleanup are intentionally limited in scope (maps folder only).
   Legacy cleanup based on mappack_manifest_old.json may remove listed legacy files.
Notes
- UI remains responsive: sync runs on a worker thread; UI updates use PostMessage.
- This file is intentionally kept as a single translation unit for easy building.
*/

//////////////////////////////////////////////////////
// Comment this out to suppress some local testing messages
#define DEBUG_MESSAGE
//////////////////////////////////////////////////////

#define WIN32_LEAN_AND_MEAN
#include <windows.h>
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
#include <functional>  // std::function
#include <memory>      // std::unique_ptr
#include <winhttp.h>
#pragma comment(lib, "winhttp.lib")
#pragma comment(lib, "User32.lib")
#pragma comment(lib, "Comctl32.lib")
#pragma comment(lib, "Comdlg32.lib")
#pragma comment(lib, "Shell32.lib")
namespace fs = std::filesystem;

// --------------------------------------------------
// App identity (window title)
// --------------------------------------------------
#define MAP_PACK_SYNC_TOOL_NAME    L"MapPack Sync Tool"
#define MAP_PACK_SYNC_TOOL_VERSION L"v0.0.1"

static const std::wstring kWindowTitle =
std::wstring(MAP_PACK_SYNC_TOOL_NAME) + L" " + MAP_PACK_SYNC_TOOL_VERSION;
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
// Remote config (manifest + root)
// --------------------------------------------------
static constexpr const char* kRemoteHost = "https://istaria-mappack.s3.us-west-2.amazonaws.com";
static constexpr const char* kRemoteRootPath = "/resources_override/";
static constexpr const char* kManifestPath = "/mappack_manifest.json";
static constexpr const char* kManifestOldPath = "/mappack_manifest_old.json";
static constexpr const wchar_t* kUpdateExeUrl = L"https://istaria-mappack.s3.us-west-2.amazonaws.com/MapPackSyncTool.exe";
static constexpr const wchar_t* kUpdateVersionUrl = L"https://istaria-mappack.s3.us-west-2.amazonaws.com/version.txt";
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
};
static AppState* g_state = nullptr;
static HANDLE g_hSingleInstanceMutex = nullptr;
// Remember the last directory used by the Save Log dialog.
static std::wstring g_lastSaveDir;
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
	int progTextY = progY + 18;
	int saveLogX = right - btnW;
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
	HDWP hdwp = BeginDeferWindowPos(12);
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

static bool OutputHasNonWhitespace()
{
	const std::wstring t = GetOutputTextW();
	for (wchar_t c : t)
	{
		if (!std::iswspace(c))
			return true;
	}
	return false;
}

static void UpdateLogActionButtonsEnabled()
{
	if (!g_state) return;
	const BOOL enable = OutputHasNonWhitespace() ? TRUE : FALSE;
	if (g_state->hCopyLogBtn) EnableWindow(g_state->hCopyLogBtn, enable);
	if (g_state->hSaveLogBtn) EnableWindow(g_state->hSaveLogBtn, enable);
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
	wchar_t* w = DupWideForPostUtf8(textUtf8);
	if (!w) return;
	PostHeapMessageW(WM_APP_LOG, w);
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
	wchar_t* w = DupWideForPostW(textWide);
	if (!w) return;
	PostHeapMessageW(WM_APP_PROGRESS_TEXT, w);
}
static void PostProgressMarqueeOn()
{
	PostMessageW(g_state->hMainWnd, WM_APP_PROGRESS_MARQ_ON, 0, 0);
}
static void PostProgressMarqueeOff()
{
	PostMessageW(g_state->hMainWnd, WM_APP_PROGRESS_MARQ_OFF, 0, 0);
}
static void PostProgressInit(size_t total)
{
	PostMessageW(g_state->hMainWnd, WM_APP_PROGRESS_INIT, (WPARAM)total, 0);
}
static void PostProgressSet(size_t pos)
{
	PostMessageW(g_state->hMainWnd, WM_APP_PROGRESS_SET, (WPARAM)pos, 0);
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
	fs::path localDeleteRoot;
	std::vector<std::string> errors;
};
struct SyncConfig
{
	std::string remoteHost;
	std::string remoteRootPath;   // NOTE: must end with '/' because JoinUrl/manifest normalization assume it.
	std::string manifestUrl;
	fs::path localBase;
	fs::path localSyncRoot;
	fs::path localDeleteRoot;
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
	r.localDeleteRoot = r.localSyncRoot / "resources" / "interface" / "maps";
	r.ok = true;
	return r;
}
// --------------------------------------------------
// SHA-256 (Windows CNG / BCrypt)
// --------------------------------------------------
static bool Sha256FileHexLower(const fs::path& filePath, std::string& outHex)
{
	outHex.clear();
	BCRYPT_ALG_HANDLE hAlg = nullptr;
	BCRYPT_HASH_HANDLE hHash = nullptr;
	NTSTATUS st = BCryptOpenAlgorithmProvider(&hAlg, BCRYPT_SHA256_ALGORITHM, nullptr, 0);
	if (st != 0) return false;
	DWORD objLen = 0, cbData = 0, hashLen = 0;
	st = BCryptGetProperty(hAlg, BCRYPT_OBJECT_LENGTH, (PUCHAR)&objLen, sizeof(objLen), &cbData, 0);
	if (st != 0) { BCryptCloseAlgorithmProvider(hAlg, 0); return false; }
	st = BCryptGetProperty(hAlg, BCRYPT_HASH_LENGTH, (PUCHAR)&hashLen, sizeof(hashLen), &cbData, 0);
	if (st != 0) { BCryptCloseAlgorithmProvider(hAlg, 0); return false; }
	std::vector<unsigned char> hashObj(objLen);
	std::vector<unsigned char> hash(hashLen);
	st = BCryptCreateHash(hAlg, &hHash, hashObj.data(), (ULONG)hashObj.size(), nullptr, 0, 0);
	if (st != 0) { BCryptCloseAlgorithmProvider(hAlg, 0); return false; }
	std::ifstream in(filePath, std::ios::binary);
	if (!in)
	{
		BCryptDestroyHash(hHash);
		BCryptCloseAlgorithmProvider(hAlg, 0);
		return false;
	}
	std::vector<unsigned char> buf(1024 * 1024);
	while (in)
	{
		in.read((char*)buf.data(), (std::streamsize)buf.size());
		std::streamsize got = in.gcount();
		if (got > 0)
		{
			st = BCryptHashData(hHash, buf.data(), (ULONG)got, 0);
			if (st != 0)
			{
				BCryptDestroyHash(hHash);
				BCryptCloseAlgorithmProvider(hAlg, 0);
				return false;
			}
		}
	}
	st = BCryptFinishHash(hHash, hash.data(), (ULONG)hash.size(), 0);
	BCryptDestroyHash(hHash);
	BCryptCloseAlgorithmProvider(hAlg, 0);
	if (st != 0) return false;
	static const char* kHex = "0123456789abcdef";
	outHex.resize(hash.size() * 2);
	for (size_t i = 0; i < hash.size(); ++i)
	{
		outHex[i * 2 + 0] = kHex[(hash[i] >> 4) & 0xF];
		outHex[i * 2 + 1] = kHex[(hash[i] >> 0) & 0xF];
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

static fs::path MakeDestPath(const fs::path& installRoot, std::string_view manifestRelPath)
{
	const std::string rel = NormalizeManifestRel(manifestRelPath);
	return installRoot / "resources_override" / "mappack" / fs::path(rel);
}
// --------------------------------------------------
// Manifest parsing
// --------------------------------------------------
struct ManifestEntry
{
	std::string remotePath;
	std::string relPath;
	std::string sha256;
};
static bool ReadJsonStringValue(const std::string& s, size_t& i, std::string& out)
{
	out.clear();
	if (i >= s.size() || s[i] != '"') return false;
	++i;
	while (i < s.size())
	{
		char c = s[i++];
		if (c == '"') return true;
		if (c == '\\')
		{
			if (i >= s.size()) return false;
			char e = s[i++];
			switch (e)
			{
			case '"': out.push_back('"'); break;
			case '\\': out.push_back('\\'); break;
			case '/': out.push_back('/'); break;
			case 'b': out.push_back('\b'); break;
			case 'f': out.push_back('\f'); break;
			case 'n': out.push_back('\n'); break;
			case 'r': out.push_back('\r'); break;
			case 't': out.push_back('	'); break;
			case 'u':
				return false;
			default:
				return false;
			}
		}
		else
		{
			out.push_back(c);
		}
	}
	return false;
}
static void SkipWs(const std::string& s, size_t& i)
{
	while (i < s.size())
	{
		char c = s[i];
		if (c == ' ' || c == '	' || c == '\r' || c == '\n') { ++i; continue; }
		break;
	}
}
static bool ParseManifestSha256(const std::string& jsonText, std::string& outBaseUrl, std::vector<ManifestEntry>& outFiles)
{
	outBaseUrl.clear();
	outFiles.clear();
	const std::string& s = jsonText;
	size_t i = 0;
	size_t posBase = s.find("\"base_url\"");
	if (posBase == std::string::npos) return false;
	i = posBase + strlen("\"base_url\"");
	SkipWs(s, i);
	if (i >= s.size() || s[i] != ':') return false;
	++i;
	SkipWs(s, i);
	std::string base;
	if (!ReadJsonStringValue(s, i, base)) return false;
	outBaseUrl = base;
	size_t posFiles = s.find("\"files\"");
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
		if (s[i] != '{') { return false; }
		++i;
		std::string path, digest;
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
			else if (key == "sha256" || key == "hash") digest = val;
		}
		if (!path.empty() && !digest.empty())
		{
			ManifestEntry e;
			e.remotePath = path;
			e.sha256 = digest;
			outFiles.push_back(e);
		}
	}
	return !outBaseUrl.empty() && !outFiles.empty();
}
// --------------------------------------------------
// WinHTTP (no redirects; treat redirects as errors)
// --------------------------------------------------
namespace AppConstants
{
	static constexpr long kManifestConnectTimeoutSec = 15L;
	static constexpr long kManifestTimeoutSec = 120L;
	static constexpr long kFileConnectTimeoutMs = 15000L;
	static constexpr long kFileTimeoutMs = 0L;
	static constexpr const char* kUserAgent = "MapPackSyncTool by Cegaiel";
	static constexpr const wchar_t* kUserAgentW = L"MapPackSyncTool by Cegaiel";
}
struct WinHttpHandleDeleter { void operator()(HINTERNET h) const noexcept { if (h) WinHttpCloseHandle(h); } };
using WinHttpHandle = std::unique_ptr<void, WinHttpHandleDeleter>;

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

static bool WinHttpGetToString_NoRedirects(
	const std::string& urlUtf8,
	std::string& outBody,
	const CancelToken& cancel,
	long connectTimeoutMs,
	long totalTimeoutMs,
	std::string* outErr,
	long* outHttp)
{
	outBody.clear();
	if (outErr) outErr->clear();
	if (outHttp) *outHttp = 0;

	std::wstring host, path;
	INTERNET_PORT port = 0;
	bool secure = false;
	if (!CrackUrlWinHttp(urlUtf8, host, path, port, secure, outErr))
		return false;

	WinHttpHandle hSession(WinHttpOpen(Utf8ToWide(AppConstants::kUserAgent).c_str(), WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0));
	if (!hSession)
	{
		if (outErr) *outErr = "WinHttpOpen failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}
	WinHttpHandle hConnect(WinHttpConnect((HINTERNET)hSession.get(), host.c_str(), port, 0));
	if (!hConnect)
	{
		if (outErr) *outErr = "WinHttpConnect failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	DWORD flags = secure ? WINHTTP_FLAG_SECURE : 0;
	WinHttpHandle hRequest(WinHttpOpenRequest((HINTERNET)hConnect.get(), L"GET", path.c_str(), nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags));
	if (!hRequest)
	{
		if (outErr) *outErr = "WinHttpOpenRequest failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	// Treat redirects as errors.
	DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_NEVER;
	(void)WinHttpSetOption((HINTERNET)hRequest.get(), WINHTTP_OPTION_REDIRECT_POLICY, &redirectPolicy, sizeof(redirectPolicy));

	// Timeouts: resolve, connect, send, receive.
	int resolveMs = connectTimeoutMs;
	int connectMs = connectTimeoutMs;
	int sendMs = connectTimeoutMs;
	int recvMs = (totalTimeoutMs <= 0) ? 0 : totalTimeoutMs;
	(void)WinHttpSetTimeouts((HINTERNET)hRequest.get(), resolveMs, connectMs, sendMs, recvMs);

	if (cancel.IsCanceled()) { if (outErr) *outErr = "Canceled"; return false; }
	if (!WinHttpSendRequest((HINTERNET)hRequest.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0))
	{
		if (outErr) *outErr = "WinHttpSendRequest failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}
	if (!WinHttpReceiveResponse((HINTERNET)hRequest.get(), nullptr))
	{
		if (outErr) *outErr = "WinHttpReceiveResponse failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	DWORD status = 0;
	DWORD statusSize = sizeof(status);
	WinHttpQueryHeaders((HINTERNET)hRequest.get(), WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize, WINHTTP_NO_HEADER_INDEX);
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

	std::string body;
	for (;;)
	{
		if (cancel.IsCanceled()) { if (outErr) *outErr = "Canceled"; return false; }
		DWORD avail = 0;
		if (!WinHttpQueryDataAvailable((HINTERNET)hRequest.get(), &avail))
		{
			if (outErr) *outErr = "WinHttpQueryDataAvailable failed (" + std::to_string(GetLastError()) + ")";
			return false;
		}
		if (avail == 0) break;
		std::vector<char> buf(avail);
		DWORD read = 0;
		if (!WinHttpReadData((HINTERNET)hRequest.get(), buf.data(), avail, &read))
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

	std::wstring host, path;
	INTERNET_PORT port = 0;
	bool secure = false;
	if (!CrackUrlWinHttp(urlUtf8, host, path, port, secure, outErr))
		return false;

	WinHttpHandle hSession(WinHttpOpen(Utf8ToWide(AppConstants::kUserAgent).c_str(), WINHTTP_ACCESS_TYPE_DEFAULT_PROXY, WINHTTP_NO_PROXY_NAME, WINHTTP_NO_PROXY_BYPASS, 0));
	if (!hSession)
	{
		if (outErr) *outErr = "WinHttpOpen failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}
	WinHttpHandle hConnect(WinHttpConnect((HINTERNET)hSession.get(), host.c_str(), port, 0));
	if (!hConnect)
	{
		if (outErr) *outErr = "WinHttpConnect failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	DWORD flags = secure ? WINHTTP_FLAG_SECURE : 0;
	WinHttpHandle hRequest(WinHttpOpenRequest((HINTERNET)hConnect.get(), L"GET", path.c_str(), nullptr, WINHTTP_NO_REFERER, WINHTTP_DEFAULT_ACCEPT_TYPES, flags));
	if (!hRequest)
	{
		if (outErr) *outErr = "WinHttpOpenRequest failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	// Treat redirects as errors.
	DWORD redirectPolicy = WINHTTP_OPTION_REDIRECT_POLICY_NEVER;
	(void)WinHttpSetOption((HINTERNET)hRequest.get(), WINHTTP_OPTION_REDIRECT_POLICY, &redirectPolicy, sizeof(redirectPolicy));

	int resolveMs = connectTimeoutMs;
	int connectMs = connectTimeoutMs;
	int sendMs = connectTimeoutMs;
	int recvMs = (totalTimeoutMs <= 0) ? 0 : totalTimeoutMs;
	(void)WinHttpSetTimeouts((HINTERNET)hRequest.get(), resolveMs, connectMs, sendMs, recvMs);

	if (cancel.IsCanceled()) { if (outErr) *outErr = "Canceled"; return false; }
	if (!WinHttpSendRequest((HINTERNET)hRequest.get(), WINHTTP_NO_ADDITIONAL_HEADERS, 0, WINHTTP_NO_REQUEST_DATA, 0, 0, 0))
	{
		if (outErr) *outErr = "WinHttpSendRequest failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}
	if (!WinHttpReceiveResponse((HINTERNET)hRequest.get(), nullptr))
	{
		if (outErr) *outErr = "WinHttpReceiveResponse failed (" + std::to_string(GetLastError()) + ")";
		return false;
	}

	DWORD status = 0;
	DWORD statusSize = sizeof(status);
	WinHttpQueryHeaders((HINTERNET)hRequest.get(), WINHTTP_QUERY_STATUS_CODE | WINHTTP_QUERY_FLAG_NUMBER, WINHTTP_HEADER_NAME_BY_INDEX, &status, &statusSize, WINHTTP_NO_HEADER_INDEX);
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

	for (;;)
	{
		if (cancel.IsCanceled()) { if (outErr) *outErr = "Canceled"; return false; }
		DWORD avail = 0;
		if (!WinHttpQueryDataAvailable((HINTERNET)hRequest.get(), &avail))
		{
			if (outErr) *outErr = "WinHttpQueryDataAvailable failed (" + std::to_string(GetLastError()) + ")";
			return false;
		}
		if (avail == 0) break;
		std::vector<unsigned char> buf(avail);
		DWORD read = 0;
		if (!WinHttpReadData((HINTERNET)hRequest.get(), buf.data(), avail, &read))
		{
			if (outErr) *outErr = "WinHttpReadData failed (" + std::to_string(GetLastError()) + ")";
			return false;
		}
		if (read == 0) break;
		if (ctx.f)
		{
			const size_t wrote = fwrite(buf.data(), 1, read, ctx.f);
			if (wrote != read)
			{
				if (outErr) *outErr = "Failed writing to temp file";
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
static int RemoveEmptyDirsBottomUp(const fs::path& root)
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
	std::string baseUrl;
	std::vector<ManifestEntry> workList;
	std::unordered_set<std::string> manifestRelSet;
};
static std::string MakeFileUrlFromBase(const std::string& base, const std::string& remotePath)
{
	std::string b = base;
	if (!b.empty() && b.back() != '/') b.push_back('/');
	if (!remotePath.empty() && remotePath.front() == '/')
		return b + remotePath.substr(1);
	return b + remotePath;
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
	std::vector<ManifestEntry> manifestFiles;
	if (!ParseManifestSha256(manifestText, out.baseUrl, manifestFiles))
	{
		PostProgressMarqueeOff();
		outErr = "Manifest parse failed.";
		return false;
	}
	PostProgressMarqueeOff();
	out.manifestRelSet.reserve(manifestFiles.size() * 2);
	out.workList = std::move(manifestFiles);
	for (auto& e : out.workList)
	{
		std::string remote = NormalizePathGeneric(e.remotePath);
		std::string rel = NormalizeManifestRel(remote);
		e.remotePath = remote;
		e.relPath = rel;
		if (!rel.empty()) out.manifestRelSet.insert(rel);
	}
	std::sort(out.workList.begin(), out.workList.end(),
		[](const ManifestEntry& a, const ManifestEntry& b) { return a.relPath < b.relPath; });
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
	Log("Parsing local files: Searching for any local files (.def/.png) that exist but are not in the manifest; Those need deleted...\r\n");
	std::vector<fs::path> localFiles;
	// Restrict deletions to: resources_override/mappack/resources/interface/maps/
	if (fs::exists(cfg.localDeleteRoot))
	{
		for (auto& entry : fs::recursive_directory_iterator(cfg.localDeleteRoot))
		{
			if (cancel.IsCanceled()) return;
			if (!entry.is_regular_file()) continue;
			std::string extLower = ToLowerAscii(entry.path().extension().string());
			if (!IsSyncedExt(extLower)) continue;
			localFiles.push_back(entry.path());
		}
	}
	else
	{
		Log("NOTE: maps folder not found; nothing to delete.\r\n");
		Log("Expected: " + PathToUtf8(cfg.localDeleteRoot) + "\r\n");
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
	Log("Parsing manifest: Searching for any local files that are missing or has changed...\r\n");
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
		std::string fileUrl = MakeFileUrlFromBase(md.baseUrl, remotePath);
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
	Log("\r\nSyncing MapPack Root:  " + PathToUtf8(cfg.localSyncRoot) + "\r\n");
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
	Log("Removing empty directories (maps folder only)...\r\n");
	const int removedDirs = RemoveEmptyDirsBottomUp(cfg.localDeleteRoot);

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
	Log("Removing empty sub-directories from MapPack 5.0 (resources_override/mappack)...\r\n");

	const int removedDirs = RemoveEmptyDirsBottomUp(cfg.localDeleteRoot);
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
		PostMessageW(g_state->hMainWnd, WM_APP_WORKER_DONE, 0, 0);
		return 0;
	}
	SyncConfig cfg;
	cfg.remoteHost = kRemoteHost;
	cfg.remoteRootPath = kRemoteRootPath;
	cfg.manifestUrl = JoinUrl(kRemoteHost, kManifestPath);
	cfg.localBase = pf.localBase;
	cfg.localSyncRoot = pf.localSyncRoot;
	cfg.localDeleteRoot = pf.localDeleteRoot;
	CancelToken cancel{ &g_state->cancelRequested };
	RunSync(cfg, cancel);
	g_state->isRunning.store(false);
	PostMessageW(g_state->hMainWnd, WM_APP_WORKER_DONE, 0, 0);
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
	EnableWindow(g_state->hBrowseBtn, FALSE);
	EnableWindow(g_state->hRunButton, FALSE);
	if (g_state->hCancelBtn) EnableWindow(g_state->hCancelBtn, TRUE);
	unsigned int tid = 0;
	unique_handle worker(reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, &WorkerThreadProc, nullptr, 0, &tid)));
	g_state->hWorkerThread = worker.release();
	if (!g_state->hWorkerThread)
	{
		g_state->isRunning.store(false);
		EnableWindow(g_state->hBrowseBtn, TRUE);
		EnableWindow(g_state->hRunButton, TRUE);
		if (g_state->hCancelBtn) EnableWindow(g_state->hCancelBtn, FALSE);
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
		PostMessageW(g_state->hMainWnd, WM_APP_WORKER_DONE, 0, 0);
		return 0;
	}

	SyncConfig cfg;
	cfg.remoteHost = kRemoteHost;
	cfg.remoteRootPath = kRemoteRootPath;
	cfg.manifestUrl = JoinUrl(kRemoteHost, kManifestPath);
	cfg.localBase = pf.localBase;
	cfg.localSyncRoot = pf.localSyncRoot;
	cfg.localDeleteRoot = pf.localDeleteRoot;

	CancelToken cancel{ &g_state->cancelRequested };
	RemoveMapPackFiles(cfg, cancel);

	g_state->isRunning.store(false);
	PostMessageW(g_state->hMainWnd, WM_APP_WORKER_DONE, 0, 0);
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

	EnableWindow(g_state->hBrowseBtn, FALSE);
	EnableWindow(g_state->hRunButton, FALSE);
	if (g_state->hDeleteBtn) EnableWindow(g_state->hDeleteBtn, FALSE);
	if (g_state->hCancelBtn) EnableWindow(g_state->hCancelBtn, TRUE);

	unsigned int tid = 0;
	unique_handle worker(reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, &WorkerThreadProcRemove, nullptr, 0, &tid)));
	g_state->hWorkerThread = worker.release();
	if (!g_state->hWorkerThread)
	{
		g_state->isRunning.store(false);
		EnableWindow(g_state->hBrowseBtn, TRUE);
		EnableWindow(g_state->hRunButton, TRUE);
		if (g_state->hDeleteBtn) EnableWindow(g_state->hDeleteBtn, TRUE);
		if (g_state->hCancelBtn) EnableWindow(g_state->hCancelBtn, FALSE);
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

static unsigned __stdcall UpdateThreadProc(void* p)
{
	UpdateResult* res = (UpdateResult*)p;
	res->ok = false;
	res->different = false;
	res->err.clear();

	fs::path curExe = GetThisExePath();
	if (curExe.empty()) { res->err = L"Cannot locate current exe"; return 0; }

	// Update availability is decided by version.txt (NOT by hashing the remote exe).
	res->localVersion = MAP_PACK_SYNC_TOOL_VERSION;
	TrimInPlace(res->localVersion);
	std::wstring verErr;
	if (!WinHttpDownloadUrlToWideString(kUpdateVersionUrl, res->remoteVersion, verErr))
	{
		res->err = L"Failed to download version.txt: " + verErr;
		return 0;
	}
	TrimInPlace(res->remoteVersion);
	if (res->remoteVersion.empty())
	{
		res->err = L"version.txt was empty";
		return 0;
	}
	if (_wcsicmp(res->remoteVersion.c_str(), res->localVersion.c_str()) != 0)
		res->different = true;

	if (!res->different)
	{
		res->ok = true;
		return 0;
	}

	// New version available: download the updated exe to a temp file in the SAME directory as the exe so any relaunch uses the same dependency neighborhood.
	fs::path tempExe = curExe.parent_path() / L"MapPackSyncTool.exe.download";
	res->downloadedTemp = tempExe;

	std::wstring dlErr;
	if (!WinHttpDownloadUrlToFile(kUpdateExeUrl, tempExe, dlErr))
	{
		res->err = L"Download failed: " + dlErr;
		return 0;
	}
	res->ok = true;
	return 0;
}

static void StartCheckForUpdates()
{
	AppState* st = g_state;
	if (!st) return;
	if (st->isUpdateRunning.exchange(true)) return;
	if (st->hCheckUpdatesBtn) EnableWindow(st->hCheckUpdatesBtn, FALSE);

	UpdateResult* res = new UpdateResult();
	uintptr_t th = _beginthreadex(nullptr, 0, UpdateThreadProc, res, 0, nullptr);
	if (!th)
	{
		st->isUpdateRunning.store(false);
		if (st->hCheckUpdatesBtn) EnableWindow(st->hCheckUpdatesBtn, TRUE);
		delete res;
		Log("ERROR: Failed to start update thread.\r\n");
		return;
	}
	st->hUpdateThread = (HANDLE)th;

	// Waiter thread: waits for the update thread then posts WM_APP+77 with UpdateResult* back to the UI thread.
	uintptr_t waiter = _beginthreadex(nullptr, 0, [](void* param)->unsigned {
		auto* pair = (std::pair<AppState*, UpdateResult*>*)param;
		WaitForSingleObject(pair->first->hUpdateThread, INFINITE);
		PostMessageW(pair->first->hMainWnd, WM_APP + 77, 0, (LPARAM)pair->second);
		delete pair;
		return 0;
		}, new std::pair<AppState*, UpdateResult*>(st, res), 0, nullptr);
	if (waiter) CloseHandle((HANDLE)waiter);
}
// --------------------------------------------------
// Window Proc
// --------------------------------------------------
LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	AppState* st = (AppState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
	if (!st) st = g_state;
	switch (msg)
	{
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
			EnableWindow(st->hBrowseBtn, TRUE);
			EnableWindow(st->hRunButton, TRUE);
			if (st->hCancelBtn) EnableWindow(st->hCancelBtn, FALSE);
			if (st->hDeleteBtn) EnableWindow(st->hDeleteBtn, TRUE);
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
	{
		AppState* st2 = (AppState*)GetWindowLongPtrW(hwnd, GWLP_USERDATA);
		if (st2 && st2->hOutput)
			LayoutMainWindow(hwnd, st2);
		return 0;
	}
	case WM_GETMINMAXINFO:
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
	case WM_COMMAND:
		if ((HWND)lParam == st->hBrowseBtn)
		{
			std::wstring path;
			if (BrowseForFolder(hwnd, path))
				SetWindowTextW(st->hFolderEdit, path.c_str());
		}
		else if ((HWND)lParam == st->hRunButton)
		{
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
			if (CopyOutputToClipboard())
				MessageBeep(MB_ICONASTERISK); // ding on successful copy
		}
		else if ((HWND)lParam == st->hSaveLogBtn)
		{
			SaveOutputToFile();
		}
		else if ((HWND)lParam == st->hCheckUpdatesBtn)
		{
			StartCheckForUpdates();
		}

		break;
	case WM_APP + 77:
	{
		std::unique_ptr<UpdateResult> res((UpdateResult*)lParam);
		if (st)
		{
			st->isUpdateRunning.store(false);
			if (st->hCheckUpdatesBtn) EnableWindow(st->hCheckUpdatesBtn, TRUE);
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
			std::wstring msg = L"You are already running the latest version.\r\n\r\nLocal: " + res->localVersion + L"\r\nRemote: " + res->remoteVersion;
			MessageBoxW(hwnd, msg.c_str(), L"MapPack Sync Tool", MB_OK | MB_ICONINFORMATION);
			if (!res->downloadedTemp.empty()) { DeleteFileW(res->downloadedTemp.c_str()); }
			return 0;
		}
		std::wstring prompt = L"New version of MapPack Sync Tool is available.\r\n\r\nCurrent version: " + res->localVersion + L"\r\nAvailable version: " + res->remoteVersion + L"\r\n\r\nProceed with the update?";
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
	case WM_CLOSE:
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
	case WM_DESTROY:
		if (st)
		{
			if (st->hTooltip) { DestroyWindow(st->hTooltip); st->hTooltip = nullptr; }
			if (st->hFontUI) { DeleteObject(st->hFontUI); st->hFontUI = nullptr; }
			if (st->hFontMono) { DeleteObject(st->hFontMono); st->hFontMono = nullptr; }
		}
		PostQuitMessage(0);
		break;
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

	// ===== DEBUG DEFAULT PATH =====
	//#if defined(DEBUG_MESSAGE) && defined(_DEBUG)
#if defined(DEBUG_MESSAGE)
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
		L"BUTTON", L"Cancel Sync",
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
		AddTooltip(g_state->hTooltip, g_state->hBrowseBtn, L"Browse for your Istaria base install folder (folder has istaria.exe)");
		AddTooltip(g_state->hTooltip, g_state->hRunButton, L"Download/Update/Sync/Install MapPack");
		AddTooltip(g_state->hTooltip, g_state->hCancelBtn, L"Cancel Sync.");
		AddTooltip(g_state->hTooltip, g_state->hDeleteBtn, L"Remove/Uninstall MapPack (New or Older versions)");
		AddTooltip(g_state->hTooltip, g_state->hFolderEdit, L"Path to your Istaria base install folder");
		AddTooltip(g_state->hTooltip, g_state->hCopyLogBtn, L"Copy the entire log to the clipboard");
		AddTooltip(g_state->hTooltip, g_state->hSaveLogBtn, L"Save the entire log to a .txt file");
		AddTooltip(g_state->hTooltip, g_state->hCheckUpdatesBtn, L"Check for updates");
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