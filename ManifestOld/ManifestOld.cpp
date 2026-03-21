// ManifestOld.cpp
// Generates a JSON manifest for files under "resources_override" next to this EXE.
// This variant does NOT compute SHA-256 (all BCrypt / hashing removed).
//
// Output file: mappack_manifest_old.json
//
// Notes:
// - This file intentionally avoids std::filesystem so it builds even if the project
//   is not set to C++17.
// - JSON output is pretty-printed and valid (no trailing commas).

#define WIN32_LEAN_AND_MEAN
#include <windows.h>

#include <algorithm>
#include <cstdint>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>

// ----------------------------
// Small helpers
// ----------------------------

static std::wstring GetExeDirectoryW()
{
    wchar_t buf[MAX_PATH] = {};
    DWORD n = GetModuleFileNameW(nullptr, buf, MAX_PATH);
    if (n == 0 || n >= MAX_PATH)
        return L".";

    std::wstring full(buf, n);
    size_t pos = full.find_last_of(L"\\/");
    if (pos == std::wstring::npos)
        return L".";
    return full.substr(0, pos);
}

static bool DirectoryExistsW(const std::wstring& path)
{
    DWORD attrs = GetFileAttributesW(path.c_str());
    return (attrs != INVALID_FILE_ATTRIBUTES) && (attrs & FILE_ATTRIBUTE_DIRECTORY);
}

static void EnsureTrailingBackslash(std::wstring& s)
{
    if (!s.empty() && s.back() != L'\\' && s.back() != L'/')
        s.push_back(L'\\');
}

static std::wstring JoinPathW(std::wstring a, const std::wstring& b)
{
    EnsureTrailingBackslash(a);
    return a + b;
}

static std::wstring ToLowerW(std::wstring s)
{
    std::transform(s.begin(), s.end(), s.begin(), [](wchar_t c) {
        return (wchar_t)CharLowerW((LPWSTR)(uintptr_t)c);
        });
    return s;
}

// UTF-16 -> UTF-8
static std::string WideToUtf8(const std::wstring& w)
{
    if (w.empty()) return {};
    int needed = WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), nullptr, 0, nullptr, nullptr);
    if (needed <= 0) return {};
    std::string out((size_t)needed, '\0');
    WideCharToMultiByte(CP_UTF8, 0, w.c_str(), (int)w.size(), &out[0], needed, nullptr, nullptr);
    return out;
}

static std::string JsonEscape(const std::string& s)
{
    std::ostringstream o;
    for (unsigned char c : s)
    {
        switch (c)
        {
        case '\"': o << "\\\""; break;
        case '\\': o << "\\\\"; break;
        case '\b': o << "\\b"; break;
        case '\f': o << "\\f"; break;
        case '\n': o << "\\n"; break;
        case '\r': o << "\\r"; break;
        case '\t': o << "\\t"; break;
        default:
            if (c < 0x20)
            {
                char buf[7];
                // \u00XX
                wsprintfA(buf, "\\u%04X", (unsigned)c);
                o << buf;
            }
            else
            {
                o << (char)c;
            }
        }
    }
    return o.str();
}

// Replace backslashes with forward slashes for JSON "path"
static std::string NormalizeRelPath(const std::wstring& relW)
{
    std::wstring tmp = relW;
    for (auto& ch : tmp)
    {
        if (ch == L'\\') ch = L'/';
    }
    return WideToUtf8(tmp);
}

// ----------------------------
// File enumeration (recursive)
// ----------------------------

struct FileEntry
{
    std::wstring absPath; // full path on disk
    std::wstring relPath; // relative to resources_override root
};

static void EnumerateFilesRecursive(
    const std::wstring& rootDir,
    const std::wstring& currentDir,
    std::vector<FileEntry>& out)
{
    std::wstring search = JoinPathW(currentDir, L"*");

    WIN32_FIND_DATAW ffd = {};
    HANDLE hFind = FindFirstFileW(search.c_str(), &ffd);
    if (hFind == INVALID_HANDLE_VALUE)
        return;

    do
    {
        const wchar_t* name = ffd.cFileName;
        if (!name || name[0] == L'\0')
            continue;
        if (wcscmp(name, L".") == 0 || wcscmp(name, L"..") == 0)
            continue;

        std::wstring abs = JoinPathW(currentDir, name);

        if (ffd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
        {
            EnumerateFilesRecursive(rootDir, abs, out);
        }
        else
        {
            // rel = abs minus rootDir + separator
            std::wstring root = rootDir;
            EnsureTrailingBackslash(root);

            std::wstring rel;
            if (abs.size() >= root.size() && _wcsnicmp(abs.c_str(), root.c_str(), root.size()) == 0)
                rel = abs.substr(root.size());
            else
                rel = name; // fallback

            out.push_back({ abs, rel });
        }

    } while (FindNextFileW(hFind, &ffd));

    FindClose(hFind);
}

// ----------------------------
// Main
// ----------------------------

int wmain()
{
    const std::wstring exeDir = GetExeDirectoryW();

    const std::wstring resourcesOverride = JoinPathW(exeDir, L"resources_override");
    if (!DirectoryExistsW(resourcesOverride))
    {
        MessageBoxW(nullptr,
            L"Expected a folder named 'resources_override' next to this EXE.\n\n"
            L"Place this EXE in your mappack folder (e.g., mappack3.6) so that:\n"
            L"  <folder>\\resources_override\\\n"
            L"exists, then run again.",
            L"Manifest generator", MB_ICONERROR | MB_OK);
        return 1;
    }
    // Keeping it empty by default.
    const std::string base_url = "";

    std::vector<FileEntry> files;
    EnumerateFilesRecursive(resourcesOverride, resourcesOverride, files);

    // Stable ordering
    std::sort(files.begin(), files.end(), [](const FileEntry& a, const FileEntry& b) {
        return _wcsicmp(a.relPath.c_str(), b.relPath.c_str()) < 0;
        });

    const std::wstring outPathW = JoinPathW(exeDir, L"mappack_manifest_old.json");
    std::ofstream out(WideToUtf8(outPathW), std::ios::binary);
    if (!out)
    {
        MessageBoxW(nullptr, L"Failed to open output file for writing:\n\nmappack_manifest_old.json",
            L"Manifest generator", MB_ICONERROR | MB_OK);
        return 2;
    }

    out << "{\n";
    out << "  \"files\": [\n";

    for (size_t i = 0; i < files.size(); ++i)
    {
        const std::string rel = NormalizeRelPath(files[i].relPath);

        out << "    { \"path\": \"" << JsonEscape(rel) << "\" }";
        if (i + 1 < files.size())
            out << ",";
        out << "\n";
    }

    out << "  ]\n";
    out << "}\n";

    out.close();
    return 0;
}