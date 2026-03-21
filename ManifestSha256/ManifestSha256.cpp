#define NOMINMAX
#include <windows.h>
#include <bcrypt.h>
#pragma comment(lib, "bcrypt.lib")

#include <filesystem>
#include <fstream>
#include <iostream>
#include <sstream>
#include <string>
#include <vector>
#include <algorithm>

namespace fs = std::filesystem;

static bool is_target_file(const fs::path& p) {
    auto ext = p.extension().wstring();
    std::transform(ext.begin(), ext.end(), ext.begin(), ::towlower);
    return (ext == L".def" || ext == L".png");
}

static std::string to_utf8(const std::wstring& ws) {
    if (ws.empty()) return {};
    int n = WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), nullptr, 0, nullptr, nullptr);
    std::string out(n, '\0');
    WideCharToMultiByte(CP_UTF8, 0, ws.c_str(), (int)ws.size(), out.data(), n, nullptr, nullptr);
    return out;
}

static std::string json_escape(const std::string& s) {
    std::string out;
    out.reserve(s.size() + 16);
    for (char c : s) {
        switch (c) {
        case '\\': out += "\\\\"; break;
        case '"':  out += "\\\""; break;
        case '\r': out += "\\r"; break;
        case '\n': out += "\\n"; break;
        case '\t': out += "\\t"; break;
        default:   out += c; break;
        }
    }
    return out;
}

static std::string bytes_to_hex_lower(const std::vector<unsigned char>& bytes) {
    static const char* kHex = "0123456789abcdef";
    std::string s(bytes.size() * 2, '\0');
    for (size_t i = 0; i < bytes.size(); ++i) {
        s[i * 2 + 0] = kHex[(bytes[i] >> 4) & 0xF];
        s[i * 2 + 1] = kHex[(bytes[i] >> 0) & 0xF];
    }
    return s;
}

static bool sha256_file_bcrypt(const fs::path& filePath, std::string& outHex) {
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
    if (!in) {
        BCryptDestroyHash(hHash);
        BCryptCloseAlgorithmProvider(hAlg, 0);
        return false;
    }

    std::vector<unsigned char> buf(1024 * 1024);
    while (in) {
        in.read((char*)buf.data(), (std::streamsize)buf.size());
        std::streamsize got = in.gcount();
        if (got > 0) {
            st = BCryptHashData(hHash, buf.data(), (ULONG)got, 0);
            if (st != 0) {
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

    outHex = bytes_to_hex_lower(hash);
    return true;
}

static fs::path get_exe_directory() {
    wchar_t buf[MAX_PATH]{};
    DWORD n = GetModuleFileNameW(nullptr, buf, MAX_PATH);
    fs::path p(buf);
    return p.parent_path();
}

int wmain() {
    // Run-from-here behavior:
    // - Assume exe is located in the mappack3.6 folder (or run from it).
    // - Use exe directory as "root".
    fs::path baseDir = get_exe_directory();

    // Required input folder:
    fs::path resourcesOverride = baseDir / L"resources_override";

    // Output file in baseDir:
    fs::path outJson = baseDir / L"mappack_manifest.json";

    std::error_code ec;
    if (!fs::exists(resourcesOverride, ec) || !fs::is_directory(resourcesOverride, ec)) {
        std::wcerr << L"ERROR: Expected folder not found:\n  " << resourcesOverride.wstring() << L"\n";
        std::wcerr << L"Place this EXE in your mappack3.6 folder (next to resources_override).\n";
        return 2;
    }

    struct Entry { std::string path; std::string hash; };
    std::vector<Entry> entries;

    // We want "resources_override/..." paths in manifest
    const std::string prefix = "resources_override/";

    for (auto it = fs::recursive_directory_iterator(resourcesOverride, fs::directory_options::skip_permission_denied, ec);
        it != fs::recursive_directory_iterator(); it.increment(ec)) {

        if (ec) { ec.clear(); continue; }
        if (!it->is_regular_file(ec) || ec) { ec.clear(); continue; }

        const fs::path p = it->path();
        if (!is_target_file(p)) continue;

        fs::path rel = fs::relative(p, resourcesOverride, ec);
        if (ec) { ec.clear(); continue; }

        std::string relUtf8 = to_utf8(rel.generic_wstring()); // forward slashes
        std::string manifestPath = prefix + relUtf8;

        std::string h;
        if (!sha256_file_bcrypt(p, h)) {
            std::wcerr << L"ERROR: sha256 failed for:\n  " << p.wstring() << L"\n";
            return 1; // fail fast
        }

        entries.push_back({ manifestPath, h });
    }

    // Deterministic sort
    std::sort(entries.begin(), entries.end(), [](const Entry& a, const Entry& b) {
        return a.path < b.path;
        });

    std::ostringstream json;
    json << "{\n";
    json << "  \"files\": [\n";
    for (size_t i = 0; i < entries.size(); ++i) {
        json << "    {\n";
        json << "      \"path\": \"" << json_escape(entries[i].path) << "\",\n";
        json << "      \"sha256\": \"" << entries[i].hash << "\"\n";
        json << "    }";
        if (i + 1 < entries.size()) json << ",";
        json << "\n";
    }
    json << "  ]\n";
    json << "}\n";

    const std::string outStr = json.str();

    std::ofstream out(outJson, std::ios::binary | std::ios::trunc);
    if (!out) {
        std::wcerr << L"ERROR: Cannot write output file:\n  " << outJson.wstring() << L"\n";
        return 2;
    }
    out.write(outStr.data(), (std::streamsize)outStr.size());
    out.close();

    std::wcout << L"Wrote " << entries.size() << L" entries to:\n  " << outJson.wstring() << L"\n";
    return 0;
}
