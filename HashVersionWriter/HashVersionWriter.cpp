// HashVersionWriter.cpp
// - Requires MapPackSyncTool.exe
// - Creates version.txt if missing
// - Line 1 = EXE FileVersion
// - Line 2 = SHA-256
// - No blank line accumulation
// - No C++17 required

#include <windows.h>
#include <bcrypt.h>
#include <iostream>
#include <fstream>
#include <vector>
#include <string>

#pragma comment(lib, "bcrypt.lib")
#pragma comment(lib, "version.lib")

static std::string Trim(const std::string& s)
{
    const char* ws = " \t\r\n";
    const auto b = s.find_first_not_of(ws);
    if (b == std::string::npos) return "";
    const auto e = s.find_last_not_of(ws);
    return s.substr(b, e - b + 1);
}

static bool FileExistsW(const wchar_t* path)
{
    DWORD attrs = GetFileAttributesW(path);
    if (attrs == INVALID_FILE_ATTRIBUTES) return false;
    return (attrs & FILE_ATTRIBUTE_DIRECTORY) == 0;
}

static bool GetExeFileVersion(const std::wstring& exePath, std::string& outVersion)
{
    DWORD handle = 0;
    DWORD size = GetFileVersionInfoSizeW(exePath.c_str(), &handle);
    if (size == 0)
        return false;

    std::vector<BYTE> buf(size);
    if (!GetFileVersionInfoW(exePath.c_str(), 0, size, buf.data()))
        return false;

    VS_FIXEDFILEINFO* ffi = nullptr;
    UINT ffiLen = 0;
    if (!VerQueryValueW(buf.data(), L"\\", reinterpret_cast<void**>(&ffi), &ffiLen) || !ffi)
        return false;

    WORD major = HIWORD(ffi->dwFileVersionMS);
    WORD minor = LOWORD(ffi->dwFileVersionMS);
    WORD build = HIWORD(ffi->dwFileVersionLS);
    WORD revision = LOWORD(ffi->dwFileVersionLS);

    if (revision == 0)
        outVersion = std::to_string(major) + "." +
        std::to_string(minor) + "." +
        std::to_string(build);
    else
        outVersion = std::to_string(major) + "." +
        std::to_string(minor) + "." +
        std::to_string(build) + "." +
        std::to_string(revision);

    return true;
}

static bool ComputeSHA256(const std::wstring& filePath, std::string& outHash)
{
    BCRYPT_ALG_HANDLE hAlg = nullptr;
    BCRYPT_HASH_HANDLE hHash = nullptr;
    DWORD hashObjectSize = 0, dataSize = 0;
    DWORD hashLen = 0;

    if (BCryptOpenAlgorithmProvider(&hAlg, BCRYPT_SHA256_ALGORITHM, nullptr, 0) != 0)
        return false;

    if (BCryptGetProperty(hAlg, BCRYPT_OBJECT_LENGTH,
        (PUCHAR)&hashObjectSize, sizeof(DWORD), &dataSize, 0) != 0)
    {
        BCryptCloseAlgorithmProvider(hAlg, 0);
        return false;
    }

    if (BCryptGetProperty(hAlg, BCRYPT_HASH_LENGTH,
        (PUCHAR)&hashLen, sizeof(DWORD), &dataSize, 0) != 0)
    {
        BCryptCloseAlgorithmProvider(hAlg, 0);
        return false;
    }

    std::vector<BYTE> hashObject(hashObjectSize);
    std::vector<BYTE> hash(hashLen);

    if (BCryptCreateHash(hAlg, &hHash, hashObject.data(),
        hashObjectSize, nullptr, 0, 0) != 0)
    {
        BCryptCloseAlgorithmProvider(hAlg, 0);
        return false;
    }

    std::ifstream file(filePath, std::ios::binary);
    if (!file)
        return false;

    std::vector<char> buffer(64 * 1024);
    while (file.good())
    {
        file.read(buffer.data(), buffer.size());
        std::streamsize read = file.gcount();
        if (read > 0)
        {
            if (BCryptHashData(hHash,
                (PUCHAR)buffer.data(),
                (ULONG)read, 0) != 0)
            {
                BCryptDestroyHash(hHash);
                BCryptCloseAlgorithmProvider(hAlg, 0);
                return false;
            }
        }
    }

    if (BCryptFinishHash(hHash, hash.data(), hashLen, 0) != 0)
    {
        BCryptDestroyHash(hHash);
        BCryptCloseAlgorithmProvider(hAlg, 0);
        return false;
    }

    BCryptDestroyHash(hHash);
    BCryptCloseAlgorithmProvider(hAlg, 0);

    static const char hex[] = "0123456789abcdef";
    outHash.clear();
    for (BYTE b : hash)
    {
        outHash.push_back(hex[b >> 4]);
        outHash.push_back(hex[b & 0xF]);
    }

    return true;
}

int main()
{
    const std::wstring exePath = L"MapPackSyncTool.exe";
    const char* versionFile = "version.txt";

    if (!FileExistsW(exePath.c_str()))
    {
        std::cerr << "ERROR: MapPackSyncTool.exe not found.\n";
        return 1;
    }

    std::string version;
    if (!GetExeFileVersion(exePath, version))
    {
        std::cerr << "ERROR: Could not extract FileVersion.\n";
        return 1;
    }

    std::string sha;
    if (!ComputeSHA256(exePath, sha))
    {
        std::cerr << "ERROR: Failed to compute SHA-256.\n";
        return 1;
    }

    version = Trim(version);
    sha = Trim(sha);

    // Read existing lines if file exists
    std::vector<std::string> lines;
    std::ifstream in(versionFile);
    if (in)
    {
        std::string line;
        while (std::getline(in, line))
            lines.push_back(line);
        in.close();
    }

    // Remove trailing blank lines
    while (!lines.empty() && Trim(lines.back()).empty())
        lines.pop_back();

    if (lines.size() < 2)
        lines.resize(2);

    lines[0] = version;  // overwrite line 1
    lines[1] = sha;      // overwrite line 2

    std::ofstream out(versionFile, std::ios::trunc);
    if (!out)
    {
        std::cerr << "ERROR: Unable to write version.txt\n";
        return 1;
    }

    for (size_t i = 0; i < lines.size(); ++i)
    {
        out << lines[i];
        if (i + 1 < lines.size())
            out << "\r\n";
    }

    out.close();

    std::cout << "version.txt updated successfully.\n";
    std::cout << "Version: " << version << "\n";
    std::cout << "SHA-256: " << sha << "\n";

    return 0;
}
