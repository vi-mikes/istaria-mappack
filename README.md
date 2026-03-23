# MapPack Sync Tool

A lightweight Windows utility for synchronizing Istaria map pack resources with a remote manifest.

- Portable (no installer required)
- No external dependencies
- Manifest-driven synchronization
- Secure update validation

---

## 📦 Features

- Sync local map pack resources with remote S3-hosted manifest
- Add / Remove / Sync operations
- Progress tracking with percentage display
- Built-in logging (copy/save)
- Automatic update checking
- Elevation detection (Administrator Mode indicator)
- INI-based configuration persistence

---

## 🖥️ Supported Operating Systems

**Minimum supported OS:** Windows Vista / Windows Server 2008

### ✅ Fully Supported
- Windows 11 (x64 or x86)
- Windows 10 (x64 or x86)

### ⚠️ Conditional / Legacy Support
- Windows 7 SP1 (best-effort support)

Requirements:
- TLS 1.2 must be enabled for WinHTTP
- Updated root certificates must be installed

If downloads fail:
- Ensure Windows 7 Service Pack 1 is installed
- Install all available Windows Updates

### ❌ Not Supported
- Windows XP and earlier

---

## 🧠 Architecture

- x86 (32-bit) build
- Runs on:
  - 32-bit Windows
  - 64-bit Windows (via WOW64)

A separate x64 build is not required.

---

## ⚙️ Runtime Dependencies

- None
- Built with static MSVC runtime (`/MT`)
- No Visual C++ Redistributable required
- No third-party DLLs required

Relies only on standard Windows components:
- Win32
- WinHTTP
- BCrypt
- RichEdit

---

## 🌐 Networking Requirements

- HTTPS access using **TLS 1.2**
- Required to download manifests and updates

---

## 🔐 Security

- Uses WinHTTP for secure network communication
- Uses Windows CNG (`BCrypt`) for hashing
- Supports Authenticode validation (production)
- SHA-256 fallback logic available for development/debug scenarios

---

## 🛠️ Notes on Windows 7

Windows 7 is considered legacy.

While the program may function on fully updated systems,  
Windows 10 or newer is strongly recommended for best reliability and security.

---

## 📄 Documentation

Documentation is provided in:
- `.docx` (primary format)
- `.rtf` (maximum compatibility)

Legacy `.doc` format is not used due to deprecation and reduced reliability.

---

## 🚀 Usage

1. Launch `MapPackSyncTool.exe`
2. Select your Istaria resources folder
3. Choose an operation:
   - Add
   - Sync
   - Remove
4. Monitor progress and logs
5. Use "Check for Updates" to verify latest version

---

## 🔧 Development Notes

- Single translation unit design
- Win32 API-based UI (no frameworks)
- Manifest-driven sync architecture
- RAII patterns for resource management
- Designed for portability and minimal dependencies

---

## 📌 Versioning

- Version is embedded via Windows `VERSIONINFO` resource
- Remote version is checked via `version.txt`
- SHA-256 hash validation included for integrity

---

## 👍 Summary

MapPack Sync Tool is designed to be:
- Simple
- Portable
- Secure
- Compatible across a wide range of Windows systems

---
