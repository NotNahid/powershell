# 🛡️ Secure Vault v1.2

A lightweight, single-file PowerShell utility that integrates directly into Windows Explorer to provide high-level folder encryption.

## 🌟 Overview
Secure Vault allows you to "Lock" and "Unlock" folders directly from your right-click context menu. It replaces heavy third-party encryption software with a transparent, 100% PowerShell-based solution using industry-standard cryptography.

## ✨ Key Features
- **Native Explorer Integration:** Adds "🛡️ Lock Folder" and "🔓 Unlock Vault" options to your right-click menu.
- **AES-256 Encryption:** Military-grade data protection using the .NET `AesManaged` class.
- **PBKDF2 Key Derivation:** Uses `Rfc2898DeriveBytes` with 10,000 iterations and a unique salt for every vault, making it resistant to brute-force attacks.
- **On-Demand Execution:** The script only runs when you click it. Zero background processes, zero RAM usage when idle.
- **Self-Contained:** A single `.ps1` file with an embedded WPF GUI. No external folders, icons, or libraries required.

## 🚀 Installation
1.  **Placement:** Save `SecureVault.ps1` to a permanent location (e.g., `C:\Scripts\SecureVault.ps1`).
2.  **Run as Admin:** Right-click the script and select **"Run with PowerShell"**.
3.  **Register:** Click the **"INSTALL EXPLORER MENU"** button in the manager window.
4.  **Confirm:** Accept the Windows UAC prompt to allow the script to update the Registry.

## 📂 Usage
### To Lock a Folder:
1.  Right-click any folder in Windows Explorer.
2.  Select **"🛡️ Lock Folder (Secure Vault)"**.
3.  Enter a strong password in the pop-up box.
4.  The folder will be compressed, encrypted, and saved as a `.vault` file.
5.  *Optional:* Confirm if you want to delete the original unencrypted folder.

### To Unlock a Vault:
1.  Right-click any `.vault` file.
2.  Select **"🔓 Unlock Vault"**.
3.  Enter your password.
4.  The original folder will be restored instantly in the same directory.

## 🛠️ Technical Details
- **Encryption Algorithm:** AES (Advanced Encryption Standard)
- **Key Size:** 256-bit
- **Mode:** CBC (Cipher Block Chaining)
- **Padding:** PKCS7
- **Vault Structure:** `[Salt (32 bytes)] + [IV (16 bytes)] + [Encrypted Data]`
- **Dependencies:** Built-in Windows .NET Framework (System.Security.Cryptography).

## ⚠️ Security Warning
**This is true encryption.** If you lose your password, the data inside the `.vault` file is **mathematically unrecoverable**. There is no "backdoor" or "forgot password" feature.

---
*Created for power users who value privacy, transparency, and system efficiency.*
