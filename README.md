# Incineration Keys

Incineration Keys is a lightweight, registry-integrated Windows system tray daemon written in Python. It provides high-performance system automation, quick document format conversion, safe temporary file cleanup, and a local network file transfer server accessible via global hotkeys. 

The application runs in the background, rendering native fade-out On-Screen Display (OSD) notifications to confirm actions and settings changes.

## Features

### Extreme Battery Mode
Toggles aggressive power saving settings, including:
* Duplicating the active power plan, capping CPU execution to 70%, and disabling Intel/AMD processor boost modes.
* Stopping resource-heavy background services like Windows Search (WSearch), Print Spooler (Spooler), Xbox Services, SysMain, and Windows Update.
* Forcefully terminating user-configured background processes such as OneDrive, SearchIndexer, and Adobe update processes.
* Dimming screen brightness and turning off standard Windows visual effects.
* Automatically restoring all services, original brightness, and the previous active power scheme when toggled off.

### Performance Mode
Optimizes the operating system for heavy workloads:
* Activates the Windows Ultimate or High Performance power scheme.
* Disables Windows visual effects and UI animations to prioritize CPU cycles.
* Suppresses system notifications and pushes Windows game-mode optimizations.
* Preserves system stability by restoring visual effects, animations, and power schemes upon deactivation.

### Smart Offline Document Converter
Converts documents locally without uploading sensitive files to external web servers:
* Conversions supported: DOCX to PDF, PDF to DOCX (using Word automation or PyMuPDF/pdf2docx layouts), and Images (PNG/JPG/WEBP) to PDF/Images.
* Features a high-fidelity Microsoft Word conversion engine that bypasses Protected View warnings for seamless offline file generation.

### Local Network File Transfer (Incinerator Drop)
Launches a localized, server-less file-sharing portal:
* Spawns a background Flask web server listening on the local area network.
* Displays a modern graphical window containing a QR code for mobile phones or secondary PCs to scan.
* Supports uploading files directly to the host PC or fetching hosted files via an offline web browser interface.

### Background System Sweep (Clean Temp)
Instantly purges junk files:
* Spawns a background thread to safely unlink temporary files and folders from the Windows TEMP directory.
* Avoids user interface freezes by reporting progress via the OSD overlay upon completion.

---

## Installation

### Prerequisites
* Windows OS
* Python 3.8 or higher
* Microsoft Word (required only for DOCX-to-PDF or Word-based PDF-to-DOCX conversions)

### Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/Sckultifacter/Incineration-keys.git
   cd Incineration-keys
   ```

2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

3. Run the application:
   ```bash
   python engine.py
   ```

---

## Configuration

Settings are dynamically read and written to `config.json`. The application validates configured hotkeys at startup to prevent conflicts with reserved Windows combinations (such as `ctrl+alt+del` or `alt+f4`).

Example configuration:
```json
{
    "hotkeys": {
        "Trigger Converter": "ctrl+alt+z",
        "Performance Mode": "ctrl+alt+p",
        "Open": "ctrl+alt+o",
        "Clean Temp": "ctrl+a+delete",
        "Incinerator Drop": "ctrl+alt+d",
        "Extreme Battery": "ctrl+alt+b"
    },
    "run_on_startup": true,
    "battery_kill_list": [
        "AdobeUpdateService.exe",
        "OneDrive.exe",
        "PhoneExperienceHost.exe",
        "SearchIndexer.exe"
    ]
}
```

### Parameter Documentation
* **hotkeys**: Custom key combinations mapped to execution triggers. Modified hotkeys are bound immediately on settings save.
* **run_on_startup**: When set to true, the application registers itself to the Windows Registry run keys (`HKCU\Software\Microsoft\Windows\CurrentVersion\Run`) to launch on boot.
* **battery_kill_list**: A list of executable names terminated immediately upon entering Extreme Battery Mode.

---

## Compilation

To package the application as a standalone executable (`.exe`) that runs silently with a system tray icon:

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```

2. Compile using the specification files:
   ```bash
   pyinstaller engine.spec
   ```
   Or:
   ```bash
   pyinstaller IncinerationKeys.spec
   ```

The compiled application will be generated in the `dist` directory.
