# Incineration Keys

Incineration Keys is a registry-integrated Windows system tray application that automates performance tuning, offline document conversion, file cleanups, and local network file sharing via global hotkeys. 

The application runs in the background and sends native fade-out On-Screen Display (OSD) notifications to confirm actions.

## Features

### Extreme Battery Mode
Toggles aggressive power saving configurations:
* Clones the current power plan, limits CPU maximum frequency to 70%, and disables CPU turbo boost.
* Suspends heavyweight background services (Windows Search, Print Spooler, Xbox services, SysMain, and Windows Update).
* Kills user-defined background apps (such as OneDrive, Adobe updates, and search indexers).
* Lowers screen brightness and disables Windows visual effects.
* Automatically restarts services, restores brightness, and reverts to the original power scheme when toggled off.

### Performance Mode
Optimizes Windows for heavy workloads:
* Activates the Ultimate Performance or High Performance power scheme.
* Disables Windows visual effects and UI animations to prioritize CPU execution.
* Suppresses system notifications and enables Windows Game Mode settings.
* Restores all original settings when disabled.

### Smart Offline Document Converter
Converts documents locally:
* Converts between DOCX, PDF, and images (PNG, JPG, WEBP).
* Employs Microsoft Word's layout engine to bypass Protected View warnings, retaining high-fidelity edits.

### Local Network File Transfer (Incinerator Drop)
Facilitates quick device-to-device transfers:
* Starts a background Flask server on the local network.
* Displays a QR code for mobile devices or other PCs to scan.
* Provides an offline web portal to upload files to the host PC or download files from it.

### Background System Sweep (Clean Temp)
Safely deletes temporary files:
* Clears files and directories in the Windows TEMP folder asynchronously.
* Employs a background thread to prevent UI freezing.

---

## Configuration

All mappings are defined in `config.json`:

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
    "battery_kill_list": ["OneDrive.exe", "SearchIndexer.exe"]
}
```
* **hotkeys**: Binds custom key combinations to triggers.
* **run_on_startup**: When true, registers the app to run on Windows boot (`HKCU\...\Run`).
* **battery_kill_list**: List of executable processes killed during Extreme Battery Mode.

---

## Quick Start

### Setup
1. Clone the repository and install dependencies:
   ```bash
   git clone https://github.com/Sckultifacter/Incineration-keys.git
   cd Incineration-keys
   pip install -r requirements.txt
   ```

2. Run the application:
   ```bash
   python engine.py
   ```
