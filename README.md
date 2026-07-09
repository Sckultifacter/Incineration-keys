# Incineration Keys

A lightweight, registry-integrated Windows tray application for system-wide hotkeys. It automates system performance tuning, offline document conversion, temp file cleanup, and local network file sharing.

## Core Features

* **Extreme Battery Mode**: Caps CPU at 70%, disables turbo boost, stops background services (Search, Spooler, Update, SysMain), dims brightness, disables visual effects, and kills user-defined background apps.
* **Performance Mode**: Activates High/Ultimate Performance power plan, disables visual animations, and runs Game Mode registry optimizations.
* **Document Converter**: Offline conversions between PDF, DOCX, and images (PNG/JPG/WEBP) using Word's layout engine and PyMuPDF.
* **Incinerator Drop**: Launches a local Flask server to transfer files to/from other local network devices via QR code.
* **Clean Temp**: Asynchronously purges the Windows TEMP directory without freezing the background thread.

---

## Configuration

Settings are managed in `config.json`. Key combinations are validated at startup to avoid system conflicts.

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

---

## Quick Start

### Installation
```bash
git clone https://github.com/Sckultifacter/Incineration-keys.git
cd Incineration-keys
pip install -r requirements.txt
```

### Running the App
```bash
python engine.py
```

### Packaging (Build Executable)
```bash
pip install pyinstaller
pyinstaller engine.spec
```
The compiled standalone executable will be generated inside the `dist` folder.
