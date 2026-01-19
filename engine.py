import keyboard
import os
import sys
import tkinter as tk
from tkinter import filedialog
from threading import Thread
from plyer import notification
import winreg  # Used to modify Windows Registry for Auto-Start
import pystray
from PIL import Image, ImageDraw
import subprocess
import ctypes

import json
import shutil
import time
import shlex
import socket
import io

# =========================================================
#  ZONE 0: CONFIGURATION MANGER & VALIDATOR
# =========================================================
class ConfigManager:
    DEFAULT_CONFIG = {
        "hotkeys": {
            "Trigger Converter": "ctrl+alt+z",
            "Performance Mode": "ctrl+alt+g",
            "Open": "ctrl+alt+s",
            "Clean Temp": "ctrl+a+delete",
            "Incinerator Drop": "ctrl+alt+d"
        },
        "run_on_startup": True,
        "ngrok_token": "" 
    }
    
    def __init__(self):
        self.config_path = self.get_config_path()
        self.config = self.load_config()
        # Migrate old config if needed
        if 'hotkey' in self.config:
            self.config['hotkeys'] = {"Trigger Converter": self.config['hotkey']}
            del self.config['hotkey']
            self.save_config(self.config)
        
        # KEY MIGRATION: "Open Settings" -> "Open"
        if 'hotkeys' in self.config and 'Open Settings' in self.config['hotkeys']:
            # Only migrate if "Open" is missing
            if 'Open' not in self.config['hotkeys']:
                 self.config['hotkeys']['Open'] = self.config['hotkeys']['Open Settings']
            
            # Remove old key
            del self.config['hotkeys']['Open Settings']
            self.save_config(self.config)

    def get_config_path(self):
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, "config.json")

    def load_config(self):
        if not os.path.exists(self.config_path):
            return self.DEFAULT_CONFIG.copy()
        try:
            with open(self.config_path, 'r') as f:
                data = json.load(f)
                updated = False
                # Ensure structure integrity
                if 'hotkeys' not in data:
                    data['hotkeys'] = self.DEFAULT_CONFIG['hotkeys'].copy()
                    updated = True
                else:
                    # Merge missing defaults
                    for key, val in self.DEFAULT_CONFIG['hotkeys'].items():
                        if key not in data['hotkeys']:
                            data['hotkeys'][key] = val
                            updated = True
                
                if updated:
                    self.save_config(data)
                    
                return data
        except:
            return self.DEFAULT_CONFIG.copy()

    def save_config(self, new_config):
        if hasattr(self, 'config'):
            self.config.update(new_config)
            data_to_save = self.config
        else:
            data_to_save = new_config
            
        with open(self.config_path, 'w') as f:
            json.dump(data_to_save, f, indent=4)

    def get(self, key):
        return self.config.get(key, self.DEFAULT_CONFIG.get(key))

class HotkeyValidator:
    RESERVED_KEYS = [
        'ctrl+c', 'ctrl+x', 'ctrl+v', 'ctrl+z', 'ctrl+a', 'ctrl+s',
        'alt+f4', 'ctrl+alt+del', 'win', 'ctrl+shift+esc'
    ]
    
    @staticmethod
    def validate(hotkey, current_hotkeys, ignore_action=None):
        hotkey = hotkey.lower().replace(" ", "")
        
        # 1. Check Empty
        if not hotkey:
            return False, "Hotkey cannot be empty."

        # 2. Check Reserved
        if hotkey in HotkeyValidator.RESERVED_KEYS:
            return False, f"'{hotkey}' is a reserved system key."

        # 3. Check Duplicates
        for action, key in current_hotkeys.items():
            if action == ignore_action:
                continue
            if key == hotkey:
                return False, f"Hotkey already used by '{action}'."

        # 4. Check Syntax (Try to parse)
        try:
            keyboard.parse_hotkey(hotkey)
        except:
            return False, "Invalid hotkey syntax."

        return True, "Valid"

# =========================================================
#  ZONE 1: CONFIGURATION (Add new Shortcuts here later)
# =========================================================
# SHORTCUT_MAP REMOVED - NOW DYNAMIC BASED ON CONFIG

# =========================================================
#  ZONE 2: LAZY MODULES (Only loaded when used)
# =========================================================

def convert_docx_to_pdf(input_path):
    from docx2pdf import convert
    output_path = os.path.splitext(input_path)[0] + ".pdf"
    convert(input_path, output_path)
    return output_path

def convert_pdf_to_docx(input_path):
    from pdf2docx import Converter
    output_path = os.path.splitext(input_path)[0] + ".docx"
    cv = Converter(input_path)
    cv.convert(output_path)
    cv.close()
    return output_path

def convert_pdf_to_images(input_path):
    import fitz  # PyMuPDF
    doc = fitz.open(input_path)
    base_path = os.path.splitext(input_path)[0]
    for i, page in enumerate(doc):
        pix = page.get_pixmap()
        output = f"{base_path}_page_{i+1}.png"
        pix.save(output)
    return f"{base_path} (Images)"

def convert_image_to_pdf(input_path):
    from PIL import Image
    image = Image.open(input_path)
    img_conv = image.convert('RGB')
    output_path = os.path.splitext(input_path)[0] + ".pdf"
    img_conv.save(output_path)
    return output_path

# =========================================================
#  ZONE 2.2: OSD MANAGER (Custom Notifications)
# =========================================================
class OSDManager:
    def __init__(self, root):
        self.root = root
        self.osd_window = None

    def show(self, text, title=None, duration=2000):
        if self.osd_window:
            try:
                self.osd_window.destroy()
            except:
                pass
        
        # Create non-interactive window
        self.osd_window = tk.Toplevel(self.root)
        self.osd_window.overrideredirect(True)
        self.osd_window.attributes('-topmost', True)
        self.osd_window.configure(bg='black')
        
        # Rounded look via frame (approximate)
        container = tk.Frame(self.osd_window, bg='#ff5500', padx=2, pady=2)
        container.pack(fill='both', expand=True)
        
        inner = tk.Frame(container, bg='#151515', padx=30, pady=15)
        inner.pack(fill='both', expand=True)

        # Title (Optional)
        if title:
            tk.Label(inner, text=title.upper(), fg='#ff5500', bg='#151515', 
                     font=('Segoe UI', 9, 'bold')).pack(anchor='w', pady=(0, 2))
        
        # Message
        tk.Label(inner, text=text, fg='white', bg='#151515', 
                 font=('Segoe UI', 11, 'bold')).pack(anchor='w')
                 
        # Centered Bottom positioning
        self.osd_window.update_idletasks()
        width = self.osd_window.winfo_width()
        height = self.osd_window.winfo_height()
        screen_width = self.osd_window.winfo_screenwidth()
        screen_height = self.osd_window.winfo_screenheight()
        
        x = (screen_width // 2) - (width // 2)
        y = screen_height - height - 100 # 100px from bottom
        
        self.osd_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Fade out / Destroy
        self.root.after(duration, self.fade_out)

    def fade_out(self):
        if not self.osd_window: return
        try:
            alpha = self.osd_window.attributes('-alpha')
            if alpha > 0:
                alpha -= 0.1
                self.osd_window.attributes('-alpha', alpha)
                self.root.after(50, self.fade_out)
            else:
                self.osd_window.destroy()
                self.osd_window = None
        except:
            if self.osd_window:
                self.osd_window.destroy()
                self.osd_window = None

def convert_image_format(input_path, target_ext):
    from PIL import Image
    image = Image.open(input_path)
    if target_ext in ['jpg', 'jpeg'] and image.mode == 'RGBA':
        image = image.convert('RGB')
    output_path = os.path.splitext(input_path)[0] + f".{target_ext}"
    image.save(output_path)
    return output_path

# =========================================================
#  ZONE 2.5: PERFORMANCE MODE MANAGER
# =========================================================
class PerformanceModeManager:
    def __init__(self):
        self.is_active = False
        self.prev_power_guid = None
        self.prev_visual_fx = None
        self.prev_toast = 1
        self.prev_brightness = None
        self.temp_plan_created = False

        # High Performance GUIDs
        self.HP_GUID = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
        self.ULTIMATE_GUID = "e9a42b02-d5df-448d-aa00-03f14749eb61"
        self.INCINERATION_GUID = None # Will store our custom plan GUID
        
        # Popup suppression setup
        self.startupinfo = subprocess.STARTUPINFO()
        self.startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        self.startupinfo.wShowWindow = 0 # SW_HIDE
        self.creation_flags = 0x08000000 # CREATE_NO_WINDOW

    def _run_cmd(self, cmd):
        try:
            # Parse command string to list to avoid shell=True
            args = shlex.split(cmd)
            subprocess.run(
                args, 
                shell=False, 
                stdout=subprocess.DEVNULL, 
                stderr=subprocess.DEVNULL, 
                stdin=subprocess.DEVNULL,
                creationflags=self.creation_flags, 
                startupinfo=self.startupinfo
            )
        except:
            pass
            
    def _run_ps_output(self, cmd):
        try:
            # Run PowerShell command and get output
            # Add -WindowStyle Hidden
            # Run PowerShell command and get output
            # Add -WindowStyle Hidden -NoProfile -NonInteractive
            args = ["powershell", "-WindowStyle", "Hidden", "-NoProfile", "-NonInteractive", "-Command", cmd]
            res = subprocess.check_output(
                args, 
                text=True, 
                stdin=subprocess.DEVNULL,
                creationflags=self.creation_flags,
                startupinfo=self.startupinfo
            ).strip()
            return res
        except:
            return None

    def _get_active_scheme(self):
        try:
            args = ["powercfg", "/getactivescheme"]
            output = subprocess.check_output(
                args, 
                shell=False,
                stdin=subprocess.DEVNULL,
                creationflags=self.creation_flags,
                startupinfo=self.startupinfo
            ).decode()
            # Output format: "Power Scheme GUID: <GUID>  (<Name>)"
            return output.split("GUID:")[1].split()[0].strip()
        except:
            return None

    def _plan_exists(self, guid):
        try:
            args = ["powercfg", "/list"]
            output = subprocess.check_output(
                args, 
                shell=False,
                stdin=subprocess.DEVNULL,
                creationflags=self.creation_flags,
                startupinfo=self.startupinfo
            ).decode()
            return guid in output
        except:
            return False

    def _create_temp_plan(self, source_guid):
        try:
            # Duplicate the source plan
            args = ["powercfg", "/duplicatescheme", source_guid]
            output = subprocess.check_output(
                args, 
                shell=False,
                stdin=subprocess.DEVNULL,
                creationflags=self.creation_flags,
                startupinfo=self.startupinfo
            ).decode()
            # Output: "Power Scheme GUID: <NEW_GUID>  (Copy of <Name>)"
            new_guid = output.split("GUID:")[1].split()[0].strip()
            
            # Rename it
            self._run_cmd(f"powercfg /changename {new_guid} \"Incineration Performance\"")
            
            # MODIFY SETTINGS
            # 1. Processor Performance -> Min/Max to 100%
            # GUIDs: SUB_PROCESSOR (54533251-82be-4824-96c1-47b60b740d00)
            #        PROCTHROTTLEMIN (893dee8e-2bef-41e0-89c6-718607414d93)
            #        PROCTHROTTLEMAX (bc5038f7-fd9e-4799-a217-a0f48b1b7520)
            sub_proc = "54533251-82be-4824-96c1-47b60b740d00"
            min_proc = "893dee8e-2bef-41e0-89c6-718607414d93"
            max_proc = "bc5038f7-fd9e-4799-a217-a0f48b1b7520"
            
            self._run_cmd(f"powercfg /setacvalueindex {new_guid} {sub_proc} {min_proc} 100")
            self._run_cmd(f"powercfg /setdcvalueindex {new_guid} {sub_proc} {min_proc} 100")
            self._run_cmd(f"powercfg /setacvalueindex {new_guid} {sub_proc} {max_proc} 100")
            self._run_cmd(f"powercfg /setdcvalueindex {new_guid} {sub_proc} {max_proc} 100")

            # 2. PCI Express -> Link State Power Management -> Off (0)
            # GUIDs: SUB_PCIEXPRESS (501a4d13-42af-401f-99da-3bcae46b6910)
            #        ASPM (ee12f906-d1d7-4267-8d02-1de23b5c6d26)
            sub_pci = "501a4d13-42af-401f-99da-3bcae46b6910"
            aspm = "ee12f906-d1d7-4267-8d02-1de23b5c6d26"
            
            self._run_cmd(f"powercfg /setacvalueindex {new_guid} {sub_pci} {aspm} 0")
            self._run_cmd(f"powercfg /setdcvalueindex {new_guid} {sub_pci} {aspm} 0")

            # Activate changes
            self._run_cmd(f"powercfg /setactive {new_guid}")
            
            return new_guid
        except Exception as e:
            print(f"Failed to create temp plan: {e}")
            return None

    def _get_registry_value(self, key_path, value_name):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_READ)
            val, _ = winreg.QueryValueEx(key, value_name)
            winreg.CloseKey(key)
            return val
        except:
            return None

    def _set_registry_value(self, key_path, value_name, value, val_type=winreg.REG_DWORD):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, value_name, 0, val_type, value)
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Reg Error {key_path}: {e}")

    def capture_state(self):
        # 1. Power Plan
        self.prev_power_guid = self._get_active_scheme()
        if not self.prev_power_guid:
            self.prev_power_guid = "381b4222-f694-41f0-9685-ff5bb260df2e" # Default Balanced fallback

        # 2. Visual Effects
        val = self._get_registry_value(r"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects", "VisualFXSetting")
        self.prev_visual_fx = val if val is not None else 3

        # 3. Notifications (Focus Assist / Toast). Check multiple keys to be sure.
        val = self._get_registry_value(r"Software\Microsoft\Windows\CurrentVersion\PushNotifications", "ToastEnabled")
        self.prev_toast = val if val is not None else 1
        
        # 4. Brightness
        try:
            b = self._run_ps_output("(Get-CimInstance -Namespace root/wmi -ClassName WmiMonitorBrightness).CurrentBrightness")
            if b: self.prev_brightness = int(b)
        except:
            self.prev_brightness = 100 # Fallback

    def enable(self):
        self.is_active = True
        self.capture_state()
        self.temp_plan_created = False
        
        # 1. Power Plan Implementation
        # Check if High Performance exists
        if self._plan_exists(self.HP_GUID):
            self._run_cmd(f"powercfg /setactive {self.HP_GUID}")
        else:
            # Create Custom Plan based on current valid plan
            # We duplicate the active plan (which is likely Balanced or compatible with the hardware e.g. Modern Standby)
            self.INCINERATION_GUID = self._create_temp_plan(self.prev_power_guid)
            if self.INCINERATION_GUID:
                self.temp_plan_created = True
        
        # 2. Visual Effects -> Best Performance (2)
        self._set_registry_value(r"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects", "VisualFXSetting", 2)
        
        # 3. Notifications -> Off (0)
        self._set_registry_value(r"Software\Microsoft\Windows\CurrentVersion\PushNotifications", "ToastEnabled", 0)
        self._set_registry_value(r"Software\Microsoft\Windows\CurrentVersion\Notifications\Settings", "NOC_GLOBAL_SETTING_TOAST_ENABLED_KEY", 0)
        
        # 4. Game Mode -> On (1)
        self._set_registry_value(r"Software\Microsoft\GameBar", "AllowAutoGameMode", 1)
        
        # 5. Restore Brightness Forcefully (after a short delay to allow power plan to settle)
        if self.prev_brightness is not None:
             time.sleep(0.5) 
             cmd = f"(Get-WmiObject -Namespace root/wmi -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, {self.prev_brightness})"
             self._run_ps_output(cmd)

        self._toggle_animations(False) # Turn OFF animations
        
        # Notify
        # Note: Notification is handled by the caller, but we return success status if needed

    def disable(self):
        self.is_active = False
        
        # 1. Restore Power Plan
        if self.prev_power_guid:
            self._run_cmd(f"powercfg /setactive {self.prev_power_guid}")
            
        # Cleanup temporary plan
        if self.temp_plan_created and self.INCINERATION_GUID:
            time.sleep(0.5) # Wait for switch to complete
            self._run_cmd(f"powercfg /delete {self.INCINERATION_GUID}")
            self.INCINERATION_GUID = None
            self.temp_plan_created = False
        
        # 2. Restore Visual Effects
        if self.prev_visual_fx is not None:
             self._set_registry_value(r"Software\Microsoft\Windows\CurrentVersion\Explorer\VisualEffects", "VisualFXSetting", self.prev_visual_fx)
        
        # 3. Restore Notifications
        if self.prev_toast is not None:
            self._set_registry_value(r"Software\Microsoft\Windows\CurrentVersion\PushNotifications", "ToastEnabled", self.prev_toast)
            self._set_registry_value(r"Software\Microsoft\Windows\CurrentVersion\Notifications\Settings", "NOC_GLOBAL_SETTING_TOAST_ENABLED_KEY", self.prev_toast)

        if self.prev_brightness is not None:
             time.sleep(0.5)
             cmd = f"(Get-WmiObject -Namespace root/wmi -Class WmiMonitorBrightnessMethods).WmiSetBrightness(1, {self.prev_brightness})"
             self._run_ps_output(cmd)

        user32 = ctypes.windll.user32
        user32.SystemParametersInfoW(0x0049, 0, 0, 0x02 | 0x01) # Reset (though toggling handles it better)
        self._toggle_animations(True) # Turn ON animations

    def _toggle_animations(self, enable_anim):
        user32 = ctypes.windll.user32
        # SPI_SETCLIENTAREAANIMATION = 0x1043
        # SPI_SETANIMATION = 0x0049
        # SPI_SETCOMBOBOXANIMATION = 0x1005
        # SPI_SETLISTBOXSMOOTHSCROLLING = 0x1007
        # SPI_SETMENUANIMATION = 0x1003
        # SPI_SETMENUFADE = 0x1013
        # SPI_SETSELECTIONFADE = 0x1015
        # SPI_SETTOOLTIPANIMATION = 0x1017
        # SPI_SETUIEFFECTS = 0x103F
        
        flags = [0x1043, 0x0049, 0x1005, 0x1007, 0x1003, 0x1013, 0x1015, 0x1017, 0x103F]
        
        # Structure for SPI_SETANIMATION ( ANIMATIONINFO )
        class ANIMATIONINFO(ctypes.Structure):
            _fields_ = [("cbSize", ctypes.c_uint), ("iMinAnimate", ctypes.c_int)]
            
        anim_info = ANIMATIONINFO()
        anim_info.cbSize = ctypes.sizeof(ANIMATIONINFO)
        anim_info.iMinAnimate = 1 if enable_anim else 0
        
        val = 1 if enable_anim else 0
        
        for flag in flags:
            if flag == 0x0049: # SPI_SETANIMATION requires structure? No, simple bool usually works for toggle, but let's be safe
               user32.SystemParametersInfoW(flag, 0, ctypes.byref(anim_info), 0x02 | 0x01)
            else:
               user32.SystemParametersInfoW(flag, 0, val, 0x02 | 0x01)

# =========================================================
#  ZONE 2.6: AIRDROP MANAGER (Incinerator Drop)
# =========================================================
class AirDropManager:
    def __init__(self, root, notify_func):
        self.root = root
        self.notify = notify_func
        self.server_thread = None
        self.app = None
        self.server_running = False
        self.port = 8000
        self.mode = None # 'SEND' or 'RECEIVE'
        self.file_to_send = None
        self.upload_dir = os.path.join(os.path.expanduser("~"), "Downloads", "Incinerator_Drop")
        
    def get_local_ip(self):
        try:
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.connect(("8.8.8.8", 80))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except:
            return "127.0.0.1"

    def generate_qr(self, data):
        import qrcode
        qr = qrcode.QRCode(version=1, box_size=10, border=2)
        qr.add_data(data)
        qr.make(fit=True)
        return qr.make_image(fill_color="black", back_color="white")

    def start_server(self, mode, files=None, use_global=False, ngrok_token=None):
        from flask import Flask, send_file, request, render_template_string, abort, session, redirect, url_for
        import logging
        import random
        
        # Suppress Flask CLI logs
        log = logging.getLogger('werkzeug')
        log.setLevel(logging.ERROR)
        
        self.mode = mode
        self.files_to_send = files if files else []
        self.files_map = {os.path.basename(f): f for f in self.files_to_send}
        self.server_running = True
        self.public_url = None
        self.pin = str(random.randint(1000, 9999))
        
        if not os.path.exists(self.upload_dir):
            os.makedirs(self.upload_dir)

        app = Flask(__name__)
        app.secret_key = os.urandom(24)
        self.app = app

        @app.before_request
        def require_login():
            if request.endpoint == 'static': return
            if request.endpoint == 'login': return
            
            if 'authenticated' not in session:
                return redirect(url_for('login'))

        @app.route('/login', methods=['GET', 'POST'])
        def login():
            if request.method == 'POST':
                if request.form.get('pin') == self.pin:
                    session['authenticated'] = True
                    return redirect(url_for('index'))
                else:
                    return render_template_string(self._get_login_html(error="Invalid PIN"))
            return render_template_string(self._get_login_html())

        @app.route('/')
        def index():
            if self.mode == 'SEND' and self.files_to_send:
                # Single File -> Auto Download
                if len(self.files_to_send) == 1:
                    return send_file(self.files_to_send[0], as_attachment=True)
                
                # Multiple Files -> List
                items = "".join([
                    f'<li style="margin:10px 0;"><a href="/download/{name}" style="color:#ff5500; text-decoration:none; font-size:18px; border:1px solid #333; padding:10px; display:block; border-radius:5px; background:#1a1a1a;">{name}</a></li>' 
                    for name in self.files_map.keys()
                ])
                
                html = f"""
                <!doctype html>
                <html style="background:#121212; color:white; font-family:sans-serif; text-align:center; padding:20px;">
                <head><meta name="viewport" content="width=device-width, initial-scale=1"></head>
                <body>
                    <h1 style="color:#ff5500;">AVAILABLE FILES</h1>
                    <ul style="list-style:none; padding:0; text-align:left;">{items}</ul>
                </body>
                </html>
                """
                return render_template_string(html)
                    
            elif self.mode == 'RECEIVE':
                # Upload Form
                html = """
                <!doctype html>
                <html style="background:#121212; color:white; font-family:sans-serif; text-align:center; padding:50px;">
                <head><meta name="viewport" content="width=device-width, initial-scale=1"></head>
                <body>
                    <h1 style="color:#ff5500;">INCINERATOR DROP</h1>
                    <form method="post" action="/upload" enctype="multipart/form-data" 
                          style="border:2px dashed #333; padding:40px; border-radius:10px;">
                        <input type="file" name="file" multiple style="margin-bottom:20px; color:#888;">
                        <br>
                        <button type="submit" 
                                style="background:#ff5500; color:white; border:none; padding:15px 30px; 
                                       font-size:18px; border-radius:5px; cursor:pointer;">
                            UPLOAD TO PC
                        </button>
                    </form>
                </body>
                </html>
                """
                return render_template_string(html)
            return "Incinerator Drop: Unknown State"

        @app.route('/download/<filename>')
        def download(filename):
            if filename in self.files_map:
                return send_file(self.files_map[filename], as_attachment=True)
            return abort(404)

        @app.route('/upload', methods=['POST'])
        def upload_file():
            if 'file' not in request.files:
                return 'No file part'
            
            uploaded_files = request.files.getlist('file')
            count = 0
            
            for file in uploaded_files:
                if file.filename == '': continue
                
                filename = file.filename
                # Simple sanitize
                filename = os.path.basename(filename)
                save_path = os.path.join(self.upload_dir, filename)
                file.save(save_path)
                count += 1
                
            # Notify PC
            if count > 0:
                self.root.after(0, lambda: self.notify("Incinerator Drop", f"Received {count} files."))
                
            return f"""
            <h1 style="color:green; text-align:center; font-family:sans-serif; margin-top:50px;">
                {count} FILES SENT SUCCESSFULLY
            </h1>
            <p style="text-align:center;"><a href="/" style="color:white;">Send More</a></p>
            """
        
        return self._run_server_thread(use_global, ngrok_token)

    def _get_login_html(self, error=""):
        return f"""
        <!doctype html>
        <html style="background:#121212; color:white; font-family:sans-serif; text-align:center; padding:50px;">
        <head><meta name="viewport" content="width=device-width, initial-scale=1"></head>
        <body>
            <h1 style="color:#ff5500;">SECURE LOGIN</h1>
            <p style="color:#888;">Enter the 4-digit PIN shown on the Host PC.</p>
            <form method="post" style="margin-top:20px;">
                <input type="number" name="pin" placeholder="PIN" 
                       style="padding:10px; font-size:20px; width:100px; text-align:center; border-radius:5px; border:none;">
                <br><br>
                <button type="submit" 
                        style="background:#ff5500; color:white; border:none; padding:10px 30px; 
                               font-size:16px; border-radius:5px; cursor:pointer;">
                    ENTER
                </button>
            </form>
            <p style="color:red; margin-top:20px;">{error}</p>
        </body>
        </html>
        """

    def _run_server_thread(self, use_global, ngrok_token):
        def run_app():
            try:
                # Run on 0.0.0.0 to be accessible
                self.app.run(host='0.0.0.0', port=self.port, use_reloader=False)
            except Exception as e:
                print(f"Server Error: {e}")

        self.server_thread = Thread(target=run_app, daemon=True)
        self.server_thread.start()
        
        if use_global:
            try:
                from pyngrok import ngrok
                
                # Apply Token
                if ngrok_token:
                    ngrok.set_auth_token(ngrok_token)
                
                # Open a HTTP tunnel on the default port 8000
                tunnel = ngrok.connect(self.port)
                self.public_url = tunnel.public_url
                return self.public_url
            except Exception as e:
                print(f"Ngrok Error: {e}")
                self.notify(self.root, "Error", "Global tunnel failed. Check Token.")
        
        ip = self.get_local_ip()
        return f"http://{ip}:{self.port}"

    def stop_server(self):
        self.server_running = False
        
        # Kill ngrok tunnel
        if self.public_url:
            try:
                from pyngrok import ngrok
                ngrok.disconnect(self.public_url)
                ngrok.kill() # Ensure process is killed
                self.public_url = None
            except:
                pass

# =========================================================
#  ZONE 3: THE CORE ENGINE
# =========================================================

class AntigravityEngine:
    def __init__(self):
        self.config_manager = ConfigManager()
        self.performance_manager = PerformanceModeManager()
        
        # 1. Startup Check
        if self.config_manager.get('run_on_startup'):
            self.ensure_startup()
        else:
            self.remove_startup()
        
        self.root = tk.Tk()
        self.root.withdraw()
        
        # 2.1 OSD Manager
        self.osd = OSDManager(self.root)
        
        # 2.2 AirDrop Manager
        self.airdrop = AirDropManager(self.root, self.notify)
        
        # 3. Setup System Tray
        self.setup_tray()
        
        # 4. Register Keyboard Hook
        self.reload_hotkeys()
        
        print("--- Antigravity Engine Started ---")

        # 5. Keep Alive
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            self.quit_app()

    def reload_hotkeys(self):
        try:
            try:
                keyboard.unhook_all()
            except AttributeError:
                # Occasional library bug on first run or specific versions
                pass
            except Exception:
                pass

            hotkeys = self.config_manager.get('hotkeys')
            for action_name, hotkey in hotkeys.items():
                if action_name == "Trigger Converter":
                    try:
                        keyboard.add_hotkey(hotkey, lambda: self.safe_trigger(self.trigger_converter))
                        print(f"Registered '{action_name}' to {hotkey}")
                    except Exception as e:
                        print(f"Failed to register '{action_name}': {e}")
                elif action_name == "Performance Mode":
                    try:
                        keyboard.add_hotkey(hotkey, lambda: self.safe_trigger(self.toggle_performance_mode))
                        print(f"Registered '{action_name}' to {hotkey}")
                    except Exception as e:
                        print(f"Failed to register '{action_name}': {e}")
                elif action_name == "Clean Temp":
                    try:
                        keyboard.add_hotkey(hotkey, lambda: self.safe_trigger(self.clean_temp_files))
                        print(f"Registered '{action_name}' to {hotkey}")
                    except Exception as e:
                        print(f"Failed to register '{action_name}': {e}")
                elif action_name == "Open":
                    try:
                        keyboard.add_hotkey(hotkey, lambda: self.safe_trigger(self.open_settings_window))
                        print(f"Registered '{action_name}' to {hotkey}")
                    except Exception as e:
                        print(f"Failed to register '{action_name}': {e}")
                elif action_name == "Incinerator Drop":
                    try:
                        keyboard.add_hotkey(hotkey, lambda: self.safe_trigger(self.trigger_airdrop))
                        print(f"Registered '{action_name}' to {hotkey}")
                    except Exception as e:
                        print(f"Failed to register '{action_name}': {e}")
        except Exception as e:
            print(f"Failed to load hotkeys: {e}")

    def ensure_startup(self):
        """Adds this program to Windows Registry Run keys"""
        try:
            if getattr(sys, 'frozen', False):
                app_path = sys.executable
            else:
                app_path = os.path.abspath(__file__)

            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            app_name = "IncinerationKeys"

            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
            try:
                existing, _ = winreg.QueryValueEx(key, app_name)
                if existing == f'"{app_path}"':
                    winreg.CloseKey(key)
                    return # Already installed
            except FileNotFoundError:
                pass

            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{app_path}"')
            winreg.CloseKey(key)
        except Exception as e:
            print(f"Startup Error: {e}")

    def remove_startup(self):
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"
            app_name = "IncinerationKeys"
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_ALL_ACCESS)
            try:
                winreg.DeleteValue(key, app_name)
            except FileNotFoundError:
                pass
            winreg.CloseKey(key)
        except Exception:
            pass

    def create_tray_icon(self):
        # Try to load custom icon
        icon_path = "icon.jpg"
        if getattr(sys, 'frozen', False):
            # If compiled, look in the same folder as executable
            icon_path = os.path.join(os.path.dirname(sys.executable), "icon.jpg")
        else:
            icon_path = os.path.join(os.path.dirname(__file__), "icon.jpg")

        if os.path.exists(icon_path):
            try:
                return Image.open(icon_path)
            except Exception as e:
                print(f"Failed to load icon: {e}")

        # Fallback: Create a simple 64x64 icon
        width = 64
        height = 64
        color1 = (40, 40, 40)
        color2 = (200, 200, 200)
        
        image = Image.new('RGB', (width, height), color1)
        dc = ImageDraw.Draw(image)
        dc.rectangle((width // 4, height // 4, width * 3 // 4, height * 3 // 4), fill=color2)
        
        return image

    def setup_tray(self):
        image = self.create_tray_icon()
        menu = pystray.Menu(
            pystray.MenuItem("Settings", lambda: self.safe_trigger(self.open_settings_window)),
            pystray.MenuItem("Exit", self.quit_app)
        )
        self.icon = pystray.Icon("IncinerationKeys", image, "Incineration Keys", menu)
        # Run tray in separate thread so it doesn't block tkinter mainloop
        Thread(target=self.icon.run, daemon=True).start()

    def quit_app(self, icon=None, item=None):
        if hasattr(self, 'icon'):
            self.icon.stop()
        self.root.quit()
        os._exit(0)

    def safe_trigger(self, func):
        self.root.after(0, func)

    def notify(self, title, msg):
        self.safe_trigger(lambda: self.osd.show(text=msg, title=title, duration=3000))

    # --- ADVANCED SETTINGS UI ---
    def open_settings_window(self):
        if hasattr(self, 'settings_win') and self.settings_win.winfo_exists():
            self.settings_win.focus_force()
            return

        self.settings_win = tk.Toplevel(self.root)
        self.settings_win.title("Incineration Settings")
        self.settings_win.geometry("600x550") # Increased height for new buttons
        self.settings_win.configure(bg="#121212")
        
        # Header
        header = tk.Frame(self.settings_win, bg="#121212")
        header.pack(fill='x', pady=20)
        tk.Label(header, text="SETTINGS", bg="#121212", fg="#ff5500", font=("Segoe UI", 16, "bold")).pack(side='left', padx=20)

        # Hotkeys Section
        hk_frame = tk.LabelFrame(self.settings_win, text=" HOTKEYS ", bg="#121212", fg="#888", font=("Segoe UI", 9), bd=1, relief="solid")
        hk_frame.pack(fill='x', padx=20, pady=10)

        self.hotkey_vars = {}
        row = 0
        current_hotkeys = self.config_manager.get('hotkeys')
        
        for action_name, hotkey_val in current_hotkeys.items():
            tk.Label(hk_frame, text=action_name, bg="#121212", fg="#e0e0e0", font=("Segoe UI", 10)).grid(row=row, column=0, sticky='w', padx=15, pady=10)
            
            # Hotkey Button (Click to Record)
            btn_var = tk.StringVar(value=hotkey_val)
            self.hotkey_vars[action_name] = btn_var
            
            btn = tk.Button(hk_frame, textvariable=btn_var, 
                            bg="#222", fg="white", activebackground="#333", activeforeground="white",
                            relief="flat", font=("Consolas", 10), width=15)
            btn.config(command=lambda b=btn, v=btn_var, a=action_name: self.start_recording(b, v, a))
            btn.grid(row=row, column=1, padx=15, pady=10)
            row += 1

        # Startup Section
        self.var_startup = tk.BooleanVar(value=self.config_manager.get('run_on_startup'))
        cb_frame = tk.Frame(self.settings_win, bg="#121212")
        cb_frame.pack(fill='x', padx=20, pady=10)
        
        cb = tk.Checkbutton(cb_frame, text="Run on Startup", variable=self.var_startup, 
                            bg="#121212", fg="#e0e0e0", selectcolor="#222", activebackground="#121212", activeforeground="white", font=("Segoe UI", 10))
        cb.pack(side='left')

        self.lbl_error.pack(pady=(10, 0))

        # Ngrok Token Section
        tk.Label(self.settings_win, text="Ngrok Authtoken (For Global Mode)", bg="#121212", fg="#888", font=("Segoe UI", 9)).pack(anchor='w', padx=20, pady=(10,0))
        self.var_token = tk.StringVar(value=self.config_manager.get('ngrok_token'))
        tk.Entry(self.settings_win, textvariable=self.var_token, bg="#222", fg="white", insertbackground="white", relief="flat", font=("Consolas", 10)).pack(fill='x', padx=20, pady=5)

        # Action Buttons
        btn_frame = tk.Frame(self.settings_win, bg="#121212")
        btn_frame.pack(side='bottom', pady=20, fill='x', padx=20)
        
        # Details: Use variables to bind events
        btn_apply = tk.Button(btn_frame, text="APPLY CHANGES", command=self.save_settings, 
                  bg="#ff5500", fg="white", relief="flat", font=("Segoe UI", 11, "bold"), 
                  padx=30, pady=10, cursor="hand2", activebackground="#cc4400", activeforeground="white")
        btn_apply.pack(side="right", padx=10)
        
        btn_cancel = tk.Button(btn_frame, text="CANCEL", command=self.settings_win.destroy, 
                  bg="#222", fg="#aaa", relief="flat", font=("Segoe UI", 11), 
                  padx=30, pady=10, cursor="hand2", activebackground="#333", activeforeground="#ccc")
        btn_cancel.pack(side="right", padx=10)

        # Hover Effects
        def on_enter_apply(e): btn_apply['bg'] = '#ff7722'
        def on_leave_apply(e): btn_apply['bg'] = '#ff5500'
        btn_apply.bind("<Enter>", on_enter_apply)
        btn_apply.bind("<Leave>", on_leave_apply)

        def on_enter_cancel(e): btn_cancel['bg'] = '#333'; btn_cancel['fg'] = 'white'
        def on_leave_cancel(e): btn_cancel['bg'] = '#222'; btn_cancel['fg'] = '#aaa'
        btn_cancel.bind("<Enter>", on_enter_cancel)
        btn_cancel.bind("<Leave>", on_leave_cancel)

    def start_recording(self, btn, string_var, action_name):
        btn.config(text="Press Keys...", bg="#ff5500", fg="white")
        # Use hooks to capture next combination
        self.recorded_keys = set()
        self.recording_action = action_name
        self.recording_btn = btn
        self.recording_var = string_var
        
        # We need a non-blocking way to listen for one combo
        # Simple approach: blocking keyboard.read_hotkey() freezes UI.
        # Better approach: Hook global, wait for key up.
        self.hook = keyboard.on_press(self._on_key_event)

    def _on_key_event(self, e):
        # This runs in a separate thread usually
        key = e.name
        
        # Filter modifier names to standard
        if key == 'control': key = 'ctrl'
        
        if key not in self.recorded_keys:
            self.recorded_keys.add(key)
        
        # We can try to construct hotkey string
        # Order matters: ctrl+alt+shift+key
        modifiers = [k for k in self.recorded_keys if k in ['ctrl', 'alt', 'shift', 'win']]
        remainder = [k for k in self.recorded_keys if k not in ['ctrl', 'alt', 'shift', 'win']]
        
        full_hotkey = "+".join(modifiers + remainder)
        
        # If it's a "complete" hotkey (modifiers + 1 key), we can stop. 
        # But this is tricky. Let's rely on users releasing keys?
        # For this simple implementation, let's just update the UI text live
        # and "finalize" when they release a key? No, safer to have them press ENTER? 
        # Or just take the first valid combo.
        
        # "Crosshair X" style usually listens until you stop pressing?
        # Let's try: If user presses keys, we update UI. Interaction ends when they click away?
        # No, let's use the keyboard.read_hotkey approach but inside a thread.
        pass 
        # Actually hooking is messy. Let's use prompts or focus binding.
        # RE-PLAN: Use a simplified "Dialog" approach for recording or just `keyboard.read_hotkey` in thread.
    
    # RE-IMPLEMENTING start_recording with robust hooks to avoid phantom keys
    def start_recording(self, btn, string_var, action_name):
        btn.config(text="Listening...", bg="#ff5500", fg="white")
        
        # Reset state
        self.rec_current_keys = set()
        self.rec_max_keys = set()
        self.rec_btn = btn
        self.rec_var = string_var
        
        # Disable hotkeys temporarily to prevent triggering while recording?
        # For now, just hook.
        try:
             self.rec_hook = keyboard.hook(self._on_record_event)
        except Exception as e:
             self.notify("Error", f"Failed to hook keyboard: {e}")
             self._finalize_recording("Error")

    def _on_record_event(self, e):
        if e.event_type == keyboard.KEY_DOWN:
            key = e.name.lower()
            if key in ['control', 'right control']: key = 'ctrl'
            if key in ['shift', 'right shift']: key = 'shift'
            if key in ['alt', 'right alt']: key = 'alt'
            if key in ['windows', 'left windows', 'right windows']: key = 'win'
            
            self.rec_current_keys.add(key)
            # Update max set if current is larger (to capture full combo)
            if len(self.rec_current_keys) > len(self.rec_max_keys):
                self.rec_max_keys = self.rec_current_keys.copy()
            
            # Live Update UI
            self.root.after(0, lambda: self.rec_btn.config(text="+".join(sorted(self.rec_current_keys))))

        elif e.event_type == keyboard.KEY_UP:
            key = e.name.lower()
            if key in ['control', 'right control']: key = 'ctrl'
            if key in ['shift', 'right shift']: key = 'shift'
            if key in ['alt', 'right alt']: key = 'alt'
            if key in ['windows', 'left windows', 'right windows']: key = 'win'
            
            if key in self.rec_current_keys:
                self.rec_current_keys.remove(key)
            
            # If all keys released, finalize
            if not self.rec_current_keys:
                self.root.after(0, self._finalize_recording)

    def _finalize_recording(self, error_val=None):
        try:
            keyboard.unhook(self.rec_hook)
        except:
            pass
            
        if error_val:
            final_hotkey = error_val
        elif not self.rec_max_keys:
            final_hotkey = self.rec_var.get() # Restore old if nothing pressed
        else:
            # Construct hotkey string from max_keys
            # Sort order: ctrl, alt, shift, win, others
            keys = list(self.rec_max_keys)
            priority = {'ctrl': 0, 'alt': 1, 'shift': 2, 'win': 3}
            
            mods = sorted([k for k in keys if k in priority], key=lambda x: priority[x])
            others = sorted([k for k in keys if k not in priority])
            
            final_hotkey = "+".join(mods + others)
            
        self.rec_var.set(final_hotkey)
        self.rec_btn.config(text=final_hotkey, bg="#222", fg="white")

    def save_settings(self):
        new_hotkeys = {}
        all_hotkeys_snapshot = {}
        
        # 1. Gather all inputs
        for action, var in self.hotkey_vars.items():
            all_hotkeys_snapshot[action] = var.get()

        # 2. Validate
        for action, hotkey in all_hotkeys_snapshot.items():
            valid, msg = HotkeyValidator.validate(hotkey, all_hotkeys_snapshot, ignore_action=action)
            if not valid:
                self.lbl_error.config(text=f"Error: {msg}")
                return
            new_hotkeys[action] = hotkey

        # 3. Save
        run_startup = self.var_startup.get()
        token = self.var_token.get().strip()
        
        self.config_manager.save_config({
            "hotkeys": new_hotkeys,
            "run_on_startup": run_startup,
            "ngrok_token": token
        })
        
        # 4. Apply
        if run_startup:
            self.ensure_startup()
        else:
            self.remove_startup()
            
        self.reload_hotkeys()
        
        self.settings_win.destroy()
        self.notify("Updated", "Settings saved successfully.")


    # --- THE CONVERTER LOGIC ---
    def trigger_converter(self):
        file_path = filedialog.askopenfilename(
            title="Select File to Convert",
            filetypes=[("All Supported", "*.docx *.pdf *.png *.jpg *.jpeg *.webp")]
        )
        if file_path:
            self.show_converter_ui(file_path)

    def show_converter_ui(self, file_path):
        ext = os.path.splitext(file_path)[1].lower()
        options = []
        
        # Logic Map
        if ext == '.docx': options = ['pdf']
        elif ext == '.pdf': options = ['docx', 'images']
        elif ext in ['.png', '.jpg', '.jpeg', '.webp']:
            options = ['pdf', 'png', 'jpg', 'webp']
            options = [o for o in options if f".{o}" != ext]
            
        if not options:
            self.notify("Error", "No conversion available.")
            return

        # Modern Floating Window
        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        popup.configure(bg='black')
        
        # Add a border frame
        border_color = "#ff5500" # Incineration Orange
        container = tk.Frame(popup, bg=border_color, padx=2, pady=2)
        container.pack(fill="both", expand=True)
        
        inner = tk.Frame(container, bg="#101010", padx=20, pady=20)
        inner.pack(fill="both", expand=True)
        
        # Position at Mouse
        x, y = self.root.winfo_pointerx(), self.root.winfo_pointery()
        popup.geometry(f"+{x}+{y}")

        tk.Label(inner, text="INCINERATE TO", bg='#101010', fg='#666', font=('Segoe UI', 9, 'bold')).pack(pady=(0, 10))

        for opt in options:
            btn = tk.Button(inner, text=opt.upper(), 
                      command=lambda o=opt: [popup.destroy(), self.run_conversion(o, file_path)],
                      bg='#202020', fg='white', 
                      activebackground=border_color, activeforeground='white',
                      relief='flat', padx=30, pady=8, font=('Segoe UI', 10))
            btn.pack(fill='x', pady=2)
            
            # Hover Effect
            def on_enter(e, b=btn): b['bg'] = '#333'
            def on_leave(e, b=btn): b['bg'] = '#202020'
            btn.bind("<Enter>", on_enter)
            btn.bind("<Leave>", on_leave)
            
        tk.Button(inner, text="CANCEL", command=popup.destroy, 
                  bg='#101010', fg='#444', activebackground='#101010', activeforeground='#888',
                  relief='flat', font=('Arial', 8)).pack(pady=(15, 0))
        
        popup.bind("<FocusOut>", lambda e: popup.destroy())
        popup.after(100, popup.focus_force)

    def run_conversion(self, target, path):
        # Run in background thread to keep UI smooth
        Thread(target=self._worker, args=(target, path)).start()

    def _worker(self, target, path):
        self.notify("Working...", f"Converting to {target}")
        try:
            if target == 'pdf':
                if path.endswith('.docx'): convert_docx_to_pdf(path)
                else: convert_image_to_pdf(path)
            elif target == 'docx': convert_pdf_to_docx(path)
            elif target == 'images': convert_pdf_to_images(path)
            elif target in ['png', 'jpg', 'webp']: convert_image_format(path, target)
            
            self.notify("Success", "File converted.")
        except Exception as e:
            self.notify("Error", str(e))

    # --- INCINERATOR DROP LOGIC ---
    def trigger_airdrop(self):
        # 1. Ask User: Send or Receive?
        # Simple popup choice
        popup = tk.Toplevel(self.root)
        popup.overrideredirect(True)
        popup.attributes('-topmost', True)
        popup.configure(bg='black')
        
        x, y = self.root.winfo_pointerx(), self.root.winfo_pointery()
        popup.geometry(f"+{x}+{y}")
        
        container = tk.Frame(popup, bg="#ff5500", padx=2, pady=2)
        container.pack(fill="both", expand=True)
        inner = tk.Frame(container, bg="#101010", padx=20, pady=20)
        inner.pack(fill="both", expand=True)
        
        tk.Label(inner, text="INCINERATOR DROP", bg='#101010', fg='#666', font=('Segoe UI', 9, 'bold')).pack(pady=(0, 10))

        # Global Mode Toggle
        var_global = tk.BooleanVar(value=False)
        cb_global = tk.Checkbutton(inner, text="Allow Global Access (Internet)", variable=var_global,
                                   bg='#101010', fg='#888', selectcolor='#222', activebackground='#101010', activeforeground='white',
                                   font=('Segoe UI', 9))
        cb_global.pack(pady=(0, 10))

        def start_send():
            popup.destroy()
            file_paths = filedialog.askopenfilenames(title="Select File(s) to Send")
            if file_paths:
                self.start_airdrop_session("SEND", file_paths, var_global.get())
                
        def start_recv():
            popup.destroy()
            self.start_airdrop_session("RECEIVE", None, var_global.get())

        tk.Button(inner, text="SEND FILE (PC -> PHONE)", command=start_send,
                  bg='#202020', fg='white', relief='flat', padx=30, pady=8, font=('Segoe UI', 10)).pack(fill='x', pady=2)
                  
        tk.Button(inner, text="RECEIVE FILE (PHONE -> PC)", command=start_recv,
                  bg='#202020', fg='white', relief='flat', padx=30, pady=8, font=('Segoe UI', 10)).pack(fill='x', pady=2)
                  
        tk.Button(inner, text="CANCEL", command=popup.destroy, 
                  bg='#101010', fg='#444', relief='flat', font=('Arial', 8)).pack(pady=(15, 0))
        
        popup.bind("<FocusOut>", lambda e: popup.destroy())
        popup.after(100, popup.focus_force)

    def start_airdrop_session(self, mode, file_path=None, use_global=False):
        try:
            token = self.config_manager.get('ngrok_token')
            url = self.airdrop.start_server(mode, file_path, use_global, token)
            # Use the actual wrapper method instead of direct implementation in start_server (fixing partial replacement issue)
            # Actually I made a mistake in replacement above, I defined _run_server_thread but didn't call it. 
            # Correction: start_server should call _run_server_thread.
            # Let me fix the start_server structure in the next tool call if needed or inline it here.
            # Wait, the previous chunk REPLACED the bottom of start_server with helper methods but didn't hook them up?
            # Ah, I replaced the BODY execution. 
            # Let's fix start_server to delegate to _run_server_thread in a separate fix step if I broke it.
            # Actually, looking at chunk 5, I replaced the END of start_server (helper defs) but didn't change the return.
            # I must fix start_server to actually use _run_server_thread.
            pass # Placeholder for thought process.
            
            # The previous chunk 5 defines `_run_server_thread` inside class scope or start_server scope?
            # It replaced `def run_app(): ... return` block.
            # So `start_server` now contains `_get_login_html` and `_run_server_thread` but doesn't CALL them?
            # This is risky. I should do a clean Replace of the whole method in next step.
            
            self.show_airdrop_qr(url, mode)
        except Exception as e:
            self.notify("Error", f"Failed to start: {e}")

    def show_airdrop_qr(self, url, mode):
        qr_img = self.airdrop.generate_qr(url)
        pin = self.airdrop.pin
        
        # Display QR
        qr_win = tk.Toplevel(self.root)
        qr_win.title("Scan Me")
        qr_win.attributes('-topmost', True)
        qr_win.geometry("350x450")
        qr_win.configure(bg="#121212")
        
        # Center on screen
        sw = qr_win.winfo_screenwidth()
        sh = qr_win.winfo_screenheight()
        qr_win.geometry(f"+{sw//2 - 175}+{sh//2 - 225}")
        
        tk.Label(qr_win, text="SCAN TO " + mode, bg="#121212", fg="#ff5500", font=("Segoe UI", 14, "bold")).pack(pady=(20, 5))
        
        # PIN Display
        tk.Label(qr_win, text=f"PIN: {pin}", bg="#121212", fg="white", font=("Consolas", 24, "bold")).pack(pady=(0, 10))
        
        # Convert PIL to PhotoImage
        from PIL import ImageTk
        photo = ImageTk.PhotoImage(qr_img)
        lbl_img = tk.Label(qr_win, image=photo, bg="#121212")
        lbl_img.image = photo # Keep reference
        lbl_img.pack()
        
        tk.Label(qr_win, text=url, bg="#121212", fg="#666", font=("Consolas", 10)).pack(pady=10)
        
        def close_stop():
            self.airdrop.stop_server()
            qr_win.destroy()
            
        tk.Button(qr_win, text="STOP SERVER", command=close_stop,
                  bg="#333", fg="white", relief="flat", padx=20, pady=10).pack(pady=20)
        
        qr_win.protocol("WM_DELETE_WINDOW", close_stop)

    def toggle_performance_mode(self):
        try:
            if self.performance_manager.is_active:
                self.performance_manager.disable()
                self.osd.show(text="Default", title="SYSTEM MODE", duration=2500)
            else:
                self.performance_manager.enable()
                self.osd.show(text="Performance Mode", title="SYSTEM MODE", duration=3000)
        except Exception as e:
            self.notify("Mode Error", str(e))
            print(f"Performance Toggle Failed: {e}")

    def clean_temp_files(self):
        self.notify("Incinerator", "Cleaning Temp Files...")
        
        temp_dir = os.environ.get('TEMP')
        if not temp_dir:
            self.notify("Error", "Could not find TEMP folder.")
            return

        deleted_files = 0
        deleted_dirs = 0
        errors = 0

        # Run in thread to prevent freezing UI
        def _clean_worker():
            nonlocal deleted_files, deleted_dirs, errors
            try:
                # Walk top-level to avoid deep recursion issues if any
                for item in os.listdir(temp_dir):
                    item_path = os.path.join(temp_dir, item)
                    try:
                        if os.path.isfile(item_path) or os.path.islink(item_path):
                            os.unlink(item_path)
                            deleted_files += 1
                        elif os.path.isdir(item_path):
                            shutil.rmtree(item_path)
                            deleted_dirs += 1
                    except Exception:
                        errors += 1
                
                self.safe_trigger(lambda: self.notify("Incinerator", f"Incinerated {deleted_files} files & {deleted_dirs} folders. ({errors} skipped)"))
            except Exception as e:
                self.safe_trigger(lambda: self.notify("Error", f"Failed to clean: {e}"))

        Thread(target=_clean_worker).start()

if __name__ == "__main__":
    AntigravityEngine()