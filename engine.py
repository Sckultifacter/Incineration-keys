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

import json

# =========================================================
#  ZONE 0: CONFIGURATION MANGER & VALIDATOR
# =========================================================
class ConfigManager:
    DEFAULT_CONFIG = {
        "hotkeys": {
            "Trigger Converter": "ctrl+alt+z"
        },
        "run_on_startup": True
    }
    
    def __init__(self):
        self.config_path = self.get_config_path()
        self.config = self.load_config()
        # Migrate old config if needed
        if 'hotkey' in self.config:
            self.config['hotkeys'] = {"Trigger Converter": self.config['hotkey']}
            del self.config['hotkey']
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
                # Ensure structure integrity
                if 'hotkeys' not in data:
                    data['hotkeys'] = self.DEFAULT_CONFIG['hotkeys'].copy()
                return data
        except:
            return self.DEFAULT_CONFIG.copy()

    def save_config(self, new_config):
        self.config.update(new_config)
        with open(self.config_path, 'w') as f:
            json.dump(self.config, f, indent=4)

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

def convert_image_format(input_path, target_ext):
    from PIL import Image
    image = Image.open(input_path)
    if target_ext in ['jpg', 'jpeg'] and image.mode == 'RGBA':
        image = image.convert('RGB')
    output_path = os.path.splitext(input_path)[0] + f".{target_ext}"
    image.save(output_path)
    return output_path

# =========================================================
#  ZONE 3: THE CORE ENGINE
# =========================================================

class AntigravityEngine:
    def __init__(self):
        self.config_manager = ConfigManager()
        
        # 1. Startup Check
        if self.config_manager.get('run_on_startup'):
            self.ensure_startup()
        else:
            self.remove_startup()
        
        # 2. Setup Hidden UI (Required for dialogs)
        self.root = tk.Tk()
        self.root.withdraw()
        
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
                if existing == app_path:
                    winreg.CloseKey(key)
                    return # Already installed
            except FileNotFoundError:
                pass

            winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, app_path)
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
        sys.exit(0)

    def safe_trigger(self, func):
        self.root.after(0, func)

    def notify(self, title, msg):
        try:
            notification.notify(title=title, message=msg, timeout=3)
        except:
            pass

    # --- ADVANCED SETTINGS UI ---
    def open_settings_window(self):
        if hasattr(self, 'settings_win') and self.settings_win.winfo_exists():
            self.settings_win.focus_force()
            return

        self.settings_win = tk.Toplevel(self.root)
        self.settings_win.title("Incineration Settings")
        self.settings_win.geometry("500x450")
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

        # Error Label
        self.lbl_error = tk.Label(self.settings_win, text="", bg="#121212", fg="red", font=("Segoe UI", 9))
        self.lbl_error.pack(pady=(10, 0))

        # Action Buttons
        btn_frame = tk.Frame(self.settings_win, bg="#121212")
        btn_frame.pack(side='bottom', pady=30)
        
        tk.Button(btn_frame, text="APPLY CHANGES", command=self.save_settings, 
                  bg="#ff5500", fg="white", relief="flat", font=("Segoe UI", 10, "bold"), padx=25, pady=8).pack(side="left", padx=10)
        tk.Button(btn_frame, text="CANCEL", command=self.settings_win.destroy, 
                  bg="#222", fg="#aaa", relief="flat", font=("Segoe UI", 10), padx=25, pady=8).pack(side="left", padx=10)

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
    
    # RE-IMPLEMENTING start_recording safer
    def start_recording(self, btn, string_var, action_name):
        btn.config(text="Listening...", bg="#ff5500", fg="white")
        Thread(target=self._record_thread, args=(btn, string_var, action_name)).start()

    def _record_thread(self, btn, string_var, action_name):
        try:
            # This blocks until a hotkey is pressed
            hotkey = keyboard.read_hotkey(suppress=False)
            # Update UI from thread
            self.root.after(0, lambda: self._finish_recording(btn, string_var, hotkey))
        except Exception:
            self.root.after(0, lambda: self._finish_recording(btn, string_var, "Error"))

    def _finish_recording(self, btn, string_var, hotkey):
        btn.config(bg="#222", fg="white")
        string_var.set(hotkey)

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
        self.config_manager.save_config({
            "hotkeys": new_hotkeys,
            "run_on_startup": run_startup
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

if __name__ == "__main__":
    AntigravityEngine()