import tkinter
from tkinter import filedialog
import customtkinter as ctk
from PIL import Image
import threading
import os
import sys
import logging
from pathlib import Path
import time
pd = None  # lazy-loaded
from datetime import datetime
import warnings
import queue
import json
import traceback
import tkinter as tk
import tkinter.messagebox as messagebox

from gui import Tooltip, SettingsWindow, MessageDialog, ContextMenu, AboutDialog
# Heavy modules will be imported lazily during initialization while native splash is visible
get_encryption_key = None
update_remote_files = None
load_nickname_map = None
smart_remap_columns_to_intended = None
parse_full_name_column = None
normalize_name = None
normalize_date = None
normalize_sex = None
normalize_city = None
load_raw_file = None
run_analysis = None
from config import HIDDEN_PASSWORD, PROVINCE_PROFILES, GLOBAL_CONFIG, THEME_COLORS, ThemeColor, APP_VERSION

if 'DEFAULT_PROVINCE' not in globals():
    DEFAULT_PROVINCE = "Oriental Mindoro"

# If built with PyInstaller --splash, update text ASAP; will close after UI init
_pyi_splash = None
try:
    import pyi_splash as _ps
    _pyi_splash = _ps
    try:
        _pyi_splash.update_text("Booting up..\nLoading...")
    except Exception:
        pass
except Exception:
    _pyi_splash = None

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

# Ensure unexpected exceptions are surfaced instead of silently closing the app
def _global_excepthook(exc_type, exc, tb):
    try:
        trace_text = ''.join(traceback.format_exception(exc_type, exc, tb))
        logging.error("Unhandled exception:\n" + trace_text)
        try:
            messagebox.showerror("Application Error", f"An unexpected error occurred.\n\n{exc}\n\nSee log for details.")
        except Exception:
            pass
    finally:
        # Do NOT call the default excepthook to avoid killing the app immediately
        # This allows Tk's mainloop to continue and the user to see the error dialog/log.
        pass

sys.excepthook = _global_excepthook

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    path = os.path.join(base_path, relative_path)
    # If the file doesn't exist in the expected location, try one level up (for PyInstaller)
    if not os.path.exists(path) and hasattr(sys, '_MEIPASS'):
        path = os.path.join(os.path.dirname(sys._MEIPASS), relative_path)
    
    return path

## Removed CTk Splash implementation (using native bootloader splash instead)

def lazy_import_heavy(progress_cb=None):
    """Import heavy modules after splash is shown to avoid late splash appearance."""
    global pd
    global get_encryption_key, update_remote_files, load_nickname_map
    global smart_remap_columns_to_intended, parse_full_name_column, normalize_name, normalize_date, normalize_sex, normalize_city
    global normalize_batch_name
    global load_raw_file, run_analysis

    try:
        if _pyi_splash:
            try: _pyi_splash.update_text("Booting up..\nLoading libraries...")
            except Exception: pass
        if progress_cb: progress_cb("Booting up..", "Loading libraries...", 12)
        if pd is None:
            import pandas as _pd
            pd = _pd
        if _pyi_splash:
            try: _pyi_splash.update_text("Booting up..\nLoading utilities...")
            except Exception: pass
        if progress_cb: progress_cb("Booting up..", "Loading utilities...", 18)
        if get_encryption_key is None:
            from data_utils import (
                get_encryption_key as _get_encryption_key,
                update_remote_files as _update_remote_files,
                load_nickname_map as _load_nickname_map,
                smart_remap_columns_to_intended as _smart_remap_columns_to_intended,
                parse_full_name_column as _parse_full_name_column,
                normalize_name as _normalize_name, normalize_date as _normalize_date,
                normalize_sex as _normalize_sex, normalize_city as _normalize_city,
                normalize_batch_name as _normalize_batch_name,
                load_raw_file as _load_raw_file
            )
            get_encryption_key = _get_encryption_key
            update_remote_files = _update_remote_files
            load_nickname_map = _load_nickname_map
            smart_remap_columns_to_intended = _smart_remap_columns_to_intended
            parse_full_name_column = _parse_full_name_column
            normalize_name = _normalize_name
            normalize_date = _normalize_date
            normalize_sex = _normalize_sex
            normalize_city = _normalize_city
            normalize_batch_name = _normalize_batch_name
            load_raw_file = _load_raw_file
        if _pyi_splash:
            try: _pyi_splash.update_text("Booting up..\nPreparing engine...")
            except Exception: pass
        if progress_cb: progress_cb("Booting up..", "Preparing engine...", 24)
        if run_analysis is None:
            from analysis_engine import run_analysis as _run_analysis
            run_analysis = _run_analysis
        if _pyi_splash:
            try: _pyi_splash.update_text("Booting up..\nLibraries ready")
            except Exception: pass
        if progress_cb: progress_cb("Booting up..", "Libraries ready", 26)
    except Exception as e:
        logging.error("Lazy import failed: %s", e, exc_info=True)

class AppData:
    def __init__(self, province_name):
        self.data_dir = Path.home() / ".splink_master_checker"
        self.data_dir.mkdir(exist_ok=True)
        
        # Get the province config, defaulting to Oriental Mindoro if not found
        province_config = PROVINCE_PROFILES.get(province_name, PROVINCE_PROFILES["Oriental Mindoro"])
        
        # Use dot notation to access dataclass attributes
        master_filename = os.path.basename(province_config.urls.master_db)
        officials_filename = os.path.basename(province_config.urls.officials)
        
        self.master_db_path = self.data_dir / master_filename
        self.master_db_meta_path = self.data_dir / f"{master_filename}.meta"
        self.officials_db_path = self.data_dir / officials_filename
        self.officials_db_meta_path = self.data_dir / f"{officials_filename}.meta"
        
        self.nickname_path = self.data_dir / "Nicknames.csv"
        self.nickname_meta_path = self.data_dir / "Nicknames.csv.meta"
        
        self.window_prefs_path = self.data_dir / "window_preferences.json"
    
    def load_window_preferences(self):
        """Load saved window size and position"""
        default_prefs = {
            "width": 520,
            "height": 580,
            "x": None,
            "y": None
        }
        
        if self.window_prefs_path.exists():
            try:
                with open(self.window_prefs_path, 'r') as f:
                    prefs = json.load(f)
                    return {**default_prefs, **prefs}
            except (json.JSONDecodeError, IOError):
                pass
        
        return default_prefs
    
    def save_window_preferences(self, width, height, x, y):
        """Save current window size and position"""
        prefs = {
            "width": width,
            "height": height,
            "x": x,
            "y": y
        }
        
        try:
            with open(self.window_prefs_path, 'w') as f:
                json.dump(prefs, f, indent=2)
        except IOError:
            pass

    def get_last_updated_str(self, file_path):
        if not file_path.exists(): return "Never"
        try:
            return time.strftime('%Y-%m-%d %H:%M', time.localtime(file_path.stat().st_mtime))
        except FileNotFoundError:
            return "Never"

class MasterCheckerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.province_name = DEFAULT_PROVINCE
        self.province_config = PROVINCE_PROFILES[self.province_name]
        
        self.title(self.province_config.title)
        
        ctk.set_appearance_mode("System")
        self.theme_name = self.province_config.theme
        self.theme_colors = THEME_COLORS.get(self.theme_name, THEME_COLORS[ThemeColor.BLUE])
        
        # Keep window hidden; finish init after heavy imports in a background thread
        self.withdraw()
        threading.Thread(target=self._background_init, daemon=True).start()

    def _background_init(self):
        """Run heavy imports in a worker thread while native splash is visible, then finish UI init."""
        try:
            lazy_import_heavy(None)
        except Exception as e:
            logging.error("Background lazy import failed: %s", e, exc_info=True)
        finally:
            # Continue UI initialization on the Tk main thread
            self.after(0, self._finish_init)

    def _finish_init(self):
        try:
            self.app_data = AppData(self.province_name)
            
            default_width, default_height = 520, 580
            app_width, app_height = default_width, default_height
            
            self.update_idletasks()
            x = (self.winfo_screenwidth() / 2) - (app_width / 2)
            y = (self.winfo_screenheight() / 2) - (app_height / 2)
            
            self.geometry(f"{app_width}x{app_height}+{int(x)}+{int(y)}")
            self.minsize(480, 400)

            try:
                logo_path = resource_path("logo.ico")
                self.logo_image = ctk.CTkImage(Image.open(logo_path), size=(20, 20))
                self.iconbitmap(logo_path)
            except Exception as e:
                self.logo_image = None
                logging.warning(f"Could not load logo.png: {e}")

            self.grid_columnconfigure(0, weight=1)
            self.grid_rowconfigure(0, weight=1)

            self.content_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
            self.content_frame.grid(row=0, column=0, sticky="nsew")
            
            self.content_frame.grid_columnconfigure(0, weight=1)
            self.content_frame.grid_rowconfigure(2, weight=1)

            self.user_filepath = None
            self.encryption_key = get_encryption_key("doleadmin")
            self.analysis_queue = None

            self._create_widgets()
            self.log_message("Welcome! Load your file to begin analysis.")
            self.update_status("main", "Ready. Please load a file.")
            self.update_status("db")
            self.update_status("nickname")
            self.update_status("officials")
            
            self.active_toplevel = None
            self._current_theme = ctk.get_appearance_mode()
            self.after(500, self._check_appearance_mode)
        except Exception as _init_err:
            logging.error("Startup error: %s", _init_err, exc_info=True)
            try:
                messagebox.showerror("Startup Error", f"The app failed to initialize.\n\n{_init_err}")
            except Exception:
                pass
        finally:
            # Show main window and close native splash
            self.deiconify()
            if _pyi_splash:
                try:
                    _pyi_splash.close()
                except Exception:
                    pass
            self.lift()
            self.focus_force()
            # Safety watchdog: ensure main window is visible even if something hid it
            self.after(2000, self._ensure_shown)

    def _check_appearance_mode(self):
        new_theme = ctk.get_appearance_mode()
        if new_theme != self._current_theme:
            self._current_theme = new_theme
            if self.active_toplevel and self.active_toplevel.winfo_exists():
                self.log_message("System theme changed. Closing open dialog...")
                self.active_toplevel.destroy()
        
        self.after(500, self._check_appearance_mode)

    def _ensure_shown(self):
        try:
            if not self.winfo_viewable():
                self.deiconify()
            self.lift()
            self.focus_force()
        except Exception:
            pass

    def _create_widgets(self):
        fg_color = (self.theme_colors.fg_color[0], self.theme_colors.fg_color[1])
        hover_color = (self.theme_colors.hover_color[0], self.theme_colors.hover_color[1])
        
        top_frame = ctk.CTkFrame(self.content_frame)
        top_frame.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")
        top_frame.grid_columnconfigure(1, weight=1)

        self.button_a = ctk.CTkButton(top_frame, text="Load Your File", command=self.select_user_file, width=135, fg_color=fg_color, hover_color=hover_color)
        self.button_a.grid(row=0, column=0, padx=10, pady=10)

        self.label_a = ctk.CTkLabel(top_frame, text="No file selected.")
        self.label_a.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        self.label_a_tooltip = Tooltip(self.label_a, "No file selected.")

        self.settings_button = ctk.CTkButton(top_frame, text="Tools & Resources", command=self.open_settings_window, width=135, fg_color=fg_color, hover_color=hover_color)
        self.settings_button.grid(row=0, column=2, padx=10, pady=10)

        self.run_button = ctk.CTkButton(self.content_frame, text="Run Analysis", command=self.run_process, state="disabled", fg_color=fg_color, hover_color=hover_color)
        self.run_button.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        # Bind Enter to trigger Run when enabled
        self.bind("<Return>", self._on_enter_pressed)
        
        self.progress_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.progress_frame.grid_columnconfigure(0, weight=1)

        self.progress_text = ctk.CTkLabel(self.progress_frame, text="Starting analysis...")
        self.progress_text.grid(row=0, column=0, padx=10, pady=(0,2), sticky="w")

        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, progress_color=fg_color)
        self.progress_bar.set(0)
        self.progress_bar.grid(row=1, column=0, padx=10, sticky="ew")

        self.log_textbox = ctk.CTkTextbox(self.content_frame, state="disabled", wrap="word", font=ctk.CTkFont(family="Courier New", size=13))
        self.log_textbox.grid(row=2, column=0, padx=10, pady=(5, 0), sticky="nsew")

        def prevent_double_click(event):
            self.log_textbox.tag_remove("sel", "1.0", "end")
            return "break"
        
        def prevent_selection(event):
            return "break"
        
        self.log_textbox.bind("<Double-Button-1>", prevent_double_click)
        self.log_textbox.bind("<Double-Button-2>", prevent_double_click)
        self.log_textbox.bind("<Double-Button-3>", prevent_double_click)
        self.log_textbox.bind("<Button-1>", prevent_double_click)
        self.log_textbox.bind("<B1-Motion>", prevent_double_click)
        
        self.progress_bar.bind("<Double-Button-1>", prevent_selection)
        self.progress_bar.bind("<Button-1>", prevent_selection)
        self.progress_bar.bind("<B1-Motion>", prevent_selection)
        self.progress_frame.bind("<Double-Button-1>", prevent_selection)
        self.progress_frame.bind("<Button-1>", prevent_selection)
        self.progress_frame.bind("<B1-Motion>", prevent_selection)
        
        self.log_textbox.configure(cursor="arrow")
        self.log_textbox._textbox.configure(insertwidth=0)
        
        self.report_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.report_frame.grid_columnconfigure(0, weight=1)

        self.report_path_label = ctk.CTkLabel(self.report_frame, text="", anchor="w")
        self.report_path_label.grid(row=0, column=0, padx=(10, 5), pady=5, sticky="ew")
        self.report_path_tooltip = Tooltip(self.report_path_label, "")

        self.open_report_button = ctk.CTkButton(self.report_frame, text="Open Report",
                                                fg_color="transparent", border_width=1,
                                                text_color=("gray10", "gray90"),
                                                border_color=("gray70", "gray30"), hover_color=hover_color)
        self.open_report_button.grid(row=0, column=1, padx=(5, 10), sticky="e")

        status_frame = ctk.CTkFrame(self.content_frame)
        status_frame.grid(row=4, column=0, padx=10, pady=(5, 5), sticky="ew")
        status_frame.grid_columnconfigure(1, weight=1)
        self.status_text_label = ctk.CTkLabel(status_frame, text="", font=("Arial", 12))
        self.status_text_label.grid(row=0, column=0, padx=(10, 0), sticky="w")
        self.db_status_label = ctk.CTkLabel(status_frame, text="", font=("Arial", 10), text_color="gray60")
        self.db_status_label.grid(row=0, column=4, padx=(10, 10), sticky="e")
        self.nickname_status_label = ctk.CTkLabel(status_frame, text="", font=("Arial", 10), text_color="gray60")
        self.nickname_status_label.grid(row=0, column=3, padx=(10, 0), sticky="e")
        self.officials_status_label = ctk.CTkLabel(status_frame, text="", font=("Arial", 10), text_color="gray60")
        self.officials_status_label.grid(row=0, column=2, padx=(10, 0), sticky="e")
        
        self.log_context_menu = ContextMenu(self)
        self.log_context_menu.add_command(label="Copy All", command=self.copy_log_to_clipboard)
        self.log_context_menu.add_separator()
        self.log_context_menu.add_command(label="Clear Log", command=self.clear_log)
        self.log_context_menu.add_separator()
        self.log_context_menu.add_command(label="Reset Window Size", command=self.reset_window_size)
        self.log_context_menu.add_separator()
        self.log_context_menu.add_command(label="About", command=self.show_about)
        
        self.log_textbox.bind("<Button-3>", self.log_context_menu.show)

    def copy_log_to_clipboard(self):
        full_log_text = self.log_textbox.get("1.0", "end-1c")
        if full_log_text:
            self.clipboard_clear()
            self.clipboard_append(full_log_text)
            self.update_status("main", "Log content copied to clipboard.")
        else:
            self.update_status("main", "Log is empty.")

    def clear_log(self):
        self.log_textbox.configure(state="normal")
        self.log_textbox.delete("1.0", "end")
        self.log_textbox.configure(state="disabled")
        self.update_status("main", "Log cleared.")

    def show_about(self):
        if self.active_toplevel is not None and self.active_toplevel.winfo_exists():
            self.active_toplevel.lift()
            return
        app_title = f"Deduplication and Official Linkage Engine v{APP_VERSION}"
        description = "A high-speed data integrity tool for identifying duplicates and linking records to official DOLE databases."
        capabilities_list = [
            "Intelligent Matching Engine: Handles typos, name variations, and nicknames while automatically cleaning data.",
            "High-Speed Parallel Processing: Uses modern multi-core CPUs to deliver fast results on large datasets.",
            "Live Database Synchronization: Ensures every analysis is performed against the most current masterfiles.",
            "Automated PDF & Excel Reporting: Instantly generates detailed reports with a summary dashboard."
        ]
        footer = "Powered by a custom Python engine for performance and accuracy."
        credits = "Programmed by A. Enage (aenage@gmail.com)\nDOLE OrMin Provincial Office"
        AboutDialog(self,
                    title_text=app_title,
                    desc_text=description,
                    capabilities=capabilities_list,
                    footer_text=footer,
                    credits_text=credits,
                    icon_image=self.logo_image,
                    theme_colors=self.theme_colors)
    
    def on_window_configure(self, event):
        if event.widget == self:
            if hasattr(self, '_save_timer'):
                self.after_cancel(self._save_timer)
            self._save_timer = self.after(500, lambda: self.app_data.save_window_preferences(
                self.winfo_width(), self.winfo_height(), self.winfo_x(), self.winfo_y()
            ))
    
    def reset_window_size(self):
        if self.state() == 'zoomed':
            self.state('normal')

        default_width, default_height = 520, 580
        
        self.update_idletasks() 
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        x = (screen_width / 2) - (default_width / 2)
        y = (screen_height / 2) - (default_height / 2)
        
        self.geometry(f"{default_width}x{default_height}+{int(x)}+{int(y)}")
        
        self.update_status("main", "Window size reset to default.")
    
    def _open_file_path(self, path_to_open):
        try:
            if sys.platform == "win32": os.startfile(os.path.normpath(path_to_open))
            elif sys.platform == "darwin": os.system(f'open "{os.path.normpath(path_to_open)}"')
            else: os.system(f'xdg-open "{os.path.normpath(path_to_open)}"')
        except Exception as e:
            self.log_message(f"❌ Could not open file: {e}")

    def update_status(self, part, text=None, state="default"):
        if part == "main":
            colors = {"default": ("gray10", "gray90"), "running": ("#FFA500", "#FF8C00"), "success": ("#2E7D32", "#66BB6A"), "error": ("#D32F2F", "#E57373")}
            self.status_text_label.configure(text=text, text_color=colors.get(state, colors["default"]))
        elif part == "db": self.db_status_label.configure(text=f"DB: {self.app_data.get_last_updated_str(self.app_data.master_db_path)}")
        elif part == "nickname": self.nickname_status_label.configure(text=f"Nick: {self.app_data.get_last_updated_str(self.app_data.nickname_path)}")
        elif part == "officials": self.officials_status_label.configure(text=f"Off: {self.app_data.get_last_updated_str(self.app_data.officials_db_path)}")
        self.update_idletasks()

    def log_message(self, message):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", message + "\n\n")
        self.log_textbox.configure(state="disabled")
        def prevent_double_click(event):
            self.log_textbox.tag_remove("sel", "1.0", "end")
            return "break"
        self.log_textbox.bind("<Double-Button-1>", prevent_double_click)
        self.log_textbox.bind("<Double-Button-2>", prevent_double_click)
        self.log_textbox.bind("<Double-Button-3>", prevent_double_click)
        self.log_textbox.see("end")
        self.update_idletasks()

    def log_final_report_path(self, path):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", "\n")
        self.log_textbox.configure(state="disabled")
        display_path = path if len(path) <= 45 else f"...{path[-42:]}"
        self.report_path_label.configure(text=f"✅ Done! Report saved to: {display_path}")
        self.report_path_tooltip.update_text(path)
        self.open_report_button.configure(command=lambda p=path: self._open_file_path(p))
        self.report_frame.grid(row=3, column=0, padx=10, pady=0, sticky="ew")
        self.log_textbox.see("end")

    def select_user_file(self):
        path = filedialog.askopenfilename(title="Select Your Data File", filetypes=(("Supported Files", "*.csv *.xlsx *.xls *.txt"), ("All files", "*.*")))
        if not path: return
        
        path = os.path.normpath(path)
        
        try:
            if os.path.getsize(path) > 100 * 1024 * 1024:
                MessageDialog(self, title="File Too Large", message="The selected file is larger than 100 MB and cannot be processed.", icon_image=self.logo_image)
                return
        except OSError:
            self.log_message("❌ Could not access the selected file.")
            return
        self.user_filepath = path
        display_path = path if len(path) <= 40 else f"...{path[-37:]}"
        self.label_a.configure(text=display_path)
        self.label_a_tooltip.update_text(path)
        self.log_message("File selected. Click 'Run Analysis' to start.")
        self.update_status("main", "Ready to run analysis.")
        self.run_button.configure(state="normal")
        # Autofocus Run button so Enter runs immediately
        self.after(100, self.run_button.focus_force)

    def open_settings_window(self):
        if self.active_toplevel is not None and self.active_toplevel.winfo_exists():
            self.active_toplevel.lift()
            return
            
        self.log_message("Opening Tools & Resources...")
        SettingsWindow(self, icon_image=self.logo_image, theme_colors=self.theme_colors)

    def run_process(self):
        if not self.user_filepath:
            self.log_message("Cannot run. Select a file first.")
            return
        self.report_frame.grid_forget()
        self.run_button.configure(state="disabled")
        self.button_a.configure(state="disabled")
        self.settings_button.configure(state="disabled")
        
        self.run_button.grid_forget()
        self.progress_frame.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.progress_bar.configure(mode="indeterminate")
        self.progress_bar.start()
        self.progress_text.configure(text="Starting analysis...")
        
        self.update_status("main", "Running analysis...", "running")
        # Convert ProvinceURLs dataclass to dict expected by update_remote_files
        province_urls = {
            "master_db": self.province_config.urls.master_db,
            "officials": self.province_config.urls.officials,
        }
        global_urls = GLOBAL_CONFIG

        self.analysis_queue = queue.Queue()
        self.after(100, self.check_progress_queue)
        
        threading.Thread(
            target=self.process_in_thread, 
            args=(province_urls, global_urls, self.analysis_queue),
            daemon=True
        ).start()

    def check_progress_queue(self):
        try:
            message = self.analysis_queue.get_nowait()
            msg_type = message[0]
            
            if msg_type == "indeterminate":
                status_text = message[1]
                self.progress_text.configure(text=status_text)

            elif msg_type == "determinate":
                if self.progress_bar.cget("mode") == "indeterminate":
                    self.progress_bar.stop()
                    self.progress_bar.configure(mode="determinate")
                
                progress_value, status_text = message[1], message[2]
                percentage = int(progress_value * 100)
                display_text = f"{status_text} ({percentage}%)"
                
                self.progress_bar.set(progress_value)
                self.progress_text.configure(text=display_text)
            
        except queue.Empty:
            pass
        finally:
            if self.run_button.cget("state") == "disabled":
                self.after(100, self.check_progress_queue)

    def process_in_thread(self, province_urls, global_urls, progress_queue):
        start_time = datetime.now()
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", UserWarning)

                has_internet = update_remote_files(self.app_data, self.encryption_key, self.log_message, province_urls, global_urls)
                self.update_status("db"); self.update_status("officials"); self.update_status("nickname")
                nickname_map = load_nickname_map(self.app_data, self.encryption_key, self.log_message)
                master_df, officials_df = None, None
                
                if self.app_data.master_db_path.exists():
                    try:
                        master_df = load_raw_file(self.app_data.master_db_path, self.encryption_key)
                        self.log_message(f"✅ [MasterDB] Loaded {len(master_df)} records from cache.")
                    except Exception as e: self.log_message(f"⚠️ [MasterDB] Could not load from cache: {e}.")
                
                if self.app_data.officials_db_path.exists():
                    try:
                        officials_df = load_raw_file(self.app_data.officials_db_path, self.encryption_key)
                        self.log_message(f"✅ [OfficialsDB] Loaded {len(officials_df)} records from cache.")
                    except Exception as e: self.log_message(f"⚠️ [OfficialsDB] Could not load from cache: {e}.")
                
                user_df = pd.DataFrame()
                if os.path.getsize(self.user_filepath) > 0:
                    
                    original_stderr = sys.stderr
                    devnull = open(os.devnull, 'w')
                    try:
                        file_ext = os.path.splitext(self.user_filepath)[1].lower()
                        if file_ext in ['.csv', '.txt']:
                            sys.stderr = devnull
                            sep = ',' if file_ext == '.csv' else '\t'
                            user_df = pd.read_csv(self.user_filepath, sep=sep, dtype=str, engine='python', on_bad_lines='warn').dropna(how='all')
                        elif file_ext in ['.xlsx', '.xls']:
                            user_df = pd.read_excel(self.user_filepath, dtype=str).dropna(how='all')
                    finally:
                        sys.stderr = original_stderr
                        devnull.close()

                dfs = {"user": user_df, "master": master_df, "officials": officials_df}
                for name, df in dfs.items():
                    if df is None or df.empty: continue
                    is_officials = (name == "officials")
                    cleaned_df = smart_remap_columns_to_intended(df, is_officials_file=is_officials)
                    cleaned_df = parse_full_name_column(cleaned_df)
                    for col in cleaned_df.columns:
                        if col in ["First Name", "Middle Name", "Last Name", "Suffix", "Position", "Barangay"]:
                            cleaned_df[col] = cleaned_df[col].apply(normalize_name)
                        elif col == "City":
                            cleaned_df[col] = cleaned_df[col].apply(normalize_city)
                        elif col == "Sex":
                            cleaned_df[col] = cleaned_df[col].apply(normalize_sex)
                        elif col == "Birthdate":
                            cleaned_df[col] = cleaned_df[col].apply(normalize_date)
                        elif col == "Contact Number":
                            cleaned_df[col] = cleaned_df[col].apply(lambda x: str(x).strip() if pd.notna(x) else '')
                        elif col == "Batch Name":
                            cleaned_df[col] = cleaned_df[col].apply(normalize_batch_name)
                    dfs[name] = cleaned_df
                
                user_df, master_df, officials_df = dfs["user"], dfs["master"], dfs["officials"]
                self.log_message(f"✅ [UserFile] Loaded and Cleaned {len(user_df)} records.")
            
            run_analysis(
                user_df, master_df, officials_df, nickname_map, 
                self.user_filepath, 
                self.province_name,  
                self.log_message, self.update_status, start_time, self.log_final_report_path,
                progress_queue
            )

        except Exception as e:
            self.log_message(f"❌ An unexpected error occurred: {e}")
            self.update_status("main", "An error occurred.", "error")
            logging.error("Error in process_in_thread", exc_info=True)
        finally:
            self.after(100, self.enable_buttons)

    def enable_buttons(self):
        self.progress_frame.grid_forget()
        self.progress_bar.stop()
        self.run_button.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        
        self.run_button.configure(state="normal")
        self.button_a.configure(state="normal")
        self.settings_button.configure(state="normal")

    def _on_enter_pressed(self, event=None):
        """If Run is enabled, pressing Enter will start analysis."""
        try:
            if self.run_button.cget("state") == "normal":
                # Use invoke to mimic a real button click
                self.run_button.invoke()
        except Exception:
            pass

if __name__ == "__main__":
    from multiprocessing import freeze_support
    freeze_support()
    app = MasterCheckerApp()
    app.mainloop()