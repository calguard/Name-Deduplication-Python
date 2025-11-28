import customtkinter as ctk
import tkinter
import os
import sys
import glob

import config
from config import GLOBAL_CONFIG
from data_utils import check_internet, smart_download_pat, get_auth_headers

class Tooltip:
    def __init__(self, widget, text):
        self.widget, self.text, self.tooltip_window, self.id = widget, text, None, None
        self.widget.bind("<Enter>", self.schedule_show)
        self.widget.bind("<Leave>", self.hide)
    def schedule_show(self, event=None): self.id = self.widget.after(500, self.show)
    def show(self):
        if self.tooltip_window or not self.text: return
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25; y += self.widget.winfo_rooty() + 25
        self.tooltip_window = ctk.CTkToplevel(self.widget)
        self.tooltip_window.wm_overrideredirect(True); self.tooltip_window.wm_geometry(f"+{x}+{y}")
        ctk.CTkLabel(self.tooltip_window, text=self.text, fg_color=("#f0f0f0", "#2b2b2b"), corner_radius=5, padx=8, pady=4).pack()
    def hide(self, event=None):
        if self.id: self.widget.after_cancel(self.id)
        if self.tooltip_window: self.tooltip_window.destroy()
        self.tooltip_window, self.id = None, None
    def update_text(self, new_text): self.text = new_text

class ContextMenu(tkinter.Menu):
    def __init__(self, master, **kwargs):
        super().__init__(master, tearoff=0, **kwargs)
        self.configure(
            font=("Segoe UI", 12),
            relief="raised",
            borderwidth=2
        )

    def show(self, event):
        try:
            self.tk_popup(event.x_root, event.y_root)
        finally:
            self.grab_release()

class CustomToplevel(ctk.CTkToplevel):
    def __init__(self, master, title="", icon_image=None):
        super().__init__(master)
        self.overrideredirect(True); self.attributes("-topmost", True); self.transient(master)
        self.grid_columnconfigure(0, weight=1); self.grid_rowconfigure(1, weight=1)
        
        if hasattr(self.master, 'active_toplevel'):
            self.master.active_toplevel = self

        title_bar = ctk.CTkFrame(self, corner_radius=0, fg_color=ctk.ThemeManager.theme["CTkFrame"]["fg_color"])
        title_bar.grid(row=0, column=0, sticky="ew"); title_bar.grid_columnconfigure(1, weight=1)
        
        if icon_image:
            icon_label = ctk.CTkLabel(title_bar, image=icon_image, text="")
            icon_label.grid(row=0, column=0, padx=(10, 5), pady=5)
            icon_label.bind("<ButtonPress-1>", self.start_move); icon_label.bind("<ButtonRelease-1>", self.stop_move); icon_label.bind("<B1-Motion>", self.do_move)
        
        title_label = ctk.CTkLabel(title_bar, text=title, font=ctk.CTkFont(weight="bold"), text_color=ctk.ThemeManager.theme["CTkLabel"]["text_color"])
        title_label.grid(row=0, column=1, padx=5, pady=5, sticky="w")
        
        close_button = ctk.CTkButton(
            title_bar, text="✕", width=40, height=28, 
            font=ctk.CTkFont(size=16), 
            text_color=("black", "white"), 
            fg_color="transparent", 
            hover_color="#c42b1c", 
            command=self.destroy
        )
        close_button.grid(row=0, column=2, sticky="e")
        
        title_bar.bind("<ButtonPress-1>", self.start_move); title_bar.bind("<ButtonRelease-1>", self.stop_move); title_bar.bind("<B1-Motion>", self.do_move)
        title_label.bind("<ButtonPress-1>", self.start_move); title_label.bind("<ButtonRelease-1>", self.stop_move); title_label.bind("<B1-Motion>", self.do_move)
        self.content_frame = ctk.CTkFrame(self, corner_radius=0)
        self.content_frame.grid(row=1, column=0, sticky="nsew")
        
    def start_move(self, event): self.x, self.y = event.x, event.y
    def stop_move(self, event): self.x, self.y = None, None
    def do_move(self, event): self.geometry(f"+{self.winfo_x() + event.x - self.x}+{self.winfo_y() + event.y - self.y}")
    
    def destroy(self):
        if hasattr(self.master, 'active_toplevel') and self.master.active_toplevel == self:
            self.master.active_toplevel = None
        super().destroy()

    def center_on_master(self, width, height):
        self.update_idletasks()
        x = self.master.winfo_x() + (self.master.winfo_width() / 2) - (width / 2)
        y = self.master.winfo_y() + (self.master.winfo_height() / 2) - (height / 2)
        self.geometry(f"{width}x{height}+{int(x)}+{int(y)}")

class PasswordDialog(CustomToplevel):
    def __init__(self, master, title="Admin Access", icon_image=None):
        super().__init__(master, title=title, icon_image=icon_image)
        self.password = None; self.protocol("WM_DELETE_WINDOW", self._on_cancel); self.after(250, lambda: self.entry.focus_force())
        self.content_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.content_frame, text="Enter Admin Password:").grid(row=0, column=0, columnspan=2, padx=20, pady=(20, 10))
        self.entry = ctk.CTkEntry(self.content_frame, show="*"); self.entry.grid(row=1, column=0, columnspan=2, padx=20, pady=5, sticky="ew")
        self.entry.bind("<Return>", self._on_ok)
        button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent"); button_frame.grid(row=2, column=0, columnspan=2, pady=(10, 20))
        # Theme-aware buttons
        theme = getattr(self.master, 'theme_colors', None)
        if theme:
            ok_fg = (theme.fg_color[0], theme.fg_color[1])
            ok_hover = (theme.hover_color[0], theme.hover_color[1])
        else:
            ok_fg = ctk.ThemeManager.theme["CTkButton"]["fg_color"]
            ok_hover = ctk.ThemeManager.theme["CTkButton"]["hover_color"]

        ctk.CTkButton(button_frame, text="OK", command=self._on_ok, width=100, fg_color=ok_fg, hover_color=ok_hover).pack(side="left", padx=5)
        ctk.CTkButton(button_frame, text="Cancel", command=self._on_cancel, width=100, fg_color="transparent", border_width=1, text_color=("gray10","gray90"), border_color=("gray70","gray30"), hover_color=ok_hover).pack(side="left", padx=5)
        self.center_on_master(width=320, height=180); self.grab_set()
    def _on_ok(self, event=None): self.password = self.entry.get(); self.destroy()
    def _on_cancel(self): self.password = None; self.destroy()
    def get_password(self): self.master.wait_window(self); return self.password

class MessageDialog(CustomToplevel):
    def __init__(self, master, title, message, icon_image=None):
        super().__init__(master, title=title, icon_image=icon_image)
        self.content_frame.grid_columnconfigure(0, weight=1); self.content_frame.grid_rowconfigure(0, weight=1)
        
        message_label = ctk.CTkLabel(self.content_frame, text=message, wraplength=350)
        message_label.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        
        # Theme-aware OK button
        theme = getattr(self.master, 'theme_colors', None)
        if theme:
            ok_fg = (theme.fg_color[0], theme.fg_color[1])
            ok_hover = (theme.hover_color[0], theme.hover_color[1])
        else:
            ok_fg = ctk.ThemeManager.theme["CTkButton"]["fg_color"]
            ok_hover = ctk.ThemeManager.theme["CTkButton"]["hover_color"]

        ok_button = ctk.CTkButton(self.content_frame, text="OK", command=self.destroy, width=100, fg_color=ok_fg, hover_color=ok_hover)
        ok_button.grid(row=1, column=0, padx=20, pady=(0, 20))

        self.after(250, ok_button.focus_force)
        self.center_on_master(width=400, height=180)

        self.grab_set()
        self.master.wait_window(self)

class AboutDialog(CustomToplevel):
    def __init__(self, master, title_text, desc_text, capabilities, footer_text, credits_text, icon_image=None, theme_colors=None):
        super().__init__(master, title="About", icon_image=icon_image)
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.rowconfigure(1, weight=1)

        # Header section with creator first, then app title and description
        header_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=(5, 5))  # Reduced top padding from 15 to 5
        header_frame.columnconfigure(0, weight=1)
        header_frame.rowconfigure(0, weight=0)  # Don't expand the header row

        # Creator credits at the very top with theme adaptation
        if theme_colors:
            credits_bg = (theme_colors.hover_color[0], theme_colors.hover_color[1])
            credits_text_color = ("white", "white")
        else:
            credits_bg = ("#e8f4fd", "#1a4a5c")
            credits_text_color = ("#1a4a5c", "#e8f4fd")
            
        credits_frame = ctk.CTkFrame(header_frame, fg_color=credits_bg, corner_radius=6)
        credits_frame.grid(row=0, column=0, sticky="ew", pady=(0, 15))
        credits_frame.columnconfigure(0, weight=1)

        credits_label = ctk.CTkLabel(
            credits_frame, 
            text=credits_text, 
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), 
            justify="center",
            text_color=credits_text_color
        )
        credits_label.grid(row=0, column=0, pady=12, padx=15, sticky="ew")

        title_label = ctk.CTkLabel(
            header_frame, 
            text=title_text, 
            font=ctk.CTkFont(family="Segoe UI", size=16, weight="bold"), 
            wraplength=390, 
            justify="center"
        )
        title_label.grid(row=1, column=0, sticky="ew")

        desc_label = ctk.CTkLabel(
            header_frame, 
            text=desc_text, 
            font=ctk.CTkFont(family="Segoe UI", size=11), 
            wraplength=360, 
            justify="center",
            text_color=("gray20", "gray80")
        )
        desc_label.grid(row=2, column=0, pady=(5, 0), sticky="ew")

        # Main content area
        main_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        main_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        main_frame.columnconfigure(0, weight=1)

        # Capabilities section - FIRST in main content
        cap_header = ctk.CTkLabel(
            main_frame, 
            text="Key Capabilities:", 
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), 
            anchor="w"
        )
        cap_header.grid(row=0, column=0, pady=(0, 8), sticky="ew")

        # Features with edge-to-edge layout like main window
        features_frame = ctk.CTkFrame(main_frame, fg_color=("gray95", "gray15"), corner_radius=0)
        features_frame.grid(row=1, column=0, pady=(0, 10), sticky="ew")
        features_frame.columnconfigure(0, weight=1)
        
        # Theme-aware check mark color
        if theme_colors:
            check_color = (theme_colors.fg_color[0], theme_colors.fg_color[1])
        else:
            check_color = ("#2a9d8f", "#2ECC71")
        
        for i, feature in enumerate(capabilities):
            line_frame = ctk.CTkFrame(features_frame, fg_color="transparent")
            line_frame.grid(row=i, column=0, sticky="ew", padx=15, pady=6)
            line_frame.columnconfigure(1, weight=1)
            
            check_label = ctk.CTkLabel(
                line_frame, 
                text="✓", 
                font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"), 
                text_color=check_color,
                width=15
            )
            check_label.grid(row=0, column=0, sticky="nw", padx=(0, 8))
            
            text_label = ctk.CTkLabel(
                line_frame, 
                text=feature, 
                font=ctk.CTkFont(family="Segoe UI", size=11), 
                wraplength=350, 
                justify="left",
                anchor="w"
            )
            text_label.grid(row=0, column=1, sticky="ew")

        # Footer section - LAST
        footer_label = ctk.CTkLabel(
            main_frame, 
            text=footer_text, 
            font=ctk.CTkFont(family="Segoe UI", size=10, slant="italic"), 
            justify="center",
            text_color=("gray40", "gray60")
        )
        footer_label.grid(row=2, column=0, pady=(2, 0), sticky="ew")
        
        # Button section
        button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        button_frame.grid(row=2, column=0, pady=(5, 10), sticky="ew")
        button_frame.columnconfigure(0, weight=1)
        
        # Theme-aware OK button
        if theme_colors:
            ok_fg = (theme_colors.fg_color[0], theme_colors.fg_color[1])
            ok_hover = (theme_colors.hover_color[0], theme_colors.hover_color[1])
        else:
            ok_fg = None
            ok_hover = None

        ok_button = ctk.CTkButton(
            button_frame,
            text="OK",
            command=self.destroy,
            width=100,
            height=32,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold"),
            fg_color=ok_fg if ok_fg else ctk.ThemeManager.theme["CTkButton"]["fg_color"],
            hover_color=ok_hover if ok_hover else ctk.ThemeManager.theme["CTkButton"]["hover_color"]
        )
        ok_button.grid(row=0, column=0)

        # Keyboard shortcuts and focus management
        # - Esc closes the dialog
        # - Enter triggers the OK button
        # - Focus the OK button shortly after the dialog is displayed
        self.bind("<Escape>", lambda e: self.destroy())
        self.bind("<Return>", lambda e: ok_button.invoke())
        self.after(100, ok_button.focus_force)

        self.center_on_master(width=420, height=450)
        self.grab_set()
        self.master.wait_window(self)

class SettingsWindow(CustomToplevel):
    def __init__(self, master, icon_image=None, theme_colors=None):
        super().__init__(master, title="Tools & Resources", icon_image=icon_image)
        self.master_app = master
        self.content_frame.grid_columnconfigure(0, weight=1)

        fg_color = (theme_colors.fg_color[0], theme_colors.fg_color[1])
        hover_color = (theme_colors.hover_color[0], theme_colors.hover_color[1])
        
        ctk.CTkButton(self.content_frame, text="Open Template", command=self.open_template, fg_color=fg_color, hover_color=hover_color).grid(row=0, column=0, padx=20, pady=(10, 5), sticky="ew")
        
        report_frame = ctk.CTkFrame(self.content_frame)
        report_frame.grid(row=1, column=0, padx=20, pady=5, sticky="ew")
        report_frame.grid_columnconfigure(0, weight=1)
        
        report_label = ctk.CTkLabel(report_frame, text="Output Report Format", font=ctk.CTkFont(weight="bold"))
        report_label.grid(row=0, column=0, columnspan=2, padx=10, pady=(5,0), sticky="w")
        
        self.report_format_var = tkinter.StringVar(value=config.REPORT_FORMAT)
        
        pdf_radio = ctk.CTkRadioButton(report_frame, text="PDF (for official documents)", variable=self.report_format_var, value="PDF", fg_color=fg_color, hover_color=hover_color)
        pdf_radio.grid(row=1, column=0, padx=20, pady=5, sticky="w")
        
        excel_radio = ctk.CTkRadioButton(report_frame, text="Excel (for data review)", variable=self.report_format_var, value="Excel", fg_color=fg_color, hover_color=hover_color)
        excel_radio.grid(row=2, column=0, padx=20, pady=5, sticky="w")
        
        # --- THEME FIX ---
        # "Clear Caches" now uses the theme's hover_color as its base color for a secondary look.
        # Its own hover color is now the primary fg_color, creating a "light up" effect.
        ctk.CTkButton(self.content_frame, text="Clear Caches", command=self.clear_all_caches, fg_color=hover_color, hover_color=fg_color).grid(row=2, column=0, padx=20, pady=5, sticky="ew")

        button_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        button_frame.grid(row=3, column=0, padx=20, pady=(10,15), sticky="e")
        
        ctk.CTkButton(button_frame, text="Save & Close", command=self.save_and_close, fg_color=fg_color, hover_color=hover_color).pack(side="right", padx=(5,0))
        
        # --- THEME FIX ---
        # The "Cancel" button's hover color is now themed.
        ctk.CTkButton(
            button_frame, 
            text="Cancel", 
            command=self.destroy,
            fg_color="transparent",
            border_width=1,
            text_color=("gray10", "gray90"),
            border_color=("gray70", "gray30"),
            hover_color=hover_color
        ).pack(side="right")

        self.center_on_master(width=320, height=270)
        self.grab_set()

    def save_and_close(self):
        selected_format = self.report_format_var.get()
        config.REPORT_FORMAT = selected_format
        self.master_app.log_message(f"✅ Default report format set to: {selected_format}")
        self.destroy()

    def open_template(self):
        template_url = GLOBAL_CONFIG["TEMPLATE_CSV_URL"]
        template_path = self.master_app.app_data.data_dir / "template.csv"
        template_meta_path = self.master_app.app_data.data_dir / "template.csv.meta"

        auth_headers = get_auth_headers(self.master_app.encryption_key)
        if not auth_headers:
             self.master_app.log_message("❌ CRITICAL: Could not prepare token for template download.")
             return

        if check_internet():
            self.master_app.log_message("🌐 Internet found.")
            status, msg = smart_download_pat(template_url, template_path, template_meta_path, auth_headers, self.master_app.encryption_key, encrypt_locally=False)
            if status == 'UP_TO_DATE': self.master_app.log_message("✅ [Template] Cache is up-to-date.")
            elif status == 'UPDATED': self.master_app.log_message("✅ [Template] Cache updated successfully.")
            else: self.master_app.log_message(f"⚠️ [Template] Could not check for updates: {msg}")
        else:
            self.master_app.log_message("⚠️ No internet connection detected.")
            if not template_path.exists():
                self.master_app.log_message("❌ [Template] No internet and no local cache found.")
                MessageDialog(self.master_app, title="Template Not Found", message="Could not download the template file. Please check your internet connection.", icon_image=self.master_app.logo_image)
                return
            else:
                self.master_app.log_message("✅ [Template] Using local cache.")

        if template_path.exists():
            try:
                self.master_app.log_message("✅ Opening template file...")
                if sys.platform == "win32": os.startfile(os.path.normpath(template_path))
                elif sys.platform == "darwin": os.system(f'open "{os.path.normpath(template_path)}"')
                else: os.system(f'xdg-open "{os.path.normpath(template_path)}"')
            except Exception as e:
                self.master_app.log_message(f"❌ Could not open template file: {e}")

    def clear_all_caches(self):
        self.master_app.log_message("🗑️ Clearing all caches...")
        deleted_any = False
        
        cache_dir = self.master_app.app_data.data_dir
        
        files_to_clear = glob.glob(os.path.join(cache_dir, '*.csv')) + glob.glob(os.path.join(cache_dir, '*.meta'))

        for f_path in files_to_clear:
            try:
                os.remove(f_path)
                deleted_any = True
            except OSError as e:
                self.master_app.log_message(f"❌ Could not delete cache file {os.path.basename(f_path)}: {e}")

        if deleted_any:
            self.master_app.log_message("✅ All caches have been cleared.")
            self.master_app.update_status("nickname"); self.master_app.update_status("db"); self.master_app.update_status("officials")
        else:
            self.master_app.log_message("ℹ️ No caches were found to clear.")