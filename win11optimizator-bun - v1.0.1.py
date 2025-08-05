# Win11Optimizator_v5.py
# Requires: Python 3.x, Administrative privileges
# Author: Adapted from original, incorporating commands from 'comenzi de pus in aplicatia Win11Optimizer.txt'
# Date: 2025-08-05
# Version: 1.0.2

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, Menu, BooleanVar
import configparser
import os
import sys
import ctypes
import subprocess
import winreg
import threading
import logging
import traceback
import time
from datetime import datetime
import urllib.request
import queue

# --- Configuration ---
LOG_FILENAME = "win11optimizator.log"
CONFIG_FILENAME = "win11optimizator.ini"
APP_NAME = "Win11Optimizator"
APP_VERSION = "1.0.1"

# --- Logging Setup ---
logging.basicConfig(
    filename=LOG_FILENAME,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logging.info(f"--- Starting {APP_NAME} v{APP_VERSION} ---")

# --- Utility Functions ---
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def request_admin_privileges():
    if not is_admin():
        try:
            ctypes.windll.shell32.ShellExecuteW(
                None, "runas", sys.executable, " ".join(sys.argv), None, 1
            )
            sys.exit(0)
        except Exception as e:
            logging.error(f"Failed to elevate privileges: {e}")
            messagebox.showerror("Error", f"Cannot obtain admin rights: {e}")
            sys.exit(1)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# --- Custom Logging Handler for Tkinter Text Widget ---
class TextHandler(logging.Handler):
    def __init__(self, text_widget):
        super().__init__()
        self.text_widget = text_widget
        self.setLevel(logging.INFO)

    def emit(self, record):
        msg = self.format(record)
        def append():
            self.text_widget.configure(state='normal')
            self.text_widget.insert(tk.END, msg + '\n')
            self.text_widget.configure(state='disabled')
            self.text_widget.see(tk.END)
        self.text_widget.after(0, append)

# --- Main Application Class ---
class Win11Optimizator:
    def __init__(self, root):
        self.root = root
        self.root.title(f"{APP_NAME} v{APP_VERSION}")
        # Setează fereastra să pornească maximizată, păstrând barele de titlu și butoanele
        self.root.state('zoomed')

        try:
            # self.root.iconbitmap(resource_path("icon.ico"))
            pass
        except tk.TclError:
            logging.warning("Icon file 'icon.ico' not found. Using default icon.")

        # --- Style Configuration ---
        self.style = ttk.Style()
        self.current_theme = "light" # Default theme
        self.configure_styles()

        # --- Menu Bar ---
        #self.create_menu()

        # --- Main Container ---
        self.main_frame = ttk.Frame(root) # Main frame for overall theme application
        self.main_frame.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        # --- Section Navigation Buttons ---
        self.nav_frame = ttk.Frame(self.main_frame) # Store nav_frame reference
        self.nav_frame.pack(fill="x", pady=(0, 10))

        self.nav_buttons = []
        sections = ["Essential Tweaks", "Advanced Tweaks", "Customize Preferences", "Install Software"]
        for i, text in enumerate(sections):
            # Create button with green style
            btn = ttk.Button(self.nav_frame, text=text, command=lambda idx=i: self.show_section(idx), style='Green.TButton')
            btn.pack(side="left", padx=5, pady=2) # Add small pady for spacing
            self.nav_buttons.append(btn)

        # --- Additional Buttons (Toggle Theme, Export, Import, About) ---
        self.theme_toggle_btn = ttk.Button(self.nav_frame, text="Light/Dark Theme", command=self.toggle_theme, style='Green.TButton')
        self.theme_toggle_btn.pack(side="left", padx=5, pady=2)

        self.export_btn = ttk.Button(self.nav_frame, text="Export", command=self.export_config, style='Green.TButton')
        self.export_btn.pack(side="left", padx=5, pady=2)

        self.import_btn = ttk.Button(self.nav_frame, text="Import", command=self.import_config, style='Green.TButton')
        self.import_btn.pack(side="left", padx=5, pady=2)

        self.about_btn = ttk.Button(self.nav_frame, text="About", command=self.show_about, style='Green.TButton')
        self.about_btn.pack(side="left", padx=5, pady=2)

        # --- Section Container ---
        self.section_container = ttk.Frame(self.main_frame)
        self.section_container.pack(fill="both", expand=True)

        # --- Initialize Sections ---
        self.sections = []
        self.section1_vars = {}
        self.section2_vars = {}
        self.section3_vars = {}
        self.section4_vars = {}

        self.create_section1()
        self.create_section2()
        self.create_section3()
        self.create_section4()

        # --- Show First Section ---
        self.show_section(0)

        # Apply initial theme to ensure everything is styled correctly from the start
        self.apply_theme(self.current_theme)

        logging.info("Application initialized successfully.")

    def configure_styles(self):
        """Configure ttk styles for light and dark themes."""
        # --- Light Theme ---
        LIGHT_BG = "#f0f0f0"
        LIGHT_FG = "black"
        LIGHT_SCROLLBAR_TROUGH = "#e0e0e0"
        LIGHT_SCROLLBAR_BG = "#c0c0c0"
        self.style.configure("Light.TFrame", background=LIGHT_BG)
        self.style.configure("Light.TLabel", background=LIGHT_BG, foreground=LIGHT_FG)
        self.style.configure("Light.TCheckbutton", background=LIGHT_BG, foreground=LIGHT_FG)
        self.style.configure("Light.Vertical.TScrollbar", troughcolor=LIGHT_SCROLLBAR_TROUGH, background=LIGHT_SCROLLBAR_BG)

        # --- Dark Theme ---
        DARK_BG = "#000000" # Pure Black
        DARK_FG = "#FFFFFF" # Pure White
        DARK_SCROLLBAR_TROUGH = "#2e2e2e"
        DARK_SCROLLBAR_BG = "#444444"
        self.style.configure("Dark.TFrame", background=DARK_BG)
        self.style.configure("Dark.TLabel", background=DARK_BG, foreground=DARK_FG)
        self.style.configure("Dark.TCheckbutton", background=DARK_BG, foreground=DARK_FG, indicatorcolor=DARK_BG)
        self.style.configure("Dark.Vertical.TScrollbar", troughcolor=DARK_SCROLLBAR_TROUGH, background=DARK_SCROLLBAR_BG)

        # --- Green Button Style (Solid green background, black text) ---
        GREEN_BG = "#4CAF50" # Material Green 500
        GREEN_BG_HOVER = "#45a049"
        GREEN_BG_PRESSED = "#3d8b40"
        BUTTON_FG = "black" # Black text for buttons

        # Use map to define different states, ensuring solid background
        self.style.configure("Green.TButton",
                             foreground=BUTTON_FG,
                             background=GREEN_BG, # Solid background
                             borderwidth=1) # Minimal border
        self.style.map("Green.TButton",
                       background=[('active', GREEN_BG_HOVER), ('pressed', GREEN_BG_PRESSED)],
                       foreground=[('pressed', BUTTON_FG), ('active', BUTTON_FG)])


    def apply_theme(self, theme_name):
        """Apply the selected theme to the application."""
        self.current_theme = theme_name
        suffix = "Dark" if theme_name == "dark" else "Light"
        DARK_BG = "#000000"
        DARK_FG = "#FFFFFF"
        LIGHT_BG = "#f0f0f0"
        LIGHT_FG = "black"

        # --- Apply base theme style ---
        bg_color = DARK_BG if theme_name == "dark" else LIGHT_BG
        fg_color = DARK_FG if theme_name == "dark" else LIGHT_FG
        self.root.configure(bg=bg_color) # Root window background
        self.main_frame.configure(style=f"{suffix}.TFrame") # Main container frame

        # --- Apply theme to Navigation Bar ---
        self.nav_frame.configure(style=f"{suffix}.TFrame")
        # Update nav buttons style (they use Green.TButton, which is static)
        # Update additional buttons style (they also use Green.TButton)

        # --- Update styles for existing widgets in sections ---
        for i, section in enumerate(self.sections):
            # --- Section Frame itself ---
            section.configure(style=f"{suffix}.TFrame")

            # --- Left Frame (Checkboxes) ---
            left_frame = section.winfo_children()[0] if section.winfo_children() else None
            if left_frame and isinstance(left_frame, ttk.Frame):
                left_frame.configure(style=f"{suffix}.TFrame")
                # Update left frame's canvas background
                canvas = left_frame.winfo_children()[0] if left_frame.winfo_children() else None
                if canvas and isinstance(canvas, tk.Canvas):
                    canvas.configure(bg=bg_color) # Match section bg
                    # Update scrollable frame
                    scrollable_frame = canvas.winfo_children()[0] if canvas.winfo_children() else None
                    if scrollable_frame and isinstance(scrollable_frame, ttk.Frame):
                        scrollable_frame.configure(style=f"{suffix}.TFrame")
                        # Update individual checkboxes and labels within the scrollable frame
                        for child in scrollable_frame.winfo_children():
                            if isinstance(child, ttk.Checkbutton):
                                child.configure(style=f"{suffix}.TCheckbutton")
                            elif isinstance(child, ttk.Label): # For section 4 categories
                                child.configure(style=f"{suffix}.TLabel")

            # --- Right Frame (Output) ---
            # Assuming right frame is the second child
            right_frame = section.winfo_children()[1] if len(section.winfo_children()) > 1 else None
            if right_frame and isinstance(right_frame, ttk.Frame):
                right_frame.configure(style=f"{suffix}.TFrame")
                # Update output label
                output_label = right_frame.winfo_children()[0] if right_frame.winfo_children() else None
                if output_label and isinstance(output_label, ttk.Label):
                     output_label.configure(style=f"{suffix}.TLabel")
                # Update output text widget
                output_text_widget = None
                for widget in right_frame.winfo_children():
                    if isinstance(widget, tk.Text):
                        output_text_widget = widget
                        break
                if output_text_widget:
                    output_text_widget.configure(
                        bg='#111111' if theme_name == 'dark' else 'white', # Slightly off-black for dark theme
                        fg=DARK_FG if theme_name == 'dark' else LIGHT_FG
                    )
                # Update output scrollbar trough and slider color
                output_scrollbar = None
                for widget in right_frame.winfo_children():
                    if isinstance(widget, ttk.Scrollbar) and widget.cget('orient') == 'vertical':
                        output_scrollbar = widget
                        break
                if output_scrollbar:
                    # Reconfigure the scrollbar style based on theme
                    if theme_name == "dark":
                        self.style.configure("Vertical.TScrollbar", troughcolor="#2e2e2e", background="#444444")
                    else:
                        self.style.configure("Vertical.TScrollbar", troughcolor="#e0e0e0", background="#c0c0c0")
                    output_scrollbar.configure(style="Vertical.TScrollbar")


            # --- Bottom Frame (Progress/Execute Button) ---
            # Find the bottom frame within the left frame
            if left_frame:
                 for child in left_frame.winfo_children():
                      if isinstance(child, ttk.Frame) and child != canvas: # Assuming canvas is the first child
                           # This should be the bottom frame
                           child.configure(style=f"{suffix}.TFrame")
                           # Update progress bar if it exists within (usually doesn't need specific styling)
                           # Update execute/install buttons (they use Green.TButton)


        # --- Update scrollbar styles for main sections (left side lists) ---
        # This is tricky because scrollbars are inside dynamic frames.
        # The best way is to reconfigure the style globally and ensure it's applied.
        # We already did this above for the output scrollbars.
        # For the section scrollbars, they should pick up the style when created/updated.
        # Let's try reconfiguring the global Vertical.TScrollbar style here again.
        if theme_name == "dark":
             self.style.configure("Vertical.TScrollbar", troughcolor="#2e2e2e", background="#444444")
             # Also explicitly set the background of the nav frame's potential background
             self.nav_frame.configure(style=f"{suffix}.TFrame") # Re-assert
        else: # light
             self.style.configure("Vertical.TScrollbar", troughcolor="#e0e0e0", background="#c0c0c0")
             self.nav_frame.configure(style=f"{suffix}.TFrame") # Re-assert

        # --- Update About Window if it exists ---
        for w in self.root.winfo_children():
            if isinstance(w, tk.Toplevel) and w.title() == "About Win11Optimizator":
                 try:
                     # Update main frame of about window
                     about_main_frame = w.winfo_children()[0] if w.winfo_children() else None
                     if about_main_frame and isinstance(about_main_frame, ttk.Frame):
                         about_main_frame.configure(style=f"{suffix}.TFrame")
                         # Update labels inside about window
                         for child in about_main_frame.winfo_children():
                             if isinstance(child, ttk.Label):
                                 child.configure(style=f"{suffix}.TLabel")
                             elif isinstance(child, tk.Text):
                                 child.configure(bg='#111111' if theme_name == 'dark' else '#f0f0f0',
                                                 fg=DARK_FG if theme_name == 'dark' else LIGHT_FG)
                 except tk.TclError:
                     pass # Ignore errors if widget structure is unexpected

        logging.info(f"Theme applied: {theme_name}")

    def toggle_theme(self):
        """Toggle between light and dark themes."""
        new_theme = "dark" if self.current_theme == "light" else "light"
        self.apply_theme(new_theme)
        logging.info(f"Theme toggled to: {new_theme}")

    #def create_menu(self):
    #    menubar = Menu(self.root)
    #    self.root.config(menu=menubar)

        # --- Help Menu ---
    #    help_menu = Menu(menubar, tearoff=0)
    #    menubar.add_cascade(label="Help", menu=help_menu)
    #    help_menu.add_command(label="About", command=self.show_about)
    #    help_menu.add_command(label="Import Configuration", command=self.import_config)
    #    help_menu.add_command(label="Export Configuration", command=self.export_config)

    def show_section(self, index):
        for i, section in enumerate(self.sections):
            if i == index:
                section.pack(fill="both", expand=True)
            else:
                section.pack_forget()

    def show_about(self):
        try:
            about_window = tk.Toplevel(self.root)
            about_window.title("About Win11Optimizator")
            about_window.geometry("500x400")
            about_window.resizable(False, False)
            # Apply theme to about window if desired (requires more setup)
            main_frame = ttk.Frame(about_window, style=f"{self.current_theme.capitalize()}.TFrame")
            main_frame.pack(fill="both", expand=True, padx=20, pady=20)
            title_label = ttk.Label(main_frame, text="Win11Optimizator", font=("Segoe UI", 18, "bold"), style=f"{self.current_theme.capitalize()}.TLabel")
            title_label.pack(pady=10)
            version_label = ttk.Label(main_frame, text=f"Version: {APP_VERSION}", style=f"{self.current_theme.capitalize()}.TLabel")
            version_label.pack()
            author_label = ttk.Label(main_frame, text="Author: EOLiann", style=f"{self.current_theme.capitalize()}.TLabel")
            author_label.pack(pady=(10, 0))
            desc_label = ttk.Label(main_frame, text="Optimize and customize Windows 11", wraplength=450, style=f"{self.current_theme.capitalize()}.TLabel")
            desc_label.pack(pady=10)
            disclaimer = tk.Text(main_frame, height=8, wrap="word", bg="#111111" if self.current_theme == "dark" else "#f0f0f0", fg="white" if self.current_theme == "dark" else "black")
            disclaimer.insert("1.0", "DISCLAIMER:\n"
                                     "This tool modifies system settings. Use at your own risk.\n"
                                     "It's highly recommended to create a system restore point before making changes.\n"
                                     "The author is not responsible for any damage caused by this tool.")
            disclaimer.config(state="disabled")
            disclaimer.pack(fill="both", expand=True, pady=10)
            close_btn = ttk.Button(main_frame, text="Close", command=about_window.destroy, style='Green.TButton')
            close_btn.pack(pady=5)
        except Exception as e:
            logging.error(f"Show about failed: {e}\n{traceback.format_exc()}")
            messagebox.showerror("Error", f"Failed to show about dialog: {e}")

    # --- Import/Export Configuration Methods ---
    def export_config(self):
        try:
            config = configparser.ConfigParser()
            config["Section1"] = {name: str(var.get()) for name, var in self.section1_vars.items()}
            config["Section2"] = {name: str(var.get()) for name, var in self.section2_vars.items()}
            config["Section3"] = {name: str(var.get()) for name, var in self.section3_vars.items()}
            config["Section4"] = {name: str(data["var"].get()) for name, data in self.section4_vars.items()}

            file_path = filedialog.asksaveasfilename(defaultextension=".ini", filetypes=[("INI files", "*.ini"), ("All files", "*.*")])
            if file_path:
                with open(file_path, 'w') as configfile:
                    config.write(configfile)
                logging.info("Export Config: Success")
                messagebox.showinfo("Export Successful", "Configuration exported successfully.")
            else:
                logging.info("Export Config: Cancelled - No file selected")
        except Exception as e:
            logging.error(f"Export Config failed: {str(e)}\n{traceback.format_exc()}")
            messagebox.showerror("Export Failed", f"Failed to export configuration: {str(e)}")

    def import_config(self):
        try:
            file_path = filedialog.askopenfilename(filetypes=[("INI files", "*.ini"), ("All files", "*.*")])
            if not file_path:
                logging.info("Import Config: Cancelled - No file selected")
                return

            config = configparser.ConfigParser()
            config.read(file_path)

            if "Section1" in config:
                for name, var in self.section1_vars.items():
                    if name in config["Section1"]:
                        var.set(config.getboolean("Section1", name, fallback=False))
            if "Section2" in config:
                for name, var in self.section2_vars.items():
                    if name in config["Section2"]:
                        var.set(config.getboolean("Section2", name, fallback=False))
            if "Section3" in config:
                for name, var in self.section3_vars.items():
                    if name in config["Section3"]:
                        var.set(config.getboolean("Section3", name, fallback=False))
            if "Section4" in config:
                for name, data in self.section4_vars.items():
                    if name in config["Section4"]:
                        data["var"].set(config.getboolean("Section4", name, fallback=False))

            logging.info("Import Config: Success")
            messagebox.showinfo("Import Successful", "Configuration imported successfully.")
        except Exception as e:
            logging.error(f"Import Config failed: {str(e)}\n{traceback.format_exc()}")
            messagebox.showerror("Import Failed", f"Failed to import configuration: {str(e)}")

    # --- Utility for Logging Commands ---
    def log_and_run_command(self, logger, command, description=""):
        """Logs a command and then executes it."""
        if description:
            logger.info(f"Executing: {description}")
        logger.info(f"Command: {command}")
        try:
            # Use shell=True for registry commands and general execution
            result = subprocess.run(command, shell=True, capture_output=True, text=True, check=True)
            logger.info(f"Success: {description or command}")
            # Optionally log stdout if needed for debugging specific commands
            # if result.stdout.strip():
            #     logger.debug(f"StdOut: {result.stdout.strip()}")
        except subprocess.CalledProcessError as e:
            logger.error(f"Failed: {description or command}")
            logger.error(f"Error: {e.stderr.strip() if e.stderr else str(e)}")
            raise # Re-raise to be caught by the calling function
        except Exception as e:
            logger.error(f"Unexpected error running command '{command}': {e}")
            raise

    # --- Section 1: Essential Tweaks ---
    def create_section1(self):
        try:
            start_time = time.time()
            frame = ttk.Frame(self.section_container, style=f"{self.current_theme.capitalize()}.TFrame")
            self.sections.append(frame)

            tweaks = [
                "Create Restore Point",
                "Delete Temporary Files",
                "Disable ConsumerFeatures",
                "Disable Telemetry",
                "Disable Activity History",
                "Disable Explorer Automatic Folder Discovery",
                "Disable GameDVR",
                "Disable Hibernation",
                "Disable Homegroup",
                "Disable Location Tracking",
                "Disable Storage Sense",
                "Disable Wifi-Sense",
                "Enable End Task With Right Click",
                "Run Disk Cleanup",
                "Change Windows Terminal default: PowerShell 5 -> PowerShell 7",
                "Disable Powershell 7 Telemetry",
                "Disable Recall",
                "Set Hibernation as default (good for laptops)",
                "Set Services to Manual",
                "Debloat Brave",
                "Debloat Edge",
                # --- New from comenzi.txt ---
                "Disable Windows Setting Sync",
                "Disable Handwriting Data Collection",
                "Disable Clipboard Data Collection",
                "Opt out of Privacy Consent",
                "Disable 3rd-party apps Telemetry (Adobe/NVIDIA/VS)",
                "Disable Lockscreen Camera Access",
                "Disable Biometrics",
                "Disable Cortana & Bing Search",
                "Disable Feedback Notifications",
                "Disable Verbose Logon",
                "Disable Sticky Keys",
                "Disable Numlock on startup",
                "Disable Snap Flyout",
                "Set Classic Right Click Menu",
                "Show File Extensions",
                "Disable Mouse Acceleration",
                "Disable Fullscreen Optimizations",
                "Disable Copilot",
                "Disable IntelliCode Remote Analysis",
                "Disable Media Player Telemetry",
                "Disable System File Access Consent",
                "Disable Account Info Access Consent",
                # --- More from comenzi.txt ---
                "Disable Language Sync",
                "Disable Credentials Sync",
                "Disable Desktop Theme Sync",
                "Disable Personalization Sync",
                "Disable Start Layout Sync",
                "Disable Application Sync",
                "Disable App Sync Setting Sync",
                "Disable Sync on Paid Network",
            ]
            self.section1_vars = {tweak: BooleanVar(value=False) for tweak in tweaks}

            # --- Layout Split ---
            left_frame = ttk.Frame(frame, style=f"{self.current_theme.capitalize()}.TFrame")
            left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

            right_frame = ttk.Frame(frame, width=400, style=f"{self.current_theme.capitalize()}.TFrame")
            right_frame.pack_propagate(False)
            right_frame.pack(side="right", fill="y", padx=(5, 0))

            # --- Left Column: Checkboxes ---
            canvas = tk.Canvas(left_frame, bg=self.style.lookup(f"{self.current_theme.capitalize()}.TFrame", 'background'), highlightthickness=0)
            scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=canvas.yview, style="Vertical.TScrollbar")
            scrollable_frame = ttk.Frame(canvas, style=f"{self.current_theme.capitalize()}.TFrame")

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # --- Bind Mousewheel to Scrollable Frame for Section 1 ---
            # Use lambda to capture the correct canvas instance
            scrollable_frame.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
            # Also bind to canvas to ensure it works if mouse is over the canvas but not directly over a widget
            canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
            # --- End Mousewheel Binding ---

            for i, tweak in enumerate(tweaks):
                chk = ttk.Checkbutton(
                    scrollable_frame,
                    text=tweak,
                    variable=self.section1_vars[tweak],
                    command=lambda t=tweak: logging.info(f"Section1 {t}: Toggled to {self.section1_vars[t].get()}"),
                    style=f"{self.current_theme.capitalize()}.TCheckbutton"
                )
                chk.grid(row=i, column=0, sticky="w", padx=10, pady=5)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            # --- Right Column: Execution Output ---
            output_label = ttk.Label(right_frame, text="Execution Output:", font=("Segoe UI", 10, "bold"), style=f"{self.current_theme.capitalize()}.TLabel")
            output_label.pack(anchor="w", padx=5, pady=(5, 0))

            self.output_text1 = tk.Text(right_frame, state='disabled', wrap='word',
                                        bg='black' if self.current_theme == 'dark' else 'white',
                                        fg='white' if self.current_theme == 'dark' else 'black')
            output_scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=self.output_text1.yview, style="Vertical.TScrollbar")
            self.output_text1.configure(yscrollcommand=output_scrollbar.set)

            self.output_text1.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
            output_scrollbar.pack(side="right", fill="y", pady=5, padx=(0, 5))

            # --- Bottom Frame (MOVED to LEFT FRAME) ---
            bottom_frame = ttk.Frame(left_frame, style=f"{self.current_theme.capitalize()}.TFrame") # Changed parent to left_frame
            bottom_frame.pack(side="bottom", fill="x", padx=10, pady=10) # Adjusted padding if needed

            self.progress1 = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate", length=300)
            self.progress1.pack(side="left", fill="x", expand=True, padx=(0, 10))

            execute_btn = ttk.Button(
                bottom_frame,
                text="Execute",
                command=lambda: threading.Thread(target=self.execute_section1, daemon=True).start(),
                style='Green.TButton'
            )
            execute_btn.pack(side="right")

            logging.info(f"Section 1 created successfully in {time.time() - start_time:.3f} seconds")
        except Exception as e:
            logging.error(f"Create section 1 failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def execute_section1(self):
        try:
            if not hasattr(self, 'output_text1'):
                logging.error("Output text widget for section 1 not found.")
                return

            exec_logger = logging.getLogger(f"Section1_Exec_{threading.get_ident()}")
            exec_logger.setLevel(logging.INFO)
            exec_logger.handlers.clear()
            text_handler = TextHandler(self.output_text1)
            formatter = logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt='%H:%M:%S')
            text_handler.setFormatter(formatter)
            exec_logger.addHandler(text_handler)

            selected = [name for name, var in self.section1_vars.items() if var.get()]
            if not selected:
                exec_logger.info("No options selected for execution.")
                logging.info("Section1: No options selected for execution.")
                return

            self.progress1["value"] = 0
            self.progress1["maximum"] = len(selected)

            self.output_text1.configure(state='normal')
            self.output_text1.delete(1.0, tk.END)
            self.output_text1.configure(state='disabled')

            for i, option in enumerate(selected):
                try:
                    exec_logger.info(f"Starting: {option}")
                    # --- Execution Logic with Command Logging ---
                    if option == "Create Restore Point":
                        self.create_restore_point(exec_logger)
                    elif option == "Delete Temporary Files":
                        self.delete_temp_files(exec_logger)
                    elif option == "Disable ConsumerFeatures":
                        self.disable_consumer_features(exec_logger)
                    elif option == "Disable Telemetry":
                        self.disable_telemetry(exec_logger)
                    elif option == "Disable Activity History":
                        self.disable_activity_history(exec_logger)
                    elif option == "Disable Explorer Automatic Folder Discovery":
                        self.disable_folder_discovery(exec_logger)
                    elif option == "Disable GameDVR":
                        self.disable_gamedvr(exec_logger)
                    elif option == "Disable Hibernation":
                        self.disable_hibernation(exec_logger)
                    elif option == "Disable Homegroup":
                        self.disable_homegroup(exec_logger)
                    elif option == "Disable Location Tracking":
                        self.disable_location_tracking(exec_logger)
                    elif option == "Disable Storage Sense":
                        self.disable_storage_sense(exec_logger)
                    elif option == "Disable Wifi-Sense":
                        self.disable_wifi_sense(exec_logger)
                    elif option == "Enable End Task With Right Click":
                        self.enable_end_task(exec_logger)
                    elif option == "Run Disk Cleanup":
                        self.run_disk_cleanup(exec_logger)
                    elif option == "Change Windows Terminal default: PowerShell 5 -> PowerShell 7":
                        self.set_powershell7_default(exec_logger)
                    elif option == "Disable Powershell 7 Telemetry":
                        self.disable_powershell7_telemetry(exec_logger)
                    elif option == "Disable Recall":
                        self.disable_recall(exec_logger)
                    elif option == "Set Hibernation as default (good for laptops)":
                        self.set_hibernation_default(exec_logger)
                    elif option == "Set Services to Manual":
                        self.set_services_manual(exec_logger)
                    elif option == "Debloat Brave":
                        self.debloat_brave(exec_logger)
                    elif option == "Debloat Edge":
                        self.debloat_edge(exec_logger)
                    # --- New Executions ---
                    elif option == "Disable Windows Setting Sync":
                        self.disable_windows_setting_sync(exec_logger)
                    elif option == "Disable Handwriting Data Collection":
                        self.disable_handwriting_data_collection(exec_logger)
                    elif option == "Disable Clipboard Data Collection":
                        self.disable_clipboard_data_collection(exec_logger)
                    elif option == "Opt out of Privacy Consent":
                        self.opt_out_privacy_consent(exec_logger)
                    elif option == "Disable 3rd-party apps Telemetry (Adobe/NVIDIA/VS)":
                        self.disable_3rd_party_telemetry(exec_logger)
                    elif option == "Disable Lockscreen Camera Access":
                        self.disable_lockscreen_camera(exec_logger)
                    elif option == "Disable Biometrics":
                        self.disable_biometrics(exec_logger)
                    elif option == "Disable Cortana & Bing Search":
                        self.disable_cortana_bing_search(exec_logger)
                    elif option == "Disable Feedback Notifications":
                        self.disable_feedback_notifications(exec_logger)
                    elif option == "Disable Verbose Logon":
                        self.disable_verbose_logon(exec_logger)
                    elif option == "Disable Sticky Keys":
                        self.disable_sticky_keys(exec_logger)
                    elif option == "Disable Numlock on startup":
                        self.disable_numlock_startup(exec_logger)
                    elif option == "Disable Snap Flyout":
                        self.disable_snap_flyout(exec_logger)
                    elif option == "Set Classic Right Click Menu":
                        self.set_classic_right_click(exec_logger)
                    elif option == "Show File Extensions":
                        self.show_file_extensions(exec_logger)
                    elif option == "Disable Mouse Acceleration":
                        self.disable_mouse_acceleration(exec_logger)
                    elif option == "Disable Fullscreen Optimizations":
                        self.disable_fullscreen_optimizations(exec_logger) # Already implemented
                    elif option == "Disable Copilot":
                        self.disable_copilot(exec_logger) # Already implemented
                    elif option == "Disable IntelliCode Remote Analysis":
                        self.disable_intellicode_remote_analysis(exec_logger)
                    elif option == "Disable Media Player Telemetry":
                        self.disable_media_player_telemetry(exec_logger)
                    elif option == "Disable System File Access Consent":
                        self.disable_system_file_access_consent(exec_logger)
                    elif option == "Disable Account Info Access Consent":
                        self.disable_account_info_access_consent(exec_logger)
                    elif option == "Disable Language Sync":
                        self.disable_language_sync(exec_logger)
                    elif option == "Disable Credentials Sync":
                        self.disable_credentials_sync(exec_logger)
                    elif option == "Disable Desktop Theme Sync":
                        self.disable_desktop_theme_sync(exec_logger)
                    elif option == "Disable Personalization Sync":
                        self.disable_personalization_sync(exec_logger)
                    elif option == "Disable Start Layout Sync":
                        self.disable_start_layout_sync(exec_logger)
                    elif option == "Disable Application Sync":
                        self.disable_application_sync(exec_logger)
                    elif option == "Disable App Sync Setting Sync":
                        self.disable_app_sync_setting_sync(exec_logger)
                    elif option == "Disable Sync on Paid Network":
                        self.disable_sync_on_paid_network(exec_logger)
                    # --- End New Executions ---
                    exec_logger.info(f"Completed: {option}")
                    logging.info(f"Section1 {option}: Enabled")
                    self.progress1["value"] = i + 1
                    self.progress1.update()
                except Exception as e:
                    error_msg = f"Failed: {option} - {str(e)}"
                    exec_logger.error(error_msg)
                    logging.error(f"Section1 {option}: Failed - {str(e)}\n{traceback.format_exc()}")

            for name, var in self.section1_vars.items():
                var.set(False)

            exec_logger.info(f"Execution Complete: Applied {len(selected)} tweaks!")
            logging.info(f"Section1: Finished applying {len(selected)} tweaks")

        except Exception as e:
            logging.error(f"Execute section 1 failed: {str(e)}\n{traceback.format_exc()}")
            if hasattr(self, 'output_text1'):
                try:
                    fatal_logger = logging.getLogger(f"Section1_Exec_{threading.get_ident()}")
                    if not fatal_logger.handlers:
                        fatal_logger.addHandler(TextHandler(self.output_text1))
                        fatal_logger.setLevel(logging.ERROR)
                    fatal_logger.error(f"Execution Failed: {str(e)}")
                except:
                    pass
            raise

    # --- Section 1 Methods (Updated to accept logger) ---
    def create_restore_point(self, logger):
        try:
            cmd = 'powershell -command "Enable-ComputerRestore -Drive $env:SystemDrive"'
            self.log_and_run_command(logger, cmd, "Enable Computer Restore")
            cmd = 'powershell -command "Checkpoint-Computer -Description \'Win11Optimizator Restore Point\' -RestorePointType \'MODIFY_SETTINGS\'"'
            self.log_and_run_command(logger, cmd, "Create Restore Point")
            logger.info("Create Restore Point: Success")
        except subprocess.CalledProcessError as e:
            logger.warning(f"Create Restore Point: Might have partially failed or requires manual check. Error: {e}")
        except Exception as e:
            logger.error(f"Create Restore Point failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def delete_temp_files(self, logger):
        try:
            cmd = 'del /q /f /s %TEMP%\\*'
            self.log_and_run_command(logger, cmd, "Delete Temporary Files")
            cmd = 'PowerShell -ExecutionPolicy Unrestricted -Command "$bin = (New-Object -ComObject Shell.Application).NameSpace(10); $bin.items()| ForEach {; Write-Host \"Deleting $($_.Name) from Recycle Bin\"; Remove-Item $_.Path -Recurse -Force; }"'
            self.log_and_run_command(logger, cmd, "Empty Recycle Bin")
            logger.info("Delete Temporary Files: Success")
        except subprocess.CalledProcessError as e:
            logger.warning(f"Delete Temporary Files: Some files might not have been deleted. Error: {e}")
        except Exception as e:
            logger.error(f"Delete Temporary Files failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_consumer_features(self, logger):
         try:
             key_path = r"SOFTWARE\Policies\Microsoft\Windows\CloudContent"
             with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                 winreg.SetValueEx(key, "DisableWindowsConsumerFeatures", 0, winreg.REG_DWORD, 1)
             logger.info("Disable ConsumerFeatures: Applied registry change.")
             logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"DisableWindowsConsumerFeatures\" /t REG_DWORD /d 1 /f")
         except Exception as e:
             logger.error(f"Disable ConsumerFeatures failed: {str(e)}\n{traceback.format_exc()}")
             raise

    def disable_telemetry(self, logger):
        try:
            key_path = r"SOFTWARE\Policies\Microsoft\Windows\DataCollection"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "AllowTelemetry", 0, winreg.REG_DWORD, 0)
            logger.info("Disable Telemetry: Applied registry change.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"AllowTelemetry\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Disable Telemetry failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_activity_history(self, logger):
        try:
             commands = [
                 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v "EnableActivityFeed" /t REG_DWORD /d 0 /f',
                 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v "PublishUserActivities" /t REG_DWORD /d 0 /f',
                 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v "UploadUserActivities" /t REG_DWORD /d 0 /f'
             ]
             for cmd in commands:
                 self.log_and_run_command(logger, cmd)
             logger.info("Disable Activity History: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Activity History failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Activity History failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_folder_discovery(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "Start_SearchFiles", 0, winreg.REG_DWORD, 0)
            logger.info("Disable Explorer Folder Discovery: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"Start_SearchFiles\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Disable Explorer Folder Discovery failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_gamedvr(self, logger):
        try:
            key_path = r"SOFTWARE\Policies\Microsoft\Windows\GameDVR"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "AllowGameDVR", 0, winreg.REG_DWORD, 0)
            logger.info("Disable GameDVR: Applied registry change.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"AllowGameDVR\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Disable GameDVR failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_hibernation(self, logger):
        try:
            cmd = "powercfg /h off"
            self.log_and_run_command(logger, cmd, "Disable Hibernation")
            logger.info("Disable Hibernation: Success")
        except subprocess.CalledProcessError as e:
            logger.warning(f"Disable Hibernation: Command might have failed. Error: {e}")
        except Exception as e:
            logger.error(f"Disable Hibernation failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_homegroup(self, logger):
        try:
            services = ["HomeGroupListener", "HomeGroupProvider"]
            for service in services:
                cmd = f'sc config "{service}" start= disabled'
                self.log_and_run_command(logger, cmd, f"Disable service {service}")
            logger.info("Disable Homegroup: Success")
        except subprocess.CalledProcessError as e:
            logger.warning(f"Disable Homegroup: Service config might have failed. Error: {e}")
        except Exception as e:
            logger.error(f"Disable Homegroup failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_location_tracking(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "Value", 0, winreg.REG_SZ, "Deny")
            logger.info("Disable Location Tracking: Applied registry change.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"Value\" /t REG_SZ /d \"Deny\" /f")
        except Exception as e:
            logger.error(f"Disable Location Tracking failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_storage_sense(self, logger):
        try:
            key_path = r"SOFTWARE\Policies\Microsoft\Windows\StorageSense"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "AllowStorageSenseGlobal", 0, winreg.REG_DWORD, 0)
            logger.info("Disable Storage Sense: Applied registry change.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"AllowStorageSenseGlobal\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Disable Storage Sense failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_wifi_sense(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\WcmSvc\wifinetworkmanager\config"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "AutoConnectAllowedOEM", 0, winreg.REG_DWORD, 0)
            logger.info("Disable Wifi-Sense: Applied registry change.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"AutoConnectAllowedOEM\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Disable Wifi-Sense failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def enable_end_task(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced\TaskbarDeveloperSettings"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "TaskbarEndTask", 0, winreg.REG_DWORD, 1)
            logger.info("Enable End Task With Right Click: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"TaskbarEndTask\" /t REG_DWORD /d 1 /f")
        except Exception as e:
            logger.error(f"Enable End Task With Right Click failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def run_disk_cleanup(self, logger):
        try:
            cmd = "cleanmgr /sagerun:1"
            logger.info("Run Disk Cleanup: Initiated (command will open Disk Cleanup window)")
            logger.info(f"Command: {cmd}")
            subprocess.Popen(cmd, shell=True) # Use Popen to not block
            logger.info("Run Disk Cleanup: Command sent.")
        except Exception as e:
            logger.error(f"Run Disk Cleanup failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def set_powershell7_default(self, logger):
        try:
            logger.warning("Change Windows Terminal default: PowerShell 5 -> PowerShell 7: Placeholder - Implementation depends on specific setup.")
        except Exception as e:
            logger.error(f"Set PowerShell 7 default failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_powershell7_telemetry(self, logger):
        try:
             logger.warning("Disable Powershell 7 Telemetry: Placeholder - Implementation depends on specific setup.")
        except Exception as e:
            logger.error(f"Disable Powershell 7 Telemetry failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_recall(self, logger):
        try:
            cmd = 'DISM /Online /Disable-Feature /FeatureName:Recall'
            self.log_and_run_command(logger, cmd, "Disable Recall Feature")
            logger.info("Disable Recall: Command executed.")
        except subprocess.CalledProcessError as e:
            logger.warning(f"Disable Recall: DISM command might have failed or feature not found. Error: {e}")
        except Exception as e:
            logger.error(f"Disable Recall failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def set_hibernation_default(self, logger):
        try:
            cmd = "powercfg /h on"
            self.log_and_run_command(logger, cmd, "Enable Hibernation")
            cmd = 'powercfg /setactive "SCHEME_MIN"'
            self.log_and_run_command(logger, cmd, "Set Power Scheme to High Performance (often uses hibernation)")
            logger.info("Set Hibernation as default: Success (set to High Performance scheme)")
        except subprocess.CalledProcessError as e:
            logger.warning(f"Set Hibernation as default: Command might have failed. Error: {e}")
        except Exception as e:
            logger.error(f"Set Hibernation as default failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def set_services_manual(self, logger):
        try:
            services = [
                "diagnosticshub.standardcollector.service", # Microsoft (R) Diagnostics Hub Standard Collector Service
                "dmwappushservice", # WAP Push Message Routing Service
                "lfsvc", # Geolocation Service
                "MapsBroker", # Downloaded Maps Manager
                "NetTcpPortSharing", # Net.Tcp Port Sharing Service
                "RemoteAccess", # Routing and Remote Access
                "RemoteRegistry", # Remote Registry
                "SharedAccess", # Internet Connection Sharing (ICS)
                "TrkWks", # Distributed Link Tracking Client
                "WbioSrvc", # Windows Biometric Service
                "WlanSvc", # WLAN AutoConfig (if not needed)
                "WMPNetworkSvc", # Windows Media Player Network Sharing Service
                "XblAuthManager", # Xbox Live Auth Manager
                "XblGameSave", # Xbox Live Game Save Service
                "XboxNetApiSvc", # Xbox Live Networking Service
            ]
            for service in services:
                try:
                    cmd = f'sc config "{service}" start= demand'
                    self.log_and_run_command(logger, cmd, f"Set service {service} to Manual")
                except subprocess.CalledProcessError:
                    logger.warning(f"Set Services to Manual: Failed to set {service} to Manual (might not exist or need higher privileges)")
            logger.info("Set Services to Manual: Attempted for listed services")
        except Exception as e:
            logger.error(f"Set Services to Manual failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def debloat_brave(self, logger):
        try:
            logger.warning("Debloat Brave: Placeholder - Implementation depends on specific Brave components to remove.")
        except Exception as e:
            logger.error(f"Debloat Brave failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def debloat_edge(self, logger):
        try:
            logger.warning("Debloat Edge: Placeholder - Implementation depends on specific Edge components to remove.")
        except Exception as e:
            logger.error(f"Debloat Edge failed: {str(e)}\n{traceback.format_exc()}")
            raise

    # --- New Methods from comenzi.txt (Updated to accept logger) ---
    def disable_windows_setting_sync(self, logger):
        try:
            commands = [
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableWindowsSettingSync" /t REG_DWORD /d 2 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableWindowsSettingSyncUserOverride" /t REG_DWORD /d 1 /f',
                'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\SettingSync" /v "SyncPolicy" /t REG_DWORD /d 5 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableSettingSync" /t REG_DWORD /d 2 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableSettingSyncUserOverride" /t REG_DWORD /d 1 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Windows Setting Sync: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Windows Setting Sync failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Windows Setting Sync failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_handwriting_data_collection(self, logger):
        try:
            commands = [
                'reg add "HKCU\\Software\\Policies\\Microsoft\\InputPersonalization" /v "RestrictImplicitInkCollection" /t REG_DWORD /d 1 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\InputPersonalization" /v "RestrictImplicitInkCollection" /t REG_DWORD /d 1 /f',
                'reg add "HKCU\\Software\\Policies\\Microsoft\\InputPersonalization" /v "RestrictImplicitTextCollection" /t REG_DWORD /d 1 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\InputPersonalization" /v "RestrictImplicitTextCollection" /t REG_DWORD /d 1 /f',
                'reg add "HKCU\\Software\\Policies\\Microsoft\\Windows\\HandwritingErrorReports" /v "PreventHandwritingErrorReports" /t REG_DWORD /d 1 /f',
                'reg add "HKLM\\Software\\Policies\\Microsoft\\Windows\\HandwritingErrorReports" /v "PreventHandwritingErrorReports" /t REG_DWORD /d 1 /f',
                'reg add "HKCU\\Software\\Policies\\Microsoft\\Windows\\TabletPC" /v "PreventHandwritingDataSharing" /t REG_DWORD /d 1 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\TabletPC" /v "PreventHandwritingDataSharing" /t REG_DWORD /d 1 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\InputPersonalization" /v "AllowInputPersonalization" /t REG_DWORD /d 0 /f',
                'reg add "HKCU\\SOFTWARE\\Microsoft\\InputPersonalization\\TrainedDataStore" /v "HarvestContacts" /t REG_DWORD /d 0 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Handwriting Data Collection: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Handwriting Data Collection failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Handwriting Data Collection failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_clipboard_data_collection(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\clipboard"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "Value", 0, winreg.REG_SZ, "Deny")
            logger.info("Disable Clipboard Data Collection: Set consent to Deny.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"Value\" /t REG_SZ /d \"Deny\" /f")
        except Exception as e:
            logger.error(f"Disable Clipboard Data Collection failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def opt_out_privacy_consent(self, logger):
        try:
            key_path = r'SOFTWARE\Microsoft\Personalization\Settings'
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "AcceptedPrivacyPolicy", 0, winreg.REG_DWORD, 0)
            logger.info("Opt out of Privacy Consent: Set AcceptedPrivacyPolicy to 0.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"AcceptedPrivacyPolicy\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Opt out of Privacy Consent failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_3rd_party_telemetry(self, logger):
        try:
             # Adobe (from existing method, updated path)
             key_path_adobe = r"SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown"
             with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path_adobe) as key:
                 winreg.SetValueEx(key, "bUsageMeasurement", 0, winreg.REG_DWORD, 0)
             logger.info("Disable 3rd-party apps Telemetry (Adobe): Applied registry change.")
             logger.info(f"Command: reg add \"HKLM\\{key_path_adobe}\" /v \"bUsageMeasurement\" /t REG_DWORD /d 0 /f")

             # NVIDIA (from comenzi.txt)
             nvidia_commands = [
                 'reg add "HKLM\\SOFTWARE\\NVIDIA Corporation\\Global\\FTS" /v "EnableRID44231" /t REG_DWORD /d 0 /f',
                 'reg add "HKLM\\SOFTWARE\\NVIDIA Corporation\\Global\\FTS" /v "EnableRID64640" /t REG_DWORD /d 0 /f',
                 'reg add "HKLM\\SOFTWARE\\NVIDIA Corporation\\Global\\FTS" /v "EnableRID66610" /t REG_DWORD /d 0 /f',
                 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\nvlddmkm\\Global\\Startup" /v "SendTelemetryData" /t REG_DWORD /d 0 /f',
                 'schtasks /change /TN "NvTmMon_{B2FE1952-0186-46C3-BAEC-A80AA35AC5B8}" /DISABLE',
                 'schtasks /change /TN "NvTmRep_{B2FE1952-0186-46C3-BAEC-A80AA35AC5B8}" /DISABLE'
             ]
             for cmd in nvidia_commands:
                 try:
                     self.log_and_run_command(logger, cmd)
                 except subprocess.CalledProcessError as e:
                     logger.warning(f"Disable 3rd-party telemetry (NVIDIA): Command failed or task not found: {cmd}. Error: {e}")

             # Visual Studio (from comenzi.txt)
             vs_commands = [
                 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\VisualStudio\\IntelliCode" /v "DisableRemoteAnalysis" /t REG_DWORD /d 1 /f',
                 'reg add "HKCU\\SOFTWARE\\Microsoft\\VSCommon\\16.0\\IntelliCode" /v "DisableRemoteAnalysis" /t REG_DWORD /d 1 /f',
                 'reg add "HKCU\\SOFTWARE\\Microsoft\\VSCommon\\17.0\\IntelliCode" /v "DisableRemoteAnalysis" /t REG_DWORD /d 1 /f',
                 'reg add "HKCU\\Software\\Microsoft\\VisualStudio\\Telemetry" /v "TurnOffSwitch" /t REG_DWORD /d 1 /f',
                 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\VisualStudio\\Feedback" /v "DisableFeedbackDialog" /t REG_DWORD /d 1 /f',
                 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\VisualStudio\\Feedback" /v "DisableEmailInput" /t REG_DWORD /d 1 /f',
                 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\VisualStudio\\Feedback" /v "DisableScreenshotCapture" /t REG_DWORD /d 1 /f'
             ]
             for cmd in vs_commands:
                 self.log_and_run_command(logger, cmd)

             logger.info("Disable 3rd-party apps Telemetry (Adobe/NVIDIA/VS): Applied registry changes and task disables.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable 3rd-party apps Telemetry failed (Registry/Task error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable 3rd-party apps Telemetry failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_lockscreen_camera(self, logger):
        try:
            key_path = r"SOFTWARE\Policies\Microsoft\Windows\Personalization"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "NoLockScreenCamera", 0, winreg.REG_DWORD, 1)
            logger.info("Disable Lockscreen Camera Access: Applied registry change.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"NoLockScreenCamera\" /t REG_DWORD /d 1 /f")
        except Exception as e:
            logger.error(f"Disable Lockscreen Camera Access failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_biometrics(self, logger):
        try:
            commands = [
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Biometrics" /v "Enabled" /t REG_DWORD /d 0 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Biometrics\\Credential Provider" /v "Enabled" /t REG_DWORD /d 0 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Biometrics: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Biometrics failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Biometrics failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_cortana_bing_search(self, logger):
        try:
            commands = [
                'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Search" /v "DisableSearchBoxSuggestions" /t REG_DWORD /d 1 /f',
                'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Search" /v "CortanaInAmbientMode" /t REG_DWORD /d 0 /f',
                'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Search" /v "BingSearchEnabled" /t REG_DWORD /d 0 /f',
                'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "ShowCortanaButton" /t REG_DWORD /d 0 /f',
                'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search" /v "CanCortanaBeEnabled" /t REG_DWORD /d 0 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" /v "ConnectedSearchUseWebOverMeteredConnections" /t REG_DWORD /d 0 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" /v "AllowCortanaAboveLock" /t REG_DWORD /d 0 /f',
                'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\SearchSettings" /v "IsDynamicSearchBoxEnabled" /t REG_DWORD /d 0 /f',
                'reg add "HKLM\\SOFTWARE\\Microsoft\\PolicyManager\\default\\Experience\\AllowCortana" /v "value" /t REG_DWORD /d 0 /f',
                'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Search" /v "CortanaEnabled" /t REG_DWORD /d 0 /f',
                'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Search" /v "CortanaEnabled" /t REG_DWORD /d 0 /f',
                'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\SearchSettings" /v "IsMSACloudSearchEnabled" /t REG_DWORD /d 0 /f',
                'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\SearchSettings" /v "IsAADCloudSearchEnabled" /t REG_DWORD /d 0 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Search" /v "AllowCloudSearch" /t REG_DWORD /d 0 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Cortana & Bing Search: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Cortana & Bing Search failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Cortana & Bing Search failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_feedback_notifications(self, logger):
        try:
            commands = [
                'reg add "HKCU\\SOFTWARE\\Microsoft\\Siuf\\Rules" /v "NumberOfSIUFInPeriod" /t REG_DWORD /d 0 /f',
                'reg delete "HKCU\\SOFTWARE\\Microsoft\\Siuf\\Rules" /v "PeriodInNanoSeconds" /f 2>nul',
                'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Policies\\DataCollection" /v "DoNotShowFeedbackNotifications" /t REG_DWORD /d 1 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" /v "DoNotShowFeedbackNotifications" /t REG_DWORD /d 1 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Feedback Notifications: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Feedback Notifications failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Feedback Notifications failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_verbose_logon(self, logger):
        try:
             key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
             with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                 winreg.SetValueEx(key, "verbosestatus", 0, winreg.REG_DWORD, 1)
             logger.info("Disable Verbose Logon: Applied registry change (Set verbosestatus=1)")
             logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"verbosestatus\" /t REG_DWORD /d 1 /f")
        except Exception as e:
             logger.error(f"Disable Verbose Logon failed: {str(e)}\n{traceback.format_exc()}")
             raise

    def disable_sticky_keys(self, logger):
        try:
            key_path = r"Control Panel\Accessibility\StickyKeys"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "Flags", 0, winreg.REG_SZ, "58")
            logger.info("Disable Sticky Keys: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"Flags\" /t REG_SZ /d \"58\" /f")
        except Exception as e:
            logger.error(f"Disable Sticky Keys failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_numlock_startup(self, logger):
        try:
            key_path = r"Control Panel\Keyboard"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "InitialKeyboardIndicators", 0, winreg.REG_SZ, "0")
            logger.info("Disable Numlock on startup: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"InitialKeyboardIndicators\" /t REG_SZ /d \"0\" /f")
        except Exception as e:
            logger.error(f"Disable Numlock on startup failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_snap_flyout(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "EnableSnapAssistFlyout", 0, winreg.REG_DWORD, 0)
            logger.info("Disable Snap Flyout: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"EnableSnapAssistFlyout\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Disable Snap Flyout failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def set_classic_right_click(self, logger):
        try:
            commands = [
                'reg add "HKCU\\Software\\Classes\\CLSID\\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\\InprocServer32" /f /ve',
                'reg add "HKCU\\Software\\Classes\\CLSID\\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\\InprocServer32" /ve /t REG_SZ /d "" /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Set Classic Right Click Menu: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Set Classic Right Click Menu failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Set Classic Right Click Menu failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def show_file_extensions(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "HideFileExt", 0, winreg.REG_DWORD, 0)
            logger.info("Show File Extensions: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"HideFileExt\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Show File Extensions failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_mouse_acceleration(self, logger):
        try:
            commands = [
                'reg add "HKCU\\Control Panel\\Mouse" /v "MouseSpeed" /t REG_SZ /d "0" /f',
                'reg add "HKCU\\Control Panel\\Mouse" /v "MouseThreshold1" /t REG_SZ /d "0" /f',
                'reg add "HKCU\\Control Panel\\Mouse" /v "MouseThreshold2" /t REG_SZ /d "0" /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Mouse Acceleration: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Mouse Acceleration failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Mouse Acceleration failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_fullscreen_optimizations(self, logger):
        try:
            commands = [
                'reg add "HKCU\\System\\GameConfigStore" /v "GameDVR_DXGIHonorFSEWindowsCompatible" /t REG_DWORD /d 1 /f',
                'reg add "HKCU\\System\\GameConfigStore" /v "GameDVR_Enabled" /t REG_DWORD /d 0 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\GameDVR" /v "AllowGameDVR" /t REG_DWORD /d 0 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Fullscreen Optimizations & Game Bar: Applied registry changes.")
        except Exception as e:
            logger.error(f"Disable Fullscreen Optimizations failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_copilot(self, logger):
        try:
            commands = [
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\WindowsCopilot" /v "TurnOffWindowsCopilot" /t REG_DWORD /d 1 /f',
                'reg add "HKCU\\Software\\Policies\\Microsoft\\Windows\\WindowsCopilot" /v "TurnOffWindowsCopilot" /t REG_DWORD /d 1 /f',
                'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Notifications\\Settings" /v "AutoOpenCopilotLargeScreens" /t REG_DWORD /d 0 /f',
                'reg add "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Advanced" /v "ShowCopilotButton" /t REG_DWORD /d 0 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Copilot: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Copilot failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Copilot failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_intellicode_remote_analysis(self, logger):
        try:
            logger.info("Disable IntelliCode Remote Analysis: Already handled in 3rd-party telemetry disable.")
        except Exception as e:
            logger.error(f"Disable IntelliCode Remote Analysis failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_media_player_telemetry(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\MediaPlayer\Preferences"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "UsageTracking", 0, winreg.REG_DWORD, 0)
            logger.info("Disable Media Player Telemetry: Set UsageTracking to 0.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"UsageTracking\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Disable Media Player Telemetry failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_system_file_access_consent(self, logger):
        try:
            paths = [
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\documentsLibrary",
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\picturesLibrary",
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\videosLibrary",
                r"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\broadFileSystemAccess"
            ]
            for path in paths:
                with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, path) as key:
                    winreg.SetValueEx(key, "Value", 0, winreg.REG_SZ, "Deny")
                logger.info(f"Disable System File Access Consent: Set consent to Deny for {path.split('\\')[-1]}.")
                logger.info(f"Command: reg add \"HKLM\\{path}\" /v \"Value\" /t REG_SZ /d \"Deny\" /f")
            logger.info("Disable System File Access Consent: Applied registry changes for docs, pics, vids, broadFS.")
        except Exception as e:
            logger.error(f"Disable System File Access Consent failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_account_info_access_consent(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\userAccountInformation"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "Value", 0, winreg.REG_SZ, "Deny")
            logger.info("Disable Account Info Access Consent: Set consent to Deny.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"Value\" /t REG_SZ /d \"Deny\" /f")
        except Exception as e:
            logger.error(f"Disable Account Info Access Consent failed: {str(e)}\n{traceback.format_exc()}")
            raise

    # --- Sync Settings (from comenzi.txt) ---
    def disable_language_sync(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\SettingSync\Groups\Language"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                 winreg.SetValueEx(key, "Enabled", 0, winreg.REG_DWORD, 0)
            logger.info("Disable Language Sync: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"Enabled\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Disable Language Sync failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_credentials_sync(self, logger):
        try:
            commands = [
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableCredentialsSettingSync" /t REG_DWORD /d 2 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableCredentialsSettingSyncUserOverride" /t REG_DWORD /d 1 /f',
                'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\SettingSync\\Groups\\Credentials" /v "Enabled" /t REG_DWORD /d 0 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Credentials Sync: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Credentials Sync failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Credentials Sync failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_desktop_theme_sync(self, logger):
        try:
            commands = [
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableDesktopThemeSettingSync" /t REG_DWORD /d 2 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableDesktopThemeSettingSyncUserOverride" /t REG_DWORD /d 1 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Desktop Theme Sync: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Desktop Theme Sync failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Desktop Theme Sync failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_personalization_sync(self, logger):
        try:
            commands = [
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisablePersonalizationSettingSync" /t REG_DWORD /d 2 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisablePersonalizationSettingSyncUserOverride" /t REG_DWORD /d 1 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Personalization Sync: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Personalization Sync failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Personalization Sync failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_start_layout_sync(self, logger):
        try:
            commands = [
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableStartLayoutSettingSync" /t REG_DWORD /d 2 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableStartLayoutSettingSyncUserOverride" /t REG_DWORD /d 1 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Start Layout Sync: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Start Layout Sync failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Start Layout Sync failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_application_sync(self, logger):
        try:
            commands = [
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableApplicationSettingSync" /t REG_DWORD /d 2 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableApplicationSettingSyncUserOverride" /t REG_DWORD /d 1 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable Application Sync: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable Application Sync failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable Application Sync failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_app_sync_setting_sync(self, logger):
        try:
            commands = [
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableAppSyncSettingSync" /t REG_DWORD /d 2 /f',
                'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\SettingSync" /v "DisableAppSyncSettingSyncUserOverride" /t REG_DWORD /d 1 /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Disable App Sync Setting Sync: Applied registry changes.")
        except subprocess.CalledProcessError as e:
            logger.error(f"Disable App Sync Setting Sync failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Disable App Sync Setting Sync failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_sync_on_paid_network(self, logger):
        try:
            key_path = r"SOFTWARE\Policies\Microsoft\Windows\SettingSync"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "DisableSyncOnPaidNetwork", 0, winreg.REG_DWORD, 1)
            logger.info("Disable Sync on Paid Network: Applied registry change.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"DisableSyncOnPaidNetwork\" /t REG_DWORD /d 1 /f")
        except Exception as e:
            logger.error(f"Disable Sync on Paid Network failed: {str(e)}\n{traceback.format_exc()}")
            raise

    # --- Section 2: Advanced Tweaks ---
    def create_section2(self):
       try:
            frame = ttk.Frame(self.section_container, style=f"{self.current_theme.capitalize()}.TFrame")
            self.sections.append(frame)

            tweaks = [
                "Block Adobe Network",
                "Remove Edge",
                "Debloat Edge",
                "Remove Microsoft Apps",
                "Install Chrome",
                "Install Firefox"
            ]
            self.section2_vars = {tweak: BooleanVar(value=False) for tweak in tweaks}

            # --- Layout Split ---
            left_frame = ttk.Frame(frame, style=f"{self.current_theme.capitalize()}.TFrame")
            left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

            right_frame = ttk.Frame(frame, width=400, style=f"{self.current_theme.capitalize()}.TFrame")
            right_frame.pack_propagate(False)
            right_frame.pack(side="right", fill="y", padx=(5, 0))

            # --- Left Column: Checkboxes ---
            canvas = tk.Canvas(left_frame, bg=self.style.lookup(f"{self.current_theme.capitalize()}.TFrame", 'background'), highlightthickness=0)
            scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=canvas.yview, style="Vertical.TScrollbar")
            scrollable_frame = ttk.Frame(canvas, style=f"{self.current_theme.capitalize()}.TFrame")

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # --- Bind Mousewheel to Scrollable Frame for Section 2 ---
            scrollable_frame.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
            canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
            # --- End Mousewheel Binding ---

            for i, tweak in enumerate(tweaks):
                chk = ttk.Checkbutton(
                    scrollable_frame,
                    text=tweak,
                    variable=self.section2_vars[tweak],
                    command=lambda t=tweak: logging.info(f"Section2 {t}: Toggled to {self.section2_vars[t].get()}"),
                    style=f"{self.current_theme.capitalize()}.TCheckbutton"
                )
                chk.grid(row=i, column=0, sticky="w", padx=10, pady=5)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            # --- Right Column: Execution Output ---
            output_label = ttk.Label(right_frame, text="Execution Output:", font=("Segoe UI", 10, "bold"), style=f"{self.current_theme.capitalize()}.TLabel")
            output_label.pack(anchor="w", padx=5, pady=(5, 0))

            self.output_text2 = tk.Text(right_frame, state='disabled', wrap='word',
                                        bg='black' if self.current_theme == 'dark' else 'white',
                                        fg='white' if self.current_theme == 'dark' else 'black')
            output_scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=self.output_text2.yview, style="Vertical.TScrollbar")
            self.output_text2.configure(yscrollcommand=output_scrollbar.set)

            self.output_text2.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
            output_scrollbar.pack(side="right", fill="y", pady=5, padx=(0, 5))

            # --- Bottom Frame (MOVED to LEFT FRAME) ---
            bottom_frame = ttk.Frame(left_frame, style=f"{self.current_theme.capitalize()}.TFrame") # Changed parent to left_frame
            bottom_frame.pack(side="bottom", fill="x", padx=10, pady=10) # Adjusted padding if needed

            self.progress2 = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate", length=300)
            self.progress2.pack(side="left", fill="x", expand=True, padx=(0, 10))

            execute_btn = ttk.Button(
                bottom_frame,
                text="Execute",
                command=lambda: threading.Thread(target=self.execute_section2, daemon=True).start(),
                style='Green.TButton'
            )
            execute_btn.pack(side="right")

            logging.info("Section 2 created successfully")
       except Exception as e:
            logging.error(f"Create section 2 failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def execute_section2(self):
        try:
            if not hasattr(self, 'output_text2'):
                logging.error("Output text widget for section 2 not found.")
                return

            exec_logger = logging.getLogger(f"Section2_Exec_{threading.get_ident()}")
            exec_logger.setLevel(logging.INFO)
            exec_logger.handlers.clear()
            text_handler = TextHandler(self.output_text2)
            formatter = logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt='%H:%M:%S')
            text_handler.setFormatter(formatter)
            exec_logger.addHandler(text_handler)

            selected = [name for name, var in self.section2_vars.items() if var.get()]
            if not selected:
                exec_logger.info("No options selected for execution.")
                logging.info("Section2: No options selected for execution.")
                return

            self.progress2["value"] = 0
            self.progress2["maximum"] = len(selected)

            self.output_text2.configure(state='normal')
            self.output_text2.delete(1.0, tk.END)
            self.output_text2.configure(state='disabled')

            for i, option in enumerate(selected):
                try:
                    exec_logger.info(f"Starting: {option}")
                    if option == "Block Adobe Network":
                        self.adobe_network_block(exec_logger)
                    elif option == "Remove Edge":
                        self.remove_edge(exec_logger) # Warning: Destructive
                    elif option == "Debloat Edge":
                        self.debloat_edge(exec_logger)
                    elif option == "Remove Microsoft Apps":
                        self.remove_microsoft_apps(exec_logger)
                    elif option == "Install Chrome":
                        self.install_software(exec_logger, "Google Chrome", "Google.Chrome")
                    elif option == "Install Firefox":
                        self.install_software(exec_logger, "Mozilla Firefox", "Mozilla.Firefox")
                    exec_logger.info(f"Completed: {option}")
                    logging.info(f"Section2 {option}: Enabled")
                    self.progress2["value"] = i + 1
                    self.progress2.update()
                except Exception as e:
                    error_msg = f"Failed: {option} - {str(e)}"
                    exec_logger.error(error_msg)
                    logging.error(f"Section2 {option}: Failed - {str(e)}\n{traceback.format_exc()}")

            for name, var in self.section2_vars.items():
                var.set(False)

            exec_logger.info(f"Execution Complete: Applied {len(selected)} tweaks!")
            logging.info(f"Section2: Finished applying {len(selected)} tweaks")

        except Exception as e:
            logging.error(f"Execute section 2 failed: {str(e)}\n{traceback.format_exc()}")
            if hasattr(self, 'output_text2'):
                try:
                    fatal_logger = logging.getLogger(f"Section2_Exec_{threading.get_ident()}")
                    if not fatal_logger.handlers:
                        fatal_logger.addHandler(TextHandler(self.output_text2))
                        fatal_logger.setLevel(logging.ERROR)
                    fatal_logger.error(f"Execution Failed: {str(e)}")
                except:
                    pass
            raise

    # --- Section 2 Methods (Updated to accept logger) ---
    def adobe_network_block(self, logger):
        try:
            key_path = r"SOFTWARE\Policies\Adobe\Acrobat Reader\DC\FeatureLockDown\cCloud"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "bAdobeSendPluginToggle", 0, winreg.REG_DWORD, 1)
            logger.info("Adobe Network Block: Applied registry change.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"bAdobeSendPluginToggle\" /t REG_DWORD /d 1 /f")
        except Exception as e:
            logger.error(f"Adobe Network Block failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def remove_edge(self, logger):
        try:
            logger.warning("Remove Edge: This is a complex and potentially unstable operation.")
            logger.warning("Remove Edge: Attempting removal via winget (may not fully succeed)...")
            logger.info("Command: winget uninstall \"Microsoft Edge\"")
            result = subprocess.run('winget uninstall "Microsoft Edge"', shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                logger.info("Remove Edge: Winget command initiated.")
                logging.info("Remove Edge: Winget command initiated.")
            else:
                logger.error("Remove Edge: Winget command failed or Edge not found via winget.")
                logging.error("Remove Edge: Winget command failed or Edge not found via winget.")
                logging.error(f"Remove Edge: Winget output: {result.stdout}\nError: {result.stderr}")
        except Exception as e:
            logger.error(f"Remove Edge failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def remove_microsoft_apps(self, logger):
        try:
            apps_to_remove = [
                "MicrosoftCorporationII.MicrosoftFamily",
                "Microsoft.OutlookForWindows",
                "Clipchamp.Clipchamp",
                "Microsoft.3DBuilder",
                "Microsoft.Microsoft3DViewer",
                "Microsoft.BingWeather",
                "Microsoft.BingSports",
                "Microsoft.BingFinance",
                "Microsoft.MicrosoftOfficeHub",
                "Microsoft.BingNews",
                "Microsoft.Office.OneNote",
                "Microsoft.Office.Sway",
                "Microsoft.WindowsPhone",
                "Microsoft.CommsPhone",
                "Microsoft.YourPhone",
                "Microsoft.Getstarted",
                "Microsoft.549981C3F5F10", # Cortana
                "Microsoft.Messaging",
                "Microsoft.WindowsSoundRecorder",
                "Microsoft.MixedReality.Portal",
                "Microsoft.WindowsFeedbackHub",
                "Microsoft.WindowsAlarms",
                "Microsoft.WindowsCamera",
                "Microsoft.MSPaint",
                "Microsoft.WindowsMaps",
                "Microsoft.MinecraftUWP",
                "Microsoft.People",
                "Microsoft.Wallet",
                "Microsoft.Print3D",
                # --- More from comenzi.txt ---
                "king.com.CandyCrushSodaSaga",
                "ShazamEntertainmentLtd.Shazam",
                "Flipboard.Flipboard",
                "9E2F88E3.Twitter", # Old Twitter package
                "ClearChannelRadioDigital.iHeartRadio",
                "D5EA27B7.Duolingo-LearnLanguagesforFree",
                "AdobeSystemsIncorporated.AdobePhotoshopExpress",
                "PandoraMediaInc.29680B314EFC2",
                "46928bounde.EclipseManager",
                "ActiproSoftwareLLC.562882FEEB491",
                "SpotifyAB.Spotify", # If removing Spotify app
                # --- Media Extensions ---
                "Microsoft.WebpImageExtension",
                "Microsoft.HEVCVideoExtension",
                "Microsoft.RawImageExtension",
                "Microsoft.WebMediaExtensions"
            ]

            failed_apps = []
            for app in apps_to_remove:
                try:
                    cmd = f'PowerShell -ExecutionPolicy Unrestricted -Command "Get-AppxPackage \\"{app}\\"| Remove-AppxPackage"'
                    logger.info(f"Removing App: {app}")
                    logger.info(f"Command: {cmd}")
                    result = subprocess.run(cmd, shell=True, capture_output=True, text=True, check=True)
                    logger.info(f"Remove Microsoft App '{app}': Success")
                except subprocess.CalledProcessError as e:
                    logger.warning(f"Remove Microsoft App '{app}': Might not be installed or failed. Error: {e}")
                    logger.warning(f"Remove Microsoft App '{app}': Stdout: {e.stdout}, Stderr: {e.stderr}")
                    failed_apps.append(app)
                except Exception as e:
                    logger.error(f"Remove Microsoft App '{app}' failed unexpectedly: {str(e)}")
                    failed_apps.append(app)

            if failed_apps:
                logger.info(f"Remove Microsoft Apps: Finished. Apps that might not have been removed or failed: {', '.join(failed_apps)}")
            else:
                logger.info("Remove Microsoft Apps: All targeted apps processed.")

        except Exception as e:
            logger.error(f"Remove Microsoft Apps failed: {str(e)}\n{traceback.format_exc()}")
            raise


    def install_software(self, logger, name, package_id):
        try:
            if not package_id:
                logger.warning(f"Install {name}: No package ID provided, skipping.")
                return
            cmd = f'winget install --id {package_id} -e --accept-source-agreements --accept-package-agreements'
            logger.info(f"Installing {name}...")
            logger.info(f"Command: {cmd}")
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode == 0:
                logger.info(f"Install {name}: Success")
            else:
                logger.error(f"Install {name}: Failed - {result.stderr}")
                raise Exception(f"Installation of {name} failed")
        except Exception as e:
            logger.error(f"Install {name} failed: {str(e)}\n{traceback.format_exc()}")
            raise

    # --- Section 3: Customize Preferences ---
    def create_section3(self):
        try:
            frame = ttk.Frame(self.section_container, style=f"{self.current_theme.capitalize()}.TFrame")
            self.sections.append(frame)

            tweaks = [
                "Disable Background Apps",
                "Disable Fullscreen Optimizations",
                "Disable Copilot",
                "Disable Intel MM",
                "Set Display for Performance",
                "Set Classic Right Click Menu",
                "Enable End Task With Right Click",
                "Hide Search Button",
                "Hide Task View Button",
                "Enable Snap Window",
                "Enable Snap Assist Flyout",
                "Enable Snap Assist Suggestion",
                "Enable Mouse Acceleration",
                "Enable Sticky Keys",
                "Show Hidden Files",
                "Show File Extensions",
                "Enable S3 Sleep"
            ]
            self.section3_vars = {tweak: BooleanVar(value=False) for tweak in tweaks}

            # --- Layout Split ---
            left_frame = ttk.Frame(frame, style=f"{self.current_theme.capitalize()}.TFrame")
            left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

            right_frame = ttk.Frame(frame, width=400, style=f"{self.current_theme.capitalize()}.TFrame")
            right_frame.pack_propagate(False)
            right_frame.pack(side="right", fill="y", padx=(5, 0))

            # --- Left Column: Checkboxes ---
            canvas = tk.Canvas(left_frame, bg=self.style.lookup(f"{self.current_theme.capitalize()}.TFrame", 'background'), highlightthickness=0)
            scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=canvas.yview, style="Vertical.TScrollbar")
            scrollable_frame = ttk.Frame(canvas, style=f"{self.current_theme.capitalize()}.TFrame")

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # --- Bind Mousewheel to Scrollable Frame for Section 3 ---
            scrollable_frame.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
            canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
            # --- End Mousewheel Binding ---

            for i, tweak in enumerate(tweaks):
                chk = ttk.Checkbutton(
                    scrollable_frame,
                    text=tweak,
                    variable=self.section3_vars[tweak],
                    command=lambda t=tweak: logging.info(f"Section3 {t}: Toggled to {self.section3_vars[t].get()}"),
                    style=f"{self.current_theme.capitalize()}.TCheckbutton"
                )
                chk.grid(row=i, column=0, sticky="w", padx=10, pady=5)

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            # --- Right Column: Execution Output ---
            output_label = ttk.Label(right_frame, text="Execution Output:", font=("Segoe UI", 10, "bold"), style=f"{self.current_theme.capitalize()}.TLabel")
            output_label.pack(anchor="w", padx=5, pady=(5, 0))

            self.output_text3 = tk.Text(right_frame, state='disabled', wrap='word',
                                        bg='black' if self.current_theme == 'dark' else 'white',
                                        fg='white' if self.current_theme == 'dark' else 'black')
            output_scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=self.output_text3.yview, style="Vertical.TScrollbar")
            self.output_text3.configure(yscrollcommand=output_scrollbar.set)

            self.output_text3.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
            output_scrollbar.pack(side="right", fill="y", pady=5, padx=(0, 5))

            # --- Bottom Frame (MOVED to LEFT FRAME) ---
            bottom_frame = ttk.Frame(left_frame, style=f"{self.current_theme.capitalize()}.TFrame") # Changed parent to left_frame
            bottom_frame.pack(side="bottom", fill="x", padx=10, pady=10) # Adjusted padding if needed

            self.progress3 = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate", length=300)
            self.progress3.pack(side="left", fill="x", expand=True, padx=(0, 10))

            execute_btn = ttk.Button(
                bottom_frame,
                text="Execute",
                command=lambda: threading.Thread(target=self.execute_section3, daemon=True).start(),
                style='Green.TButton'
            )
            execute_btn.pack(side="right")

            logging.info("Section 3 created successfully")
        except Exception as e:
            logging.error(f"Create section 3 failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def execute_section3(self):
        try:
            if not hasattr(self, 'output_text3'):
                logging.error("Output text widget for section 3 not found.")
                return

            exec_logger = logging.getLogger(f"Section3_Exec_{threading.get_ident()}")
            exec_logger.setLevel(logging.INFO)
            exec_logger.handlers.clear()
            text_handler = TextHandler(self.output_text3)
            formatter = logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt='%H:%M:%S')
            text_handler.setFormatter(formatter)
            exec_logger.addHandler(text_handler)

            selected = [name for name, var in self.section3_vars.items() if var.get()]
            if not selected:
                exec_logger.info("No options selected for execution.")
                logging.info("Section3: No options selected for execution.")
                return

            self.progress3["value"] = 0
            self.progress3["maximum"] = len(selected)

            self.output_text3.configure(state='normal')
            self.output_text3.delete(1.0, tk.END)
            self.output_text3.configure(state='disabled')

            for i, option in enumerate(selected):
                try:
                    exec_logger.info(f"Starting: {option}")
                    if option == "Disable Background Apps":
                        self.disable_background_apps(exec_logger)
                    elif option == "Disable Fullscreen Optimizations":
                        self.disable_fullscreen_optimizations(exec_logger) # Already implemented
                    elif option == "Disable Copilot":
                        self.disable_copilot(exec_logger) # Already implemented
                    elif option == "Disable Intel MM":
                        self.disable_intel_mm(exec_logger)
                    elif option == "Set Display for Performance":
                        self.set_display_performance(exec_logger)
                    elif option == "Set Classic Right Click Menu":
                        self.set_classic_right_click(exec_logger) # Already implemented
                    elif option == "Enable End Task With Right Click":
                        self.enable_end_task(exec_logger) # Already implemented
                    elif option == "Hide Search Button":
                        self.hide_search_button(exec_logger)
                    elif option == "Hide Task View Button":
                        self.hide_task_view_button(exec_logger)
                    elif option == "Enable Snap Window":
                        self.enable_snap_window(exec_logger)
                    elif option == "Enable Snap Assist Flyout":
                        self.enable_snap_assist_flyout(exec_logger)
                    elif option == "Enable Snap Assist Suggestion":
                        self.enable_snap_assist_suggestion(exec_logger)
                    elif option == "Enable Mouse Acceleration":
                        self.enable_mouse_acceleration(exec_logger)
                    elif option == "Enable Sticky Keys":
                        self.enable_sticky_keys(exec_logger)
                    elif option == "Show Hidden Files":
                        self.show_hidden_files(exec_logger)
                    elif option == "Show File Extensions":
                        self.show_file_extensions(exec_logger) # Already implemented
                    elif option == "Enable S3 Sleep":
                        self.enable_s3_sleep(exec_logger)
                    exec_logger.info(f"Completed: {option}")
                    logging.info(f"Section3 {option}: Enabled")
                    self.progress3["value"] = i + 1
                    self.progress3.update()
                except Exception as e:
                    error_msg = f"Failed: {option} - {str(e)}"
                    exec_logger.error(error_msg)
                    logging.error(f"Section3 {option}: Failed - {str(e)}\n{traceback.format_exc()}")

            for name, var in self.section3_vars.items():
                var.set(False)

            exec_logger.info(f"Execution Complete: Applied {len(selected)} tweaks!")
            logging.info(f"Section3: Finished applying {len(selected)} tweaks")

        except Exception as e:
            logging.error(f"Execute section 3 failed: {str(e)}\n{traceback.format_exc()}")
            if hasattr(self, 'output_text3'):
                try:
                    fatal_logger = logging.getLogger(f"Section3_Exec_{threading.get_ident()}")
                    if not fatal_logger.handlers:
                        fatal_logger.addHandler(TextHandler(self.output_text3))
                        fatal_logger.setLevel(logging.ERROR)
                    fatal_logger.error(f"Execution Failed: {str(e)}")
                except:
                    pass
            raise

    # --- Section 3 Methods (Updated to accept logger) ---
    def disable_background_apps(self, logger):
        try:
            key_path = r"SOFTWARE\Policies\Microsoft\Windows\AppPrivacy"
            with winreg.CreateKey(winreg.HKEY_LOCAL_MACHINE, key_path) as key:
                winreg.SetValueEx(key, "LetAppsRunInBackground", 0, winreg.REG_DWORD, 2)
            logger.info("Disable background apps: Applied registry change.")
            logger.info(f"Command: reg add \"HKLM\\{key_path}\" /v \"LetAppsRunInBackground\" /t REG_DWORD /d 2 /f")
        except Exception as e:
            logger.error(f"Disable background apps failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def disable_intel_mm(self, logger):
        try:
            logger.warning("disable_intel_mm: Placeholder function, implementation needed (e.g., disable Intel ME service)")
        except Exception as e:
            logger.error(f"Disable Intel MM failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def set_display_performance(self, logger):
        try:
            key_path = r"Control Panel\Desktop"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "VisualFXSetting", 0, winreg.REG_DWORD, 2) # 2 = Adjust for best performance
            logger.info("Set display for performance: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"VisualFXSetting\" /t REG_DWORD /d 2 /f")
        except Exception as e:
            logger.error(f"Set display for performance failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def hide_search_button(self, logger):
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Search"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "SearchboxTaskbarMode", 0, winreg.REG_DWORD, 0)
            logger.info("Hide Search Button: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"SearchboxTaskbarMode\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Hide Search Button failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def hide_task_view_button(self, logger):
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "ShowTaskViewButton", 0, winreg.REG_DWORD, 0)
            logger.info("Hide Task View Button: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"ShowTaskViewButton\" /t REG_DWORD /d 0 /f")
        except Exception as e:
            logger.error(f"Hide Task View Button failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def enable_snap_window(self, logger):
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "EnableSnapBar", 0, winreg.REG_DWORD, 1)
            logger.info("Enable Snap Window: Applied registry change (Set EnableSnapBar=1)")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"EnableSnapBar\" /t REG_DWORD /d 1 /f")
        except Exception as e:
            logger.error(f"Enable Snap Window failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def enable_snap_assist_flyout(self, logger):
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "EnableSnapAssistFlyout", 0, winreg.REG_DWORD, 1)
            logger.info("Enable Snap Assist Flyout: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"EnableSnapAssistFlyout\" /t REG_DWORD /d 1 /f")
        except Exception as e:
            logger.error(f"Enable Snap Assist Flyout failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def enable_snap_assist_suggestion(self, logger):
        try:
            key_path = r"Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "EnableSnapAssistSuggestion", 0, winreg.REG_DWORD, 1)
            logger.info("Enable Snap Assist Suggestion: Applied registry change (Set EnableSnapAssistSuggestion=1)")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"EnableSnapAssistSuggestion\" /t REG_DWORD /d 1 /f")
        except Exception as e:
            logger.error(f"Enable Snap Assist Suggestion failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def enable_mouse_acceleration(self, logger):
        try:
            commands = [
                'reg add "HKCU\\Control Panel\\Mouse" /v "MouseSpeed" /t REG_SZ /d "1" /f',
                'reg add "HKCU\\Control Panel\\Mouse" /v "MouseThreshold1" /t REG_SZ /d "6" /f',
                'reg add "HKCU\\Control Panel\\Mouse" /v "MouseThreshold2" /t REG_SZ /d "10" /f'
            ]
            for cmd in commands:
                self.log_and_run_command(logger, cmd)
            logger.info("Enable Mouse Acceleration: Applied registry changes (reverted disable settings).")
        except subprocess.CalledProcessError as e:
            logger.error(f"Enable Mouse Acceleration failed (Registry error): {str(e)}")
            raise
        except Exception as e:
            logger.error(f"Enable Mouse Acceleration failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def enable_sticky_keys(self, logger):
        try:
            key_path = r"Control Panel\Accessibility\StickyKeys"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "Flags", 0, winreg.REG_SZ, "506")
            logger.info("Enable Sticky Keys: Applied registry change (Reverted to default Flags=506)")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"Flags\" /t REG_SZ /d \"506\" /f")
        except Exception as e:
            logger.error(f"Enable Sticky Keys failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def show_hidden_files(self, logger):
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
                winreg.SetValueEx(key, "Hidden", 0, winreg.REG_DWORD, 1)
            logger.info("Show Hidden Files: Applied registry change.")
            logger.info(f"Command: reg add \"HKCU\\{key_path}\" /v \"Hidden\" /t REG_DWORD /d 1 /f")
        except Exception as e:
            logger.error(f"Show Hidden Files failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def enable_s3_sleep(self, logger):
        try:
            logger.info("Enable S3 Sleep: Command 'powercfg /setactive SCHEME_CURRENT' executed. Ensure S3 is enabled in BIOS/UEFI and power plan settings.")
            logger.info("Command: powercfg /setactive \"SCHEME_CURRENT\"")
            subprocess.run('powercfg /setactive "SCHEME_CURRENT"', shell=True, check=True)
        except subprocess.CalledProcessError as e:
            logger.warning(f"Enable S3 Sleep: Command might have failed or S3 not supported. Error: {e}")
        except Exception as e:
            logger.error(f"Enable S3 sleep failed: {str(e)}\n{traceback.format_exc()}")
            raise

    # --- Section 4: Install Software ---
    def create_section4(self):
        try:
            frame = ttk.Frame(self.section_container, style=f"{self.current_theme.capitalize()}.TFrame")
            self.sections.append(frame)

            software_categories = {
                "Browsers": [
                    ("Google Chrome", "Google.Chrome"),
                    ("Mozilla Firefox", "Mozilla.Firefox"),
                    ("Brave Browser", "Brave.Brave")
                ],
                "Development": [
                    ("Git", "Git.Git"),
                    ("Node.js LTS", "OpenJS.NodeJS.LTS"),
                    ("Python 3", "Python.Python.3"),
                    ("Visual Studio Code", "Microsoft.VisualStudioCode")
                ],
                "Documents": [
                    ("Adobe Acrobat Reader", "Adobe.Acrobat.Reader.64-bit"),
                    ("LibreOffice", "TheDocumentFoundation.LibreOffice"),
                    ("Notepad++", "Notepad++.Notepad++"),
                    ("OnlyOffice", "ONLYOFFICE.DesktopEditors")
                ],
                "Multimedia": [
                    ("GIMP (Image Editor)", "GIMP.GIMP"),
                    ("OBS Studio", "OBSProject.OBSStudio"),
                    ("Paint.NET", "dotPDN.PaintDotNet"),
                    ("VLC (Video Player)", "VideoLAN.VLC")
                ],
                "Utilities": [
                    ("7-Zip", "7zip.7zip"),
                    ("AnyDesk", "AnyDeskSoftwareGmbH.AnyDesk"),
                    ("Bitwarden", "Bitwarden.Bitwarden"),
                    ("Malwarebytes", "Malwarebytes.Malwarebytes"),
                    ("TeamViewer", "TeamViewer.TeamViewer"),
                    ("qBittorrent", "qBittorrent.qBittorrent")
                ]
            }

            self.section4_vars = {}
            row_counter = 0

            # --- Layout Split ---
            left_frame = ttk.Frame(frame, style=f"{self.current_theme.capitalize()}.TFrame")
            left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

            right_frame = ttk.Frame(frame, width=400, style=f"{self.current_theme.capitalize()}.TFrame")
            right_frame.pack_propagate(False)
            right_frame.pack(side="right", fill="y", padx=(5, 0))

            # --- Left Column: Checkboxes ---
            canvas = tk.Canvas(left_frame, bg=self.style.lookup(f"{self.current_theme.capitalize()}.TFrame", 'background'), highlightthickness=0)
            scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=canvas.yview, style="Vertical.TScrollbar")
            scrollable_frame = ttk.Frame(canvas, style=f"{self.current_theme.capitalize()}.TFrame")

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # --- Bind Mousewheel to Scrollable Frame for Section 4 ---
            scrollable_frame.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
            canvas.bind("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1*(event.delta/120)), "units"))
            # --- End Mousewheel Binding ---

            for category, apps in software_categories.items():
                cat_label = ttk.Label(scrollable_frame, text=category, font=("Segoe UI", 10, "bold"), style=f"{self.current_theme.capitalize()}.TLabel")
                cat_label.grid(row=row_counter, column=0, sticky="w", padx=10, pady=(10, 2))
                row_counter += 1
                for name, package_id in apps:
                    var = BooleanVar(value=False)
                    self.section4_vars[name] = {"var": var, "package_id": package_id}
                    chk = ttk.Checkbutton(
                        scrollable_frame,
                        text=name,
                        variable=var,
                        command=lambda n=name: logging.info(f"Section4 {n}: Toggled to {self.section4_vars[n]['var'].get()}"),
                        style=f"{self.current_theme.capitalize()}.TCheckbutton"
                    )
                    chk.grid(row=row_counter, column=0, sticky="w", padx=20, pady=2)
                    row_counter += 1

            canvas.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")

            # --- Right Column: Execution Output ---
            output_label = ttk.Label(right_frame, text="Installation Output:", font=("Segoe UI", 10, "bold"), style=f"{self.current_theme.capitalize()}.TLabel")
            output_label.pack(anchor="w", padx=5, pady=(5, 0))

            self.output_text4 = tk.Text(right_frame, state='disabled', wrap='word',
                                        bg='black' if self.current_theme == 'dark' else 'white',
                                        fg='white' if self.current_theme == 'dark' else 'black')
            output_scrollbar = ttk.Scrollbar(right_frame, orient="vertical", command=self.output_text4.yview, style="Vertical.TScrollbar")
            self.output_text4.configure(yscrollcommand=output_scrollbar.set)

            self.output_text4.pack(side="left", fill="both", expand=True, padx=(5, 0), pady=5)
            output_scrollbar.pack(side="right", fill="y", pady=5, padx=(0, 5))

            # --- Bottom Frame (MOVED to LEFT FRAME) ---
            bottom_frame = ttk.Frame(left_frame, style=f"{self.current_theme.capitalize()}.TFrame") # Changed parent to left_frame
            bottom_frame.pack(side="bottom", fill="x", padx=10, pady=10) # Adjusted padding if needed

            self.progress4 = ttk.Progressbar(bottom_frame, orient="horizontal", mode="determinate", length=300)
            self.progress4.pack(side="left", fill="x", expand=True, padx=(0, 10))

            execute_btn = ttk.Button(
                bottom_frame,
                text="Install Selected",
                command=lambda: threading.Thread(target=self.execute_section4, daemon=True).start(),
                style='Green.TButton'
            )
            execute_btn.pack(side="right")

            logging.info("Section 4 created successfully")
        except Exception as e:
            logging.error(f"Create section 4 failed: {str(e)}\n{traceback.format_exc()}")
            raise

    def execute_section4(self):
        try:
            if not hasattr(self, 'output_text4'):
                logging.error("Output text widget for section 4 not found.")
                return

            exec_logger = logging.getLogger(f"Section4_Exec_{threading.get_ident()}")
            exec_logger.setLevel(logging.INFO)
            exec_logger.handlers.clear()
            text_handler = TextHandler(self.output_text4)
            formatter = logging.Formatter("[%(asctime)s] %(levelname)s: %(message)s", datefmt='%H:%M:%S')
            text_handler.setFormatter(formatter)
            exec_logger.addHandler(text_handler)

            selected = [(name, data["package_id"]) for name, data in self.section4_vars.items() if data["var"].get() and data["package_id"]]
            if not selected:
                exec_logger.info("No software selected for installation.")
                logging.info("Section4: No software selected for installation.")
                return

            self.progress4["value"] = 0
            self.progress4["maximum"] = len(selected)

            self.output_text4.configure(state='normal')
            self.output_text4.delete(1.0, tk.END)
            self.output_text4.configure(state='disabled')

            for i, (name, package_id) in enumerate(selected):
                try:
                    exec_logger.info(f"Installing: {name}")
                    self.install_software(exec_logger, name, package_id) # Reuse existing method
                    exec_logger.info(f"Installed: {name}")
                    logging.info(f"Section4 {name}: Installed")
                    self.progress4["value"] = i + 1
                    self.progress4.update()
                except Exception as e:
                    error_msg = f"Failed to install {name}: {str(e)}"
                    exec_logger.error(error_msg)
                    logging.error(f"Section4 {name}: Failed - {str(e)}\n{traceback.format_exc()}")

            # Deselect checkboxes after installation
            for data in self.section4_vars.values():
                data["var"].set(False)

            exec_logger.info(f"Installation Complete: Installed {len(selected)} applications!")
            logging.info(f"Section4: Finished installing {len(selected)} applications")

        except Exception as e:
            logging.error(f"Execute section 4 failed: {str(e)}\n{traceback.format_exc()}")
            if hasattr(self, 'output_text4'):
                try:
                    fatal_logger = logging.getLogger(f"Section4_Exec_{threading.get_ident()}")
                    if not fatal_logger.handlers:
                        fatal_logger.addHandler(TextHandler(self.output_text4))
                        fatal_logger.setLevel(logging.ERROR)
                    fatal_logger.error(f"Installation Failed: {str(e)}")
                except:
                    pass
            raise

# --- Main Execution ---
if __name__ == "__main__":
    request_admin_privileges()
    root = tk.Tk()
    app = Win11Optimizator(root)
    root.mainloop()
    logging.info("--- Application closed ---")
