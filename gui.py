"""
Prepaid Expense Workpaper Automation – GUI Module
All UI layout, widgets, and visual logic live here.
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import threading
import requests
import sys
import os
import subprocess

CURRENT_VERSION = "v1.1.0"
GITHUB_REPO = "shuhb021/Prepaid-Tool"


# ─────────────────── Color Palette ───────────────────
DARK_THEME    = "#3A1C5E"   # Header
MID_THEME     = "#6B37A8"   # Button hover
ACCENT_THEME  = "#512882"   # Brand Purple
LIGHT_THEME   = "#EDE5F5"   # Nav active hover & Progress BG
BANNER_BG     = "#F0E6FF"
BANNER_TEXT   = "#512882"
MAIN_BG       = "#F4F5F7"   # Light Gray/White Main Background
CARD_BG       = "#FFFFFF"   # Pure White Cards
STAT_HEADER   = "#512882"
BORDER_CLR    = "#D4C8E8"
MUTED_TEXT    = "#6B7280"
GREEN_OK      = "#16A34A"
STATUS_COLOR  = "#856CA9"
TEXT_DARK     = "#111827"   # Dark text for light background
SEPARATOR_CLR = "#E5E7EB"

NAV_ITEMS = []


class PrepaidApp(ctk.CTk):
    """Modern Prepaid Expense Workpaper Automation tool.

    Parameters
    ----------
    process_fn : callable
        The backend processing function with signature::

            process_fn(basefile, client_name, fy_start, fy_end,
                       method, sample_n, progress_callback)
            -> (total_rows, sample_total, addition_count,
                addition_sum, addition_pct, output_path)

        ``progress_callback(value: float, text: str)`` is called by the
        backend to report progress (0.0 – 1.0).
    """

    def __init__(self, process_fn=None):
        super().__init__()

        self.process_fn = process_fn

        self.title("Prepaid Expense Workpaper Automation")
        self.geometry("1100x720")
        self.minsize(1050, 680)
        self.configure(fg_color=MAIN_BG)

        ctk.set_appearance_mode("light")

        # Layout: 2 columns (sidebar | main), 2 rows (header | body)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._build_header()
        self._build_sidebar()
        self._build_main()

        # Check for updates in background
        threading.Thread(target=self._check_updates_bg, daemon=True).start()

    # ─────────────── Auto-Updater ───────────────
    def _check_updates_bg(self):
        try:
            url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
            response = requests.get(url, timeout=5)
            response.raise_for_status()
            data = response.json()
            latest_version = data.get("tag_name")
            
            if latest_version and latest_version != CURRENT_VERSION:
                self.after(0, self._prompt_update, latest_version, data)
        except requests.exceptions.HTTPError as e:
            if e.response.status_code != 404:
                print("Update check failed:", e)
        except requests.exceptions.RequestException:
            pass # Ignore network errors silently
        except Exception as e:
            print("Update check failed:", e)

    def _prompt_update(self, latest_version, data):
        ans = messagebox.askyesno(
            "Update Available",
            f"A new version ({latest_version}) is available. You are currently on {CURRENT_VERSION}.\n\nDo you want to update now?"
        )
        if ans:
            assets = data.get("assets", [])
            exe_url = None
            for asset in assets:
                if asset["name"].endswith(".exe"):
                    exe_url = asset["browser_download_url"]
                    break
            
            if exe_url:
                threading.Thread(target=self._download_and_update, args=(exe_url,), daemon=True).start()
            else:
                messagebox.showerror("Update Error", "No executable found in the latest release.")

    def _download_and_update(self, exe_url):
        self.after(0, lambda: self.btn_generate.configure(state="disabled", text="Downloading Update..."))
        try:
            response = requests.get(exe_url, stream=True)
            response.raise_for_status()
            
            new_exe_path = "Prepaid Tool_new.exe"
            with open(new_exe_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            bat_path = "update.bat"
            current_exe = sys.executable
            if not getattr(sys, 'frozen', False):
                self.after(0, lambda: messagebox.showinfo("Update", "Update downloaded, but you are not running the compiled .exe"))
                self.after(0, lambda: self.btn_generate.configure(state="normal", text="  \U0001F4CB  Generate Workpaper"))
                return
                
            exe_name = os.path.basename(current_exe)
            
            bat_content = f"""@echo off
echo Updating Prepaid Tool... Please wait.
timeout /t 2 /nobreak > NUL
move /y "{new_exe_path}" "{exe_name}"
start "" "{exe_name}"
del "%~f0"
"""
            with open(bat_path, "w", encoding="utf-8") as f:
                f.write(bat_content)
                
            subprocess.Popen([bat_path], shell=True)
            self.after(0, self.destroy)
            
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Update Failed", f"Failed to download update: {e}"))
            self.after(0, lambda: self.btn_generate.configure(state="normal", text="  \U0001F4CB  Generate Workpaper"))

    # ─────────────── Header ───────────────
    def _build_header(self):
        header = ctk.CTkFrame(self, fg_color=DARK_THEME, corner_radius=0, height=45)
        header.grid(row=0, column=0, columnspan=2, sticky="ew")
        header.grid_propagate(False)

        ctk.CTkLabel(
            header,
            text="   Prepaid Expense Workpaper Automation",
            font=ctk.CTkFont(family="Segoe UI", size=15, weight="bold"),
            text_color="white",
        ).pack(side="left", padx=10, pady=8)

    # ─────────────── Sidebar ───────────────
    def _build_sidebar(self):
        # Outer wrapper gives a purple right-edge accent line
        sidebar_wrapper = ctk.CTkFrame(self, fg_color=ACCENT_THEME, corner_radius=0, width=224)
        sidebar_wrapper.grid(row=1, column=0, sticky="nsew")
        sidebar_wrapper.grid_propagate(False)

        # Inner white panel (3px right margin exposes the purple accent)
        sidebar = ctk.CTkFrame(sidebar_wrapper, fg_color=CARD_BG, corner_radius=0)
        sidebar.pack(side="left", fill="both", expand=True, padx=(0, 3))

        ctk.CTkLabel(
            sidebar,
            text="Walker Chandiok\n& Co LLP",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=TEXT_DARK,
            justify="left",
        ).pack(anchor="w", padx=22, pady=(22, 28))

        self.nav_buttons = []
        for i, (icon, label) in enumerate(NAV_ITEMS):
            is_active = i == 0
            btn = ctk.CTkButton(
                sidebar,
                text=f" {icon}  {label}",
                font=ctk.CTkFont(family="Segoe UI", size=14),
                fg_color=ACCENT_THEME if is_active else "transparent",
                hover_color=LIGHT_THEME,
                text_color="white" if is_active else ACCENT_THEME,
                anchor="w",
                height=42,
                corner_radius=8,
                command=lambda idx=i: self._nav_click(idx),
            )
            btn.pack(fill="x", padx=14, pady=3)
            self.nav_buttons.append(btn)

        ctk.CTkLabel(
            sidebar,
            text="Walker Chandiok\n& Co LLP",
            font=ctk.CTkFont(family="Segoe UI", size=11),
            text_color=MUTED_TEXT,
            justify="left",
        ).pack(side="bottom", anchor="w", padx=22, pady=22)

    def _nav_click(self, idx):
        for i, btn in enumerate(self.nav_buttons):
            if i == idx:
                btn.configure(fg_color=ACCENT_THEME, text_color="white")
            else:
                btn.configure(fg_color="transparent", text_color=ACCENT_THEME)

    # ─────────────── Main Content ───────────────
    def _build_main(self):
        main = ctk.CTkFrame(self, fg_color=MAIN_BG, corner_radius=0)
        main.grid(row=1, column=1, sticky="nsew")
        main.grid_columnconfigure(0, weight=3)
        main.grid_columnconfigure(1, weight=2)
        main.grid_rowconfigure(1, weight=1)

        # ── Top bar: Client + Status ──
        top_bar = ctk.CTkFrame(main, fg_color=MAIN_BG, height=40)
        top_bar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=25, pady=(15, 5))

        self.client_header = ctk.CTkLabel(
            top_bar,
            text="Client: \u2014",
            font=ctk.CTkFont(family="Segoe UI", size=14, weight="bold"),
            text_color=TEXT_DARK,
        )
        self.client_header.pack(side="left")

        self.status_dot = ctk.CTkLabel(
            top_bar,
            text="\u25CF  Not Started",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=STATUS_COLOR,
        )
        self.status_dot.pack(side="right")

        # ── Left – scrollable form ──
        left = ctk.CTkScrollableFrame(main, fg_color=MAIN_BG, corner_radius=0)
        left.grid(row=1, column=0, sticky="nsew", padx=(25, 10), pady=5)

        # Banner
        banner = ctk.CTkFrame(left, fg_color=BANNER_BG, corner_radius=10, height=50)
        banner.pack(fill="x", pady=(0, 18))
        ctk.CTkLabel(
            banner,
            text="  \U0001F4CB  Automate prepaid expense calculations and generate workpapers efficiently.",
            font=ctk.CTkFont(family="Segoe UI", size=12),
            text_color=BANNER_TEXT,
            wraplength=500,
        ).pack(padx=15, pady=12, anchor="w")

        # ── Select Input File ──
        self._section_label(left, "Select Input File")
        file_row = ctk.CTkFrame(left, fg_color="transparent")
        file_row.pack(fill="x", pady=(0, 12))

        self.entry_file = ctk.CTkEntry(
            file_row, height=38, corner_radius=8,
            border_color=BORDER_CLR, fg_color=CARD_BG,
            placeholder_text="Select an Excel file...",
        )
        self.entry_file.pack(side="left", fill="x", expand=True, padx=(0, 8))

        ctk.CTkButton(
            file_row, text="Browse", width=90, height=38,
            fg_color=ACCENT_THEME, hover_color=MID_THEME,
            corner_radius=8, command=self._browse_file,
        ).pack(side="right")

        # ── Client Name ──
        self._section_label(left, "Client Name")
        self.entry_client = ctk.CTkEntry(
            left, height=38, corner_radius=8,
            border_color=BORDER_CLR, fg_color=CARD_BG,
            placeholder_text="Enter client name",
        )
        self.entry_client.pack(fill="x", pady=(0, 12))
        self.entry_client.bind("<FocusOut>", self._on_client_entered)

        # ── FY Start Date ──
        self._section_label(left, "FY Start Date (DD-MM-YYYY)")
        self.entry_fy_start = ctk.CTkEntry(
            left, height=38, corner_radius=8,
            border_color=BORDER_CLR, fg_color=CARD_BG,
            placeholder_text="DD-MM-YYYY",
        )
        self.entry_fy_start.pack(fill="x", pady=(0, 12))
        self.entry_fy_start.bind("<FocusOut>", self._on_start_entered)

        # ── FY End Date ──
        self._section_label(left, "FY End Date (DD-MM-YYYY)")
        self.entry_fy_end = ctk.CTkEntry(
            left, height=38, corner_radius=8,
            border_color=BORDER_CLR, fg_color=CARD_BG,
            placeholder_text="DD-MM-YYYY",
        )
        self.entry_fy_end.pack(fill="x", pady=(0, 15))
        self.entry_fy_end.bind("<FocusOut>", self._on_end_entered)

        # ── Sampling Card ──
        sampling_card = ctk.CTkFrame(
            left, fg_color=CARD_BG, corner_radius=10,
            border_color=BORDER_CLR, border_width=1,
        )
        sampling_card.pack(fill="x", pady=(0, 15))

        self.sampling_method = ctk.StringVar(value="top")

        ctk.CTkRadioButton(
            sampling_card, text="Top-N Sampling",
            variable=self.sampling_method, value="top",
            font=ctk.CTkFont(size=13),
            fg_color=ACCENT_THEME, hover_color=MID_THEME,
            border_color=BORDER_CLR,
        ).pack(anchor="w", padx=20, pady=(15, 5))

        ctk.CTkFrame(sampling_card, fg_color=BORDER_CLR, height=1).pack(
            fill="x", padx=20, pady=5
        )

        ctk.CTkRadioButton(
            sampling_card, text="Random Sampling",
            variable=self.sampling_method, value="random",
            font=ctk.CTkFont(size=13),
            fg_color=ACCENT_THEME, hover_color=MID_THEME,
            border_color=BORDER_CLR,
        ).pack(anchor="w", padx=20, pady=5)

        ctk.CTkLabel(
            sampling_card, text="Enter the no of values to be picked:",
            font=ctk.CTkFont(size=12), text_color=TEXT_DARK,
        ).pack(anchor="w", padx=20, pady=(10, 4))

        self.entry_top_n = ctk.CTkEntry(
            sampling_card, height=35, width=120, corner_radius=8,
            border_color=BORDER_CLR, fg_color=CARD_BG,
        )
        self.entry_top_n.pack(anchor="w", padx=20, pady=(0, 15))

        # ── Generate Button ──
        self.btn_generate = ctk.CTkButton(
            left,
            text="  \U0001F4CB  Generate Workpaper",
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color=ACCENT_THEME, hover_color=MID_THEME,
            height=45, corner_radius=10,
            command=self._run_process,
        )
        self.btn_generate.pack(anchor="w", pady=(5, 20))

        # ── Right – Statistics Panel ──
        right = ctk.CTkFrame(
            main, fg_color=CARD_BG, corner_radius=12,
            border_color=BORDER_CLR, border_width=1,
        )
        right.grid(row=1, column=1, sticky="nsew", padx=(10, 25), pady=5)

        ctk.CTkLabel(
            right, text="Statistics",
            font=ctk.CTkFont(size=15, weight="bold"),
            text_color=STAT_HEADER,
        ).pack(anchor="w", padx=20, pady=(18, 12))

        stat_items = [
            ("total_rows",     "Total Rows Processed:"),
            ("audit_period",   "Audit Period:"),
            ("addition_count", "No of Total Addition in the\nCurrent Year:"),
            ("addition_sum",   "Amount of Total Addition in the\nCurrent Year:"),
            ("sample_total",   "Sample Amount Total:"),
            ("sample_pct",     "Sample Amount % of Total:"),
            ("status",         "Status:"),
        ]

        self.stat_labels = {}
        for key, label_text in stat_items:
            row_f = ctk.CTkFrame(right, fg_color="transparent")
            row_f.pack(fill="x", padx=20, pady=6)

            ctk.CTkLabel(
                row_f, text=label_text,
                font=ctk.CTkFont(size=12), text_color="#4A4A5A",
                justify="left", wraplength=180,
            ).pack(side="left", anchor="nw")

            val_color = STATUS_COLOR if key == "status" else TEXT_DARK
            val_text  = "Not Started" if key == "status" else "-"
            lbl = ctk.CTkLabel(
                row_f, text=val_text,
                font=ctk.CTkFont(
                    size=12, weight="bold" if key == "status" else "normal"
                ),
                text_color=val_color,
            )
            lbl.pack(side="right", anchor="ne")
            self.stat_labels[key] = lbl

            if key != "status":
                ctk.CTkFrame(right, fg_color=SEPARATOR_CLR, height=1).pack(
                    fill="x", padx=20
                )

        # ── Progress bar (bottom) ──
        self.progress = ctk.CTkProgressBar(
            main, fg_color=LIGHT_THEME,
            progress_color=ACCENT_THEME,
            height=6, corner_radius=3,
        )
        self.progress.grid(
            row=2, column=0, columnspan=2, sticky="ew", padx=25, pady=(5, 5)
        )
        self.progress.set(0)

        self.status_label = ctk.CTkLabel(
            main, text="Waiting for input...",
            font=ctk.CTkFont(size=11), text_color=MUTED_TEXT,
        )
        self.status_label.grid(
            row=3, column=0, columnspan=2, sticky="w", padx=25, pady=(0, 10)
        )

    # ─────────────── UI Helpers ───────────────
    @staticmethod
    def _section_label(parent, text):
        ctk.CTkLabel(
            parent, text=text,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=TEXT_DARK,
        ).pack(anchor="w", pady=(0, 4))

    def _browse_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.entry_file.delete(0, "end")
            self.entry_file.insert(0, path)
            self._update_progress(0.20, "Base file selected")

    def _on_client_entered(self, _event=None):
        name = self.entry_client.get()
        if name:
            self.client_header.configure(text=f"Client: {name}")
            self._update_progress(0.40, "Client name captured")

    def _on_start_entered(self, _event=None):
        if self.entry_fy_start.get():
            self._update_progress(0.60, "FY start date captured")

    def _on_end_entered(self, _event=None):
        if self.entry_fy_end.get():
            self._update_progress(0.80, "Audit period completed")

    def _update_progress(self, value, text):
        current = self.progress.get()
        if value > current or value == 0:
            self.progress.set(value)
            self.status_label.configure(text=text)
            self.update_idletasks()

    def _update_status_dot(self, text, color=STATUS_COLOR):
        self.status_dot.configure(text=f"\u25CF  {text}", text_color=color)

    def _show_statistics(self, total_rows, sample_total, addition_count,
                         addition_sum, addition_pct):
        self.stat_labels["total_rows"].configure(text=str(total_rows))
        fy_s = self.entry_fy_start.get()
        fy_e = self.entry_fy_end.get()
        self.stat_labels["audit_period"].configure(text=f"{fy_s} to {fy_e}")
        self.stat_labels["addition_count"].configure(
            text=str(round(addition_count, 2))
        )
        self.stat_labels["addition_sum"].configure(
            text=str(round(addition_sum, 2))
        )
        self.stat_labels["sample_total"].configure(
            text=str(round(sample_total, 2))
        )
        self.stat_labels["sample_pct"].configure(
            text=f"{round(addition_pct, 2)}%"
        )
        self.stat_labels["status"].configure(
            text="Completed \u2705", text_color=GREEN_OK
        )

    def _processing_complete(self, total_rows, sample_total, addition_count,
                             addition_sum, addition_pct, output_path):
        self._update_progress(1.0, "Processing completed \u2705")
        self._show_statistics(
            total_rows, sample_total, addition_count, addition_sum, addition_pct
        )
        self._update_status_dot("Completed", GREEN_OK)
        self.btn_generate.configure(
            state="normal", text="  \U0001F4CB  Generate Workpaper"
        )
        messagebox.showinfo(
            "Success",
            f"\u2705 Workpaper Generated!\nSaved at:\n{output_path}",
        )

    def _reset_ui(self):
        self.btn_generate.configure(
            state="normal", text="  \U0001F4CB  Generate Workpaper"
        )
        self._update_progress(0, "Waiting for input...")
        self._update_status_dot("Not Started")

    # ─────────────── Thread Launcher ───────────────
    def _run_process(self):
        if self.process_fn is None:
            messagebox.showerror("Error", "No processing function configured.")
            return

        basefile     = self.entry_file.get()
        client_name  = self.entry_client.get()
        fy_start     = self.entry_fy_start.get()
        fy_end       = self.entry_fy_end.get()
        method       = self.sampling_method.get()
        sample_n_str = self.entry_top_n.get()

        # ── Quick validation (UI-side) ──
        if not all([basefile, client_name, fy_start, fy_end]):
            messagebox.showerror("Error", "Please fill all fields")
            return
        if not sample_n_str:
            messagebox.showerror("Error", "Please enter the values for Sampling")
            return
        try:
            sample_n = int(sample_n_str)
        except ValueError:
            messagebox.showerror("Error", "Sample size must be an integer")
            return

        self.btn_generate.configure(state="disabled", text="  \u23F3  Processing...")
        self._update_status_dot("Processing...", "#F59E0B")

        def _thread_target():
            def progress_cb(value, text):
                self.after(0, self._update_progress, value, text)

            try:
                result = self.process_fn(
                    basefile, client_name, fy_start, fy_end,
                    method, sample_n, progress_cb,
                )
                self.after(0, self._processing_complete, *result)
            except Exception as e:
                self.after(
                    0, messagebox.showerror, "Error",
                    f"An error occurred:\n{str(e)}",
                )
                self.after(0, self._reset_ui)

        thread = threading.Thread(target=_thread_target, daemon=True)
        thread.start()
