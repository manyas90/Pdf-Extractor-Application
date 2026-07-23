# ==========================================================
# gui.py
# PDF TO WORD CONVERTER PRO - VERSION 3.1 (Interactive)
# Material Design UI + working conversion logic + extra
# interactivity: accent color cycling, tooltips, drag & drop,
# multi-page preview navigation, toast notifications.
# ==========================================================

import os
import time
import fitz  # PyMuPDF
import threading
import subprocess

import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox

# ----------------------------------------------------------
# Optional drag-and-drop support.
# If tkinterdnd2 isn't installed, the app runs exactly as
# before (no drag & drop) — everything else still works.
# ----------------------------------------------------------
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    HAS_DND = True
except ImportError:
    HAS_DND = False

# ----------------------------------------------------------
# Helper modules — real implementations if present, otherwise
# safe fallbacks so the UI still runs standalone.
# ----------------------------------------------------------
try:
    from converter import PDFConverter
except ImportError:
    class PDFConverter:
        """Fallback converter — replace with your real pdf2docx-based one.
        Implements every method the GUI calls so the app is runnable
        standalone for UI testing even without the real module."""

        def convert(self, pdf_path, output_path, callback=None):
            time.sleep(1.5)
            return True

        def get_pdf_information(self, pdf_path):
            doc = fitz.open(pdf_path)
            info = {
                "file_name": os.path.basename(pdf_path),
                "file_size": os.path.getsize(pdf_path),
                "total_pages": len(doc),
            }
            doc.close()
            return info

        def set_progress_callback(self, callback):
            self._progress_cb = callback

        def convert_exact_layout(self, pdf_path, output_path):
            cb = getattr(self, "_progress_cb", None)
            for pct in (20, 45, 70, 90, 100):
                time.sleep(0.15)
                if cb:
                    cb(pct)
            with open(output_path, "wb"):
                pass
            return True

        def export_text_only(self, pdf_path):
            doc = fitz.open(pdf_path)
            pages = [page.get_text() for page in doc]
            doc.close()
            return pages

        def export_images_only(self, pdf_path, folder):
            return []

        def export_tables_only(self, pdf_path):
            return []

try:
    from pdf_info import PDFInfo
    HAS_PDF_INFO = True
except ImportError:
    HAS_PDF_INFO = False

    class PDFInfo:
        """Fallback: read basic metadata straight from the PDF with PyMuPDF."""
        def __init__(self, file_path):
            self.file_name = os.path.basename(file_path)
            self.file_size = os.path.getsize(file_path)
            doc = fitz.open(file_path)
            self.total_pages = len(doc)
            doc.close()

try:
    from utils import format_size
except ImportError:
    def format_size(size_bytes):
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        return f"{size_bytes / (1024 * 1024):.2f} MB"


# ----------------------------------------------------------
# Material Design 3 Theme System
# ----------------------------------------------------------
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

THEMES = {
    "light": {
        "bg": "#F8F9FA",
        "surface": "#FFFFFF",
        "header_bg": "#FFFFFF",
        "primary": "#1A73E8",
        "primary_hover": "#1557B0",
        "secondary": "#F1F3F4",
        "secondary_hover": "#E8EAED",
        "accent": "#EA4335",
        "accent_hover": "#C5221F",
        "text_main": "#202124",
        "text_sub": "#5F6368",
        "border": "#E0E0E0",
        "status_ready": "#34A853",
        "status_working": "#F9AB00",
        "status_error": "#EA4335",
    },
    "dark": {
        "bg": "#1E1E1E",
        "surface": "#252526",
        "header_bg": "#2D2D2D",
        "primary": "#8AB4F8",
        "primary_hover": "#AECBFA",
        "secondary": "#3C4043",
        "secondary_hover": "#4A4D51",
        "accent": "#F28B82",
        "accent_hover": "#EE675C",
        "text_main": "#E8EAED",
        "text_sub": "#9AA0A6",
        "border": "#3C4043",
        "status_ready": "#81C995",
        "status_working": "#FDD663",
        "status_error": "#F28B82",
    }
}

# ----------------------------------------------------------
# Interactive Accent Color Palettes
# Click the 🎨 swatch button in the header to cycle through
# these — every button in the app updates live, in both
# light and dark mode.
# ----------------------------------------------------------
ACCENT_ORDER = ["blue", "purple", "teal", "orange", "pink"]

ACCENT_PALETTES = {
    "blue": {
        "light": {"primary": "#1A73E8", "primary_hover": "#1557B0"},
        "dark":  {"primary": "#8AB4F8", "primary_hover": "#AECBFA"},
    },
    "purple": {
        "light": {"primary": "#8E24AA", "primary_hover": "#6A1B7A"},
        "dark":  {"primary": "#CE93D8", "primary_hover": "#E1BEE7"},
    },
    "teal": {
        "light": {"primary": "#00897B", "primary_hover": "#00695C"},
        "dark":  {"primary": "#80CBC4", "primary_hover": "#B2DFDB"},
    },
    "orange": {
        "light": {"primary": "#FB8C00", "primary_hover": "#E65100"},
        "dark":  {"primary": "#FFCC80", "primary_hover": "#FFE0B2"},
    },
    "pink": {
        "light": {"primary": "#D81B60", "primary_hover": "#AD1457"},
        "dark":  {"primary": "#F48FB1", "primary_hover": "#F8BBD0"},
    },
}


# ==============================================================
# Small reusable UI helpers: tooltips + toast notifications
# ==============================================================
class ToolTip:
    """Lightweight hover tooltip for any widget (button, label, ...)."""

    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self.show, add="+")
        widget.bind("<Leave>", self.hide, add="+")

    def show(self, event=None):
        if self.tip or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 8
        except Exception:
            return
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        try:
            self.tip.wm_attributes("-topmost", True)
        except Exception:
            pass
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(
            self.tip, text=self.text, background="#202124", foreground="#FFFFFF",
            font=("Segoe UI", 9), padx=8, pady=4, borderwidth=0
        ).pack()

    def hide(self, event=None):
        if self.tip:
            try:
                self.tip.destroy()
            except Exception:
                pass
            self.tip = None


class Toast(ctk.CTkFrame):
    """Small auto-dismissing banner used for non-blocking success
    confirmations, so quick actions (extract images/tables etc.) don't
    force the user to click 'OK' on a dialog every time."""

    def __init__(self, parent, message, color, text_color="#FFFFFF", duration=2600, offset=0):
        super().__init__(parent, corner_radius=10, fg_color=color)
        ctk.CTkLabel(
            self, text=message, text_color=text_color,
            font=ctk.CTkFont(family="Segoe UI", size=12, weight="bold")
        ).pack(padx=18, pady=10)
        self.place(relx=0.985, rely=0.93 - offset, anchor="se")
        self.after(duration, self._dismiss)

    def _dismiss(self):
        try:
            self.destroy()
        except Exception:
            pass


class PDFToWordGUI:

    def __init__(self, root):
        self.root = root

        # -----------------------------------
        # Window Setup
        # -----------------------------------
        self.root.title("PDF To Word Converter Pro v3.1")
        self.root.geometry("1280x820")
        self.root.minsize(1100, 720)

        # -----------------------------------
        # App State
        # -----------------------------------
        self.pdf_path = ""
        self.output_path = ""
        self.dark_mode = False
        self.converter = PDFConverter()
        self.start_time = None
        self.theme = THEMES["light"]
        self.busy = False  # True while converting/extracting

        # Interactive accent color state (see ACCENT_PALETTES above)
        self.accent_index = 0
        self.accent_name = ACCENT_ORDER[self.accent_index]

        # Multi-page live preview state
        self.preview_doc = None
        self.preview_page = 0
        self.preview_page_count = 0

        # Track every widget that needs recoloring on theme/accent change:
        # list of (widget, {option: theme_key})
        self._themed = []
        self._toast_count = 0

        # -----------------------------------
        # Typography
        # -----------------------------------
        self.font_title = ctk.CTkFont(family="Segoe UI", size=18, weight="bold")
        self.font_header = ctk.CTkFont(family="Segoe UI", size=14, weight="bold")
        self.font_bold = ctk.CTkFont(family="Segoe UI", size=12, weight="bold")
        self.font_normal = ctk.CTkFont(family="Segoe UI", size=12)
        self.font_small = ctk.CTkFont(family="Segoe UI", size=11)

        # -----------------------------------
        # Root Container
        # -----------------------------------
        self.container = ctk.CTkFrame(self.root, fg_color=self.theme["bg"], corner_radius=0)
        self.container.pack(fill="both", expand=True)
        self._track(self.container, fg_color="bg")

        self.container.grid_columnconfigure(0, weight=1)
        self.container.grid_rowconfigure(1, weight=1)

        self.build_header()
        self.build_workspace()
        self.build_status_footer()
        self.bind_shortcuts()
        self.setup_drag_and_drop()

    # ==========================================================
    # Theming / interactivity helpers
    # ==========================================================
    def _track(self, widget, **opt_to_key):
        """Register a widget + which of its options map to which theme key,
        so accent/theme changes can recolor everything consistently."""
        self._themed.append((widget, opt_to_key))

    def _add_press_feedback(self, widget):
        """Give a button a quick, tactile 'pressed' look — it dips to its
        hover color the instant the mouse goes down, then springs back to
        its normal fg_color on release, on top of the usual hover effect."""
        try:
            normal = widget.cget("fg_color")
            pressed = widget.cget("hover_color")
        except Exception:
            return
        widget.bind("<ButtonPress-1>", lambda e, w=widget, c=pressed: w.configure(fg_color=c))
        widget.bind("<ButtonRelease-1>", lambda e, w=widget, c=normal: w.configure(fg_color=c))

    def _tip(self, widget, text):
        ToolTip(widget, text)

    def show_toast(self, message, kind="success"):
        color = self.theme["status_ready"] if kind == "success" else self.theme["status_error"]
        offset = self._toast_count * 0.09
        self._toast_count += 1
        toast = Toast(self.container, message, color=color, offset=offset)
        # Free up the stacking slot once this toast disappears.
        self.root.after(2650, lambda: setattr(self, "_toast_count", max(0, self._toast_count - 1)))

    # ==========================================================
    # 1. HEADER BAR
    # ==========================================================
    def build_header(self):
        self.header = ctk.CTkFrame(
            self.container, height=58, corner_radius=0,
            fg_color=self.theme["header_bg"], border_width=1, border_color=self.theme["border"]
        )
        self.header.grid(row=0, column=0, sticky="ew")
        self.header.grid_propagate(False)
        self._track(self.header, fg_color="header_bg", border_color="border")

        self.brand_box = ctk.CTkFrame(self.header, fg_color="transparent")
        self.brand_box.pack(side="left", padx=20, pady=10)

        self.logo_lbl = ctk.CTkLabel(
            self.brand_box, text="📑", font=ctk.CTkFont(size=22), text_color=self.theme["primary"]
        )
        self.logo_lbl.pack(side="left", padx=(0, 10))
        self._track(self.logo_lbl, text_color="primary")

        self.title_lbl = ctk.CTkLabel(
            self.brand_box, text="PDF EXTRACTOR PRO ",
            font=self.font_title, text_color=self.theme["text_main"]
        )
        self.title_lbl.pack(side="left")
        self._track(self.title_lbl, text_color="text_main")

        self.theme_btn = ctk.CTkButton(
            self.header, text="🌙  Dark Mode", width=110, height=34, corner_radius=17,
            fg_color=self.theme["secondary"], hover_color=self.theme["secondary_hover"],
            text_color=self.theme["text_main"], font=self.font_small, command=self.toggle_theme
        )
        self.theme_btn.pack(side="right", padx=(4, 20), pady=12)
        self._track(self.theme_btn, fg_color="secondary", hover_color="secondary_hover", text_color="text_main")
        self._tip(self.theme_btn, "Switch between light and dark appearance")

        # Accent color swatch — click to cycle button colors live.
        self.color_btn = ctk.CTkButton(
            self.header, text="🎨", width=44, height=34, corner_radius=17,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            text_color="#FFFFFF", font=self.font_small, command=self.cycle_accent_color
        )
        self.color_btn.pack(side="right", padx=(0, 4), pady=12)
        self._add_press_feedback(self.color_btn)
        self._tip(self.color_btn, "Cycle the app's accent color")

    # ==========================================================
    # 2. WORKSPACE
    # ==========================================================
    def build_workspace(self):
        self.body = ctk.CTkFrame(self.container, fg_color="transparent")
        self.body.grid(row=1, column=0, sticky="nsew", padx=18, pady=16)

        self.body.grid_columnconfigure(0, weight=4)
        self.body.grid_columnconfigure(1, weight=6)
        self.body.grid_rowconfigure(0, weight=1)
        self.body.grid_rowconfigure(1, weight=1)

        # ---------------- CARD 1: DOCUMENT INPUT ----------------
        self.card_doc = ctk.CTkFrame(
            self.body, fg_color=self.theme["surface"], corner_radius=12,
            border_width=1, border_color=self.theme["border"]
        )
        self.card_doc.grid(row=0, column=0, sticky="nsew", padx=(0, 8), pady=(0, 8))
        self._track(self.card_doc, fg_color="surface", border_color="border")

        lbl = ctk.CTkLabel(self.card_doc, text="📄 DOCUMENT INPUT", font=self.font_header,
                            text_color=self.theme["primary"])
        lbl.pack(anchor="w", padx=16, pady=(14, 8))
        self._track(lbl, text_color="primary")

        self.btn_open_pdf = ctk.CTkButton(
            self.card_doc, text="📂  Open PDF File", height=40, corner_radius=8,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_bold, command=self.select_pdf
        )
        self.btn_open_pdf.pack(fill="x", padx=16, pady=(4, 6))
        self._track(self.btn_open_pdf, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_open_pdf)
        self._tip(self.btn_open_pdf, "Browse for a PDF (Ctrl+O)")

        self.entry_pdf = ctk.CTkEntry(
            self.card_doc, placeholder_text="No document selected...", height=32,
            corner_radius=6, state="disabled", font=self.font_small
        )
        self.entry_pdf.pack(fill="x", padx=16, pady=(0, 4))

        # Drag & drop hint — only shown when tkinterdnd2 is available.
        self.dnd_hint = ctk.CTkLabel(
            self.card_doc,
            text="or drag & drop a PDF here" if HAS_DND else "",
            font=self.font_small, text_color=self.theme["text_sub"]
        )
        self.dnd_hint.pack(anchor="w", padx=16, pady=(0, 8))
        self._track(self.dnd_hint, text_color="text_sub")

        self.btn_output_loc = ctk.CTkButton(
            self.card_doc, text="💾  Set Output Location", height=38, corner_radius=8,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_bold, command=self.select_output
        )
        self.btn_output_loc.pack(fill="x", padx=16, pady=(0, 6))
        self._track(self.btn_output_loc, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_output_loc)
        self._tip(self.btn_output_loc, "Choose where the .docx should be saved (Ctrl+S)")

        self.entry_output = ctk.CTkEntry(
            self.card_doc, placeholder_text="Default: Same as input directory", height=32,
            corner_radius=6, state="disabled", font=self.font_small
        )
        self.entry_output.pack(fill="x", padx=16, pady=(0, 14))

        # ---------------- CARD 2: METADATA ----------------
        self.card_info = ctk.CTkFrame(
            self.body, fg_color=self.theme["surface"], corner_radius=12,
            border_width=1, border_color=self.theme["border"]
        )
        self.card_info.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(8, 0))
        self._track(self.card_info, fg_color="surface", border_color="border")

        lbl2 = ctk.CTkLabel(self.card_info, text="ℹ  FILE METADATA", font=self.font_header,
                             text_color=self.theme["primary"])
        lbl2.pack(anchor="w", padx=16, pady=(14, 8))
        self._track(lbl2, text_color="primary")

        self.info_box = ctk.CTkFrame(self.card_info, fg_color="transparent")
        self.info_box.pack(fill="both", expand=True, padx=16, pady=(0, 14))

        self.lbl_name_val = self.make_info_row("File Name:", "None")
        self.lbl_size_val = self.make_info_row("File Size:", "-")
        self.lbl_pages_val = self.make_info_row("Total Pages:", "-")

        # ---------------- CARD 3: LIVE PREVIEW ----------------
        self.card_preview = ctk.CTkFrame(
            self.body, fg_color=self.theme["surface"], corner_radius=12,
            border_width=1, border_color=self.theme["border"]
        )
        self.card_preview.grid(row=0, column=1, sticky="nsew", padx=(8, 0), pady=(0, 8))
        self._track(self.card_preview, fg_color="surface", border_color="border")

        preview_head = ctk.CTkFrame(self.card_preview, fg_color="transparent")
        preview_head.pack(fill="x", padx=16, pady=(12, 6))

        lbl3 = ctk.CTkLabel(preview_head, text="🖼  LIVE PDF PREVIEW", font=self.font_header,
                             text_color=self.theme["primary"])
        lbl3.pack(side="left")
        self._track(lbl3, text_color="primary")

        self.lbl_page_indicator = ctk.CTkLabel(
            preview_head, text="", font=self.font_small, text_color=self.theme["text_sub"]
        )
        self.lbl_page_indicator.pack(side="right")
        self._track(self.lbl_page_indicator, text_color="text_sub")

        self.preview_canvas = ctk.CTkLabel(
            self.card_preview,
            text="[ First Page Thumbnail Preview ]\n\nLoad a PDF file to render thumbnail",
            font=self.font_normal, text_color=self.theme["text_sub"]
        )
        self.preview_canvas.pack(fill="both", expand=True, padx=16, pady=(0, 8))
        self._track(self.preview_canvas, text_color="text_sub")

        # Prev / Next page navigation for multi-page PDFs.
        preview_nav = ctk.CTkFrame(self.card_preview, fg_color="transparent")
        preview_nav.pack(fill="x", padx=16, pady=(0, 14))

        self.btn_prev_page = ctk.CTkButton(
            preview_nav, text="◀  Prev", width=90, height=32, corner_radius=8,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_small, command=self.show_prev_page, state="disabled"
        )
        self.btn_prev_page.pack(side="left")
        self._track(self.btn_prev_page, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_prev_page)
        self._tip(self.btn_prev_page, "Previous page")

        self.btn_next_page = ctk.CTkButton(
            preview_nav, text="Next  ▶", width=90, height=32, corner_radius=8,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_small, command=self.show_next_page, state="disabled"
        )
        self.btn_next_page.pack(side="right")
        self._track(self.btn_next_page, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_next_page)
        self._tip(self.btn_next_page, "Next page")

        # ---------------- CARD 4: ACTIONS ----------------
        self.card_actions = ctk.CTkFrame(
            self.body, fg_color=self.theme["surface"], corner_radius=12,
            border_width=1, border_color=self.theme["border"]
        )
        self.card_actions.grid(row=1, column=1, sticky="nsew", padx=(8, 0), pady=(8, 0))
        self._track(self.card_actions, fg_color="surface", border_color="border")

        lbl4 = ctk.CTkLabel(self.card_actions, text="⚡  PRIMARY ACTIONS", font=self.font_header,
                             text_color=self.theme["primary"])
        lbl4.pack(anchor="w", padx=16, pady=(14, 8))
        self._track(lbl4, text_color="primary")

        row1 = ctk.CTkFrame(self.card_actions, fg_color="transparent")
        row1.pack(fill="x", padx=16, pady=4)

        self.btn_convert = ctk.CTkButton(
            row1, text="🔄  Convert to Word", height=42, corner_radius=8,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_bold, command=self.start_conversion
        )
        self.btn_convert.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self._track(self.btn_convert, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_convert)
        self._tip(self.btn_convert, "Convert the loaded PDF to a Word document (Ctrl+R)")

        self.btn_extract = ctk.CTkButton(
            row1, text="📝  Extract Text", height=42, corner_radius=8,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_bold, command=self.extract_text
        )
        self.btn_extract.pack(side="left", fill="x", expand=True, padx=4)
        self._track(self.btn_extract, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_extract)
        self._tip(self.btn_extract, "Save just the plain text to a .txt file")

        self.btn_save = ctk.CTkButton(
            row1, text="💾  Save As...", height=42, corner_radius=8,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_bold, command=self.select_output
        )
        self.btn_save.pack(side="left", fill="x", expand=True, padx=(4, 0))
        self._track(self.btn_save, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_save)
        self._tip(self.btn_save, "Pick a custom output location")

        # ==========================
        # Second Row
        # ==========================

        row2 = ctk.CTkFrame(
            self.card_actions,
            fg_color="transparent"
        )
        row2.pack(fill="x", padx=16, pady=(6, 14))

        # ---------------- Open Output ----------------

        self.btn_open = ctk.CTkButton(
            row2,
            text="📖 Open Output",
            height=40,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_bold,
            command=self.open_output_file
        )
        self.btn_open.pack(side="left", expand=True, fill="x", padx=4)
        self._track(self.btn_open, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_open)
        self._tip(self.btn_open, "Open the converted file")

        # ---------------- Open Folder ----------------

        self.btn_folder = ctk.CTkButton(
            row2,
            text="📁 Open Folder",
            height=40,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_bold,
            command=self.open_folder
        )
        self.btn_folder.pack(side="left", expand=True, fill="x", padx=4)
        self._track(self.btn_folder, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_folder)
        self._tip(self.btn_folder, "Open the containing folder")

        # ---------------- Extract Images ----------------

        self.btn_images = ctk.CTkButton(
            row2,
            text="🖼 Extract Images",
            height=40,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_bold,
            command=self.extract_images
        )
        self.btn_images.pack(side="left", expand=True, fill="x", padx=4)
        self._track(self.btn_images, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_images)
        self._tip(self.btn_images, "Pull every embedded image out of the PDF")

        # ---------------- Extract Tables ----------------

        self.btn_tables = ctk.CTkButton(
            row2,
            text="📊 Extract Tables",
            height=40,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_bold,
            command=self.extract_tables
        )
        self.btn_tables.pack(side="left", expand=True, fill="x", padx=4)
        self._track(self.btn_tables, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_tables)
        self._tip(self.btn_tables, "Detect and count tables in the PDF")

        # ---------------- Clear ----------------

        self.btn_clear = ctk.CTkButton(
            row2,
            text="🧹 Clear All",
            height=40,
            fg_color=self.theme["primary"], hover_color=self.theme["primary_hover"],
            font=self.font_bold,
            command=self.clear_all
        )
        self.btn_clear.pack(side="left", expand=True, fill="x", padx=4)
        self._track(self.btn_clear, fg_color="primary", hover_color="primary_hover")
        self._add_press_feedback(self.btn_clear)
        self._tip(self.btn_clear, "Reset everything (Ctrl+L)")

    def make_info_row(self, label_text, default_val):
        row = ctk.CTkFrame(self.info_box, fg_color="transparent")
        row.pack(fill="x", pady=4)

        lab = ctk.CTkLabel(row, text=label_text, font=self.font_bold,
                            text_color=self.theme["text_sub"], width=90, anchor="w")
        lab.pack(side="left")
        self._track(lab, text_color="text_sub")

        val_lbl = ctk.CTkLabel(row, text=default_val, font=self.font_normal,
                                text_color=self.theme["text_main"], anchor="w")
        val_lbl.pack(side="left", fill="x", expand=True)
        self._track(val_lbl, text_color="text_main")
        return val_lbl

    # ==========================================================
    # 3. STATUS & PROGRESS FOOTER
    # ==========================================================
    def build_status_footer(self):
        self.footer = ctk.CTkFrame(
            self.container, height=68, corner_radius=0,
            fg_color=self.theme["header_bg"], border_width=1, border_color=self.theme["border"]
        )
        self.footer.grid(row=2, column=0, sticky="ew")
        self.footer.grid_propagate(False)
        self._track(self.footer, fg_color="header_bg", border_color="border")

        bar_box = ctk.CTkFrame(self.footer, fg_color="transparent")
        bar_box.pack(fill="x", padx=20, pady=(8, 2))

        lbl = ctk.CTkLabel(bar_box, text="Progress:", font=self.font_bold, text_color=self.theme["text_sub"])
        lbl.pack(side="left", padx=(0, 10))
        self._track(lbl, text_color="text_sub")

        self.progress_bar = ctk.CTkProgressBar(
            bar_box, height=10, corner_radius=5, progress_color=self.theme["primary"]
        )
        self.progress_bar.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.progress_bar.set(0)
        self._track(self.progress_bar, progress_color="primary")

        self.lbl_percent = ctk.CTkLabel(bar_box, text="0%", font=self.font_bold,
                                         text_color=self.theme["primary"], width=45)
        self.lbl_percent.pack(side="right")
        self._track(self.lbl_percent, text_color="primary")

        meta_box = ctk.CTkFrame(self.footer, fg_color="transparent")
        meta_box.pack(fill="x", padx=20, pady=(0, 6))

        self.lbl_status = ctk.CTkLabel(meta_box, text="Status: Ready", font=self.font_normal,
                                        text_color=self.theme["status_ready"])
        self.lbl_status.pack(side="left")

        self.lbl_timer = ctk.CTkLabel(meta_box, text="Elapsed Time: 00:00:00", font=self.font_normal,
                                       text_color=self.theme["text_sub"])
        self.lbl_timer.pack(side="right")
        self._track(self.lbl_timer, text_color="text_sub")

    # ==========================================================
    # KEYBOARD SHORTCUTS
    # ==========================================================
    def bind_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self.select_pdf())
        self.root.bind("<Control-s>", lambda e: self.select_output())
        self.root.bind("<Control-r>", lambda e: self.start_conversion())
        self.root.bind("<Control-l>", lambda e: self.clear_all())

    # ==========================================================
    # DRAG & DROP
    # ==========================================================
    def setup_drag_and_drop(self):
        """Let the user drop a PDF straight onto the Document Input card.
        No-op if tkinterdnd2 isn't installed — the rest of the app is
        unaffected either way."""
        if not HAS_DND:
            return
        try:
            for widget in (self.card_doc, self.entry_pdf, self.dnd_hint):
                widget.drop_target_register(DND_FILES)
                widget.dnd_bind("<<Drop>>", self._on_drop_pdf)
        except Exception:
            pass  # DnD not fully supported in this environment — fail quietly

    def _on_drop_pdf(self, event):
        if self.busy:
            return
        raw = event.data.strip()
        # Paths with spaces arrive wrapped in {curly braces}; split safely.
        paths = []
        buf = ""
        in_brace = False
        for ch in raw:
            if ch == "{":
                in_brace = True
                buf = ""
            elif ch == "}":
                in_brace = False
                paths.append(buf)
                buf = ""
            elif ch == " " and not in_brace:
                if buf:
                    paths.append(buf)
                    buf = ""
            else:
                buf += ch
        if buf:
            paths.append(buf)

        pdf_paths = [p for p in paths if p.lower().endswith(".pdf")]
        if not pdf_paths:
            messagebox.showwarning("Warning", "Please drop a .pdf file.")
            return
        self._load_pdf(pdf_paths[0])

    # ==========================================================
    # LOGIC & EVENT HANDLERS (real, working implementations)
    # ==========================================================
    def select_pdf(self):
        if self.busy:
            return
        filename = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
        if not filename:
            return
        self._load_pdf(filename)

    def _load_pdf(self, filename):
        """Shared load path for both the file dialog and drag & drop."""
        self.pdf_path = filename

        self.entry_pdf.configure(state="normal")
        self.entry_pdf.delete(0, "end")
        self.entry_pdf.insert(0, filename)
        self.entry_pdf.configure(state="disabled")

        # Auto-suggest an output path, same as v1
        suggested_out = os.path.splitext(filename)[0] + ".docx"
        self.output_path = suggested_out
        self.entry_output.configure(state="normal")
        self.entry_output.delete(0, "end")
        self.entry_output.insert(0, suggested_out)
        self.entry_output.configure(state="disabled")

        try:
            info = self.converter.get_pdf_information(
                filename
            )

            self.lbl_name_val.configure(
                text=info.get("file_name", os.path.basename(filename))
            )

            self.lbl_size_val.configure(
                text=format_size(
                    info.get("file_size", os.path.getsize(filename))
                )
            )

            self.lbl_pages_val.configure(
                text=str(
                    info.get("total_pages", 0)
                )
            )

            self.lbl_status.configure(text="Status: PDF Loaded Successfully", text_color=self.theme["status_ready"])
        except Exception as e:
            messagebox.showerror("Error", str(e))
            self.lbl_status.configure(text="Status: Failed to Load PDF", text_color=self.theme["status_error"])
            return

        self.render_preview(filename)

    def render_preview(self, pdf_file):
        # Close any previously open preview document first.
        if self.preview_doc is not None:
            try:
                self.preview_doc.close()
            except Exception:
                pass
            self.preview_doc = None

        try:
            doc = fitz.open(pdf_file)
            self.preview_doc = doc
            self.preview_page_count = len(doc)
            self.preview_page = 0
            self._render_current_page()
        except Exception as e:
            self.preview_canvas.configure(
                image=None,
                text=f"Preview Unavailable\n{e}"
            )
            self.lbl_page_indicator.configure(text="")
            self.btn_prev_page.configure(state="disabled")
            self.btn_next_page.configure(state="disabled")

    def _render_current_page(self):
        if self.preview_doc is None:
            return
        try:
            from PIL import Image

            page = self.preview_doc.load_page(self.preview_page)
            pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))

            img = Image.frombytes(
                "RGB",
                [pix.width, pix.height],
                pix.samples
            )
            img.thumbnail((450, 350))

            preview = ctk.CTkImage(
                light_image=img,
                dark_image=img,
                size=img.size
            )

            self.preview_canvas.configure(image=preview, text="")
            self.preview_canvas.image = preview

            self.lbl_page_indicator.configure(
                text=f"Page {self.preview_page + 1} / {self.preview_page_count}"
            )
            self.btn_prev_page.configure(state="normal" if self.preview_page > 0 else "disabled")
            self.btn_next_page.configure(
                state="normal" if self.preview_page < self.preview_page_count - 1 else "disabled"
            )
        except Exception as e:
            self.preview_canvas.configure(image=None, text=f"Preview Unavailable\n{e}")

    def show_prev_page(self):
        if self.preview_doc is None or self.preview_page <= 0:
            return
        self.preview_page -= 1
        self._render_current_page()

    def show_next_page(self):
        if self.preview_doc is None or self.preview_page >= self.preview_page_count - 1:
            return
        self.preview_page += 1
        self._render_current_page()

    def select_output(self):
        if self.busy:
            return
        filename = filedialog.asksaveasfilename(
            title="Save Word File", defaultextension=".docx", filetypes=[("Word Document", "*.docx")]
        )
        if filename:
            self.output_path = filename
            self.entry_output.configure(state="normal")
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, filename)
            self.entry_output.configure(state="disabled")

    # ---------------- Convert ----------------
    def start_conversion(self):
        if self.busy:
            return
        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF document first.")
            return
        if not self.output_path:
            messagebox.showwarning("Warning", "Please choose an output location.")
            return

        self.busy = True
        self.btn_convert.configure(state="disabled")
        self.btn_extract.configure(state="disabled")

        self.progress_bar.set(0)
        self.lbl_percent.configure(text="0%")
        self.lbl_status.configure(text="Status: Converting Document...", text_color=self.theme["status_working"])

        self.start_time = time.time()
        self._animate_progress()
        self.update_timer()

        threading.Thread(target=self._convert_worker, daemon=True).start()

    def _animate_progress(self):
        """Simple indeterminate-style animation while the worker thread runs."""
        if not self.busy:
            return
        current = self.progress_bar.get()
        nxt = 0.08 if current >= 0.9 else current + 0.08
        self.progress_bar.set(nxt)
        self.lbl_percent.configure(text=f"{int(nxt * 100)}%")
        self.root.after(150, self._animate_progress)

    def _progress_callback(self, value):

        self.root.after(
            0,
            lambda: (
                self.progress_bar.set(value / 100),
                self.lbl_percent.configure(text=f"{value}%")
            )
        )


    def _convert_worker(self):

        try:

            self.converter.set_progress_callback(
                self._progress_callback
            )

            result =self.converter.convert_exact_layout(
                self.pdf_path,
                self.output_path
            )

            if result:

                self.root.after(
                    0,
                    self._on_convert_success
                )

            else:

                self.root.after(
                    0,
                    self._on_convert_error,
                    "Conversion Failed"
                )

        except Exception as e:

            self.root.after(
                0,
                self._on_convert_error,
                str(e)
            )

    def _on_convert_success(self):
        self.busy = False
        self.progress_bar.set(1.0)
        self.lbl_percent.configure(text="100%")
        self.lbl_status.configure(text="Status: Completed Successfully!", text_color=self.theme["status_ready"])
        self.btn_convert.configure(state="normal")
        self.btn_extract.configure(state="normal")
        self.show_toast(f"Saved to {os.path.basename(self.output_path)}", kind="success")

    def _on_convert_error(self, error_message):
        self.busy = False
        self.lbl_status.configure(text="Status: Conversion Failed", text_color=self.theme["status_error"])
        self.btn_convert.configure(state="normal")
        self.btn_extract.configure(state="normal")
        messagebox.showerror("Conversion Error", error_message)

    # ---------------- Extract Text ----------------
    def _extract_worker(self, txt_file):

        try:

            text = self.converter.export_text_only(
                self.pdf_path
            )

            with open(
                txt_file,
                "w",
                encoding="utf-8"
            ) as f:

                for paragraph in text:

                    f.write(paragraph)
                    f.write("\n\n")

            self.root.after(
                0,
                self._on_extract_success,
                txt_file
            )

        except Exception as e:

            self.root.after(
                0,
                self._on_extract_error,
                str(e)
            )

    def _on_extract_success(self, txt_file):
        self.busy = False
        self.progress_bar.set(1.0)
        self.lbl_percent.configure(text="100%")
        self.lbl_status.configure(text="Status: Text Extraction Completed", text_color=self.theme["status_ready"])
        self.btn_convert.configure(state="normal")
        self.btn_extract.configure(state="normal")
        self.show_toast(f"Text extracted to {os.path.basename(txt_file)}", kind="success")

    def _on_extract_error(self, error_message):
        self.busy = False
        self.lbl_status.configure(text="Status: Text Extraction Failed", text_color=self.theme["status_error"])
        self.btn_convert.configure(state="normal")
        self.btn_extract.configure(state="normal")
        messagebox.showerror("Error", error_message)

    def extract_images(self):

        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return

        folder = filedialog.askdirectory()

        if not folder:
            return

        images = self.converter.export_images_only(
            self.pdf_path,
            folder
        )

        self.show_toast(f"{len(images)} image(s) extracted", kind="success")

    def extract_tables(self):

        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return

        tables = self.converter.export_tables_only(
            self.pdf_path
        )

        self.show_toast(f"{len(tables)} table(s) found", kind="success")

    def extract_text(self):

        if self.busy:
            return

        if not self.pdf_path:
            messagebox.showwarning("Warning", "Please select a PDF file first.")
            return

        txt_file = filedialog.asksaveasfilename(
            title="Save Text File",
            defaultextension=".txt",
            filetypes=[("Text File", "*.txt")]
        )

        if not txt_file:
            return

        self.busy = True

        self.btn_convert.configure(state="disabled")
        self.btn_extract.configure(state="disabled")

        self.progress_bar.set(0)

        self.lbl_percent.configure(text="0%")

        self.lbl_status.configure(
            text="Status: Extracting Text..."
        )

        self.start_time = time.time()

        self.update_timer()

        threading.Thread(
            target=self._extract_worker,
            args=(txt_file,),
            daemon=True
        ).start()

    # ---------------- Timer ----------------
    def update_timer(self):
        if not self.busy:
            return
        elapsed = int(time.time() - self.start_time)
        hrs, rem = divmod(elapsed, 3600)
        mins, secs = divmod(rem, 60)
        self.lbl_timer.configure(text=f"Elapsed Time: {hrs:02}:{mins:02}:{secs:02}")
        self.root.after(1000, self.update_timer)

    # ---------------- Open / Folder ----------------
    def open_output_file(self):
        if not self.output_path or not os.path.exists(self.output_path):
            messagebox.showwarning("Warning", "No output file available yet.")
            return
        try:
            if os.name == "nt":
                os.startfile(self.output_path)
            else:
                subprocess.Popen(["xdg-open", self.output_path])
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def open_folder(self):
        target = self.output_path or self.pdf_path
        if not target:
            messagebox.showwarning("Warning", "No file/folder available.")
            return
        folder = os.path.dirname(target) or "."
        if not os.path.exists(folder):
            messagebox.showerror("Error", "Folder not found.")
            return
        try:
            if os.name == "nt":
                os.startfile(folder)
            else:
                subprocess.Popen(["xdg-open", folder])
        except Exception as e:
            messagebox.showerror("Error", str(e))

    # ---------------- Clear ----------------
    def clear_all(self):
        if self.busy:
            return

        self.pdf_path = ""
        self.output_path = ""

        self.entry_pdf.configure(state="normal")
        self.entry_pdf.delete(0, "end")
        self.entry_pdf.configure(state="disabled")

        self.entry_output.configure(state="normal")
        self.entry_output.delete(0, "end")
        self.entry_output.configure(state="disabled")

        self.lbl_name_val.configure(text="None")
        self.lbl_size_val.configure(text="-")
        self.lbl_pages_val.configure(text="-")

        # Clear Preview
        if self.preview_doc is not None:
            try:
                self.preview_doc.close()
            except Exception:
                pass
            self.preview_doc = None
        self.preview_page = 0
        self.preview_page_count = 0

        self.preview_canvas.configure(
            image=None,
            text="[ First Page Thumbnail Preview ]\n\nLoad a PDF file to render thumbnail"
        )
        self.preview_canvas.image = None
        self.lbl_page_indicator.configure(text="")
        self.btn_prev_page.configure(state="disabled")
        self.btn_next_page.configure(state="disabled")

        self.progress_bar.set(0)
        self.lbl_percent.configure(text="0%")
        self.lbl_status.configure(text="Status: Ready", text_color=self.theme["status_ready"])
        self.lbl_timer.configure(text="Elapsed Time: 00:00:00")

    # ---------------- Theme ----------------
    def toggle_theme(self):
        self.dark_mode = not self.dark_mode
        mode = "dark" if self.dark_mode else "light"
        ctk.set_appearance_mode(mode)
        self.theme = THEMES[mode]

        for widget, mapping in self._themed:
            opts = {opt: self.theme[key] for opt, key in mapping.items()}
            try:
                widget.configure(**opts)
            except Exception:
                pass  # widget may have been destroyed (e.g. cleared preview image)

        # Status label colors depend on current state, not a fixed theme key —
        # refresh them relative to whatever they're currently showing.
        current_status_text = self.lbl_status.cget("text")
        if "Failed" in current_status_text:
            self.lbl_status.configure(text_color=self.theme["status_error"])
        elif "Converting" in current_status_text or "Extracting" in current_status_text:
            self.lbl_status.configure(text_color=self.theme["status_working"])
        else:
            self.lbl_status.configure(text_color=self.theme["status_ready"])

        self.theme_btn.configure(text="☀️  Light Mode" if self.dark_mode else "🌙  Dark Mode")

        # Re-apply the current accent on top of the new light/dark base
        # theme, so switching Dark Mode doesn't reset your chosen color.
        self._apply_accent(flash=False)

    # ---------------- Interactive Accent Color ----------------
    def cycle_accent_color(self):
        """Cycle to the next accent palette and repaint every button that
        keys off 'primary' / 'primary_hover' — the live, clickable
        button-color changer."""
        self.accent_index = (self.accent_index + 1) % len(ACCENT_ORDER)
        self.accent_name = ACCENT_ORDER[self.accent_index]
        self._apply_accent(flash=True)

    def _apply_accent(self, flash=True):
        mode = "dark" if self.dark_mode else "light"
        accent = ACCENT_PALETTES[self.accent_name][mode]

        # Update the live theme dict so any future widgets built from
        # self.theme also pick up the new accent.
        self.theme["primary"] = accent["primary"]
        self.theme["primary_hover"] = accent["primary_hover"]

        # Repaint every widget registered against "primary"/"primary_hover".
        for widget, mapping in self._themed:
            opts = {opt: self.theme[key] for opt, key in mapping.items()
                     if key in ("primary", "primary_hover")}
            if not opts:
                continue
            try:
                widget.configure(**opts)
            except Exception:
                pass

        # The swatch button previews the new color directly.
        self.color_btn.configure(fg_color=accent["primary"], hover_color=accent["primary_hover"])
        self._add_press_feedback(self.color_btn)  # refresh captured colors

        if flash:
            # Quick flash so the click feels responsive even before the
            # cursor moves off the button (hover_color would otherwise
            # be the only visual cue).
            self.color_btn.configure(fg_color="#FFFFFF")
            self.root.after(90, lambda: self.color_btn.configure(fg_color=accent["primary"]))


if __name__ == "__main__":
    if HAS_DND:
        class DnDCTk(ctk.CTk, TkinterDnD.DnDWrapper):
            """Combines customtkinter's CTk window with tkinterdnd2's
            drag-and-drop support."""
            def __init__(self, *args, **kwargs):
                ctk.CTk.__init__(self, *args, **kwargs)
                self.TkdndVersion = TkinterDnD._require(self)

        app_root = DnDCTk()
    else:
        app_root = ctk.CTk()

    app = PDFToWordGUI(app_root)
    app_root.mainloop()