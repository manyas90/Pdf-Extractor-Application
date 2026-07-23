# ==========================================================
# gui.py
# PDF TO WORD CONVERTER PRO - VERSION 3.0 (Merged)
# Material Design UI (from v2) + working conversion logic (from v1)
# ==========================================================

import os
import time
import fitz  # PyMuPDF
import threading
import subprocess

import customtkinter as ctk
from tkinter import filedialog, messagebox

# ----------------------------------------------------------
# Helper modules — real implementations if present, otherwise
# safe fallbacks so the UI still runs standalone.
# ----------------------------------------------------------
try:
    from converter import PDFConverter
except ImportError:
    class PDFConverter:
        """Fallback converter — replace with your real pdf2docx-based one."""
        def convert(self, pdf_path, output_path, callback=None):
            time.sleep(1.5)
            return True

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


class PDFToWordGUI:

    def __init__(self, root):
        self.root = root

        # -----------------------------------
        # Window Setup
        # -----------------------------------
        self.root.title("PDF To Word Converter Pro v3.0")
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

        # Track every widget that needs recoloring on theme toggle:
        # list of (widget, {option: theme_key})
        self._themed = []

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

    # ==========================================================
    # Theming helper
    # ==========================================================
    def _track(self, widget, **opt_to_key):
        """Register a widget + which of its options map to which theme key,
        so toggle_theme can recolor everything consistently."""
        self._themed.append((widget, opt_to_key))

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
            self.brand_box, text="PDF EXTRACTOR PRO  |  Dashboard 3.0",
            font=self.font_title, text_color=self.theme["text_main"]
        )
        self.title_lbl.pack(side="left")
        self._track(self.title_lbl, text_color="text_main")

        self.theme_btn = ctk.CTkButton(
            self.header, text="🌙  Dark Mode", width=110, height=34, corner_radius=17,
            fg_color=self.theme["secondary"], hover_color=self.theme["secondary_hover"],
            text_color=self.theme["text_main"], font=self.font_small, command=self.toggle_theme
        )
        self.theme_btn.pack(side="right", padx=20, pady=12)
        self._track(self.theme_btn, fg_color="secondary", hover_color="secondary_hover", text_color="text_main")

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

        self.entry_pdf = ctk.CTkEntry(
            self.card_doc, placeholder_text="No document selected...", height=32,
            corner_radius=6, state="disabled", font=self.font_small
        )
        self.entry_pdf.pack(fill="x", padx=16, pady=(0, 12))

        self.btn_output_loc = ctk.CTkButton(
            self.card_doc, text="💾  Set Output Location", height=38, corner_radius=8,
            fg_color=self.theme["secondary"], hover_color=self.theme["secondary_hover"],
            text_color=self.theme["text_main"], font=self.font_bold, command=self.select_output
        )
        self.btn_output_loc.pack(fill="x", padx=16, pady=(0, 6))
        self._track(self.btn_output_loc, fg_color="secondary", hover_color="secondary_hover", text_color="text_main")

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

        lbl3 = ctk.CTkLabel(self.card_preview, text="🖼  LIVE PDF PREVIEW", font=self.font_header,
                             text_color=self.theme["primary"])
        lbl3.pack(anchor="w", padx=16, pady=(12, 6))
        self._track(lbl3, text_color="primary")

        self.preview_canvas = ctk.CTkLabel(
            self.card_preview,
            text="[ First Page Thumbnail Preview ]\n\nLoad a PDF file to render thumbnail",
            font=self.font_normal, text_color=self.theme["text_sub"]
        )
        self.preview_canvas.pack(fill="both", expand=True, padx=16, pady=(0, 14))
        self._track(self.preview_canvas, text_color="text_sub")

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

        self.btn_extract = ctk.CTkButton(
            row1, text="📝  Extract Text", height=42, corner_radius=8,
            fg_color=self.theme["secondary"], hover_color=self.theme["secondary_hover"],
            text_color=self.theme["text_main"], font=self.font_bold, command=self.extract_text
        )
        self.btn_extract.pack(side="left", fill="x", expand=True, padx=4)
        self._track(self.btn_extract, fg_color="secondary", hover_color="secondary_hover", text_color="text_main")

        self.btn_save = ctk.CTkButton(
            row1, text="💾  Save As...", height=42, corner_radius=8,
            fg_color=self.theme["secondary"], hover_color=self.theme["secondary_hover"],
            text_color=self.theme["text_main"], font=self.font_bold, command=self.select_output
        )
        # ==========================
        # Second Row
        # ==========================

        row2 = ctk.CTkFrame(
            self.card_actions,
            fg_color="transparent"
        )
        row2.pack(fill="x", padx=16, pady=(6,14))

        # ---------------- Open Output ----------------

        self.btn_open = ctk.CTkButton(
            row2,
            text="📖 Open Output",
            height=40,
            command=self.open_output_file
        )
        self.btn_open.pack(side="left", expand=True, fill="x", padx=4)

        # ---------------- Open Folder ----------------

        self.btn_folder = ctk.CTkButton(
            row2,
            text="📁 Open Folder",
            height=40,
            command=self.open_folder
        )
        self.btn_folder.pack(side="left", expand=True, fill="x", padx=4)

        # ---------------- Extract Images ----------------

        self.btn_images = ctk.CTkButton(
            row2,
            text="🖼 Extract Images",
            height=40,
            command=self.extract_images
        )
        self.btn_images.pack(side="left", expand=True, fill="x", padx=4)

        # ---------------- Extract Tables ----------------

        self.btn_tables = ctk.CTkButton(
            row2,
            text="📊 Extract Tables",
            height=40,
            command=self.extract_tables
        )
        self.btn_tables.pack(side="left", expand=True, fill="x", padx=4)

        # ---------------- Clear ----------------

        self.btn_clear = ctk.CTkButton(
            row2,
            text="🧹 Clear All",
            height=40,
            command=self.clear_all
        )
        self.btn_clear.pack(side="left", expand=True, fill="x", padx=4)

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
    # LOGIC & EVENT HANDLERS (real, working implementations)
    # ==========================================================
    def select_pdf(self):
        if self.busy:
            return
        filename = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF Files", "*.pdf")])
        if not filename:
            return

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
        try:
            from PIL import Image

            doc = fitz.open(pdf_file)
            page = doc.load_page(0)

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

            self.preview_canvas.configure(
                image=preview,
                text=""
            )

            self.preview_canvas.image = preview

            doc.close()

        except Exception as e:
            self.preview_canvas.configure(
                image=None,
                text=f"Preview Unavailable\n{e}"
            )
            
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
        messagebox.showinfo("Success", f"File saved to:\n{self.output_path}")

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
        messagebox.showinfo("Success", f"Text extracted to:\n{txt_file}")

    def _on_extract_error(self, error_message):
        self.busy = False
        self.lbl_status.configure(text="Status: Text Extraction Failed", text_color=self.theme["status_error"])
        self.btn_convert.configure(state="normal")
        self.btn_extract.configure(state="normal")
        messagebox.showerror("Error", error_message)

    def extract_images(self):

        if not self.pdf_path:
            return

        folder = filedialog.askdirectory()

        if not folder:
            return

        images = self.converter.export_images_only(
            self.pdf_path,
            folder
        )

        messagebox.showinfo(
            "Success",
            f"{len(images)} Images Extracted"
        )
        
    def extract_tables(self):

        if not self.pdf_path:
            return

        tables = self.converter.export_tables_only(
            self.pdf_path
        )

        messagebox.showinfo(
            "Tables",
            f"{len(tables)} Tables Found"
        )
    
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
        self.preview_canvas.configure(
            image=None,
            text="[ First Page Thumbnail Preview ]\n\nLoad a PDF file to render thumbnail"
        )

        # Remove image reference
        self.preview_canvas.image = None
            
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


if __name__ == "__main__":
    app_root = ctk.CTk()
    app = PDFToWordGUI(app_root)
    app_root.mainloop()