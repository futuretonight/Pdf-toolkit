"""
pdf toolkit V3  ·  PC Edition
────────────────────────────────────────────────────
Full-featured document converter with compression, PDF tools, OCR, and more.
Supports: CSV, XLSX, DOCX, PDF, Images, TXT

Requirements (pip install):
    pandas pillow reportlab openpyxl python-docx pdf2image pypdf psutil
    pdfplumber pikepdf docx2pdf tkinterdnd2
    pymupdf pypdfium2 numpy opencv-python
"""

import os
import sys
import threading
import subprocess
import shutil
from datetime import datetime
import tempfile
import traceback
import time
import hashlib
import sqlite3
import struct
import concurrent.futures
from pathlib import Path

# ── Tkinter ──────────────────────────────────────────────────────────────────
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ImportError:
    print("ERROR: Tkinter not available. Install python3-tk")
    sys.exit(1)

# ── Optional heavy deps ───────────────────────────────────────────────────────
try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False
    A4 = (595.27, 841.89)

try:
    from PIL import Image, ImageDraw, ImageFont
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import docx
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

try:
    from pdf2image import convert_from_path
    PDF2IMAGE_AVAILABLE = True
except ImportError:
    PDF2IMAGE_AVAILABLE = False

try:
    from docx2pdf import convert as docx2pdf_convert
    DOCX2PDF_AVAILABLE = True
except ImportError:
    DOCX2PDF_AVAILABLE = False

try:
    import pypdf
    PYPDF_AVAILABLE = True
except ImportError:
    PYPDF_AVAILABLE = False

try:
    import pikepdf
    PIKEPDF_AVAILABLE = True
except ImportError:
    PIKEPDF_AVAILABLE = False

try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    TkBase = TkinterDnD.Tk
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    TkBase = tk.Tk
    DND_FILES = None

# ── PyMuPDF — primary renderer (pip install pymupdf) ─────────────────────────
# Try the modern top-level name first (pymupdf >= 1.24).  The old alias
# "import fitz" can accidentally resolve to an unrelated "frontend" package
# that also registers a "fitz" module, causing a RuntimeError on import.
try:
    import pymupdf as fitz            # pymupdf >= 1.24  (preferred)
    PYMUPDF_AVAILABLE = True
except ImportError:
    try:
        # Older pymupdf installs (< 1.24) only expose the fitz name.
        # Guard against the frontend-package collision by verifying fitz.open exists.
        import fitz as _fitz_test
        if not callable(getattr(_fitz_test, "open", None)):
            raise ImportError("fitz resolved to wrong package (frontend collision)")
        fitz = _fitz_test
        PYMUPDF_AVAILABLE = True
    except (ImportError, RuntimeError):
        PYMUPDF_AVAILABLE = False

# ── pypdfium2 — secondary renderer (pip install pypdfium2) ───────────────────
try:
    import pypdfium2 as pdfium
    PYPDFIUM2_AVAILABLE = True
except ImportError:
    PYPDFIUM2_AVAILABLE = False

# ── NumPy ─────────────────────────────────────────────────────────────────────
try:
    import numpy as np
    NUMPY_AVAILABLE = True
except ImportError:
    NUMPY_AVAILABLE = False
    np = None

# ── OpenCV (pip install opencv-python) ───────────────────────────────────────
try:
    import cv2
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False
    cv2 = None


# =====================================================================
#  DARK THEME PALETTE & STYLE
# =====================================================================

DARK = {
    "bg":        "#12121f",
    "panel":     "#1c1c2e",
    "surface":   "#252538",
    "border":    "#3a3a5c",
    "accent":    "#7c5cfc",
    "accent2":   "#00d4ff",
    "text":      "#e2e8f0",
    "text_dim":  "#8892a4",
    "success":   "#10b981",
    "warning":   "#f59e0b",
    "danger":    "#ef4444",
    "log_bg":    "#0d0d18",
    "log_fg":    "#39ff14",
    "btn":       "#2e2e50",
    "btn_hover": "#3d3d70",
    "entry":     "#1a1a2e",
    "select":    "#7c5cfc",
}


def apply_dark_theme(root):
    """Configure ttk dark theme on top of 'clam'."""
    style = ttk.Style(root)
    style.theme_use("clam")

    bg   = DARK["bg"]
    pan  = DARK["panel"]
    surf = DARK["surface"]
    bdr  = DARK["border"]
    acc  = DARK["accent"]
    txt  = DARK["text"]
    dim  = DARK["text_dim"]
    btn  = DARK["btn"]
    ent  = DARK["entry"]

    # Root window
    root.configure(bg=bg)

    # Frame / LabelFrame
    style.configure("TFrame",       background=bg)
    style.configure("TLabelframe",  background=pan, bordercolor=bdr,
                    relief="flat", padding=4)
    style.configure("TLabelframe.Label", background=pan, foreground=acc,
                    font=("Consolas", 9, "bold"))

    # Notebook tabs
    style.configure("TNotebook",    background=bg,  borderwidth=0)
    style.configure("TNotebook.Tab", background=surf, foreground=dim,
                    font=("Consolas", 9, "bold"), padding=(12, 5),
                    borderwidth=0)
    style.map("TNotebook.Tab",
              background=[("selected", pan), ("active", btn)],
              foreground=[("selected", acc), ("active", txt)],
              expand=[("selected", [1, 1, 1, 0])])

    # Labels
    style.configure("TLabel", background=bg, foreground=txt,
                    font=("Segoe UI", 9))
    style.configure("Dim.TLabel", background=bg, foreground=dim,
                    font=("Segoe UI", 8))
    style.configure("Title.TLabel", background=bg, foreground=acc,
                    font=("Segoe UI", 14, "bold"))
    style.configure("Section.TLabel", background=pan, foreground=acc,
                    font=("Segoe UI", 10, "bold"))

    # Buttons
    style.configure("TButton", background=btn, foreground=txt,
                    font=("Segoe UI", 9), borderwidth=0,
                    relief="flat", padding=(10, 5))
    style.map("TButton",
              background=[("active", DARK["btn_hover"]), ("pressed", acc)],
              foreground=[("active", txt)])

    style.configure("Accent.TButton", background=acc, foreground="#ffffff",
                    font=("Segoe UI", 9, "bold"), padding=(14, 6))
    style.map("Accent.TButton",
              background=[("active", "#6b4de8"), ("pressed", "#5a3dd8")])

    style.configure("Danger.TButton", background=DARK["danger"],
                    foreground="#ffffff", font=("Segoe UI", 9, "bold"),
                    padding=(10, 5))

    # Entry / Combobox
    style.configure("TEntry", fieldbackground=ent, foreground=txt,
                    insertcolor=txt, bordercolor=bdr, lightcolor=bdr,
                    darkcolor=bdr)
    style.configure("TCombobox", fieldbackground=ent, foreground=txt,
                    selectbackground=DARK["select"], selectforeground=txt,
                    background=btn, bordercolor=bdr)
    style.map("TCombobox",
              fieldbackground=[("readonly", ent)],
              foreground=[("readonly", txt)])

    # Scrollbar
    style.configure("TScrollbar", background=surf, troughcolor=bg,
                    bordercolor=bg, arrowcolor=dim)

    # Progressbar
    style.configure("TProgressbar", troughcolor=surf, background=acc,
                    borderwidth=0)

    # Separator
    style.configure("TSeparator", background=bdr)

    # Listbox (tk, not ttk — set directly when creating)
    LB_OPTS = dict(bg=surf, fg=txt, selectbackground=acc,
                   selectforeground="#fff", bd=0, highlightthickness=0,
                   activestyle="none", font=("Consolas", 9))

    return LB_OPTS   # caller can unpack when building Listboxes


# =====================================================================
#  UTILITY FUNCTIONS
# =====================================================================

def now():
    return datetime.now().strftime("%H:%M:%S")


def log_append(log_widget, text, tag="normal"):
    if log_widget is None:
        return
    if threading.current_thread() is not threading.main_thread():
        try:
            log_widget.after(0, lambda: log_append(log_widget, text, tag))
        except Exception:
            pass
        return
    try:
        log_widget.configure(state="normal")
        ts = f"[{now()}] "
        log_widget.insert(tk.END, ts, "ts")
        log_widget.insert(tk.END, text + "\n", tag)
        log_widget.yview_moveto(1.0)
        log_widget.configure(state="disabled")
    except Exception:
        pass


def safe_filename(path):
    base = os.path.basename(path)
    name, _ = os.path.splitext(base)
    return name


def ensure_dir(path):
    directory = os.path.dirname(path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory, exist_ok=True)


def detect_file_type(path):
    ext = os.path.splitext(path)[1].lower().strip(".")
    return ext if ext else "unknown"


def open_path(path):
    """Open file or folder in the OS native explorer/finder."""
    target = path if os.path.isdir(path) else os.path.dirname(path)
    if not target or not os.path.exists(target):
        return
    if sys.platform.startswith("win"):
        if os.path.isfile(path):
            subprocess.Popen(["explorer", "/select,", path])
        else:
            os.startfile(target)
    elif sys.platform.startswith("darwin"):
        subprocess.Popen(["open", "-R", path] if os.path.isfile(path) else ["open", target])
    else:
        subprocess.Popen(["xdg-open", target])


def human_size(nbytes):
    for unit in ("B", "KB", "MB", "GB"):
        if nbytes < 1024:
            return f"{nbytes:.1f} {unit}"
        nbytes /= 1024
    return f"{nbytes:.1f} TB"


def human_time(seconds):
    seconds = max(0, int(round(seconds)))
    if seconds < 60:
        return f"{seconds}s"
    minutes, seconds = divmod(seconds, 60)
    if minutes < 60:
        return f"{minutes}m {seconds:02d}s"
    hours, minutes = divmod(minutes, 60)
    return f"{hours}h {minutes:02d}m"


def total_existing_size(paths):
    total = 0
    for path in paths:
        try:
            if path and os.path.exists(path):
                total += os.path.getsize(path)
        except OSError:
            pass
    return total


def output_dir_or_source_dir(selected_dir, source_path):
    selected_dir = (selected_dir or "").strip()
    if selected_dir:
        os.makedirs(selected_dir, exist_ok=True)
        return selected_dir
    return os.path.dirname(source_path)


def pdf_output_path(src, suffix, out_dir=None):
    folder = output_dir_or_source_dir(out_dir, src)
    return os.path.join(folder, f"{safe_filename(src)}{suffix}.pdf")


def estimate_compressed_size(total_bytes, engine, quality):
    """Return a rough output-size estimate for the selected compression engine."""
    if total_bytes <= 0:
        return 0
    if "pikepdf" in engine:
        factor = 0.90
    else:
        factors = {
            "screen": 0.35,
            "ebook": 0.55,
            "printer": 0.75,
            "prepress": 0.90,
        }
        factor = factors.get(quality, 0.55)
    return int(total_bytes * factor)


def estimate_job_time(total_bytes, mb_per_second):
    if total_bytes <= 0:
        return 1
    return max(1, total_bytes / (1024 * 1024) / mb_per_second)


def size_delta(before, after):
    if before <= 0:
        return "n/a"
    saved = (1 - after / before) * 100
    return f"{saved:.1f}% saved" if saved >= 0 else f"{abs(saved):.1f}% larger"


# =====================================================================
#  DATAFRAME → IMAGE / PDF  (Pillow textbbox fix)
# =====================================================================

def _resolve_font(size=14):
    """Load a decent font, falling back to default."""
    candidates = [
        "DejaVuSans.ttf",
        "arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
        "C:/Windows/Fonts/arial.ttf",
    ]
    for path in candidates:
        try:
            return ImageFont.truetype(path, size)
        except Exception:
            pass
    return ImageFont.load_default()


def df_to_image(df, out_path, log_widget=None, max_width=5120):
    if not PIL_AVAILABLE:
        raise RuntimeError("Pillow not installed.")
    if pd is None:
        raise RuntimeError("Pandas not installed.")

    font = _resolve_font(14)
    pad = 12
    dummy = Image.new("RGB", (10, 10))
    draw = ImageDraw.Draw(dummy)

    columns = list(df.columns)
    rows = df.astype(str).values.tolist()

    def text_w(text):
        bbox = draw.textbbox((0, 0), str(text), font=font)
        return bbox[2] - bbox[0]

    def text_h(text):
        bbox = draw.textbbox((0, 0), str(text), font=font)
        return bbox[3] - bbox[1]

    col_widths = []
    for i, col in enumerate(columns):
        w = text_w(str(col)) + pad * 2
        for row in rows:
            w = max(w, text_w(str(row[i])) + pad * 2)
        col_widths.append(min(w, 320))

    row_h = text_h("Ag") + pad
    hdr_h = row_h + 8
    total_w = sum(col_widths)
    total_h = hdr_h + row_h * len(rows) + 10

    if total_w > max_width:
        scale = max_width / total_w
        col_widths = [max(int(w * scale), 20) for w in col_widths]
        total_w = sum(col_widths)

    # Dark-themed table image
    img = Image.new("RGB", (total_w, total_h), DARK["panel"])
    draw = ImageDraw.Draw(img)

    x = 0
    for i, col in enumerate(columns):
        w = col_widths[i]
        draw.rectangle([x, 0, x + w - 1, hdr_h], fill="#2e2e50", outline=DARK["border"])
        draw.text((x + pad, 6), str(col)[:40], font=font, fill=DARK["accent"])
        x += w

    y = hdr_h
    for ri, row in enumerate(rows):
        fill = DARK["surface"] if ri % 2 == 0 else "#202035"
        x = 0
        for i, cell in enumerate(row):
            w = col_widths[i]
            draw.rectangle([x, y, x + w - 1, y + row_h - 1], fill=fill, outline=DARK["border"])
            draw.text((x + pad, y + pad // 2), str(cell)[:50], font=font, fill=DARK["text"])
            x += w
        y += row_h

    ensure_dir(out_path)
    img.save(out_path)
    if log_widget:
        log_append(log_widget, f"Image saved: {out_path}", "success")


def df_to_pdf(df, out_path, log_widget=None):
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ReportLab not installed.")
    if pd is None:
        raise RuntimeError("Pandas not installed.")

    data = [list(df.columns)] + [[str(x) for x in row] for row in df.itertuples(index=False)]
    ensure_dir(out_path)
    doc = SimpleDocTemplate(out_path, pagesize=A4)
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#2e2e50")),
        ("TEXTCOLOR",  (0, 0), (-1, 0), colors.HexColor("#7c5cfc")),
        ("BACKGROUND", (0, 1), (-1, -1), colors.white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5ff")]),
        ("GRID",       (0, 0), (-1, -1), 0.4, colors.HexColor("#ccccdd")),
        ("ALIGN",      (0, 0), (-1, -1), "LEFT"),
        ("FONTSIZE",   (0, 0), (-1, -1), 9),
        ("VALIGN",     (0, 0), (-1, -1), "TOP"),
        ("FONTNAME",   (0, 0), (-1, 0), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
    ]))
    doc.build([table])
    if log_widget:
        log_append(log_widget, f"PDF saved: {out_path}", "success")


# =====================================================================
#  FORMAT CONVERTERS
# =====================================================================

def csv_to_pdf(src, out, log_widget=None):
    if pd is None: raise RuntimeError("Pandas required")
    df_to_pdf(pd.read_csv(src), out, log_widget)

def csv_to_image(src, out, log_widget=None):
    if pd is None: raise RuntimeError("Pandas required")
    df_to_image(pd.read_csv(src), out, log_widget)

def csv_to_xlsx(src, out, log_widget=None):
    if pd is None: raise RuntimeError("Pandas required")
    df = pd.read_csv(src)
    ensure_dir(out)
    df.to_excel(out, index=False)
    if log_widget: log_append(log_widget, f"XLSX saved: {out}", "success")

def xlsx_to_pdf(src, out, log_widget=None):
    if pd is None: raise RuntimeError("Pandas required")
    df_to_pdf(pd.read_excel(src), out, log_widget)

def xlsx_to_image(src, out, log_widget=None):
    if pd is None: raise RuntimeError("Pandas required")
    df_to_image(pd.read_excel(src), out, log_widget)

def xlsx_to_csv(src, out, log_widget=None):
    if pd is None: raise RuntimeError("Pandas required")
    df = pd.read_excel(src)
    ensure_dir(out)
    df.to_csv(out, index=False)
    if log_widget: log_append(log_widget, f"CSV saved: {out}", "success")

def txt_to_pdf(src, out, log_widget=None):
    if not REPORTLAB_AVAILABLE: raise RuntimeError("ReportLab required")
    with open(src, "r", encoding="utf-8", errors="ignore") as f:
        text = f.read()
    ensure_dir(out)
    doc = SimpleDocTemplate(out, pagesize=A4)
    styles = getSampleStyleSheet()
    para = Paragraph(text.replace("\n", "<br/>"), styles["Normal"])
    doc.build([para])
    if log_widget: log_append(log_widget, f"PDF saved: {out}", "success")

def image_to_pdf(src, out, log_widget=None):
    if not PIL_AVAILABLE: raise RuntimeError("Pillow required")
    img = Image.open(src)
    if img.mode == "RGBA":
        img = img.convert("RGB")
    ensure_dir(out)
    img.save(out, "PDF")
    if log_widget: log_append(log_widget, f"PDF saved: {out}", "success")

def pdf_to_images(src, out_dir, dpi=200, log_widget=None):
    if not PDF2IMAGE_AVAILABLE:
        raise RuntimeError("pdf2image not installed. Also needs: poppler")
    os.makedirs(out_dir, exist_ok=True)
    pages = convert_from_path(src, dpi=dpi)
    outfiles = []
    for i, page in enumerate(pages, 1):
        path = os.path.join(out_dir, f"{safe_filename(src)}_page_{i}.png")
        page.save(path, "PNG")
        outfiles.append(path)
        if log_widget: log_append(log_widget, f"  Extracted page {i}/{len(pages)}")
    return outfiles

def docx_to_pdf(src, out, log_widget=None):
    if DOCX2PDF_AVAILABLE:
        try:
            tmp = tempfile.mkdtemp()
            docx2pdf_convert(src, tmp)
            produced = os.path.join(tmp, safe_filename(src) + ".pdf")
            if os.path.exists(produced):
                ensure_dir(out)
                shutil.move(produced, out)
                shutil.rmtree(tmp, ignore_errors=True)
                if log_widget: log_append(log_widget, f"PDF saved: {out}", "success")
                return
        except Exception as e:
            if log_widget: log_append(log_widget, f"docx2pdf failed: {e}", "warn")

    if shutil.which("soffice"):
        try:
            tmp = tempfile.mkdtemp()
            subprocess.run(["soffice", "--headless", "--convert-to", "pdf",
                            "--outdir", tmp, src], check=True, capture_output=True)
            produced = os.path.join(tmp, safe_filename(src) + ".pdf")
            if os.path.exists(produced):
                ensure_dir(out)
                shutil.move(produced, out)
                shutil.rmtree(tmp, ignore_errors=True)
                if log_widget: log_append(log_widget, f"PDF saved via LibreOffice: {out}", "success")
                return
        except Exception as e:
            if log_widget: log_append(log_widget, f"LibreOffice failed: {e}", "warn")

    # Text-only fallback
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx not installed")
    document = docx.Document(src)
    text = "\n".join(p.text for p in document.paragraphs)
    tmp_txt = os.path.join(tempfile.gettempdir(), "_tmp_docx.txt")
    with open(tmp_txt, "w", encoding="utf-8") as f:
        f.write(text)
    txt_to_pdf(tmp_txt, out, log_widget)


def docx_to_images(src, out_dir, log_widget=None):
    tmp_pdf = os.path.join(tempfile.gettempdir(), safe_filename(src) + "_tmp.pdf")
    docx_to_pdf(src, tmp_pdf, log_widget)
    return pdf_to_images(tmp_pdf, out_dir, log_widget=log_widget)


# =====================================================================
#  PDF-SPECIFIC PC TOOLS
# =====================================================================

def compress_pdf_pikepdf(src, out, log_widget=None):
    """Lossless PDF compression using pikepdf (no Ghostscript needed)."""
    if not PIKEPDF_AVAILABLE:
        raise RuntimeError("pikepdf not installed: pip install pikepdf")
    with pikepdf.open(src) as pdf:
        pdf.save(out, compress_streams=True, object_stream_mode=pikepdf.ObjectStreamMode.generate)
    before = os.path.getsize(src)
    after  = os.path.getsize(out)
    ratio  = (1 - after / before) * 100 if before else 0
    if log_widget:
        log_append(log_widget,
                   f"Compressed: {human_size(before)} → {human_size(after)}  ({ratio:.1f}% saved)",
                   "success")


def _find_ghostscript():
    """Return the Ghostscript executable name, or None if not found.

    On Windows (32-bit or 64-bit) sys.platform == 'win32', so we try
    the console variants in order of preference.
    """
    if sys.platform == "win32":
        candidates = ["gswin64c", "gswin32c", "gs"]
    else:
        candidates = ["gs"]
    for name in candidates:
        if shutil.which(name):
            return name
    return None


def compress_pdf_ghostscript(src, out, quality="ebook", log_widget=None, dpi=None):
    """High-ratio compression using Ghostscript (lossy, much smaller).

    quality  : one of screen / ebook / printer / prepress  (ignored when dpi is given)
    dpi      : integer 30-600; when provided, overrides quality with explicit image resolution
    """
    gs_map = {"screen": "/screen", "ebook": "/ebook", "printer": "/printer", "prepress": "/prepress"}
    gs_q   = gs_map.get(quality, "/ebook")
    gs_bin = _find_ghostscript()
    if not gs_bin:
        raise RuntimeError(
            "Ghostscript not found (tried gswin64c, gswin32c, gs). "
            "Install from https://www.ghostscript.com/releases/ and make sure it's on PATH.")
    cmd = [gs_bin, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.5",
           f"-dPDFSETTINGS={gs_q}", "-dNOPAUSE", "-dQUIET", "-dBATCH"]
    if dpi is not None:
        # explicit DPI overrides the preset PDFSETTINGS image resolution
        dpi = max(30, min(600, int(dpi)))
        cmd += [
            f"-dColorImageResolution={dpi}",
            f"-dGrayImageResolution={dpi}",
            f"-dMonoImageResolution={min(dpi * 2, 1200)}",
        ]
    cmd += [f"-sOutputFile={out}", src]
    subprocess.run(cmd, check=True)
    if log_widget:
        before = os.path.getsize(src)
        after  = os.path.getsize(out)
        ratio  = (1 - after / before) * 100 if before else 0
        log_append(log_widget,
                   f"GS compressed: {human_size(before)} → {human_size(after)}  ({ratio:.1f}% saved)",
                   "success")


def split_pdf(src, out_dir, mode="pages", ranges=None, log_widget=None):
    """Split a PDF.

    mode='pages'  → one file per page
    mode='range'  → extract pages in ranges like '1-3,5,7-9'
    """
    if not PYPDF_AVAILABLE:
        raise RuntimeError("pypdf not installed")
    os.makedirs(out_dir, exist_ok=True)
    reader = pypdf.PdfReader(src)
    total  = len(reader.pages)
    base   = safe_filename(src)
    out_files = []

    if mode == "pages":
        for i, page in enumerate(reader.pages, 1):
            writer = pypdf.PdfWriter()
            writer.add_page(page)
            out = os.path.join(out_dir, f"{base}_page_{i:03d}.pdf")
            with open(out, "wb") as f:
                writer.write(f)
            out_files.append(out)
            if log_widget: log_append(log_widget, f"  Wrote page {i}/{total}")
    elif mode == "range" and ranges:
        # parse "1-3,5,7-9"
        pages_set = set()
        for part in ranges.split(","):
            part = part.strip()
            if "-" in part:
                a, b = part.split("-", 1)
                pages_set.update(range(int(a) - 1, int(b)))
            else:
                pages_set.add(int(part) - 1)
        pages_sorted = sorted(p for p in pages_set if 0 <= p < total)
        writer = pypdf.PdfWriter()
        for p in pages_sorted:
            writer.add_page(reader.pages[p])
        out = os.path.join(out_dir, f"{base}_extract.pdf")
        with open(out, "wb") as f:
            writer.write(f)
        out_files.append(out)
        if log_widget: log_append(log_widget, f"Extracted {len(pages_sorted)} pages", "success")

    return out_files


def merge_pdfs(pdf_list, out_path, log_widget=None):
    if not PYPDF_AVAILABLE:
        raise RuntimeError("pypdf not installed")
    writer = pypdf.PdfWriter()
    for path in pdf_list:
        try:
            writer.append(path)
            if log_widget: log_append(log_widget, f"  Appended: {os.path.basename(path)}")
        except Exception as e:
            if log_widget: log_append(log_widget, f"  Failed: {path} — {e}", "warn")
    ensure_dir(out_path)
    with open(out_path, "wb") as f:
        writer.write(f)
    if log_widget: log_append(log_widget, f"Merged → {out_path}", "success")
    return out_path


def extract_text_pdfplumber(src, out_txt, log_widget=None):
    if not PDFPLUMBER_AVAILABLE:
        raise RuntimeError("pdfplumber not installed: pip install pdfplumber")
    text_parts = []
    with pdfplumber.open(src) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            t = page.extract_text() or ""
            text_parts.append(f"{'─'*60}\nPAGE {i}\n{'─'*60}\n{t}")
            if log_widget: log_append(log_widget, f"  Extracted page {i}/{len(pdf.pages)}")
    full = "\n\n".join(text_parts)
    ensure_dir(out_txt)
    with open(out_txt, "w", encoding="utf-8") as f:
        f.write(full)
    if log_widget: log_append(log_widget, f"Text saved: {out_txt}", "success")
    return full


def extract_tables_pdfplumber(src, out_xlsx, log_widget=None):
    if not PDFPLUMBER_AVAILABLE:
        raise RuntimeError("pdfplumber not installed")
    if pd is None:
        raise RuntimeError("pandas not installed")
    all_tables = []
    with pdfplumber.open(src) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            tables = page.extract_tables()
            for j, tbl in enumerate(tables):
                if tbl:
                    df = pd.DataFrame(tbl[1:], columns=tbl[0])
                    df.insert(0, "_page", i)
                    df.insert(1, "_table", j + 1)
                    all_tables.append(df)
            if log_widget and tables:
                log_append(log_widget, f"  Page {i}: {len(tables)} table(s)")
    if not all_tables:
        raise RuntimeError("No tables found in PDF")
    combined = pd.concat(all_tables, ignore_index=True)
    ensure_dir(out_xlsx)
    combined.to_excel(out_xlsx, index=False)
    if log_widget:
        log_append(log_widget, f"Tables saved: {out_xlsx}  ({len(all_tables)} tables total)", "success")
    return combined


def encrypt_pdf(src, out, user_pw, owner_pw="", log_widget=None):
    if PIKEPDF_AVAILABLE:
        with pikepdf.open(src) as pdf:
            enc = pikepdf.Encryption(owner=owner_pw or user_pw, user=user_pw)
            pdf.save(out, encryption=enc)
    elif PYPDF_AVAILABLE:
        r = pypdf.PdfReader(src)
        w = pypdf.PdfWriter()
        for p in r.pages: w.add_page(p)
        w.encrypt(user_pw, owner_pw or user_pw)
        with open(out, "wb") as f: w.write(f)
    else:
        raise RuntimeError("pikepdf or pypdf required")
    if log_widget: log_append(log_widget, f"Encrypted PDF saved: {out}", "success")


def decrypt_pdf(src, out, password, log_widget=None):
    if PIKEPDF_AVAILABLE:
        with pikepdf.open(src, password=password) as pdf:
            pdf.save(out)
    elif PYPDF_AVAILABLE:
        r = pypdf.PdfReader(src)
        r.decrypt(password)
        w = pypdf.PdfWriter()
        for p in r.pages: w.add_page(p)
        with open(out, "wb") as f: w.write(f)
    else:
        raise RuntimeError("pikepdf or pypdf required")
    if log_widget: log_append(log_widget, f"Decrypted PDF saved: {out}", "success")


def add_watermark_text(src, out, text, log_widget=None):
    """Burn a diagonal text watermark into every page using reportlab + pypdf."""
    if not REPORTLAB_AVAILABLE or not PYPDF_AVAILABLE:
        raise RuntimeError("reportlab and pypdf both required for watermarking")
    from reportlab.pdfgen import canvas as rl_canvas
    import io, math

    # Build watermark page in memory
    packet = io.BytesIO()
    c = rl_canvas.Canvas(packet, pagesize=A4)
    w, h = A4
    c.setFont("Helvetica-Bold", 48)
    c.setFillColorRGB(0.7, 0.7, 0.7, alpha=0.3)
    c.saveState()
    c.translate(w / 2, h / 2)
    c.rotate(45)
    c.drawCentredString(0, 0, text)
    c.restoreState()
    c.save()
    packet.seek(0)

    wm_reader = pypdf.PdfReader(packet)
    wm_page   = wm_reader.pages[0]

    reader = pypdf.PdfReader(src)
    writer = pypdf.PdfWriter()
    for page in reader.pages:
        page.merge_page(wm_page)
        writer.add_page(page)
    with open(out, "wb") as f:
        writer.write(f)
    if log_widget: log_append(log_widget, f"Watermarked PDF saved: {out}", "success")


def get_pdf_metadata(src):
    """Return metadata dict for a PDF."""
    if PIKEPDF_AVAILABLE:
        with pikepdf.open(src) as pdf:
            meta = dict(pdf.docinfo) if pdf.docinfo else {}
            return {str(k): str(v) for k, v in meta.items()}, len(pdf.pages)
    if PYPDF_AVAILABLE:
        r = pypdf.PdfReader(src)
        m = r.metadata or {}
        return {str(k): str(v) for k, v in m.items()}, len(r.pages)
    return {}, 0


def rotate_pdf(src, out, degrees, page_range="all", log_widget=None):
    if not PYPDF_AVAILABLE:
        raise RuntimeError("pypdf required")
    reader = pypdf.PdfReader(src)
    writer = pypdf.PdfWriter()
    total = len(reader.pages)
    if page_range == "all":
        indices = range(total)
    else:
        indices = [int(x) - 1 for x in page_range.split(",") if x.strip().isdigit()]
    for i, page in enumerate(reader.pages):
        if i in indices:
            page.rotate(degrees)
        writer.add_page(page)
    with open(out, "wb") as f:
        writer.write(f)
    if log_widget: log_append(log_widget, f"Rotated {degrees}° → {out}", "success")


def restore_pdf_quality(
    src, out,
    render_dpi=300,
    denoise_strength=10,
    sharpen_strength=120,
    use_clahe=True,
    bilateral_d=0,
    mode="document",
    log_widget=None,
):
    """Re-render every page at render_dpi then run the full OpenCV restoration
    pipeline.  Output is a PDF built from the processed page images.

    Pipeline:
      1. Render each page at render_dpi using the best available renderer
         (PyMuPDF > pypdfium2 > pdf2image / poppler)
      2. apply_restoration_pipeline — denoise + CLAHE + sharpen + bilateral
      3. Reassemble into a PDF via Pillow

    render_dpi       : 72 – 600  (higher = sharper, bigger file)
    denoise_strength : 0 = off, 1-30  (fastNlMeans h-param)
    sharpen_strength : 0 = off, 1-200 (unsharp percent)
    use_clahe        : local contrast enhancement — best for text
    bilateral_d      : 0 = off, 5-15 = edge-preserving final pass
    mode             : "document" | "photo" | "mixed"
    """
    if not PIL_AVAILABLE:
        raise RuntimeError("Pillow not installed: pip install pillow")

    n_pages = count_pdf_pages(src)
    log_append(log_widget,
               f"Restore: {n_pages} page(s) @ {render_dpi} dpi  "
               f"denoise={denoise_strength}  sharpen={sharpen_strength}  "
               f"clahe={use_clahe}  bilateral={bilateral_d}  mode={mode}", "info")

    processed = []
    for i in range(1, n_pages + 1):
        raw = render_pdf_page_pil(src, i, dpi=render_dpi)
        enhanced = apply_restoration_pipeline(
            raw,
            denoise_strength=denoise_strength,
            sharpen_strength=sharpen_strength,
            use_clahe=use_clahe,
            bilateral_d=bilateral_d,
            mode=mode,
        )
        processed.append(enhanced)
        log_append(log_widget, f"  page {i}/{n_pages}")

    ensure_dir(out)
    processed[0].save(
        out, "PDF", resolution=render_dpi,
        save_all=True, append_images=processed[1:],
    )
    before = os.path.getsize(src)
    after  = os.path.getsize(out)
    log_append(log_widget,
               f"Restored: {human_size(before)} → {human_size(after)}  →  {out}",
               "success")


def run_ocrmypdf(src, out, log_widget=None):
    if not shutil.which("ocrmypdf"):
        raise RuntimeError("ocrmypdf not found. Install: pip install ocrmypdf  (also needs tesseract)")
    result = subprocess.run(["ocrmypdf", "--force-ocr", src, out],
                            capture_output=True, text=True)
    if result.returncode == 0:
        if log_widget: log_append(log_widget, f"OCR completed: {out}", "success")
    else:
        if log_widget: log_append(log_widget, f"OCR error: {result.stderr}", "error")
        raise RuntimeError(result.stderr)


# =====================================================================
#  PDF CRACKING BACKEND
# =====================================================================

def _find_tool(*names):
    """Return the first name found on PATH, or None."""
    for n in names:
        p = shutil.which(n)
        if p:
            return p
    return None


def _find_pdf2john():
    """Locate pdf2john in all known locations."""
    # explicit executables
    direct = _find_tool("pdf2john", "pdf2john.pl", "pdf2john.py")
    if direct:
        return direct
    # john install dirs on common platforms
    search_dirs = []
    if sys.platform == "win32":
        for base in (os.environ.get("ProgramFiles", "C:\\Program Files"),
                     os.environ.get("ProgramFiles(x86)", ""),
                     os.path.expanduser("~")):
            search_dirs.append(os.path.join(base, "John", "run"))
            search_dirs.append(os.path.join(base, "john", "run"))
    else:
        search_dirs += ["/usr/share/john", "/usr/lib/john",
                        "/opt/john/run", "/snap/john/current/run"]
    for d in search_dirs:
        for name in ("pdf2john.pl", "pdf2john.py", "pdf2john"):
            p = os.path.join(d, name)
            if os.path.exists(p):
                return p
    return None


def _find_john():
    return _find_tool("john", "john.exe")


def _find_hashcat():
    return _find_tool("hashcat", "hashcat.exe", "hashcat64.exe",
                      "hashcat32.exe")


def extract_pdf_hash(pdf_path, log_widget=None):
    """Run pdf2john to extract the hash string from a PDF.

    Returns (hash_string, hash_file_path) on success, (None, None) on failure.
    The hash file is written to a temp dir; caller is responsible for cleanup.
    """
    p2j = _find_pdf2john()
    if not p2j:
        log_append(log_widget,
                   "pdf2john not found.  Install John the Ripper and ensure "
                   "pdf2john.pl / pdf2john.py is on PATH or in the JtR run/ dir.",
                   "error")
        return None, None

    hash_file = os.path.join(tempfile.gettempdir(),
                             f"pdftk_{os.getpid()}.hash")
    try:
        # pdf2john can be a Perl script, Python script, or native binary
        if p2j.endswith(".pl"):
            cmd = ["perl", p2j, pdf_path]
        elif p2j.endswith(".py"):
            cmd = [sys.executable, p2j, pdf_path]
        else:
            cmd = [p2j, pdf_path]

        result = subprocess.run(cmd, capture_output=True, text=True,
                                timeout=30)
        raw = (result.stdout or "").strip()
        if not raw:
            log_append(log_widget,
                       f"pdf2john produced no output.  stderr: "
                       f"{result.stderr[:200]}", "error")
            return None, None

        with open(hash_file, "w", encoding="utf-8") as f:
            f.write(raw + "\n")

        log_append(log_widget,
                   f"Hash extracted → {hash_file}", "info")
        log_append(log_widget,
                   f"  {raw[:80]}{'...' if len(raw) > 80 else ''}", "info")
        return raw, hash_file

    except Exception as exc:
        log_append(log_widget, f"Hash extraction failed: {exc}", "error")
        if os.path.exists(hash_file):
            try: os.remove(hash_file)
            except: pass
        return None, None


def _john_show(hash_file):
    """Return cracked password from john --show, or None."""
    john = _find_john()
    if not john:
        return None
    try:
        r = subprocess.run([john, "--show", hash_file],
                           capture_output=True, text=True, timeout=10)
        for line in r.stdout.splitlines():
            if ":" in line and not line.startswith("#"):
                parts = line.split(":")
                if len(parts) >= 2 and parts[1]:
                    return parts[1]
    except Exception:
        pass
    return None


def crack_with_john(
    pdf_path,
    hash_file,
    mode="wordlist",          # "wordlist" | "incremental" | "markov" | "mask"
    wordlist=None,
    mask=None,                # e.g. ?a?a?a?a?a?a
    markov_level=200,
    session_name=None,
    restore_session=False,
    log_widget=None,
    line_cb=None,             # called with each stdout line (for live log)
    cancel_flag=None,         # threading.Event — set to stop
):
    """Run John the Ripper against a pre-extracted hash file.

    Modes
    ─────
    wordlist    --wordlist=<file>  (or stdin if wordlist is None → falls back
                                    to built-in john wordlist)
    incremental --incremental  (brute-force, no wordlist needed)
    markov      --markov=<level>  (human-pattern aware, ~200 is good default)
    mask        --mask=<pattern>  (e.g. ?u?l?l?l?d?d = Cap+4lower+2digit)

    Returns the cracked password string or None.
    """
    john = _find_john()
    if not john:
        log_append(log_widget,
                   "john not found on PATH.  Install John the Ripper.", "error")
        return None

    cmd = [john]

    # session management
    if restore_session and session_name:
        cmd += [f"--restore={session_name}"]
    else:
        if session_name:
            cmd += [f"--session={session_name}"]

        if mode == "wordlist":
            if wordlist and os.path.exists(wordlist):
                cmd += [f"--wordlist={wordlist}"]
                cmd += ["--rules"]            # mangle rules improve coverage
            else:
                cmd += ["--wordlist"]         # use john's built-in wordlist
        elif mode == "incremental":
            cmd += ["--incremental"]
        elif mode == "markov":
            cmd += [f"--markov={markov_level}"]
        elif mode == "mask":
            if not mask:
                log_append(log_widget, "Mask mode requires a mask pattern.", "error")
                return None
            cmd += [f"--mask={mask}"]

        cmd += [hash_file]

    log_append(log_widget, f"JtR cmd: {' '.join(cmd)}", "info")

    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1)

        for line in proc.stdout:
            line = line.rstrip()
            if line:
                log_append(log_widget, f"  JtR: {line}")
                if line_cb:
                    line_cb(line)
            if cancel_flag and cancel_flag.is_set():
                proc.terminate()
                log_append(log_widget, "JtR terminated by user.", "warn")
                return None

        proc.wait()
        pw = _john_show(hash_file)
        if pw:
            log_append(log_widget, f"✓ PASSWORD: {pw}", "success")
        else:
            log_append(log_widget, "JtR: no password found in this run.", "warn")
        return pw

    except Exception as exc:
        log_append(log_widget, f"JtR error: {exc}", "error")
        return None


def detect_pdf_hash_mode(hash_string):
    """Detect the hashcat hash-mode number from a pdf2john hash string.

    Returns (mode_int, description_str).
    """
    # pdf2john hash format: filename:$pdf$<ver>*<rev>*<enc>*...
    import re
    m = re.search(r'\$pdf\$(\d+)\*(\d+)\*(\d+)', hash_string)
    if not m:
        return 10500, "PDF 1.1–1.3 (mode 10400/10500 — fallback)"

    ver = int(m.group(1))
    rev = int(m.group(2))
    bits = int(m.group(3))

    # mapping based on hashcat hash-mode table (as of hashcat 6.2.6)
    if ver == 2 and bits == 40:
        return 10400, "PDF 1.1–1.3  v2/r2  40-bit RC4  (Adobe Acrobat 3–4)"
    if ver == 3 and bits == 128:
        return 10500, "PDF 1.4–1.6  v3/r3  128-bit RC4  (Adobe Acrobat 5–7)"
    if ver == 4 and rev == 4:
        return 10600, "PDF 1.5–1.7  v4/r4  128-bit RC4  (Adobe Acrobat 7–8)"
    if ver == 5 and rev == 5:
        return 10700, "PDF 1.7 Level 3  v5/r5  256-bit AES  (Adobe Acrobat 9)"
    if ver == 5 and rev == 6:
        return 10700, "PDF 1.7 Level 8  v5/r6  256-bit AES  (Adobe Acrobat X–XI)"
    # newest / unknown
    return 10700, f"PDF v{ver}/r{rev}  {bits}-bit  (autodetected)"


_HASHCAT_MODE_MAP = {
    "PDF 1.1–1.3 (Adobe 3–4)  40-bit RC4":         10400,
    "PDF 1.4–1.6 (Adobe 5–7)  128-bit RC4":        10500,
    "PDF 1.5–1.7 (Adobe 7–8)  128-bit RC4 v4r4":   10600,
    "PDF 1.7 Level 3 (Adobe 9)  256-bit AES v5r5":  10700,
    "PDF 1.7 Level 8 (Adobe X–XI)  256-bit AES v5r6": 10700,
    "PDF 2.0 (Adobe DC)  256-bit AES":              25400,
    "Auto-detect from hash":                         None,
}


def crack_with_hashcat(
    hash_file,
    attack_mode=3,            # 3 = mask  (brute-force)
    hash_mode=10500,          # hashcat -m value
    mask=None,                # e.g. ?a?a?a?a?a?a
    wordlist=None,            # used in attack mode 0
    rules_file=None,
    workload=3,               # 1-4, 3 = default
    log_widget=None,
    line_cb=None,
    cancel_flag=None,
):
    """Run hashcat against a pdf2john hash file.

    attack_mode 0 = wordlist, 3 = mask (brute-force), 6 = hybrid wordlist+mask
    """
    hc = _find_hashcat()
    if not hc:
        log_append(log_widget,
                   "hashcat not found on PATH.  Install from "
                   "https://hashcat.net/hashcat/", "error")
        return None

    cmd = [hc,
           f"-m{hash_mode}",
           f"-a{attack_mode}",
           f"-w{workload}",
           "--status", "--status-timer=5",
           "--potfile-disable",
           hash_file]

    if attack_mode == 0:
        if not wordlist:
            log_append(log_widget, "Wordlist required for attack mode 0.", "error")
            return None
        cmd.append(wordlist)
        if rules_file:
            cmd += ["-r", rules_file]
    elif attack_mode == 3:
        default_mask = mask or "?a?a?a?a?a?a"
        cmd.append(default_mask)
    elif attack_mode == 6:
        if not wordlist:
            log_append(log_widget, "Wordlist required for hybrid mode.", "error")
            return None
        cmd += [wordlist, mask or "?a?a?a"]

    log_append(log_widget, f"hashcat cmd: {' '.join(cmd)}", "info")

    found_pw = None
    try:
        proc = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1)

        for line in proc.stdout:
            line = line.rstrip()
            if line:
                log_append(log_widget, f"  hc: {line}")
                if line_cb:
                    line_cb(line)
                # hashcat prints cracked hash:password
                if ":" in line and not line.startswith("["):
                    parts = line.rsplit(":", 1)
                    if len(parts) == 2 and parts[1]:
                        found_pw = parts[1].strip()
            if cancel_flag and cancel_flag.is_set():
                proc.terminate()
                log_append(log_widget, "hashcat terminated by user.", "warn")
                return None

        proc.wait()

    except Exception as exc:
        log_append(log_widget, f"hashcat error: {exc}", "error")
        return None

    if found_pw:
        log_append(log_widget, f"✓ PASSWORD: {found_pw}", "success")
    else:
        log_append(log_widget, "hashcat: no password found in this run.", "warn")
    return found_pw


def probe_gpu():
    """Return True if hashcat can see a GPU (runs hashcat -I)."""
    hc = _find_hashcat()
    if not hc:
        return False
    try:
        r = subprocess.run([hc, "-I"], capture_output=True, text=True, timeout=10)
        out = r.stdout + r.stderr
        return any(kw in out.lower()
                   for kw in ("opencl", "cuda", "metal", "gpu", "nvidia",
                               "amd", "intel arc"))
    except Exception:
        return False


# =====================================================================
#  PDF PREVIEW RENDER HELPERS
# =====================================================================

def render_pdf_page_pil(src, page_num=1, dpi=150):
    """Render a single PDF page to a PIL RGB Image (page_num is 1-indexed).

    Renderer priority:
      1. PyMuPDF (fitz)   — fastest, no system deps, sub-pixel AA
      2. pypdfium2        — Google's PDFium engine, also no system deps
      3. pdf2image        — needs poppler (system dep), slowest
    Raises RuntimeError if none are available.
    """
    if PYMUPDF_AVAILABLE:
        doc  = fitz.open(src)
        page = doc[page_num - 1]                           # 0-indexed
        mat  = fitz.Matrix(dpi / 72.0, dpi / 72.0)
        pix  = page.get_pixmap(matrix=mat, alpha=False, colorspace=fitz.csRGB)
        doc.close()
        if PIL_AVAILABLE:
            return Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
        # numpy path when PIL not available
        if NUMPY_AVAILABLE:
            arr = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.height, pix.width, 3)
            return arr   # caller must handle ndarray
        raise RuntimeError("Need Pillow or numpy to decode PyMuPDF pixmap")

    if PYPDFIUM2_AVAILABLE:
        pdf  = pdfium.PdfDocument(src)
        page = pdf[page_num - 1]
        scale = dpi / 72.0
        bmp  = page.render(scale=scale, rotation=0)
        img  = bmp.to_pil()
        pdf.close()
        return img.convert("RGB") if img.mode != "RGB" else img

    if PDF2IMAGE_AVAILABLE:
        pages = convert_from_path(
            src, dpi=dpi, first_page=page_num, last_page=page_num,
            thread_count=2, fmt="ppm")
        if not pages:
            raise RuntimeError(f"Page {page_num} not found in {src!r}")
        img = pages[0]
        return img.convert("RGB") if img.mode != "RGB" else img

    raise RuntimeError(
        "No PDF renderer found.\n"
        "Install one of:  pip install pymupdf  |  pip install pypdfium2  |  "
        "pip install pdf2image  (+ poppler)")


def count_pdf_pages(src):
    """Return page count — uses the fastest available method."""
    if PYMUPDF_AVAILABLE:
        try:
            doc = fitz.open(src)
            n   = len(doc)
            doc.close()
            return n
        except Exception:
            pass
    if PYPDFIUM2_AVAILABLE:
        try:
            pdf = pdfium.PdfDocument(src)
            n   = len(pdf)
            pdf.close()
            return n
        except Exception:
            pass
    if PYPDF_AVAILABLE:
        try:
            return len(pypdf.PdfReader(src).pages)
        except Exception:
            pass
    if PIKEPDF_AVAILABLE:
        try:
            with pikepdf.open(src) as pdf:
                return len(pdf.pages)
        except Exception:
            pass
    return 1


def simulate_gs_compression(img, target_dpi, ref_dpi=150):
    """Visually simulate Ghostscript lossy compression for preview.

    Down-samples then up-samples with NEAREST so block artefacts and
    pixelation are clearly visible.  Returns img unchanged when
    target_dpi >= ref_dpi (no degradation expected).
    Uses OpenCV when available for a more realistic JPEG-artefact simulation.
    """
    if target_dpi >= ref_dpi:
        return img

    if CV2_AVAILABLE and NUMPY_AVAILABLE and PIL_AVAILABLE:
        arr   = np.array(img, dtype=np.uint8)
        bgr   = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
        scale = target_dpi / ref_dpi
        h, w  = bgr.shape[:2]
        small = cv2.resize(bgr, (max(1, int(w * scale)), max(1, int(h * scale))),
                           interpolation=cv2.INTER_AREA)
        # simulate JPEG block artefacts with a low-quality re-encode
        q = max(1, min(95, int(target_dpi / 6)))
        _, enc = cv2.imencode(".jpg", small, [cv2.IMWRITE_JPEG_QUALITY, q])
        small  = cv2.imdecode(enc, cv2.IMREAD_COLOR)
        out    = cv2.resize(small, (w, h), interpolation=cv2.INTER_NEAREST)
        return Image.fromarray(cv2.cvtColor(out, cv2.COLOR_BGR2RGB))

    if not PIL_AVAILABLE:
        return img
    w, h  = img.size
    scale = target_dpi / ref_dpi
    small = img.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
    return small.resize((w, h), Image.NEAREST)


# ──────────────────────────────────────────────────────────────────────────────
#  FULL OPENCV RESTORATION PIPELINE
# ──────────────────────────────────────────────────────────────────────────────

def _pil_to_bgr(img):
    """PIL RGB → numpy BGR."""
    return cv2.cvtColor(np.array(img, dtype=np.uint8), cv2.COLOR_RGB2BGR)


def _bgr_to_pil(arr):
    """numpy BGR → PIL RGB."""
    return Image.fromarray(cv2.cvtColor(arr, cv2.COLOR_BGR2RGB))


def apply_restoration_pipeline(
    img,
    denoise_strength=10,    # 0 = off, 1-30 = h param for fastNlMeans
    sharpen_strength=120,   # 0 = off, percent for UnsharpMask / custom kernel
    use_clahe=True,         # local contrast enhancement (great for scanned docs)
    bilateral_d=0,          # 0 = off, 5-9 = bilateral filter diameter
    mode="document",        # "document" | "photo" | "mixed"
):
    """Multi-stage image restoration using OpenCV + Pillow.

    Stages (run only when the relevant strength > 0 / flag is True):
      1. fastNlMeansDenoisingColored  — removes JPEG/compression noise
      2. CLAHE in LAB space          — local contrast boost (text legibility)
      3. Sharpening
           document mode → Laplacian unsharp mask (hard edges, crisp text)
           photo mode    → Pillow UnsharpMask (smoother, film-like)
           mixed         → OpenCV detail enhance
      4. Bilateral filter            — edge-preserving final smooth

    Falls back gracefully to Pillow-only when cv2 / numpy not available.
    """
    if not PIL_AVAILABLE:
        return img

    # ── Pillow-only fast path ────────────────────────────────────────
    if not CV2_AVAILABLE or not NUMPY_AVAILABLE:
        from PIL import ImageFilter, ImageEnhance
        result = img
        if sharpen_strength > 0:
            pct = sharpen_strength
            result = result.filter(
                ImageFilter.UnsharpMask(radius=1.5, percent=pct, threshold=2))
        if use_clahe:
            result = ImageEnhance.Contrast(result).enhance(1.25)
        return result

    # ── Full OpenCV pipeline ─────────────────────────────────────────
    bgr = _pil_to_bgr(img)

    # Stage 1 — denoise
    if denoise_strength > 0:
        h = int(denoise_strength)
        bgr = cv2.fastNlMeansDenoisingColored(
            bgr, None,
            h=h, hColor=max(1, h // 2),
            templateWindowSize=7,
            searchWindowSize=21,
        )

    # Stage 2 — CLAHE in LAB (boosts local contrast without blowing out highlights)
    if use_clahe:
        lab        = cv2.cvtColor(bgr, cv2.COLOR_BGR2LAB)
        l, a, b    = cv2.split(lab)
        clip_limit = 2.5 if mode == "document" else 1.8
        clahe      = cv2.createCLAHE(clipLimit=clip_limit, tileGridSize=(8, 8))
        l          = clahe.apply(l)
        bgr        = cv2.cvtColor(cv2.merge([l, a, b]), cv2.COLOR_LAB2BGR)

    # Stage 3 — sharpening
    if sharpen_strength > 0:
        if mode == "document":
            # Laplacian-based unsharp mask — maximises text edge crispness
            blur      = cv2.GaussianBlur(bgr, (0, 0), sigmaX=1.5)
            alpha     = sharpen_strength / 100.0
            bgr       = cv2.addWeighted(bgr, 1 + alpha, blur, -alpha, 0)
        elif mode == "photo":
            # Softer unsharp mask via Pillow (avoids haloing on gradients)
            pil_tmp   = _bgr_to_pil(bgr)
            from PIL import ImageFilter
            pil_tmp   = pil_tmp.filter(
                ImageFilter.UnsharpMask(
                    radius=2.0, percent=sharpen_strength, threshold=3))
            bgr       = _pil_to_bgr(pil_tmp)
        else:
            # Mixed → OpenCV detail enhance (preserves textures)
            sigma_s   = max(10, min(200, sharpen_strength * 2))
            sigma_r   = 0.15
            bgr       = cv2.detailEnhance(bgr, sigma_s=sigma_s, sigma_r=sigma_r)

    # Stage 4 — bilateral filter (edge-preserving smooth to kill leftover noise)
    if bilateral_d >= 5:
        d         = int(bilateral_d)
        sigma_c   = 40 if mode == "document" else 60
        bgr       = cv2.bilateralFilter(bgr, d, sigma_c, sigma_c)

    return _bgr_to_pil(bgr)


def apply_sharpen_filter(img, sharpen=True):
    """Legacy shim — delegates to the full pipeline."""
    if not sharpen or img is None:
        return img
    return apply_restoration_pipeline(
        img, denoise_strength=0, sharpen_strength=120,
        use_clahe=False, bilateral_d=0, mode="document")


def fit_pil_to_box(img, box_w, box_h):
    """Resize a PIL image to fit inside (box_w × box_h) keeping aspect ratio.

    Never upscales beyond the image's native size.
    """
    if not PIL_AVAILABLE or img is None:
        return img
    iw, ih = img.size
    if iw <= 0 or ih <= 0:
        return img
    scale = min(box_w / iw, box_h / ih, 1.0)
    return img.resize(
        (max(1, int(iw * scale)), max(1, int(ih * scale))),
        Image.LANCZOS,
    )


# =====================================================================
#  PREVIEW PANE WIDGET  (reusable before/after PDF viewer)
# =====================================================================

# =====================================================================
#  SCROLLABLE FRAME HELPER  (used by restore + compress left panels)
# =====================================================================

class ScrollableFrame(ttk.Frame):
    """A vertically-scrollable ttk.Frame.

    Use  .inner  as the parent for child widgets.
    Mousewheel is bound automatically on Windows/macOS/Linux.
    """

    def __init__(self, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self._canvas = tk.Canvas(self, bg=DARK["bg"],
                                  highlightthickness=0, bd=0)
        self._sb     = ttk.Scrollbar(self, orient="vertical",
                                      command=self._canvas.yview)
        self._canvas.configure(yscrollcommand=self._sb.set)

        self._canvas.grid(row=0, column=0, sticky="nsew")
        self._sb.grid(row=0, column=1, sticky="ns")

        self.inner = ttk.Frame(self._canvas)
        self._win_id = self._canvas.create_window(
            (0, 0), window=self.inner, anchor="nw")

        self.inner.bind("<Configure>", self._on_inner_configure)
        self._canvas.bind("<Configure>", self._on_canvas_configure)

        # mousewheel — three platform variants
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self._canvas.bind(seq, self._on_wheel, add="+")
            self.inner.bind(seq, self._on_wheel, add="+")

    def _on_inner_configure(self, _e=None):
        self._canvas.configure(
            scrollregion=self._canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        # stretch inner frame to canvas width so pack/grid fill="x" works
        self._canvas.itemconfigure(self._win_id, width=event.width)

    def _on_wheel(self, event):
        if event.num == 4:        delta = -1
        elif event.num == 5:      delta =  1
        else:                     delta = -int(event.delta / 120)
        self._canvas.yview_scroll(delta, "units")

    def scroll_top(self):
        self._canvas.yview_moveto(0)


# =====================================================================
#  HYPERDIMENSIONAL COMPUTING (HDC) ENGINE
# =====================================================================
#
#  Dimension: 8192 bits  (not 4096, not 10 000)
#  ─────────────────────────────────────────────
#  • 10 000 (Kanerva SDM) has no advantage over a power-of-2 here.
#    numpy SIMD operates on 64-bit lanes: 8192 = 128 × uint64 is
#    cache-line-aligned and ~20% faster per XOR than 10 000.
#  • 4096 = 64 × uint64 (512 B/page, 500 KB/1 000 pages).
#    Works, but majority-vote bundling loses accuracy faster and
#    Hamming-distance resolution is halved vs 8192.
#  • 8192 = 128 × uint64 (1 KB/page, 1 MB/1 000 pages). Capacity
#    ~2^4096 quasi-orthogonal vectors — effectively infinite for PDF.
#    SIMD-aligned. This is the sweet spot.  Chosen.
#
#  Structural DNA per page (no pixel data ever touches RAM):
#  ─────────────────────────────────────────────────────────
#  PageVec = bind(KEY_FONT,   bundle(font   vecs))   # XOR association
#          ⊕ bind(KEY_TEXT,   bundle(text   vecs))
#          ⊕ bind(KEY_LAYOUT, layout_vec)
#          ⊕ bind(KEY_IMAGE,  bundle(image  vecs))
#
#  Each component vec is generated deterministically via SHAKE-256.
#  Cache: SQLite at ~/.pdftoolkit/hdc_cache.db (~1 KB × n_pages).
# =====================================================================

HDC_DIM   = 8192
HDC_WORDS = HDC_DIM // 64   # 128 × uint64
HDC_BYTES = HDC_DIM // 8    # 1024 bytes

_KEY_FONT_ARR   = None
_KEY_TEXT_ARR   = None
_KEY_LAYOUT_ARR = None
_KEY_IMAGE_ARR  = None


def _make_role_key(seed_str):
    raw = hashlib.shake_256(seed_str.encode()).digest(HDC_BYTES)
    return np.frombuffer(raw, dtype=np.uint64).copy()


def _init_role_keys():
    global _KEY_FONT_ARR, _KEY_TEXT_ARR, _KEY_LAYOUT_ARR, _KEY_IMAGE_ARR
    if _KEY_FONT_ARR is None and NUMPY_AVAILABLE:
        _KEY_FONT_ARR   = _make_role_key("KEY_FONT::v1")
        _KEY_TEXT_ARR   = _make_role_key("KEY_TEXT::v1")
        _KEY_LAYOUT_ARR = _make_role_key("KEY_LAYOUT::v1")
        _KEY_IMAGE_ARR  = _make_role_key("KEY_IMAGE::v1")


class HDCVec:
    """8192-bit binary HDC vector — 128 × numpy uint64 (1 024 bytes).

    bind(other)      →  XOR  (reversible association)
    bundle(vecs)     →  majority vote  (superposition)
    hamming(other)   →  differing bits  (0 = identical)
    similarity()     →  1 − hamming/DIM  (0.5 = random)
    permute(k)       →  cyclic bit-rotate  (encodes order)
    from_string(s)   →  deterministic via SHAKE-256
    to_bytes() / from_bytes()  →  SQLite serialisation
    """
    __slots__ = ("v",)

    def __init__(self, arr=None):
        if not NUMPY_AVAILABLE:
            raise RuntimeError("numpy required for HDC engine")
        if arr is None:
            self.v = np.random.randint(
                0, np.iinfo(np.uint64).max, HDC_WORDS, dtype=np.uint64)
        else:
            self.v = np.asarray(arr, dtype=np.uint64)
            if len(self.v) != HDC_WORDS:
                raise ValueError(f"Expected {HDC_WORDS} words, got {len(self.v)}")

    def bind(self, other):
        return HDCVec(self.v ^ other.v)

    def __xor__(self, other):
        return self.bind(other)

    def hamming(self, other):
        return int(np.unpackbits((self.v ^ other.v).view(np.uint8)).sum())

    def similarity(self, other):
        return 1.0 - self.hamming(other) / HDC_DIM

    def permute(self, k=1):
        bits = np.unpackbits(self.v.view(np.uint8))
        bits = np.roll(bits, k)
        return HDCVec(np.packbits(bits).view(np.uint64).copy())

    @classmethod
    def bundle(cls, vecs):
        """Majority-vote superposition.  Deterministic tie-breaking."""
        if not vecs:
            return cls()
        if len(vecs) == 1:
            return cls(vecs[0].v.copy())
        stacked = np.stack([v.v for v in vecs])
        bits    = np.unpackbits(stacked.view(np.uint8), axis=1)
        counts  = bits.sum(axis=0, dtype=np.int32)
        n       = len(vecs)
        tie_raw = hashlib.shake_256(b"BUNDLE_TIE_v1").digest(HDC_BYTES)
        tie     = np.unpackbits(np.frombuffer(tie_raw, dtype=np.uint8))
        result  = np.where(counts > n // 2, np.uint8(1),
                  np.where(counts < n // 2, np.uint8(0),
                           tie.astype(np.uint8)))
        return cls(np.packbits(result).view(np.uint64).copy())

    @classmethod
    def from_string(cls, s):
        raw = hashlib.shake_256(s.encode("utf-8", errors="replace")).digest(HDC_BYTES)
        return cls(np.frombuffer(raw, dtype=np.uint64).copy())

    @classmethod
    def from_int(cls, n):
        return cls.from_string(str(n))

    def to_bytes(self):
        return self.v.tobytes()

    @classmethod
    def from_bytes(cls, b):
        return cls(np.frombuffer(b, dtype=np.uint64).copy())

    def __repr__(self):
        return f"<HDCVec {HDC_DIM}d pop={int(np.unpackbits(self.v.view(np.uint8)).sum())}>"


# ── Structural DNA extractor ─────────────────────────────────────────

class PDFStructuralParser:
    """Extract Structural DNA from a PDF page — no pixels, no RAM bloat."""

    @staticmethod
    def extract_pymupdf(doc, idx):
        page   = doc[idx]
        fonts  = [f[3] or f[1] for f in page.get_fonts(full=True)]
        words  = [w[4] for w in page.get_text("words")]
        blocks = page.get_text("dict")["blocks"]
        images = page.get_images(full=False)
        rect   = page.rect
        n_text = sum(1 for b in blocks if b.get("type") == 0)
        n_img  = len(images)
        return {
            "fonts":  fonts[:20],
            "words":  words[:60],
            "layout": (int(rect.width // 72) * 72,
                       int(rect.height // 72) * 72,
                       min(n_text, 12), min(n_img, 8)),
            "images": [img[0] for img in images[:16]],
        }

    @staticmethod
    def extract_pypdf(reader, idx):
        page  = reader.pages[idx]
        fonts = []
        try:
            res = page.get("/Resources")
            if res and "/Font" in res:
                fonts = list(res["/Font"].keys())
        except Exception:
            pass
        text = (page.extract_text() or "").split()[:60]
        box  = page.mediabox
        w    = float(getattr(box, "width",  box[2]))
        h    = float(getattr(box, "height", box[3]))
        return {
            "fonts":  fonts[:20],
            "words":  text,
            "layout": (int(w // 72) * 72, int(h // 72) * 72, 0, 0),
            "images": [],
        }

    @staticmethod
    def encode(dna):
        _init_role_keys()

        def _comp(items, fallback):
            return (HDCVec.bundle([HDCVec.from_string(str(x)) for x in items])
                    if items else HDCVec.from_string(fallback))

        F = _comp(dna["fonts"],  "_nofont_")
        T = _comp(dna["words"],  "_notext_")
        L = HDCVec.from_string(str(dna["layout"]))
        I = _comp(dna["images"], "_noimg_")

        return (HDCVec(_KEY_FONT_ARR).bind(F)
                ^ HDCVec(_KEY_TEXT_ARR).bind(T)
                ^ HDCVec(_KEY_LAYOUT_ARR).bind(L)
                ^ HDCVec(_KEY_IMAGE_ARR).bind(I))


# ── Document index with SQLite cache ─────────────────────────────────

_HDC_CACHE_PATH = Path.home() / ".pdftoolkit" / "hdc_cache.db"


class HDCDocumentIndex:
    """Builds, caches, and queries per-page HDC vectors.

    Cache key: (canonical_path, mtime, fsize, sha256[:16] of first 4 KB)
    Storage  : ~/.pdftoolkit/hdc_cache.db  (SQLite, ~1 KB per page)
    """

    _db_conn = None
    _db_lock = threading.Lock()

    @classmethod
    def _conn(cls):
        if cls._db_conn is None:
            _HDC_CACHE_PATH.parent.mkdir(parents=True, exist_ok=True)
            cls._db_conn = sqlite3.connect(
                str(_HDC_CACHE_PATH), check_same_thread=False)
            cls._db_conn.execute("""
                CREATE TABLE IF NOT EXISTS page_vectors (
                    path TEXT, mtime REAL, fsize INTEGER, fhash TEXT,
                    page_num INTEGER, vector BLOB,
                    PRIMARY KEY (path, mtime, fsize, fhash, page_num))""")
            cls._db_conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_doc
                ON page_vectors (path, mtime, fsize, fhash)""")
            cls._db_conn.commit()
        return cls._db_conn

    @staticmethod
    def _fp(path):
        s = Path(path).stat()
        with open(path, "rb") as f:
            h = hashlib.sha256(f.read(4096)).hexdigest()[:16]
        return s.st_mtime, s.st_size, h

    def __init__(self, path):
        self.path    = str(Path(path).resolve())
        self.n_pages = 0
        self.vectors = {}          # {page_num(1-indexed): HDCVec}
        self._mt, self._fs, self._fh = self._fp(path)
        self._built  = False

    # cache I/O

    def load_from_cache(self):
        rows = self._conn().execute(
            "SELECT page_num,vector FROM page_vectors "
            "WHERE path=? AND mtime=? AND fsize=? AND fhash=?",
            (self.path, self._mt, self._fs, self._fh)).fetchall()
        if not rows:
            return False
        self.vectors = {r[0]: HDCVec.from_bytes(r[1]) for r in rows}
        self.n_pages = max(self.vectors) if self.vectors else 0
        self._built  = True
        return True

    def save_to_cache(self):
        rows = [(self.path, self._mt, self._fs, self._fh,
                 pg, vec.to_bytes()) for pg, vec in self.vectors.items()]
        with self._db_lock:
            self._conn().executemany(
                "INSERT OR REPLACE INTO page_vectors VALUES (?,?,?,?,?,?)", rows)
            self._conn().commit()

    def clear_cache(self):
        with self._db_lock:
            self._conn().execute(
                "DELETE FROM page_vectors WHERE path=?", (self.path,))
            self._conn().commit()

    # build

    def build(self, progress_cb=None, cancel_flag=None):
        if not NUMPY_AVAILABLE:
            return
        if PYMUPDF_AVAILABLE:
            doc = fitz.open(self.path)
            self.n_pages = len(doc)
            for i in range(self.n_pages):
                if cancel_flag and cancel_flag.is_set():
                    doc.close(); return
                try:
                    dna = PDFStructuralParser.extract_pymupdf(doc, i)
                    self.vectors[i + 1] = PDFStructuralParser.encode(dna)
                except Exception:
                    self.vectors[i + 1] = HDCVec.from_int(i + 1)
                if progress_cb:
                    progress_cb(i + 1, self.n_pages)
            doc.close()
        elif PYPDF_AVAILABLE:
            reader = pypdf.PdfReader(self.path)
            self.n_pages = len(reader.pages)
            for i in range(self.n_pages):
                if cancel_flag and cancel_flag.is_set():
                    return
                try:
                    dna = PDFStructuralParser.extract_pypdf(reader, i)
                    self.vectors[i + 1] = PDFStructuralParser.encode(dna)
                except Exception:
                    self.vectors[i + 1] = HDCVec.from_int(i + 1)
                if progress_cb:
                    progress_cb(i + 1, self.n_pages)
        self._built = True

    # query

    def nearest_pages(self, page_num, n=8):
        cur = self.vectors.get(page_num)
        if cur is None:
            return []
        results = sorted(
            [(pg, cur.similarity(v)) for pg, v in self.vectors.items()
             if pg != page_num],
            key=lambda x: -x[1])
        return results[:n]

    def xor_navigate(self, page_num, query_vec):
        """Semantic jump: XOR current page toward query, find nearest real page."""
        cur = self.vectors.get(page_num)
        if cur is None:
            return page_num
        target = cur ^ query_vec
        return min(self.vectors.items(),
                   key=lambda kv: target.hamming(kv[1]))[0]

    def similarity_map(self, current_page):
        cur = self.vectors.get(current_page)
        if cur is None:
            return {}
        return {pg: cur.similarity(v) for pg, v in self.vectors.items()}


# ── Predictive pre-renderer ──────────────────────────────────────────

class PredictivePreRenderer:
    """Background ThreadPoolExecutor that pre-renders likely-next pages.

    Combines sequential prediction (±1, +2, +3) with HDC semantic
    prediction (nearest_pages by Hamming similarity).
    """

    def __init__(self, path, index, dpi=150, max_workers=2):
        self._path    = path
        self._index   = index
        self._dpi     = dpi
        self._cache   = {}
        self._lock    = threading.Lock()
        self._pending = set()
        self._pool    = concurrent.futures.ThreadPoolExecutor(
                            max_workers=max_workers,
                            thread_name_prefix="prerender")
        self._alive   = True

    def stop(self):
        self._alive = False
        self._pool.shutdown(wait=False, cancel_futures=True)

    def get(self, page_num):
        with self._lock:
            return self._cache.get(page_num)

    def evict_far(self, cur, radius=12):
        with self._lock:
            for p in [k for k in self._cache if abs(k - cur) > radius]:
                del self._cache[p]

    def prime(self, cur, n_semantic=4):
        if not self._alive:
            return
        n = self._index.n_pages
        cands = {cur + d for d in (1, 2, 3, -1) if 1 <= cur + d <= n}
        for pg, _ in self._index.nearest_pages(cur, n_semantic):
            if 1 <= pg <= n:
                cands.add(pg)
        for pg in cands:
            with self._lock:
                if pg in self._cache or pg in self._pending:
                    continue
                self._pending.add(pg)
            self._pool.submit(self._render, pg)

    def _render(self, pg):
        try:
            img = render_pdf_page_pil(self._path, pg, self._dpi)
        except Exception:
            img = None
        with self._lock:
            self._pending.discard(pg)
            if img is not None:
                self._cache[pg] = img


# =====================================================================
#  PDF VIEWER PANE
# =====================================================================

class PDFViewerPane:
    """Full-featured HDC-powered single-document PDF viewer.

    Layout
    ──────
    PanedWindow (horizontal)
    ├── LEFT  — scrollable thumbnail strip
    │           similarity bars coloured by HDC Hamming distance
    └── RIGHT — main canvas + nav bar + HDC status bar
                zoom/pan/drag (same model as PreviewPane)
                keyboard: arrow keys / PgUp / PgDn
    """

    THUMB_W   = 110
    THUMB_H   = 150
    THUMB_DPI = 18
    VIEW_DPI  = 150

    def __init__(self, parent, log_widget):
        self.log           = log_widget
        self._path         = None
        self._index        = None
        self._prerender    = None
        self._n_pages      = 0
        self._cur_page     = 1
        self._pil_main     = None
        self._photo_main   = None
        self._thumb_photos = {}
        self._thumb_pil    = {}
        self._thumb_frames = {}
        self._index_cancel = threading.Event()
        self._prerender_on = tk.BooleanVar(value=False)
        self._vp = {"zoom": 1.0, "ox": 0, "oy": 0,
                    "drag_x": None, "drag_y": None}
        self._render_gen   = 0

        self._thumb_pool = concurrent.futures.ThreadPoolExecutor(
            max_workers=2, thread_name_prefix="thumb")

        self.frame = ttk.Frame(parent)
        self._build_ui()

    # ── UI ────────────────────────────────────────────────────────────

    def _build_ui(self):
        outer = tk.PanedWindow(
            self.frame, orient="horizontal", sashwidth=5,
            bg=DARK["border"], bd=0, sashrelief="flat")
        outer.pack(fill="both", expand=True)

        left  = tk.Frame(outer, bg=DARK["bg"])
        right = ttk.Frame(outer)
        outer.add(left,  minsize=self.THUMB_W + 30, width=self.THUMB_W + 34)
        outer.add(right, minsize=320)

        # ── thumbnail strip ────────────────────────────────────────────
        lhdr = tk.Frame(left, bg=DARK["panel"])
        lhdr.pack(fill="x")
        tk.Label(lhdr, text="  PAGES", bg=DARK["panel"],
                 fg=DARK["accent"], font=("Consolas", 8, "bold")
                 ).pack(side="left", pady=4, padx=4)
        self._pg_count_lbl = tk.Label(
            lhdr, text="", bg=DARK["panel"],
            fg=DARK["text_dim"], font=("Segoe UI", 7))
        self._pg_count_lbl.pack(side="right", padx=4)

        tsf = tk.Frame(left, bg=DARK["bg"])
        tsf.pack(fill="both", expand=True)
        self._tc = tk.Canvas(tsf, bg=DARK["bg"],
                              highlightthickness=0,
                              width=self.THUMB_W + 26)
        tsb = ttk.Scrollbar(tsf, orient="vertical",
                             command=self._tc.yview)
        self._tc.configure(yscrollcommand=tsb.set)
        self._tc.pack(side="left", fill="both", expand=True)
        tsb.pack(side="right", fill="y")

        self._ti = tk.Frame(self._tc, bg=DARK["bg"])
        self._tw = self._tc.create_window((0, 0), window=self._ti,
                                           anchor="nw")
        self._ti.bind("<Configure>",
                      lambda _e: self._tc.configure(
                          scrollregion=self._tc.bbox("all")))
        self._tc.bind("<Configure>",
                      lambda e: self._tc.itemconfigure(
                          self._tw, width=e.width))
        for s in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self._tc.bind(s, self._thumb_scroll)
            self._ti.bind(s, self._thumb_scroll)

        # ── nav bar ────────────────────────────────────────────────────
        nav = tk.Frame(right, bg=DARK["panel"])
        nav.pack(fill="x", padx=4, pady=(4, 0))

        def _nb(text, cmd, w=None):
            kw = {"width": w} if w else {}
            ttk.Button(nav, text=text, command=cmd, **kw).pack(
                side="left", padx=1, pady=2)

        _nb("📂 Open", self.open_file)
        ttk.Separator(nav, orient="vertical").pack(
            side="left", fill="y", padx=6, pady=3)
        _nb("◄", lambda: self.goto(self._cur_page - 1), w=2)

        self._page_var = tk.IntVar(value=1)
        self._spin = ttk.Spinbox(nav, from_=1, to=9999, width=5,
                                  textvariable=self._page_var,
                                  command=lambda: self.goto(
                                      self._page_var.get()))
        self._spin.pack(side="left", padx=2)
        self._spin.bind("<Return>",
                        lambda _e: self.goto(self._page_var.get()))

        self._total_lbl = tk.Label(nav, text="/ –",
                                    bg=DARK["panel"], fg=DARK["text_dim"],
                                    font=("Segoe UI", 8))
        self._total_lbl.pack(side="left")
        _nb("►", lambda: self.goto(self._cur_page + 1), w=2)

        ttk.Separator(nav, orient="vertical").pack(
            side="left", fill="y", padx=6, pady=3)
        _nb("−", lambda: self._zoom_delta(-0.18), w=2)
        self._zoom_lbl = tk.Label(nav, text="100%",
                                   bg=DARK["panel"], fg=DARK["accent2"],
                                   font=("Consolas", 8, "bold"), width=5)
        self._zoom_lbl.pack(side="left")
        _nb("+", lambda: self._zoom_delta(+0.18), w=2)
        _nb("⊡ Fit", self._zoom_fit)

        ttk.Separator(nav, orient="vertical").pack(
            side="left", fill="y", padx=6, pady=3)
        tk.Checkbutton(
            nav, text="⚡ Pre-render",
            variable=self._prerender_on,
            command=self._on_prerender_toggle,
            bg=DARK["panel"], fg=DARK["text"],
            selectcolor=DARK["surface"],
            activebackground=DARK["panel"],
            font=("Segoe UI", 8),
        ).pack(side="left", padx=4)

        self._nav_status = tk.Label(
            nav, text="  Open a PDF to begin",
            bg=DARK["panel"], fg=DARK["text_dim"],
            font=("Consolas", 8))
        self._nav_status.pack(side="left", padx=6)

        # ── main canvas ────────────────────────────────────────────────
        self._cnv = tk.Canvas(right, bg=DARK["log_bg"],
                               highlightthickness=0, cursor="crosshair")
        self._cnv.pack(fill="both", expand=True, padx=4, pady=4)

        for s in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self._cnv.bind(s, self._wheel)
        self._cnv.bind("<ButtonPress-1>",   self._drag_start)
        self._cnv.bind("<B1-Motion>",       self._drag_move)
        self._cnv.bind("<ButtonRelease-1>", self._drag_end)
        self._cnv.bind("<Double-Button-1>", lambda _e: self._zoom_fit())
        self._cnv.bind("<Configure>",       self._on_cnv_configure)
        self._cnv.bind("<Left>",  lambda _e: self.goto(self._cur_page - 1))
        self._cnv.bind("<Right>", lambda _e: self.goto(self._cur_page + 1))
        self._cnv.bind("<Prior>", lambda _e: self.goto(self._cur_page - 1))
        self._cnv.bind("<Next>",  lambda _e: self.goto(self._cur_page + 1))

        # ── HDC status bar ─────────────────────────────────────────────
        hbar = tk.Frame(right, bg=DARK["surface"], height=22)
        hbar.pack(fill="x", padx=4, pady=(0, 4))
        hbar.pack_propagate(False)

        self._hdc_lbl = tk.Label(
            hbar,
            text=f"  HDC engine  ·  {HDC_DIM}-bit vectors  ·  "
                 f"1 KB/page  ·  SQLite cache  ·  ready",
            bg=DARK["surface"], fg=DARK["text_dim"],
            font=("Consolas", 7), anchor="w")
        self._hdc_lbl.pack(side="left", fill="both", expand=True, padx=4)

        self._sim_bar = tk.Canvas(hbar, bg=DARK["surface"],
                                   highlightthickness=0,
                                   width=180, height=22)
        self._sim_bar.pack(side="right", padx=4)

        # ── PDF flag panel (collapsible, below HDC bar) ────────────────
        self._flag_panel = PDFFlagPanel(right, log_widget, self)

    # ── open / index ──────────────────────────────────────────────────

    def open_file(self, path=None):
        if path is None:
            path = filedialog.askopenfilename(
                filetypes=[("PDF", "*.pdf")])
        if not path or not os.path.exists(path):
            return

        if self._prerender:
            self._prerender.stop()
            self._prerender = None
        self._index_cancel.set()
        self._index_cancel = threading.Event()

        self._path     = path
        self._cur_page = 1
        self._thumb_photos.clear()
        self._thumb_pil.clear()
        self._thumb_frames.clear()
        for w in self._ti.winfo_children():
            w.destroy()

        self._n_pages = count_pdf_pages(path)
        self._total_lbl.config(text=f"/ {self._n_pages}")
        self._pg_count_lbl.config(text=f"{self._n_pages} pg")
        self._spin.configure(to=self._n_pages)
        self._set_status("Loading …", DARK["accent2"])

        self._build_thumb_slots()
        self.goto(1, force=True)
        self._start_hdc_build(path)

        # load into flag panel (background — pikepdf open can be slow on big files)
        threading.Thread(
            target=lambda: self.frame.after(
                100, lambda: self._flag_panel.load_pdf(path)),
            daemon=True).start()

    def _start_hdc_build(self, path):
        if not NUMPY_AVAILABLE:
            self._set_status("numpy missing — HDC disabled", DARK["warning"])
            return
        cancel = self._index_cancel

        def _worker():
            idx = HDCDocumentIndex(path)
            # 1. try cache (may return in <5ms)
            if idx.load_from_cache():
                self.frame.after(0, lambda: self._on_index_ready(idx, True))
                return
            # 2. build fresh
            def _prog(d, t):
                if not cancel.is_set():
                    self.frame.after(0, lambda dd=d, tt=t:
                        self._set_status(
                            f"HDC indexing {dd}/{tt} …", DARK["accent2"]))
            idx.build(progress_cb=_prog, cancel_flag=cancel)
            if cancel.is_set():
                return
            self.frame.after(0, lambda: self._on_index_ready(idx, False))
            # offer to cache
            self.frame.after(200, lambda: self._offer_cache(idx))

        threading.Thread(target=_worker, daemon=True).start()

    def _offer_cache(self, idx):
        if messagebox.askyesno(
            "Save HDC Index",
            f"Cache the structural index for\n"
            f"{os.path.basename(self._path)}?\n\n"
            f"Size: {idx.n_pages} KB  "
            f"({idx.n_pages} pages × 1 KB)\n\n"
            f"Next open will be instant (loaded from disk).",
            icon="question"):
            threading.Thread(target=idx.save_to_cache, daemon=True).start()

    def _on_index_ready(self, idx, from_cache):
        self._index = idx
        src = "cache ⚡" if from_cache else "built"
        self._set_status(
            f"HDC ready ({src})  ·  {idx.n_pages} KB  ·  {HDC_DIM}-bit",
            DARK["success"])
        self._update_hdc_bar()
        self._update_thumb_sims()
        if self._prerender_on.get():
            self._init_prerender()

    # ── thumbnails ───────────────────────────────────────────────────

    def _build_thumb_slots(self):
        for pg in range(1, self._n_pages + 1):
            slot = tk.Frame(self._ti, bg=DARK["bg"], cursor="hand2")
            slot.pack(pady=(4, 0), padx=6)

            cnv = tk.Canvas(slot, bg=DARK["surface"],
                            width=self.THUMB_W, height=self.THUMB_H,
                            highlightthickness=1,
                            highlightbackground=DARK["border"])
            cnv.pack()
            cnv.create_text(self.THUMB_W // 2, self.THUMB_H // 2,
                            text=str(pg), fill=DARK["text_dim"],
                            font=("Consolas", 10))

            sim_c = tk.Canvas(slot, bg=DARK["surface"],
                              width=self.THUMB_W, height=4,
                              highlightthickness=0)
            sim_c.pack()

            lbl = tk.Label(slot, text=str(pg), bg=DARK["bg"],
                           fg=DARK["text_dim"], font=("Segoe UI", 7))
            lbl.pack()

            self._thumb_frames[pg] = {"slot": slot, "cnv": cnv,
                                       "sim": sim_c, "lbl": lbl}
            for w in (slot, cnv, lbl):
                w.bind("<Button-1>",
                       lambda _e, p=pg: self.goto(p))
            for s in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                for w in (slot, cnv, lbl, sim_c):
                    w.bind(s, self._thumb_scroll)

            self._thumb_pool.submit(self._render_thumb, pg)

        self._tc.yview_moveto(0)

    def _render_thumb(self, pg):
        try:
            img = render_pdf_page_pil(self._path, pg, self.THUMB_DPI)
            sc  = min(self.THUMB_W / img.width, self.THUMB_H / img.height)
            img = img.resize((max(1, int(img.width * sc)),
                               max(1, int(img.height * sc))),
                              Image.LANCZOS)
            self.frame.after(0, lambda p=pg, i=img: self._place_thumb(p, i))
        except Exception:
            pass

    def _place_thumb(self, pg, pil_img):
        if pg not in self._thumb_frames:
            return
        try:
            from PIL import ImageTk
            photo = ImageTk.PhotoImage(pil_img)
        except Exception:
            return
        self._thumb_pil[pg]    = pil_img
        self._thumb_photos[pg] = photo
        tf = self._thumb_frames[pg]
        tf["cnv"].delete("all")
        tf["cnv"].create_image(self.THUMB_W // 2, self.THUMB_H // 2,
                                image=photo, anchor="center")
        if pg == self._cur_page:
            tf["cnv"].configure(highlightbackground=DARK["accent"],
                                highlightthickness=2)

    def _update_thumb_sims(self):
        if not self._index or not self._index.vectors:
            return
        smap = self._index.similarity_map(self._cur_page)
        for pg, tf in self._thumb_frames.items():
            sim  = smap.get(pg, 0.5)
            frac = max(0.0, (sim - 0.5) * 2.0)   # 0.5→0, 1.0→1
            sc   = tf["sim"]
            sc.delete("all")
            w = int(self.THUMB_W * frac)
            if w > 0:
                col = "#00c8c8" if frac > 0.55 else "#445566"
                sc.create_rectangle(0, 0, w, 4, fill=col, outline="")
            if pg == self._cur_page:
                tf["lbl"].config(text=f"► {pg}", fg=DARK["accent"])
                tf["cnv"].configure(highlightbackground=DARK["accent"],
                                    highlightthickness=2)
            else:
                pct = int(frac * 100)
                tf["lbl"].config(
                    text=f"{pg}" + (f" {pct}%" if pct > 25 else ""),
                    fg=DARK["text_dim"])
                tf["cnv"].configure(highlightbackground=DARK["border"],
                                    highlightthickness=1)

    def _thumb_scroll(self, event):
        if event.num == 4:       d = -1
        elif event.num == 5:     d =  1
        elif event.delta > 0:    d = -1
        else:                    d =  1
        self._tc.yview_scroll(d, "units")

    def _scroll_to_thumb(self, pg):
        if self._n_pages > 1:
            frac = (pg - 1) / max(self._n_pages - 1, 1)
            self._tc.yview_moveto(max(0.0, frac - 0.12))

    # ── main canvas ───────────────────────────────────────────────────

    def goto(self, pg, force=False):
        if not self._path:
            return
        pg = max(1, min(pg, self._n_pages))
        if pg == self._cur_page and not force:
            return
        self._cur_page = pg
        self._page_var.set(pg)
        self._render_gen += 1
        gen = self._render_gen
        self._set_status(f"Loading page {pg} …", DARK["accent2"])

        # pre-render hit?
        pre = self._prerender.get(pg) if self._prerender else None
        if pre is not None:
            self._pil_main = pre
            self._zoom_fit()
            self._set_status(f"Page {pg}/{self._n_pages}  [⚡ pre-rendered]",
                             DARK["success"])
        else:
            threading.Thread(target=self._render_main,
                             args=(pg, gen), daemon=True).start()

        self._update_thumb_sims()
        self._scroll_to_thumb(pg)
        self._update_hdc_bar()

        if self._prerender and self._prerender_on.get():
            threading.Thread(target=self._prerender.prime,
                             args=(pg,), daemon=True).start()
            self._prerender.evict_far(pg)

    def _render_main(self, pg, gen):
        try:
            img = render_pdf_page_pil(self._path, pg, self.VIEW_DPI)
        except Exception as exc:
            self.frame.after(0, lambda: self._set_status(
                f"Render error: {exc}", DARK["danger"]))
            return
        self.frame.after(0, lambda p=pg, i=img, g=gen:
                         self._on_rendered(p, i, g))

    def _on_rendered(self, pg, img, gen):
        if gen != self._render_gen:
            return
        self._pil_main = img
        self._zoom_fit()
        self._set_status(
            f"Page {pg}/{self._n_pages}  "
            f"[{img.width}×{img.height} @ {self.VIEW_DPI} dpi]",
            DARK["success"])

    def _draw_main(self):
        img = self._pil_main
        if img is None or not self._cnv.winfo_exists():
            return
        vp = self._vp
        cw = max(self._cnv.winfo_width(),  300)
        ch = max(self._cnv.winfo_height(), 300)
        base  = min(cw / img.width, ch / img.height, 1.0)
        total = base * vp["zoom"]
        nw = max(1, int(img.width  * total))
        nh = max(1, int(img.height * total))
        rsmp = Image.LANCZOS if total < 2 else Image.NEAREST
        try:
            from PIL import ImageTk
            photo = ImageTk.PhotoImage(img.resize((nw, nh), rsmp))
        except Exception:
            return
        self._photo_main = photo
        self._cnv.delete("all")
        self._cnv.create_image(cw // 2 + vp["ox"],
                                ch // 2 + vp["oy"],
                                image=photo, anchor="center")
        self._zoom_lbl.config(text=f"{int(vp['zoom']*100)}%")

    # ── zoom / pan ────────────────────────────────────────────────────

    def _wheel(self, event):
        if event.num == 4:    f = 1.12
        elif event.num == 5:  f = 1/1.12
        elif event.delta > 0: f = 1.12
        else:                 f = 1/1.12
        vp = self._vp
        nz = max(0.05, min(20.0, vp["zoom"] * f))
        s  = nz / vp["zoom"]
        vp["ox"] = int(event.x - s * (event.x - vp["ox"]))
        vp["oy"] = int(event.y - s * (event.y - vp["oy"]))
        vp["zoom"] = nz
        self._draw_main()

    def _zoom_delta(self, d):
        self._vp["zoom"] = max(0.05, min(20.0, self._vp["zoom"] + d))
        self._draw_main()

    def _zoom_fit(self, *_):
        self._vp.update({"zoom": 1.0, "ox": 0, "oy": 0})
        self._draw_main()

    def _drag_start(self, e):
        self._vp["drag_x"] = e.x;  self._vp["drag_y"] = e.y
        self._cnv.configure(cursor="fleur")
        self._cnv.focus_set()

    def _drag_move(self, e):
        vp = self._vp
        if vp["drag_x"] is None: return
        vp["ox"] += e.x - vp["drag_x"];  vp["drag_x"] = e.x
        vp["oy"] += e.y - vp["drag_y"];  vp["drag_y"] = e.y
        self._draw_main()

    def _drag_end(self, _e):
        self._vp["drag_x"] = self._vp["drag_y"] = None
        self._cnv.configure(cursor="crosshair")

    def _on_cnv_configure(self, _e):
        if self._pil_main:
            self._zoom_fit()

    # ── pre-renderer ──────────────────────────────────────────────────

    def _on_prerender_toggle(self):
        if self._prerender_on.get():
            self._init_prerender()
        else:
            if self._prerender:
                self._prerender.stop()
                self._prerender = None

    def _init_prerender(self):
        if not self._path or not self._index:
            return
        if self._prerender:
            self._prerender.stop()
        self._prerender = PredictivePreRenderer(
            self._path, self._index, self.VIEW_DPI)
        threading.Thread(target=self._prerender.prime,
                         args=(self._cur_page,), daemon=True).start()

    # ── HDC bar ───────────────────────────────────────────────────────

    def _update_hdc_bar(self):
        if not self._index or not self._index.vectors:
            return
        n    = self._index.n_pages
        smap = self._index.similarity_map(self._cur_page)
        top5 = sorted(smap.values(), reverse=True)[:5]
        frac = max(0.0, (sum(top5) / len(top5) - 0.5) * 2) if top5 else 0

        self._hdc_lbl.config(
            text=f"  HDC {HDC_DIM}-bit  ·  {n} KB  ·  "
                 f"pg {self._cur_page}/{n}  ·  "
                 f"top-5 semantic sim {int(frac*100)}%")

        sc = self._sim_bar
        sc.delete("all")
        sc.create_text(4, 11, anchor="w", text="sim",
                       fill=DARK["text_dim"], font=("Consolas", 7))
        bx = 28; bw = 142
        sc.create_rectangle(bx, 6, bx+bw, 16,
                             fill=DARK["surface"], outline=DARK["border"])
        fw = int(bw * frac)
        if fw:
            col = "#00c8c8" if frac > 0.6 else "#405060"
            sc.create_rectangle(bx, 6, bx+fw, 16, fill=col, outline="")

    # ── misc ──────────────────────────────────────────────────────────

    def _set_status(self, text, color=None):
        try:
            self._nav_status.config(
                text=f"  {text}", fg=color or DARK["text_dim"])
        except Exception:
            pass

    def destroy(self):
        self._index_cancel.set()
        if self._prerender:
            self._prerender.stop()
        self._thumb_pool.shutdown(wait=False)


# =====================================================================
#  PDF REPAIR ENGINE  —  pikepdf-based XREF rebuild + object surgery
# =====================================================================

# PDF permission bit positions (ISO 32000-1 Table 22)
_PERM_BITS = [
    (2,  "Print (low quality)"),
    (3,  "Modify contents"),
    (4,  "Copy / extract text"),
    (5,  "Add / modify annotations"),
    (8,  "Fill interactive form fields"),
    (9,  "Extract for accessibility"),
    (10, "Assemble document"),
    (11, "Print (high quality)"),
]

# All known PDF stream filter names
_STREAM_FILTERS = [
    "FlateDecode", "DCTDecode", "JPXDecode", "CCITTFaxDecode",
    "JBIG2Decode", "LZWDecode", "RunLengthDecode", "ASCII85Decode",
    "ASCIIHexDecode", "Crypt",
]


class PDFRepairEngine:
    """XREF rebuild, corruption triage, and structural repair via pikepdf.

    pikepdf calls libqpdf under the hood.  When you open a PDF with
    suppress_warnings=True and attempt_recovery=True, libqpdf scans
    the entire binary stream looking for ``N 0 obj`` headers and
    rebuilds the XREF from scratch — recovering files that Acrobat,
    Chrome and even Ghostscript reject with "File corrupted".

    Extra repairs performed here beyond raw pikepdf open:
      • Linearisation removal  (linearised PDFs are harder to edit)
      • Object stream flattening  (exposes buried objects for inspection)
      • Orphan object detection    (objects not reachable from the catalog)
      • Truncation detection       (EOF marker missing or misplaced)
    """

    def __init__(self, path):
        self.path      = path
        self.pdf       = None        # pikepdf.Pdf once opened
        self.warnings  = []
        self.n_objects = 0
        self.n_orphans = 0
        self.xref_ok   = False
        self.linearised= False
        self.encrypted = False
        self.version   = ""
        self._opened   = False

    def open(self, password=""):
        """Open with full recovery.  Returns (success, message)."""
        if not PIKEPDF_AVAILABLE:
            return False, "pikepdf not installed (pip install pikepdf)"
        import pikepdf
        _PwErr = getattr(pikepdf, "PasswordError",
                 getattr(getattr(pikepdf, "_core", pikepdf),
                         "PasswordError", Exception))
        try:
            kwargs = {"suppress_warnings": False, "attempt_recovery": True}
            if password:
                kwargs["password"] = password
            self.pdf       = pikepdf.open(self.path, **kwargs)
            self.warnings  = [str(w) for w in self.pdf.get_warnings()]
            self.n_objects = len(self.pdf.objects)
            self.version   = str(self.pdf.pdf_version)
            try:
                self.linearised = "/Linearized" in self.pdf.Root
            except Exception:
                self.linearised = False
            self.encrypted = self.pdf.is_encrypted
            self.xref_ok   = len(self.warnings) == 0
            self._opened   = True
            return True, f"Opened OK  ·  v{self.version}  ·  {self.n_objects} objects"
        except _PwErr:
            return False, "Password required or incorrect"
        except Exception as exc:
            return False, f"Recovery failed: {exc}"

    def triage(self):
        """Return a list of (severity, message) tuples describing the file."""
        if not self._opened:
            return [("error", "File not opened")]
        results = []
        if self.warnings:
            for w in self.warnings:
                results.append(("warn", f"XREF warning: {w}"))
        else:
            results.append(("ok", "XREF table intact — no warnings"))

        if self.linearised:
            results.append(("info", "Linearised (web-optimised) — will be removed on save"))
        if self.encrypted:
            results.append(("warn", "Encrypted — permission flags may be enforced"))

        # check for truncation: last object number vs max xref entry
        try:
            max_obj = max(int(str(o.objgen[0])) for o in self.pdf.objects
                          if o is not None)
            if max_obj > self.n_objects + 50:
                results.append(("warn",
                    f"Gap detected: max obj# {max_obj} but only "
                    f"{self.n_objects} objects — possible truncation"))
        except Exception:
            pass

        # orphan detection: walk catalog and mark reachable
        try:
            reachable = set()
            def _walk(obj, depth=0):
                if depth > 40:
                    return
                try:
                    key = str(getattr(obj, "objgen", id(obj)))
                except Exception:
                    return
                if key in reachable:
                    return
                reachable.add(key)
                try:
                    if hasattr(obj, "items"):
                        for _, v in obj.items():
                            _walk(v, depth + 1)
                    elif hasattr(obj, "__iter__") and not isinstance(
                            obj, (str, bytes)):
                        for v in obj:
                            _walk(v, depth + 1)
                except Exception:
                    pass

            root = getattr(self.pdf, "Root", None)
            if root is not None:
                _walk(root)
            self.n_orphans = max(0, self.n_objects - len(reachable))
            if self.n_orphans > 0:
                results.append(("info",
                    f"{self.n_orphans} orphan objects "
                    f"(not reachable from catalog — safe to remove)"))
            else:
                results.append(("ok", "No orphan objects"))
        except Exception as exc:
            results.append(("info", f"Orphan scan skipped: {exc}"))

        return results

    def save_repaired(self, out_path, linearise=False,
                      remove_orphans=False):
        """Save a repaired, flattened copy."""
        if not self._opened:
            raise RuntimeError("File not opened")
        import pikepdf
        save_opts = {
            "fix_metadata_version": True,
            "linearize": linearise,
            "object_stream_mode": pikepdf.ObjectStreamMode.generate,
            "compress_streams": True,
        }
        self.pdf.save(out_path, **save_opts)

    def get_permission_int(self):
        """Return the raw /P integer from the encryption dictionary (trailer)."""
        if not self._opened or not self.pdf.is_encrypted:
            return None
        try:
            # /Encrypt lives in the trailer, NOT in the catalog Root
            enc = self.pdf.trailer.get("/Encrypt")
            if enc is not None:
                p = enc.get("/P")
                return int(p) if p is not None else -4
        except Exception:
            pass
        return None

    def set_permission_int(self, p_int):
        """Write a new /P value to the trailer encryption dictionary."""
        if not self._opened:
            return
        import pikepdf
        try:
            enc = self.pdf.trailer.get("/Encrypt")
            if enc is not None:
                enc["/P"] = pikepdf.Integer(p_int)
        except Exception:
            pass

    def get_info_dict(self):
        """Return {key: str} for all /Info entries."""
        if not self._opened:
            return {}
        result = {}
        try:
            info = self.pdf.docinfo
            for k in info:
                try:
                    result[str(k)] = str(info[k])
                except Exception:
                    result[str(k)] = "?"
        except Exception:
            pass
        return result

    def set_info_field(self, key, value):
        """Write a single /Info field."""
        if not self._opened:
            return
        try:
            self.pdf.docinfo[key] = value
        except Exception:
            pass

    def get_object_summary(self, max_objs=300):
        """Return list of (objnum, type_str, size_bytes) for inspection."""
        if not self._opened:
            return []
        import pikepdf
        rows = []
        for i, obj in enumerate(self.pdf.objects):
            if obj is None or i >= max_objs:
                continue
            try:
                gen     = obj.objgen
                objnum  = int(gen[0])
                t       = ""
                subtype = ""
                size    = 0
                try:
                    t       = str(obj.get("/Type",    ""))
                    subtype = str(obj.get("/Subtype", ""))
                except Exception:
                    pass
                try:
                    # read_raw_bytes() only exists on stream objects
                    if hasattr(obj, "read_raw_bytes"):
                        size = len(obj.read_raw_bytes())
                    elif obj.is_stream:
                        size = len(obj.read_bytes())
                except Exception:
                    pass
                label = f"{t} {subtype}".strip() or type(obj).__name__
                rows.append((objnum, label, size))
            except Exception:
                continue
        return sorted(rows, key=lambda r: r[0])

    def get_filter_summary(self):
        """Return {filter_name: count} across all stream objects."""
        if not self._opened:
            return {}
        counts = {}
        for obj in self.pdf.objects:
            if obj is None:
                continue
            try:
                f = obj.get("/Filter")
                if f is None:
                    continue
                names = [str(f)] if not hasattr(f, "__iter__") else [str(x) for x in f]
                for n in names:
                    n = n.lstrip("/")
                    counts[n] = counts.get(n, 0) + 1
            except Exception:
                continue
        return counts

    def close(self):
        if self.pdf:
            try:
                self.pdf.close()
            except Exception:
                pass
            self.pdf = None


# =====================================================================
#  FLAG INSPECTOR WIDGET  —  embedded inside PDFViewerPane
# =====================================================================

class PDFFlagPanel:
    """Collapsible bottom panel inside PDFViewerPane.

    Four tabs:
      XREF / Health  — triage results + repair button
      Permissions    — checkbox grid for all 8 permission bits
      Objects        — scrollable table of every PDF object
      Metadata       — /Info dict editor + filter map
    """

    PANEL_H = 220    # collapsed → 28, expanded → PANEL_H

    def __init__(self, parent, log_widget, viewer_pane):
        self._log    = log_widget
        self._viewer = viewer_pane
        self._engine = None          # PDFRepairEngine once opened
        self._perm_vars = {}         # {bit: tk.BooleanVar}
        self._expanded  = False

        # ── outer collapsible frame ───────────────────────────────────
        self.outer = tk.Frame(parent, bg=DARK["surface"])
        self.outer.pack(fill="x", padx=4, pady=(0, 4))

        # toggle header
        hdr = tk.Frame(self.outer, bg=DARK["surface"], height=28)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        self._toggle_lbl = tk.Label(
            hdr, text="▶  PDF Structure & Flags",
            bg=DARK["surface"], fg=DARK["accent"],
            font=("Consolas", 8, "bold"), anchor="w", cursor="hand2")
        self._toggle_lbl.pack(side="left", fill="x", expand=True, padx=8)
        self._toggle_lbl.bind("<Button-1>", lambda _e: self.toggle())

        self._health_badge = tk.Label(
            hdr, text="  no file  ",
            bg=DARK["surface"], fg=DARK["text_dim"],
            font=("Consolas", 7))
        self._health_badge.pack(side="right", padx=8)

        # collapsible body
        self._body = tk.Frame(self.outer, bg=DARK["bg"],
                               height=self.PANEL_H)
        # body is NOT packed initially — toggled on demand

        # notebook inside body
        self._nb = ttk.Notebook(self._body)
        self._nb.pack(fill="both", expand=True, padx=4, pady=4)

        self._build_xref_tab()
        self._build_perm_tab()
        self._build_objects_tab()
        self._build_meta_tab()

    # ── collapse / expand ────────────────────────────────────────────

    def toggle(self):
        self._expanded = not self._expanded
        if self._expanded:
            self._body.pack(fill="x")
            self.outer.configure(height=self.PANEL_H + 28)
            self._toggle_lbl.config(text="▼  PDF Structure & Flags")
        else:
            self._body.pack_forget()
            self._toggle_lbl.config(text="▶  PDF Structure & Flags")

    # ── XREF / Health tab ────────────────────────────────────────────

    def _build_xref_tab(self):
        tab = ttk.Frame(self._nb)
        self._nb.add(tab, text="⚕ Health / XREF")

        ctrl = ttk.Frame(tab)
        ctrl.pack(fill="x", padx=6, pady=4)

        ttk.Button(ctrl, text="🔧 Repair & Save copy",
                   command=self._do_repair).pack(side="left", padx=(0, 6))
        ttk.Button(ctrl, text="⟳ Re-triage",
                   command=self._do_triage).pack(side="left", padx=(0, 6))

        self._pw_entry = ttk.Entry(ctrl, width=18, show="*")
        self._pw_entry.pack(side="left", padx=(12, 2))
        ttk.Label(ctrl, text="password (if encrypted)",
                  style="Dim.TLabel").pack(side="left")

        # triage results list
        sf = ScrollableFrame(tab)
        sf.pack(fill="both", expand=True, padx=6)
        self._triage_inner = sf.inner

    def _do_triage(self):
        if not self._engine or not self._engine._opened:
            return
        for w in self._triage_inner.winfo_children():
            w.destroy()
        results = self._engine.triage()
        icons = {"ok": ("✓", DARK["success"]),
                 "warn": ("⚠", DARK["warning"]),
                 "info": ("ℹ", DARK["accent2"]),
                 "error": ("✗", DARK["danger"])}
        for sev, msg in results:
            icon, col = icons.get(sev, ("·", DARK["text_dim"]))
            row = tk.Frame(self._triage_inner, bg=DARK["bg"])
            row.pack(fill="x", pady=1)
            tk.Label(row, text=icon, bg=DARK["bg"], fg=col,
                     font=("Consolas", 9, "bold"), width=2).pack(side="left")
            tk.Label(row, text=msg, bg=DARK["bg"], fg=DARK["text"],
                     font=("Segoe UI", 8), anchor="w").pack(
                         side="left", fill="x", expand=True)
        # update health badge
        worst = "ok"
        for sev, _ in results:
            if sev == "error":  worst = "error"; break
            if sev == "warn":   worst = "warn"
        badge_txt = {"ok": " ✓ Healthy ", "warn": " ⚠ Warnings ",
                     "error": " ✗ Errors "}
        badge_col = {"ok": DARK["success"], "warn": DARK["warning"],
                     "error": DARK["danger"]}
        self._health_badge.config(
            text=badge_txt.get(worst, ""),
            fg=badge_col.get(worst, DARK["text_dim"]))

    def _do_repair(self):
        if not self._engine or not self._engine._opened:
            messagebox.showwarning("No file", "Open a PDF first.")
            return
        src = self._engine.path
        out = src.replace(".pdf", "_repaired.pdf")
        out = filedialog.asksaveasfilename(
            initialfile=os.path.basename(out),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        try:
            self._engine.save_repaired(out)
            log_append(self._log, f"Repaired PDF saved: {out}", "success")
            messagebox.showinfo("Saved", f"Repaired PDF:\n{out}")
        except Exception as exc:
            messagebox.showerror("Repair failed", str(exc))

    # ── Permissions tab ──────────────────────────────────────────────

    def _build_perm_tab(self):
        tab = ttk.Frame(self._nb)
        self._nb.add(tab, text="🔒 Permissions")

        info = tk.Label(
            tab,
            text="  Permission flags live in the /P integer of the "
                 "encryption dictionary.\n"
                 "  Changing them here and saving removes the enforcement "
                 "— only works on owner-level opens.",
            bg=DARK["bg"], fg=DARK["text_dim"],
            font=("Segoe UI", 8), justify="left", anchor="w")
        info.pack(fill="x", padx=6, pady=(6, 4))

        grid = tk.Frame(tab, bg=DARK["bg"])
        grid.pack(fill="x", padx=6)

        for i, (bit, desc) in enumerate(_PERM_BITS):
            var = tk.BooleanVar(value=True)
            self._perm_vars[bit] = var
            row, col = divmod(i, 2)
            cb = tk.Checkbutton(
                grid, text=f"Bit {bit:2d}  {desc}",
                variable=var,
                bg=DARK["bg"], fg=DARK["text"],
                selectcolor=DARK["surface"],
                activebackground=DARK["bg"],
                activeforeground=DARK["accent"],
                font=("Segoe UI", 8),
            )
            cb.grid(row=row, column=col, sticky="w", padx=12, pady=2)

        btn_f = ttk.Frame(tab)
        btn_f.pack(fill="x", padx=6, pady=6)
        ttk.Button(btn_f, text="✓ Unlock All",
                   command=self._perm_unlock_all).pack(side="left", padx=(0,4))
        ttk.Button(btn_f, text="✗ Lock All",
                   command=self._perm_lock_all).pack(side="left", padx=(0,4))
        ttk.Button(btn_f, text="💾 Apply & Save copy",
                   command=self._perm_apply).pack(side="left")

        self._perm_raw_lbl = tk.Label(
            tab, text="  /P = n/a",
            bg=DARK["bg"], fg=DARK["text_dim"],
            font=("Consolas", 8))
        self._perm_raw_lbl.pack(anchor="w", padx=6)

    def _load_permissions(self):
        p = self._engine.get_permission_int() if self._engine else None
        if p is None:
            for var in self._perm_vars.values():
                var.set(True)
            self._perm_raw_lbl.config(text="  /P = not encrypted")
            return
        for bit, var in self._perm_vars.items():
            var.set(bool(p & (1 << (bit - 1))))
        self._perm_raw_lbl.config(
            text=f"  /P = {p}  (0x{p & 0xFFFFFFFF:08X})  binary={p:032b}"[-60:])

    def _perm_unlock_all(self):
        for var in self._perm_vars.values():
            var.set(True)

    def _perm_lock_all(self):
        for var in self._perm_vars.values():
            var.set(False)

    def _perm_apply(self):
        if not self._engine or not self._engine._opened:
            return
        # build new /P integer — low bits 0-1 must always be 0,
        # bits 2-11 are the permission bits, upper bits set to 1
        new_p = -4   # 0xFFFFFFFC base (bits 0-1 = 0, rest = 1)
        for bit, var in self._perm_vars.items():
            if var.get():
                new_p |=  (1 << (bit - 1))
            else:
                new_p &= ~(1 << (bit - 1))
        out = filedialog.asksaveasfilename(
            initialfile=os.path.basename(
                self._engine.path).replace(".pdf", "_flags.pdf"),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        try:
            self._engine.set_permission_int(new_p)
            self._engine.save_repaired(out)
            log_append(self._log,
                       f"Flags written: /P={new_p}  saved → {out}", "success")
            messagebox.showinfo("Saved",
                                f"Permission flags updated.\n/P = {new_p}\n\n{out}")
            self._load_permissions()
        except Exception as exc:
            messagebox.showerror("Failed", str(exc))

    # ── Objects tab ──────────────────────────────────────────────────

    def _build_objects_tab(self):
        tab = ttk.Frame(self._nb)
        self._nb.add(tab, text="🗂 Objects")

        ctrl = ttk.Frame(tab)
        ctrl.pack(fill="x", padx=6, pady=4)
        self._obj_filter = ttk.Entry(ctrl, width=22)
        self._obj_filter.pack(side="left", padx=(0, 6))
        self._obj_filter.insert(0, "filter by type…")
        self._obj_filter.bind("<FocusIn>",
            lambda _e: self._obj_filter.delete(0, "end")
                       if self._obj_filter.get() == "filter by type…"
                       else None)
        ttk.Button(ctrl, text="🔍 Filter",
                   command=self._obj_apply_filter).pack(side="left")
        self._obj_count_lbl = tk.Label(ctrl, text="",
                                        bg=DARK["bg"], fg=DARK["text_dim"],
                                        font=("Segoe UI", 8))
        self._obj_count_lbl.pack(side="left", padx=8)

        # treeview
        cols = ("objnum", "type", "size")
        self._obj_tree = ttk.Treeview(
            tab, columns=cols, show="headings", height=6)
        for col, hdr, w in zip(cols,
                                ("Obj #", "Type / Subtype", "Stream bytes"),
                                (70, 240, 100)):
            self._obj_tree.heading(col, text=hdr,
                                    command=lambda c=col: self._obj_sort(c))
            self._obj_tree.column(col, width=w, anchor="w")
        self._obj_tree.pack(fill="both", expand=True, padx=6, pady=(0, 4))

        obj_sb = ttk.Scrollbar(tab, orient="vertical",
                                command=self._obj_tree.yview)
        self._obj_tree.configure(yscrollcommand=obj_sb.set)

        self._obj_rows = []    # full list, cached for filtering

    def _load_objects(self):
        self._obj_rows = self._engine.get_object_summary() \
            if self._engine and self._engine._opened else []
        self._obj_apply_filter()

    def _obj_apply_filter(self):
        q = self._obj_filter.get().strip().lower()
        if q in ("", "filter by type…"):
            rows = self._obj_rows
        else:
            rows = [r for r in self._obj_rows if q in r[1].lower()]
        self._obj_tree.delete(*self._obj_tree.get_children())
        for objnum, typ, size in rows:
            size_str = f"{size:,}" if size else "—"
            self._obj_tree.insert("", "end",
                                   values=(objnum, typ, size_str))
        self._obj_count_lbl.config(
            text=f"{len(rows)} / {len(self._obj_rows)} objects")

    def _obj_sort(self, col):
        items = [(self._obj_tree.set(k, col), k)
                 for k in self._obj_tree.get_children()]
        try:
            items.sort(key=lambda t: int(t[0].replace(",", "")))
        except ValueError:
            items.sort()
        for idx, (_, k) in enumerate(items):
            self._obj_tree.move(k, "", idx)

    # ── Metadata tab ─────────────────────────────────────────────────

    def _build_meta_tab(self):
        tab = ttk.Frame(self._nb)
        self._nb.add(tab, text="📋 Metadata")

        top = ttk.Frame(tab)
        top.pack(fill="x", padx=6, pady=4)
        ttk.Button(top, text="💾 Save changes",
                   command=self._meta_save).pack(side="left", padx=(0, 6))
        ttk.Button(top, text="✕ Clear all metadata",
                   command=self._meta_clear).pack(side="left")

        # filter map
        fmap_lbl = tk.Label(tab, text="  Stream filters:",
                             bg=DARK["bg"], fg=DARK["text_dim"],
                             font=("Consolas", 8))
        fmap_lbl.pack(anchor="w", padx=6)
        self._fmap_lbl = tk.Label(tab, text="",
                                   bg=DARK["bg"], fg=DARK["accent2"],
                                   font=("Consolas", 8), anchor="w")
        self._fmap_lbl.pack(fill="x", padx=14)

        ttk.Separator(tab, orient="horizontal").pack(
            fill="x", padx=6, pady=4)

        # /Info fields grid
        self._meta_entries = {}
        fields = ["/Title", "/Author", "/Subject", "/Keywords",
                  "/Creator", "/Producer", "/CreationDate", "/ModDate"]
        meta_grid = ttk.Frame(tab)
        meta_grid.pack(fill="x", padx=6)
        meta_grid.columnconfigure(1, weight=1)
        for row, key in enumerate(fields):
            tk.Label(meta_grid, text=f"{key}:",
                     bg=DARK["bg"], fg=DARK["text_dim"],
                     font=("Consolas", 8), width=14, anchor="e").grid(
                         row=row, column=0, sticky="e", padx=(0, 6),
                         pady=1)
            ent = ttk.Entry(meta_grid)
            ent.grid(row=row, column=1, sticky="ew", pady=1)
            self._meta_entries[key] = ent

    def _load_metadata(self):
        info = self._engine.get_info_dict() if self._engine and \
               self._engine._opened else {}
        for key, ent in self._meta_entries.items():
            ent.delete(0, "end")
            ent.insert(0, info.get(key, ""))

        fmap = self._engine.get_filter_summary() if self._engine and \
               self._engine._opened else {}
        self._fmap_lbl.config(
            text="  ".join(f"{k}: {v}" for k, v in sorted(fmap.items()))
                 or "(none)")

    def _meta_save(self):
        if not self._engine or not self._engine._opened:
            return
        for key, ent in self._meta_entries.items():
            val = ent.get().strip()
            if val:
                self._engine.set_info_field(key, val)
        out = filedialog.asksaveasfilename(
            initialfile=os.path.basename(
                self._engine.path).replace(".pdf", "_meta.pdf"),
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")])
        if not out:
            return
        try:
            self._engine.save_repaired(out)
            log_append(self._log, f"Metadata saved → {out}", "success")
            messagebox.showinfo("Saved", f"Metadata written.\n{out}")
        except Exception as exc:
            messagebox.showerror("Failed", str(exc))

    def _meta_clear(self):
        for ent in self._meta_entries.values():
            ent.delete(0, "end")

    # ── Called by PDFViewerPane on open ──────────────────────────────

    def load_pdf(self, path, password=""):
        """Open or re-open a PDF into the engine.  Called from viewer."""
        if self._engine:
            self._engine.close()
        self._engine = PDFRepairEngine(path)
        pw = password or self._pw_entry.get().strip()
        ok, msg = self._engine.open(pw)
        log_append(self._log,
                   f"FlagPanel: {msg}", "success" if ok else "warn")
        if ok:
            self._do_triage()
            self._load_permissions()
            self._load_objects()
            self._load_metadata()
        else:
            self._health_badge.config(
                text=" ⚠ open failed ", fg=DARK["danger"])



class PreviewPane:
    """Side-by-side before/after PDF preview with zoom/pan magnifier.

    Fixes vs previous version
    ─────────────────────────
    • Jitter eliminated: _on_canvas_resize only re-fits the already-cached
      PIL images (instant, main thread).  Full re-renders only happen when
      slider values change or the user clicks Refresh.
    • PIL images are cached as _pil_b / _pil_a so resize is O(1).
    • Zoom/pan:
        - Mouse wheel over either canvas → zoom in/out (10 % per tick)
        - Click-drag → pan
        - Double-click → reset zoom & pan
        - Ctrl+scroll → zoom the OTHER canvas (compare at different zooms)
    • Zoom indicator label shown under each canvas.
    • Busy-guard prevents stale renders overwriting a newer one.
    """

    RENDER_DPI = 150
    CANVAS_H   = 370

    def __init__(self, parent, log_widget):
        self.log = log_widget

        self._src       = None
        self._n_pages   = 1
        self._before_fn = None
        self._after_fn  = None

        # cached PIL images (set in _finish, read in _redraw_cached)
        self._pil_b = None
        self._pil_a = None
        # cached PhotoImages (kept to prevent GC)
        self._photo_b = None
        self._photo_a = None

        self._render_id  = None
        self._busy       = False
        self._render_gen = 0   # generation counter — stale renders are dropped

        # per-canvas zoom / pan state  { canvas: {"zoom": float, "ox": int, "oy": int} }
        self._vp = {}

        # ── outer frame ───────────────────────────────────────────────
        self.frame = ttk.LabelFrame(parent, text="  ◈  Live Preview")

        # ── top control bar ───────────────────────────────────────────
        ctrl = tk.Frame(self.frame, bg=DARK["panel"])
        ctrl.pack(fill="x", padx=6, pady=(4, 0))

        tk.Label(ctrl, text="Page:", bg=DARK["panel"],
                 fg=DARK["text_dim"], font=("Segoe UI", 8)).pack(side="left")

        self._page_var = tk.IntVar(value=1)
        self._spin = ttk.Spinbox(ctrl, from_=1, to=9999, width=4,
                                  textvariable=self._page_var,
                                  command=self.refresh)
        self._spin.pack(side="left", padx=(4, 2))
        self._spin.bind("<Return>", lambda _e: self.refresh())

        self._pg_lbl = tk.Label(ctrl, text="/ –", bg=DARK["panel"],
                                 fg=DARK["text_dim"], font=("Segoe UI", 8))
        self._pg_lbl.pack(side="left")

        self._status_lbl = tk.Label(ctrl, text="  Select a PDF to preview",
                                     bg=DARK["panel"], fg=DARK["text_dim"],
                                     font=("Consolas", 8))
        self._status_lbl.pack(side="left", padx=10)

        ttk.Button(ctrl, text="⟳ Refresh",
                   command=self.refresh).pack(side="right", padx=4)

        # zoom hint
        tk.Label(ctrl,
                 text="scroll=zoom  ·  drag=pan  ·  dbl-click=reset",
                 bg=DARK["panel"], fg=DARK["text_dim"],
                 font=("Segoe UI", 7)).pack(side="right", padx=6)

        # ── canvas area ───────────────────────────────────────────────
        ca = tk.Frame(self.frame, bg=DARK["bg"])
        ca.pack(fill="both", expand=True, padx=6, pady=(4, 6))
        ca.columnconfigure(0, weight=1)
        ca.columnconfigure(1, weight=1)
        ca.rowconfigure(0, weight=1)

        def _make_panel(col, label):
            wrap = tk.Frame(ca, bg=DARK["border"], bd=1, relief="flat")
            wrap.grid(row=0, column=col, sticky="nsew",
                      padx=(0, 3) if col == 0 else (3, 0))
            wrap.rowconfigure(1, weight=1)
            wrap.columnconfigure(0, weight=1)

            tk.Label(wrap, text=f"  {label}  ",
                     bg=DARK["surface"], fg=DARK["accent"],
                     font=("Consolas", 9, "bold")).grid(
                         row=0, column=0, sticky="ew")

            cnv = tk.Canvas(wrap, bg=DARK["log_bg"],
                            highlightthickness=0, height=self.CANVAS_H,
                            cursor="crosshair")
            cnv.grid(row=1, column=0, sticky="nsew")

            # zoom label
            zlbl = tk.Label(wrap, text="100%",
                            bg=DARK["surface"], fg=DARK["text_dim"],
                            font=("Consolas", 8))
            zlbl.grid(row=2, column=0, sticky="ew")

            # init viewport state
            self._vp[cnv] = {"zoom": 1.0, "ox": 0, "oy": 0,
                              "drag_x": None, "drag_y": None,
                              "zlbl": zlbl}

            # ── zoom / pan bindings ───────────────────────────────────
            cnv.bind("<MouseWheel>",
                     lambda e, c=cnv: self._wheel(e, c))
            cnv.bind("<Button-4>",
                     lambda e, c=cnv: self._wheel(e, c))
            cnv.bind("<Button-5>",
                     lambda e, c=cnv: self._wheel(e, c))
            cnv.bind("<ButtonPress-1>",
                     lambda e, c=cnv: self._drag_start(e, c))
            cnv.bind("<B1-Motion>",
                     lambda e, c=cnv: self._drag_move(e, c))
            cnv.bind("<ButtonRelease-1>",
                     lambda e, c=cnv: self._drag_end(e, c))
            cnv.bind("<Double-Button-1>",
                     lambda e, c=cnv: self._zoom_reset(c))
            cnv.bind("<Configure>",
                     lambda e, c=cnv: self._on_canvas_configure(e, c))

            return cnv

        self._cnv_b = _make_panel(0, "BEFORE")
        self._cnv_a = _make_panel(1, "AFTER")

        self._placeholder(self._cnv_b, "Before")
        self._placeholder(self._cnv_a, "After")

    # ── zoom / pan internals ─────────────────────────────────────────

    def _wheel(self, event, cnv):
        vp = self._vp[cnv]
        if event.num == 4:        factor = 1.10
        elif event.num == 5:      factor = 1 / 1.10
        elif event.delta > 0:     factor = 1.10
        else:                     factor = 1 / 1.10

        # zoom towards the cursor position
        mx, my = event.x, event.y
        new_zoom = max(0.1, min(10.0, vp["zoom"] * factor))
        scale    = new_zoom / vp["zoom"]
        vp["ox"] = int(mx - scale * (mx - vp["ox"]))
        vp["oy"] = int(my - scale * (my - vp["oy"]))
        vp["zoom"] = new_zoom
        self._redraw_one(cnv)

    def _drag_start(self, event, cnv):
        vp = self._vp[cnv]
        vp["drag_x"] = event.x
        vp["drag_y"] = event.y
        cnv.configure(cursor="fleur")

    def _drag_move(self, event, cnv):
        vp = self._vp[cnv]
        if vp["drag_x"] is None:
            return
        dx = event.x - vp["drag_x"]
        dy = event.y - vp["drag_y"]
        vp["ox"] += dx
        vp["oy"] += dy
        vp["drag_x"] = event.x
        vp["drag_y"] = event.y
        self._redraw_one(cnv)

    def _drag_end(self, event, cnv):
        vp = self._vp[cnv]
        vp["drag_x"] = vp["drag_y"] = None
        cnv.configure(cursor="crosshair")

    def _zoom_reset(self, cnv):
        vp = self._vp[cnv]
        vp["zoom"] = 1.0
        vp["ox"]   = 0
        vp["oy"]   = 0
        self._redraw_one(cnv)

    def _on_canvas_configure(self, event, cnv):
        """Window resize: re-fit the CACHED image without re-rendering."""
        # reset pan/zoom so image stays centred after resize
        vp = self._vp[cnv]
        vp["ox"] = 0
        vp["oy"] = 0
        pil = self._pil_b if cnv is self._cnv_b else self._pil_a
        if pil is not None:
            self._redraw_one(cnv, force_pil=pil)
        # do NOT call schedule_refresh — that would re-render

    # ── drawing helpers ──────────────────────────────────────────────

    def _placeholder(self, cnv, label=""):
        cnv.delete("all")
        w = max(cnv.winfo_width(), 80)
        h = max(cnv.winfo_height(), 40)
        cnv.create_text(w // 2, h // 2, anchor="center",
                        text=f"{'No preview — ' + label if label else 'No preview'}\n"
                             "select a PDF to begin",
                        fill=DARK["text_dim"], font=("Consolas", 9),
                        justify="center")

    def _set_status(self, text, color=None):
        try:
            self._status_lbl.config(
                text=f"  {text}", fg=color or DARK["text_dim"])
        except Exception:
            pass

    def _redraw_one(self, cnv, force_pil=None):
        """Refit + zoom/pan a cached PIL image onto cnv. Main thread only."""
        pil = force_pil or (
            self._pil_b if cnv is self._cnv_b else self._pil_a)
        if pil is None:
            return

        vp  = self._vp[cnv]
        cw  = max(cnv.winfo_width(),  120)
        ch  = max(cnv.winfo_height(), self.CANVAS_H)

        # fit to canvas first, then apply zoom
        base_scale = min(cw / pil.width, ch / pil.height, 1.0)
        total_zoom = base_scale * vp["zoom"]
        nw = max(1, int(pil.width  * total_zoom))
        nh = max(1, int(pil.height * total_zoom))

        resample = Image.LANCZOS if total_zoom < 2.0 else Image.NEAREST
        try:
            resized = pil.resize((nw, nh), resample)
        except Exception:
            return

        try:
            from PIL import ImageTk
            photo = ImageTk.PhotoImage(resized)
        except Exception:
            return

        # store to prevent GC
        if cnv is self._cnv_b:
            self._photo_b = photo
        else:
            self._photo_a = photo

        x = cw // 2 + vp["ox"]
        y = ch // 2 + vp["oy"]
        cnv.delete("all")
        cnv.create_image(x, y, image=photo, anchor="center")

        # zoom label
        pct = int(vp["zoom"] * 100)
        vp["zlbl"].config(
            text=f"{pct}%  ({nw}×{nh}px)",
            fg=DARK["accent2"] if pct != 100 else DARK["text_dim"])

    def _redraw_cached(self):
        """Re-fit both cached PIL images without re-rendering. Instant."""
        if self._pil_b is not None:
            self._zoom_reset(self._cnv_b)
        if self._pil_a is not None:
            self._zoom_reset(self._cnv_a)

    # ── Public API ───────────────────────────────────────────────────

    def set_source(self, src):
        self._src = src
        if src and os.path.exists(src):
            try:
                self._n_pages = count_pdf_pages(src)
                self._pg_lbl.config(text=f"/ {self._n_pages}")
                if self._page_var.get() > self._n_pages:
                    self._page_var.set(self._n_pages)
            except Exception:
                self._n_pages = 1
                self._pg_lbl.config(text="/ ?")

    def set_before_fn(self, fn):
        self._before_fn = fn

    def set_after_fn(self, fn):
        self._after_fn = fn

    def schedule_refresh(self, delay_ms=500):
        """Debounced re-render. Safe to call on every slider tick."""
        if self._render_id is not None:
            try:
                self.frame.after_cancel(self._render_id)
            except Exception:
                pass
        self._render_id = self.frame.after(delay_ms, self.refresh)

    def refresh(self):
        """Launch a background render of both panels."""
        if not self._src or not os.path.exists(self._src):
            self._set_status("No PDF selected", DARK["warning"])
            return
        if self._busy:
            # Don't stack renders — just reschedule once
            self.schedule_refresh(350)
            return

        self._busy = True
        self._render_gen += 1
        gen = self._render_gen
        self._set_status("Rendering …", DARK["accent2"])
        page = max(1, min(self._page_var.get(), self._n_pages))

        before_fn = self._before_fn
        after_fn  = self._after_fn
        src       = self._src

        def _worker():
            b_img = a_img = None
            try:
                if before_fn:
                    b_img = before_fn(src, page)
            except Exception as exc:
                self.frame.after(0, lambda: self._set_status(
                    f"Before: {exc}", DARK["danger"]))
            try:
                if after_fn:
                    a_img = after_fn(src, page)
            except Exception as exc:
                self.frame.after(0, lambda: self._set_status(
                    f"After: {exc}", DARK["danger"]))

            # drop stale result if a newer render was already started
            self.frame.after(0, lambda: self._finish(b_img, a_img, gen))

        threading.Thread(target=_worker, daemon=True).start()

    def _finish(self, b_img, a_img, gen):
        self._busy = False
        if gen != self._render_gen:
            return   # stale — discard silently
        self._pil_b = b_img
        self._pil_a = a_img
        # reset zoom/pan so new image fills canvas cleanly
        for cnv in (self._cnv_b, self._cnv_a):
            vp = self._vp[cnv]
            vp["zoom"] = 1.0
            vp["ox"]   = 0
            vp["oy"]   = 0
        self._redraw_one(self._cnv_b, force_pil=b_img)
        self._redraw_one(self._cnv_a, force_pil=a_img)
        self._set_status("Ready  ·  scroll to zoom  ·  drag to pan  ·  dbl-click to reset",
                         DARK["success"])




# =====================================================================
#  DARK SCROLLED TEXT HELPER
# =====================================================================

def make_log(parent):
    """Build a dark-themed, color-tagged log widget."""
    st = scrolledtext.ScrolledText(
        parent,
        height=10,
        state="disabled",
        bg=DARK["log_bg"],
        fg=DARK["log_fg"],
        font=("Consolas", 9),
        insertbackground=DARK["log_fg"],
        selectbackground=DARK["accent"],
        relief="flat",
        bd=0,
    )
    st.tag_config("ts",      foreground=DARK["text_dim"])
    st.tag_config("normal",  foreground=DARK["log_fg"])
    st.tag_config("success", foreground=DARK["success"])
    st.tag_config("warn",    foreground=DARK["warning"])
    st.tag_config("error",   foreground=DARK["danger"])
    st.tag_config("info",    foreground=DARK["accent2"])
    return st


# =====================================================================
#  REUSABLE WIDGETS
# =====================================================================

def make_file_row(parent, label, lb_opts, types=None, row=0):
    """A label + entry + browse button row. Returns the entry widget."""
    ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0,8), pady=4)
    entry = ttk.Entry(parent)
    entry.grid(row=row, column=1, sticky="ew", pady=4)

    def browse():
        if types:
            f = filedialog.askopenfilename(filetypes=types)
        else:
            f = filedialog.askopenfilename()
        if f:
            entry.delete(0, tk.END)
            entry.insert(0, f)

    ttk.Button(parent, text="Browse", command=browse).grid(row=row, column=2, padx=(8,0), pady=4)
    parent.columnconfigure(1, weight=1)
    return entry


def make_listbox(parent, lb_opts, height=10, scrollbar=True):
    frame = ttk.Frame(parent)
    lb = tk.Listbox(frame, height=height, **lb_opts)
    lb.pack(side="left", fill="both", expand=True)
    if scrollbar:
        sb = ttk.Scrollbar(frame, command=lb.yview)
        sb.pack(side="right", fill="y")
        lb.config(yscrollcommand=sb.set)
    return frame, lb


def make_progress(parent):
    pb = ttk.Progressbar(parent, mode="indeterminate", length=400)
    return pb


# =====================================================================
#  MAIN APPLICATION
# =====================================================================

class ConverterApp(TkBase):

    def __init__(self):
        super().__init__()
        self.title("Abyss Toolkit  ·  PC Edition")
        self.geometry("1060x820")
        self.minsize(800, 600)
        self.last_output = None
        self.lb_opts = apply_dark_theme(self)
        self._build_ui()
        self.after(300, self._post_init)
        # Global mousewheel handler — routes wheel events to whichever
        # scrollable canvas is under the cursor, regardless of focus.
        # This is the correct fix for "wheel only works on the scrollbar".
        self.bind_all("<MouseWheel>", self._global_wheel, add="+")
        self.bind_all("<Button-4>",   self._global_wheel, add="+")
        self.bind_all("<Button-5>",   self._global_wheel, add="+")

    def _global_wheel(self, event):
        """Route mousewheel to the nearest scrollable ancestor of event.widget."""
        if event.num == 4:    delta = -1
        elif event.num == 5:  delta =  1
        elif event.delta > 0: delta = -1
        else:                 delta =  1

        w = event.widget
        while w is not None:
            try:
                # ttk.Treeview has its own yview
                if isinstance(w, ttk.Treeview):
                    w.yview_scroll(delta, "units")
                    return
                # Any Canvas that has a yview scrollcommand configured
                if isinstance(w, tk.Canvas):
                    # check if it is actually scrollable (has content)
                    try:
                        lo, hi = w.yview()
                        if lo > 0.0001 or hi < 0.9999:
                            w.yview_scroll(delta, "units")
                            return
                        # even at top/bottom, try scrolling so user
                        # gets tactile feedback that this IS the scroll target
                        # only if scrollregion is configured
                        sr = w.cget("scrollregion")
                        if sr and sr != "":
                            w.yview_scroll(delta, "units")
                            return
                    except Exception:
                        pass
            except Exception:
                pass
            try:
                w = w.master
            except Exception:
                break

    # ── Layout skeleton ─────────────────────────────────────────────

    def _build_ui(self):
        # ── Header bar ──────────────────────────────────────────────
        hdr = tk.Frame(self, bg=DARK["panel"], height=48)
        hdr.pack(fill="x")
        hdr.pack_propagate(False)

        tk.Label(hdr, text="  ⬡  ABYSS TOOLKIT  ·  PC EDITION",
                 bg=DARK["panel"], fg=DARK["accent"],
                 font=("Consolas", 13, "bold")).pack(side="left", pady=10)

        self.stats_frame = tk.Frame(hdr, bg=DARK["panel"])
        self.stats_frame.pack(side="right", padx=14)

        self.lbl_cpu = tk.Label(self.stats_frame, text="CPU –", bg=DARK["panel"],
                                fg=DARK["text_dim"], font=("Consolas", 8))
        self.lbl_cpu.pack(side="left", padx=6)
        self.lbl_ram = tk.Label(self.stats_frame, text="RAM –", bg=DARK["panel"],
                                fg=DARK["text_dim"], font=("Consolas", 8))
        self.lbl_ram.pack(side="left", padx=6)

        sep = tk.Frame(self, bg=DARK["border"], height=1)
        sep.pack(fill="x")

        # ── Notebook ────────────────────────────────────────────────
        main_pane = tk.PanedWindow(
            self, orient="vertical", sashwidth=6, bg=DARK["border"], bd=0,
            sashrelief="flat")
        main_pane.pack(fill="both", expand=True, padx=6, pady=6)
        self.main_pane = main_pane

        nb = ttk.Notebook(main_pane)
        main_pane.add(nb, minsize=360)
        self.nb = nb

        self._tabs = {}
        tab_defs = [
            ("convert",  "⇄  Convert"),
            ("compress", "⊟  Compress"),
            ("restore",  "✦  Restore"),
            ("viewer",   "📄  Viewer"),
            ("split",    "✂  Split"),
            ("merge",    "⊕  Merge"),
            ("extract",  "⊞  Extract"),
            ("protect",  "⚿  Protect"),
            ("rotate",   "↻  Rotate"),
            ("crack",    "⚡  Crack"),
            ("tools",    "⚙  Tools"),
        ]
        for key, label in tab_defs:
            frame = ttk.Frame(nb)
            nb.add(frame, text=label)
            self._tabs[key] = frame

        # ── Log area ────────────────────────────────────────────────
        log_outer = tk.Frame(main_pane, bg=DARK["panel"])
        main_pane.add(log_outer, minsize=170)

        tk.Label(log_outer, text="  SYSTEM LOG", bg=DARK["panel"],
                 fg=DARK["accent"], font=("Consolas", 8, "bold")).pack(anchor="w")

        self.log_widget = make_log(log_outer)
        self.log_widget.pack(fill="both", expand=True, padx=4, pady=(2, 4))

        # ── Build each tab ──────────────────────────────────────────
        self._build_convert_tab()
        self._build_compress_tab()
        self._build_restore_tab()
        self._build_viewer_tab()
        self._build_split_tab()
        self._build_merge_tab()
        self._build_extract_tab()
        self._build_protect_tab()
        self._build_rotate_tab()
        self._build_crack_tab()
        self._build_tools_tab()

    def _post_init(self):
        log_append(self.log_widget, "Abyss Toolkit PC Edition initialised", "info")
        log_append(self.log_widget, f"Python {sys.version.split()[0]}  |  Platform: {sys.platform}", "info")
        self._check_deps()
        if PSUTIL_AVAILABLE:
            self._update_stats()

    def _check_deps(self):
        missing = []
        checks = [
            ("pandas",      pd is not None),
            ("PIL/Pillow",  PIL_AVAILABLE),
            ("reportlab",   REPORTLAB_AVAILABLE),
            ("pypdf",       PYPDF_AVAILABLE),
            ("pikepdf",     PIKEPDF_AVAILABLE),
            ("pdfplumber",  PDFPLUMBER_AVAILABLE),
            ("pdf2image",   PDF2IMAGE_AVAILABLE),
            ("openpyxl",    OPENPYXL_AVAILABLE),
            ("python-docx", DOCX_AVAILABLE),
            ("psutil",      PSUTIL_AVAILABLE),
        ]
        for name, ok in checks:
            if not ok:
                missing.append(name)
        if missing:
            log_append(self.log_widget,
                       f"MISSING packages: {', '.join(missing)}  "
                       f"(pip install {' '.join(missing)})", "warn")
        else:
            log_append(self.log_widget, "All Python packages present ✓", "success")

    def _update_stats(self):
        if not PSUTIL_AVAILABLE:
            return
        try:
            cpu = psutil.cpu_percent(interval=0.1)
            ram = psutil.virtual_memory().percent
            self.lbl_cpu.config(text=f"CPU {cpu:.0f}%",
                                fg=DARK["danger"] if cpu > 80 else DARK["text_dim"])
            self.lbl_ram.config(text=f"RAM {ram:.0f}%",
                                fg=DARK["danger"] if ram > 85 else DARK["text_dim"])
        except Exception:
            pass
        self.after(2000, self._update_stats)

    # ── helper: run in thread with progress bar ──────────────────────

    def _run(self, fn, pb=None):
        """Run fn() in a daemon thread, spinning pb if given."""
        def show_progress():
            if not pb:
                return
            manager = getattr(pb, "_progress_manager", "pack")
            try:
                if manager == "grid":
                    pb.grid()
                else:
                    pb.pack(fill="x", padx=4, pady=4)
                pb.start(12)
            except Exception as e:
                log_append(self.log_widget, f"Progress bar error: {e}", "warn")

        def hide_progress():
            if not pb:
                return
            manager = getattr(pb, "_progress_manager", "pack")
            try:
                pb.stop()
                if manager == "grid":
                    pb.grid_remove()
                else:
                    pb.pack_forget()
            except Exception:
                pass

        def worker():
            self.after(0, show_progress)
            try:
                fn()
            except Exception as e:
                log_append(self.log_widget, f"Unexpected worker error: {e}", "error")
                self._show_error("Error", str(e))
            finally:
                self.after(0, hide_progress)
        threading.Thread(target=worker, daemon=True).start()

    def _show_info(self, title, message):
        self._show_message(messagebox.showinfo, title, message)

    def _show_warning(self, title, message):
        self._show_message(messagebox.showwarning, title, message)

    def _show_error(self, title, message):
        self._show_message(messagebox.showerror, title, message)

    def _show_message(self, func, title, message):
        if threading.current_thread() is threading.main_thread():
            func(title, message)
        else:
            self.after(0, lambda: func(title, message))

    def _set_label(self, label, text, color=None):
        if threading.current_thread() is not threading.main_thread():
            self.after(0, lambda: self._set_label(label, text, color))
            return
        kwargs = {"text": text}
        if color:
            kwargs["foreground"] = color
        label.config(**kwargs)

    # ================================================================
    #  TAB: CONVERT
    # ================================================================

    def _build_convert_tab(self):
        tab = self._tabs["convert"]
        f   = ttk.Frame(tab)
        f.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(f, text="File Converter", style="Section.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        self._cv_input = make_file_row(f, "Input File:", self.lb_opts, row=1)

        # detected type badge
        ttk.Label(f, text="Detected:").grid(row=2, column=0, sticky="w", pady=4)
        self._cv_type = ttk.Label(f, text="–", foreground=DARK["accent2"])
        self._cv_type.grid(row=2, column=1, sticky="w")
        self._cv_input.bind("<FocusOut>",
            lambda e: self._cv_type.config(text=detect_file_type(self._cv_input.get()).upper()))

        ttk.Label(f, text="Output Format:").grid(row=3, column=0, sticky="w", pady=4)
        fmts = ["pdf", "png", "jpg", "xlsx", "csv", "txt"]
        self._cv_fmt = ttk.Combobox(f, values=fmts, state="readonly", width=10)
        self._cv_fmt.set("pdf")
        self._cv_fmt.grid(row=3, column=1, sticky="w")

        ttk.Label(f, text="Output Folder:").grid(row=4, column=0, sticky="w", pady=4)
        self._cv_outdir = ttk.Entry(f)
        self._cv_outdir.grid(row=4, column=1, sticky="ew", pady=4)
        ttk.Button(f, text="Pick", command=lambda: self._browse_dir(self._cv_outdir)).grid(
            row=4, column=2, padx=(8, 0))

        pb = make_progress(f)
        btn_frame = ttk.Frame(f)
        btn_frame.grid(row=5, column=0, columnspan=3, sticky="w", pady=10)

        ttk.Button(btn_frame, text="CONVERT", style="Accent.TButton",
                   command=lambda: self._start_convert(pb)).pack(side="left")
        ttk.Button(btn_frame, text="Open Folder",
                   command=lambda: open_path(self._cv_outdir.get() or self._cv_input.get())).pack(
                       side="left", padx=10)

        pb.grid(row=6, column=0, columnspan=3, sticky="ew")
        pb._progress_manager = "grid"
        pb.grid_remove()

        # results
        res_lf = ttk.LabelFrame(f, text="Output Files")
        res_lf.grid(row=7, column=0, columnspan=3, sticky="nsew", pady=8)
        f.rowconfigure(7, weight=1)
        f.columnconfigure(1, weight=1)

        _, self._cv_list = make_listbox(res_lf, self.lb_opts, height=8)
        self._cv_list.master.pack(fill="both", expand=True, padx=5, pady=5)

        # DnD support
        if DND_AVAILABLE:
            self._cv_input.drop_target_register(DND_FILES)
            self._cv_input.dnd_bind("<<Drop>>", lambda e: self._cv_input.insert(
                0, e.data.strip("{}")))

    def _start_convert(self, pb):
        src = self._cv_input.get().strip()
        if not src or not os.path.exists(src):
            self._show_error("Error", "Select a valid input file.")
            return
        fmt = self._cv_fmt.get().lower()
        out_dir = self._cv_outdir.get().strip() or os.path.dirname(src)
        name = safe_filename(src)
        ftype = detect_file_type(src)
        self._cv_type.config(text=ftype.upper())
        self._run(lambda: self._do_convert(src, fmt, out_dir, name, ftype), pb)

    def _do_convert(self, src, fmt, out_dir, name, ftype):
        log_append(self.log_widget, f"Converting  {ftype.upper()} → {fmt.upper()}  …")
        self.after(0, lambda: self._cv_list.delete(0, tk.END))
        outputs = []

        try:
            if ftype == "csv":
                if fmt == "pdf":   csv_to_pdf(src, f := os.path.join(out_dir, f"{name}.pdf"),   self.log_widget); outputs.append(f)
                elif fmt == "xlsx": csv_to_xlsx(src, f := os.path.join(out_dir, f"{name}.xlsx"), self.log_widget); outputs.append(f)
                elif fmt in ("png","jpg"): csv_to_image(src, f := os.path.join(out_dir, f"{name}.{fmt}"), self.log_widget); outputs.append(f)
            elif ftype in ("xlsx","xls"):
                if fmt == "pdf":   xlsx_to_pdf(src, f := os.path.join(out_dir, f"{name}.pdf"),   self.log_widget); outputs.append(f)
                elif fmt == "csv": xlsx_to_csv(src, f := os.path.join(out_dir, f"{name}.csv"),   self.log_widget); outputs.append(f)
                elif fmt in ("png","jpg"): xlsx_to_image(src, f := os.path.join(out_dir, f"{name}.{fmt}"), self.log_widget); outputs.append(f)
            elif ftype == "txt":
                if fmt == "pdf":   txt_to_pdf(src, f := os.path.join(out_dir, f"{name}.pdf"), self.log_widget); outputs.append(f)
            elif ftype in ("png","jpg","jpeg","bmp","gif","webp","tiff"):
                if fmt == "pdf":   image_to_pdf(src, f := os.path.join(out_dir, f"{name}.pdf"), self.log_widget); outputs.append(f)
                elif fmt in ("png","jpg","bmp"):
                    img = Image.open(src)
                    if fmt == "jpg" and img.mode == "RGBA": img = img.convert("RGB")
                    out = os.path.join(out_dir, f"{name}.{fmt}")
                    img.save(out); outputs.append(out)
            elif ftype == "docx":
                if fmt == "pdf":   docx_to_pdf(src, f := os.path.join(out_dir, f"{name}.pdf"), self.log_widget); outputs.append(f)
                elif fmt in ("png","jpg"):
                    imgs = docx_to_images(src, os.path.join(out_dir, f"{name}_imgs"), self.log_widget)
                    outputs.extend(imgs)
            elif ftype == "pdf":
                if fmt in ("png","jpg"):
                    imgs = pdf_to_images(src, os.path.join(out_dir, f"{name}_imgs"), log_widget=self.log_widget)
                    outputs.extend(imgs)
                else:
                    raise RuntimeError("PDF → PDF conversion is unnecessary.")
            else:
                raise RuntimeError(f"Unsupported input type: {ftype}")

            for o in outputs:
                self.last_output = o
                self.after(0, lambda o=o: self._cv_list.insert(tk.END, o))

            log_append(self.log_widget, f"Done — {len(outputs)} file(s) created", "success")
            self._show_info("Conversion Complete", f"{len(outputs)} file(s) created.")

        except Exception as e:
            log_append(self.log_widget, f"ERROR: {e}", "error")
            self._show_error("Conversion Failed", str(e))

    # ================================================================
    #  TAB: VIEWER  (HDC-powered PDF viewer)
    # ================================================================

    def _build_viewer_tab(self):
        tab = self._tabs["viewer"]
        self._viewer_pane = PDFViewerPane(tab, self.log_widget)
        self._viewer_pane.frame.pack(fill="both", expand=True)

    # ================================================================
    #  TAB: COMPRESS
    # ================================================================

    def _build_compress_tab(self):
        tab = self._tabs["compress"]

        pane = tk.PanedWindow(
            tab, orient="horizontal", sashwidth=6,
            bg=DARK["border"], bd=0, sashrelief="flat")
        pane.pack(fill="both", expand=True, padx=4, pady=4)

        ctrl_outer = ttk.Frame(pane)
        prev_f     = ttk.Frame(pane)
        pane.add(ctrl_outer, minsize=300, width=380)
        pane.add(prev_f,     minsize=340)

        # ── LEFT: scrollable control panel ───────────────────────────
        sf = ScrollableFrame(ctrl_outer)
        sf.pack(fill="both", expand=True)
        cf = sf.inner       # all controls go here

        ttk.Label(cf, text="Batch PDF Compression",
                  style="Section.TLabel").pack(anchor="w", padx=10,
                  pady=(10, 6))

        list_frame, self._cmp_lb = make_listbox(cf, self.lb_opts, height=6)
        list_frame.pack(fill="x", padx=10)

        btn_row = ttk.Frame(cf)
        btn_row.pack(fill="x", padx=10, pady=(4, 2))
        ttk.Button(btn_row, text="＋ Add PDFs",
                   command=self._cmp_add_files).pack(side="left", padx=(0, 4))
        ttk.Button(btn_row, text="✕ Clear",
                   command=lambda: self._cmp_lb.delete(0, tk.END)).pack(side="left")
        ttk.Label(btn_row, text="← click to preview",
                  style="Dim.TLabel").pack(side="right")

        # Engine selector
        eng_f = ttk.Frame(cf)
        eng_f.pack(fill="x", padx=10, pady=(10, 2))
        ttk.Label(eng_f, text="Engine:").pack(side="left")
        self._cmp_engine = ttk.Combobox(
            eng_f, values=["pikepdf (lossless)", "Ghostscript (lossy)"],
            state="readonly", width=24)
        self._cmp_engine.set("pikepdf (lossless)")
        self._cmp_engine.pack(side="left", padx=8)

        # GS DPI slider — fixed-width container stops cascade resize
        ttk.Separator(cf, orient="horizontal").pack(fill="x", padx=10, pady=8)
        ttk.Label(cf, text="GS Image DPI  (Ghostscript engine only)",
                  style="Dim.TLabel").pack(anchor="w", padx=10)

        self._cmp_dpi_var = tk.IntVar(value=150)
        self._cmp_dpi_lbl = tk.Label(
            cf, text="150 dpi — balanced",
            bg=DARK["bg"], fg=DARK["accent2"],
            font=("Consolas", 9, "bold"), anchor="w")
        self._cmp_dpi_lbl.pack(anchor="w", padx=10)

        # Use a fixed-width container so the Scale never forces a reflow
        dpi_container = tk.Frame(cf, bg=DARK["bg"], height=28)
        dpi_container.pack(fill="x", padx=10)
        dpi_container.pack_propagate(False)

        def _on_cmp_dpi(val):
            v = int(float(val))
            if v <= 72:    tag = "screen (tiny, very pixelated)"
            elif v <= 120: tag = "web / ebook"
            elif v <= 200: tag = "balanced"
            elif v <= 300: tag = "printer quality"
            else:          tag = "prepress / hi-res"
            self._cmp_dpi_lbl.config(text=f"{v} dpi — {tag}")
            self._cmp_preview.schedule_refresh(500)

        tk.Scale(dpi_container, from_=30, to=600, orient="horizontal",
                 variable=self._cmp_dpi_var, command=_on_cmp_dpi,
                 bg=DARK["bg"], fg=DARK["text"], troughcolor=DARK["surface"],
                 activebackground=DARK["accent"], highlightthickness=0,
                 bd=0, resolution=5, showvalue=False,
                 ).place(relx=0, rely=0, relwidth=1.0, relheight=1.0)

        hint_f = tk.Frame(cf, bg=DARK["bg"])
        hint_f.pack(fill="x", padx=10)
        tk.Label(hint_f, text="← smaller/pixelated",
                 bg=DARK["bg"], fg=DARK["text_dim"],
                 font=("Segoe UI", 7)).pack(side="left")
        tk.Label(hint_f, text="higher quality →",
                 bg=DARK["bg"], fg=DARK["text_dim"],
                 font=("Segoe UI", 7)).pack(side="right")

        ttk.Separator(cf, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # Output folder
        of_f = ttk.Frame(cf)
        of_f.pack(fill="x", padx=10, pady=4)
        ttk.Label(of_f, text="Output Folder:").pack(side="left")
        self._cmp_outdir = ttk.Entry(of_f)
        self._cmp_outdir.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(of_f, text="Save To…",
                   command=lambda: self._browse_dir(self._cmp_outdir)).pack(side="left")

        self._cmp_status = ttk.Label(
            cf,
            text="Output folder optional — blank saves beside each source PDF.",
            style="Dim.TLabel")
        self._cmp_status.pack(fill="x", padx=10, pady=(0, 4))

        pb = make_progress(cf)
        pb._progress_manager = "pack"

        def _do_compress():
            files = list(self._cmp_lb.get(0, tk.END))
            if not files:
                self._show_warning("Empty List", "Add PDF files first.")
                return
            engine  = self._cmp_engine.get()
            gs_dpi  = self._cmp_dpi_var.get()
            if gs_dpi <= 72:    gs_q = "screen"
            elif gs_dpi <= 150: gs_q = "ebook"
            elif gs_dpi <= 250: gs_q = "printer"
            else:               gs_q = "prepress"
            out_dir = self._cmp_outdir.get().strip()
            before_total   = total_existing_size(files)
            expected_total = estimate_compressed_size(before_total, engine, gs_q)
            mbps = 18 if "pikepdf" in engine else 6
            eta  = estimate_job_time(before_total, mbps)
            status = (f"Input: {human_size(before_total)}  |  "
                      f"Expected: ~{human_size(expected_total)}  |  ETA: ~{human_time(eta)}")
            self._set_label(self._cmp_status, status, DARK["accent2"])
            log_append(self.log_widget, status, "info")

            def worker():
                started = time.perf_counter()
                outputs = []
                for fp in files:
                    out = pdf_output_path(fp, "_compressed", out_dir)
                    try:
                        if "pikepdf" in engine:
                            compress_pdf_pikepdf(fp, out, self.log_widget)
                        else:
                            compress_pdf_ghostscript(fp, out, gs_q,
                                                     self.log_widget, dpi=gs_dpi)
                        outputs.append(out)
                        self.last_output = out
                    except Exception as e:
                        log_append(self.log_widget,
                                   f"Failed {os.path.basename(fp)}: {e}", "error")
                elapsed = time.perf_counter() - started
                after_total = total_existing_size(outputs)
                final = (f"Actual: {human_size(after_total)}  |  "
                         f"Original: {human_size(before_total)}  |  "
                         f"{size_delta(before_total, after_total)}  |  Time: {human_time(elapsed)}")
                self._set_label(self._cmp_status, final, DARK["success"])
                log_append(self.log_widget, final, "success")
                self._show_info("Done", f"Compression complete.\n\n{final}")
            self._run(worker, pb)

        ttk.Button(cf, text="START BATCH COMPRESS", style="Accent.TButton",
                   command=_do_compress).pack(anchor="w", padx=10, pady=(4, 4))
        pb.pack(fill="x", padx=12)
        pb.pack_forget()
        ttk.Label(cf,
                  text="pikepdf = lossless (DPI slider ignored)\n"
                       "Ghostscript = lossy, slider controls image DPI",
                  style="Dim.TLabel").pack(anchor="w", padx=10, pady=(0, 10))

        # ── RIGHT: preview pane ──────────────────────────────────────
        self._cmp_preview = PreviewPane(prev_f, self.log_widget)
        self._cmp_preview.frame.pack(fill="both", expand=True, padx=(4, 8), pady=10)

        self._cmp_preview.set_before_fn(
            lambda src, pg: render_pdf_page_pil(
                src, pg, PreviewPane.RENDER_DPI))

        def _cmp_after(src, pg):
            img = render_pdf_page_pil(src, pg, PreviewPane.RENDER_DPI)
            return simulate_gs_compression(img, self._cmp_dpi_var.get(),
                                           PreviewPane.RENDER_DPI)
        self._cmp_preview.set_after_fn(_cmp_after)
        self._cmp_lb.bind("<<ListboxSelect>>", self._cmp_lb_select)

    def _cmp_add_files(self):
        paths = filedialog.askopenfilenames(filetypes=[("PDF", "*.pdf")])
        for p in paths:
            self._cmp_lb.insert(tk.END, p)
        # auto-preview the first newly added file
        if paths:
            self._cmp_preview.set_source(paths[0])
            self._cmp_preview.refresh()

    def _cmp_lb_select(self, _event):
        sel = self._cmp_lb.curselection()
        if not sel:
            return
        path = self._cmp_lb.get(sel[0])
        if os.path.exists(path):
            self._cmp_preview.set_source(path)
            self._cmp_preview.refresh()



    # ================================================================
    #  TAB: RESTORE  (de-pixelate / de-artifact over-compressed PDFs)
    # ================================================================

    def _build_restore_tab(self):
        tab = self._tabs["restore"]

        pane = tk.PanedWindow(
            tab, orient="horizontal", sashwidth=6,
            bg=DARK["border"], bd=0, sashrelief="flat")
        pane.pack(fill="both", expand=True, padx=4, pady=4)

        ctrl_outer = ttk.Frame(pane)
        prev_f     = ttk.Frame(pane)
        pane.add(ctrl_outer, minsize=300, width=400)
        pane.add(prev_f,     minsize=340)

        # ── LEFT: scrollable so controls never clip ───────────────────
        sf = ScrollableFrame(ctrl_outer)
        sf.pack(fill="both", expand=True)
        cf = sf.inner                       # attach all widgets here

        ttk.Label(cf, text="PDF Quality Restore",
                  style="Section.TLabel").pack(anchor="w", padx=10, pady=(10, 2))
        ttk.Label(cf,
                  text="Re-renders pages at high DPI then runs:\n"
                       "denoise → CLAHE contrast → sharpen → bilateral smooth.\n"
                       "All sliders update the preview in real time.",
                  style="Dim.TLabel").pack(anchor="w", padx=10, pady=(0, 8))

        # File picker
        fp_f = ttk.Frame(cf)
        fp_f.pack(fill="x", padx=10, pady=(0, 4))
        ttk.Label(fp_f, text="Input PDF:").pack(side="left")
        self._rs_input = ttk.Entry(fp_f)
        self._rs_input.pack(side="left", fill="x", expand=True, padx=8)

        def _rs_browse():
            p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
            if p:
                self._rs_input.delete(0, tk.END)
                self._rs_input.insert(0, p)
                self._rs_preview.set_source(p)
                self._rs_preview.refresh()

        ttk.Button(fp_f, text="Browse…", command=_rs_browse).pack(side="left")

        of_f = ttk.Frame(cf)
        of_f.pack(fill="x", padx=10, pady=4)
        ttk.Label(of_f, text="Output Folder:").pack(side="left")
        self._rs_outdir = ttk.Entry(of_f)
        self._rs_outdir.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(of_f, text="Save To…",
                   command=lambda: self._browse_dir(self._rs_outdir)).pack(side="left")

        ttk.Separator(cf, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # Mode preset
        mode_f = ttk.Frame(cf)
        mode_f.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(mode_f, text="Content mode:").pack(side="left")
        self._rs_mode = ttk.Combobox(
            mode_f,
            values=["document  (text, diagrams, sharp edges)",
                    "photo     (natural images, gradients)",
                    "mixed     (detail-preserve both)"],
            state="readonly", width=36)
        self._rs_mode.set("document  (text, diagrams, sharp edges)")
        self._rs_mode.pack(side="left", padx=8)
        self._rs_mode.bind("<<ComboboxSelected>>",
                           lambda _e: self._rs_preview.schedule_refresh(500))

        ttk.Separator(cf, orient="horizontal").pack(fill="x", padx=10, pady=8)

        # ── Slider factory — uses .place() so Scale never drives layout ──
        def _slider_row(label, var, from_, to_, res=1,
                        left_hint="", right_hint=""):
            outer = tk.Frame(cf, bg=DARK["bg"])
            outer.pack(fill="x", padx=10, pady=(0, 6))

            # header: label on left, live value on right
            hdr = tk.Frame(outer, bg=DARK["bg"])
            hdr.pack(fill="x")
            tk.Label(hdr, text=label, bg=DARK["bg"], fg=DARK["text"],
                     font=("Segoe UI", 9)).pack(side="left")
            val_lbl = tk.Label(hdr, text=str(var.get()),
                                bg=DARK["bg"], fg=DARK["accent2"],
                                font=("Consolas", 9, "bold"), width=6, anchor="e")
            val_lbl.pack(side="right")

            # fixed-height container — Scale placed inside so it never
            # causes a cascade geometry reflow when the window is resized
            track = tk.Frame(outer, bg=DARK["bg"], height=26)
            track.pack(fill="x")
            track.pack_propagate(False)

            def _cb(v, _l=val_lbl):
                _l.config(text=str(int(float(v))))
                self._rs_preview.schedule_refresh(500)

            tk.Scale(track, from_=from_, to=to_, orient="horizontal",
                     variable=var, command=_cb, resolution=res,
                     bg=DARK["bg"], fg=DARK["text"],
                     troughcolor=DARK["surface"],
                     activebackground=DARK["accent"],
                     highlightthickness=0, bd=0, showvalue=False,
                     ).place(relx=0, rely=0, relwidth=1.0, relheight=1.0)

            if left_hint or right_hint:
                hr = tk.Frame(outer, bg=DARK["bg"])
                hr.pack(fill="x")
                tk.Label(hr, text=left_hint, bg=DARK["bg"],
                         fg=DARK["text_dim"], font=("Segoe UI", 7)).pack(side="left")
                tk.Label(hr, text=right_hint, bg=DARK["bg"],
                         fg=DARK["text_dim"], font=("Segoe UI", 7)).pack(side="right")

        self._rs_dpi_var       = tk.IntVar(value=300)
        self._rs_denoise_var   = tk.IntVar(value=10)
        self._rs_sharpen_var   = tk.IntVar(value=120)
        self._rs_bilateral_var = tk.IntVar(value=0)

        _slider_row("Render DPI  (higher = sharper, larger output file)",
                    self._rs_dpi_var, 72, 600, res=5,
                    left_hint="72 — screen", right_hint="600 — print / archival")
        _slider_row("Denoise strength  (fastNlMeans h-param — removes JPEG noise)",
                    self._rs_denoise_var, 0, 30, res=1,
                    left_hint="0 — off", right_hint="30 — heavy (may blur detail)")
        _slider_row("Sharpening strength  (Laplacian unsharp — crispness)",
                    self._rs_sharpen_var, 0, 250, res=5,
                    left_hint="0 — off", right_hint="250 — very crisp edges")
        _slider_row("Bilateral filter  (edge-preserving final smooth)",
                    self._rs_bilateral_var, 0, 15, res=1,
                    left_hint="0 — off", right_hint="15 — smooth edges")

        ttk.Separator(cf, orient="horizontal").pack(fill="x", padx=10, pady=8)

        self._rs_clahe_var = tk.BooleanVar(value=True)
        tk.Checkbutton(
            cf,
            text="CLAHE contrast boost (LAB space — greatly improves "
                 "text legibility on faded / over-compressed PDFs)",
            variable=self._rs_clahe_var,
            command=lambda: self._rs_preview.schedule_refresh(500),
            bg=DARK["bg"], fg=DARK["text"], selectcolor=DARK["surface"],
            activebackground=DARK["bg"], activeforeground=DARK["accent"],
            font=("Segoe UI", 9), wraplength=340, justify="left",
        ).pack(anchor="w", padx=10)

        ttk.Separator(cf, orient="horizontal").pack(fill="x", padx=10, pady=8)

        self._rs_status = ttk.Label(
            cf, text="Adjust sliders — preview updates automatically.",
            style="Dim.TLabel")
        self._rs_status.pack(fill="x", padx=10, pady=(0, 4))

        pb_rs = make_progress(cf)
        pb_rs._progress_manager = "pack"

        def _rs_mode_str():
            m = self._rs_mode.get()
            if m.startswith("photo"):  return "photo"
            if m.startswith("mixed"):  return "mixed"
            return "document"

        def _do_restore():
            src = self._rs_input.get().strip()
            if not src or not os.path.exists(src):
                self._show_error("Error", "Select a valid PDF first.")
                return
            dpi     = self._rs_dpi_var.get()
            denoise = self._rs_denoise_var.get()
            sharpen = self._rs_sharpen_var.get()
            clahe   = self._rs_clahe_var.get()
            bilat   = self._rs_bilateral_var.get()
            mode    = _rs_mode_str()
            out_dir = output_dir_or_source_dir(self._rs_outdir.get(), src)
            out     = pdf_output_path(src, f"_restored_{dpi}dpi_{mode}", out_dir)
            self._set_label(self._rs_status,
                            f"Restoring @ {dpi} dpi …", DARK["accent2"])
            log_append(self.log_widget,
                       f"Restore: {os.path.basename(src)} "
                       f"dpi={dpi} denoise={denoise} sharpen={sharpen} "
                       f"clahe={clahe} bilateral={bilat} mode={mode}")

            def worker():
                try:
                    restore_pdf_quality(
                        src, out,
                        render_dpi=dpi, denoise_strength=denoise,
                        sharpen_strength=sharpen, use_clahe=clahe,
                        bilateral_d=bilat, mode=mode,
                        log_widget=self.log_widget,
                    )
                    self.last_output = out
                    self._set_label(self._rs_status,
                                    f"Done — {os.path.basename(out)}",
                                    DARK["success"])
                    self._show_info("Restore Complete", f"Saved:\n{out}")
                except Exception as e:
                    log_append(self.log_widget, f"Restore failed: {e}", "error")
                    self._show_error("Restore Failed", str(e))
            self._run(worker, pb_rs)

        ttk.Button(cf, text="RESTORE PDF QUALITY", style="Accent.TButton",
                   command=_do_restore).pack(anchor="w", padx=10, pady=(4, 4))
        pb_rs.pack(fill="x", padx=12)
        pb_rs.pack_forget()

        libs = []
        if PYMUPDF_AVAILABLE:     libs.append("PyMuPDF ✓")
        elif PYPDFIUM2_AVAILABLE: libs.append("pypdfium2 ✓")
        elif PDF2IMAGE_AVAILABLE: libs.append("pdf2image ✓")
        else:                     libs.append("⚠ no renderer!")
        if CV2_AVAILABLE:         libs.append("OpenCV ✓")
        else:                     libs.append("OpenCV ✗  pip install opencv-python")
        ttk.Label(cf, text="  ·  ".join(libs),
                  style="Dim.TLabel").pack(anchor="w", padx=10, pady=(0, 12))

        # ── RIGHT: live before/after preview ─────────────────────────
        self._rs_preview = PreviewPane(prev_f, self.log_widget)
        self._rs_preview.frame.pack(fill="both", expand=True, padx=(4, 8), pady=10)

        self._rs_preview.set_before_fn(
            lambda src, pg: render_pdf_page_pil(src, pg, PreviewPane.RENDER_DPI))

        def _rs_after(src, pg):
            raw = render_pdf_page_pil(src, pg, PreviewPane.RENDER_DPI)
            return apply_restoration_pipeline(
                raw,
                denoise_strength=self._rs_denoise_var.get(),
                sharpen_strength=self._rs_sharpen_var.get(),
                use_clahe=self._rs_clahe_var.get(),
                bilateral_d=self._rs_bilateral_var.get(),
                mode=_rs_mode_str(),
            )
        self._rs_preview.set_after_fn(_rs_after)



    # ================================================================
    #  TAB: SPLIT
    # ================================================================

    def _build_split_tab(self):
        tab = self._tabs["split"]
        f   = ttk.Frame(tab)
        f.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(f, text="PDF Splitter", style="Section.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0,10))
        f.columnconfigure(1, weight=1)

        self._sp_input = make_file_row(f, "Input PDF:", self.lb_opts,
                                       [("PDF", "*.pdf")], row=1)

        ttk.Label(f, text="Output Folder:").grid(row=2, column=0, sticky="w", pady=4)
        self._sp_outdir = ttk.Entry(f)
        self._sp_outdir.grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Button(f, text="Save To...",
                   command=lambda: self._browse_dir(self._sp_outdir)).grid(
                       row=2, column=2, padx=(8,0))

        ttk.Label(f, text="Mode:").grid(row=3, column=0, sticky="w", pady=4)
        self._sp_mode = ttk.Combobox(f, values=["One file per page", "Extract page range"],
                                      state="readonly", width=22)
        self._sp_mode.set("One file per page")
        self._sp_mode.grid(row=3, column=1, sticky="w")

        ttk.Label(f, text="Page Range:").grid(row=4, column=0, sticky="w", pady=4)
        self._sp_range = ttk.Entry(f)
        self._sp_range.insert(0, "e.g. 1-3,5,8-10")
        self._sp_range.grid(row=4, column=1, sticky="ew", pady=4)
        ttk.Label(f, text="(used only in range mode)", style="Dim.TLabel").grid(
            row=4, column=2, padx=(8,0))

        self._sp_status = ttk.Label(
            f,
            text="Output folder is optional; blank saves split PDFs beside the source PDF.",
            style="Dim.TLabel")
        self._sp_status.grid(row=5, column=0, columnspan=3, sticky="w", pady=(2, 4))

        pb = make_progress(f)
        pb.grid(row=7, column=0, columnspan=3, sticky="ew", pady=4)
        pb._progress_manager = "grid"
        pb.grid_remove()

        def _do_split():
            src = self._sp_input.get().strip()
            if not src or not os.path.exists(src):
                self._show_error("Error", "Select a valid PDF.")
                return
            out_dir = output_dir_or_source_dir(self._sp_outdir.get(), src)
            mode    = "pages" if "per page" in self._sp_mode.get() else "range"
            rng     = self._sp_range.get().strip()
            before_total = total_existing_size([src])
            eta = estimate_job_time(before_total, 30)
            status = (f"Input: {human_size(before_total)}  |  Expected: ~{human_size(before_total)}  "
                      f"|  ETA: ~{human_time(eta)}")
            self._set_label(self._sp_status, status, DARK["accent2"])
            log_append(self.log_widget, status, "info")
            log_append(self.log_widget, f"Splitting {os.path.basename(src)} …")

            def worker():
                try:
                    started = time.perf_counter()
                    files = split_pdf(src, out_dir, mode, rng, self.log_widget)
                    elapsed = time.perf_counter() - started
                    after_total = total_existing_size(files)
                    self.last_output = files[-1] if files else out_dir
                    final_status = (f"Actual: {human_size(after_total)} across {len(files)} file(s)  "
                                    f"|  Time: {human_time(elapsed)}")
                    self._set_label(self._sp_status, final_status, DARK["success"])
                    log_append(self.log_widget, f"Split → {len(files)} file(s)  in  {out_dir}", "success")
                    log_append(self.log_widget, final_status, "success")
                    self._show_info("Done", f"{len(files)} file(s) created in:\n{out_dir}\n\n{final_status}")
                except Exception as e:
                    log_append(self.log_widget, f"Split failed: {e}", "error")
                    self._show_error("Split Failed", str(e))
            self._run(worker, pb)

        ttk.Button(f, text="SPLIT PDF", style="Accent.TButton",
                   command=_do_split).grid(row=6, column=0, columnspan=3, sticky="w", pady=10)

    # ================================================================
    #  TAB: MERGE
    # ================================================================

    def _build_merge_tab(self):
        tab = self._tabs["merge"]
        f   = ttk.Frame(tab)
        f.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(f, text="PDF Merger", style="Section.TLabel").pack(anchor="w", pady=(0,8))

        list_frame, self._mg_lb = make_listbox(f, self.lb_opts, height=10)
        list_frame.pack(fill="both", expand=True)

        btn_row = ttk.Frame(f)
        btn_row.pack(fill="x", pady=6)
        ttk.Button(btn_row, text="＋ Add PDFs",
                   command=lambda: [self._mg_lb.insert(tk.END, p)
                                    for p in filedialog.askopenfilenames(
                                        filetypes=[("PDF","*.pdf")])]).pack(side="left", padx=(0,4))
        ttk.Button(btn_row, text="↑ Up",
                   command=lambda: self._listbox_move(self._mg_lb, -1)).pack(side="left", padx=2)
        ttk.Button(btn_row, text="↓ Down",
                   command=lambda: self._listbox_move(self._mg_lb, 1)).pack(side="left", padx=2)
        ttk.Button(btn_row, text="✕ Remove",
                   command=lambda: self._mg_lb.delete(
                       self._mg_lb.curselection()[0])
                   if self._mg_lb.curselection() else None).pack(side="left", padx=2)
        ttk.Button(btn_row, text="Clear All",
                   command=lambda: self._mg_lb.delete(0, tk.END)).pack(side="left", padx=2)

        out_row = ttk.Frame(f)
        out_row.pack(fill="x", pady=4)
        ttk.Label(out_row, text="Output Folder:").pack(side="left")
        self._mg_outdir = ttk.Entry(out_row)
        self._mg_outdir.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(out_row, text="Save To...",
                   command=lambda: self._browse_dir(self._mg_outdir)).pack(side="left")

        name_row = ttk.Frame(f)
        name_row.pack(fill="x", pady=4)
        ttk.Label(name_row, text="Output Name:").pack(side="left")
        self._mg_outname = ttk.Entry(name_row, width=28)
        self._mg_outname.insert(0, "merged.pdf")
        self._mg_outname.pack(side="left", padx=8)

        self._mg_status = ttk.Label(
            f,
            text="Output folder is optional; blank saves beside the first source PDF.",
            style="Dim.TLabel")
        self._mg_status.pack(fill="x", pady=(2, 4))

        pb = make_progress(f)
        pb._progress_manager = "pack"

        def _do_merge():
            files = list(self._mg_lb.get(0, tk.END))
            if len(files) < 2:
                self._show_warning("Too Few Files", "Add at least 2 PDFs.")
                return
            out_dir = output_dir_or_source_dir(self._mg_outdir.get(), files[0])
            out_name = self._mg_outname.get().strip() or f"{safe_filename(files[0])}_merged.pdf"
            if not out_name.lower().endswith(".pdf"):
                out_name += ".pdf"
            out = os.path.join(out_dir, out_name)
            before_total = total_existing_size(files)
            eta = estimate_job_time(before_total, 25)
            status = (f"Input: {human_size(before_total)}  |  Expected: ~{human_size(before_total)}  "
                      f"|  ETA: ~{human_time(eta)}")
            self._set_label(self._mg_status, status, DARK["accent2"])
            log_append(self.log_widget, status, "info")
            log_append(self.log_widget, f"Merging {len(files)} PDFs …")

            def worker():
                try:
                    started = time.perf_counter()
                    merge_pdfs(files, out, self.log_widget)
                    elapsed = time.perf_counter() - started
                    after_total = total_existing_size([out])
                    self.last_output = out
                    final_status = (f"Actual: {human_size(after_total)}  |  Original: {human_size(before_total)}  "
                                    f"|  {size_delta(before_total, after_total)}  |  Time: {human_time(elapsed)}")
                    self._set_label(self._mg_status, final_status, DARK["success"])
                    log_append(self.log_widget, final_status, "success")
                    self._show_info("Done", f"Merged PDF saved:\n{out}\n\n{final_status}")
                except Exception as e:
                    log_append(self.log_widget, f"Merge failed: {e}", "error")
                    self._show_error("Merge Failed", str(e))
            self._run(worker, pb)

        ttk.Button(f, text="MERGE PDFs", style="Accent.TButton",
                   command=_do_merge).pack(pady=8)
        pb.pack(fill="x", padx=4)
        pb.pack_forget()
        ttk.Label(f, text="Drag items to reorder (or use ↑↓ buttons)", style="Dim.TLabel").pack()

    # ================================================================
    #  TAB: EXTRACT
    # ================================================================

    def _build_extract_tab(self):
        tab = self._tabs["extract"]
        f   = ttk.Frame(tab)
        f.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(f, text="PDF Data Extraction  (powered by pdfplumber)",
                  style="Section.TLabel").grid(row=0, column=0, columnspan=3, sticky="w", pady=(0,10))
        f.columnconfigure(1, weight=1)

        self._ex_input = make_file_row(f, "Input PDF:", self.lb_opts,
                                       [("PDF", "*.pdf")], row=1)

        ttk.Label(f, text="Output File:").grid(row=2, column=0, sticky="w", pady=4)
        self._ex_out = ttk.Entry(f)
        self._ex_out.grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Button(f, text="Pick", command=lambda: self._browse_save(
            self._ex_out, [("Text","*.txt"),("Excel","*.xlsx")])).grid(
                row=2, column=2, padx=(8,0))

        ttk.Label(f, text="Mode:").grid(row=3, column=0, sticky="w", pady=4)
        self._ex_mode = ttk.Combobox(f,
            values=["Extract Text → TXT", "Extract Tables → XLSX", "PDF Metadata"],
            state="readonly", width=28)
        self._ex_mode.set("Extract Text → TXT")
        self._ex_mode.grid(row=3, column=1, sticky="w")

        # live text preview
        prev_lf = ttk.LabelFrame(f, text="Preview")
        prev_lf.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=8)
        f.rowconfigure(5, weight=1)

        self._ex_preview = tk.Text(prev_lf, bg=DARK["log_bg"], fg=DARK["text"],
                                    font=("Consolas", 9), state="disabled",
                                    relief="flat", wrap="word",
                                    insertbackground=DARK["text"])
        sb = ttk.Scrollbar(prev_lf, command=self._ex_preview.yview)
        self._ex_preview.config(yscrollcommand=sb.set)
        self._ex_preview.pack(side="left", fill="both", expand=True, padx=2, pady=2)
        sb.pack(side="right", fill="y")

        pb = make_progress(f)
        pb.grid(row=6, column=0, columnspan=3, sticky="ew")
        pb._progress_manager = "grid"
        pb.grid_remove()

        def _do_extract():
            src = self._ex_input.get().strip()
            if not src or not os.path.exists(src):
                self._show_error("Error", "Select a valid PDF.")
                return
            mode = self._ex_mode.get()
            out  = self._ex_out.get().strip()
            log_append(self.log_widget, f"Extracting from {os.path.basename(src)} …")

            def worker():
                try:
                    out_path = out
                    if "Text" in mode:
                        if not out_path:
                            out_path = os.path.join(os.path.dirname(src), f"{safe_filename(src)}_text.txt")
                        text = extract_text_pdfplumber(src, out_path, self.log_widget)
                        self._set_preview(text[:8000] + ("\n…(truncated)" if len(text) > 8000 else ""))
                        self.last_output = out_path
                        self._show_info("Done", f"Text extracted:\n{out_path}")
                    elif "Tables" in mode:
                        if not out_path:
                            out_path = os.path.join(os.path.dirname(src), f"{safe_filename(src)}_tables.xlsx")
                        df = extract_tables_pdfplumber(src, out_path, self.log_widget)
                        self._set_preview(df.head(30).to_string())
                        self.last_output = out_path
                        self._show_info("Done", f"Tables extracted:\n{out_path}")
                    elif "Metadata" in mode:
                        meta, pages = get_pdf_metadata(src)
                        info = f"Pages: {pages}\n\n" + "\n".join(f"{k}: {v}" for k, v in meta.items())
                        self._set_preview(info)
                        log_append(self.log_widget, f"Metadata: {pages} pages, {len(meta)} fields", "success")
                except Exception as e:
                    log_append(self.log_widget, f"Extraction failed: {e}", "error")
                    self._show_error("Extraction Failed", str(e))
            self._run(worker, pb)

        ttk.Button(f, text="EXTRACT", style="Accent.TButton",
                   command=_do_extract).grid(row=4, column=0, columnspan=3, sticky="w", pady=8)

    def _set_preview(self, text):
        if threading.current_thread() is not threading.main_thread():
            self.after(0, lambda: self._set_preview(text))
            return
        self._ex_preview.config(state="normal")
        self._ex_preview.delete("1.0", tk.END)
        self._ex_preview.insert(tk.END, text)
        self._ex_preview.config(state="disabled")

    # ================================================================
    #  TAB: PROTECT
    # ================================================================

    def _build_protect_tab(self):
        tab = self._tabs["protect"]
        f   = ttk.Frame(tab)
        f.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(f, text="PDF Protection & Watermarking", style="Section.TLabel").pack(anchor="w", pady=(0,10))

        # ── Encrypt section ──────────────────────────────────────────
        enc_lf = ttk.LabelFrame(f, text="Encrypt / Decrypt")
        enc_lf.pack(fill="x", pady=6)
        eg = ttk.Frame(enc_lf)
        eg.pack(fill="x", padx=10, pady=8)
        eg.columnconfigure(1, weight=1)

        self._pr_enc_in = make_file_row(eg, "PDF File:", self.lb_opts,
                                        [("PDF","*.pdf")], row=0)
        ttk.Label(eg, text="User Password:").grid(row=1, column=0, sticky="w", pady=4)
        self._pr_enc_pw = ttk.Entry(eg, show="•")
        self._pr_enc_pw.grid(row=1, column=1, sticky="ew", pady=4)
        ttk.Label(eg, text="Owner Password\n(optional):", style="Dim.TLabel").grid(
            row=2, column=0, sticky="w", pady=4)
        self._pr_enc_ow = ttk.Entry(eg, show="•")
        self._pr_enc_ow.grid(row=2, column=1, sticky="ew", pady=4)

        btn_row = ttk.Frame(enc_lf)
        btn_row.pack(anchor="w", padx=10, pady=(0,8))

        def _do_enc():
            src = self._pr_enc_in.get().strip()
            pw  = self._pr_enc_pw.get()
            if not src or not pw:
                self._show_error("Error", "Select PDF and enter password.")
                return
            out = src.replace(".pdf", "_encrypted.pdf")
            try:
                encrypt_pdf(src, out, pw, self._pr_enc_ow.get(), self.log_widget)
                self._show_info("Done", f"Encrypted:\n{out}")
            except Exception as e:
                self._show_error("Failed", str(e))

        def _do_dec():
            src = self._pr_enc_in.get().strip()
            pw  = self._pr_enc_pw.get()
            if not src or not pw:
                self._show_error("Error", "Select PDF and enter password.")
                return
            out = src.replace(".pdf", "_decrypted.pdf")
            try:
                decrypt_pdf(src, out, pw, self.log_widget)
                self._show_info("Done", f"Decrypted:\n{out}")
            except Exception as e:
                self._show_error("Failed", str(e))

        ttk.Button(btn_row, text="Encrypt PDF", style="Accent.TButton",
                   command=_do_enc).pack(side="left", padx=(0,6))
        ttk.Button(btn_row, text="Remove Encryption",
                   command=_do_dec).pack(side="left")

        # ── Watermark section ────────────────────────────────────────
        wm_lf = ttk.LabelFrame(f, text="Text Watermark")
        wm_lf.pack(fill="x", pady=6)
        wg = ttk.Frame(wm_lf)
        wg.pack(fill="x", padx=10, pady=8)
        wg.columnconfigure(1, weight=1)

        self._pr_wm_in = make_file_row(wg, "PDF File:", self.lb_opts,
                                       [("PDF","*.pdf")], row=0)
        ttk.Label(wg, text="Watermark Text:").grid(row=1, column=0, sticky="w", pady=4)
        self._pr_wm_txt = ttk.Entry(wg)
        self._pr_wm_txt.insert(0, "CONFIDENTIAL")
        self._pr_wm_txt.grid(row=1, column=1, sticky="ew", pady=4)

        def _do_wm():
            src = self._pr_wm_in.get().strip()
            txt = self._pr_wm_txt.get().strip()
            if not src or not txt:
                self._show_error("Error", "Select PDF and enter watermark text.")
                return
            out = src.replace(".pdf", "_watermarked.pdf")
            try:
                add_watermark_text(src, out, txt, self.log_widget)
                self._show_info("Done", f"Watermarked:\n{out}")
            except Exception as e:
                self._show_error("Failed", str(e))

        ttk.Button(wm_lf, text="Apply Watermark", style="Accent.TButton",
                   command=_do_wm).pack(anchor="w", padx=10, pady=(0, 8))

    # ================================================================
    #  TAB: ROTATE
    # ================================================================

    def _build_rotate_tab(self):
        tab = self._tabs["rotate"]
        f   = ttk.Frame(tab)
        f.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(f, text="PDF Page Rotation", style="Section.TLabel").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0,10))
        f.columnconfigure(1, weight=1)

        self._rt_input = make_file_row(f, "Input PDF:", self.lb_opts,
                                       [("PDF","*.pdf")], row=1)
        ttk.Label(f, text="Degrees:").grid(row=2, column=0, sticky="w", pady=4)
        self._rt_deg = ttk.Combobox(f, values=["90", "180", "270"], state="readonly", width=8)
        self._rt_deg.set("90")
        self._rt_deg.grid(row=2, column=1, sticky="w")

        ttk.Label(f, text="Pages:").grid(row=3, column=0, sticky="w", pady=4)
        self._rt_pages = ttk.Entry(f)
        self._rt_pages.insert(0, "all")
        self._rt_pages.grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Label(f, text="all  or  comma list: 1,2,5", style="Dim.TLabel").grid(
            row=3, column=2, padx=(8,0))

        def _do_rotate():
            src = self._rt_input.get().strip()
            if not src or not os.path.exists(src):
                self._show_error("Error", "Select a valid PDF.")
                return
            out = src.replace(".pdf", f"_rotated{self._rt_deg.get()}.pdf")
            try:
                rotate_pdf(src, out, int(self._rt_deg.get()),
                           self._rt_pages.get().strip(), self.log_widget)
                self._show_info("Done", f"Rotated PDF:\n{out}")
            except Exception as e:
                self._show_error("Failed", str(e))

        ttk.Button(f, text="ROTATE", style="Accent.TButton",
                   command=_do_rotate).grid(row=4, column=0, columnspan=3, sticky="w", pady=10)

    # ================================================================
    #  TAB: CRACK
    # ================================================================

    def _build_crack_tab(self):
        tab = self._tabs["crack"]

        # scrollable so nothing clips on small windows
        sf = ScrollableFrame(tab)
        sf.pack(fill="both", expand=True)
        cf = sf.inner

        # ── header ───────────────────────────────────────────────────
        ttk.Label(cf, text="PDF Password Cracker",
                  style="Section.TLabel").pack(anchor="w", padx=14, pady=(12, 2))

        # tool availability badges
        tool_f = tk.Frame(cf, bg=DARK["bg"])
        tool_f.pack(anchor="w", padx=14, pady=(0, 8))
        self._ck_tool_badges = tool_f
        self._refresh_tool_badges()

        ttk.Separator(cf, orient="horizontal").pack(fill="x", padx=14, pady=(0, 8))

        # ── PDF input ─────────────────────────────────────────────────
        sec = ttk.LabelFrame(cf, text="  Target PDF")
        sec.pack(fill="x", padx=14, pady=(0, 8))

        fp_f = ttk.Frame(sec)
        fp_f.pack(fill="x", padx=10, pady=8)
        ttk.Label(fp_f, text="Locked PDF:").pack(side="left")
        self._ck_pdf = ttk.Entry(fp_f)
        self._ck_pdf.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(fp_f, text="Browse…",
                   command=lambda: self._ck_browse_pdf()).pack(side="left")

        # PDF version / protection info
        info_f = ttk.Frame(sec)
        info_f.pack(fill="x", padx=10, pady=(0, 8))
        ttk.Button(info_f, text="🔍 Probe PDF",
                   command=self._ck_probe_pdf).pack(side="left")
        self._ck_pdf_info = tk.Label(
            info_f, text="  (probe to detect encryption version)",
            bg=DARK["bg"], fg=DARK["text_dim"],
            font=("Consolas", 8), anchor="w")
        self._ck_pdf_info.pack(side="left", padx=8)

        # ── Hash extraction ───────────────────────────────────────────
        sec2 = ttk.LabelFrame(cf, text="  Step 1 — Extract Hash  (pdf2john)")
        sec2.pack(fill="x", padx=14, pady=(0, 8))
        s2f = ttk.Frame(sec2)
        s2f.pack(fill="x", padx=10, pady=8)
        ttk.Button(s2f, text="Extract Hash",
                   command=self._ck_extract_hash).pack(side="left")
        self._ck_hash_lbl = tk.Label(
            s2f, text="  not extracted yet",
            bg=DARK["bg"], fg=DARK["text_dim"],
            font=("Consolas", 8), anchor="w")
        self._ck_hash_lbl.pack(side="left", padx=8)
        self._ck_hash_str  = None   # raw hash string
        self._ck_hash_file = None   # path to .hash file

        # ── Engine selector ───────────────────────────────────────────
        sec3 = ttk.LabelFrame(cf, text="  Step 2 — Choose Engine")
        sec3.pack(fill="x", padx=14, pady=(0, 8))

        self._ck_engine = tk.StringVar(value="john")
        eng_f = ttk.Frame(sec3)
        eng_f.pack(fill="x", padx=10, pady=(8, 4))

        for val, label in (("john",    "CPU — John the Ripper"),
                           ("hashcat", "GPU — hashcat")):
            tk.Radiobutton(
                eng_f, text=label, variable=self._ck_engine, value=val,
                command=self._ck_engine_changed,
                bg=DARK["bg"], fg=DARK["text"],
                selectcolor=DARK["surface"],
                activebackground=DARK["bg"],
                activeforeground=DARK["accent"],
                font=("Segoe UI", 9),
            ).pack(side="left", padx=12)

        self._ck_gpu_lbl = tk.Label(
            eng_f, text="",
            bg=DARK["bg"], fg=DARK["text_dim"],
            font=("Segoe UI", 8))
        self._ck_gpu_lbl.pack(side="left", padx=8)

        # probe GPU in background
        threading.Thread(target=self._ck_probe_gpu, daemon=True).start()

        # ── Attack mode ───────────────────────────────────────────────
        sec4 = ttk.LabelFrame(cf, text="  Step 3 — Attack Mode")
        sec4.pack(fill="x", padx=14, pady=(0, 8))

        # John modes
        self._ck_john_frame = ttk.Frame(sec4)
        self._ck_john_frame.pack(fill="x", padx=10, pady=8)

        ttk.Label(self._ck_john_frame, text="Mode:").grid(
            row=0, column=0, sticky="w", pady=4)
        self._ck_john_mode = ttk.Combobox(
            self._ck_john_frame,
            values=["wordlist", "incremental", "markov", "mask"],
            state="readonly", width=14)
        self._ck_john_mode.set("incremental")
        self._ck_john_mode.grid(row=0, column=1, sticky="w", padx=8)
        self._ck_john_mode.bind("<<ComboboxSelected>>",
                                lambda _e: self._ck_mode_changed())

        mode_hints = {
            "wordlist":    "tests passwords from a list — fastest when you have a good list",
            "incremental": "brute-force all character combinations — no wordlist needed",
            "markov":      "human-pattern statistical model — great for real passwords",
            "mask":        "constrained brute-force — use when you know partial structure",
        }
        self._ck_mode_hint = tk.Label(
            self._ck_john_frame, text=mode_hints["incremental"],
            bg=DARK["bg"], fg=DARK["text_dim"],
            font=("Segoe UI", 8), wraplength=480, justify="left")
        self._ck_mode_hint.grid(row=0, column=2, sticky="w", padx=12)

        # wordlist row
        self._ck_wl_row = ttk.Frame(self._ck_john_frame)
        self._ck_wl_row.grid(row=1, column=0, columnspan=3, sticky="ew",
                              pady=(4, 0))
        ttk.Label(self._ck_wl_row, text="Wordlist:").pack(side="left")
        self._ck_wl = ttk.Entry(self._ck_wl_row)
        self._ck_wl.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(self._ck_wl_row, text="Browse…",
                   command=lambda: self._browse_entry(
                       self._ck_wl,
                       [("Text", "*.txt"), ("All", "*.*")])).pack(side="left")
        ttk.Label(self._ck_wl_row,
                  text="optional — blank uses JtR built-in list",
                  style="Dim.TLabel").pack(side="left", padx=8)

        # markov level
        self._ck_markov_row = ttk.Frame(self._ck_john_frame)
        self._ck_markov_row.grid(row=2, column=0, columnspan=3, sticky="ew",
                                  pady=(4, 0))
        ttk.Label(self._ck_markov_row, text="Markov level:").pack(side="left")
        self._ck_markov_var = tk.IntVar(value=200)
        ttk.Spinbox(self._ck_markov_row, from_=50, to=400, width=6,
                    textvariable=self._ck_markov_var).pack(
                        side="left", padx=8)
        ttk.Label(self._ck_markov_row,
                  text="200 = default  ·  higher = more combinations",
                  style="Dim.TLabel").pack(side="left", padx=4)

        # mask row (shared between john mask + hashcat -a3)
        self._ck_mask_row = ttk.Frame(self._ck_john_frame)
        self._ck_mask_row.grid(row=3, column=0, columnspan=3, sticky="ew",
                                pady=(4, 0))
        ttk.Label(self._ck_mask_row, text="Mask:").pack(side="left")
        self._ck_mask = ttk.Entry(self._ck_mask_row, width=26)
        self._ck_mask.insert(0, "?a?a?a?a?a?a")
        self._ck_mask.pack(side="left", padx=8)
        ttk.Label(self._ck_mask_row,
                  text="?l=lower  ?u=upper  ?d=digit  ?s=symbol  ?a=all",
                  style="Dim.TLabel").pack(side="left", padx=4)

        # preset mask buttons — my own idea
        preset_f = ttk.Frame(self._ck_john_frame)
        preset_f.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(4, 0))
        ttk.Label(preset_f, text="Presets:", style="Dim.TLabel").pack(side="left")
        for label, mask_val in (
            ("6-char any",   "?a?a?a?a?a?a"),
            ("8-char any",   "?a?a?a?a?a?a?a?a"),
            ("Cap+5l+2d",    "?u?l?l?l?l?l?d?d"),
            ("6 digits",     "?d?d?d?d?d?d"),
            ("word+4d",      "?l?l?l?l?l?d?d?d?d"),
        ):
            ttk.Button(preset_f, text=label,
                       command=lambda m=mask_val: (
                           self._ck_mask.delete(0, "end"),
                           self._ck_mask.insert(0, m))
                       ).pack(side="left", padx=2)

        # hashcat attack modes (hidden until engine=hashcat)
        self._ck_hc_frame = ttk.Frame(sec4)
        # not packed yet — shown by _ck_engine_changed

        ttk.Label(self._ck_hc_frame, text="Attack:").grid(
            row=0, column=0, sticky="w", pady=4)
        self._ck_hc_attack = ttk.Combobox(
            self._ck_hc_frame,
            values=["0 — wordlist", "3 — mask (brute-force)",
                    "6 — hybrid (wordlist + mask)"],
            state="readonly", width=28)
        self._ck_hc_attack.set("3 — mask (brute-force)")
        self._ck_hc_attack.grid(row=0, column=1, sticky="w", padx=8)

        ttk.Label(self._ck_hc_frame, text="Hash mode:").grid(
            row=1, column=0, sticky="w", pady=4)
        self._ck_hc_mode = ttk.Combobox(
            self._ck_hc_frame,
            values=list(_HASHCAT_MODE_MAP.keys()),
            state="readonly", width=48)
        self._ck_hc_mode.set("Auto-detect from hash")
        self._ck_hc_mode.grid(row=1, column=1, columnspan=2,
                               sticky="ew", padx=8)

        ttk.Label(self._ck_hc_frame, text="Mask:").grid(
            row=2, column=0, sticky="w", pady=4)
        self._ck_hc_mask = ttk.Entry(self._ck_hc_frame, width=26)
        self._ck_hc_mask.insert(0, "?a?a?a?a?a?a")
        self._ck_hc_mask.grid(row=2, column=1, sticky="w", padx=8)
        ttk.Label(self._ck_hc_frame,
                  text="used in -a3 and -a6",
                  style="Dim.TLabel").grid(row=2, column=2,
                  sticky="w", padx=4)

        ttk.Label(self._ck_hc_frame, text="Workload:").grid(
            row=3, column=0, sticky="w", pady=4)
        self._ck_hc_workload = ttk.Combobox(
            self._ck_hc_frame,
            values=["1 — low", "2 — default", "3 — high", "4 — nightmare"],
            state="readonly", width=20)
        self._ck_hc_workload.set("3 — high")
        self._ck_hc_workload.grid(row=3, column=1, sticky="w", padx=8)

        self._ck_john_frame.columnconfigure(2, weight=1)
        self._ck_hc_frame.columnconfigure(2, weight=1)

        # run initial mode toggle
        self._ck_mode_changed()
        self._ck_engine_changed()

        # ── Session management ────────────────────────────────────────
        sec5 = ttk.LabelFrame(cf, text="  Session Management")
        sec5.pack(fill="x", padx=14, pady=(0, 8))
        s5f = ttk.Frame(sec5)
        s5f.pack(fill="x", padx=10, pady=8)

        ttk.Label(s5f, text="Session name:").pack(side="left")
        self._ck_session = ttk.Entry(s5f, width=20)
        self._ck_session.insert(0, "pdftk_crack")
        self._ck_session.pack(side="left", padx=8)

        self._ck_restore_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            s5f, text="Restore previous session",
            variable=self._ck_restore_var,
            bg=DARK["bg"], fg=DARK["text"], selectcolor=DARK["surface"],
            activebackground=DARK["bg"],
            font=("Segoe UI", 9),
        ).pack(side="left", padx=12)

        ttk.Label(
            sec5,
            text="  JtR saves progress automatically every few minutes.  "
                 "Tick 'Restore' to continue a previous run instead of restarting.",
            style="Dim.TLabel",
        ).pack(anchor="w", padx=10, pady=(0, 8))

        # ── Status + controls ─────────────────────────────────────────
        ttk.Separator(cf, orient="horizontal").pack(fill="x", padx=14, pady=4)

        act_f = ttk.Frame(cf)
        act_f.pack(fill="x", padx=14, pady=8)

        self._ck_start_btn = ttk.Button(
            act_f, text="⚡ START ATTACK", style="Danger.TButton",
            command=self._ck_start)
        self._ck_start_btn.pack(side="left", padx=(0, 8))

        self._ck_stop_btn = ttk.Button(
            act_f, text="■ STOP",
            command=self._ck_stop, state="disabled")
        self._ck_stop_btn.pack(side="left", padx=(0, 8))

        ttk.Button(act_f, text="🔑 Show cracked passwords",
                   command=self._ck_show).pack(side="left")

        self._ck_result_lbl = tk.Label(
            cf, text="",
            bg=DARK["bg"], fg=DARK["text_dim"],
            font=("Consolas", 10, "bold"),
            anchor="w")
        self._ck_result_lbl.pack(fill="x", padx=14, pady=(4, 0))

        pb = make_progress(cf)
        pb._progress_manager = "pack"
        pb.pack(fill="x", padx=14, pady=(4, 8))
        pb.pack_forget()

        self._ck_pb         = pb
        self._ck_cancel_ev  = threading.Event()

        ttk.Label(cf,
                  text="ℹ  Install: John the Ripper  ·  hashcat  ·  "
                       "pdf2john.pl (ships with JtR in run/ dir)\n"
                       "⚠  Only use on PDFs you own or have explicit permission to unlock.",
                  style="Dim.TLabel").pack(anchor="w", padx=14, pady=(0, 12))

    # ── crack tab helpers ─────────────────────────────────────────────

    def _refresh_tool_badges(self):
        for w in self._ck_tool_badges.winfo_children():
            w.destroy()
        checks = [
            ("pdf2john", bool(_find_pdf2john())),
            ("john",     bool(_find_john())),
            ("hashcat",  bool(_find_hashcat())),
        ]
        for name, found in checks:
            col  = DARK["success"] if found else DARK["danger"]
            mark = "✓" if found else "✗"
            tk.Label(self._ck_tool_badges,
                     text=f" {mark} {name} ",
                     bg=col, fg="#000000" if found else "#ffffff",
                     font=("Consolas", 8, "bold"),
                     relief="flat", padx=4, pady=2,
                     ).pack(side="left", padx=3)

    def _ck_browse_pdf(self):
        p = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
        if p:
            self._ck_pdf.delete(0, "end")
            self._ck_pdf.insert(0, p)
            self._ck_pdf_info.config(text="  (click Probe PDF for details)",
                                      fg=DARK["text_dim"])
            self._ck_hash_str  = None
            self._ck_hash_file = None
            self._ck_hash_lbl.config(text="  not extracted yet",
                                      fg=DARK["text_dim"])

    def _ck_probe_pdf(self):
        path = self._ck_pdf.get().strip()
        if not path or not os.path.exists(path):
            self._show_error("Error", "Select a PDF first.")
            return
        try:
            # ── try pikepdf first (works even on encrypted files) ────────
            if PIKEPDF_AVAILABLE:
                import pikepdf
                # Normalise PasswordError across pikepdf versions
                _PwErr = getattr(pikepdf, "PasswordError",
                         getattr(getattr(pikepdf, "_core", pikepdf),
                                 "PasswordError", Exception))
                try:
                    with pikepdf.open(path, suppress_warnings=True,
                                      attempt_recovery=True) as pk:
                        trailer = pk.trailer
                        enc_ref = trailer.get("/Encrypt")
                        is_enc  = enc_ref is not None

                        ver = str(pk.pdf_version)
                        n   = len(pk.objects)

                        if is_enc:
                            try:
                                p_val = int(pk.trailer["/Encrypt"].get("/P", -4))
                                bits  = [desc for bit, desc in _PERM_BITS
                                         if p_val & (1 << (bit - 1))]
                                perm_str = ", ".join(bits) if bits else "none"
                            except Exception:
                                perm_str = "unknown"
                            self._ck_pdf_info.config(
                                text=f"  🔒 Encrypted  ·  PDF v{ver}  ·  "
                                     f"{n} objects  ·  Perms: {perm_str[:60]}",
                                fg=DARK["warning"])
                        else:
                            self._ck_pdf_info.config(
                                text=f"  ✓ Not encrypted  ·  PDF v{ver}  ·  "
                                     f"{n} objects",
                                fg=DARK["success"])
                        return
                except _PwErr:
                    self._ck_pdf_info.config(
                        text="  🔒 Encrypted (password required to probe internals)",
                        fg=DARK["warning"])
                    return

            # ── PyMuPDF fallback ─────────────────────────────────────────
            if PYMUPDF_AVAILABLE:
                doc  = fitz.open(path)
                info = doc.metadata or {}          # guard None
                enc  = doc.is_encrypted
                perm = doc.permissions             # int in newer PyMuPDF
                doc.close()
                producer = (info.get("producer") or "unknown")[:40]
                if enc:
                    self._ck_pdf_info.config(
                        text=f"  🔒 Encrypted  ·  Producer: {producer}  ·  "
                             f"Permissions: {perm}",
                        fg=DARK["warning"])
                else:
                    self._ck_pdf_info.config(
                        text=f"  ✓ Not encrypted  ·  Producer: {producer}",
                        fg=DARK["success"])
            else:
                self._ck_pdf_info.config(
                    text="  Install pymupdf or pikepdf to probe PDF details",
                    fg=DARK["text_dim"])
        except Exception as exc:
            self._ck_pdf_info.config(
                text=f"  Probe error: {exc}", fg=DARK["danger"])

    def _ck_probe_gpu(self):
        has_gpu = probe_gpu()
        col  = DARK["success"] if has_gpu else DARK["text_dim"]
        text = "  GPU detected — hashcat recommended" if has_gpu else \
               "  No GPU detected — CPU (john) recommended"
        self.after(0, lambda: self._ck_gpu_lbl.config(text=text, fg=col))

    def _ck_extract_hash(self):
        path = self._ck_pdf.get().strip()
        if not path or not os.path.exists(path):
            self._show_error("Error", "Select a PDF first.")
            return
        self._ck_hash_lbl.config(text="  extracting…", fg=DARK["accent2"])

        def _worker():
            h_str, h_file = extract_pdf_hash(path, self.log_widget)
            if h_str:
                self._ck_hash_str  = h_str
                self._ck_hash_file = h_file
                # auto-detect mode
                mode_id, mode_desc = detect_pdf_hash_mode(h_str)
                short = h_str[:55] + ("…" if len(h_str) > 55 else "")
                self.after(0, lambda: (
                    self._ck_hash_lbl.config(
                        text=f"  ✓ {short}",
                        fg=DARK["success"]),
                    self._ck_pdf_info.config(
                        text=f"  🔒 {mode_desc}  (hashcat -m {mode_id})",
                        fg=DARK["warning"]),
                    # pre-select the matching hashcat mode
                    self._ck_hc_mode.set(
                        next((k for k, v in _HASHCAT_MODE_MAP.items()
                              if v == mode_id), "Auto-detect from hash")),
                ))
            else:
                self.after(0, lambda: self._ck_hash_lbl.config(
                    text="  ✗ extraction failed — check log",
                    fg=DARK["danger"]))

        threading.Thread(target=_worker, daemon=True).start()

    def _ck_engine_changed(self):
        eng = self._ck_engine.get()
        if eng == "john":
            self._ck_hc_frame.pack_forget()
            self._ck_john_frame.pack(fill="x", padx=10, pady=8)
        else:
            self._ck_john_frame.pack_forget()
            self._ck_hc_frame.pack(fill="x", padx=10, pady=8)

    def _ck_mode_changed(self):
        mode = self._ck_john_mode.get()
        hints = {
            "wordlist":    "tests passwords from a list — fastest when you have a good list",
            "incremental": "brute-force all combinations — slow but exhaustive, no wordlist needed",
            "markov":      "human-pattern Markov chains — excellent for real-world passwords",
            "mask":        "constrained brute-force — use when you know partial structure",
        }
        self._ck_mode_hint.config(text=hints.get(mode, ""))
        self._ck_wl_row.grid_remove()
        self._ck_markov_row.grid_remove()
        self._ck_mask_row.grid_remove()
        if mode == "wordlist":    self._ck_wl_row.grid()
        elif mode == "markov":    self._ck_markov_row.grid()
        elif mode == "mask":      self._ck_mask_row.grid()

    def _browse_entry(self, entry, filetypes):
        p = filedialog.askopenfilename(filetypes=filetypes)
        if p:
            entry.delete(0, "end")
            entry.insert(0, p)

    def _ck_start(self):
        pdf = self._ck_pdf.get().strip()
        if not pdf or not os.path.exists(pdf):
            self._show_error("Error", "Select a valid PDF first.")
            return

        # ensure hash is extracted
        if not self._ck_hash_file or not os.path.exists(
                self._ck_hash_file or ""):
            ans = messagebox.askyesno(
                "Extract Hash First?",
                "The hash has not been extracted yet.\n"
                "Extract it now and then start cracking?")
            if not ans:
                return
            self._ck_extract_hash()
            # wait a moment then retry — user will just click Start again
            self._ck_result_lbl.config(
                text="Hash extracted — click START ATTACK again.",
                fg=DARK["accent2"])
            return

        self._ck_cancel_ev.clear()
        self._ck_start_btn.config(state="disabled")
        self._ck_stop_btn.config(state="normal")
        self._ck_result_lbl.config(text="⟳ Cracking in progress…",
                                    fg=DARK["warning"])
        self._ck_pb.pack(fill="x", padx=14)
        self._ck_pb.start(12)

        eng       = self._ck_engine.get()
        session   = self._ck_session.get().strip() or "pdftk_crack"
        restore   = self._ck_restore_var.get()
        hash_file = self._ck_hash_file
        cancel    = self._ck_cancel_ev

        def _warn_session_active():
            """Called when user tries to close while cracking."""
            return messagebox.askyesno(
                "Session Running",
                f"A cracking session '{session}' is active.\n\n"
                "Save progress? (John stores .rec files automatically)\n\n"
                "Yes = stop cleanly (progress saved)\n"
                "No  = abandon immediately")

        def worker():
            pw = None
            try:
                if eng == "john":
                    mode    = self._ck_john_mode.get()
                    wl      = self._ck_wl.get().strip() or None
                    markov  = self._ck_markov_var.get()
                    mask    = self._ck_mask.get().strip() or None
                    pw = crack_with_john(
                        pdf, hash_file,
                        mode=mode, wordlist=wl,
                        mask=mask, markov_level=markov,
                        session_name=session,
                        restore_session=restore,
                        log_widget=self.log_widget,
                        cancel_flag=cancel,
                    )
                else:
                    # hashcat
                    atk_str  = self._ck_hc_attack.get()
                    atk_mode = int(atk_str.split(" ")[0])
                    mode_key = self._ck_hc_mode.get()
                    hc_mode  = _HASHCAT_MODE_MAP.get(mode_key)
                    if hc_mode is None and self._ck_hash_str:
                        hc_mode, _ = detect_pdf_hash_mode(self._ck_hash_str)
                    hc_mode  = hc_mode or 10500
                    wl_str   = self._ck_wl.get().strip() or None
                    mask_str = self._ck_hc_mask.get().strip() or None
                    wl_num   = int(self._ck_hc_workload.get().split(" ")[0])
                    pw = crack_with_hashcat(
                        hash_file,
                        attack_mode=atk_mode,
                        hash_mode=hc_mode,
                        mask=mask_str,
                        wordlist=wl_str,
                        workload=wl_num,
                        log_widget=self.log_widget,
                        cancel_flag=cancel,
                    )
            except Exception as exc:
                log_append(self.log_widget, f"Cracker error: {exc}", "error")
            finally:
                self.after(0, lambda p=pw: self._ck_done(p))

        threading.Thread(target=worker, daemon=True).start()

    def _ck_stop(self):
        save = messagebox.askyesno(
            "Stop Cracking Session",
            f"Stop the session '{self._ck_session.get()}'?\n\n"
            "John the Ripper automatically saves .rec progress files.\n"
            "You can resume later by ticking 'Restore previous session'.\n\n"
            "Yes = signal stop (progress saved)\n"
            "No  = keep running")
        if save:
            self._ck_cancel_ev.set()
            self._ck_stop_btn.config(state="disabled")

    def _ck_done(self, pw):
        self._ck_pb.stop()
        self._ck_pb.pack_forget()
        self._ck_start_btn.config(state="normal")
        self._ck_stop_btn.config(state="disabled")
        if pw:
            self._ck_result_lbl.config(
                text=f"✓ PASSWORD FOUND: {pw}", fg=DARK["success"])
            messagebox.showinfo("Password Found",
                                f"The password is:\n\n{pw}")
        else:
            self._ck_result_lbl.config(
                text="✗ Not found in this run.  Try a different mode or wordlist.",
                fg=DARK["danger"])

    def _ck_show(self):
        """Show all passwords john has cracked so far."""
        john = _find_john()
        if not john:
            self._show_error("Not installed", "john not found on PATH.")
            return
        hf = self._ck_hash_file
        if not hf or not os.path.exists(hf):
            self._show_warning("No hash", "Extract the hash first.")
            return
        try:
            r = subprocess.run([john, "--show", hf],
                               capture_output=True, text=True, timeout=10)
            messagebox.showinfo("Cracked Passwords", r.stdout or "None found yet.")
        except Exception as exc:
            self._show_error("Error", str(exc))

    # ================================================================
    #  TAB: TOOLS
    # ================================================================

    def _build_tools_tab(self):
        tab = self._tabs["tools"]
        f   = ttk.Frame(tab)
        f.pack(fill="both", expand=True, padx=14, pady=14)

        ttk.Label(f, text="Tools & Info", style="Section.TLabel").pack(anchor="w", pady=(0,10))

        # ── OCR ──────────────────────────────────────────────────────
        ocr_lf = ttk.LabelFrame(f, text="OCR — Make Scanned PDF Searchable")
        ocr_lf.pack(fill="x", pady=6)

        def _do_ocr():
            src = filedialog.askopenfilename(filetypes=[("PDF","*.pdf")])
            if not src: return
            out = src.replace(".pdf", "_ocr.pdf")
            def worker():
                try:
                    run_ocrmypdf(src, out, self.log_widget)
                    self.last_output = out
                    self._show_info("Done", f"OCR PDF:\n{out}")
                except Exception as e:
                    self._show_error("OCR Failed", str(e))
            self._run(worker)

        ttk.Label(ocr_lf,
                  text="Needs: pip install ocrmypdf  +  system tesseract",
                  style="Dim.TLabel").pack(anchor="w", padx=10, pady=(6,0))
        ttk.Button(ocr_lf, text="Run OCR on PDF", command=_do_ocr).pack(anchor="w", padx=10, pady=8)

        # ── Package status ───────────────────────────────────────────
        pkg_lf = ttk.LabelFrame(f, text="Package Status")
        pkg_lf.pack(fill="x", pady=6)

        pkgs = [
            ("pandas",      pd is not None,       "pip install pandas"),
            ("Pillow",      PIL_AVAILABLE,         "pip install pillow"),
            ("reportlab",   REPORTLAB_AVAILABLE,   "pip install reportlab"),
            ("pypdf",       PYPDF_AVAILABLE,       "pip install pypdf"),
            ("pikepdf",     PIKEPDF_AVAILABLE,     "pip install pikepdf"),
            ("pdfplumber",  PDFPLUMBER_AVAILABLE,  "pip install pdfplumber"),
            ("pdf2image",   PDF2IMAGE_AVAILABLE,   "pip install pdf2image  (+poppler)"),
            ("openpyxl",    OPENPYXL_AVAILABLE,    "pip install openpyxl"),
            ("python-docx", DOCX_AVAILABLE,        "pip install python-docx"),
            ("docx2pdf",    DOCX2PDF_AVAILABLE,    "pip install docx2pdf"),
            ("psutil",      PSUTIL_AVAILABLE,      "pip install psutil"),
            ("tkinterdnd2", DND_AVAILABLE,         "pip install tkinterdnd2"),
        ]
        grid = ttk.Frame(pkg_lf)
        grid.pack(fill="x", padx=10, pady=8)

        for i, (name, ok, install) in enumerate(pkgs):
            row, col = divmod(i, 2)
            color = DARK["success"] if ok else DARK["danger"]
            mark  = "✓" if ok else "✗"
            tip   = "" if ok else f"  →  {install}"
            tk.Label(grid, text=f"{mark}  {name}{tip}",
                     bg=DARK["bg"], fg=color,
                     font=("Consolas", 9)).grid(
                         row=row, column=col, sticky="w", padx=12, pady=2)

        # ── Open last output ─────────────────────────────────────────
        ttk.Button(f, text="Reveal Last Output in Explorer",
                   command=lambda: (open_path(self.last_output)
                                    if self.last_output else
                                    self._show_info("Nothing", "No output yet."))).pack(
                                        anchor="w", pady=10)

    # ================================================================
    #  HELPERS
    # ================================================================

    def _browse_dir(self, entry):
        d = filedialog.askdirectory()
        if d:
            entry.delete(0, tk.END)
            entry.insert(0, d)

    def _browse_save(self, entry, filetypes):
        f = filedialog.asksaveasfilename(filetypes=filetypes)
        if f:
            entry.delete(0, tk.END)
            entry.insert(0, f)

    def _listbox_move(self, lb, direction):
        sel = lb.curselection()
        if not sel:
            return
        i = sel[0]
        j = i + direction
        if j < 0 or j >= lb.size():
            return
        a, b = lb.get(i), lb.get(j)
        lb.delete(i)
        lb.insert(i, b)
        lb.delete(j)
        lb.insert(j, a)
        lb.selection_set(j)

# =====================================================================
#  ENTRY POINT
# =====================================================================

def main():
    app = ConverterApp()
    app.mainloop()


if __name__ == "__main__":
    main()
