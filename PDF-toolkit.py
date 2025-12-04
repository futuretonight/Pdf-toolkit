#!/usr/bin/env python3

"""
Universal Converter + Abyss Toolkit (Tkinter Edition) - ULTIMATE FIXED VERSION
Full-featured document converter with compression, cracking, OCR, and merge tools.
Supports: CSV, XLSX, DOCX, PDF, Images, TXT

Features:
- Advanced PDF/document conversion (20+ formats)
- Batch PDF compression with quality presets
- PDF password unlock (manual + John the Ripper cracking)
- Password strength indicator with real-time feedback
- OCR for scanned PDFs
- PDF merging and manipulation
- System monitoring (CPU/RAM)
- Termux Android integration

Requirements:
    pip install pandas pillow reportlab openpyxl python-docx pdf2image pypdf psutil
    apt install ghostscript poppler-utils tesseract-ocr (Linux/Termux)
    pkg install john                 # For PDF password cracking
"""

import os
import sys
import threading
import subprocess
import shutil
from datetime import datetime
import tempfile
import traceback
import re

# Tkinter
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
except ImportError as e:
    print("ERROR: Tkinter not available. Install python3-tk")
    sys.exit(1)

# Optional dependencies
try:
    import pandas as pd
except ImportError:
    pd = None

try:
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
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
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False


# =====================================================================
# UTILITY FUNCTIONS
# =====================================================================

def now():
    """Return current time as HH:MM:SS string."""
    return datetime.now().strftime("%H:%M:%S")


def log_append(log_widget, text):
    """Append text to scrolled text widget with timestamp."""
    if log_widget is None:
        return
    try:
        log_widget.configure(state='normal')
        log_widget.insert(tk.END, f"[{now()}] {text}\n")
        log_widget.yview_moveto(1.0)
        log_widget.configure(state='disabled')
    except Exception:
        pass


def safe_filename(path):
    """Extract filename without extension."""
    base = os.path.basename(path)
    name, _ = os.path.splitext(base)
    return name


def ensure_dir(path):
    """Ensure directory exists for given file path."""
    directory = os.path.dirname(path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory, exist_ok=True)


def detect_file_type(path):
    """Detect file type by extension."""
    ext = os.path.splitext(path)[1].lower().strip(".")
    return ext if ext else "unknown"


def check_password_strength(password):
    """
    Check password strength and return (score, label, color).
    Score: 0-5
    """
    if not password:
        return 0, "No Password", "gray"

    score = 0
    length = len(password)

    # Length scoring
    if length >= 8:
        score += 1
    if length >= 12:
        score += 1
    if length >= 16:
        score += 1

    # Character variety
    if re.search(r'[a-z]', password):
        score += 1
    if re.search(r'[A-Z]', password):
        score += 1
    if re.search(r'[0-9]', password):
        score += 1
    if re.search(r'[^a-zA-Z0-9]', password):
        score += 1

    # Cap at 5
    score = min(score, 5)

    # Labels and colors
    if score <= 1:
        return score, "Very Weak", "red"
    elif score == 2:
        return score, "Weak", "orange"
    elif score == 3:
        return score, "Moderate", "yellow"
    elif score == 4:
        return score, "Strong", "lightgreen"
    else:
        return score, "Very Strong", "green"


def create_password_entry_with_toggle(parent, log_widget=None, with_strength=False):
    """
    Create password entry with show/hide toggle button.
    Returns (frame, entry, toggle_btn, [strength_label]) tuple.
    """
    frame = ttk.Frame(parent)

    entry = ttk.Entry(frame, show="•", width=30)
    entry.pack(side='left', fill='x', expand=True, padx=(0, 5))

    toggle_btn = ttk.Button(frame, text="👁", width=3)
    toggle_btn.pack(side='left')

    strength_label = None
    if with_strength:
        strength_label = ttk.Label(frame, text="", width=12)
        strength_label.pack(side='left', padx=5)

        def update_strength(*args):
            pwd = entry.get()
            score, label, color = check_password_strength(pwd)
            strength_label.config(text=label, foreground=color)

        entry.bind('<KeyRelease>', update_strength)

    def toggle_visibility():
        if entry.cget('show') == '•':
            entry.config(show='')
            toggle_btn.config(text="👁‍🗨")
        else:
            entry.config(show='•')
            toggle_btn.config(text="👁")

    toggle_btn.config(command=toggle_visibility)

    if with_strength:
        return frame, entry, toggle_btn, strength_label
    return frame, entry, toggle_btn


# =====================================================================
# CONVERSION FUNCTIONS - DATAFRAME TO IMAGE/PDF
# =====================================================================

def df_to_image(df, out_path, log_widget=None, max_width=4096):
    """Convert pandas DataFrame to image using PIL."""
    if not PIL_AVAILABLE:
        raise RuntimeError("Pillow not installed. Install: pip install pillow")

    if pd is None:
        raise RuntimeError("Pandas not installed. Install: pip install pandas")

    # Load font
    try:
        font = ImageFont.truetype("DejaVuSans.ttf", 16)
    except Exception:
        try:
            font = ImageFont.truetype("/system/fonts/DroidSans.ttf", 16)
        except Exception:
            font = ImageFont.load_default()

    pad = 10
    dummy_img = Image.new("RGB", (10, 10))
    draw = ImageDraw.Draw(dummy_img)

    columns = list(df.columns)
    rows = df.astype(str).values.tolist()

    # Calculate column widths (using getbbox for Pillow 10.0+ compatibility)
    col_widths = []
    for i, col in enumerate(columns):
        try:
            # Try modern Pillow method
            bbox = draw.textbbox((0, 0), str(col), font=font)
            w = (bbox[2] - bbox[0]) + pad * 2
        except Exception:
            # Fallback for older Pillow
            w = draw.textsize(str(col), font=font)[0] + pad * 2

        for row in rows:
            cell_text = str(row[i])
            try:
                bbox = draw.textbbox((0, 0), cell_text, font=font)
                cell_w = (bbox[2] - bbox[0]) + pad * 2
            except Exception:
                cell_w = draw.textsize(cell_text, font=font)[0] + pad * 2
            w = max(w, cell_w)
        col_widths.append(min(w, 300))  # Max 300px per column

    try:
        bbox = draw.textbbox((0, 0), "Ag", font=font)
        row_height = (bbox[3] - bbox[1]) + pad
    except Exception:
        row_height = draw.textsize("Ag", font=font)[1] + pad

    header_height = row_height + 10

    total_width = sum(col_widths)
    total_height = header_height + row_height * len(rows)

    # Scale down if too wide
    if total_width > max_width:
        scale = max_width / total_width
        col_widths = [int(w * scale) for w in col_widths]
        total_width = max_width

    # Create image
    img = Image.new("RGB", (total_width, total_height), "white")
    draw = ImageDraw.Draw(img)

    # Draw header
    x = 0
    for i, col in enumerate(columns):
        w = col_widths[i]
        draw.rectangle([x, 0, x + w, header_height], fill=(220, 220, 220), outline="black")
        draw.text((x + pad, 5), str(col), font=font, fill="black")
        x += w

    # Draw rows
    y = header_height
    for row in rows:
        x = 0
        for i, cell in enumerate(row):
            w = col_widths[i]
            draw.rectangle([x, y, x + w, y + row_height], outline="black", fill="white")
            cell_text = str(cell)[:50]  # Truncate long text
            draw.text((x + pad, y + 5), cell_text, font=font, fill="black")
            x += w
        y += row_height

    ensure_dir(out_path)
    img.save(out_path)
    if log_widget:
        log_append(log_widget, f"Image saved: {out_path}")


def df_to_pdf(df, out_path, log_widget=None):
    """Convert pandas DataFrame to PDF using ReportLab."""
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ReportLab not installed. Install: pip install reportlab")

    if pd is None:
        raise RuntimeError("Pandas not installed. Install: pip install pandas")

    data = [list(df.columns)]
    for row in df.itertuples(index=False):
        data.append([str(x) for x in row])

    ensure_dir(out_path)
    doc = SimpleDocTemplate(out_path, pagesize=A4)

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTSIZE', (0, 0), (-1, -1), 10),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))

    doc.build([table])
    if log_widget:
        log_append(log_widget, f"PDF saved: {out_path}")


# =====================================================================
# FORMAT-SPECIFIC CONVERSION FUNCTIONS
# =====================================================================

def csv_to_pdf(src, out, log_widget=None):
    """Convert CSV to PDF."""
    if pd is None:
        raise RuntimeError("Pandas required for CSV conversion")
    df = pd.read_csv(src)
    df_to_pdf(df, out, log_widget)


def csv_to_image(src, out, log_widget=None):
    """Convert CSV to image."""
    if pd is None:
        raise RuntimeError("Pandas required for CSV conversion")
    df = pd.read_csv(src)
    df_to_image(df, out, log_widget)


def csv_to_xlsx(src, out, log_widget=None):
    """Convert CSV to XLSX."""
    if pd is None:
        raise RuntimeError("Pandas required for CSV conversion")
    df = pd.read_csv(src)
    ensure_dir(out)
    df.to_excel(out, index=False)
    if log_widget:
        log_append(log_widget, f"XLSX saved: {out}")


def xlsx_to_pdf(src, out, log_widget=None):
    """Convert XLSX to PDF."""
    if pd is None:
        raise RuntimeError("Pandas required for XLSX conversion")
    df = pd.read_excel(src)
    df_to_pdf(df, out, log_widget)


def xlsx_to_image(src, out, log_widget=None):
    """Convert XLSX to image."""
    if pd is None:
        raise RuntimeError("Pandas required for XLSX conversion")
    df = pd.read_excel(src)
    df_to_image(df, out, log_widget)


def xlsx_to_csv(src, out, log_widget=None):
    """Convert XLSX to CSV."""
    if pd is None:
        raise RuntimeError("Pandas required for XLSX conversion")
    df = pd.read_excel(src)
    ensure_dir(out)
    df.to_csv(out, index=False)
    if log_widget:
        log_append(log_widget, f"CSV saved: {out}")


def txt_to_pdf(src, out, log_widget=None):
    """Convert text file to PDF."""
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ReportLab required for TXT to PDF conversion")

    with open(src, 'r', encoding='utf-8', errors='ignore') as f:
        text = f.read()

    ensure_dir(out)
    doc = SimpleDocTemplate(out, pagesize=A4)
    styles = getSampleStyleSheet()

    # Replace newlines with HTML breaks
    text_html = text.replace('\n', '<br/>')
    para = Paragraph(text_html, styles["Normal"])

    doc.build([para])
    if log_widget:
        log_append(log_widget, f"PDF saved: {out}")


def image_to_pdf(src, out, log_widget=None):
    """Convert image to PDF."""
    if not PIL_AVAILABLE:
        raise RuntimeError("Pillow required for image conversion")

    img = Image.open(src)
    if img.mode == "RGBA":
        img = img.convert("RGB")

    ensure_dir(out)
    img.save(out, "PDF")
    if log_widget:
        log_append(log_widget, f"PDF saved: {out}")


def pdf_to_images(src, out_dir, log_widget=None):
    """Convert PDF pages to images."""
    if not PDF2IMAGE_AVAILABLE:
        raise RuntimeError("pdf2image not installed. Install: pip install pdf2image")

    ensure_dir(out_dir)

    try:
        pages = convert_from_path(src)
        outfiles = []
        for i, page in enumerate(pages, 1):
            out_path = os.path.join(out_dir, f"{safe_filename(src)}_page_{i}.png")
            page.save(out_path, "PNG")
            outfiles.append(out_path)
            if log_widget:
                log_append(log_widget, f"Extracted page {i}")
        return outfiles
    except Exception as e:
        raise RuntimeError(f"PDF extraction failed: {e}")


def docx_to_pdf(src, out, log_widget=None):
    """Convert DOCX to PDF."""
    # Try docx2pdf first (Windows compatible)
    if DOCX2PDF_AVAILABLE:
        try:
            tmp_dir = tempfile.mkdtemp()
            docx2pdf_convert(src, tmp_dir)
            produced = os.path.join(tmp_dir, safe_filename(src) + ".pdf")
            if os.path.exists(produced):
                ensure_dir(out)
                shutil.move(produced, out)
                shutil.rmtree(tmp_dir, ignore_errors=True)
                if log_widget:
                    log_append(log_widget, f"PDF saved: {out}")
                return
        except Exception as e:
            if log_widget:
                log_append(log_widget, f"docx2pdf failed: {e}")

    # Try LibreOffice
    if shutil.which("soffice"):
        try:
            tmp_dir = tempfile.mkdtemp()
            cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", tmp_dir, src]
            subprocess.run(cmd, check=True, capture_output=True)
            produced = os.path.join(tmp_dir, safe_filename(src) + ".pdf")
            if os.path.exists(produced):
                ensure_dir(out)
                shutil.move(produced, out)
                shutil.rmtree(tmp_dir, ignore_errors=True)
                if log_widget:
                    log_append(log_widget, f"PDF saved: {out}")
                return
        except Exception as e:
            if log_widget:
                log_append(log_widget, f"LibreOffice conversion failed: {e}")

    # Fallback: extract text and create simple PDF
    if not DOCX_AVAILABLE:
        raise RuntimeError("python-docx not installed. Install: pip install python-docx")

    document = docx.Document(src)
    paragraphs = "\n".join([p.text for p in document.paragraphs])

    txt_file = os.path.join(tempfile.gettempdir(), "tmp_docx.txt")
    with open(txt_file, "w", encoding='utf-8') as f:
        f.write(paragraphs)

    txt_to_pdf(txt_file, out, log_widget)


def docx_to_images(src, out_dir, log_widget=None):
    """Convert DOCX to images via PDF."""
    tmp_pdf = os.path.join(tempfile.gettempdir(), safe_filename(src) + "_tmp.pdf")
    docx_to_pdf(src, tmp_pdf, log_widget)
    return pdf_to_images(tmp_pdf, out_dir, log_widget)


# =====================================================================
# PDF PASSWORD FUNCTIONS
# =====================================================================

def unlock_pdf_manual(pdf_path, password, log_widget):
    """
    Unlock a password-protected PDF using known password.
    Returns path to unlocked PDF.
    """
    if not PYPDF_AVAILABLE:
        raise RuntimeError("pypdf not installed. Install: pip install pypdf")

    try:
        log_append(log_widget, f"Attempting to unlock: {os.path.basename(pdf_path)}")

        reader = pypdf.PdfReader(pdf_path)

        # Check if encrypted
        if not reader.is_encrypted:
            log_append(log_widget, "PDF is not encrypted")
            messagebox.showinfo("Not Encrypted", "This PDF is not password-protected")
            return None

        # Try to decrypt
        if not reader.decrypt(password):
            log_append(log_widget, "❌ Incorrect password")
            return None

        log_append(log_widget, "✔ Password correct, creating unlocked version...")

        # Create unlocked version
        writer = pypdf.PdfWriter()

        for page in reader.pages:
            writer.add_page(page)

        # Generate output path
        base = os.path.splitext(pdf_path)[0]
        out_path = f"{base}_unlocked.pdf"

        # Handle file exists
        counter = 1
        while os.path.exists(out_path):
            out_path = f"{base}_unlocked_{counter}.pdf"
            counter += 1

        ensure_dir(out_path)
        with open(out_path, "wb") as output_file:
            writer.write(output_file)

        log_append(log_widget, f"✔ Unlocked PDF saved: {out_path}")
        return out_path

    except Exception as e:
        log_append(log_widget, f"❌ Unlock failed: {e}")
        return None


def crack_pdf_with_john(pdf_path, wordlist, log_widget):
    """
    Attempt to crack PDF password using John the Ripper.
    Handles pdf2john, pdf2john.py, pdf2john.pl across Termux/Linux/macOS.
    """
    # Try to find pdf2john in multiple forms
    pdf2john_cmd = None

    # Check common locations and names
    candidates = [
        "pdf2john",           # Most common (symlink or Python version)
        "pdf2john.py",        # Explicit Python version
        "pdf2john.pl",        # Perl version (older)
        "/data/data/com.termux/files/usr/share/john/pdf2john.py",  # Termux Python
        "/data/data/com.termux/files/usr/share/john/pdf2john.pl",  # Termux Perl
        "/usr/share/john/pdf2john.py",   # Linux Python
        "/usr/share/john/pdf2john.pl",   # Linux Perl
    ]

    for candidate in candidates:
        if os.path.exists(candidate):
            pdf2john_cmd = candidate
            break
        elif shutil.which(candidate):
            pdf2john_cmd = candidate
            break

    if not pdf2john_cmd:
        log_append(log_widget, "❌ pdf2john not found. Install: pkg install john")
        return None

    log_append(log_widget, f"Using: {pdf2john_cmd}")

    hash_file = pdf_path + ".hash"

    # Extract hash
    try:
        log_append(log_widget, "Extracting PDF hash...")

        # Handle Python vs Perl versions
        if pdf2john_cmd.endswith(".py"):
            cmd = ["python3", pdf2john_cmd, pdf_path]
        elif pdf2john_cmd.endswith(".pl"):
            cmd = ["perl", pdf2john_cmd, pdf_path]
        else:
            cmd = [pdf2john_cmd, pdf_path]

        with open(hash_file, "w") as hf:
            result = subprocess.run(cmd, stdout=hf, stderr=subprocess.PIPE, text=True)
            if result.returncode != 0:
                log_append(log_widget, f"Hash extraction failed: {result.stderr}")
                return None

    except Exception as e:
        log_append(log_widget, f"Hash extraction error: {e}")
        return None

    # Verify hash file exists and has content
    if not os.path.exists(hash_file) or os.path.getsize(hash_file) == 0:
        log_append(log_widget, "❌ Hash file empty. PDF may not be encrypted.")
        return None

    # Run John the Ripper
    john_cmd = ["john", hash_file]
    if wordlist and os.path.exists(wordlist):
        john_cmd.append(f"--wordlist={wordlist}")
        log_append(log_widget, f"Using wordlist: {os.path.basename(wordlist)}")
    else:
        log_append(log_widget, "No wordlist — using default John modes")

    try:
        log_append(log_widget, "🔥 Running John the Ripper...")
        proc = subprocess.Popen(
            john_cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True
        )

        for line in proc.stdout:
            log_append(log_widget, f"JOHN: {line.rstrip()}")

        proc.wait()

        # Show cracked passwords
        result = subprocess.run(
            ["john", "--show", hash_file],
            capture_output=True,
            text=True
        )

        for line in result.stdout.splitlines():
            if ":" in line and not line.startswith("0 password"):
                parts = line.split(":")
                if len(parts) >= 2:
                    password = parts[1].strip()
                    log_append(log_widget, f"🔑 PASSWORD FOUND: {password}")
                    return password

        log_append(log_widget, "❌ No password found")
        return None

    except FileNotFoundError:
        log_append(log_widget, "❌ 'john' command not found. Install: pkg install john")
        return None
    except Exception as e:
        log_append(log_widget, f"Cracking error: {e}")
        return None
    finally:
        # Cleanup hash file
        if os.path.exists(hash_file):
            try:
                os.remove(hash_file)
            except Exception:
                pass


# =====================================================================
# ABYSS TOOLKIT FUNCTIONS
# =====================================================================

def compress_pdf_list(file_list, dpi, log_widget, delete_original=False):
    """Compress PDFs using Ghostscript."""
    dpi_setting = str(int(dpi))

    for file_path in file_list:
        base, ext = os.path.splitext(file_path)
        out_path = f"{base}_compressed{ext}"

        counter = 1
        final_out_path = out_path
        while os.path.exists(final_out_path):
            final_out_path = f"{base}_compressed_{counter}{ext}"
            counter += 1

        cmd = [
            "gs", "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
            "-dPDFSETTINGS=/ebook",  # Use a base setting
            f"-dColorImageResolution={dpi_setting}",
            f"-dGrayImageResolution={dpi_setting}",
            f"-dMonoImageResolution={dpi_setting}",
            "-dColorImageDownsampleType=/Bicubic",
            "-dGrayImageDownsampleType=/Bicubic",
            "-dMonoImageDownsampleType=/Bicubic",
            "-dNOPAUSE", "-dQUIET", "-dBATCH",
            f"-sOutputFile={final_out_path}", file_path
        ]

        log_append(log_widget, f"Compressing: {os.path.basename(file_path)}")

        try:
            # Run Ghostscript
            result = subprocess.run(cmd, check=True, capture_output=True, text=True)
            log_append(log_widget, f"Saved: {final_out_path}")

            # Delete original if requested and compression was successful
            if delete_original:
                try:
                    os.remove(file_path)
                    log_append(log_widget, f"Deleted original: {os.path.basename(file_path)}")
                except OSError as e:
                    log_append(log_widget, f"Failed to delete {os.path.basename(file_path)}: {e}")

        except subprocess.CalledProcessError as e:
            log_append(log_widget, f"Compression failed for {os.path.basename(file_path)}: {e.stderr}")
        except FileNotFoundError:
            log_append(log_widget, "Ghostscript not found. Install: apt install ghostscript")
            break


def merge_pdfs(pdf_list, out_path, log_widget):
    """Merge multiple PDFs into one."""
    if not PYPDF_AVAILABLE:
        raise RuntimeError("pypdf not installed. Install: pip install pypdf")

    writer = pypdf.PdfWriter()

    for pdf_path in pdf_list:
        try:
            writer.append(pdf_path)
            log_append(log_widget, f"Appended: {os.path.basename(pdf_path)}")
        except Exception as e:
            log_append(log_widget, f"Failed to append {pdf_path}: {e}")

    ensure_dir(out_path)
    with open(out_path, "wb") as output_file:
        writer.write(output_file)

    log_append(log_widget, f"Merged PDF saved: {out_path}")
    return out_path


def run_ocrmypdf(src, out, log_widget):
    """Run OCR on PDF using ocrmypdf."""
    if not shutil.which("ocrmypdf"):
        raise RuntimeError("ocrmypdf not found. Install: pip install ocrmypdf")

    cmd = ["ocrmypdf", "--force-ocr", src, out]
    log_append(log_widget, "Starting OCR process...")

    try:
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode == 0:
            log_append(log_widget, f"OCR completed: {out}")
        else:
            log_append(log_widget, f"OCR error: {result.stderr}")
    except Exception as e:
        log_append(log_widget, f"OCR failed: {e}")
        raise


def termux_share_file(path, log_widget):
    """Share file using Termux API."""
    if shutil.which("termux-share"):
        subprocess.Popen(["termux-share", "-a", "send", path])
        log_append(log_widget, f"Shared via Termux: {path}")
    else:
        log_append(log_widget, "termux-share not found")


# =====================================================================
# MAIN APPLICATION
# =====================================================================

class ConverterApp(tk.Tk):
    """Main application window."""

    def __init__(self):
        super().__init__()
        self.title("Universal Converter + Abyss Toolkit")
        self.minsize(700, 500)

        self.last_output = None
        self._abyss_injected = False

        self.create_ui()
        self.after(500, self.inject_abyss_features)

    def create_ui(self):
        """Create main UI layout."""
        # Title bar
        title_frame = ttk.Frame(self)
        title_frame.pack(fill='x', padx=8, pady=6)

        ttk.Label(
            title_frame,
            text="UNIVERSAL CONVERTER TOOLKIT",
            font=("Arial", 16, "bold")
        ).pack(side='left')

        # System stats (will be populated later)
        self.stats_frame = ttk.Frame(title_frame)
        self.stats_frame.pack(side='right')

        # Notebook (tabs)
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # Create tabs
        self.tab_convert = ttk.Frame(self.notebook)
        self.tab_compress = ttk.Frame(self.notebook)
        self.tab_crack = ttk.Frame(self.notebook)
        self.tab_tools = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_convert, text="Convert")
        self.notebook.add(self.tab_compress, text="Compress")
        self.notebook.add(self.tab_crack, text="Crack")
        self.notebook.add(self.tab_tools, text="Tools")

        # Build Convert tab
        self.build_convert_tab()

        # Placeholder labels for other tabs (will be built by inject_abyss_features)
        ttk.Label(self.tab_compress, text="Loading...").pack(pady=20)
        ttk.Label(self.tab_crack, text="Loading...").pack(pady=20)
        ttk.Label(self.tab_tools, text="Loading...").pack(pady=20)

        # Global log
        log_frame = ttk.LabelFrame(self, text="SYSTEM LOG")
        log_frame.pack(fill='both', expand=True, padx=10, pady=10)

        self.log_widget = scrolledtext.ScrolledText(
            log_frame,
            height=12,
            state='disabled',
            bg="#111111",
            fg="#00ff00",
            font=("Courier", 9)
        )
        self.log_widget.pack(fill='both', expand=True, padx=5, pady=5)

    def build_convert_tab(self):
        """Build the Convert tab UI."""
        container = ttk.Frame(self.tab_convert)
        container.pack(fill='both', expand=True, padx=10, pady=10)

        # Input file selection
        input_frame = ttk.Frame(container)
        input_frame.pack(fill='x', pady=5)

        ttk.Label(input_frame, text="Input File:").pack(side='left')
        self.input_entry = ttk.Entry(input_frame)
        self.input_entry.pack(side='left', fill='x', expand=True, padx=8)
        ttk.Button(input_frame, text="Browse", command=self.browse_input).pack(side='left')

        # Detected type
        type_frame = ttk.Frame(container)
        type_frame.pack(fill='x', pady=5)

        ttk.Label(type_frame, text="Detected Type:").pack(side='left')
        self.type_label = ttk.Label(type_frame, text="None", foreground="blue")
        self.type_label.pack(side='left', padx=8)

        # Output options
        output_frame = ttk.Frame(container)
        output_frame.pack(fill='x', pady=5)

        ttk.Label(output_frame, text="Output Format:").pack(side='left')
        self.format_combo = ttk.Combobox(
            output_frame,
            values=["pdf", "png", "jpg", "xlsx", "csv"],
            state='readonly',
            width=10
        )
        self.format_combo.set("pdf")
        self.format_combo.pack(side='left', padx=8)

        ttk.Label(output_frame, text="Output Folder:").pack(side='left', padx=(20, 0))
        self.output_folder_entry = ttk.Entry(output_frame, width=25)
        self.output_folder_entry.pack(side='left', padx=8)
        ttk.Button(output_frame, text="Pick Folder", command=self.browse_output).pack(side='left')

        # Convert button
        button_frame = ttk.Frame(container)
        button_frame.pack(fill='x', pady=10)

        self.convert_button = ttk.Button(
            button_frame,
            text="CONVERT",
            command=self.start_conversion
        )
        self.convert_button.pack(side='left')

        ttk.Button(
            button_frame,
            text="Open Output Folder",
            command=self.open_output_folder
        ).pack(side='left', padx=10)

        # Results list
        results_frame = ttk.LabelFrame(container, text="Output Files")
        results_frame.pack(fill='both', expand=True, pady=10)

        self.results_listbox = tk.Listbox(results_frame, height=10)
        self.results_listbox.pack(fill='both', expand=True, padx=5, pady=5)

    def browse_input(self):
        """Browse for input file."""
        filename = filedialog.askopenfilename(title="Select input file")
        if filename:
            self.input_entry.delete(0, tk.END)
            self.input_entry.insert(0, filename)

            file_type = detect_file_type(filename)
            self.type_label.config(text=file_type.upper())
            log_append(self.log_widget, f"Selected: {os.path.basename(filename)} ({file_type})")

    def browse_output(self):
        """Browse for output folder."""
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_folder_entry.delete(0, tk.END)
            self.output_folder_entry.insert(0, folder)

    def open_output_folder(self):
        """Open output folder in file manager."""
        folder = self.output_folder_entry.get().strip()
        if not folder or not os.path.exists(folder):
            messagebox.showwarning("Warning", "Output folder does not exist")
            return

        if sys.platform.startswith("win"):
            os.startfile(folder)
        elif sys.platform.startswith("darwin"):
            subprocess.Popen(["open", folder])
        else:
            subprocess.Popen(["xdg-open", folder])

    def start_conversion(self):
        """Start conversion process in background thread."""
        input_file = self.input_entry.get().strip()

        if not input_file or not os.path.exists(input_file):
            messagebox.showerror("Error", "Please select a valid input file")
            return

        output_format = self.format_combo.get().lower()
        output_folder = self.output_folder_entry.get().strip()

        if not output_folder:
            output_folder = os.path.dirname(input_file)

        self.convert_button.config(state='disabled')
        self.results_listbox.delete(0, tk.END)

        thread = threading.Thread(
            target=self.conversion_worker,
            args=(input_file, output_format, output_folder),
            daemon=True
        )
        thread.start()

    def conversion_worker(self, input_file, output_format, output_folder):
        """Worker thread for conversion."""
        try:
            file_type = detect_file_type(input_file)
            log_append(self.log_widget, f"Converting {file_type} → {output_format}")

            outputs = self.perform_conversion(input_file, file_type, output_format, output_folder)

            for output in outputs:
                self.results_listbox.insert(tk.END, output)
                self.last_output = output

            log_append(self.log_widget, f"Conversion complete: {len(outputs)} file(s)")
            messagebox.showinfo("Success", f"Created {len(outputs)} file(s)")

        except Exception as e:
            error_msg = str(e)
            log_append(self.log_widget, f"ERROR: {error_msg}")
            log_append(self.log_widget, traceback.format_exc())
            messagebox.showerror("Conversion Failed", error_msg)

        finally:
            self.convert_button.config(state='normal')

    def perform_conversion(self, src, file_type, out_format, out_dir):
        """Perform the actual conversion based on file types."""
        name = safe_filename(src)
        outputs = []

        # CSV conversions
        if file_type == "csv":
            if out_format == "pdf":
                out = os.path.join(out_dir, f"{name}.pdf")
                csv_to_pdf(src, out, self.log_widget)
                outputs.append(out)
            elif out_format in ("png", "jpg", "jpeg"):
                ext = "jpg" if out_format == "jpeg" else out_format
                out = os.path.join(out_dir, f"{name}.{ext}")
                csv_to_image(src, out, self.log_widget)
                outputs.append(out)
            elif out_format == "xlsx":
                out = os.path.join(out_dir, f"{name}.xlsx")
                csv_to_xlsx(src, out, self.log_widget)
                outputs.append(out)
            elif out_format == "csv":
                outputs.append(src)

        # XLSX conversions
        elif file_type in ("xlsx", "xls"):
            if out_format == "pdf":
                out = os.path.join(out_dir, f"{name}.pdf")
                xlsx_to_pdf(src, out, self.log_widget)
                outputs.append(out)
            elif out_format in ("png", "jpg", "jpeg"):
                ext = "jpg" if out_format == "jpeg" else out_format
                out = os.path.join(out_dir, f"{name}.{ext}")
                xlsx_to_image(src, out, self.log_widget)
                outputs.append(out)
            elif out_format == "csv":
                out = os.path.join(out_dir, f"{name}.csv")
                xlsx_to_csv(src, out, self.log_widget)
                outputs.append(out)
            elif out_format == "xlsx":
                outputs.append(src)

        # TXT conversions
        elif file_type == "txt":
            if out_format == "pdf":
                out = os.path.join(out_dir, f"{name}.pdf")
                txt_to_pdf(src, out, self.log_widget)
                outputs.append(out)

        # Image conversions
        elif file_type in ("png", "jpg", "jpeg", "bmp", "gif"):
            if out_format == "pdf":
                out = os.path.join(out_dir, f"{name}.pdf")
                image_to_pdf(src, out, self.log_widget)
                outputs.append(out)
            elif out_format in ("png", "jpg", "jpeg"):
                ext = "jpg" if out_format == "jpeg" else out_format
                out = os.path.join(out_dir, f"{name}.{ext}")
                img = Image.open(src)
                if ext == "jpg" and img.mode == "RGBA":
                    img = img.convert("RGB")
                img.save(out)
                outputs.append(out)

        # DOCX conversions
        elif file_type == "docx":
            if out_format == "pdf":
                out = os.path.join(out_dir, f"{name}.pdf")
                docx_to_pdf(src, out, self.log_widget)
                outputs.append(out)
            elif out_format in ("png", "jpg", "jpeg"):
                tmp_dir = os.path.join(out_dir, f"{name}_temp")
                imgs = docx_to_images(src, tmp_dir, self.log_widget)
                ext = "jpg" if out_format == "jpeg" else out_format

                for i, img_path in enumerate(imgs, 1):
                    img = Image.open(img_path)
                    final = os.path.join(out_dir, f"{name}_page_{i}.{ext}")
                    if ext == "jpg" and img.mode == "RGBA":
                        img = img.convert("RGB")
                    img.save(final)
                    outputs.append(final)

        # PDF conversions
        elif file_type == "pdf":
            if out_format in ("png", "jpg", "jpeg"):
                tmp_dir = os.path.join(out_dir, f"{name}_temp")
                imgs = pdf_to_images(src, tmp_dir, self.log_widget)
                ext = "jpg" if out_format == "jpeg" else out_format

                for i, img_path in enumerate(imgs, 1):
                    img = Image.open(img_path)
                    final = os.path.join(out_dir, f"{name}_page_{i}.{ext}")
                    if ext == "jpg" and img.mode == "RGBA":
                        img = img.convert("RGB")
                    img.save(final)
                    outputs.append(final)
            elif out_format == "pdf":
                outputs.append(src)

        else:
            raise RuntimeError(f"Unsupported file type: {file_type}")

        return outputs

    def inject_abyss_features(self):
        """Inject Abyss toolkit features into tabs."""
        if self._abyss_injected:
            return
        self._abyss_injected = True

        # Add system stats
        if PSUTIL_AVAILABLE:
            self.lbl_cpu = ttk.Label(self.stats_frame, text="CPU: 0%")
            self.lbl_cpu.pack(side='left', padx=5)
            self.lbl_ram = ttk.Label(self.stats_frame, text="RAM: 0%")
            self.lbl_ram.pack(side='left', padx=5)
            ttk.Label(self.stats_frame, text="| ABYSS v1.0", foreground="red").pack(side='left', padx=5)
            self.update_system_stats()

        # Build Compress tab
        self.build_compress_tab()

        # Build Crack tab
        self.build_crack_tab()

        # Build Tools tab
        self.build_tools_tab()

    def update_system_stats(self):
        """Update system statistics."""
        if not PSUTIL_AVAILABLE:
            return

        try:
            cpu = psutil.cpu_percent(interval=0.1)
            ram = psutil.virtual_memory().percent
            self.lbl_cpu.config(text=f"CPU: {cpu:.1f}%")
            self.lbl_ram.config(text=f"RAM: {ram:.1f}%")
        except Exception:
            pass

        self.after(2000, self.update_system_stats)

    def build_compress_tab(self):
        """Build Compress tab UI."""
        for widget in self.tab_compress.winfo_children():
            widget.destroy()

        container = ttk.Frame(self.tab_compress)
        container.pack(fill='both', expand=True, padx=10, pady=10)

        ttk.Label(container, text="Batch PDF Compression Queue", font=("Arial", 12, "bold")).pack(anchor='w', pady=5)

        # File list
        list_frame = ttk.Frame(container)
        list_frame.pack(fill='both', expand=True, pady=5)

        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side='right', fill='y')

        self.compress_listbox = tk.Listbox(list_frame, height=12, yscrollcommand=scrollbar.set)
        self.compress_listbox.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.compress_listbox.yview)

        # Buttons
        btn_frame = ttk.Frame(container)
        btn_frame.pack(fill='x', pady=5)

        def add_files():
            files = filedialog.askopenfilenames(title="Select PDF files", filetypes=[("PDF files", "*.pdf")])
            for f in files:
                self.compress_listbox.insert(tk.END, f)

        def add_folder():
            folder = filedialog.askdirectory(title="Select folder with PDFs")
            if not folder:
                return

            log_append(self.log_widget, f"Scanning for PDFs in: {folder}")
            count = 0
            for root, _, files in os.walk(folder):
                for filename in files:
                    if filename.lower().endswith(".pdf"):
                        full_path = os.path.join(root, filename)
                        self.compress_listbox.insert(tk.END, full_path)
                        count += 1
            log_append(self.log_widget, f"Added {count} PDF(s) from folder.")

        ttk.Button(btn_frame, text="Add Files", command=add_files).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Add Folder", command=add_folder).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Clear List", command=lambda: self.compress_listbox.delete(0, tk.END)).pack(side='left')

        # Quality selection
        quality_frame = ttk.Frame(container)
        quality_frame.pack(fill='x', pady=5)

        ttk.Label(quality_frame, text="Image DPI (lower is smaller):").pack(side='left')
        self.dpi_slider = ttk.Scale(
            quality_frame,
            from_=72,
            to=300,
            orient='horizontal',
            length=200,
        )
        self.dpi_slider.set(150)
        self.dpi_slider.pack(side='left', padx=10, fill='x', expand=True)

        self.dpi_label = ttk.Label(quality_frame, text="150 DPI", width=8)
        self.dpi_label.pack(side='left')

        def update_dpi_label(val):
            self.dpi_label.config(text=f"{int(float(val))} DPI")

        self.dpi_slider.config(command=update_dpi_label)

        # Options frame (for checkbox)
        options_frame = ttk.Frame(container)
        options_frame.pack(fill='x', pady=5)

        self.delete_original_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame,
            text="Delete original files after compression",
            variable=self.delete_original_var
        ).pack(side='left', padx=5)

        # Start button
        def start_compress():
            files = list(self.compress_listbox.get(0, tk.END))
            if not files:
                messagebox.showwarning("No Files", "Please add PDF files to compress")
                return

            delete_original = self.delete_original_var.get()
            if delete_original:
                if not messagebox.askyesno("Confirm Delete", "Are you sure you want to permanently delete the original files after compression? This cannot be undone."):
                    return

            dpi = self.dpi_slider.get()
            threading.Thread(
                target=lambda: compress_pdf_list(files, dpi, self.log_widget, delete_original),
                daemon=True
            ).start()

        ttk.Button(
            container,
            text="START BATCH COMPRESSION",
            command=start_compress
        ).pack(pady=10)

        ttk.Label(container, text="Requires Ghostscript: apt install ghostscript", foreground="gray").pack()

    def build_crack_tab(self):
        """Build Crack tab UI with manual unlock + cracking."""
        for widget in self.tab_crack.winfo_children():
            widget.destroy()

        container = ttk.Frame(self.tab_crack)
        container.pack(fill='both', expand=True, padx=10, pady=10)

        # Header
        ttk.Label(
            container,
            text="PDF Password Tools",
            font=("Arial", 12, "bold")
        ).pack(anchor='w', pady=5)

        # PDF file selection (shared between both tools)
        pdf_frame = ttk.LabelFrame(container, text="Select PDF File")
        pdf_frame.pack(fill='x', pady=10)

        inner_pdf = ttk.Frame(pdf_frame)
        inner_pdf.pack(fill='x', padx=5, pady=5)

        ttk.Label(inner_pdf, text="PDF File:").pack(side='left')
        self.crack_pdf_entry = ttk.Entry(inner_pdf)
        self.crack_pdf_entry.pack(side='left', fill='x', expand=True, padx=8)

        def browse_pdf():
            f = filedialog.askopenfilename(
                title="Select locked PDF",
                filetypes=[("PDF files", "*.pdf")]
            )
            if f:
                self.crack_pdf_entry.delete(0, tk.END)
                self.crack_pdf_entry.insert(0, f)

        ttk.Button(inner_pdf, text="Browse", command=browse_pdf).pack(side='left')

        # ==========================================
        # MANUAL UNLOCK (I KNOW THE PASSWORD)
        # ==========================================

        unlock_frame = ttk.LabelFrame(container, text="Manual Unlock (I Know the Password)")
        unlock_frame.pack(fill='x', pady=10)

        unlock_inner = ttk.Frame(unlock_frame)
        unlock_inner.pack(fill='x', padx=5, pady=5)

        ttk.Label(unlock_inner, text="Password:").pack(side='left', padx=5)

        # Create password entry with show/hide toggle and strength meter
        pwd_frame, self.known_pwd_entry, pwd_toggle, pwd_strength = create_password_entry_with_toggle(
            unlock_inner,
            with_strength=True
        )
        pwd_frame.pack(side='left', fill='x', expand=True, padx=5)

        def unlock_with_known_password():
            pdf = self.crack_pdf_entry.get().strip()
            password = self.known_pwd_entry.get().strip()

            if not pdf or not os.path.exists(pdf):
                messagebox.showerror("Error", "Please select a valid PDF file")
                return

            if not password:
                messagebox.showerror("Error", "Please enter a password")
                return

            log_append(self.log_widget, "Attempting to unlock PDF...")

            def worker():
                try:
                    out_path = unlock_pdf_manual(pdf, password, self.log_widget)
                    if out_path:
                        self.last_output = out_path
                        self.after(0, lambda: messagebox.showinfo(
                            "Success",
                            f"✔ Unlocked PDF saved:\n\n{out_path}"
                        ))
                    else:
                        self.after(0, lambda: messagebox.showerror(
                            "Failed",
                            "❌ Incorrect password or unable to unlock PDF"
                        ))
                except Exception as e:
                    self.after(0, lambda: messagebox.showerror(
                        "Error",
                        f"Unlock failed: {e}"
                    ))

            threading.Thread(target=worker, daemon=True).start()

        ttk.Button(
            unlock_inner,
            text="Unlock PDF",
            command=unlock_with_known_password
        ).pack(side='left', padx=5)

        # ==========================================
        # CRACK PASSWORD (DON'T KNOW PASSWORD)
        # ==========================================

        crack_frame = ttk.LabelFrame(container, text="Crack Password (Don't Know Password)")
        crack_frame.pack(fill='x', pady=10)

        ttk.Label(
            crack_frame,
            text="Uses John the Ripper to crack password-protected PDFs"
        ).pack(anchor='w', padx=5, pady=5)

        # Wordlist selection
        wl_frame = ttk.Frame(crack_frame)
        wl_frame.pack(fill='x', padx=5, pady=5)

        ttk.Label(wl_frame, text="Wordlist (optional):").pack(side='left')
        self.crack_wl_entry = ttk.Entry(wl_frame)
        self.crack_wl_entry.pack(side='left', fill='x', expand=True, padx=8)

        def browse_wl():
            f = filedialog.askopenfilename(
                title="Select wordlist",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
            )
            if f:
                self.crack_wl_entry.delete(0, tk.END)
                self.crack_wl_entry.insert(0, f)

        ttk.Button(wl_frame, text="Browse", command=browse_wl).pack(side='left')

        # Result label
        self.crack_result = ttk.Label(
            crack_frame,
            text="",
            foreground="blue",
            font=("Arial", 10)
        )
        self.crack_result.pack(pady=5)

        # Password result with copy button and strength meter
        result_frame = ttk.Frame(crack_frame)
        result_frame.pack(fill='x', padx=5, pady=5)

        self.cracked_pwd_entry = ttk.Entry(result_frame, state='readonly')
        self.cracked_pwd_entry.pack(side='left', fill='x', expand=True, padx=5)

        def copy_password():
            pwd = self.cracked_pwd_entry.get()
            if pwd:
                self.clipboard_clear()
                self.clipboard_append(pwd)
                messagebox.showinfo("Copied", "Password copied to clipboard")

        self.copy_pwd_btn = ttk.Button(
            result_frame,
            text="Copy",
            command=copy_password,
            state='disabled',
            width=8
        )
        self.copy_pwd_btn.pack(side='left', padx=2)

        self.cracked_pwd_strength = ttk.Label(result_frame, text="", width=12)
        self.cracked_pwd_strength.pack(side='left', padx=5)

        # Start crack button
        def start_crack():
            pdf = self.crack_pdf_entry.get().strip()
            if not pdf or not os.path.exists(pdf):
                messagebox.showerror("Error", "Please select a valid PDF file")
                return

            wl = self.crack_wl_entry.get().strip() or None

            self.crack_result.config(text="⏳ Cracking in progress...", foreground="orange")
            self.cracked_pwd_entry.config(state='normal')
            self.cracked_pwd_entry.delete(0, tk.END)
            self.cracked_pwd_entry.config(state='readonly')
            self.copy_pwd_btn.config(state='disabled')
            self.cracked_pwd_strength.config(text="")

            def worker():
                result = crack_pdf_with_john(pdf, wl, self.log_widget)

                if result:
                    # Update strength meter
                    score, label, color = check_password_strength(result)
                    self.after(0, lambda: self.crack_result.config(
                        text="✔ Password found!",
                        foreground="green"
                    ))
                    self.after(0, lambda: self.cracked_pwd_entry.config(state='normal'))
                    self.after(0, lambda: self.cracked_pwd_entry.delete(0, tk.END))
                    self.after(0, lambda: self.cracked_pwd_entry.insert(0, result))
                    self.after(0, lambda: self.cracked_pwd_entry.config(state='readonly'))
                    self.after(0, lambda: self.copy_pwd_btn.config(state='normal'))
                    self.after(0, lambda: self.cracked_pwd_strength.config(text=label, foreground=color))

                    # Ask if user wants to unlock
                    def ask_unlock():
                        if messagebox.askyesno(
                            "Password Found",
                            f"Password found: {result}\nStrength: {label}\n\nGenerate unlocked PDF?"
                        ):
                            out_path = unlock_pdf_manual(pdf, result, self.log_widget)
                            if out_path:
                                self.last_output = out_path
                                messagebox.showinfo("Success", f"✔ Unlocked PDF saved:\n\n{out_path}")

                    self.after(0, ask_unlock)

                else:
                    self.after(0, lambda: self.crack_result.config(
                        text="❌ Password not found",
                        foreground="red"
                    ))

            threading.Thread(target=worker, daemon=True).start()

        ttk.Button(
            crack_frame,
            text="START ATTACK",
            command=start_crack
        ).pack(pady=10)

        # Requirements note
        req_frame = ttk.Frame(container)
        req_frame.pack(fill='x', pady=10)

        ttk.Label(
            req_frame,
            text="Requirements:",
            font=("Arial", 9, "bold")
        ).pack(anchor='w')

        ttk.Label(
            req_frame,
            text="• Manual Unlock: pip install pypdf",
            foreground="gray"
        ).pack(anchor='w', padx=20)

        ttk.Label(
            req_frame,
            text="• Cracking: pkg install john (Termux) or apt install john (Linux)",
            foreground="gray"
        ).pack(anchor='w', padx=20)

    def build_tools_tab(self):
        """Build Tools tab UI."""
        for widget in self.tab_tools.winfo_children():
            widget.destroy()

        container = ttk.Frame(self.tab_tools)
        container.pack(fill='both', expand=True, padx=10, pady=10)

        ttk.Label(container, text="Additional Tools", font=("Arial", 12, "bold")).pack(anchor='w', pady=5)

        # OCR Tool
        ocr_frame = ttk.LabelFrame(container, text="OCR - Make PDF Searchable")
        ocr_frame.pack(fill='x', pady=10, padx=5)

        ttk.Label(ocr_frame, text="Convert scanned PDFs to searchable text").pack(anchor='w', padx=10, pady=5)

        def run_ocr():
            f = filedialog.askopenfilename(title="Select PDF to OCR", filetypes=[("PDF files", "*.pdf")])
            if not f:
                return

            out = f.replace(".pdf", "_ocr.pdf")

            def worker():
                try:
                    run_ocrmypdf(f, out, self.log_widget)
                    self.last_output = out
                    messagebox.showinfo("Success", f"OCR completed:\n{out}")
                except Exception as e:
                    messagebox.showerror("OCR Failed", str(e))

            threading.Thread(target=worker, daemon=True).start()

        ttk.Button(ocr_frame, text="Run OCR on PDF", command=run_ocr).pack(padx=10, pady=5)

        # PDF Merge Tool
        merge_frame = ttk.LabelFrame(container, text="PDF Merge")
        merge_frame.pack(fill='x', pady=10, padx=5)

        ttk.Label(merge_frame, text="Combine multiple PDFs into one file").pack(anchor='w', padx=10, pady=5)

        def merge_pdfs_ui():
            files = filedialog.askopenfilenames(title="Select PDFs to merge (in order)", filetypes=[("PDF files", "*.pdf")])
            if len(files) < 2:
                messagebox.showwarning("Not Enough Files", "Select at least 2 PDF files")
                return

            out = filedialog.asksaveasfilename(
                title="Save merged PDF as",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")]
            )
            if not out:
                return

            def worker():
                try:
                    merge_pdfs(files, out, self.log_widget)
                    self.last_output = out
                    messagebox.showinfo("Success", f"PDFs merged:\n{out}")
                except Exception as e:
                    messagebox.showerror("Merge Failed", str(e))

            threading.Thread(target=worker, daemon=True).start()

        ttk.Button(merge_frame, text="Merge PDFs", command=merge_pdfs_ui).pack(padx=10, pady=5)

        # Termux Share
        share_frame = ttk.LabelFrame(container, text="Termux Share")
        share_frame.pack(fill='x', pady=10, padx=5)

        ttk.Label(share_frame, text="Share last converted file to Android apps").pack(anchor='w', padx=10, pady=5)

        def share_last():
            if not self.last_output or not os.path.exists(self.last_output):
                messagebox.showwarning("No File", "No recent output file to share")
                return

            termux_share_file(self.last_output, self.log_widget)

        ttk.Button(share_frame, text="Share Last Output", command=share_last).pack(padx=10, pady=5)

        # Package info
        info_frame = ttk.LabelFrame(container, text="Installed Packages")
        info_frame.pack(fill='x', pady=10, padx=5)

        packages = [
            ("Pandas", pd is not None),
            ("ReportLab", REPORTLAB_AVAILABLE),
            ("Pillow", PIL_AVAILABLE),
            ("OpenPyXL", OPENPYXL_AVAILABLE),
            ("python-docx", DOCX_AVAILABLE),
            ("pdf2image", PDF2IMAGE_AVAILABLE),
            ("pypdf", PYPDF_AVAILABLE),
            ("psutil", PSUTIL_AVAILABLE)
        ]

        for name, installed in packages:
            status = "✔ Installed" if installed else "❌ Missing"
            color = "green" if installed else "red"
            lbl = ttk.Label(info_frame, text=f"{name}: {status}", foreground=color)
            lbl.pack(anchor='w', padx=10)


# =====================================================================
# MAIN ENTRY POINT
# =====================================================================

def main():
    """Main entry point."""
    app = ConverterApp()

    # Log startup
    log_append(app.log_widget, "Universal Converter + Abyss Toolkit initialized")
    log_append(app.log_widget, f"Python {sys.version.split()[0]} on {sys.platform}")

    # Check critical dependencies
    if not REPORTLAB_AVAILABLE:
        log_append(app.log_widget, "WARNING: ReportLab not installed - PDF generation disabled")
    if not PIL_AVAILABLE:
        log_append(app.log_widget, "WARNING: Pillow not installed - Image operations disabled")
    if not pd:
        log_append(app.log_widget, "WARNING: Pandas not installed - CSV/XLSX disabled")
    if not PYPDF_AVAILABLE:
        log_append(app.log_widget, "WARNING: pypdf not installed - PDF unlock/merge disabled")

    app.mainloop()


if __name__ == "__main__":
    main()
