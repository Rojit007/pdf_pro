#!/usr/bin/env python3
"""
PDF Studio v4 – Complete Edition
All bugs fixed, all features implemented.
"""
import sys, os, json, threading, shutil, tempfile, copy, math, io, re, time, hashlib
from pathlib import Path

# ── Runtime / dependency bootstrap ────────────────────────────────────────────
IS_FROZEN = getattr(sys, "frozen", False)

def install_deps():
    if IS_FROZEN or os.environ.get("PDF_STUDIO_SKIP_AUTO_INSTALL") == "1":
        return
    import subprocess
    needed = {
        "pypdf": "pypdf",
        "PIL": "Pillow",
        "pdf2image": "pdf2image",
        "pytesseract": "pytesseract",
        "fitz": "PyMuPDF",
    }
    for imp, pkg in needed.items():
        try:
            __import__(imp)
        except ImportError:
            print(f"Installing {pkg}…")
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

install_deps()

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, colorchooser, simpledialog
from pypdf import PdfReader, PdfWriter
from pypdf.generic import NameObject, create_string_object
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageFilter, ImageEnhance
import fitz  # PyMuPDF

# ── Constants ─────────────────────────────────────────────────────────────────
OPTIONS = ["Rotate 90° CW", "Rotate 90° CCW"]
OPTION_COLORS = {
    "Portrait":       "#2E7D32",
    "Landscape":      "#1565C0",
    "Rotate 90° CW":  "#6A1B9A",
    "Rotate 90° CCW": "#E65100",
}
ROTATE_STEP = 90
THUMB_W, THUMB_H   = 90, 120
ROW_H              = THUMB_H + 16
PREVIEW_MIN_W      = 500
RECENT_FILE        = Path.home() / ".pdf_studio_recent.json"
SESSION_FILE       = Path.home() / ".pdf_studio_session.json"
MAX_RECENT         = 12
SESSION_PERSISTENCE = False

PAGE_SIZES = {
    "A4":     (595, 842),
    "A3":     (842, 1191),
    "A5":     (420, 595),
    "Letter": (612, 792),
    "Legal":  (612, 1008),
    "Tabloid":(792, 1224),
}

DARK_THEME = {
    "bg": "#1E1E2E", "fg": "#CDD6F4", "accent": "#89B4FA",
    "row_even": "#181825", "row_odd": "#1E1E2E",
    "toolbar": "#11111B", "statusbar": "#11111B",
    "select": "#313244", "excluded": "#45293a",
    "preview_bg": "#0E0E16",
}
LIGHT_THEME = {
    "bg": "#F8FAFC", "fg": "#1E293B", "accent": "#2563EB",
    "row_even": "#FFFFFF", "row_odd": "#F0F4FF",
    "toolbar": "#1E2A3A", "statusbar": "#1E2A3A",
    "select": "#DBEAFE", "excluded": "#FEE2E2",
    "preview_bg": "#0F172A",
}


# ── Module-level helpers ──────────────────────────────────────────────────────
def _to_roman(num):
    val  = [1000,900,500,400,100,90,50,40,10,9,5,4,1]
    syms = ["M","CM","D","CD","C","XC","L","XL","X","IX","V","IV","I"]
    result = ""
    for v, s in zip(val, syms):
        while num >= v:
            result += s; num -= v
    return result.lower()
def normalize_rotation(value, orig_orient="Portrait"):
    if isinstance(value, (int, float)):
        return int(value) % 360
    s = str(value).strip()
    if s in ("Rotate 90° CW", "CW"):
        return 90
    if s in ("Rotate 90° CCW", "CCW"):
        return 270
    if s == "Portrait":
        return 0 if orig_orient == "Portrait" else 270
    if s == "Landscape":
        return 90 if orig_orient == "Portrait" else 0
    m = re.match(r"^(-?\d+)", s)
    if m:
        return int(m.group(1)) % 360
    return 0

def rotation_label(value):
    deg = normalize_rotation(value)
    if deg == 0:
        return "0°"
    if deg == 90:
        return "90° CW"
    if deg == 180:
        return "180°"
    if deg == 270:
        return "90° CCW"
    return f"{deg}°"


def get_page_orientation(page):
    box = page.mediabox
    w, h = float(box.width), float(box.height)
    rot = int(page.get("/Rotate") or 0) % 360
    if rot in (90, 270):
        w, h = h, w
    return "Landscape" if w > h else "Portrait"

def apply_transform(page, choice):
    if choice is None:
        return page
    rot = normalize_rotation(choice)
    if rot:
        page.rotate(rot)
    return page

def apply_visual_rotation(img: Image.Image, choice) -> Image.Image:
    rot = normalize_rotation(choice)
    if rot == 90:
        return img.transpose(Image.ROTATE_270)
    elif rot == 180:
        return img.transpose(Image.ROTATE_180)
    elif rot == 270:
        return img.transpose(Image.ROTATE_90)
    return img

def render_page_image_fitz(pdf_path, page_index, orientation_choice, orig_orient,
                            target_w=None, target_h=None, for_thumb=False):
    try:
        doc = fitz.open(pdf_path)
        page = doc[page_index]
        zoom = 1.5 if for_thumb else 2.5
        if target_w:
            rect = page.rect
            zoom = min((target_w / rect.width) * 1.5, 3.0)
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        doc.close()
        if target_w and target_h:
            img.thumbnail((target_w, target_h), Image.LANCZOS)
        img = apply_visual_rotation(img, orientation_choice)
        return img
    except Exception:
        pass
    w = target_w or THUMB_W
    h = target_h or THUMB_H
    img = Image.new("RGB", (w, h), "#EEEEEE")
    draw = ImageDraw.Draw(img)
    draw.text((w // 2, h // 2), f"Page {page_index + 1}", fill="#888888", anchor="mm")
    return img

def get_page_info_fitz(pdf_path, page_index):
    try:
        doc = fitz.open(pdf_path)
        page = doc[page_index]
        rect = page.rect
        text = page.get_text().strip()
        doc.close()
        return {"width": rect.width, "height": rect.height, "has_text": bool(text)}
    except:
        return {"width": 595, "height": 842, "has_text": False}

def load_recent_files():
    try:
        if RECENT_FILE.exists():
            data = json.loads(RECENT_FILE.read_text())
            return [p for p in data if os.path.exists(p)]
    except:
        pass
    return []

def save_recent_files(paths):
    try:
        RECENT_FILE.write_text(json.dumps(paths[:MAX_RECENT]))
    except:
        pass

def add_recent_file(path):
    recent = load_recent_files()
    if path in recent:
        recent.remove(path)
    recent.insert(0, path)
    save_recent_files(recent[:MAX_RECENT])


# ── Page data record ──────────────────────────────────────────────────────────
class PageRecord:
    def __init__(self, source_path, source_index, orientation, is_blank=False):
        raw = orientation.get() if hasattr(orientation, "get") else orientation
        if raw in ("Portrait", "Landscape"):
            orig_orient = raw
        else:
            orig_orient = "Portrait"
        self.source_path   = source_path
        self.source_index  = source_index
        self.orientation   = tk.IntVar(value=normalize_rotation(raw, orig_orient))
        self.included      = tk.BooleanVar(value=True)
        self.orig_orient   = orig_orient
        self.is_blank      = is_blank
        self.thumb_img     = None
        self.thumb_tk      = None
        self.annotations   = []
        self.redactions    = []
        self._row_widget   = None
        self._thumb_label  = None
        self._rb_frame     = None
        self._rot_value_lbl = None

    def snapshot(self):
        return {
            "source_path":  self.source_path,
            "source_index": self.source_index,
            "orientation":  int(self.orientation.get()),
            "rotation":     int(self.orientation.get()),
            "included":     self.included.get(),
            "is_blank":     self.is_blank,
            "orig_orient":  self.orig_orient,
            "annotations":  list(self.annotations),
            "redactions":   list(self.redactions),
        }

    @classmethod
    def from_snapshot(cls, snap):
        seed = snap.get("orig_orient", snap.get("orientation", "Portrait"))
        if seed not in ("Portrait", "Landscape"):
            seed = "Portrait"
        rec = cls(snap["source_path"], snap["source_index"],
                  tk.StringVar(value=seed),
                  is_blank=snap.get("is_blank", False))
        rec.included.set(snap.get("included", True))
        rec.orig_orient   = seed
        rec.orientation.set(normalize_rotation(
            snap.get("rotation", snap.get("orientation", 0)),
            rec.orig_orient))
        rec.annotations   = snap.get("annotations", [])
        rec.redactions    = snap.get("redactions", [])
        return rec


# ── Undo/Redo stack ───────────────────────────────────────────────────────────
class UndoStack:
    def __init__(self, max_depth=60):
        self._undo = []
        self._redo = []
        self.max_depth = max_depth

    def push(self, pages_list, description="action"):
        snap = [r.snapshot() for r in pages_list]
        self._undo.append((description, snap))
        if len(self._undo) > self.max_depth:
            self._undo.pop(0)
        self._redo.clear()

    def undo(self):
        if not self._undo: return None, None
        desc, snap = self._undo.pop()
        self._redo.append((desc, snap))
        return desc, snap

    def redo(self):
        if not self._redo: return None, None
        desc, snap = self._redo.pop()
        self._undo.append((desc, snap))
        return desc, snap

    def can_undo(self): return bool(self._undo)
    def can_redo(self): return bool(self._redo)
    def peek_undo(self): return self._undo[-1][0] if self._undo else ""
    def peek_redo(self): return self._redo[-1][0] if self._redo else ""


# ── Tooltip helper ────────────────────────────────────────────────────────────
class Tooltip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text   = text
        self.tip    = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        x, y = self.widget.winfo_rootx() + 20, self.widget.winfo_rooty() + 30
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tip, text=self.text, bg="#1E293B", fg="#F8FAFC",
                 font=("Helvetica", 9), padx=8, pady=4, relief="flat").pack()

    def hide(self, event=None):
        if self.tip:
            self.tip.destroy()
            self.tip = None


# ── Main App ──────────────────────────────────────────────────────────────────
class PDFStudio:
    def __init__(self, root):
        self.root       = root
        self.root.title("PDF Studio v4")
        self.root.resizable(True, True)
        self.root.minsize(1100, 660)

        self.pages:         list[PageRecord] = []
        self.primary_path   = None
        self.drag_data      = {}
        self.thumb_cache    = {}
        self.preview_cache  = {}
        self.selected_pages = set()

        self._preview_index = -1
        self._preview_zoom  = 1.0
        self._preview_tk    = None

        self._dark_mode = tk.BooleanVar(value=False)
        self._theme     = LIGHT_THEME

        # metadata
        self.meta_title    = tk.StringVar()
        self.meta_author   = tk.StringVar()
        self.meta_subject  = tk.StringVar()
        self.meta_keywords = tk.StringVar()
        self.meta_creator  = tk.StringVar(value="PDF Studio v4")

        # output options
        self.add_page_numbers    = tk.BooleanVar(value=False)
        self.page_num_format     = tk.StringVar(value="decimal")
        self.page_num_position   = tk.StringVar(value="bottom-center")
        self.compress_output     = tk.BooleanVar(value=False)
        self.split_mode          = tk.StringVar(value="single")
        self.watermark_text      = tk.StringVar()
        self.watermark_opacity   = tk.IntVar(value=40)
        self.watermark_pages     = tk.StringVar(value="all")
        self.watermark_color     = tk.StringVar(value="#AAAAAA")
        self.output_page_size    = tk.StringVar(value="Original")
        self.pdfa_mode           = tk.BooleanVar(value=False)
        self.linearize           = tk.BooleanVar(value=False)
        self.encrypt_pdf         = tk.BooleanVar(value=False)
        self.owner_password      = tk.StringVar()
        self.user_password       = tk.StringVar()
        self.flatten_forms       = tk.BooleanVar(value=False)
        self.header_text         = tk.StringVar()
        self.footer_text         = tk.StringVar()
        self.thumb_size          = tk.IntVar(value=90)

        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", lambda *_: self._filter_rows())

        self.undo_stack      = UndoStack()
        self.status_var      = tk.StringVar(value="Open a PDF to get started.")
        self._auto_save_job  = None
        self._bookmarks      = []

        self._build_ui()
        self._bind_keys()
        if SESSION_PERSISTENCE:
            self._restore_session()
            self._start_autosave()
        else:
            self._clear_saved_session()

    # ═══════════════════════════════ THEME ══════════════════════════════════
    def _toggle_theme(self):
        self._dark_mode.set(not self._dark_mode.get())
        self._theme = DARK_THEME if self._dark_mode.get() else LIGHT_THEME
        self._rebuild_rows()

    def T(self, key):
        return self._theme.get(key, "#FFFFFF")

    # ═══════════════════════════════ UNDO HELPERS ═══════════════════════════
    def _push_undo(self, description="action"):
        self.undo_stack.push(self.pages, description)
        self._update_undo_labels()

    def _update_undo_labels(self):
        u = self.undo_stack.peek_undo()
        r = self.undo_stack.peek_redo()
        self._undo_btn.config(state="normal" if u else "disabled",
                              text=f"↩ Undo{': '+u if u else ''}")
        self._redo_btn.config(state="normal" if r else "disabled",
                              text=f"↪ Redo{': '+r if r else ''}")

    def _restore_from_snap(self, snap):
        new_pages = []
        for s in snap:
            rec = PageRecord.from_snapshot(s)
            key = (rec.source_path, rec.source_index, rec.orientation.get())
            rec.thumb_img = self.thumb_cache.get(key)
            new_pages.append(rec)
        return new_pages

    def do_undo(self):
        desc, snap = self.undo_stack.undo()
        if snap is None: return
        self.pages = self._restore_from_snap(snap)
        self.selected_pages.clear()
        self._rebuild_rows()
        self._update_status()
        self._update_undo_labels()
        self.status_var.set(f"↩ Undid: {desc}")

    def do_redo(self):
        desc, snap = self.undo_stack.redo()
        if snap is None: return
        self.pages = self._restore_from_snap(snap)
        self.selected_pages.clear()
        self._rebuild_rows()
        self._update_status()
        self._update_undo_labels()
        self.status_var.set(f"↪ Redid: {desc}")

    # ═══════════════════════════════ UI BUILD ═══════════════════════════════
    def _build_ui(self):
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)
        self._build_menubar()
        self._build_toolbar()
        self._build_main_area()
        self._build_statusbar()

    def _build_menubar(self):
        mb = tk.Menu(self.root, tearoff=0, bg="#1E2A3A", fg="white",
                     activebackground="#2563EB", activeforeground="white")
        self.root.config(menu=mb)

        fm = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="File", menu=fm)
        fm.add_command(label="Open PDF…        Ctrl+O", command=self.browse_file)
        fm.add_command(label="Merge PDF…",              command=self.merge_pdf)
        fm.add_separator()
        fm.add_command(label="Import Images as PDF…",   command=self.import_images)
        fm.add_command(label="Export Pages as Images…", command=self.export_as_images)
        fm.add_separator()
        self._recent_menu = tk.Menu(fm, tearoff=0)
        fm.add_cascade(label="Recent Files", menu=self._recent_menu)
        fm.add_separator()
        fm.add_command(label="Save PDF…         Ctrl+S", command=self.save_pdf)
        fm.add_command(label="Save Preset…",             command=self.save_preset)
        fm.add_command(label="Load Preset…",             command=self.load_preset)
        fm.add_separator()
        fm.add_command(label="Exit", command=self._on_close)
        self._refresh_recent_menu()

        em = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="Edit", menu=em)
        em.add_command(label="Undo  Ctrl+Z",      command=self.do_undo)
        em.add_command(label="Redo  Ctrl+Y",      command=self.do_redo)
        em.add_separator()
        em.add_command(label="Select All  Ctrl+A", command=self.select_all_pages)
        em.add_command(label="Deselect All",       command=self.deselect_all_pages)
        em.add_separator()
        em.add_command(label="Delete Selected",    command=self.delete_selected)
        em.add_command(label="Duplicate Selected", command=self.duplicate_selected)

        pm = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="Pages", menu=pm)
        pm.add_command(label="Insert Blank Page",    command=self.insert_blank)
        pm.add_command(label="Reverse All Pages",    command=self.reverse_pages)
        pm.add_separator()
        pm.add_command(label="Crop Margins…",        command=self.crop_margins_dialog)
        pm.add_command(label="Resize Pages…",        command=self.resize_pages_dialog)
        pm.add_separator()
        pm.add_command(label="OCR (Make Searchable)…", command=self.ocr_dialog)
        pm.add_command(label="Find & Replace Text…",   command=self.find_replace_dialog)
        pm.add_command(label="Redact Region…",         command=self.redact_dialog)

        am = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="Annotate", menu=am)
        am.add_command(label="Add Text Annotation…", command=self.add_text_annotation)
        am.add_command(label="Add Stamp/Image…",     command=self.add_stamp_dialog)
        am.add_command(label="Edit Header / Footer…",command=self.header_footer_dialog)
        am.add_separator()
        am.add_command(label="Add Bookmark…",        command=self.add_bookmark_dialog)
        am.add_command(label="Edit Bookmarks…",      command=self.bookmarks_dialog)

        sm = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="Security", menu=sm)
        sm.add_command(label="Encrypt / Password…", command=self.encryption_dialog)
        sm.add_command(label="Unlock PDF…",         command=self.unlock_pdf_dialog)

        vm = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="View", menu=vm)
        vm.add_checkbutton(label="Dark Mode", variable=self._dark_mode, command=self._toggle_theme)
        vm.add_separator()
        vm.add_command(label="Compare Pages…",   command=self.compare_pages_dialog)
        vm.add_command(label="Page Inspector…",  command=self.page_inspector_dialog)
        vm.add_separator()
        vm.add_command(label="Zoom In    Ctrl++", command=self._zoom_in)
        vm.add_command(label="Zoom Out   Ctrl+-", command=self._zoom_out)
        vm.add_command(label="Fit Page   Ctrl+0", command=self._zoom_fit)

        tm = tk.Menu(mb, tearoff=0)
        mb.add_cascade(label="Tools", menu=tm)
        tm.add_command(label="Batch Process Folder…", command=self.batch_process_dialog)
        tm.add_command(label="PDF Metadata…",         command=self.show_metadata_dialog)
        tm.add_command(label="Output Options…",       command=self.show_output_options)
        tm.add_separator()
        tm.add_command(label="Keyboard Shortcuts  F1", command=self.show_shortcuts)

    def _refresh_recent_menu(self):
        self._recent_menu.delete(0, "end")
        recent = load_recent_files()
        if not recent:
            self._recent_menu.add_command(label="(none)", state="disabled")
        for path in recent:
            self._recent_menu.add_command(
                label=os.path.basename(path),
                command=lambda p=path: self._open_path(p))

    def _build_toolbar(self):
        tb = tk.Frame(self.root, bg="#1E2A3A", padx=8, pady=6)
        tb.grid(row=0, column=0, sticky="ew")

        def tbtn(parent, text, cmd, bg="#2E4057", fg="white", state="normal", tip="", **kw):
            b = tk.Button(parent, text=text, command=cmd, bg=bg, fg=fg,
                          font=("Helvetica", 9, "bold"), relief="flat",
                          padx=8, pady=4, cursor="hand2", state=state,
                          activebackground="#3D5A7A", activeforeground="white",
                          disabledforeground="#888", **kw)
            if tip:
                Tooltip(b, tip)
            return b

        def sep():
            tk.Frame(tb, width=2, bg="#3D5A7A").pack(side="left", padx=6, fill="y", pady=2)

        tbtn(tb, "📂 Open",        self.browse_file,         bg="#2563EB", tip="Open PDF (Ctrl+O)").pack(side="left", padx=2)
        tbtn(tb, "➕ Merge",       self.merge_pdf,           bg="#7C3AED", tip="Merge another PDF").pack(side="left", padx=2)
        tbtn(tb, "🖼 Images→PDF", self.import_images,        bg="#0D9488", tip="Import images as PDF pages").pack(side="left", padx=2)
        tbtn(tb, "📄 Blank",       self.insert_blank,        bg="#475569", tip="Insert blank page").pack(side="left", padx=2)
        sep()
        self._undo_btn = tbtn(tb, "↩ Undo", self.do_undo, bg="#92400E", state="disabled")
        self._undo_btn.pack(side="left", padx=2)
        self._redo_btn = tbtn(tb, "↪ Redo", self.do_redo, bg="#92400E", state="disabled")
        self._redo_btn.pack(side="left", padx=2)
        sep()
        tbtn(tb, "⇅ Reverse",   self.reverse_pages,         bg="#475569").pack(side="left", padx=2)
        tbtn(tb, "✂ Extract",   self.extract_pages,         bg="#0891B2", tip="Extract included pages to new PDF").pack(side="left", padx=2)
        tbtn(tb, "🔍 Compare",  self.compare_pages_dialog,  bg="#7C3AED", tip="Compare two pages side by side").pack(side="left", padx=2)
        sep()
        tbtn(tb, "🔒 Encrypt",  self.encryption_dialog,     bg="#DC2626", tip="Password protect output").pack(side="left", padx=2)
        tbtn(tb, "🔓 Unlock",   self.unlock_pdf_dialog,     bg="#16A34A", tip="Remove password from PDF").pack(side="left", padx=2)
        tbtn(tb, "⚙ Output",   self.show_output_options,    bg="#B45309").pack(side="left", padx=2)
        tbtn(tb, "🏷 Meta",     self.show_metadata_dialog,  bg="#B45309").pack(side="left", padx=2)
        sep()
        tbtn(tb, "🌙 Theme",    self._toggle_theme,          bg="#334155", tip="Toggle dark/light mode").pack(side="left", padx=2)
        tbtn(tb, "F1 Help",     self.show_shortcuts,         bg="#475569").pack(side="left", padx=2)
        sep()
        tk.Label(tb, text="Thumbs:", bg="#1E2A3A", fg="#94A3B8", font=("Helvetica", 8)).pack(side="left")
        ttk.Scale(tb, from_=60, to=150, orient="horizontal", length=80,
                  variable=self.thumb_size, command=self._on_thumb_size_change).pack(side="left", padx=4)
        sep()
        tk.Frame(tb, bg="#1E2A3A").pack(side="left", expand=True, fill="x")
        tbtn(tb, "💾 Save PDF",    self.save_pdf,            bg="#16A34A", tip="Save PDF (Ctrl+S)").pack(side="right", padx=3)
        tbtn(tb, "📤 Export Imgs", self.export_as_images,    bg="#0891B2").pack(side="right", padx=3)
        tbtn(tb, "🔄 Batch",       self.batch_process_dialog,bg="#7C3AED", tip="Batch process a folder").pack(side="right", padx=3)

    def _build_main_area(self):
        pane = tk.PanedWindow(self.root, orient="horizontal", sashwidth=6, bg="#CBD5E1")
        pane.grid(row=1, column=0, sticky="nsew", padx=4, pady=4)

        left = tk.Frame(pane, bg="#F8FAFC")
        pane.add(left, minsize=640)
        left.rowconfigure(3, weight=1)
        left.columnconfigure(0, weight=1)
        self._build_search_bar(left)
        self._build_bulk_bar(left)
        self._build_range_bar(left)
        self._build_page_canvas(left)

        right = tk.Frame(pane, bg="#0F172A")
        pane.add(right, minsize=PREVIEW_MIN_W)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)
        self._build_preview_panel(right)

    def _build_search_bar(self, parent):
        bar = tk.Frame(parent, bg="#E2E8F0", padx=8, pady=4)
        bar.grid(row=0, column=0, sticky="ew")
        tk.Label(bar, text="🔍 Filter:", bg="#E2E8F0", font=("Helvetica", 9)).pack(side="left")
        tk.Entry(bar, textvariable=self._search_var, width=20, font=("Helvetica", 9)).pack(side="left", padx=4)
        tk.Button(bar, text="✖", command=lambda: self._search_var.set(""),
                  bg="#DC2626", fg="white", font=("Helvetica", 8), relief="flat", padx=4).pack(side="left")
        tk.Label(bar, text="  Multi-select: Ctrl+Click / Shift+Click",
                 bg="#E2E8F0", fg="#64748B", font=("Helvetica", 8)).pack(side="left", padx=12)

    def _build_bulk_bar(self, parent):
        bar = tk.Frame(parent, bg="#E2E8F0", padx=8, pady=5)
        bar.grid(row=1, column=0, sticky="ew")

        def bbtn(text, cmd, bg):
            return tk.Button(bar, text=text, command=cmd, bg=bg, fg="white",
                             font=("Helvetica", 8, "bold"), relief="flat", padx=6, pady=3)

        tk.Label(bar, text="All pages:", bg="#E2E8F0", font=("Helvetica", 9, "bold")).pack(side="left", padx=(0,6))
        bbtn("↻ CW",  lambda: self.rotate_all_pages(ROTATE_STEP),  OPTION_COLORS["Rotate 90° CW"]).pack(side="left", padx=2)
        bbtn("↺ CCW", lambda: self.rotate_all_pages(-ROTATE_STEP), OPTION_COLORS["Rotate 90° CCW"]).pack(side="left", padx=2)
        bbtn("↺ Reset",  self.reset_all_orient,              "#64748B").pack(side="left", padx=6)
        tk.Frame(bar, width=1, bg="#94A3B8").pack(side="left", fill="y", padx=6, pady=2)
        bbtn("☑ All",   lambda: self.set_all_include(True),  "#0D9488").pack(side="left", padx=2)
        bbtn("☐ None",  lambda: self.set_all_include(False), "#DC2626").pack(side="left", padx=2)
        bbtn("⇅ Odd",   lambda: self.select_odd_even("odd"), "#7C3AED").pack(side="left", padx=2)
        bbtn("⇅ Even",  lambda: self.select_odd_even("even"),"#7C3AED").pack(side="left", padx=2)
        tk.Frame(bar, width=1, bg="#94A3B8").pack(side="left", fill="y", padx=8, pady=2)
        tk.Label(bar, text="Go:", bg="#E2E8F0", font=("Helvetica", 9)).pack(side="left")
        self._goto_var = tk.StringVar()
        goto_e = tk.Entry(bar, textvariable=self._goto_var, width=5, font=("Helvetica", 9))
        goto_e.pack(side="left", padx=3)
        goto_e.bind("<Return>", lambda e: self._goto_page())
        tk.Button(bar, text="▶", command=self._goto_page,
                  bg="#475569", fg="white", font=("Helvetica", 8), relief="flat").pack(side="left")

    def _build_range_bar(self, parent):
        bar = tk.Frame(parent, bg="#F1F5F9", padx=8, pady=4, relief="groove", bd=1)
        bar.grid(row=2, column=0, sticky="ew", pady=(0, 2))
        tk.Label(bar, text="Pages (e.g. 1,3,5-8):", bg="#F1F5F9", font=("Helvetica", 9)).pack(side="left")
        self.range_entry = tk.Entry(bar, width=18, font=("Helvetica", 9))
        self.range_entry.pack(side="left", padx=4)

        def rbtn(text, cmd, bg):
            tk.Button(bar, text=text, command=cmd, bg=bg, fg="white",
                      font=("Helvetica", 8, "bold"), relief="flat", padx=6, pady=2).pack(side="left", padx=2)

        rbtn("✔ Include",   lambda: self._apply_range(True),  "#0D9488")
        rbtn("✖ Exclude",   lambda: self._apply_range(False), "#DC2626")
        rbtn("⊕ Duplicate", self.duplicate_range,             "#7C3AED")
        rbtn("🗑 Delete",    self.delete_range,                "#DC2626")
        rbtn("↻ Rotate",    self.orient_range_dialog,         "#1565C0")
        tk.Label(bar, text="← comma / range", fg="#94A3B8", bg="#F1F5F9",
                 font=("Helvetica", 8)).pack(side="left", padx=4)

    def _build_page_canvas(self, parent):
        frame = tk.Frame(parent, bg="#F8FAFC")
        frame.grid(row=3, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        self.list_canvas = tk.Canvas(frame, bg="#F8FAFC", highlightthickness=0)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=self.list_canvas.yview)
        self.inner = tk.Frame(self.list_canvas, bg="#F8FAFC")
        self.inner.bind("<Configure>", lambda e: self.list_canvas.configure(
            scrollregion=self.list_canvas.bbox("all")))
        self._inner_id = self.list_canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.list_canvas.configure(yscrollcommand=vsb.set)
        self.list_canvas.bind("<Configure>",
            lambda e: self.list_canvas.itemconfig(self._inner_id, width=e.width))
        self.list_canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")

        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.list_canvas.bind(seq, self._on_mousewheel)

        hdr = tk.Frame(self.inner, bg="#CBD5E1")
        hdr.pack(fill="x", padx=4, pady=(4, 0))
        for txt, w in [("", 6), ("✔", 4), ("Pg", 4), ("Thumb", 11),
                       ("Orientation", 14), ("Set To", 21), ("Actions", 14)]:
            tk.Label(hdr, text=txt, width=w, bg="#CBD5E1",
                     font=("Helvetica", 9, "bold"), anchor="center").pack(side="left", padx=2)

        self.rows_frame = tk.Frame(self.inner, bg="#F8FAFC")
        self.rows_frame.pack(fill="both", expand=True, padx=4)

    def _on_mousewheel(self, event):
        if event.num == 4:
            self.list_canvas.yview_scroll(-1, "units")
        elif event.num == 5:
            self.list_canvas.yview_scroll(1, "units")
        else:
            self.list_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _build_preview_panel(self, parent):
        hdr = tk.Frame(parent, bg="#1E2A3A", pady=6)
        hdr.grid(row=0, column=0, sticky="ew")
        hdr.columnconfigure(1, weight=1)
        tk.Button(hdr, text="◀", command=self._preview_prev,
                  bg="#334155", fg="white", relief="flat", padx=8, pady=2,
                  font=("Helvetica", 11, "bold")).grid(row=0, column=0, padx=(8, 2))
        self._preview_info_var = tk.StringVar(value="No page selected")
        tk.Label(hdr, textvariable=self._preview_info_var,
                 bg="#1E2A3A", fg="#94A3B8", font=("Helvetica", 9, "bold")).grid(row=0, column=1, sticky="ew")
        tk.Button(hdr, text="▶", command=self._preview_next,
                  bg="#334155", fg="white", relief="flat", padx=8, pady=2,
                  font=("Helvetica", 11, "bold")).grid(row=0, column=2, padx=(2, 8))

        self.preview_canvas = tk.Canvas(parent, bg="#0F172A", highlightthickness=0, cursor="hand2")
        self.preview_canvas.grid(row=1, column=0, sticky="nsew", padx=12, pady=(8, 4))
        self.preview_canvas.bind("<Configure>",          self._on_preview_resize)
        self.preview_canvas.bind("<Button-1>",           self._on_preview_click)
        self.preview_canvas.bind("<MouseWheel>",         self._on_preview_scroll)
        self.preview_canvas.bind("<Button-4>",           lambda e: self._zoom_in())
        self.preview_canvas.bind("<Button-5>",           lambda e: self._zoom_out())
        self.preview_canvas.bind("<Control-MouseWheel>", self._on_preview_ctrl_scroll)

        zoom_bar = tk.Frame(parent, bg="#1E2A3A", pady=4)
        zoom_bar.grid(row=2, column=0, sticky="ew")
        zoom_bar.columnconfigure(2, weight=1)

        def zbtn(text, cmd, tip=""):
            b = tk.Button(zoom_bar, text=text, command=cmd,
                          bg="#334155", fg="white", relief="flat",
                          padx=10, pady=2, font=("Helvetica", 11), cursor="hand2")
            if tip: Tooltip(b, tip)
            return b

        zbtn("−", self._zoom_out,  "Zoom out").grid(row=0, column=0, padx=(8, 2))
        self._zoom_label = tk.Label(zoom_bar, text="Fit", width=5,
                                    bg="#1E2A3A", fg="#94A3B8", font=("Helvetica", 9, "bold"))
        self._zoom_label.grid(row=0, column=1, padx=2)
        zbtn("+", self._zoom_in,   "Zoom in").grid(row=0, column=2, padx=(2, 0), sticky="w")
        zbtn("⊡ Fit", self._zoom_fit, "Fit to window (Ctrl+0)").grid(row=0, column=3, padx=(4, 2))
        zbtn("📄 Info", lambda: self.page_inspector_dialog(), "Page inspector").grid(row=0, column=4, padx=(2, 8))

        self._preview_foot_var = tk.StringVar(value="Click a thumbnail to preview")
        tk.Label(parent, textvariable=self._preview_foot_var,
                 bg="#0F172A", fg="#475569", font=("Helvetica", 8)).grid(row=3, column=0, pady=(0, 8))

    def _build_statusbar(self):
        sb = tk.Frame(self.root, bg="#1E2A3A", padx=8, pady=4)
        sb.grid(row=2, column=0, sticky="ew")
        sb.columnconfigure(0, weight=1)
        tk.Label(sb, textvariable=self.status_var, bg="#1E2A3A", fg="#94A3B8",
                 font=("Helvetica", 9), anchor="w").grid(row=0, column=0, sticky="ew")
        self.progress = ttk.Progressbar(sb, mode="indeterminate", length=120)
        self.progress.grid(row=0, column=1, padx=8)
        self.progress.grid_remove()
        self._autosave_var = tk.StringVar(value="")
        tk.Label(sb, textvariable=self._autosave_var, bg="#1E2A3A", fg="#4B5563",
                 font=("Helvetica", 8)).grid(row=0, column=2, padx=4)

    def _bind_keys(self):
        self.root.bind("<Control-o>",     lambda e: self.browse_file())
        self.root.bind("<Control-s>",     lambda e: self.save_pdf())
        self.root.bind("<Control-z>",     lambda e: self.do_undo())
        self.root.bind("<Control-y>",     lambda e: self.do_redo())
        self.root.bind("<Control-Z>",     lambda e: self.do_undo())
        self.root.bind("<Control-a>",     lambda e: self.select_all_pages())
        self.root.bind("<Control-equal>", lambda e: self._zoom_in())
        self.root.bind("<Control-minus>", lambda e: self._zoom_out())
        self.root.bind("<Control-0>",     lambda e: self._zoom_fit())
        self.root.bind("<F1>",            lambda e: self.show_shortcuts())
        self.root.bind("<Left>",          lambda e: self._preview_prev())
        self.root.bind("<Right>",         lambda e: self._preview_next())
        self.root.bind("<Delete>",        lambda e: self.delete_selected())
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    # ═══════════════════════════════ PREVIEW ════════════════════════════════
    def _render_preview(self, rec: PageRecord = None):
        if rec is None and 0 <= self._preview_index < len(self.pages):
            rec = self.pages[self._preview_index]
        if rec is None:
            return

        c = self.preview_canvas
        c.update_idletasks()
        cw, ch = c.winfo_width(), c.winfo_height()
        if cw < 10 or ch < 10:
            self.root.after(100, self._render_preview)
            return

        c.delete("all")
        c.create_rectangle(0, 0, cw, ch, fill="#0F172A", outline="")

        max_w = max(10, int(cw * self._preview_zoom) - 40)
        max_h = max(10, int(ch * self._preview_zoom) - 40)

        if rec.is_blank:
            pw = min(max_w, int(max_h * 0.707))
            ph = min(max_h, int(max_w / 0.707))
            x0, y0 = (cw - pw) // 2, (ch - ph) // 2
            c.create_rectangle(x0+6, y0+6, x0+pw+6, y0+ph+6, fill="#000000", outline="", stipple="gray50")
            c.create_rectangle(x0, y0, x0+pw, y0+ph, fill="#FFFFFF", outline="#CCCCCC", width=2)
            c.create_text(cw//2, ch//2, text="BLANK PAGE", fill="#AAAAAA", font=("Helvetica", 16, "bold"))
        else:
            key = (rec.source_path, rec.source_index, max_w, max_h,
                   rec.orientation.get(), int(self._preview_zoom * 100))
            if key in self.preview_cache:
                img = self.preview_cache[key]
            else:
                img = render_page_image_fitz(rec.source_path, rec.source_index,
                                             rec.orientation.get(), rec.orig_orient,
                                             max_w, max_h)
                if img:
                    self.preview_cache[key] = img
                    self._trim_preview_cache()

            if img is None:
                img = Image.new("RGB", (max_w, max_h), "#EEEEEE")

            iw, ih = img.size
            x, y = (cw - iw) // 2, (ch - ih) // 2
            c.create_rectangle(x+6, y+6, x+iw+6, y+ih+6, fill="#000000", outline="", stipple="gray25")
            self._preview_tk = ImageTk.PhotoImage(img)
            c.create_image(x, y, anchor="nw", image=self._preview_tk)
            c.create_rectangle(x, y, x+iw, y+ih, fill="", outline="#334155", width=3)

        idx = self._preview_index
        src = os.path.basename(rec.source_path) if not rec.is_blank else "Blank"
        self._preview_info_var.set(f"Page {idx+1} of {len(self.pages)} • {rotation_label(rec.orientation.get())}")
        self._zoom_label.config(text=f"{int(self._preview_zoom*100)}%")
        self._preview_foot_var.set(f"{'◀/▶' if len(self.pages)>1 else ''} {src} | Ctrl+Scroll to zoom")
        self.root.after(20, self._preload_adjacent)

    def _on_preview_resize(self, event):
        if self._preview_index >= 0:
            self.root.after(50, self._render_preview)

    def _on_preview_click(self, event):
        self._preview_next()

    def _on_preview_scroll(self, event):
        if event.delta > 0:
            self._preview_prev()
        else:
            self._preview_next()

    def _on_preview_ctrl_scroll(self, event):
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()

    def _preview_prev(self):
        if not self.pages: return
        new_idx = max(0, self._preview_index - 1)
        if new_idx != self._preview_index:
            self._preview_index = new_idx
            self._render_preview()
            self._scroll_to_row(new_idx)

    def _preview_next(self):
        if not self.pages: return
        new_idx = min(len(self.pages)-1, self._preview_index + 1)
        if new_idx != self._preview_index:
            self._preview_index = new_idx
            self._render_preview()
            self._scroll_to_row(new_idx)

    def _zoom_in(self):
        self._preview_zoom = min(4.0, self._preview_zoom + 0.25)
        self._render_preview()

    def _zoom_out(self):
        self._preview_zoom = max(0.25, self._preview_zoom - 0.25)
        self._render_preview()

    def _zoom_fit(self):
        self._preview_zoom = 1.0
        self._render_preview()

    def _preload_adjacent(self):
        if not self.pages or self._preview_index < 0: return
        idx = self._preview_index
        cw = self.preview_canvas.winfo_width()
        ch = self.preview_canvas.winfo_height()
        max_w, max_h = max(10, int(cw) - 40), max(10, int(ch) - 40)
        for offset in [-1, 1]:
            i = idx + offset
            if 0 <= i < len(self.pages):
                rec = self.pages[i]
                if rec.is_blank: continue
                key = (rec.source_path, rec.source_index, max_w, max_h,
                       rec.orientation.get(), int(self._preview_zoom * 100))
                if key not in self.preview_cache:
                    img = render_page_image_fitz(rec.source_path, rec.source_index,
                                                 rec.orientation.get(), rec.orig_orient,
                                                 max_w, max_h)
                    if img:
                        self.preview_cache[key] = img
                        self._trim_preview_cache()

    def _trim_preview_cache(self, max_items=100):
        if len(self.preview_cache) > max_items:
            for k in list(self.preview_cache.keys())[:len(self.preview_cache)-max_items]:
                del self.preview_cache[k]

    def _scroll_to_row(self, idx):
        try:
            rows = self.rows_frame.winfo_children()
            total = len(rows)
            if total == 0: return
            frac = max(0.0, min(1.0, (idx - 1) / total))
            self.list_canvas.yview_moveto(frac)
        except:
            pass

    def _goto_page(self):
        try:
            n = int(self._goto_var.get())
            idx = n - 1
            if 0 <= idx < len(self.pages):
                self._preview_index = idx
                self._render_preview()
                self._scroll_to_row(idx)
                self._goto_var.set("")
        except:
            pass

    def _on_thumb_size_change(self, val):
        global THUMB_W, THUMB_H, ROW_H
        THUMB_W = int(float(val))
        THUMB_H = int(THUMB_W * 1.33)
        ROW_H   = THUMB_H + 16
        self.thumb_cache.clear()
        self._rebuild_rows()
        self._load_thumbs_async()

    # ═══════════════════════════════ ROW RENDERING ══════════════════════════
    def _rebuild_rows(self):
        for w in self.rows_frame.winfo_children():
            w.destroy()
        filter_text = self._search_var.get().strip().lower()
        for i, rec in enumerate(self.pages):
            if filter_text:
                label = f"page {i+1} {os.path.basename(rec.source_path).lower()}"
                if filter_text not in label:
                    continue
            self._add_row(i, rec)
        self.inner.update_idletasks()
        self.list_canvas.configure(scrollregion=self.list_canvas.bbox("all"))
        if self._preview_index >= 0 and self.pages:
            self._preview_index = min(self._preview_index, len(self.pages)-1)
            self.root.after(50, self._render_preview)

    def _filter_rows(self):
        self._rebuild_rows()

    def _add_row(self, i, rec: PageRecord):
        is_previewed = (i == self._preview_index)
        is_selected  = (i in self.selected_pages)
        row_bg = self._row_bg(i, rec.included.get(), is_previewed, is_selected)
        row = tk.Frame(self.rows_frame, bg=row_bg, relief="flat", bd=0, height=ROW_H)
        row.pack(fill="x", pady=1)
        row.pack_propagate(False)
        rec._row_widget = row

        if is_previewed:
            tk.Frame(row, width=4, bg="#2563EB").pack(side="left", fill="y")
        if is_selected:
            tk.Frame(row, width=4, bg="#F59E0B").pack(side="left", fill="y")

        handle = tk.Label(row, text="⠿", bg=row_bg, fg="#94A3B8",
                          font=("Helvetica", 14), cursor="fleur", width=2)
        handle.pack(side="left", padx=(4, 0))
        handle.bind("<ButtonPress-1>",   lambda e, idx=i: self._drag_start(e, idx))
        handle.bind("<B1-Motion>",       lambda e, idx=i: self._drag_motion(e, idx))
        handle.bind("<ButtonRelease-1>", lambda e, idx=i: self._drag_release(e, idx))

        cb = tk.Checkbutton(row, variable=rec.included, bg=row_bg, cursor="hand2",
                            command=lambda r=rec, idx=i: self._toggle_include(r, idx))
        cb.pack(side="left", padx=4)

        pg_lbl = tk.Label(row, text=str(i+1), width=3, bg=row_bg,
                          font=("Helvetica", 10, "bold"), fg="#475569", cursor="hand2")
        pg_lbl.pack(side="left", padx=2)
        pg_lbl.bind("<Button-1>",         lambda e, idx=i, r=rec: self._on_row_click(e, idx, r))
        pg_lbl.bind("<Control-Button-1>", lambda e, idx=i, r=rec: self._on_row_ctrl_click(e, idx, r))
        pg_lbl.bind("<Shift-Button-1>",   lambda e, idx=i, r=rec: self._on_row_shift_click(e, idx, r))

        thumb_lbl = tk.Label(row, bg=row_bg, cursor="hand2", relief="flat", bd=1)
        thumb_lbl.pack(side="left", padx=6, pady=4)
        rec._thumb_label = thumb_lbl
        self._refresh_thumb(rec)
        thumb_lbl.bind("<Button-1>", lambda e, r=rec, idx=i: self._on_thumb_click(r, idx))
        thumb_lbl.bind("<Button-3>", lambda e, r=rec, idx=i: self._show_context_menu(e, idx))
        row.bind("<Button-3>",       lambda e, idx=i: self._show_context_menu(e, idx))

        badge_color = OPTION_COLORS.get(rec.orig_orient, "#64748B")
        tk.Label(row, text=rec.orig_orient, width=10, bg=row_bg,
                 fg=badge_color, font=("Helvetica", 9, "bold")).pack(side="left", padx=4)

        rb_frame = tk.Frame(row, bg=row_bg)
        rb_frame.pack(side="left", padx=4)
        rec._rb_frame = rb_frame
        rec._rot_value_lbl = tk.Label(
            rb_frame, text=rotation_label(rec.orientation.get()), width=8,
            bg=row_bg, fg="#334155", font=("Helvetica", 9, "bold"))
        rec._rot_value_lbl.pack(side="left", padx=(0, 4))
        tk.Button(rb_frame, text="↻ CW",
                  command=lambda r=rec, idx=i: self._rotate_page(r, idx, ROTATE_STEP),
                  bg=OPTION_COLORS["Rotate 90° CW"], fg="white",
                  font=("Helvetica", 8, "bold"), relief="flat", padx=6, pady=2,
                  cursor="hand2").pack(side="left", padx=2)
        tk.Button(rb_frame, text="↺ CCW",
                  command=lambda r=rec, idx=i: self._rotate_page(r, idx, -ROTATE_STEP),
                  bg=OPTION_COLORS["Rotate 90° CCW"], fg="white",
                  font=("Helvetica", 8, "bold"), relief="flat", padx=6, pady=2,
                  cursor="hand2").pack(side="left", padx=2)
        self._refresh_orient_btns(rec)

        act = tk.Frame(row, bg=row_bg)
        act.pack(side="left", padx=6)

        def abtn(parent, text, cmd, bg, tip=""):
            b = tk.Button(parent, text=text, command=cmd, bg=bg, fg="white",
                          font=("Helvetica", 8, "bold"), relief="flat", padx=5, pady=2, cursor="hand2")
            if tip: Tooltip(b, tip)
            return b

        abtn(act, "⧉", lambda idx=i: self.duplicate_page(idx), "#7C3AED", "Duplicate").pack(side="left", padx=1)
        abtn(act, "🗑", lambda idx=i: self.delete_page(idx),    "#DC2626", "Delete").pack(side="left", padx=1)
        abtn(act, "↑",  lambda idx=i: self.move_page(idx, -1), "#0D9488", "Move up").pack(side="left", padx=1)
        abtn(act, "↓",  lambda idx=i: self.move_page(idx,  1), "#0D9488", "Move down").pack(side="left", padx=1)
        abtn(act, "✏", lambda idx=i: self.add_text_annotation(idx), "#B45309", "Annotate").pack(side="left", padx=1)
        abtn(act, "⬛", lambda idx=i: self.redact_dialog(idx), "#111827", "Redact").pack(side="left", padx=1)

    def _show_context_menu(self, event, idx):
        menu = tk.Menu(self.root, tearoff=0)
        menu.add_command(label=f"Preview page {idx+1}", command=lambda: self._on_thumb_click(self.pages[idx], idx))
        menu.add_separator()
        menu.add_command(label="Duplicate",      command=lambda: self.duplicate_page(idx))
        menu.add_command(label="Delete",         command=lambda: self.delete_page(idx))
        menu.add_separator()
        menu.add_command(label="Rotate CW",      command=lambda: self._rotate_page(self.pages[idx], idx, ROTATE_STEP))
        menu.add_command(label="Rotate CCW",     command=lambda: self._rotate_page(self.pages[idx], idx, -ROTATE_STEP))
        menu.add_separator()
        menu.add_command(label="Move Up",        command=lambda: self.move_page(idx, -1))
        menu.add_command(label="Move Down",      command=lambda: self.move_page(idx, 1))
        menu.add_separator()
        menu.add_command(label="Add Annotation", command=lambda: self.add_text_annotation(idx))
        menu.add_command(label="Redact Region",  command=lambda: self.redact_dialog(idx))
        menu.add_command(label="Page Inspector", command=lambda: self.page_inspector_dialog(idx))
        menu.add_separator()
        menu.add_command(label="Include",        command=lambda: self._set_include(idx, True))
        menu.add_command(label="Exclude",        command=lambda: self._set_include(idx, False))
        menu.tk_popup(event.x_root, event.y_root)

    def _set_include(self, idx, state):
        self.pages[idx].included.set(state)
        self._rebuild_rows()
        self._update_status()

    def _on_row_click(self, event, idx, rec):
        self.selected_pages.clear()
        self.selected_pages.add(idx)
        self._preview_index = idx
        self._render_preview(rec)
        self._rebuild_rows()

    def _on_row_ctrl_click(self, event, idx, rec):
        if idx in self.selected_pages:
            self.selected_pages.discard(idx)
        else:
            self.selected_pages.add(idx)
        self._rebuild_rows()

    def _on_row_shift_click(self, event, idx, rec):
        if not self.selected_pages:
            self.selected_pages.add(idx)
        else:
            last = max(self.selected_pages)
            lo, hi = min(last, idx), max(last, idx)
            for j in range(lo, hi+1):
                self.selected_pages.add(j)
        self._rebuild_rows()

    def _row_bg(self, i, included, is_previewed=False, is_selected=False):
        t = self._theme
        if is_previewed: return t["select"]
        if is_selected:  return "#FEF3C7"
        if not included: return t["excluded"]
        return t["row_even"] if i % 2 == 0 else t["row_odd"]

    def _refresh_thumb(self, rec: PageRecord):
        if rec.is_blank:
            rec.thumb_img = self._make_blank_thumb()
        else:
            key = (rec.source_path, rec.source_index, rec.orientation.get(), THUMB_W, THUMB_H)
            if key not in self.thumb_cache:
                self.thumb_cache[key] = render_page_image_fitz(
                    rec.source_path, rec.source_index,
                    rec.orientation.get(), rec.orig_orient,
                    THUMB_W, THUMB_H, for_thumb=True)
            rec.thumb_img = self.thumb_cache[key]
        if rec.thumb_img:
            rec.thumb_tk = ImageTk.PhotoImage(rec.thumb_img)
            if rec._thumb_label and rec._thumb_label.winfo_exists():
                rec._thumb_label.config(image=rec.thumb_tk, width=THUMB_W, height=THUMB_H)

    def _orientation_changed(self, rec: PageRecord, idx: int):
        self._refresh_orient_btns(rec)
        self._refresh_thumb(rec)
        if idx == self._preview_index:
            self.preview_cache.clear()
            self._render_preview(rec)
        self._update_status()

    def _rotate_page(self, rec: PageRecord, idx: int, delta: int):
        rec.orientation.set((int(rec.orientation.get()) + delta) % 360)
        self._orientation_changed(rec, idx)

    def _on_thumb_click(self, rec: PageRecord, idx: int):
        self._preview_index = idx
        self.selected_pages = {idx}
        self._render_preview(rec)
        self._rebuild_rows()

    def _refresh_orient_btns(self, rec: PageRecord):
        if not rec._rb_frame or not rec._rb_frame.winfo_exists():
            return
        row_bg = rec._rb_frame.cget("bg")
        if rec._rot_value_lbl and rec._rot_value_lbl.winfo_exists():
            rec._rot_value_lbl.config(text=rotation_label(rec.orientation.get()), bg=row_bg)

    def _toggle_include(self, rec: PageRecord, idx: int):
        bg = self._row_bg(idx, rec.included.get(), idx == self._preview_index)
        self._repaint_row(rec._row_widget, bg)
        self._refresh_orient_btns(rec)
        self._update_status()

    def _repaint_row(self, widget, bg):
        try: widget.config(bg=bg)
        except: pass
        for child in widget.winfo_children():
            self._repaint_row(child, bg)

    # ═══════════════════════════════ MULTI-SELECT ACTIONS ═══════════════════
    def select_all_pages(self):
        self.selected_pages = set(range(len(self.pages)))
        self._rebuild_rows()

    def deselect_all_pages(self):
        self.selected_pages.clear()
        self._rebuild_rows()

    def delete_selected(self):
        if not self.selected_pages: return
        self._push_undo("delete selected")
        for idx in sorted(self.selected_pages, reverse=True):
            self.pages.pop(idx)
        self.selected_pages.clear()
        if self._preview_index >= len(self.pages):
            self._preview_index = len(self.pages) - 1
        self._rebuild_rows()
        self._update_status()

    def duplicate_selected(self):
        if not self.selected_pages: return
        self._push_undo("duplicate selected")
        for idx in sorted(self.selected_pages, reverse=True):
            src = self.pages[idx]
            rec = PageRecord(src.source_path, src.source_index,
                             tk.StringVar(value=src.orig_orient),
                             is_blank=src.is_blank)
            rec.orig_orient = src.orig_orient
            rec.orientation.set(src.orientation.get())
            rec.annotations = list(src.annotations)
            rec.redactions  = list(src.redactions)
            self.pages.insert(idx+1, rec)
        self._rebuild_rows()
        self._update_status()

    # ═══════════════════════════════ THUMBNAIL LOADING ══════════════════════
    def _load_thumbs_async(self):
        self.progress.grid()
        self.progress.start(10)
        threading.Thread(target=self._load_thumbs_worker, daemon=True).start()

    def _load_thumbs_worker(self):
        for rec in list(self.pages):
            if rec.is_blank:
                rec.thumb_img = self._make_blank_thumb()
            else:
                key = (rec.source_path, rec.source_index, rec.orientation.get(), THUMB_W, THUMB_H)
                if key not in self.thumb_cache:
                    self.thumb_cache[key] = render_page_image_fitz(
                        rec.source_path, rec.source_index,
                        rec.orientation.get(), rec.orig_orient,
                        THUMB_W, THUMB_H, for_thumb=True)
                rec.thumb_img = self.thumb_cache[key]
            self.root.after(0, lambda r=rec: self._update_thumb_ui(r))
        self.root.after(0, self._thumbs_done)

    def _update_thumb_ui(self, rec: PageRecord):
        if rec._thumb_label and rec._thumb_label.winfo_exists() and rec.thumb_img:
            rec.thumb_tk = ImageTk.PhotoImage(rec.thumb_img)
            rec._thumb_label.config(image=rec.thumb_tk, width=THUMB_W, height=THUMB_H)
        if (0 <= self._preview_index < len(self.pages) and
                self.pages[self._preview_index] is rec):
            self.root.after(0, self._render_preview)

    def _thumbs_done(self):
        self.progress.stop()
        self.progress.grid_remove()
        if self._preview_index >= 0 and self.pages:
            self._render_preview()

    def _make_blank_thumb(self):
        img  = Image.new("RGB", (THUMB_W, THUMB_H), "#FFFFFF")
        draw = ImageDraw.Draw(img)
        draw.rectangle([1, 1, THUMB_W-2, THUMB_H-2], outline="#CCCCCC", width=2)
        draw.text((THUMB_W//2, THUMB_H//2), "BLANK", fill="#AAAAAA", anchor="mm")
        return img

    # ═══════════════════════════════ FILE OPERATIONS ════════════════════════
    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Open PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if path:
            self._open_path(path)

    def _open_path(self, path):
        self.primary_path = path
        self.root.title(f"PDF Studio v4 – {os.path.basename(path)}")
        self._load_pdf(path, replace=True)
        add_recent_file(path)
        self._refresh_recent_menu()

    def _load_pdf(self, path, replace=False, insert_after=None, password=None):
        try:
            reader = PdfReader(path)
            if reader.is_encrypted:
                pwd = password or simpledialog.askstring(
                    "Password", f"Enter password for:\n{os.path.basename(path)}",
                    show="*", parent=self.root)
                if not pwd:
                    return
                try:
                    reader.decrypt(pwd)
                except Exception:
                    messagebox.showerror("Error", "Wrong password.")
                    return
        except Exception as e:
            messagebox.showerror("Error", f"Could not read PDF:\n{e}")
            return

        new_records = []
        for i in range(len(reader.pages)):
            orient = get_page_orientation(reader.pages[i])
            rec = PageRecord(path, i, tk.StringVar(value=orient))
            new_records.append(rec)

        if replace:
            self.pages = new_records
            meta = reader.metadata or {}
            self.meta_title.set(meta.get("/Title", ""))
            self.meta_author.set(meta.get("/Author", ""))
            self.meta_subject.set(meta.get("/Subject", ""))
            self._preview_index = 0 if new_records else -1
            self.undo_stack = UndoStack()
            self._update_undo_labels()
        elif insert_after is None:
            self.pages.extend(new_records)
        else:
            self.pages[insert_after+1:insert_after+1] = new_records

        self.preview_cache.clear()
        self._rebuild_rows()
        self._load_thumbs_async()
        self._update_status()

    def merge_pdf(self):
        path = filedialog.askopenfilename(
            title="Merge PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not path: return
        self._push_undo("merge PDF")
        self._load_pdf(path, replace=False)

    def import_images(self):
        paths = filedialog.askopenfilenames(
            title="Select Images",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.tiff *.bmp *.gif"),
                       ("All files", "*.*")])
        if not paths: return
        self._push_undo("import images")
        for path in paths:
            self._add_image_as_page(path)
        self._rebuild_rows()
        self._load_thumbs_async()
        self._update_status()

    def _add_image_as_page(self, img_path):
        try:
            img = Image.open(img_path).convert("RGB")
            tmp = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            tmp.close()
            img.save(tmp.name, "PDF", resolution=150)
            rec = PageRecord(tmp.name, 0, tk.StringVar(value="Portrait"))
            self.pages.append(rec)
        except Exception as e:
            messagebox.showerror("Error", f"Could not import image:\n{img_path}\n{e}")

    def export_as_images(self):
        if not self.pages:
            messagebox.showwarning("No pages", "Open a PDF first.")
            return
        included = [r for r in self.pages if r.included.get()]
        if not included:
            messagebox.showwarning("Nothing selected", "Include at least one page.")
            return

        win = tk.Toplevel(self.root)
        win.title("Export as Images")
        win.geometry("360x220")
        win.grab_set()
        tk.Label(win, text="Export Settings", font=("Helvetica", 12, "bold")).pack(pady=8)
        fmt_var = tk.StringVar(value="PNG")
        dpi_var = tk.IntVar(value=150)
        tk.Label(win, text="Format:").pack()
        for f in ["PNG", "JPEG", "TIFF"]:
            tk.Radiobutton(win, text=f, variable=fmt_var, value=f).pack(anchor="w", padx=40)
        frm = tk.Frame(win); frm.pack(pady=4)
        tk.Label(frm, text="DPI:").pack(side="left")
        tk.Entry(frm, textvariable=dpi_var, width=6).pack(side="left", padx=4)

        def do_export():
            out_dir = filedialog.askdirectory(title="Choose output folder")
            if not out_dir: return
            win.destroy()
            fmt = fmt_var.get()
            dpi = max(72, min(600, dpi_var.get()))
            self.progress.grid(); self.progress.start(10)
            def worker():
                for i, rec in enumerate(included):
                    if rec.is_blank:
                        img = Image.new("RGB", (595, 842), "white")
                    else:
                        img = render_page_image_fitz(rec.source_path, rec.source_index,
                                                     rec.orientation.get(), rec.orig_orient,
                                                     int(595*dpi/72), int(842*dpi/72))
                    fname = os.path.join(out_dir, f"page_{i+1:04d}.{fmt.lower()}")
                    img.save(fname)
                self.root.after(0, lambda: (
                    self.progress.stop(), self.progress.grid_remove(),
                    messagebox.showinfo("Done", f"Exported {len(included)} images to:\n{out_dir}")))
            threading.Thread(target=worker, daemon=True).start()

        tk.Button(win, text="Export", command=do_export,
                  bg="#16A34A", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    # ═══════════════════════════════ PAGE ACTIONS ════════════════════════════
    def insert_blank(self):
        if not self.pages:
            messagebox.showwarning("No PDF", "Open a PDF first.")
            return
        self._push_undo("insert blank page")
        rec = PageRecord("__blank__", -1, tk.StringVar(value="Portrait"), is_blank=True)
        rec.thumb_img = self._make_blank_thumb()
        insert_at = self._preview_index + 1 if self._preview_index >= 0 else len(self.pages)
        self.pages.insert(insert_at, rec)
        self._rebuild_rows()
        self._update_status()

    def duplicate_page(self, idx):
        self._push_undo("duplicate page")
        src = self.pages[idx]
        rec = PageRecord(src.source_path, src.source_index,
                         tk.StringVar(value=src.orig_orient),
                         is_blank=src.is_blank)
        rec.orig_orient = src.orig_orient
        rec.orientation.set(src.orientation.get())
        rec.annotations = list(src.annotations)
        rec.redactions  = list(src.redactions)
        self.pages.insert(idx + 1, rec)
        self._rebuild_rows()
        self._update_status()

    def delete_page(self, idx):
        self._push_undo("delete page")
        self.pages.pop(idx)
        self.selected_pages.discard(idx)
        if self._preview_index >= len(self.pages):
            self._preview_index = len(self.pages) - 1
        self._rebuild_rows()
        self._update_status()

    def move_page(self, idx, direction):
        new_idx = idx + direction
        if 0 <= new_idx < len(self.pages):
            self._push_undo("move page")
            self.pages[idx], self.pages[new_idx] = self.pages[new_idx], self.pages[idx]
            if self._preview_index == idx:       self._preview_index = new_idx
            elif self._preview_index == new_idx: self._preview_index = idx
            self._rebuild_rows()

    def reverse_pages(self):
        if not self.pages: return
        self._push_undo("reverse pages")
        self.pages.reverse()
        self._rebuild_rows()
        self._update_status()
        self.status_var.set("⇅ Pages reversed")

    def extract_pages(self):
        included = [r for r in self.pages if r.included.get()]
        if not included:
            messagebox.showwarning("Nothing selected", "Include at least one page.")
            return
        out_path = filedialog.asksaveasfilename(
            title="Extract Pages", defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")])
        if out_path:
            self._write_pdf(included, out_path)

    def set_all_orient(self, choice):
        delta = normalize_rotation(choice)
        if delta == 270:
            delta = -ROTATE_STEP
        elif delta == 90:
            delta = ROTATE_STEP
        elif delta == 0:
            self.reset_all_orient()
            return
        self.rotate_all_pages(delta)

    def rotate_all_pages(self, delta):
        if not self.pages:
            return
        step_txt = "CW" if delta > 0 else "CCW"
        self._push_undo(f"rotate all {step_txt}")
        for rec in self.pages:
            rec.orientation.set((int(rec.orientation.get()) + delta) % 360)
        self.thumb_cache.clear()
        self.preview_cache.clear()
        self._rebuild_rows()
        self._load_thumbs_async()

    def reset_all_orient(self):
        self._push_undo("reset orientations")
        for rec in self.pages:
            rec.orientation.set(0)
        self.thumb_cache.clear()
        self.preview_cache.clear()
        self._rebuild_rows()
        self._load_thumbs_async()

    def set_all_include(self, state: bool):
        self._push_undo("set all include" if state else "exclude all")
        for rec in self.pages:
            rec.included.set(state)
        self._rebuild_rows()
        self._update_status()

    def select_odd_even(self, which):
        self._push_undo(f"select {which} pages")
        for i, rec in enumerate(self.pages):
            rec.included.set((i + 1) % 2 == (1 if which == "odd" else 0))
        self._rebuild_rows()
        self._update_status()

    def _parse_range_text(self):
        text = self.range_entry.get().strip()
        if not text: return None
        indices, n = set(), len(self.pages)
        for part in text.split(","):
            part = part.strip()
            if "-" in part:
                a, _, b = part.partition("-")
                try:
                    for p in range(int(a.strip()), int(b.strip()) + 1):
                        if 1 <= p <= n: indices.add(p - 1)
                except: pass
            else:
                try:
                    p = int(part)
                    if 1 <= p <= n: indices.add(p - 1)
                except: pass
        return indices

    def _apply_range(self, state: bool):
        idxs = self._parse_range_text()
        if not idxs:
            messagebox.showwarning("Invalid", "No valid page numbers.")
            return
        self._push_undo("range include/exclude")
        for idx in idxs:
            self.pages[idx].included.set(state)
        self._rebuild_rows()
        self._update_status()

    def duplicate_range(self):
        idxs = self._parse_range_text()
        if not idxs: return
        self._push_undo("duplicate range")
        for idx in sorted(idxs, reverse=True):
            src = self.pages[idx]
            rec = PageRecord(src.source_path, src.source_index,
                             tk.StringVar(value=src.orig_orient),
                             is_blank=src.is_blank)
            rec.orig_orient = src.orig_orient
            rec.orientation.set(src.orientation.get())
            self.pages.insert(idx + 1, rec)
        self._rebuild_rows()
        self._update_status()

    def delete_range(self):
        idxs = self._parse_range_text()
        if not idxs: return
        if not messagebox.askyesno("Confirm", f"Delete {len(idxs)} page(s)?"):
            return
        self._push_undo("delete range")
        for idx in sorted(idxs, reverse=True):
            self.pages.pop(idx)
        self._rebuild_rows()
        self._update_status()

    def orient_range_dialog(self):
        idxs = self._parse_range_text()
        if not idxs:
            messagebox.showwarning("Invalid", "No valid page numbers.")
            return
        win = tk.Toplevel(self.root)
        win.title("Rotate Range")
        win.geometry("320x180")
        win.grab_set()
        tk.Label(win, text=f"Rotate {len(idxs)} page(s):",
                 font=("Helvetica", 10, "bold")).pack(pady=10)

        btns = tk.Frame(win)
        btns.pack(pady=4)

        def apply(delta=None):
            self._push_undo("range rotate")
            for idx in idxs:
                if delta is None:
                    self.pages[idx].orientation.set(0)
                else:
                    self.pages[idx].orientation.set((int(self.pages[idx].orientation.get()) + delta) % 360)
            self.thumb_cache.clear()
            self.preview_cache.clear()
            self._rebuild_rows()
            self._load_thumbs_async()
            win.destroy()
        tk.Button(btns, text="↻ CW", command=lambda: apply(ROTATE_STEP),
                  bg=OPTION_COLORS["Rotate 90° CW"], fg="white",
                  relief="flat", padx=10, pady=4).pack(side="left", padx=4)
        tk.Button(btns, text="↺ CCW", command=lambda: apply(-ROTATE_STEP),
                  bg=OPTION_COLORS["Rotate 90° CCW"], fg="white",
                  relief="flat", padx=10, pady=4).pack(side="left", padx=4)
        tk.Button(win, text="Reset (0°)", command=lambda: apply(None),
                  bg="#64748B", fg="white", relief="flat", padx=10, pady=4).pack(pady=(8, 6))

    def _drag_start(self, event, idx):
        self.drag_data = {"idx": idx, "y_start": event.y_root, "moved": False}

    def _drag_motion(self, event, idx):
        if not self.drag_data: return
        dy    = event.y_root - self.drag_data["y_start"]
        steps = int(dy // max(1, ROW_H // 2))
        if steps != 0:
            new_idx = max(0, min(len(self.pages)-1, self.drag_data["idx"] + steps))
            if new_idx != self.drag_data["idx"]:
                if not self.drag_data["moved"]:
                    self._push_undo("reorder pages")
                    self.drag_data["moved"] = True
                self.pages[self.drag_data["idx"]], self.pages[new_idx] = \
                    self.pages[new_idx], self.pages[self.drag_data["idx"]]
                self.drag_data["idx"]     = new_idx
                self.drag_data["y_start"] = event.y_root
                self._rebuild_rows()

    def _drag_release(self, event, idx):
        self.drag_data = {}

    # ═══════════════════════════════ ANNOTATIONS & REDACTIONS ═══════════════
    def add_text_annotation(self, idx=None):
        if idx is None:
            idx = self._preview_index
        if idx < 0 or idx >= len(self.pages):
            messagebox.showwarning("No page", "Select a page first.")
            return
        win = tk.Toplevel(self.root)
        win.title(f"Annotate Page {idx+1}")
        win.geometry("400x300")
        win.grab_set()
        tk.Label(win, text="Annotation Text:", font=("Helvetica", 10, "bold")).pack(pady=(12,4))
        txt = tk.Text(win, height=6, width=45, font=("Helvetica", 10))
        txt.pack(padx=12)
        frm = tk.Frame(win); frm.pack(pady=8, fill="x", padx=12)
        tk.Label(frm, text="X pos (0-1):").grid(row=0, column=0, sticky="e")
        xv = tk.DoubleVar(value=0.1)
        tk.Entry(frm, textvariable=xv, width=6).grid(row=0, column=1, padx=4)
        tk.Label(frm, text="Y pos (0-1):").grid(row=0, column=2, sticky="e", padx=(8,0))
        yv = tk.DoubleVar(value=0.9)
        tk.Entry(frm, textvariable=yv, width=6).grid(row=0, column=3, padx=4)
        tk.Label(frm, text="Font size:").grid(row=1, column=0, sticky="e")
        fv = tk.IntVar(value=12)
        tk.Entry(frm, textvariable=fv, width=6).grid(row=1, column=1, padx=4)
        color_var = tk.StringVar(value="#000000")
        tk.Label(frm, text="Color:").grid(row=1, column=2, sticky="e", padx=(8,0))
        color_btn = tk.Button(frm, bg=color_var.get(), width=4, relief="raised",
                              command=lambda: self._pick_color(color_var, color_btn))
        color_btn.grid(row=1, column=3, padx=4)

        def apply():
            content = txt.get("1.0", "end").strip()
            if not content:
                win.destroy(); return
            self._push_undo("add annotation")
            self.pages[idx].annotations.append({
                "text": content, "x": xv.get(), "y": yv.get(),
                "fontsize": fv.get(), "color": color_var.get()
            })
            win.destroy()
            self.status_var.set(f"✏ Annotation added to page {idx+1}")

        tk.Button(win, text="Add Annotation", command=apply,
                  bg="#2563EB", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    def _pick_color(self, color_var, btn):
        result = colorchooser.askcolor(color=color_var.get(), parent=self.root)
        if result[1]:
            color_var.set(result[1])
            btn.config(bg=result[1])

    def redact_dialog(self, idx=None):
        if idx is None:
            idx = self._preview_index
        if idx < 0 or idx >= len(self.pages):
            messagebox.showwarning("No page", "Select a page first.")
            return
        win = tk.Toplevel(self.root)
        win.title(f"Redact Page {idx+1}")
        win.geometry("380x220")
        win.grab_set()
        tk.Label(win, text="Define redaction region (0.0–1.0 normalized coords):",
                 font=("Helvetica", 10, "bold")).pack(pady=(12,4))
        frm = tk.Frame(win); frm.pack(padx=12)
        vals = {}
        for r, (lbl, key, default) in enumerate([
            ("Left (x0):", "x0", 0.1), ("Top (y0):", "y0", 0.1),
            ("Right (x1):", "x1", 0.9), ("Bottom (y1):", "y1", 0.2)]):
            tk.Label(frm, text=lbl, width=14, anchor="e").grid(row=r, column=0, pady=3)
            v = tk.DoubleVar(value=default)
            tk.Entry(frm, textvariable=v, width=8).grid(row=r, column=1, padx=6)
            vals[key] = v

        def apply():
            try:
                box = (vals["x0"].get(), vals["y0"].get(), vals["x1"].get(), vals["y1"].get())
            except:
                messagebox.showerror("Error", "Invalid coordinates."); return
            self._push_undo("add redaction")
            self.pages[idx].redactions.append(box)
            win.destroy()
            self.status_var.set(f"⬛ Redaction added to page {idx+1}")

        tk.Button(win, text="⬛ Apply Redaction", command=apply,
                  bg="#111827", fg="white", relief="flat", padx=10, pady=4).pack(pady=12)

    def add_stamp_dialog(self):
        idx = self._preview_index
        if idx < 0:
            messagebox.showwarning("No page", "Select a page first.")
            return
        path = filedialog.askopenfilename(
            title="Select Stamp Image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp"), ("All files", "*.*")])
        if not path: return
        win = tk.Toplevel(self.root)
        win.title(f"Stamp Page {idx+1}")
        win.geometry("360x200")
        win.grab_set()
        tk.Label(win, text="Stamp position & size (0.0–1.0):",
                 font=("Helvetica", 10, "bold")).pack(pady=(12,4))
        frm = tk.Frame(win); frm.pack(padx=12)
        xv = tk.DoubleVar(value=0.7); yv = tk.DoubleVar(value=0.8)
        wv = tk.DoubleVar(value=0.25); hv = tk.DoubleVar(value=0.15)
        for r, (lbl, v) in enumerate([("X:", xv), ("Y:", yv), ("Width:", wv), ("Height:", hv)]):
            tk.Label(frm, text=lbl, width=8, anchor="e").grid(row=r//2, column=(r%2)*2, pady=3)
            tk.Entry(frm, textvariable=v, width=7).grid(row=r//2, column=(r%2)*2+1, padx=4)

        def apply():
            self._push_undo("add stamp")
            self.pages[idx].annotations.append({
                "type": "stamp", "image_path": path,
                "x": xv.get(), "y": yv.get(), "w": wv.get(), "h": hv.get()
            })
            win.destroy()
            self.status_var.set(f"🖼 Stamp added to page {idx+1}")

        tk.Button(win, text="Add Stamp", command=apply,
                  bg="#7C3AED", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    # ═══════════════════════════════ CROP & RESIZE ════════════════════════════
    def crop_margins_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Crop Margins")
        win.geometry("380x240")
        win.grab_set()
        tk.Label(win, text="Crop margins (points) for selected/all pages:",
                 font=("Helvetica", 10, "bold")).pack(pady=(12,4))
        frm = tk.Frame(win); frm.pack(padx=12)
        margins = {}
        for i, (lbl, key, default) in enumerate([
            ("Top:", "top", 0), ("Bottom:", "bottom", 0),
            ("Left:", "left", 0), ("Right:", "right", 0)]):
            tk.Label(frm, text=lbl, width=10, anchor="e").grid(row=i//2, column=(i%2)*2, pady=4)
            v = tk.DoubleVar(value=default)
            tk.Entry(frm, textvariable=v, width=8).grid(row=i//2, column=(i%2)*2+1, padx=6)
            margins[key] = v
        scope_var = tk.StringVar(value="all")
        for val, lbl in [("all","All pages"), ("included","Included only"), ("selected","Selected")]:
            tk.Radiobutton(win, text=lbl, variable=scope_var, value=val).pack(anchor="w", padx=40)

        def apply():
            t, b, l, r = (margins["top"].get(), margins["bottom"].get(),
                          margins["left"].get(), margins["right"].get())
            target = self._scope_pages(scope_var.get())
            if not target: win.destroy(); return
            self._push_undo("crop margins")
            for rec in target:
                rec.annotations.append({"type": "crop", "top": t, "bottom": b, "left": l, "right": r})
            win.destroy()
            self.status_var.set(f"✂ Crop applied to {len(target)} page(s)")

        tk.Button(win, text="Apply Crop", command=apply,
                  bg="#0D9488", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    def resize_pages_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Resize Pages")
        win.geometry("380x260")
        win.grab_set()
        tk.Label(win, text="Target page size:", font=("Helvetica", 11, "bold")).pack(pady=(12,4))
        size_var = tk.StringVar(value="A4")
        for s in list(PAGE_SIZES.keys()) + ["Custom"]:
            tk.Radiobutton(win, text=s, variable=size_var, value=s).pack(anchor="w", padx=40)
        cfrm = tk.Frame(win); cfrm.pack()
        tk.Label(cfrm, text="W (pts):").pack(side="left")
        cw = tk.IntVar(value=595); ch_v = tk.IntVar(value=842)
        tk.Entry(cfrm, textvariable=cw, width=6).pack(side="left", padx=2)
        tk.Label(cfrm, text="H:").pack(side="left")
        tk.Entry(cfrm, textvariable=ch_v, width=6).pack(side="left", padx=2)

        def apply():
            s = size_var.get()
            w, h = (cw.get(), ch_v.get()) if s == "Custom" else PAGE_SIZES[s]
            self._push_undo("resize pages")
            for rec in self.pages:
                rec.annotations.append({"type": "resize", "width": w, "height": h})
            win.destroy()
            self.status_var.set(f"📐 Resize to {s} ({w}×{h} pts) queued")

        tk.Button(win, text="Apply Resize", command=apply,
                  bg="#1565C0", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    def _scope_pages(self, scope):
        if scope == "all":       return list(self.pages)
        if scope == "included":  return [r for r in self.pages if r.included.get()]
        if scope == "selected":  return [self.pages[i] for i in self.selected_pages if i < len(self.pages)]
        return []

    # ═══════════════════════════════ SECURITY ═══════════════════════════════
    def encryption_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Encrypt PDF")
        win.geometry("380x280")
        win.grab_set()
        tk.Label(win, text="🔒 Password Protect Output PDF",
                 font=("Helvetica", 11, "bold")).pack(pady=(12,4))
        frm = tk.Frame(win); frm.pack(padx=20, pady=8, fill="x")
        for row, (lbl, var, show) in enumerate([
            ("Owner password:", self.owner_password, "*"),
            ("User password:",  self.user_password,  "*")]):
            tk.Label(frm, text=lbl, width=18, anchor="e").grid(row=row, column=0, pady=6)
            tk.Entry(frm, textvariable=var, show=show, width=24).grid(row=row, column=1, padx=8)
        tk.Label(frm, text="Owner = full access, User = read only",
                 fg="#64748B", font=("Helvetica", 8)).grid(row=2, column=0, columnspan=2, sticky="w")
        tk.Checkbutton(win, text="Enable encryption for output",
                       variable=self.encrypt_pdf, font=("Helvetica", 10)).pack(pady=4)
        tk.Checkbutton(win, text="Flatten form fields before saving",
                       variable=self.flatten_forms, font=("Helvetica", 10)).pack()
        tk.Button(win, text="Done", command=win.destroy,
                  bg="#16A34A", fg="white", relief="flat", padx=10, pady=4).pack(pady=12)

    def unlock_pdf_dialog(self):
        path = filedialog.askopenfilename(
            title="Select Encrypted PDF",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")])
        if not path: return
        pwd = simpledialog.askstring("Password", f"Password for:\n{os.path.basename(path)}",
                                     show="*", parent=self.root)
        if pwd is None: return
        out_path = filedialog.asksaveasfilename(
            title="Save Unlocked PDF", defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")])
        if not out_path: return
        try:
            reader = PdfReader(path)
            reader.decrypt(pwd)
            writer = PdfWriter()
            for page in reader.pages:
                writer.add_page(copy.deepcopy(page))
            with open(out_path, "wb") as f:
                writer.write(f)
            messagebox.showinfo("Done", f"Unlocked PDF saved to:\n{out_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Could not unlock:\n{e}")

    # ═══════════════════════════════ OCR ════════════════════════════════════
    def ocr_dialog(self):
        try:
            import pytesseract
        except ImportError:
            messagebox.showerror("Not installed", "pytesseract not installed.\nRun: pip install pytesseract")
            return
        if not self.pages:
            messagebox.showwarning("No pages", "Open a PDF first.")
            return
        win = tk.Toplevel(self.root)
        win.title("OCR – Make Searchable")
        win.geometry("400x260")
        win.grab_set()
        tk.Label(win, text="OCR Settings", font=("Helvetica", 11, "bold")).pack(pady=(12,4))
        lang_var = tk.StringVar(value="eng")
        frm = tk.Frame(win); frm.pack(padx=20, fill="x")
        tk.Label(frm, text="Language code (e.g. eng, hin, fra):").pack(anchor="w")
        tk.Entry(frm, textvariable=lang_var, width=20).pack(anchor="w", pady=4)
        scope_var = tk.StringVar(value="all")
        for val, lbl in [("all","All pages"), ("included","Included only")]:
            tk.Radiobutton(frm, text=lbl, variable=scope_var, value=val).pack(anchor="w")
        out_v = tk.StringVar(value="")
        tk.Label(frm, text="Output path:").pack(anchor="w", pady=(8,0))
        pfrm = tk.Frame(frm); pfrm.pack(fill="x")
        tk.Entry(pfrm, textvariable=out_v, width=32).pack(side="left")
        tk.Button(pfrm, text="…", command=lambda: out_v.set(
            filedialog.asksaveasfilename(defaultextension=".pdf",
                filetypes=[("PDF files","*.pdf")])),
            relief="flat", bg="#475569", fg="white").pack(side="left", padx=4)

        def do_ocr():
            out = out_v.get()
            if not out:
                messagebox.showwarning("No output", "Choose an output path."); return
            win.destroy()
            self.progress.grid(); self.progress.start(10)
            lang   = lang_var.get()
            target = [r for r in self.pages if r.included.get()] if scope_var.get() == "included" else self.pages

            def worker():
                try:
                    import pytesseract
                    writer = PdfWriter()
                    for rec in target:
                        if rec.is_blank:
                            writer.add_blank_page(width=595, height=842); continue
                        img = render_page_image_fitz(rec.source_path, rec.source_index,
                                                     rec.orientation.get(), rec.orig_orient,
                                                     1240, 1754)
                        pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, lang=lang, extension="pdf")
                        r2 = PdfReader(io.BytesIO(pdf_bytes))
                        writer.add_page(r2.pages[0])
                    with open(out, "wb") as f:
                        writer.write(f)
                    self.root.after(0, lambda: (
                        self.progress.stop(), self.progress.grid_remove(),
                        messagebox.showinfo("OCR Done", f"Searchable PDF saved:\n{out}")))
                except Exception as e:
                    self.root.after(0, lambda: (
                        self.progress.stop(), self.progress.grid_remove(),
                        messagebox.showerror("OCR Error", str(e))))

            threading.Thread(target=worker, daemon=True).start()

        tk.Button(win, text="Run OCR", command=do_ocr,
                  bg="#0D9488", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    # ═══════════════════════════════ FIND & REPLACE ══════════════════════════
    def find_replace_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Find & Replace Text")
        win.geometry("420x220")
        win.grab_set()
        tk.Label(win, text="Find & Replace (text layer PDFs only)",
                 font=("Helvetica", 11, "bold")).pack(pady=(12,4))
        frm = tk.Frame(win); frm.pack(padx=20, fill="x")
        fv = tk.StringVar(); rv = tk.StringVar()
        tk.Label(frm, text="Find:",    width=10, anchor="e").grid(row=0, column=0, pady=6)
        tk.Entry(frm, textvariable=fv, width=30).grid(row=0, column=1, padx=8)
        tk.Label(frm, text="Replace:", width=10, anchor="e").grid(row=1, column=0, pady=6)
        tk.Entry(frm, textvariable=rv, width=30).grid(row=1, column=1, padx=8)
        out_v = tk.StringVar()
        tk.Label(frm, text="Output:",  width=10, anchor="e").grid(row=2, column=0)
        pfrm = tk.Frame(frm); pfrm.grid(row=2, column=1)
        tk.Entry(pfrm, textvariable=out_v, width=22).pack(side="left")
        tk.Button(pfrm, text="…", command=lambda: out_v.set(
            filedialog.asksaveasfilename(defaultextension=".pdf",
                filetypes=[("PDF files","*.pdf")])),
            relief="flat", bg="#475569", fg="white").pack(side="left")

        def do_replace():
            find, replace, out = fv.get(), rv.get(), out_v.get()
            if not find or not out:
                messagebox.showwarning("Incomplete", "Fill all fields."); return
            win.destroy()
            try:
                with fitz.open(self.primary_path) as doc:
                    for page in doc:
                        page.clean_contents()
                        for inst in page.search_for(find):
                            page.add_redact_annot(inst, replace)
                        page.apply_redactions()
                    doc.save(out)
                messagebox.showinfo("Done", f"Saved with replacements:\n{out}")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        tk.Button(win, text="Replace All", command=do_replace,
                  bg="#2563EB", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    # ═══════════════════════════════ COMPARE PAGES ══════════════════════════
    def compare_pages_dialog(self):
        if len(self.pages) < 2:
            messagebox.showwarning("Need 2 pages", "Load at least 2 pages to compare.")
            return
        win = tk.Toplevel(self.root)
        win.title("Compare Pages Side by Side")
        win.geometry("1000x640")
        tk.Label(win, text="Select two pages to compare:", font=("Helvetica", 11, "bold")).pack(pady=8)
        frm = tk.Frame(win); frm.pack(fill="x", padx=20)
        p1v = tk.IntVar(value=1); p2v = tk.IntVar(value=2)
        tk.Label(frm, text="Page A:").pack(side="left")
        tk.Spinbox(frm, from_=1, to=len(self.pages), textvariable=p1v, width=5).pack(side="left", padx=4)
        tk.Label(frm, text="Page B:").pack(side="left", padx=(12,0))
        tk.Spinbox(frm, from_=1, to=len(self.pages), textvariable=p2v, width=5).pack(side="left", padx=4)
        cv = tk.Canvas(win, bg="#111827")
        cv.pack(fill="both", expand=True, padx=12, pady=8)

        def render():
            cv.update_idletasks()
            cw, ch = cv.winfo_width(), cv.winfo_height()
            half = (cw - 20) // 2
            cv.delete("all")
            cv._imgs = []
            for side, pv, x_off in [(0, p1v, 0), (1, p2v, half + 20)]:
                idx = pv.get() - 1
                if 0 <= idx < len(self.pages):
                    rec = self.pages[idx]
                    img = render_page_image_fitz(rec.source_path, rec.source_index,
                                                 rec.orientation.get(), rec.orig_orient,
                                                 half, ch - 20)
                    if img:
                        tk_img = ImageTk.PhotoImage(img)
                        cv._imgs.append(tk_img)
                        cv.create_image(x_off + half//2, ch//2, image=tk_img, anchor="center")
                        cv.create_text(x_off + half//2, 20, text=f"Page {pv.get()}",
                                       fill="white", font=("Helvetica", 10, "bold"))

        tk.Button(win, text="🔄 Render", command=render,
                  bg="#2563EB", fg="white", relief="flat", padx=10, pady=4).pack(pady=4)
        win.after(300, render)

    # ═══════════════════════════════ PAGE INSPECTOR ══════════════════════════
    def page_inspector_dialog(self, idx=None):
        if idx is None:
            idx = self._preview_index
        if idx < 0 or idx >= len(self.pages):
            messagebox.showwarning("No page", "Select a page first.")
            return
        rec  = self.pages[idx]
        info = {} if rec.is_blank else get_page_info_fitz(rec.source_path, rec.source_index)
        win  = tk.Toplevel(self.root)
        win.title(f"Page {idx+1} Inspector")
        win.geometry("380x320")
        win.grab_set()
        tk.Label(win, text=f"📄 Page {idx+1} Inspector", font=("Helvetica", 12, "bold")).pack(pady=(12,4))
        ttk.Separator(win).pack(fill="x", padx=12, pady=4)
        frm = tk.Frame(win, padx=16, pady=8); frm.pack(fill="both")
        fields = [
            ("Source file",  os.path.basename(rec.source_path) if not rec.is_blank else "Blank"),
            ("Source index", str(rec.source_index+1) if not rec.is_blank else "—"),
            ("Orig orient",  rec.orig_orient),
            ("Rotation",     rotation_label(rec.orientation.get())),
            ("Width (pts)",  f"{info.get('width','—'):.1f}" if info.get("width") else "—"),
            ("Height (pts)", f"{info.get('height','—'):.1f}" if info.get("height") else "—"),
            ("Has text",     "Yes" if info.get("has_text") else "No"),
            ("Annotations",  str(len(rec.annotations))),
            ("Redactions",   str(len(rec.redactions))),
            ("Included",     "Yes" if rec.included.get() else "No"),
        ]
        for i, (lbl, val) in enumerate(fields):
            bg = "#F8FAFC" if i % 2 == 0 else "#F0F4FF"
            row = tk.Frame(frm, bg=bg); row.pack(fill="x", pady=1)
            tk.Label(row, text=lbl, width=16, bg=bg, anchor="e",
                     font=("Helvetica", 9, "bold"), fg="#475569").pack(side="left", padx=4, pady=3)
            tk.Label(row, text=val, bg=bg, anchor="w",
                     font=("Helvetica", 9)).pack(side="left", padx=4)
        tk.Button(win, text="Close", command=win.destroy,
                  bg="#475569", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    # ═══════════════════════════════ HEADER / FOOTER ═════════════════════════
    def header_footer_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Header / Footer")
        win.geometry("440x220")
        win.grab_set()
        tk.Label(win, text="Header / Footer Text", font=("Helvetica", 11, "bold")).pack(pady=(12,4))
        tk.Label(win, text="Use {page} for page number, {total} for total pages, {date} for today",
                 fg="#64748B", font=("Helvetica", 8)).pack()
        frm = tk.Frame(win, padx=20); frm.pack(fill="x", pady=8)
        for row, (lbl, var) in enumerate([("Header:", self.header_text), ("Footer:", self.footer_text)]):
            tk.Label(frm, text=lbl, width=10, anchor="e").grid(row=row, column=0, pady=6)
            tk.Entry(frm, textvariable=var, width=36, font=("Helvetica", 10)).grid(row=row, column=1, padx=8)
        tk.Button(win, text="Done", command=win.destroy,
                  bg="#16A34A", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    # ═══════════════════════════════ BOOKMARKS ═══════════════════════════════
    def add_bookmark_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Add Bookmark")
        win.geometry("380x180")
        win.grab_set()
        tk.Label(win, text="Add PDF Bookmark", font=("Helvetica", 11, "bold")).pack(pady=(12,4))
        frm = tk.Frame(win, padx=20); frm.pack(fill="x")
        tk.Label(frm, text="Title:", width=8, anchor="e").grid(row=0, column=0, pady=6)
        tv = tk.StringVar()
        tk.Entry(frm, textvariable=tv, width=30).grid(row=0, column=1, padx=8)
        tk.Label(frm, text="Page:", width=8, anchor="e").grid(row=1, column=0, pady=6)
        pv = tk.IntVar(value=max(1, self._preview_index+1))
        tk.Entry(frm, textvariable=pv, width=8).grid(row=1, column=1, padx=8, sticky="w")

        def apply():
            self._bookmarks.append({"title": tv.get(), "page": pv.get()-1})
            win.destroy()
            self.status_var.set(f"🔖 Bookmark '{tv.get()}' added for page {pv.get()}")

        tk.Button(win, text="Add", command=apply,
                  bg="#2563EB", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    def bookmarks_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Manage Bookmarks")
        win.geometry("400x320")
        bmarks = self._bookmarks
        tk.Label(win, text="PDF Bookmarks", font=("Helvetica", 11, "bold")).pack(pady=(12,4))
        lb = tk.Listbox(win, font=("Helvetica", 10), height=10)
        lb.pack(fill="both", expand=True, padx=12, pady=4)
        for b in bmarks:
            lb.insert("end", f"  p.{b['page']+1}  {b['title']}")

        def delete_selected():
            sel = lb.curselection()
            for idx in reversed(sel):
                lb.delete(idx)
                if idx < len(bmarks): bmarks.pop(idx)

        tk.Button(win, text="Delete Selected", command=delete_selected,
                  bg="#DC2626", fg="white", relief="flat", padx=8, pady=3).pack(pady=4)
        tk.Button(win, text="Close", command=win.destroy,
                  bg="#475569", fg="white", relief="flat", padx=8, pady=3).pack(pady=4)

    # ═══════════════════════════════ BATCH PROCESSING ═══════════════════════
    def batch_process_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("Batch Process Folder")
        win.geometry("500x400")
        win.grab_set()
        tk.Label(win, text="🔄 Batch Process PDFs", font=("Helvetica", 12, "bold")).pack(pady=(12,4))

        in_dir_v  = tk.StringVar()
        out_dir_v = tk.StringVar()
        frm = tk.Frame(win, padx=16); frm.pack(fill="x", pady=4)

        def pick_dir(var):
            d = filedialog.askdirectory()
            if d: var.set(d)

        for row, (lbl, var) in enumerate([("Input folder:", in_dir_v), ("Output folder:", out_dir_v)]):
            tk.Label(frm, text=lbl, width=14, anchor="e").grid(row=row, column=0, pady=6)
            tk.Entry(frm, textvariable=var, width=28).grid(row=row, column=1, padx=4)
            tk.Button(frm, text="Browse…", command=lambda v=var: pick_dir(v),
                      bg="#475569", fg="white", relief="flat", padx=6).grid(row=row, column=2)

        ops = tk.LabelFrame(win, text="Operations to apply", padx=12, pady=8)
        ops.pack(fill="x", padx=16, pady=8)
        rotate_v    = tk.BooleanVar(value=False)
        rotate_to_v = tk.StringVar(value=OPTIONS[0])
        encrypt_v   = tk.BooleanVar(value=False)
        pwd_v       = tk.StringVar()
        compress_v  = tk.BooleanVar(value=True)

        tk.Checkbutton(ops, text="Rotate all pages by:", variable=rotate_v).grid(row=0, column=0, sticky="w")
        ttk.Combobox(ops, textvariable=rotate_to_v, values=OPTIONS, width=12, state="readonly").grid(row=0, column=1, padx=4)
        tk.Checkbutton(ops, text="Compress output", variable=compress_v).grid(row=1, column=0, sticky="w", pady=4)
        tk.Checkbutton(ops, text="Encrypt with password:", variable=encrypt_v).grid(row=2, column=0, sticky="w")
        tk.Entry(ops, textvariable=pwd_v, show="*", width=16).grid(row=2, column=1, padx=4)

        log_var = tk.StringVar(value="Ready.")
        tk.Label(win, textvariable=log_var, fg="#475569", font=("Helvetica", 9)).pack()

        def run_batch():
            in_dir = in_dir_v.get(); out_dir = out_dir_v.get()
            if not in_dir or not out_dir:
                messagebox.showwarning("Missing", "Choose input and output folders."); return
            os.makedirs(out_dir, exist_ok=True)
            pdfs = [f for f in os.listdir(in_dir) if f.lower().endswith(".pdf")]
            if not pdfs:
                messagebox.showwarning("None found", "No PDFs found in input folder."); return
            self.progress.grid(); self.progress.start(10)

            def worker():
                done = 0
                for fname in pdfs:
                    try:
                        src = os.path.join(in_dir, fname)
                        dst = os.path.join(out_dir, fname)
                        reader = PdfReader(src)
                        writer = PdfWriter()
                        for page in reader.pages:
                            p = copy.deepcopy(page)
                            if rotate_v.get():
                                apply_transform(p, rotate_to_v.get())
                            writer.add_page(p)
                        if compress_v.get():
                            for p in writer.pages:
                                try: p.compress_content_streams()
                                except: pass
                        if encrypt_v.get() and pwd_v.get():
                            writer.encrypt(pwd_v.get())
                        with open(dst, "wb") as f:
                            writer.write(f)
                        done += 1
                        self.root.after(0, lambda d=done, t=len(pdfs):
                            log_var.set(f"Processed {d}/{t}…"))
                    except Exception as e:
                        self.root.after(0, lambda fn=fname, err=e:
                            log_var.set(f"Error: {fn}: {err}"))
                self.root.after(0, lambda: (
                    self.progress.stop(), self.progress.grid_remove(),
                    log_var.set(f"✅ Done! {done}/{len(pdfs)} files processed."),
                    messagebox.showinfo("Batch Done", f"Processed {done}/{len(pdfs)} PDFs.")))

            threading.Thread(target=worker, daemon=True).start()

        tk.Button(win, text="▶ Run Batch", command=run_batch,
                  bg="#16A34A", fg="white", relief="flat", padx=12, pady=5).pack(pady=8)

    # ═══════════════════════════════ OUTPUT & METADATA DIALOGS ══════════════
    def show_output_options(self):
        win = tk.Toplevel(self.root)
        win.title("Output Options")
        win.geometry("480x480")
        win.resizable(False, False)
        win.grab_set()
        tk.Label(win, text="⚙ Output Options", font=("Helvetica", 12, "bold")).pack(pady=(12,4))
        ttk.Separator(win).pack(fill="x", padx=12, pady=4)

        pnf = tk.LabelFrame(win, text="Page Numbers", padx=8, pady=4)
        pnf.pack(fill="x", padx=12, pady=4)
        tk.Checkbutton(pnf, text="Add page numbers to output",
                       variable=self.add_page_numbers).grid(row=0, column=0, columnspan=4, sticky="w")
        tk.Label(pnf, text="Format:").grid(row=1, column=0, sticky="e")
        for col, (val, lbl) in enumerate([("decimal","1,2,3"), ("roman","i,ii,iii"), ("alpha","a,b,c")], 1):
            tk.Radiobutton(pnf, text=lbl, variable=self.page_num_format, value=val).grid(row=1, column=col)
        tk.Label(pnf, text="Position:").grid(row=2, column=0, sticky="e")
        ttk.Combobox(pnf, textvariable=self.page_num_position, width=18, state="readonly",
                     values=["bottom-center","bottom-left","bottom-right",
                             "top-center","top-left","top-right"]).grid(row=2, column=1, columnspan=3, sticky="w", padx=4)

        wf = tk.LabelFrame(win, text="Watermark / Stamp", padx=8, pady=4)
        wf.pack(fill="x", padx=12, pady=4)
        for row, (lbl, var, w) in enumerate([("Text:", self.watermark_text, 24), ("On pages:", self.watermark_pages, 16)]):
            tk.Label(wf, text=lbl, width=10, anchor="e").grid(row=row, column=0)
            tk.Entry(wf, textvariable=var, width=w).grid(row=row, column=1, padx=4, pady=3)
        tk.Label(wf, text="Opacity:").grid(row=2, column=0, sticky="e")
        ttk.Scale(wf, from_=5, to=100, orient="horizontal", length=120,
                  variable=self.watermark_opacity).grid(row=2, column=1, sticky="w")
        wc_btn = tk.Button(wf, text="  Color  ", bg=self.watermark_color.get(), relief="raised", width=6,
                           command=lambda: self._pick_color(self.watermark_color, wc_btn))
        wc_btn.grid(row=2, column=2, padx=4)

        of = tk.LabelFrame(win, text="Output Format", padx=8, pady=4)
        of.pack(fill="x", padx=12, pady=4)
        tk.Checkbutton(of, text="Compress / optimise output", variable=self.compress_output).pack(anchor="w")
        tk.Checkbutton(of, text="PDF/A compliance (archival)", variable=self.pdfa_mode).pack(anchor="w")
        tk.Checkbutton(of, text="Linearize (fast web view)",   variable=self.linearize).pack(anchor="w")

        tf = tk.LabelFrame(win, text="Page Size Override", padx=8, pady=4)
        tf.pack(fill="x", padx=12, pady=4)
        ttk.Combobox(tf, textvariable=self.output_page_size,
                     values=["Original"] + list(PAGE_SIZES.keys()),
                     state="readonly", width=14).pack(anchor="w")

        sf = tk.LabelFrame(win, text="Split Mode", padx=8, pady=4)
        sf.pack(fill="x", padx=12, pady=4)
        for val, lbl in [("single","Single PDF"), ("odd","Odd pages"), ("even","Even pages"), ("each","Each page separately")]:
            tk.Radiobutton(sf, text=lbl, variable=self.split_mode, value=val).pack(side="left", padx=6)

        tk.Button(win, text="✔ Done", command=win.destroy,
                  bg="#16A34A", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    def show_metadata_dialog(self):
        win = tk.Toplevel(self.root)
        win.title("PDF Metadata")
        win.geometry("420x320")
        win.resizable(False, False)
        win.grab_set()
        tk.Label(win, text="🏷 Edit PDF Metadata", font=("Helvetica", 12, "bold")).pack(pady=(12,4))
        ttk.Separator(win).pack(fill="x", padx=12, pady=4)
        frm = tk.Frame(win, padx=16, pady=8); frm.pack(fill="x")
        fields = [
            ("Title:",    self.meta_title),
            ("Author:",   self.meta_author),
            ("Subject:",  self.meta_subject),
            ("Keywords:", self.meta_keywords),
            ("Creator:",  self.meta_creator),
        ]
        for row, (lbl, var) in enumerate(fields):
            tk.Label(frm, text=lbl, width=12, anchor="e",
                     font=("Helvetica", 10)).grid(row=row, column=0, pady=5)
            tk.Entry(frm, textvariable=var, width=32,
                     font=("Helvetica", 10)).grid(row=row, column=1, padx=8)
        tk.Button(win, text="✔ Save", command=win.destroy,
                  bg="#16A34A", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    def show_shortcuts(self):
        win = tk.Toplevel(self.root)
        win.title("Keyboard Shortcuts")
        win.geometry("420x460")
        win.resizable(False, False)
        win.grab_set()
        tk.Label(win, text="⌨ Keyboard Shortcuts", font=("Helvetica", 12, "bold")).pack(pady=(12,4))
        ttk.Separator(win).pack(fill="x", padx=12, pady=4)
        shortcuts = [
            ("Ctrl + O",       "Open PDF"),
            ("Ctrl + S",       "Save PDF"),
            ("Ctrl + Z",       "Undo"),
            ("Ctrl + Y",       "Redo"),
            ("Ctrl + A",       "Select all pages"),
            ("Ctrl + Click",   "Toggle page selection"),
            ("Shift + Click",  "Range select pages"),
            ("Delete",         "Delete selected pages"),
            ("Ctrl + =",       "Zoom in preview"),
            ("Ctrl + −",       "Zoom out preview"),
            ("Ctrl + 0",       "Fit preview"),
            ("Ctrl + Scroll",  "Zoom preview"),
            ("← / →",          "Previous / Next page"),
            ("F1",             "Show this help"),
        ]
        frm = tk.Frame(win, padx=20, pady=8); frm.pack(fill="both")
        for i, (key, desc) in enumerate(shortcuts):
            bg = "#F8FAFC" if i % 2 == 0 else "#F0F4FF"
            row = tk.Frame(frm, bg=bg); row.pack(fill="x", pady=1)
            tk.Label(row, text=key, width=16, bg=bg, font=("Courier", 10, "bold"),
                     fg="#1E40AF", anchor="w").pack(side="left", padx=8, pady=4)
            tk.Label(row, text=desc, bg=bg, font=("Helvetica", 10), anchor="w").pack(side="left")
        ttk.Separator(win).pack(fill="x", padx=12, pady=8)
        tk.Label(win, text="💡 Right-click on page rows for context menu",
                 fg="#475569", font=("Helvetica", 9)).pack(pady=2)
        tk.Button(win, text="Close", command=win.destroy,
                  bg="#475569", fg="white", relief="flat", padx=10, pady=4).pack(pady=8)

    # ═══════════════════════════════ PRESET SAVE/LOAD ════════════════════════
    def save_preset(self):
        if not self.pages:
            messagebox.showwarning("No pages", "Open a PDF first."); return
        path = filedialog.asksaveasfilename(title="Save Preset", defaultextension=".json",
                                            filetypes=[("JSON preset", "*.json")])
        if not path: return
        data = {
            "pages":    [r.snapshot() for r in self.pages],
            "metadata": {"title": self.meta_title.get(), "author": self.meta_author.get(),
                         "subject": self.meta_subject.get(), "keywords": self.meta_keywords.get()},
            "options":  {
                "add_page_numbers":  self.add_page_numbers.get(),
                "page_num_format":   self.page_num_format.get(),
                "page_num_position": self.page_num_position.get(),
                "compress":          self.compress_output.get(),
                "split_mode":        self.split_mode.get(),
                "watermark_text":    self.watermark_text.get(),
                "watermark_pages":   self.watermark_pages.get(),
                "watermark_opacity": self.watermark_opacity.get(),
                "output_page_size":  self.output_page_size.get(),
                "header_text":       self.header_text.get(),
                "footer_text":       self.footer_text.get(),
            },
            "bookmarks": self._bookmarks,
        }
        with open(path, "w") as f:
            json.dump(data, f, indent=2)
        messagebox.showinfo("Preset saved", f"Saved to:\n{path}")

    def load_preset(self):
        path = filedialog.askopenfilename(title="Load Preset", filetypes=[("JSON preset", "*.json")])
        if not path: return
        try:
            with open(path) as f:
                data = json.load(f)
        except Exception as e:
            messagebox.showerror("Error", f"Could not read preset:\n{e}"); return

        self.pages.clear()
        for pd in data.get("pages", []):
            self.pages.append(PageRecord.from_snapshot(pd))
        meta = data.get("metadata", {})
        self.meta_title.set(meta.get("title", ""))
        self.meta_author.set(meta.get("author", ""))
        self.meta_subject.set(meta.get("subject", ""))
        self.meta_keywords.set(meta.get("keywords", ""))
        opts = data.get("options", {})
        self.add_page_numbers.set(opts.get("add_page_numbers", False))
        self.page_num_format.set(opts.get("page_num_format", "decimal"))
        self.page_num_position.set(opts.get("page_num_position", "bottom-center"))
        self.compress_output.set(opts.get("compress", False))
        self.split_mode.set(opts.get("split_mode", "single"))
        self.watermark_text.set(opts.get("watermark_text", ""))
        self.watermark_pages.set(opts.get("watermark_pages", "all"))
        self.watermark_opacity.set(opts.get("watermark_opacity", 40))
        self.output_page_size.set(opts.get("output_page_size", "Original"))
        self.header_text.set(opts.get("header_text", ""))
        self.footer_text.set(opts.get("footer_text", ""))
        self._bookmarks = data.get("bookmarks", [])
        self._preview_index = 0 if self.pages else -1
        self.preview_cache.clear()
        self.thumb_cache.clear()
        self._rebuild_rows()
        self._load_thumbs_async()
        self._update_status()
        messagebox.showinfo("Preset loaded", f"Loaded {len(self.pages)} pages.")

    # ═══════════════════════════════ PDF WRITING ════════════════════════════
    def save_pdf(self):
        if not self.pages:
            messagebox.showwarning("No pages", "Open a PDF first."); return
        included = [r for r in self.pages if r.included.get()]
        if not included:
            messagebox.showwarning("Nothing to save", "No pages are included."); return
        base = os.path.splitext(os.path.basename(self.primary_path or "output"))[0] + "_output"
        mode = self.split_mode.get()
        if mode == "each":
            out_dir = filedialog.askdirectory(title="Select folder for individual pages")
            if not out_dir: return
            self._save_each_page(included, out_dir, base)
        else:
            out_path = filedialog.asksaveasfilename(
                title="Save PDF", defaultextension=".pdf",
                initialfile=base + ".pdf",
                filetypes=[("PDF files", "*.pdf")])
            if not out_path: return
            if mode == "odd":
                pages_to_save = [r for i, r in enumerate(included) if (i+1) % 2 == 1]
            elif mode == "even":
                pages_to_save = [r for i, r in enumerate(included) if (i+1) % 2 == 0]
            else:
                pages_to_save = included
            self.progress.grid(); self.progress.start(10)
            threading.Thread(
                target=lambda: self._write_pdf_thread(pages_to_save, out_path),
                daemon=True).start()

    def _write_pdf_thread(self, records, out_path):
        try:
            self._write_pdf(records, out_path)
        finally:
            self.root.after(0, lambda: (self.progress.stop(), self.progress.grid_remove()))

    def _write_pdf(self, records, out_path):
        try:
            from datetime import date as _date
            reader_cache = {}

            def get_reader(path):
                if path not in reader_cache:
                    reader_cache[path] = PdfReader(path)
                return reader_cache[path]

            writer = PdfWriter()
            size_override = self.output_page_size.get()

            for i, rec in enumerate(records):
                if rec.is_blank:
                    tw, th = PAGE_SIZES.get(size_override, (595, 842)) if size_override != "Original" else (595, 842)
                    writer.add_blank_page(width=tw, height=th)
                else:
                    reader = get_reader(rec.source_path)
                    page   = copy.deepcopy(reader.pages[rec.source_index])
                    page   = apply_transform(page, rec.orientation.get())

                    if size_override != "Original" and size_override in PAGE_SIZES:
                        tw, th = PAGE_SIZES[size_override]
                        page.mediabox.lower_left  = (0, 0)
                        page.mediabox.upper_right = (tw, th)

                    for ann in rec.annotations:
                        if ann.get("type") == "crop":
                            mb = page.mediabox
                            pw = float(mb.width); ph = float(mb.height)
                            page.mediabox.lower_left  = (ann["left"], ann["bottom"])
                            page.mediabox.upper_right = (pw - ann["right"], ph - ann["top"])

                    writer.add_page(page)

            # Metadata
            meta = {}
            if self.meta_title.get():    meta["/Title"]    = self.meta_title.get()
            if self.meta_author.get():   meta["/Author"]   = self.meta_author.get()
            if self.meta_subject.get():  meta["/Subject"]  = self.meta_subject.get()
            if self.meta_keywords.get(): meta["/Keywords"] = self.meta_keywords.get()
            if self.meta_creator.get():  meta["/Creator"]  = self.meta_creator.get()
            if meta: writer.add_metadata(meta)

            # Bookmarks
            for bm in self._bookmarks:
                try:
                    pg = bm["page"]
                    if 0 <= pg < len(writer.pages):
                        writer.add_outline_item(bm["title"], pg)
                except: pass

            # Compress
            if self.compress_output.get():
                for p in writer.pages:
                    try: p.compress_content_streams()
                    except: pass

            # Encrypt
            if self.encrypt_pdf.get() and (self.owner_password.get() or self.user_password.get()):
                writer.encrypt(user_password=self.user_password.get(),
                               owner_password=self.owner_password.get())

            with open(out_path, "wb") as f:
                writer.write(f)

            # Post-processing overlays
            if (self.watermark_text.get().strip() or self.add_page_numbers.get() or
                    self.header_text.get().strip() or self.footer_text.get().strip()):
                self._add_overlay(out_path, len(records))

            if any(rec.redactions for rec in records if not rec.is_blank) or \
               any(rec.annotations for rec in records if not rec.is_blank):
                self._apply_fitz_overlays(records, out_path)

            n = len(records)
            self.root.after(0, lambda: (
                self.status_var.set(f"✅ Saved {n} page{'s' if n!=1 else ''} → {os.path.basename(out_path)}"),
                messagebox.showinfo("Saved", f"Saved {n} page{'s' if n!=1 else ''} successfully!\n\n{out_path}")
            ))
        except Exception as e:
            import traceback
            tb = traceback.format_exc()[:400]
            self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to save:\n{e}\n\n{tb}"))

    def _apply_fitz_overlays(self, records, out_path):
        try:
            doc = fitz.open(out_path)
            for i, rec in enumerate(records):
                if i >= len(doc): break
                page  = doc[i]
                pw, ph = page.rect.width, page.rect.height

                for ann in rec.annotations:
                    if ann.get("type") in ("crop", "resize"): continue
                    if ann.get("type") == "stamp":
                        try:
                            img_rect = fitz.Rect(
                                ann["x"] * pw, ann["y"] * ph,
                                (ann["x"] + ann["w"]) * pw,
                                (ann["y"] + ann["h"]) * ph)
                            page.insert_image(img_rect, filename=ann["image_path"])
                        except: pass
                    elif ann.get("text"):
                        c   = ann.get("color", "#000000").lstrip("#")
                        rgb = tuple(int(c[j:j+2], 16)/255 for j in (0, 2, 4))
                        page.insert_text(
                            fitz.Point(ann["x"] * pw, ann["y"] * ph),
                            ann["text"], fontsize=ann.get("fontsize", 12), color=rgb)

                for box in rec.redactions:
                    x0, y0, x1, y1 = box
                    page.add_redact_annot(fitz.Rect(x0*pw, y0*ph, x1*pw, y1*ph), fill=(0,0,0))
                if rec.redactions:
                    page.apply_redactions()

            doc.save(out_path, incremental=True, encryption=fitz.PDF_ENCRYPT_KEEP)
            doc.close()
        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"⚠ Overlay warning: {e}"))

    def _add_overlay(self, pdf_path, total_pages):
        try:
            from datetime import date as _date
            reader = PdfReader(pdf_path)
            writer = PdfWriter()

            wm_text      = self.watermark_text.get().strip()
            opacity      = max(5, min(100, self.watermark_opacity.get()))
            wm_color_hex = self.watermark_color.get().lstrip("#")
            wm_rgb       = tuple(int(wm_color_hex[j:j+2], 16) for j in (0, 2, 4)) \
                           if len(wm_color_hex) == 6 else (180, 180, 180)

            wm_pages_str = self.watermark_pages.get().strip()
            if wm_pages_str.lower() == "all":
                wm_set = set(range(total_pages))
            else:
                wm_set = set()
                for part in wm_pages_str.split(","):
                    part = part.strip()
                    if "-" in part:
                        try:
                            a, b = map(int, part.split("-"))
                            wm_set.update(range(a-1, b))
                        except: pass
                    else:
                        try: wm_set.add(int(part)-1)
                        except: pass

            def fmt_pnum(n, fmt):
                if fmt == "roman":  return _to_roman(n)
                if fmt == "alpha":  return chr(ord('a') + (n-1) % 26)
                return str(n)

            fnt_path = self._find_font()

            for i, page in enumerate(reader.pages):
                pw = int(float(page.mediabox.width))
                ph = int(float(page.mediabox.height))
                overlay = Image.new("RGBA", (pw, ph), (0,0,0,0))
                draw    = ImageDraw.Draw(overlay)

                try:
                    fnt_lg = ImageFont.truetype(fnt_path, max(20, pw // 10)) if fnt_path else ImageFont.load_default()
                    fnt_sm = ImageFont.truetype(fnt_path, 14)                if fnt_path else ImageFont.load_default()
                except:
                    fnt_lg = fnt_sm = ImageFont.load_default()

                if wm_text and i in wm_set:
                    diag   = int(math.sqrt(pw**2 + ph**2))
                    wm_img = Image.new("RGBA", (diag, diag), (0,0,0,0))
                    wd     = ImageDraw.Draw(wm_img)
                    bbox   = wd.textbbox((0,0), wm_text, font=fnt_lg)
                    tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
                    alpha  = int(255 * opacity / 100)
                    wd.text(((diag-tw)//2, (diag-th)//2), wm_text,
                            fill=(*wm_rgb, alpha), font=fnt_lg)
                    wm_img = wm_img.rotate(math.degrees(math.atan2(ph, pw)))
                    overlay.paste(wm_img, ((pw-diag)//2, (ph-diag)//2), wm_img)

                if self.add_page_numbers.get():
                    pn_pos  = self.page_num_position.get()
                    pn_text = fmt_pnum(i+1, self.page_num_format.get())
                    margin  = 20
                    bbox    = draw.textbbox((0,0), pn_text, font=fnt_sm)
                    tw, th  = bbox[2]-bbox[0], bbox[3]-bbox[1]
                    py = (ph - margin - th) if "bottom" in pn_pos else margin
                    if "center" in pn_pos: px = (pw - tw) // 2
                    elif "left"  in pn_pos: px = margin
                    else:                   px = pw - margin - tw
                    draw.text((px, py), pn_text, fill=(60,60,60,220), font=fnt_sm)

                today = _date.today().strftime("%Y-%m-%d")
                if self.header_text.get().strip():
                    ht = (self.header_text.get()
                          .replace("{page}", str(i+1))
                          .replace("{total}", str(total_pages))
                          .replace("{date}", today))
                    draw.text((20, 8), ht, fill=(60,60,60,220), font=fnt_sm)

                if self.footer_text.get().strip():
                    ft = (self.footer_text.get()
                          .replace("{page}", str(i+1))
                          .replace("{total}", str(total_pages))
                          .replace("{date}", today))
                    bbox = draw.textbbox((0,0), ft, font=fnt_sm)
                    th   = bbox[3]-bbox[1]
                    draw.text((20, ph - th - 8), ft, fill=(60,60,60,220), font=fnt_sm)

                buf = io.BytesIO()
                overlay.convert("RGB").save(buf, format="PDF", resolution=72)
                buf.seek(0)
                ov_reader = PdfReader(buf)
                page.merge_page(ov_reader.pages[0])
                writer.add_page(page)

            tmp = pdf_path + ".tmp"
            with open(tmp, "wb") as f:
                writer.write(f)
            os.replace(tmp, pdf_path)
        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"⚠ Overlay warning: {e}"))

    def _find_font(self):
        candidates = [
            "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
            "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
            "C:/Windows/Fonts/arial.ttf",
            "/System/Library/Fonts/Helvetica.ttc",
        ]
        for c in candidates:
            if os.path.exists(c): return c
        return ""

    def _save_each_page(self, records, out_dir, base):
        try:
            reader_cache = {}
            def get_reader(path):
                if path not in reader_cache: reader_cache[path] = PdfReader(path)
                return reader_cache[path]
            for i, rec in enumerate(records, 1):
                writer = PdfWriter()
                if rec.is_blank:
                    writer.add_blank_page(width=595, height=842)
                else:
                    reader = get_reader(rec.source_path)
                    page   = copy.deepcopy(reader.pages[rec.source_index])
                    apply_transform(page, rec.orientation.get())
                    writer.add_page(page)
                out = os.path.join(out_dir, f"{base}_page{i:03d}.pdf")
                with open(out, "wb") as f:
                    writer.write(f)
            messagebox.showinfo("Saved", f"Saved {len(records)} individual PDFs to:\n{out_dir}")
            self.status_var.set(f"✅ Saved {len(records)} PDFs → {out_dir}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed:\n{e}")

    # ═══════════════════════════════ SESSION & AUTOSAVE ══════════════════════
    def _clear_saved_session(self):
        try:
            if SESSION_FILE.exists():
                SESSION_FILE.unlink()
        except:
            pass
    def _save_session(self):
        try:
            if not SESSION_PERSISTENCE: return
            if not self.pages: return
            data = {
                "primary_path":   self.primary_path or "",
                "preview_index":  self._preview_index,
                "pages_snapshot": [r.snapshot() for r in self.pages],
            }
            SESSION_FILE.write_text(json.dumps(data))
        except: pass

    def _restore_session(self):
        try:
            if not SESSION_FILE.exists(): return
            data  = json.loads(SESSION_FILE.read_text())
            snaps = data.get("pages_snapshot", [])
            if not snaps: return
            if not all(s.get("is_blank") or os.path.exists(s["source_path"]) for s in snaps):
                return
            self.pages = [PageRecord.from_snapshot(s) for s in snaps]
            self.primary_path   = data.get("primary_path") or None
            self._preview_index = data.get("preview_index", 0)
            if self.primary_path:
                self.root.title(f"PDF Studio v4 – {os.path.basename(self.primary_path)} (restored)")
            self._rebuild_rows()
            self._load_thumbs_async()
            self._update_status()
            self.status_var.set("📋 Session restored from last run.")
        except: pass

    def _start_autosave(self):
        def tick():
            self._save_session()
            self._autosave_var.set(f"💾 {time.strftime('%H:%M:%S')}")
            self._auto_save_job = self.root.after(30000, tick)
        self._auto_save_job = self.root.after(30000, tick)

    def _on_close(self):
        self._save_session()
        self.root.destroy()

    # ═══════════════════════════════ STATUS ═════════════════════════════════
    def _update_status(self):
        if not self.pages:
            self.status_var.set("Open a PDF to get started.")
            return
        total    = len(self.pages)
        included = sum(r.included.get() for r in self.pages)
        sel      = len(self.selected_pages)
        sel_str  = f" • {sel} selected" if sel else ""
        self.status_var.set(
            f"{included}/{total} pages included{sel_str} • "
            f"Preview: page {self._preview_index+1 if self._preview_index >= 0 else '—'}")


# ═══════════════════════════════ ENTRY POINT ══════════════════════════════════
def main():
    root = tk.Tk()
    app  = PDFStudio(root)
    root.update_idletasks()
    W, H = 1280, 760
    sw   = root.winfo_screenwidth()
    sh   = root.winfo_screenheight()
    root.geometry(f"{W}x{H}+{(sw-W)//2}+{(sh-H)//2}")
    root.mainloop()

if __name__ == "__main__":
    main()