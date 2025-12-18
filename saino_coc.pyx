# cython: language_level=3
# saino_coc.pyx
# Saino+ Comic Ultimate Converter
# Developed by: im_abi

import sys
import os
import re
import shutil
import tempfile
import zipfile
import subprocess
import json
import time
import gc
from pathlib import Path
from typing import List, Dict, Optional

# Cython imports for speed
from libc.stdlib cimport malloc, free

from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QListWidget, QLabel,
    QMessageBox, QHBoxLayout, QAbstractItemView, QSpinBox, QFormLayout, QGroupBox,
    QCheckBox, QProgressBar, QInputDialog, QDialog, QComboBox, QFrame, QSizePolicy
)
from PySide6.QtGui import QFont, QIcon, QDragEnterEvent, QDropEvent, QColor, QPalette, QLinearGradient, QBrush, QPainter
from PySide6.QtCore import Qt, QThread, Signal, QSize, QTimer

from PIL import Image, ImageFile
ImageFile.LOAD_TRUNCATED_IMAGES = True

# --- Optional Libraries ---
HAS_IMG2PDF = False
try:
    import img2pdf
    HAS_IMG2PDF = True
except ImportError: pass

HAS_PATOOL = False
try:
    import patoolib
    HAS_PATOOL = True
except ImportError: pass

HAS_FITZ = False
try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except ImportError: pass

# --- Constants & Config ---
cdef tuple IMAGE_EXTS = ('.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tiff')
cdef tuple CONTAINER_EXTS = ('.zip', '.cbz', '.rar', '.cbr', '.7z')
CONFIG_PATH = Path.home() / ".saino_comic_config.json"
SESSION_PATH = Path.home() / ".saino_comic_session.json"

DEFAULT_CONFIG = {
    "language": "fa",
    "sort_mode": "Manual",
    "dpi_enabled": False,
    "dpi_value": 300,
    "quality_default": 90,
    "grayscale": False
}

# --- Styling (Saino Dark Theme) ---
STYLESHEET = """
QWidget {
    background-color: #121212;
    color: #e0e0e0;
    font-family: 'Segoe UI', 'Tahoma', sans-serif;
    font-size: 14px;
}
QFrame#Sidebar {
    background-color: #1a1a1a;
    border-right: 1px solid #333;
}
QGroupBox {
    border: 1px solid #444;
    border-radius: 8px;
    margin-top: 22px;
    font-weight: bold;
    color: #00e5ff;
    background-color: #1e1e1e;
}
QGroupBox::title {
    subcontrol-origin: margin;
    left: 10px;
    padding: 0 5px;
}
QPushButton {
    background-color: #2d2d2d;
    border: 1px solid #3e3e3e;
    border-radius: 6px;
    padding: 8px 15px;
    color: #ddd;
}
QPushButton:hover {
    background-color: #3d3d3d;
    border-color: #00e5ff;
    color: #00e5ff;
}
QPushButton#AccentButton {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #006064, stop:1 #00acc1);
    color: white;
    font-weight: bold;
    border: none;
}
QPushButton#AccentButton:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #00838f, stop:1 #00bcd4);
}
QPushButton#DangerButton {
    background-color: #b71c1c;
    color: white; border: none;
}
QPushButton#DangerButton:hover { background-color: #d32f2f; }

QListWidget {
    background-color: #181818;
    border: 1px solid #333;
    border-radius: 6px;
    font-size: 13px;
}
QListWidget::item { padding: 6px; }
QListWidget::item:selected {
    background-color: #004d40;
    border: 1px solid #00e5ff;
    color: white;
}
QProgressBar {
    border: 1px solid #444;
    border-radius: 6px;
    text-align: center;
    background-color: #111;
    color: white;
}
QProgressBar::chunk {
    background-color: #00e5ff;
    border-radius: 5px;
}
QComboBox, QSpinBox, QLineEdit {
    background-color: #252525;
    border: 1px solid #3e3e3e;
    padding: 5px; border-radius: 4px; color: white;
}
QLabel#Branding {
    font-family: 'Consolas', monospace;
    color: #00bcd4;
    font-size: 11px;
    font-weight: bold;
    padding: 8px;
}
"""

STRINGS = {
    "en": {
        "app_title": "Saino COC - Ultimate Comic Converter",
        "add": "Add Files / Folders",
        "convert": "START CONVERSION",
        "delete": "Remove Selected",
        "clear": "Clear All",
        "sort": "Sort Order:",
        "settings": "Configuration",
        "no_src": "List is empty! Add some files.",
        "separate": "Convert Separately",
        "merge": "Merge All to One",
        "format": "Output Format",
        "name_req": "Enter Output Filename:",
        "dest": "Select Destination Folder",
        "done": "Process Completed!",
        "err": "Error: ",
        "proc": "Processing: ",
        "save": "Saving Final File...",
        "manual": "Manual", "natural": "Natural Name", "date": "Date Added"
    },
    "fa": {
        "app_title": "ساینو COC - مبدل حرفه‌ای کمیک",
        "add": "افزودن فایل / پوشه",
        "convert": "شروع تبدیل",
        "delete": "حذف انتخاب شده",
        "clear": "پاکسازی لیست",
        "sort": "ترتیب مرتب‌سازی:",
        "settings": "تنظیمات و پیکربندی",
        "no_src": "لیست خالی است! فایلی اضافه کنید.",
        "separate": "تبدیل جداگانه (تکی)",
        "merge": "ادغام همه در یک فایل",
        "format": "فرمت خروجی",
        "name_req": "نام فایل نهایی را وارد کنید:",
        "dest": "انتخاب پوشه مقصد",
        "done": "عملیات با موفقیت پایان یافت!",
        "err": "خطا: ",
        "proc": "در حال پردازش: ",
        "save": "در حال ذخیره‌سازی نهایی...",
        "manual": "دستی", "natural": "نام طبیعی (Natural)", "date": "تاریخ افزودن"
    }
}

# --- Global Config ---
cdef dict CONFIG = DEFAULT_CONFIG.copy()
if CONFIG_PATH.exists():
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            CONFIG.update(json.load(f))
    except: pass

def save_config():
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(CONFIG, f, ensure_ascii=False, indent=2)
    except: pass

cpdef str tr(str key):
    return STRINGS.get(CONFIG.get("language", "fa"), STRINGS["fa"]).get(key, key)

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]

def get_7z_path():
    if sys.platform == "win32":
        paths = ["C:\\Program Files\\7-Zip\\7z.exe", "C:\\Program Files (x86)\\7-Zip\\7z.exe"]
        for p in paths:
            if os.path.exists(p): return p
    return "7z"

# --- Optimized Worker ---
class ConversionWorker(QThread):
    progress = Signal(int)
    status = Signal(str)
    finished = Signal(list)

    def __init__(self, sources, out_dir, mode_merge, out_fmt, out_name, quality, dpi, grayscale):
        super().__init__()
        self.sources = sources
        self.out_dir = out_dir
        self.mode_merge = mode_merge
        self.out_fmt = out_fmt
        self.out_name = out_name
        self.quality = quality
        self.dpi = dpi
        self.grayscale = grayscale
        self.cancel_flag = False
        self._temp_dirs = []

    def stop(self):
        self.cancel_flag = True

    def run(self):
        created = []
        try:
            total = len(self.sources)
            merge_pool = [] # Stores paths of extracted images

            for idx, src in enumerate(self.sources):
                if self.cancel_flag: break
                
                label = src.get("label", "Unknown")
                self.status.emit(f"{tr('proc')} {label}")
                
                # Extract/Render images
                imgs = self._get_images(src)
                if not imgs: continue

                if self.mode_merge:
                    merge_pool.extend(imgs)
                    self.progress.emit(int((idx + 1) / total * 50)) # 50% for extraction
                else:
                    # Individual Mode
                    out_path = self._save_file(imgs, os.path.splitext(label)[0])
                    if out_path: created.append(out_path)
                    self.progress.emit(int((idx + 1) / total * 100))
                    
                    # Memory Cleanup
                    imgs.clear()
                    gc.collect()

            # Merge Phase
            if self.mode_merge and merge_pool and not self.cancel_flag:
                self.status.emit(tr("save"))
                out_path = self._save_file(merge_pool, self.out_name)
                if out_path: created.append(out_path)
                self.progress.emit(100)

        except Exception as e:
            self.status.emit(f"{tr('err')} {str(e)}")
        finally:
            self._cleanup()
            self.finished.emit(created)

    def _get_images(self, src):
        ptype = src['type']
        path = src['path']
        if src.get('content_override'): return src['content_override']
        
        imgs = []
        if ptype == 'folder':
            for r, _, f in os.walk(path):
                for file in sorted(f, key=natural_sort_key):
                    if file.lower().endswith(IMAGE_EXTS): imgs.append(os.path.join(r, file))
        
        elif ptype == 'archive':
            tmp = tempfile.mkdtemp(prefix="saino_tmp_")
            self._temp_dirs.append(tmp)
            
            # 1. Patool
            if HAS_PATOOL:
                try: patoolib.extract_archive(path, outdir=tmp, verbose=False)
                except: pass
            # 2. 7zip
            else:
                exe = get_7z_path()
                subprocess.run([exe, "x", "-y", f"-o{tmp}", path], 
                               stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                               creationflags=0x08000000 if sys.platform=="win32" else 0)
            
            for r, _, f in os.walk(tmp):
                for file in sorted(f, key=natural_sort_key):
                    if file.lower().endswith(IMAGE_EXTS): imgs.append(os.path.join(r, file))

        elif ptype == 'pdf':
            if HAS_FITZ:
                tmp = tempfile.mkdtemp(prefix="saino_pdf_")
                self._temp_dirs.append(tmp)
                try:
                    doc = fitz.open(path)
                    for i, page in enumerate(doc):
                        if self.cancel_flag: break
                        pix = page.get_pixmap(dpi=self.dpi)
                        fn = os.path.join(tmp, f"p_{i:05d}.jpg")
                        pix.save(fn)
                        imgs.append(fn)
                    doc.close()
                except: pass
        
        elif ptype == 'image':
            imgs.append(path)
            
        return imgs

    def _save_file(self, images, name):
        if not images: return None
        fpath = os.path.join(self.out_dir, f"{name}.{self.out_fmt.lower()}")
        
        if self.out_fmt == 'CBZ':
            try:
                with zipfile.ZipFile(fpath, 'w', zipfile.ZIP_STORED) as z:
                    for i, img in enumerate(images):
                        if self.cancel_flag: return None
                        ext = os.path.splitext(img)[1]
                        z.write(img, f"p_{i:05d}{ext}")
                return fpath
            except: return None
            
        elif self.out_fmt == 'PDF':
            # Fast Path: img2pdf (Direct JPEG embedding)
            if HAS_IMG2PDF and not self.grayscale:
                try:
                    with open(fpath, "wb") as f:
                        f.write(img2pdf.convert(images))
                    return fpath
                except: pass # Fallback
            
            # Robust Path: PIL
            try:
                # Optimized logic: Convert first, then append others using generator to save RAM
                first_img = None
                pil_list = []
                
                for img in images:
                    if self.cancel_flag: break
                    try:
                        im = Image.open(img)
                        im = im.convert("L") if self.grayscale else im.convert("RGB")
                        
                        if first_img is None:
                            first_img = im
                        else:
                            pil_list.append(im)
                    except: continue

                if first_img:
                    first_img.save(fpath, "PDF", save_all=True, append_images=pil_list, quality=self.quality)
                    first_img.close()
                    for p in pil_list: p.close()
                    return fpath
            except Exception as e:
                print(e)
                return None
        return None

    def _cleanup(self):
        for d in self._temp_dirs:
            try: shutil.rmtree(d, ignore_errors=True)
            except: pass
        gc.collect()

# --- Main Window ---
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(tr("app_title"))
        self.resize(1000, 680)
        self.setAcceptDrops(True)
        self.sources = []
        self.init_ui()
        self.load_session()
        self.refresh_ui()

    def init_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0,0,0,0)
        
        # Sidebar
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setFixedWidth(270)
        side_lay = QVBoxLayout(sidebar)
        
        grp = QGroupBox(tr("settings"))
        form = QFormLayout()
        
        self.cmb_lang = QComboBox()
        self.cmb_lang.addItems(["فارسی", "English"])
        self.cmb_lang.setCurrentIndex(0 if CONFIG["language"] == "fa" else 1)
        self.cmb_lang.currentIndexChanged.connect(self.change_lang)

        self.cmb_sort = QComboBox()
        self.cmb_sort.addItems([tr("manual"), tr("natural"), tr("date")])
        self.cmb_sort.currentIndexChanged.connect(self.apply_sort)

        self.spin_qual = QSpinBox()
        self.spin_qual.setRange(10, 100)
        self.spin_qual.setValue(CONFIG.get("quality_default", 90))
        self.spin_qual.setSuffix("%")

        self.chk_dpi = QCheckBox("Custom DPI")
        self.spin_dpi = QSpinBox()
        self.spin_dpi.setRange(72, 600)
        self.spin_dpi.setValue(CONFIG.get("dpi_value", 300))
        self.spin_dpi.setEnabled(False)
        self.chk_dpi.stateChanged.connect(lambda s: self.spin_dpi.setEnabled(s!=0))

        self.chk_gray = QCheckBox("Grayscale")
        self.chk_gray.setChecked(CONFIG.get("grayscale", False))

        form.addRow(QLabel("Lang/زبان"), self.cmb_lang)
        form.addRow(tr("sort"), self.cmb_sort)
        form.addRow("Quality:", self.spin_qual)
        form.addRow(self.chk_dpi, self.spin_dpi)
        form.addRow(self.chk_gray)
        grp.setLayout(form)

        self.btn_convert = QPushButton(tr("convert"))
        self.btn_convert.setObjectName("AccentButton")
        self.btn_convert.setMinimumHeight(50)
        self.btn_convert.clicked.connect(self.start_convert)

        side_lay.addWidget(grp)
        side_lay.addStretch()
        side_lay.addWidget(self.btn_convert)

        # Main Area
        main_area = QFrame()
        main_lay = QVBoxLayout(main_area)

        top_bar = QHBoxLayout()
        self.btn_add = QPushButton(tr("add"))
        self.btn_del = QPushButton(tr("delete"))
        self.btn_del.setObjectName("DangerButton")
        self.btn_clr = QPushButton(tr("clear"))

        top_bar.addWidget(self.btn_add)
        top_bar.addStretch()
        top_bar.addWidget(self.btn_del)
        top_bar.addWidget(self.btn_clr)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_widget.setDragEnabled(True)
        self.list_widget.setAcceptDrops(True)
        self.list_widget.setDragDropMode(QAbstractItemView.InternalMove)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.lbl_status = QLabel("")
        self.lbl_status.setAlignment(Qt.AlignCenter)
        self.lbl_status.setStyleSheet("color: #00e5ff;")

        # Branding Footer
        lbl_brand = QLabel("ساخته شده در 𝕊𝕒𝕚𝕟𝕠™ برنامه نویس 𝕚𝕞_𝕒𝕓𝕚🌙")
        lbl_brand.setObjectName("Branding")
        lbl_brand.setAlignment(Qt.AlignRight | Qt.AlignBottom)
        lbl_brand.setTextInteractionFlags(Qt.TextSelectableByMouse)

        main_lay.addLayout(top_bar)
        main_lay.addWidget(self.list_widget)
        main_lay.addWidget(self.lbl_status)
        main_lay.addWidget(self.progress)
        main_lay.addWidget(lbl_brand)

        layout.addWidget(sidebar)
        layout.addWidget(main_area)

        # Connections
        self.btn_add.clicked.connect(self.add_files)
        self.btn_del.clicked.connect(self.delete_sel)
        self.btn_clr.clicked.connect(self.clear_all)

    # --- Logic ---
    def load_session(self):
        if SESSION_PATH.exists():
            try:
                with open(SESSION_PATH, 'r') as f:
                    data = json.load(f)
                    self.sources = data.get('sources', [])
            except: pass

    def save_session(self):
        # Update config
        CONFIG['language'] = "fa" if self.cmb_lang.currentIndex() == 0 else "en"
        CONFIG['grayscale'] = self.chk_gray.isChecked()
        save_config()
        # Update session
        clean = [{k:v for k,v in s.items() if k != 'temp'} for s in self.sources]
        try:
            with open(SESSION_PATH, 'w') as f:
                json.dump({'sources': clean}, f)
        except: pass

    def refresh_ui(self):
        self.list_widget.clear()
        for i, s in enumerate(self.sources):
            icon = "📁" if s['type'] == 'folder' else "📄"
            if s['type'] == 'archive': icon = "📦"
            self.list_widget.addItem(f"{i+1}. {icon}  {s['label']}")
        
        # Update Texts
        self.setWindowTitle(tr("app_title"))
        self.btn_add.setText(tr("add"))
        self.btn_convert.setText(tr("convert"))
        self.btn_del.setText(tr("delete"))
        self.btn_clr.setText(tr("clear"))

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, tr("add"), "", "Files (*.zip *.cbz *.rar *.pdf *.jpg *.png)")
        for f in files: self._add_path(f)
        self.refresh_ui()
        self.save_session()

    def _add_path(self, path):
        if any(s['path'] == path for s in self.sources): return
        ptype = "other"
        if os.path.isdir(path): ptype = "folder"
        elif path.lower().endswith(CONTAINER_EXTS): ptype = "archive"
        elif path.lower().endswith('.pdf'): ptype = "pdf"
        elif path.lower().endswith(IMAGE_EXTS): ptype = "image"
        
        self.sources.append({
            "path": path, "type": ptype, "label": os.path.basename(path),
            "added_at": time.time(), "content_override": None
        })

    def delete_sel(self):
        rows = sorted([i.row() for i in self.list_widget.selectedIndexes()], reverse=True)
        for r in rows: 
            if r < len(self.sources): del self.sources[r]
        self.refresh_ui()
        self.save_session()

    def clear_all(self):
        self.sources.clear()
        self.refresh_ui()
        self.save_session()

    def apply_sort(self):
        mode = self.cmb_sort.currentIndex()
        if mode == 1: self.sources.sort(key=lambda x: natural_sort_key(x['label']))
        elif mode == 2: self.sources.sort(key=lambda x: x['added_at'])
        self.refresh_ui()

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls(): e.accept()
    
    def dropEvent(self, e):
        for u in e.mimeData().urls(): self._add_path(u.toLocalFile())
        self.refresh_ui()
        self.save_session()

    def change_lang(self):
        CONFIG['language'] = "fa" if self.cmb_lang.currentIndex() == 0 else "en"
        self.refresh_ui()

    def start_convert(self):
        if not self.sources:
            QMessageBox.warning(self, "!", tr("no_src"))
            return
            
        # 1. Mode
        dlg = QMessageBox(self)
        dlg.setText(tr("settings"))
        b_sep = dlg.addButton(tr("separate"), QMessageBox.ActionRole)
        b_mrg = dlg.addButton(tr("merge"), QMessageBox.ActionRole)
        dlg.addButton(tr("cancel"), QMessageBox.RejectRole)
        dlg.exec()
        if dlg.clickedButton() is None: return
        is_merge = (dlg.clickedButton() == b_mrg)

        # 2. Format
        dlg2 = QMessageBox(self)
        dlg2.setText(tr("format"))
        b_pdf = dlg2.addButton("PDF", QMessageBox.ActionRole)
        b_cbz = dlg2.addButton("CBZ", QMessageBox.ActionRole)
        dlg2.exec()
        fmt = "PDF" if dlg2.clickedButton() == b_pdf else "CBZ"

        # 3. Name (Merge only)
        final_name = "Merged_Output"
        if is_merge:
            txt, ok = QInputDialog.getText(self, tr("format"), tr("name_req"))
            if ok and txt.strip(): final_name = txt.strip()
            else: return

        # 4. Folder
        out_dir = QFileDialog.getExistingDirectory(self, tr("dest"))
        if not out_dir: return

        self.btn_convert.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setValue(0)
        
        self.worker = ConversionWorker(
            self.sources, out_dir, is_merge, fmt, final_name,
            self.spin_qual.value(),
            self.spin_dpi.value() if self.chk_dpi.isChecked() else 300,
            self.chk_gray.isChecked()
        )
        self.worker.progress.connect(self.progress.setValue)
        self.worker.status.connect(self.lbl_status.setText)
        self.worker.finished.connect(self.on_finish)
        self.worker.start()

    def on_finish(self, files):
        self.btn_convert.setEnabled(True)
        self.progress.setVisible(False)
        self.lbl_status.setText(tr("done"))
        msg = f"{tr('done')}\n" + "\n".join([os.path.basename(f) for f in files])
        QMessageBox.information(self, "Saino COC", msg)

    def closeEvent(self, e):
        self.save_session()
        e.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setStyleSheet(STYLESHEET)
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
