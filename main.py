#!/usr/bin/env python3
# main.py
# Saino COC v4.1 - Final (Nuitka-ready)
# Author: im_abi + assistant
# Description: Full-featured Saino COC application: supports folders, archives, PDF input,
# Contents editor, merge/separate, session, localization, memory-safe PDF output,
# preserves library default DPI unless "Use custom DPI" is checked, cancel-safe, temp cleanup.

# ------------------ Build notes (Nuitka) ------------------
# Recommended Nuitka command for a single-file executable (Windows example):
#   nuitka --onefile --enable-plugin=pyside6 --include-package=pillow --output-dir=dist main.py
# For icons and resources: add --windows-icon-from-ico=icon.ico
# If using patool or fitz, ensure those packages are available in the build environment.
# Test the script in Python before building with Nuitka.
# ---------------------------------------------------------

import sys, os, re, json, shutil, tempfile, zipfile, subprocess, compileall, gc
from pathlib import Path
from typing import List, Dict, Optional

from PySide6.QtWidgets import (
    QApplication, QWidget, QLabel, QListWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QProgressBar, QComboBox, QSpinBox, QCheckBox, QFormLayout, QGroupBox,
    QMessageBox, QInputDialog, QDialog, QSlider, QTabWidget, QSplitter, QSizePolicy,
    QAbstractItemView
)
from PySide6.QtGui import QPixmap, QFont
from PySide6.QtCore import Qt, QThread, Signal

from PIL import Image, ImageFile, ImageEnhance
ImageFile.LOAD_TRUNCATED_IMAGES = True

# Optional libs
try:
    import img2pdf
    HAS_IMG2PDF = True
except Exception:
    HAS_IMG2PDF = False

try:
    try:
        from pypdf import PdfMerger
        HAS_PYPDF = True
    except Exception:
        from PyPDF2 import PdfMerger  # type: ignore
        HAS_PYPDF = True
except Exception:
    HAS_PYPDF = False

try:
    import fitz
    HAS_FITZ = True
except Exception:
    HAS_FITZ = False

try:
    import patoolib
    HAS_PATOOL = True
except Exception:
    HAS_PATOOL = False

# ---------------- constants & config ----------------
APP_NAME = "Saino COC"
VERSION = "v4.1"
HOME = Path.home()
CONFIG_PATH = HOME / ".saino_coc_v4_config.json"
SESSION_PATH = HOME / ".saino_coc_v4_session.json"
IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tiff', '.tif'}
CONTAINER_EXTS = {'.zip', '.cbz', '.rar', '.cbr', '.7z', '.tar'}
PRIORITY_EXTS = set(list(CONTAINER_EXTS) + ['.pdf'])

DEFAULT_CONFIG = {
    "language": "fa",
    "quality": 90,
    "use_custom_dpi": False,
    "dpi": 300,
    "sort_mode": "Manual",
    "enhancement": {"brightness":1.0, "contrast":1.0, "sharpness":1.0, "resize_w":0},
    "grayscale": False
}

# -------------- helpers ----------------
def load_json(path: Path, default):
    try:
        if path.exists():
            with open(path, 'r', encoding='utf-8') as f:
                return json.load(f)
    except Exception:
        pass
    return default.copy()

def save_json(path: Path, obj):
    try:
        with open(path, 'w', encoding='utf-8') as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

CONFIG = load_json(CONFIG_PATH, DEFAULT_CONFIG)
SESSION = load_json(SESSION_PATH, {"sources": []})

STRINGS = {
    "en": {
        "add":"Add","convert":"Convert","delete":"Delete","clear":"Clear All","sort":"Sort:",
        "settings":"Settings","no_sources_msg":"No sources added.","separate":"Separate per source","merge":"Merge into one",
        "output_format":"Output format","pdf":"PDF","cbz":"CBZ","merged_filename":"Merged filename",
        "proposed_filename":"Proposed filename (edit if you want):","invalid_name":"Filename cannot be empty.",
        "choose_output_folder":"Choose output folder","created":"Created:","no_outputs":"No outputs produced (error or cancelled).",
        "extract_failed":"Cannot extract archive for inspection.","cannot_extract":"Cannot extract archive: install 7z or patool.",
        "pdf_render_req":"PDF processing requires PyMuPDF (fitz).","ok":"OK","cancel":"Cancel","open_folder":"Open containing folder",
        "open_file":"Open first output file","done":"Done","app_name":APP_NAME
    },
    "fa": {
        "add":"افزودن","convert":"تبدیل","delete":"حذف","clear":"پاک کردن همه","sort":"مرتب‌سازی:",
        "settings":"تنظیمات","no_sources_msg":"هیچ منبعی اضافه نشده است.","separate":"هر منبع جدا (تکی)","merge":"ادغام در یک فایل",
        "output_format":"فرمت خروجی","pdf":"PDF","cbz":"CBZ","merged_filename":"نام فایل ادغام",
        "proposed_filename":"نام پیشنهادی (در صورت نیاز ویرایش کنید):","invalid_name":"نام نمی‌تواند خالی باشد.",
        "choose_output_folder":"پوشه خروجی را انتخاب کنید","created":"ساخته شد:","no_outputs":"هیچ فایلی ساخته نشد (خطا یا لغو).",
        "extract_failed":"استخراج آرشیو برای بازبینی ممکن نیست.","cannot_extract":"استخراج آرشیو ممکن نیست: 7z یا patool نصب کنید.",
        "pdf_render_req":"پردازش PDF نیاز به PyMuPDF (fitz) دارد.","ok":"تأیید","cancel":"انصراف","open_folder":"باز کردن پوشه حاوی خروجی",
        "open_file":"باز کردن اولین فایل خروجی","done":"پایان","app_name":APP_NAME
    }
}

def tr(key: str) -> str:
    lang = CONFIG.get('language', 'fa')
    return STRINGS.get(lang, STRINGS['fa']).get(key, key)

def natural_sort_key(s: str):
    return [int(x) if x.isdigit() else x.lower() for x in re.split(r'(\d+)', s)]

def extract_last_number(s: str) -> Optional[int]:
    nums = re.findall(r'(\d+)', s)
    return int(nums[-1]) if nums else None

def remove_numbers_from_name(s: str) -> str:
    return re.sub(r'\d+', '', s).strip().strip('_- ')

def run_7z_extract(archive_path: str, dest_dir: str) -> bool:
    exe = shutil.which('7z') or shutil.which('7za')
    if not exe:
        return False
    try:
        cmd = [exe, 'x', '-y', f'-o{dest_dir}', archive_path]
        p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return p.returncode == 0
    except Exception:
        return False

# ---------------- source & session ----------------
_next_id = 1

def make_source(path: str) -> Dict:
    global _next_id
    p = Path(path)
    ext = p.suffix.lower()
    if p.is_dir(): typ = 'folder'
    elif ext == '.pdf': typ = 'pdf'
    elif ext in CONTAINER_EXTS: typ = 'archive'
    elif ext in IMAGE_EXTS: typ = 'image'
    else: typ = 'other'
    src = {'id': _next_id, 'path': path, 'type': typ, 'label': p.name, 'temp': None, 'content_override': None, 'added_at': None}
    _next_id += 1
    return src

def session_save_sources(sources: List[Dict]):
    serial = []
    for s in sources:
        copy = {k:v for k,v in s.items() if k != 'temp'}
        serial.append(copy)
    SESSION['sources'] = serial
    save_json(SESSION_PATH, SESSION)

def session_load_sources() -> List[Dict]:
    return [s.copy() for s in SESSION.get('sources', [])]

# ---------------- Contents Editor (same as classic) ----------------
class ContentsEditor(QDialog):
    def __init__(self, parent, src: Dict):
        super().__init__(parent)
        self.src = src
        self.setWindowTitle(self.src.get('label', 'Contents'))
        self.resize(700, 420)
        self.working_dir: Optional[str] = None
        self.files: List[str] = []

        self.list_widget = QListWidget(); self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.sort_combo = QComboBox(); self.sort_combo.addItems(['Manual','Name (natural)','Number']); self.sort_combo.currentIndexChanged.connect(self.on_sort)
        btn_up = QPushButton('⬆'); btn_down = QPushButton('⬇'); btn_delete = QPushButton(tr('delete'))
        btn_ok = QPushButton(tr('ok') if tr('ok') else 'OK'); btn_cancel = QPushButton(tr('cancel') if tr('cancel') else 'Cancel')
        btn_up.clicked.connect(self.move_up); btn_down.clicked.connect(self.move_down); btn_delete.clicked.connect(self.delete_selected)
        btn_ok.clicked.connect(self.accept); btn_cancel.clicked.connect(self.reject)

        top = QHBoxLayout(); top.addWidget(QLabel(tr('sort'))); top.addWidget(self.sort_combo); top.addStretch()
        btn_row = QHBoxLayout(); btn_row.addWidget(btn_up); btn_row.addWidget(btn_down); btn_row.addWidget(btn_delete); btn_row.addStretch(); btn_row.addWidget(btn_ok); btn_row.addWidget(btn_cancel)
        layout = QVBoxLayout(self); layout.addLayout(top); layout.addWidget(self.list_widget); layout.addLayout(btn_row)
        self.prepare()

    def prepare(self):
        p = Path(self.src['path'])
        if self.src['type'] == 'archive':
            tmp = tempfile.mkdtemp(prefix='ex_contents_'); success=False
            if HAS_PATOOL:
                try:
                    patoolib.extract_archive(self.src['path'], outdir=tmp, interactive=False); success=True
                except Exception:
                    success=False
            if not success:
                if run_7z_extract(self.src['path'], tmp): success=True
            if not success:
                try: shutil.unpack_archive(self.src['path'], tmp); success=True
                except Exception: success=False
            if not success:
                QMessageBox.warning(self, tr('extract_failed'), tr('extract_failed')); self.working_dir=None; return
            self.working_dir = tmp; self.src['temp'] = tmp
        else:
            self.working_dir = self.src['path']

        if self.src.get('content_override'):
            self.files = [f for f in self.src['content_override'] if os.path.exists(f)]
        else:
            collected=[]
            if os.path.isdir(self.working_dir):
                for name in sorted(os.listdir(self.working_dir), key=natural_sort_key):
                    full = os.path.join(self.working_dir, name)
                    if os.path.isdir(full):
                        has_pr=False; has_im=False
                        for r,d,fs in os.walk(full):
                            for ff in fs:
                                ext=os.path.splitext(ff)[1].lower()
                                if ext in PRIORITY_EXTS: has_pr=True; break
                                if ext in IMAGE_EXTS: has_im=True
                            if has_pr: break
                        if has_pr or has_im: collected.append(full)
                    else:
                        ext=os.path.splitext(name)[1].lower()
                        if ext in PRIORITY_EXTS or ext in IMAGE_EXTS: collected.append(full)
            else:
                if os.path.isfile(self.working_dir): collected=[self.working_dir]
            self.files = collected
        self.reload()

    def reload(self):
        self.list_widget.clear()
        for p in self.files:
            label = os.path.basename(p)
            if os.path.isdir(p): label=f'[Folder] {label}'
            self.list_widget.addItem(label)

    def move_up(self):
        sel = sorted({idx.row() for idx in self.list_widget.selectedIndexes()})
        if not sel: return
        i = sel[0]
        if i>0:
            block=[self.files[j] for j in sel]
            for j in reversed(sel): del self.files[j]
            insert_at = sel[0]-1
            for k,item in enumerate(block): self.files.insert(insert_at+k, item)
            self.reload(); self.list_widget.setCurrentRow(insert_at)

    def move_down(self):
        sel = sorted({idx.row() for idx in self.list_widget.selectedIndexes()})
        if not sel: return
        i = sel[-1]
        if i < len(self.files)-1:
            block=[self.files[j] for j in sel]
            for j in reversed(sel): del self.files[j]
            insert_at = sel[-1]+1 - len(block) + 1
            for k,item in enumerate(block): self.files.insert(insert_at+k, item)
            self.reload(); self.list_widget.setCurrentRow(insert_at+len(block)-1)

    def delete_selected(self):
        sel = sorted({idx.row() for idx in self.list_widget.selectedIndexes()}, reverse=True)
        for i in sel:
            if 0<=i<len(self.files): del self.files[i]
        self.reload()

    def on_sort(self, _):
        mode = self.sort_combo.currentText()
        if mode=='Manual': return
        if mode=='Name (natural)': self.files.sort(key=lambda p: natural_sort_key(os.path.basename(p)))
        elif mode=='Number':
            def keyfn(p):
                n = extract_last_number(os.path.basename(p))
                return (n if n is not None else 10**9, natural_sort_key(os.path.basename(p)))
            self.files.sort(key=keyfn)
        self.reload()

    def accept(self):
        self.src['content_override'] = self.files.copy(); super().accept()

# ---------------- Worker ----------------
class BatchWorker(QThread):
    progress = Signal(int)
    message = Signal(str)
    detailed = Signal(dict)
    finished = Signal(list)

    def __init__(self, sources: List[Dict], out_dir: str, merge: bool=False, out_format: str='PDF', quality: int=90, dpi: int=300):
        super().__init__(); self.sources=sources; self.out_dir=out_dir; self.merge=merge; self.out_format=out_format.upper(); self.quality=quality; self.dpi=dpi
        self._temp_dirs=[]; self._cancel=False

    def cancel(self): self._cancel=True

    def run(self):
        outputs=[]
        try:
            os.makedirs(self.out_dir, exist_ok=True)
            total=len(self.sources); processed=0
            for si, src in enumerate(self.sources, start=1):
                if self._cancel: break
                label = src.get('label', os.path.basename(src.get('path','')))
                self.message.emit(f"{si}/{total}: {label}")
                items = src.get('content_override') or [src['path']]
                all_images=[]
                for it in items:
                    if self._cancel: break
                    imgs = self._gather_images_for_item(it)
                    all_images.extend(imgs)
                page_total=len(all_images)
                self.detailed.emit({'source_index':si,'source_total':total,'page':0,'page_total':page_total})
                if page_total==0:
                    processed+=1; self.progress.emit(int(processed/total*100)); continue
                base_name = remove_numbers_from_name(label) or f'source{si}'
                out_name = base_name if not self.merge else 'merged'
                if self._cancel: break
                if self.out_format=='CBZ':
                    op = self._make_cbz(all_images, self.out_dir, out_name)
                    if op: outputs.append(op)
                else:
                    op = self._make_pdf(all_images, self.out_dir, out_name, si, total)
                    if op: outputs.append(op)
                if self._cancel: break
                processed += 1; self.progress.emit(int(processed/total*100))
            # merge
            if self.merge and self.out_format=='PDF' and outputs and not self._cancel and HAS_PYPDF:
                try:
                    merged_path=os.path.join(self.out_dir,'merged.pdf'); merger=PdfMerger()
                    for p in outputs:
                        if self._cancel: break
                        merger.append(p)
                    if not self._cancel:
                        with open(merged_path,'wb') as fh: merger.write(fh)
                        merger.close(); outputs=[merged_path]; self.message.emit('Merged into merged.pdf')
                except Exception as e: self.message.emit(f'Merge failed: {e}')
            self._cleanup(); self.finished.emit(outputs if not self._cancel else [])
        except Exception as e:
            self._cleanup(); self.message.emit(f'Worker error: {e}'); self.finished.emit([])

    def _gather_images_for_item(self, item: str) -> List[str]:
        imgs=[]
        if os.path.isdir(item):
            for name in sorted(os.listdir(item), key=natural_sort_key):
                if self._cancel: break
                full=os.path.join(item,name)
                if os.path.isdir(full): imgs.extend(self._gather_images_for_item(full))
                else:
                    ext=os.path.splitext(name)[1].lower()
                    if ext in IMAGE_EXTS: imgs.append(full)
                    elif ext in PRIORITY_EXTS: imgs.extend(self._gather_images_for_item(full))
        elif os.path.isfile(item):
            ext=os.path.splitext(item)[1].lower()
            if ext in IMAGE_EXTS: imgs.append(item)
            elif ext=='.pdf':
                if not HAS_FITZ: self.message.emit(tr('pdf_render_req')); return []
                tmp=tempfile.mkdtemp(prefix='pdfimg_'); self._temp_dirs.append(tmp)
                try:
                    doc=fitz.open(item); use_custom=CONFIG.get('use_custom_dpi',False); target_dpi=int(self.dpi)
                    for i,page in enumerate(doc, start=1):
                        if self._cancel: break
                        pix = page.get_pixmap(dpi=target_dpi) if use_custom else page.get_pixmap()
                        out=os.path.join(tmp, f"{i:04}.jpg"); pix.save(out); imgs.append(out)
                    doc.close()
                except Exception as e: self.message.emit(f'PDF render error: {e}')
            elif ext in CONTAINER_EXTS:
                tmp=tempfile.mkdtemp(prefix='ex_'); self._temp_dirs.append(tmp); success=False
                if HAS_PATOOL:
                    try: patoolib.extract_archive(item, outdir=tmp, interactive=False); success=True
                    except Exception: success=False
                if not success:
                    if run_7z_extract(item,tmp): success=True
                if not success:
                    try: shutil.unpack_archive(item,tmp); success=True
                    except Exception: success=False
                if not success: self.message.emit(tr('cannot_extract')); return []
                for root,_,files in os.walk(tmp):
                    for f in sorted(files, key=natural_sort_key):
                        if self._cancel: break
                        if os.path.splitext(f)[1].lower() in IMAGE_EXTS: imgs.append(os.path.join(root,f))
        return imgs

    def _make_cbz(self, images: List[str], out_dir: str, base_name: str) -> Optional[str]:
        try:
            os.makedirs(out_dir, exist_ok=True)
            out_path=os.path.join(out_dir, f"{base_name}.cbz")
            with zipfile.ZipFile(out_path,'w',compression=zipfile.ZIP_STORED) as z:
                for i,p in enumerate(images, start=1):
                    if self._cancel: 
                        try: z.close()
                        except: pass
                        try: os.remove(out_path)
                        except: pass
                        return None
                    arc=f"{i:04}{os.path.splitext(p)[1].lower()}"
                    try: z.write(p, arc)
                    except Exception:
                        try:
                            im=Image.open(p); buf=tempfile.NamedTemporaryFile(delete=False,suffix='.jpg')
                            im.convert('RGB').save(buf.name, format='JPEG', quality=self.quality); im.close(); z.write(buf.name, arc); os.unlink(buf.name)
                        except Exception:
                            pass
            return out_path
        except Exception as e:
            self.message.emit(f'CBZ error: {e}'); return None

    def _make_pdf(self, images: List[str], out_dir: str, base_name: str, source_index: int, total_sources: int) -> Optional[str]:
        try:
            if not images: return None
            os.makedirs(out_dir, exist_ok=True)
            out_path=os.path.join(out_dir, f"{base_name}.pdf")
            if self._cancel: return None
            tmpdir=tempfile.mkdtemp(prefix='pdfout_'); self._temp_dirs.append(tmpdir); temp_files=[]
            for idx, src in enumerate(images, start=1):
                if self._cancel: break
                proc = self._process_image_to_temp(src, tmpdir, idx)
                if proc: temp_files.append(proc)
                self.detailed.emit({'source_index':source_index,'source_total':total_sources,'page':idx,'page_total':len(images)})
            if self._cancel: return None
            if HAS_IMG2PDF:
                try:
                    with open(out_path,'wb') as fh: fh.write(img2pdf.convert([str(p) for p in temp_files]))
                    return out_path
                except Exception as e:
                    self.message.emit(f'img2pdf failed, fallback PIL: {e}')
            # PIL fallback
            first=None; others=[]
            for i,p in enumerate(temp_files):
                if self._cancel: break
                try:
                    im=Image.open(p).convert('RGB')
                    if i==0: first=im
                    else: others.append(im)
                except Exception:
                    continue
            if self._cancel:
                try: 
                    if first: first.close()
                except: pass
                for im in others:
                    try: im.close()
                    except: pass
                return None
            if first:
                first.save(out_path, save_all=True, append_images=others, quality=self.quality)
                try: first.close()
                except: pass
                for im in others:
                    try: im.close()
                    except: pass
                return out_path
            return None
        finally:
            pass

    def _process_image_to_temp(self, src_path: str, tmpdir: str, idx: int) -> Optional[str]:
        try:
            im = Image.open(src_path)
            if CONFIG.get('grayscale', False): im = im.convert('L')
            else:
                if im.mode not in ('RGB','L'): im = im.convert('RGB')
                else: im = im.copy()
            rw = CONFIG.get('enhancement', {}).get('resize_w', 0)
            if rw and im.width > rw:
                h = int(im.height * (rw / im.width)); im = im.resize((rw,h), Image.Resampling.LANCZOS)
            enh = CONFIG.get('enhancement', {}); br = enh.get('brightness',1.0); co = enh.get('contrast',1.0); sh = enh.get('sharpness',1.0)
            if br != 1.0: im = ImageEnhance.Brightness(im).enhance(br)
            if co != 1.0: im = ImageEnhance.Contrast(im).enhance(co)
            if sh != 1.0: im = ImageEnhance.Sharpness(im).enhance(sh)
            out = os.path.join(tmpdir, f"{idx:05d}.jpg")
            save_kwargs={'quality': self.quality}
            if CONFIG.get('use_custom_dpi', False): save_kwargs['dpi']=(int(self.dpi), int(self.dpi))
            im.save(out, format='JPEG', **save_kwargs)
            try: im.close()
            except: pass
            return out
        except Exception as e:
            self.message.emit(f'Image process error: {e}'); return None

    def _cleanup(self):
        for d in list(self._temp_dirs):
            try: shutil.rmtree(d, ignore_errors=True)
            except Exception: pass
        self._temp_dirs=[]; gc.collect()

# ---------------- MainWindow (old layout vibe) ----------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} {VERSION}")
        # smaller window, not fullscreen, compact layout similar to original
        self.resize(980, 700)
        self.sources: List[Dict] = session_load_sources()
        self.worker: Optional[BatchWorker] = None
        self._build_ui()
        self._refresh_list()

    def _build_ui(self):
        font = QFont('Segoe UI', 10); self.setFont(font)
        main = QVBoxLayout(self)

        # top toolbar (compact)
        top = QHBoxLayout()
        btn_add = QPushButton(tr('add')); btn_add.clicked.connect(self.action_add)
        btn_del = QPushButton(tr('delete')); btn_del.clicked.connect(self.action_delete)
        btn_clear = QPushButton(tr('clear')); btn_clear.clicked.connect(self.action_clear)
        top.addWidget(btn_add); top.addWidget(btn_del); top.addWidget(btn_clear); top.addStretch()
        main.addLayout(top)

        # central split: list (left) + actions (right)
        splitter = QSplitter(Qt.Horizontal)
        left = QWidget(); llay = QVBoxLayout(left)
        self.list_widget = QListWidget(); self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_widget.itemDoubleClicked.connect(self.action_edit_contents)
        llay.addWidget(self.list_widget)

        # action bar beneath list (compact)
        action_row = QHBoxLayout()
        self.combo_mode = QComboBox(); self.combo_mode.addItems([tr('separate'), tr('merge')])
        self.combo_fmt = QComboBox(); self.combo_fmt.addItems(['PDF','CBZ','FOLDER'])
        self.btn_start = QPushButton('START PROCESSING'); self.btn_start.setStyleSheet('font-weight:700; background:#006064; color:white')
        self.btn_start.setMinimumHeight(44); self.btn_start.clicked.connect(self.action_convert)
        action_row.addWidget(QLabel('Mode:')); action_row.addWidget(self.combo_mode); action_row.addWidget(QLabel('Format:')); action_row.addWidget(self.combo_fmt); action_row.addStretch(); action_row.addWidget(self.btn_start)
        llay.addLayout(action_row)

        splitter.addWidget(left)

        # right: preview + log + settings vertical
        right = QWidget(); rlay = QVBoxLayout(right)
        self.preview_label = QLabel('Preview'); self.preview_label.setAlignment(Qt.AlignCenter); self.preview_label.setMinimumSize(280,360)
        rlay.addWidget(self.preview_label)
        self.log_list = QListWidget(); self.log_list.setMaximumHeight(180)
        rlay.addWidget(QLabel('Log:')); rlay.addWidget(self.log_list)

        # settings compact
        grp = QGroupBox(tr('settings')); f = QFormLayout()
        self.quality_spin = QSpinBox(); self.quality_spin.setRange(10,100); self.quality_spin.setValue(CONFIG.get('quality',90))
        self.custom_dpi_cb = QCheckBox('Use custom DPI'); self.custom_dpi_cb.setChecked(CONFIG.get('use_custom_dpi',False))
        self.dpi_spin = QSpinBox(); self.dpi_spin.setRange(72,1200); self.dpi_spin.setValue(CONFIG.get('dpi',300)); self.dpi_spin.setEnabled(self.custom_dpi_cb.isChecked())
        self.custom_dpi_cb.toggled.connect(self.dpi_spin.setEnabled)
        self.gray_cb = QCheckBox('Grayscale'); self.gray_cb.setChecked(CONFIG.get('grayscale', False))
        self.resize_spin = QSpinBox(); self.resize_spin.setRange(0,8000); self.resize_spin.setValue(CONFIG.get('enhancement',{}).get('resize_w',0)); self.resize_spin.setSuffix(' px')
        self.b_slide = QSlider(Qt.Horizontal); self.b_slide.setRange(5,30); self.b_slide.setValue(int(CONFIG.get('enhancement',{}).get('brightness',1.0)*10))
        self.c_slide = QSlider(Qt.Horizontal); self.c_slide.setRange(5,30); self.c_slide.setValue(int(CONFIG.get('enhancement',{}).get('contrast',1.0)*10))
        self.s_slide = QSlider(Qt.Horizontal); self.s_slide.setRange(5,40); self.s_slide.setValue(int(CONFIG.get('enhancement',{}).get('sharpness',1.0)*10))
        f.addRow('Quality:', self.quality_spin); f.addRow(self.custom_dpi_cb, self.dpi_spin); f.addRow(self.gray_cb); f.addRow('Resize:', self.resize_spin)
        f.addRow('Brightness:', self.b_slide); f.addRow('Contrast:', self.c_slide); f.addRow('Sharpness:', self.s_slide)
        btn_save = QPushButton('Save settings'); btn_save.clicked.connect(self.save_settings); f.addRow(btn_save)
        grp.setLayout(f); rlay.addWidget(grp)

        splitter.addWidget(right)
        main.addWidget(splitter)

        # footer: progress + cancel + attribution (small and centered)
        fb = QHBoxLayout(); self.progress = QProgressBar(); self.progress.setVisible(False); self.cancel_btn = QPushButton('Cancel'); self.cancel_btn.setEnabled(False); self.cancel_btn.clicked.connect(self.action_cancel)
        fb.addWidget(self.progress); fb.addWidget(self.cancel_btn)
        main.addLayout(fb)

        footer_lbl = QLabel('ساخته شده توسط 𝕊𝕒𝕚𝕟𝕠™ دولوپر𝕚𝕞_𝕒𝕓𝕚🌙')
        footer_lbl.setAlignment(Qt.AlignRight); footer_lbl.setStyleSheet('color:#00bcd4; font-family: Consolas; font-size:11px')
        main.addWidget(footer_lbl)

        # store refs
        self.enh_b_slider = self.b_slide; self.enh_c_slider = self.c_slide; self.enh_s_slider = self.s_slide

    # UI actions
    def action_add(self):
        mb = QMessageBox(self); mb.setWindowTitle(tr('add')); mb.setText(tr('add'))
        files_btn = mb.addButton('Select files (images/pdf/archives)', QMessageBox.ActionRole); folder_btn = mb.addButton('Add folder', QMessageBox.ActionRole); mb.addButton(tr('cancel'), QMessageBox.RejectRole)
        mb.exec(); clicked = mb.clickedButton()
        if clicked == files_btn:
            files,_ = QFileDialog.getOpenFileNames(self,'Select files','', 'Supported (*.zip *.cbz *.cbr *.rar *.pdf *.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff)')
            for f in files:
                if not f: continue; s = make_source(f); self.sources.append(s)
            self.apply_sort(); self._refresh_list(); session_save_sources(self.sources)
        elif clicked == folder_btn:
            folder = QFileDialog.getExistingDirectory(self,'Select folder');
            if not folder: return; added = self._scan_and_add_folder(folder)
            if not added: QMessageBox.information(self,'', tr('no_sources_msg'))
            self.apply_sort(); self._refresh_list(); session_save_sources(self.sources)

    def _scan_and_add_folder(self, folder_path: str) -> bool:
        found=[]; folder_path=os.path.abspath(folder_path)
        for root,dirs,files in os.walk(folder_path):
            for d in dirs:
                sub=os.path.join(root,d)
                has_pr=False; has_im=False
                for r2,d2,f2 in os.walk(sub):
                    for ff in f2:
                        ext=os.path.splitext(ff)[1].lower()
                        if ext in PRIORITY_EXTS: has_pr=True; break
                        if ext in IMAGE_EXTS: has_im=True
                    if has_pr: break
                if has_pr: found.append(sub)
        top_pr=False; top_im=False
        for f in os.listdir(folder_path):
            ext=os.path.splitext(f)[1].lower();
            if ext in PRIORITY_EXTS: top_pr=True
            if ext in IMAGE_EXTS: top_im=True
        if found:
            for sub in sorted(set(found), key=natural_sort_key): self.sources.append(make_source(sub))
            return True
        if top_pr or top_im:
            self.sources.append(make_source(folder_path)); return True
        return False

    def action_delete(self):
        sel=sorted({idx.row() for idx in self.list_widget.selectedIndexes()}, reverse=True)
        for r in sel:
            if 0<=r<len(self.sources): s=self.sources.pop(r);
            if s.get('temp'):
                try: shutil.rmtree(s['temp'], ignore_errors=True)
                except: pass
        self._refresh_list(); session_save_sources(self.sources)

    def action_clear(self):
        for s in self.sources:
            if s.get('temp'):
                try: shutil.rmtree(s['temp'], ignore_errors=True)
                except: pass
        self.sources.clear(); self._refresh_list(); session_save_sources(self.sources)

    def action_edit_contents(self, item):
        row=self.list_widget.currentRow();
        if row<0 or row>=len(self.sources): return
        src=self.sources[row]; dlg=ContentsEditor(self, src)
        if dlg.exec(): self.apply_sort(); self._refresh_list(); session_save_sources(self.sources)

    def action_convert(self):
        if not self.sources: QMessageBox.warning(self,'', tr('no_sources_msg')); return
        mb=QMessageBox(self); mb.setWindowTitle(tr('settings')); mb.setText(tr('settings'))
        sep=mb.addButton(tr('separate'), QMessageBox.ActionRole); merg=mb.addButton(tr('merge'), QMessageBox.ActionRole); mb.addButton(tr('cancel'), QMessageBox.RejectRole); mb.exec()
        if mb.clickedButton() is None: return
        do_merge=(mb.clickedButton()==merg)
        mb2=QMessageBox(self); mb2.setWindowTitle(tr('output_format')); mb2.setText(tr('output_format'))
        pdf_b=mb2.addButton(tr('pdf'), QMessageBox.ActionRole); cbz_b=mb2.addButton(tr('cbz'), QMessageBox.ActionRole); mb2.addButton(tr('cancel'), QMessageBox.RejectRole); mb2.exec()
        if mb2.clickedButton() is None: return
        out_fmt = 'PDF' if mb2.clickedButton()==pdf_b else 'CBZ'
        final_name=None
        if do_merge:
            base=remove_numbers_from_name(os.path.basename(self.sources[0]['path'])) or 'output'
            text, ok = QInputDialog.getText(self, tr('merged_filename'), tr('proposed_filename'), text=base+'_merged')
            if not ok: return
            final_name=text.strip();
            if final_name=='': QMessageBox.warning(self,'Invalid', tr('invalid_name')); return
        out_dir = QFileDialog.getExistingDirectory(self, tr('choose_output_folder'))
        if not out_dir: return
        run_cfg = {
            'quality': self.quality_spin.value(), 'use_custom_dpi': self.custom_dpi_cb.isChecked(), 'dpi': self.dpi_spin.value(),
            'enhancement': {'brightness': self.enh_b_slider.value()/10.0, 'contrast': self.enh_c_slider.value()/10.0, 'sharpness': self.enh_s_slider.value()/10.0, 'resize_w': self.resize_spin.value()},
            'grayscale': self.gray_cb.isChecked()
        }
        CONFIG.update(run_cfg)
        self.progress.setVisible(True); self.progress.setValue(0); self.cancel_btn.setEnabled(True)
        thread_sources=[s.copy() for s in self.sources]
        worker=BatchWorker(thread_sources, out_dir, merge=do_merge, out_format=out_fmt, quality=run_cfg['quality'], dpi=run_cfg['dpi'])
        self.worker=worker
        worker.progress.connect(self.progress.setValue); worker.message.connect(self._log); worker.detailed.connect(self._on_detailed); worker.finished.connect(self._on_finished)
        worker.start()

    def action_cancel(self):
        if self.worker: self.worker.cancel(); self._log('Cancel requested...'); self.cancel_btn.setEnabled(False)

    def _on_detailed(self, info: dict):
        si = info.get('source_index'); st = info.get('source_total'); p = info.get('page'); pt = info.get('page_total')
        self._log(f"{si}/{st} — {p}/{pt}")

    def _on_finished(self, outputs: List[str]):
        self.progress.setVisible(False); self.cancel_btn.setEnabled(False)
        if not outputs: QMessageBox.information(self, tr('done'), tr('no_outputs')); return
        msg = tr('created') + '\n' + '\n'.join(outputs)
        dlg = QMessageBox(self); dlg.setWindowTitle(tr('done')); dlg.setText(msg)
        open_f = dlg.addButton(tr('open_folder'), QMessageBox.ActionRole); open_file = dlg.addButton(tr('open_file'), QMessageBox.ActionRole)
        dlg.addButton(tr('cancel'), QMessageBox.RejectRole); dlg.exec()
        if dlg.clickedButton() == open_f:
            try:
                if sys.platform.startswith('win'): os.startfile(os.path.dirname(outputs[0]))
                elif sys.platform == 'darwin': subprocess.call(['open', os.path.dirname(outputs[0])])
                else: subprocess.call(['xdg-open', os.path.dirname(outputs[0])])
            except Exception:
                pass
        elif dlg.clickedButton() == open_file:
            try:
                f = outputs[0]
                if sys.platform.startswith('win'): os.startfile(f)
                elif sys.platform == 'darwin': subprocess.call(['open', f])
                else: subprocess.call(['xdg-open', f])
            except Exception:
                pass
        session_save_sources(self.sources)

    def _log(self, txt: str):
        self.log_list.addItem(txt); self.log_list.scrollToBottom()

    def _refresh_list(self):
        self.list_widget.clear()
        for i, s in enumerate(self.sources, start=1):
            icon = '📁' if s['type']=='folder' else ('📦' if s['type']=='archive' else ('🖼️' if s['type']=='image' else '📄'))
            self.list_widget.addItem(f"{i}. {icon} {s.get('label', os.path.basename(s.get('path','')))}")

    def on_sort_changed(self, _): self.apply_sort(); self._refresh_list(); CONFIG['sort_mode']=self.sort_combo.currentText(); save_json(CONFIG_PATH, CONFIG)

    def apply_sort(self):
        mode=self.sort_combo.currentText()
        if mode=='Manual': return
        if mode=='Name (natural)': self.sources.sort(key=lambda s: natural_sort_key(s.get('label','')))
        elif mode=='Added time': self.sources.sort(key=lambda s: s.get('added_at') or 0)
        elif mode=='Number':
            def k(s): n=extract_last_number(s.get('label','')); return (n if n is not None else 10**9, natural_sort_key(s.get('label','')))
            self.sources.sort(key=k)

    def save_settings(self):
        CONFIG['quality']=self.quality_spin.value(); CONFIG['use_custom_dpi']=self.custom_dpi_cb.isChecked(); CONFIG['dpi']=self.dpi_spin.value()
        CONFIG['enhancement']={'brightness':self.enh_b_slider.value()/10.0,'contrast':self.enh_c_slider.value()/10.0,'sharpness':self.enh_s_slider.value()/10.0,'resize_w':self.resize_spin.value()}
        CONFIG['grayscale']=self.gray_cb.isChecked(); save_json(CONFIG_PATH, CONFIG); QMessageBox.information(self,'Settings','Saved')

    def closeEvent(self, ev):
        CONFIG['sort_mode']=self.sort_combo.currentText(); save_json(CONFIG_PATH, CONFIG); session_save_sources(self.sources)
        for s in self.sources:
            if s.get('temp'):
                try: shutil.rmtree(s['temp'], ignore_errors=True)
                except: pass
        try: compileall.compile_file(__file__, force=False, quiet=1)
        except Exception: pass
        ev.accept()

# ---------------- run ----------------
def main():
    app = QApplication(sys.argv); w = MainWindow(); w.show(); sys.exit(app.exec())

if __name__ == '__main__':
    main()
