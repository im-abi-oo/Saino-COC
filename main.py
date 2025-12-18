# batch_converter_saino_final.py
# Final fixes: removed dry-run, fixed language, fixed cancel behavior, compiled pyc on exit, app name Saino+ Comic

import sys, os, re, shutil, tempfile, zipfile, subprocess, json, time, compileall
from pathlib import Path
from typing import List, Dict, Optional

from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QListWidget, QLabel,
    QMessageBox, QHBoxLayout, QAbstractItemView, QSpinBox, QFormLayout, QGroupBox,
    QCheckBox, QProgressBar, QInputDialog, QDialog, QComboBox, QSizePolicy
)
from PySide6.QtGui import QKeyEvent
from PySide6.QtCore import Qt, QThread, Signal

from PIL import Image

# optional libs (used if available)
try:
    import img2pdf; HAS_IMG2PDF = True
except Exception:
    HAS_IMG2PDF = False

try:
    from pypdf import PdfMerger; HAS_PYPDF = True
except Exception:
    try:
        from PyPDF2 import PdfMerger; HAS_PYPDF = True
    except Exception:
        HAS_PYPDF = False

try:
    import patoolib; HAS_PATOOL = True
except Exception:
    HAS_PATOOL = False

try:
    import fitz; HAS_FITZ = True
except Exception:
    HAS_FITZ = False

# ---------------- constants and config ----------------
IMAGE_EXTS = tuple(Image.registered_extensions().keys())
CONTAINER_EXTS = ('.zip', '.cbz', '.rar', '.cbr')
PRIORITY_EXTS = tuple(list(CONTAINER_EXTS) + ['.pdf'])

CONFIG_PATH = Path.home() / ".batch_converter_config.json"
SESSION_PATH = Path.home() / ".batch_converter_session.json"

DEFAULT_CONFIG = {
    "language": "fa",
    "sort_mode": "Manual",
    "dpi_enabled": False,
    "dpi_value": 300,
    "quality_default": 95,
    "app_name": "Saino COC"
}

def load_json(path: Path, default):
    try:
        if path.exists():
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return default.copy()

def save_json(path: Path, obj):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(obj, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

CONFIG = load_json(CONFIG_PATH, DEFAULT_CONFIG)
SESSION = load_json(SESSION_PATH, {"sources": []})

# ---------- localization (strings used throughout UI) ----------
STRINGS = {
    "en": {
        "add": "Add",
        "convert": "Convert",
        "delete": "Delete",
        "move_up": "Move Up",
        "move_down": "Move Down",
        "clear": "Clear All",
        "sort": "Sort:",
        "settings": "Settings",
        "no_sources_msg": "No sources added.",
        "separate": "Separate per source",
        "merge": "Merge into one",
        "output_format": "Choose output format",
        "pdf": "PDF",
        "cbz": "CBZ",
        "merged_filename": "Merged filename",
        "proposed_filename": "Proposed filename (edit if you want):",
        "invalid_name": "Filename cannot be empty.",
        "choose_output_folder": "Choose output folder",
        "created": "Created:",
        "no_outputs": "No outputs produced (error or cancelled).",
        "extract_failed": "Cannot extract archive for inspection.",
        "cannot_extract": "Cannot extract archive: install 7z or patool.",
        "pdf_render_req": "PDF processing requires PyMuPDF (fitz).",
        "ok": "OK",
        "cancel": "Cancel",
        "open_folder": "Open containing folder",
        "open_file": "Open first output file",
        "done": "Done",
        "app_name": CONFIG.get("app_name", "Saino+ Comic")
    },
    "fa": {
        "add": "افزودن",
        "convert": "تبدیل",
        "delete": "حذف",
        "move_up": "جابه‌جایی بالا",
        "move_down": "جابه‌جایی پایین",
        "clear": "پاک کردن همه",
        "sort": "مرتب‌سازی:",
        "settings": "تنظیمات",
        "no_sources_msg": "هیچ منبعی اضافه نشده است.",
        "separate": "هر منبع جدا (تکی)",
        "merge": "ادغام در یک فایل",
        "output_format": "فرمت خروجی را انتخاب کنید",
        "pdf": "PDF",
        "cbz": "CBZ",
        "merged_filename": "نام فایل ادغام",
        "proposed_filename": "نام پیشنهادی (در صورت نیاز ویرایش کنید):",
        "invalid_name": "نام نمی‌تواند خالی باشد.",
        "choose_output_folder": "پوشه خروجی را انتخاب کنید",
        "created": "ساخته شد:",
        "no_outputs": "هیچ فایلی ساخته نشد (خطا یا لغو).",
        "extract_failed": "استخراج آرشیو برای بازبینی ممکن نیست.",
        "cannot_extract": "استخراج آرشیو ممکن نیست: 7z یا patool نصب کنید.",
        "pdf_render_req": "پردازش PDF نیاز به PyMuPDF (fitz) دارد.",
        "ok": "تأیید",
        "cancel": "انصراف",
        "open_folder": "باز کردن پوشه حاوی خروجی",
        "open_file": "باز کردن اولین فایل خروجی",
        "done": "پایان",
        "app_name": CONFIG.get("app_name", "Saino+ Comic")
    }
}

def tr(key: str) -> str:
    lang = CONFIG.get("language", "fa")
    return STRINGS.get(lang, STRINGS["fa"]).get(key, key)

# ---------- helpers ----------
def natural_sort_key(s: str):
    return [int(x) if x.isdigit() else x.lower() for x in re.split(r'(\d+)', s)]

def extract_last_number(s: str) -> Optional[int]:
    nums = re.findall(r'(\d+)', s)
    if not nums:
        return None
    return int(nums[-1])

def remove_numbers_from_name(s: str):
    return re.sub(r'\d+', '', s).strip().strip('_- ')

def run_7z_extract(archive_path: str, dest_dir: str) -> bool:
    try:
        cmd = ["7z", "x", "-y", f"-o{dest_dir}", archive_path]
        p = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return p.returncode == 0
    except Exception:
        return False

# ---------- Source model ----------
_next_id = 1
def make_source(path: str) -> Dict:
    global _next_id
    p = Path(path)
    ext = p.suffix.lower()
    if p.is_dir():
        typ = "folder"
    elif ext == ".pdf":
        typ = "pdf"
    elif ext in CONTAINER_EXTS:
        typ = "archive"
    elif ext in IMAGE_EXTS:
        typ = "image"
    else:
        typ = "other"
    src = {"id": _next_id, "path": path, "type": typ, "label": os.path.basename(path),
           "temp": None, "content_override": None, "added_at": time.time()}
    _next_id += 1
    return src

def session_save_sources(sources: List[Dict]):
    serial = []
    for s in sources:
        copy = {k:v for k,v in s.items() if k!='temp'}
        serial.append(copy)
    SESSION['sources'] = serial
    save_json(SESSION_PATH, SESSION)

def session_load_sources() -> List[Dict]:
    loaded = []
    for s in SESSION.get('sources', []):
        src = s.copy()
        loaded.append(src)
    return loaded

# ---------- Contents Editor (supports multi-select delete) ----------
class ContentsEditor(QDialog):
    def __init__(self, parent, src: Dict):
        super().__init__(parent)
        self.src = src
        self.setWindowTitle(self.src.get("label","Contents"))
        self.resize(700,420)
        self.working_dir = None
        self.files: List[str] = []

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)

        self.sort_combo = QComboBox()
        self.sort_combo.addItems(["Manual","Name (natural)","Number"])
        self.sort_combo.currentIndexChanged.connect(self.on_sort)

        btn_up = QPushButton("⬆")
        btn_down = QPushButton("⬇")
        btn_delete = QPushButton(tr("delete"))
        btn_apply = QPushButton(tr("ok") if tr("ok") else "OK")
        btn_cancel = QPushButton(tr("cancel") if tr("cancel") else "Cancel")

        btn_up.clicked.connect(self.move_up)
        btn_down.clicked.connect(self.move_down)
        btn_delete.clicked.connect(self.delete_selected)
        btn_apply.clicked.connect(self.accept)
        btn_cancel.clicked.connect(self.reject)

        top = QHBoxLayout(); top.addWidget(QLabel(tr("sort"))); top.addWidget(self.sort_combo); top.addStretch()
        btn_row = QHBoxLayout(); btn_row.addWidget(btn_up); btn_row.addWidget(btn_down); btn_row.addWidget(btn_delete); btn_row.addStretch(); btn_row.addWidget(btn_apply); btn_row.addWidget(btn_cancel)

        layout = QVBoxLayout()
        layout.addLayout(top)
        layout.addWidget(self.list_widget)
        layout.addLayout(btn_row)
        self.setLayout(layout)

        self.prepare()

    def prepare(self):
        p = Path(self.src["path"])
        if self.src["type"]=="archive":
            tmp = tempfile.mkdtemp(prefix="ex_contents_")
            success=False
            if HAS_PATOOL:
                try:
                    patoolib.extract_archive(self.src["path"], outdir=tmp, verbose=False); success=True
                except Exception:
                    success=False
            if not success:
                if run_7z_extract(self.src["path"], tmp): success=True
            if not success:
                try:
                    shutil.unpack_archive(self.src["path"], tmp); success=True
                except Exception:
                    success=False
            if not success:
                QMessageBox.warning(self, tr("extract_failed"), tr("extract_failed"))
                self.working_dir=None; return
            self.working_dir=tmp; self.src["temp"]=tmp
        else:
            self.working_dir=self.src["path"]

        if self.src.get("content_override"):
            self.files = [f for f in self.src["content_override"] if os.path.exists(f)]
        else:
            collected=[]
            if os.path.isdir(self.working_dir):
                for name in sorted(os.listdir(self.working_dir), key=natural_sort_key):
                    full = os.path.join(self.working_dir, name)
                    if os.path.isdir(full):
                        # include if has priority or images
                        has_pr=False; has_im=False
                        for r,d,fs in os.walk(full):
                            for ff in fs:
                                ext=os.path.splitext(ff)[1].lower()
                                if ext in PRIORITY_EXTS: has_pr=True; break
                                if ext in IMAGE_EXTS: has_im=True
                            if has_pr: break
                        if has_pr: collected.append(full)
                        elif has_im: collected.append(full)
                    else:
                        ext=os.path.splitext(name)[1].lower()
                        if ext in PRIORITY_EXTS or ext in IMAGE_EXTS:
                            collected.append(full)
            else:
                if os.path.isfile(self.working_dir): collected=[self.working_dir]
            self.files = collected
        self.reload()

    def reload(self):
        self.list_widget.clear()
        for p in self.files:
            label = os.path.basename(p)
            if os.path.isdir(p): label=f"[Folder] {label}"
            self.list_widget.addItem(label)

    def move_up(self):
        sel = sorted({idx.row() for idx in self.list_widget.selectedIndexes()})
        if not sel: return
        i = sel[0]
        if i>0:
            # swap the block up by one
            block = [self.files[j] for j in sel]
            for j in sel:
                del self.files[j]
            insert_at = sel[0]-1
            for k, item in enumerate(block):
                self.files.insert(insert_at+k, item)
            self.reload()
            self.list_widget.setCurrentRow(insert_at)

    def move_down(self):
        sel = sorted({idx.row() for idx in self.list_widget.selectedIndexes()})
        if not sel: return
        i = sel[-1]
        if i < len(self.files)-1:
            block = [self.files[j] for j in sel]
            for j in reversed(sel):
                del self.files[j]
            insert_at = sel[-1]+1 - len(block) + 1
            for k, item in enumerate(block):
                self.files.insert(insert_at+k, item)
            self.reload()
            self.list_widget.setCurrentRow(insert_at+len(block)-1)

    def delete_selected(self):
        sel = sorted({idx.row() for idx in self.list_widget.selectedIndexes()}, reverse=True)
        for i in sel:
            if 0<=i<len(self.files): del self.files[i]
        self.reload()

    def on_sort(self, _):
        mode = self.sort_combo.currentText()
        if mode=="Manual": return
        if mode=="Name (natural)":
            self.files.sort(key=lambda p: natural_sort_key(os.path.basename(p)))
        elif mode=="Number":
            def keyfn(p):
                n = extract_last_number(os.path.basename(p))
                return (n if n is not None else 10**9, natural_sort_key(os.path.basename(p)))
            self.files.sort(key=keyfn)
        self.reload()

    def accept(self):
        self.src["content_override"]=self.files.copy()
        super().accept()

# ---------- Worker thread (cancel-safe, checks _cancel frequently) ----------
class BatchConvertThread(QThread):
    progress = Signal(int)
    message = Signal(str)
    finished_signal = Signal(list)
    detailed = Signal(dict)

    def __init__(self, sources: List[Dict], out_dir: str, merge: bool=False, out_format: str='PDF', quality: int=95, dpi: int=300):
        super().__init__()
        self.sources = sources
        self.out_dir = out_dir
        self.merge = merge
        self.out_format = out_format.upper()
        self.quality = quality
        self.dpi = dpi
        self._temp_dirs: List[str] = []
        self._cancel = False

    def cancel(self):
        self._cancel = True

    def run(self):
        outputs=[]
        try:
            os.makedirs(self.out_dir, exist_ok=True)
            total_sources = len(self.sources)
            processed_sources = 0
            for si, src in enumerate(self.sources, start=1):
                if self._cancel: break
                label = src.get("label", os.path.basename(src.get("path","")))
                self.message.emit(f"{tr('settings')} {si}/{total_sources}: {label}")
                items = src.get("content_override") or [src["path"]]
                # gather images/pages for this source
                all_images=[]
                for it in items:
                    if self._cancel: break
                    imgs = self._gather_images_for_item(it)
                    all_images.extend(imgs)
                page_total = len(all_images)
                self.detailed.emit({"source_index":si,"source_total":total_sources,"page":0,"page_total":page_total})
                if page_total==0:
                    self.message.emit(f"No pages in {label}")
                    processed_sources += 1
                    self.progress.emit(int(processed_sources/total_sources*100))
                    continue
                base_name = remove_numbers_from_name(label) or f"source{src.get('id',si)}"
                out_name = base_name if not self.merge else "merged"
                # check cancel before heavy operations
                if self._cancel:
                    break
                if self.out_format=='CBZ':
                    outp = self._make_cbz(all_images, self.out_dir, out_name)
                    if outp: outputs.append(outp)
                    if self._cancel: break
                else:
                    outp = self._make_pdf_from_images_with_progress(all_images, self.out_dir, out_name, si, total_sources)
                    if outp: outputs.append(outp)
                    if self._cancel: break
                processed_sources += 1
                self.progress.emit(int(processed_sources/total_sources*100))
            # merge if requested and not cancelled
            if self.merge and self.out_format=='PDF' and outputs and not self._cancel and HAS_PYPDF:
                try:
                    merged=os.path.join(self.out_dir,"merged.pdf")
                    merger=PdfMerger()
                    for p in outputs:
                        if self._cancel: break
                        merger.append(p)
                    if not self._cancel:
                        with open(merged,"wb") as f:
                            merger.write(f)
                        merger.close()
                        outputs=[merged]
                        self.message.emit("Merged into merged.pdf")
                except Exception as e:
                    self.message.emit(f"Merge failed: {e}")
            self._cleanup()
            self.finished_signal.emit(outputs if not self._cancel else [])
        except Exception as e:
            self._cleanup()
            self.message.emit(f"Worker error: {e}")
            self.finished_signal.emit([])

    def _gather_images_for_item(self, item):
        imgs=[]
        if os.path.isdir(item):
            for name in sorted(os.listdir(item), key=natural_sort_key):
                if self._cancel: break
                full = os.path.join(item, name)
                if os.path.isdir(full):
                    imgs.extend(self._gather_images_for_item(full))
                else:
                    ext = os.path.splitext(name)[1].lower()
                    if ext in IMAGE_EXTS:
                        imgs.append(full)
                    elif ext in PRIORITY_EXTS:
                        imgs.extend(self._gather_images_for_item(full))
        elif os.path.isfile(item):
            ext = os.path.splitext(item)[1].lower()
            if ext in IMAGE_EXTS:
                imgs.append(item)
            elif ext=='.pdf':
                if not HAS_FITZ:
                    self.message.emit(tr("pdf_render_req"))
                    return []
                tmp=tempfile.mkdtemp(prefix="pdfimg_"); self._temp_dirs.append(tmp)
                try:
                    doc=fitz.open(item)
                    for i,page in enumerate(doc, start=1):
                        if self._cancel: break
                        pix=page.get_pixmap(dpi=self.dpi)
                        pth=os.path.join(tmp,f"{i:04}.jpg"); pix.save(pth); imgs.append(pth)
                    doc.close()
                except Exception as e:
                    self.message.emit(f"PDF render error: {e}")
            elif ext in CONTAINER_EXTS:
                tmp=tempfile.mkdtemp(prefix="ex_"); self._temp_dirs.append(tmp)
                success=False
                if HAS_PATOOL:
                    try:
                        patoolib.extract_archive(item, outdir=tmp, verbose=False); success=True
                    except Exception: success=False
                if not success:
                    if run_7z_extract(item,tmp): success=True
                if not success:
                    try:
                        shutil.unpack_archive(item,tmp); success=True
                    except Exception: success=False
                if not success:
                    self.message.emit(tr("cannot_extract")); return []
                for root,_,files in os.walk(tmp):
                    for f in sorted(files, key=natural_sort_key):
                        if self._cancel: break
                        if os.path.splitext(f)[1].lower() in IMAGE_EXTS:
                            imgs.append(os.path.join(root,f))
        return imgs

    def _make_cbz(self, images, out_dir, base_name):
        try:
            os.makedirs(out_dir, exist_ok=True)
            out_path=os.path.join(out_dir, f"{base_name}.cbz")
            with zipfile.ZipFile(out_path,'w',compression=zipfile.ZIP_STORED) as z:
                for i,p in enumerate(images, start=1):
                    if self._cancel:
                        # abort and remove partial file
                        try: z.close()
                        except: pass
                        try: os.remove(out_path)
                        except: pass
                        return None
                    arc=f"{i:04}{os.path.splitext(p)[1].lower()}"
                    try: z.write(p,arc)
                    except: pass
            return out_path
        except Exception as e:
            self.message.emit(f"CBZ error: {e}"); return None

    def _make_pdf_from_images_with_progress(self, images, out_dir, base_name, source_index, total_sources):
        try:
            if not images: return None
            os.makedirs(out_dir, exist_ok=True)
            out_path=os.path.join(out_dir, f"{base_name}.pdf")
            # check cancel before heavy atomic call
            if self._cancel: return None
            if HAS_IMG2PDF:
                try:
                    # img2pdf is atomic; check cancel just before
                    if self._cancel: return None
                    with open(out_path,"wb") as f:
                        f.write(img2pdf.convert([str(p) for p in images]))
                    # emit per-page done
                    for i in range(1,len(images)+1):
                        if self._cancel: break
                        self.detailed.emit({"source_index":source_index,"source_total":total_sources,"page":i,"page_total":len(images)})
                    if self._cancel:
                        try: os.remove(out_path)
                        except: pass
                        return None
                    return out_path
                except Exception as e:
                    self.message.emit(f"img2pdf failed: {e} -> fallback PIL")
            # PIL fallback: build multipage while checking cancel
            first=Image.open(images[0])
            if first.mode!="RGB": first=first.convert("RGB")
            others=[]
            self.detailed.emit({"source_index":source_index,"source_total":total_sources,"page":1,"page_total":len(images)})
            for idx,p in enumerate(images[1:], start=2):
                if self._cancel:
                    try: first.close()
                    except: pass
                    for im in others:
                        try: im.close()
                        except: pass
                    return None
                try:
                    im=Image.open(p)
                    if im.mode!="RGB": im=im.convert("RGB")
                    others.append(im)
                    self.detailed.emit({"source_index":source_index,"source_total":total_sources,"page":idx,"page_total":len(images)})
                except Exception:
                    continue
            if self._cancel:
                try: first.close()
                except: pass
                for im in others:
                    try: im.close()
                    except: pass
                return None
            first.save(out_path, save_all=True, append_images=others, quality=self.quality)
            try: first.close()
            except: pass
            for im in others:
                try: im.close()
                except: pass
            return out_path
        except Exception as e:
            self.message.emit(f"PDF error: {e}")
            return None

    def _cleanup(self):
        for d in self._temp_dirs:
            try: shutil.rmtree(d, ignore_errors=True)
            except: pass
        self._temp_dirs=[]

# ---------- Main UI (no preview, multi-delete enabled) ----------
class ImageToPDF(QWidget):
    def __init__(self):
        super().__init__()
        app_name = CONFIG.get("app_name","Saino+ Comic")
        self.setWindowTitle(app_name)
        self.resize(1000,640)
        # restore session sources
        self.sources: List[Dict] = session_load_sources()
        self.sort_mode = CONFIG.get("sort_mode","Manual")

        # widgets
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_widget.currentItemChanged.connect(self.on_selection_changed)
        self.list_widget.itemDoubleClicked.connect(self.on_double)

        self.progress_bar = QProgressBar(); self.progress_bar.setVisible(False)
        self.status_label = QLabel("")

        # buttons
        self.btn_add = QPushButton(tr("add")); self._full(self.btn_add); self.btn_add.clicked.connect(self.add_sources)
        self.btn_convert = QPushButton(tr("convert")); self._full(self.btn_convert); self.btn_convert.clicked.connect(self.convert_dialog)
        self.btn_cancel_op = QPushButton(tr("cancel")); self._full(self.btn_cancel_op); self.btn_cancel_op.clicked.connect(self.cancel_operation); self.btn_cancel_op.setVisible(False)
        self.btn_delete = QPushButton(tr("delete")); self._full(self.btn_delete); self.btn_delete.clicked.connect(self.delete_selected)
        self.btn_move_up = QPushButton(tr("move_up")); self._full(self.btn_move_up); self.btn_move_up.clicked.connect(self.move_up)
        self.btn_move_down = QPushButton(tr("move_down")); self._full(self.btn_move_down); self.btn_move_down.clicked.connect(self.move_down)
        self.btn_clear = QPushButton(tr("clear")); self._full(self.btn_clear); self.btn_clear.clicked.connect(self.clear_all)

        # controls
        self.sort_combo = QComboBox(); self.sort_combo.addItems(["Manual","Name (natural)","Added time","Number"])
        try:
            idx=self.sort_combo.findText(CONFIG.get("sort_mode","Manual"))
            if idx>=0: self.sort_combo.setCurrentIndex(idx); self.sort_mode=self.sort_combo.currentText()
        except: pass
        self.sort_combo.currentIndexChanged.connect(self.on_sort_changed)

        self.lang_combo = QComboBox(); self.lang_combo.addItems(["فارسی","English"])
        self.lang_combo.setCurrentIndex(0 if CONFIG.get("language","fa")=="fa" else 1)
        self.lang_combo.currentIndexChanged.connect(self.on_lang_changed)

        self.quality_spin = QSpinBox(); self.quality_spin.setRange(1,100); self.quality_spin.setValue(CONFIG.get("quality_default",95)); self.quality_spin.setSuffix("%")
        self.dpi_cb = QCheckBox("Custom DPI"); self.dpi_cb.setChecked(CONFIG.get("dpi_enabled",False)); self.dpi_cb.stateChanged.connect(self.on_dpi_cb)
        self.dpi_spin = QSpinBox(); self.dpi_spin.setRange(50,1200); self.dpi_spin.setValue(CONFIG.get("dpi_value",300)); self.dpi_spin.setEnabled(CONFIG.get("dpi_enabled",False))

        # layout
        ctrl_group = QGroupBox(tr("settings"))
        ctrl_layout = QVBoxLayout()
        row1=QHBoxLayout(); row1.addWidget(self.btn_add); row1.addWidget(self.btn_convert); row1.addWidget(self.btn_cancel_op)
        row2=QHBoxLayout(); row2.addWidget(self.btn_delete); row2.addWidget(self.btn_clear); row2.addWidget(self.btn_move_up); row2.addWidget(self.btn_move_down)
        row3=QHBoxLayout(); row3.addWidget(QLabel(tr("sort"))); row3.addWidget(self.sort_combo); row3.addStretch(); row3.addWidget(self.lang_combo)
        ctrl_layout.addLayout(row1); ctrl_layout.addLayout(row2); ctrl_layout.addLayout(row3)
        form=QFormLayout(); form.addRow("PDF Quality:", self.quality_spin); form.addRow(self.dpi_cb, self.dpi_spin)
        ctrl_layout.addLayout(form); ctrl_layout.addWidget(self.progress_bar); ctrl_layout.addWidget(self.status_label)
        ctrl_group.setLayout(ctrl_layout)

        left = QVBoxLayout(); left.addWidget(ctrl_group); left.addWidget(self.list_widget)
        main = QHBoxLayout(); main.addLayout(left,60)
        self.setLayout(main)

        self.batch_thread: Optional[BatchConvertThread]=None
        self.refresh_list_widget()

    def _full(self, w):
        w.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed); w.setMinimumHeight(36)

    def on_dpi_cb(self,s):
        CONFIG['dpi_enabled'] = (s==Qt.Checked); save_json(CONFIG_PATH,CONFIG)

    def on_lang_changed(self, idx):
        CONFIG['language'] = "fa" if idx==0 else "en"; save_json(CONFIG_PATH,CONFIG)
        self.update_texts()

    def update_texts(self):
        # update button texts according to language
        self.btn_add.setText(tr("add"))
        self.btn_convert.setText(tr("convert"))
        self.btn_delete.setText(tr("delete"))
        self.btn_move_up.setText(tr("move_up"))
        self.btn_move_down.setText(tr("move_down"))
        self.btn_clear.setText(tr("clear"))
        # group titles
        # window title
        self.setWindowTitle(CONFIG.get("app_name","Saino+ Comic"))

    def on_sort_changed(self,_):
        self.sort_mode = self.sort_combo.currentText(); CONFIG['sort_mode']=self.sort_mode; save_json(CONFIG_PATH,CONFIG)
        self.apply_sort(); self.refresh_list_widget()

    def refresh_list_widget(self):
        self.list_widget.clear()
        for i,src in enumerate(self.sources, start=1):
            label = f"{i} - {src.get('label',os.path.basename(src.get('path','')))}"
            self.list_widget.addItem(label)

    def apply_sort(self):
        mode=self.sort_mode
        if mode=="Manual": return
        if mode=="Name (natural)":
            self.sources.sort(key=lambda s: natural_sort_key(s.get('label','')))
        elif mode=="Added time":
            self.sources.sort(key=lambda s: s.get('added_at',0))
        elif mode=="Number":
            def k(s):
                n = extract_last_number(s.get('label',''))
                return (n if n is not None else 10**9, natural_sort_key(s.get('label','')))
            self.sources.sort(key=k)

    def add_sources(self):
        mb=QMessageBox(self); mb.setWindowTitle(tr("add")); mb.setText(tr("add"))
        files_btn=mb.addButton("Select files (images/pdf/archives)", QMessageBox.ActionRole)
        folder_btn=mb.addButton("Add folder", QMessageBox.ActionRole); mb.addButton(tr("cancel"), QMessageBox.RejectRole); mb.exec()
        clicked=mb.clickedButton()
        if clicked==files_btn:
            files,_=QFileDialog.getOpenFileNames(self,"Select files","", "Supported (*.zip *.cbz *.cbr *.rar *.pdf *.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff)")
            for f in files:
                if not f: continue
                s=make_source(f); self.sources.append(s)
            self.apply_sort(); self.refresh_list_widget(); session_save_sources(self.sources)
        elif clicked==folder_btn:
            folder=QFileDialog.getExistingDirectory(self,"Select folder")
            if not folder: return
            added=self._scan_and_add_folder(folder)
            if not added: QMessageBox.information(self,"",tr("no_sources_msg"))
            self.apply_sort(); self.refresh_list_widget(); session_save_sources(self.sources)

    def _scan_and_add_folder(self, folder_path):
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
            ext=os.path.splitext(f)[1].lower()
            if ext in PRIORITY_EXTS: top_pr=True
            if ext in IMAGE_EXTS: top_im=True
        if found:
            for sub in sorted(set(found), key=natural_sort_key):
                s=make_source(sub); self.sources.append(s)
            return True
        if top_pr:
            s=make_source(folder_path); self.sources.append(s); return True
        if top_im:
            s=make_source(folder_path); self.sources.append(s); return True
        return False

    def on_selection_changed(self, cur, prev):
        pass

    def on_double(self, item):
        row=self.list_widget.currentRow()
        if row<0 or row>=len(self.sources): return
        src=self.sources[row]
        dlg=ContentsEditor(self, src)
        if dlg.exec():
            self.apply_sort(); self.refresh_list_widget(); session_save_sources(self.sources)

    def delete_selected(self):
        sel = sorted({idx.row() for idx in self.list_widget.selectedIndexes()}, reverse=True)
        for r in sel:
            if 0<=r<len(self.sources):
                s=self.sources.pop(r)
                if s.get("temp"):
                    try: shutil.rmtree(s["temp"], ignore_errors=True)
                    except: pass
        self.refresh_list_widget(); session_save_sources(self.sources)

    def move_up(self):
        sel = sorted({idx.row() for idx in self.list_widget.selectedIndexes()})
        if not sel: return
        i=min(sel)
        if i>0:
            self.sources[i-1], self.sources[i] = self.sources[i], self.sources[i-1]
            self.refresh_list_widget()
            self.list_widget.setCurrentRow(i-1); session_save_sources(self.sources)

    def move_down(self):
        sel = sorted({idx.row() for idx in self.list_widget.selectedIndexes()})
        if not sel: return
        i=max(sel)
        if i < len(self.sources)-1:
            self.sources[i+1], self.sources[i] = self.sources[i], self.sources[i+1]
            self.refresh_list_widget()
            self.list_widget.setCurrentRow(i+1); session_save_sources(self.sources)

    def clear_all(self):
        for s in self.sources:
            if s.get("temp"):
                try: shutil.rmtree(s["temp"], ignore_errors=True)
                except: pass
        self.sources.clear(); self.refresh_list_widget(); session_save_sources(self.sources)

    def convert_dialog(self):
        if not self.sources:
            QMessageBox.warning(self,"",tr("no_sources_msg")); return
        mb=QMessageBox(self); mb.setWindowTitle(tr("settings")); mb.setText(tr("settings"))
        sep=mb.addButton(tr("separate"), QMessageBox.ActionRole)
        merge=mb.addButton(tr("merge"), QMessageBox.ActionRole); mb.addButton(tr("cancel"), QMessageBox.RejectRole); mb.exec()
        if mb.clickedButton() is None: return
        do_merge = (mb.clickedButton()==merge)
        mb2=QMessageBox(self); mb2.setWindowTitle(tr("output_format")); mb2.setText(tr("output_format"))
        pdf_b=mb2.addButton(tr("pdf"), QMessageBox.ActionRole); cbz_b=mb2.addButton(tr("cbz"), QMessageBox.ActionRole); mb2.addButton(tr("cancel"), QMessageBox.RejectRole); mb2.exec()
        if mb2.clickedButton() is None: return
        out_fmt = 'PDF' if mb2.clickedButton()==pdf_b else 'CBZ'
        final_name=None
        if do_merge:
            maxn=0
            for s in self.sources:
                cand = s.get("content_override") or [s["path"]]
                for c in cand:
                    if os.path.isdir(c):
                        for f in os.listdir(c):
                            nums=re.findall(r'(\d+)', f)
                            if nums: maxn=max(maxn,int(nums[-1]))
                    else:
                        nums=re.findall(r'(\d+)', os.path.basename(c))
                        if nums: maxn=max(maxn,int(nums[-1]))
            tens = (maxn//10)*10 if maxn>0 else 0
            base = remove_numbers_from_name(os.path.basename(self.sources[0]['path'])) or "output"
            proposed=f"{base}_ch{tens}_vol{(tens//10) if tens>0 else 0}"
            text,ok = QInputDialog.getText(self,tr("merged_filename"), tr("proposed_filename"), text=proposed)
            if not ok: return
            final_name=text.strip()
            if final_name=="": QMessageBox.warning(self,"Invalid",tr("invalid_name")); return
        out_dir = QFileDialog.getExistingDirectory(self,tr("choose_output_folder"))
        if not out_dir: return
        quality = self.quality_spin.value()
        dpi = self.dpi_spin.value() if self.dpi_cb.isChecked() else CONFIG.get("dpi_value",300)
        # start convert
        self.progress_bar.setVisible(True); self.progress_bar.setValue(0); self.status_label.setText("")
        self.btn_cancel_op.setVisible(True); self.btn_cancel_op.setEnabled(True)
        thread_sources=[s.copy() for s in self.sources]
        self.batch_thread = BatchConvertThread(thread_sources, out_dir, merge=do_merge, out_format=out_fmt, quality=quality, dpi=dpi)
        self.batch_thread.progress.connect(lambda v: self.progress_bar.setValue(v))
        self.batch_thread.message.connect(lambda m: self.status_label.setText(m))
        self.batch_thread.detailed.connect(self.on_detailed_progress)
        self.batch_thread.finished_signal.connect(self.on_finished)
        self.batch_thread.start()

    def on_detailed_progress(self, info: dict):
        si = info.get("source_index"); st = info.get("source_total")
        p = info.get("page"); pt = info.get("page_total")
        self.status_label.setText(f"{si}/{st} — {p}/{pt}")

    def on_finished(self, outputs: List[str]):
        self.progress_bar.setVisible(False); self.btn_cancel_op.setVisible(False)
        if not outputs:
            QMessageBox.information(self,tr("done"), tr("no_outputs")); return
        msg = tr("created") + "\n" + "\n".join(outputs)
        dlg=QMessageBox(self); dlg.setWindowTitle(tr("done")); dlg.setText(msg)
        open_f = dlg.addButton(tr("open_folder"), QMessageBox.ActionRole)
        open_file = dlg.addButton(tr("open_file"), QMessageBox.ActionRole)
        dlg.addButton(tr("cancel"), QMessageBox.RejectRole); dlg.exec()
        if dlg.clickedButton()==open_f:
            try:
                if sys.platform.startswith("win"): os.startfile(os.path.dirname(outputs[0]))
                elif sys.platform=="darwin": subprocess.call(["open", os.path.dirname(outputs[0])])
                else: subprocess.call(["xdg-open", os.path.dirname(outputs[0])])
            except: pass
        elif dlg.clickedButton()==open_file:
            try:
                f=outputs[0]
                if sys.platform.startswith("win"): os.startfile(f)
                elif sys.platform=="darwin": subprocess.call(["open", f])
                else: subprocess.call(["xdg-open", f])
            except: pass
        session_save_sources(self.sources)

    def cancel_operation(self):
        if self.batch_thread:
            self.batch_thread.cancel()
            self.status_label.setText(tr("cancel") + "...")
            self.btn_cancel_op.setEnabled(False)

    def closeEvent(self, ev):
        CONFIG['sort_mode']=self.sort_mode
        CONFIG['quality_default']=self.quality_spin.value()
        CONFIG['dpi_enabled']=self.dpi_cb.isChecked()
        CONFIG['dpi_value']=self.dpi_spin.value()
        save_json(CONFIG_PATH, CONFIG)
        session_save_sources(self.sources)
        # cleanup temps
        for s in self.sources:
            if s.get("temp"):
                try: shutil.rmtree(s["temp"], ignore_errors=True)
                except: pass
        # compile pyc for this file for small optimization
        try:
            compileall.compile_file(__file__, force=False, quiet=1)
        except Exception:
            pass
        ev.accept()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    w = ImageToPDF()
    # ensure app name/title reflects config
    app_name = CONFIG.get("app_name", "Saino+ Comic")
    w.setWindowTitle(app_name)
    w.show()
    sys.exit(app.exec())
