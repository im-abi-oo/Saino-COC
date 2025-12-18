# main.py
# Saino COC Ultimate v3.2 FINAL - Polished UI, Memory-safe, DPI default-preserve
# Developer: im_abi + assistant (finalized)

import sys, os, re, shutil, tempfile, zipfile, subprocess, json, io, gc
from pathlib import Path
from PIL import Image, ImageFile, ImageEnhance
ImageFile.LOAD_TRUNCATED_IMAGES = True

from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog, QListWidget, QLabel,
    QMessageBox, QHBoxLayout, QAbstractItemView, QSpinBox, QFormLayout, QGroupBox,
    QCheckBox, QProgressBar, QInputDialog, QDialog, QComboBox, QFrame, QTabWidget,
    QScrollArea, QSlider, QSizePolicy, QSplitter, QStyle, QToolButton
)
from PySide6.QtGui import QFont, QPixmap, QAction
from PySide6.QtCore import Qt, QThread, Signal, QSize

# Optional libs
try:
    import img2pdf
    HAS_IMG2PDF = True
except Exception:
    HAS_IMG2PDF = False

try:
    import fitz  # PyMuPDF
    HAS_FITZ = True
except Exception:
    HAS_FITZ = False

# ---------------- CONFIG ----------------
APP_NAME = "Saino COC Ultimate"
VERSION = "3.2 FINAL"
CONFIG_PATH = Path.home() / ".saino_coc_v3_config.json"
SESSION_PATH = Path.home() / ".saino_coc_v3_session.json"
IMAGE_EXTS = {'.jpg', '.jpeg', '.png', '.webp', '.bmp', '.tiff', '.tif'}
CONTAINER_EXTS = {'.zip', '.cbz', '.rar', '.cbr', '.7z', '.tar'}

DEFAULT_CONFIG = {
    "language": "fa",
    "quality": 90,
    "dpi": 300,
    "use_custom_dpi": False,
    "grayscale": False,
    "enhancement": {
        "brightness": 1.0,
        "contrast": 1.0,
        "sharpness": 1.0,
        "resize_w": 0
    }
}

cconf = DEFAULT_CONFIG.copy()

# ---------------- STYLES ----------------
STYLESHEET = """
QWidget { background: #0f1115; color: #e6eef2; font-family: 'Segoe UI', Roboto, Arial; }
QTabWidget::pane { border: none; }
QGroupBox { border: 1px solid #263238; border-radius: 8px; margin-top: 18px; color: #80deea; }
QListWidget { background: #071016; border: 1px solid #263238; }
QPushButton#Accent { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #006064, stop:1 #00acc1); color: white; padding:8px 14px; border-radius:8px; }
QPushButton { background: #12171a; border: 1px solid #263238; padding:6px 10px; border-radius:6px; }
QPushButton:hover { border-color: #00e5ff; }
QProgressBar { background:#071016; border:1px solid #263238; height:18px; }
QProgressBar::chunk { background: #00e5ff; }
QLabel#Title { font-size:16pt; color:#00e5ff; font-weight:700 }
"""

# ---------------- UTIL ----------------
def load_config():
    global cconf
    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                cconf.update(json.load(f))
        except Exception:
            pass


def save_config():
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(cconf, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]


def get_7z_path():
    if sys.platform == 'win32':
        p1 = Path("C:/Program Files/7-Zip/7z.exe")
        if p1.exists(): return str(p1)
    return '7z'

# ---------------- IMAGE PROCESSOR ----------------
class ImageProcessor:
    @staticmethod
    def process(path, cfg):
        try:
            with Image.open(path) as im:
                # Mode
                im = im.convert('L') if cfg.get('grayscale') else (im.convert('RGB') if im.mode not in ('RGB','L') else im.copy())

                # Resize
                rw = cfg['enhancement'].get('resize_w', 0)
                if rw and im.width > rw:
                    h = int(im.height * (rw / im.width))
                    im = im.resize((rw, h), Image.Resampling.LANCZOS)

                # Enhancements
                br = cfg['enhancement'].get('brightness', 1.0)
                if br != 1.0: im = ImageEnhance.Brightness(im).enhance(br)
                co = cfg['enhancement'].get('contrast', 1.0)
                if co != 1.0: im = ImageEnhance.Contrast(im).enhance(co)
                sh = cfg['enhancement'].get('sharpness', 1.0)
                if sh != 1.0: im = ImageEnhance.Sharpness(im).enhance(sh)

                return im.copy()
        except Exception as e:
            print('ImageProcessor error:', e)
            return None

# ---------------- WORKER ----------------
class SuperWorker(QThread):
    progress = Signal(int)
    log = Signal(str)
    finished = Signal(list)

    def __init__(self, sources, out_dir, merge, fmt, out_name, cfg):
        super().__init__()
        self.sources = sources
        self.out_dir = out_dir
        self.merge = merge
        self.fmt = fmt
        self.out_name = out_name
        self.cfg = cfg
        self.cancel = False
        self.temp_dirs = []

    def stop(self):
        self.cancel = True

    def run(self):
        outputs = []
        try:
            images_buffer = []
            total = len(self.sources)
            for i, s in enumerate(self.sources):
                if self.cancel: break
                self.log.emit(f"Processing [{i+1}/{total}]: {s['label']}")
                imgs = self._gather_images(s)
                if not imgs: continue
                if self.merge:
                    images_buffer.extend(imgs)
                    self.progress.emit(int((i+1)/total*40))
                else:
                    base = os.path.splitext(s['label'])[0]
                    out = self._create_output(imgs, base)
                    if out: outputs.append(out)
                    self.progress.emit(int((i+1)/total*100))
                    imgs.clear(); gc.collect()

            if self.merge and images_buffer and not self.cancel:
                self.log.emit('Finalizing merged output...')
                final = self._create_output(images_buffer, self.out_name)
                if final: outputs.append(final)
                self.progress.emit(100)

        except Exception as e:
            self.log.emit(f'Error: {e}')
        finally:
            self._cleanup()
            self.finished.emit(outputs)

    def _gather_images(self, src):
        ptype = src['type']
        path = src['path']
        imgs = []

        if ptype == 'folder':
            for r, _, files in os.walk(path):
                for f in sorted(files, key=natural_sort_key):
                    if Path(f).suffix.lower() in IMAGE_EXTS:
                        imgs.append(os.path.join(r, f))

        elif ptype == 'archive':
            tmp = tempfile.mkdtemp(prefix='saino_ext_')
            self.temp_dirs.append(tmp)
            extracted = False
            try:
                if shutil.which('7z'):
                    exe = get_7z_path()
                    subprocess.run([exe, 'x', '-y', f'-o{tmp}', path], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    extracted = True
            except Exception:
                extracted = False
            if not extracted:
                try:
                    import patoolib
                    patoolib.extract_archive(path, outdir=tmp, interactive=False)
                    extracted = True
                except Exception:
                    pass

            for r, _, files in os.walk(tmp):
                for f in sorted(files, key=natural_sort_key):
                    if Path(f).suffix.lower() in IMAGE_EXTS:
                        imgs.append(os.path.join(r, f))

        elif ptype == 'pdf':
            if HAS_FITZ:
                tmp = tempfile.mkdtemp(prefix='saino_pdf_')
                self.temp_dirs.append(tmp)
                try:
                    doc = fitz.open(path)
                    # If user set custom DPI, use it; otherwise preserve default by not passing dpi.
                    use_custom = self.cfg.get('use_custom_dpi', False)
                    target_dpi = int(self.cfg.get('dpi', 300))
                    for idx, page in enumerate(doc):
                        if self.cancel: break
                        if use_custom:
                            pix = page.get_pixmap(dpi=target_dpi)
                        else:
                            pix = page.get_pixmap()  # Keep library default resolution
                        out = os.path.join(tmp, f'p_{idx:05d}.jpg')
                        pix.save(out)
                        imgs.append(out)
                    doc.close()
                except Exception as e:
                    print('PDF extract err:', e)
            else:
                self.log.emit('PyMuPDF not installed — cannot extract PDF pages')

        elif ptype == 'image':
            imgs.append(path)

        return imgs

    def _create_output(self, image_paths, name):
        if not image_paths: return None
        target = os.path.join(self.out_dir, name)

        if self.fmt == 'FOLDER':
            dest = target
            os.makedirs(dest, exist_ok=True)
            for idx, src in enumerate(image_paths):
                if self.cancel: return None
                processed = ImageProcessor.process(src, self.cfg)
                if processed:
                    outp = os.path.join(dest, f'{idx:05d}.jpg')
                    processed.save(outp, quality=self.cfg.get('quality', 90))
                    processed.close()
            return dest

        elif self.fmt == 'CBZ':
            full = target + '.cbz'
            with zipfile.ZipFile(full, 'w', zipfile.ZIP_STORED) as zf:
                for idx, src in enumerate(image_paths):
                    if self.cancel: break
                    proc = ImageProcessor.process(src, self.cfg)
                    if proc:
                        buf = io.BytesIO()
                        proc.save(buf, format='JPEG', quality=self.cfg.get('quality', 90))
                        zf.writestr(f'{idx:05d}.jpg', buf.getvalue())
                        proc.close()
            return full

        elif self.fmt == 'PDF':
            full = target + '.pdf'
            # Convert using img2pdf (memory efficient) if available
            tmp_dir = tempfile.mkdtemp(prefix='saino_pdfout_')
            try:
                temp_files = []
                for idx, src in enumerate(image_paths):
                    if self.cancel: break
                    proc = ImageProcessor.process(src, self.cfg)
                    if proc:
                        outp = os.path.join(tmp_dir, f'{idx:05d}.jpg')
                        save_kwargs = {'quality': self.cfg.get('quality', 90)}
                        # Only set dpi if user enabled custom DPI
                        if self.cfg.get('use_custom_dpi'):
                            dpi = int(self.cfg.get('dpi', 300))
                            save_kwargs['dpi'] = (dpi, dpi)
                        proc.save(outp, **save_kwargs)
                        proc.close()
                        temp_files.append(outp)
                    self.progress.emit(int((idx+1)/len(image_paths)*80))

                if not self.cancel:
                    if HAS_IMG2PDF:
                        with open(full, 'wb') as f:
                            f.write(img2pdf.convert(temp_files))
                    else:
                        # Fallback: use PIL to make PDF (may be memory-heavier)
                        first = None
                        others = []
                        for p in temp_files:
                            try:
                                im = Image.open(p).convert('RGB')
                                if first is None:
                                    first = im
                                else:
                                    others.append(im)
                            except Exception:
                                pass
                        if first:
                            first.save(full, 'PDF', save_all=True, append_images=others)
                            first.close()
                            for o in others: o.close()
                return full if not self.cancel else None
            finally:
                try: shutil.rmtree(tmp_dir, ignore_errors=True)
                except: pass

        return None

    def _cleanup(self):
        for d in self.temp_dirs:
            try: shutil.rmtree(d, ignore_errors=True)
            except: pass
        gc.collect()

# ---------------- MAIN UI ----------------
class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        load_config()
        self.setWindowTitle(f"{APP_NAME} {VERSION}")
        self.resize(1200, 780)
        self.sources = []
        self._setup_ui()
        self._load_session()
        self._refresh_list()

    def _setup_ui(self):
        self.setStyleSheet(STYLESHEET)
        main = QVBoxLayout(self)

        title = QLabel(f"{APP_NAME} - {VERSION}")
        title.setObjectName('Title')
        title.setAlignment(Qt.AlignLeft)
        main.addWidget(title)

        splitter = QSplitter(Qt.Horizontal)

        # Left: List + Toolbar
        left = QWidget()
        llay = QVBoxLayout(left)
        toolbar = QHBoxLayout()
        btn_add = QPushButton('Add Files/Folders')
        btn_add.clicked.connect(self.add_files)
        btn_del = QPushButton('Remove')
        btn_del.clicked.connect(self.delete_items)
        btn_clr = QPushButton('Clear')
        btn_clr.clicked.connect(self.clear_all)
        toolbar.addWidget(btn_add); toolbar.addWidget(btn_del); toolbar.addWidget(btn_clr)

        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.list_widget.itemSelectionChanged.connect(self._on_select)

        llay.addLayout(toolbar)
        llay.addWidget(self.list_widget)

        # Action area
        act = QHBoxLayout()
        self.combo_mode = QComboBox(); self.combo_mode.addItems(['Separate','Merge'])
        self.combo_fmt = QComboBox(); self.combo_fmt.addItems(['PDF','CBZ','FOLDER'])
        self.btn_start = QPushButton('Start Processing'); self.btn_start.setObjectName('Accent'); self.btn_start.setMinimumHeight(44)
        self.btn_start.clicked.connect(self.start_processing)
        act.addWidget(QLabel('Mode:')); act.addWidget(self.combo_mode)
        act.addWidget(QLabel('Format:')); act.addWidget(self.combo_fmt)
        act.addStretch(); act.addWidget(self.btn_start)

        llay.addLayout(act)
        splitter.addWidget(left)

        # Right: Preview + Log + Settings in Tabs
        right = QTabWidget()

        # Preview tab
        tab_preview = QWidget(); pv_l = QVBoxLayout(tab_preview)
        self.preview_label = QLabel('Select an item to preview'); self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setMinimumSize(320,400)
        pv_l.addWidget(self.preview_label)

        # Log list
        self.log_list = QListWidget()
        pv_l.addWidget(QLabel('Log:'))
        pv_l.addWidget(self.log_list)

        right.addTab(tab_preview, 'Preview')

        # Settings tab
        tab_set = QWidget(); s_l = QVBoxLayout(tab_set)
        form = QFormLayout()
        self.spin_qual = QSpinBox(); self.spin_qual.setRange(10,100); self.spin_qual.setValue(cconf.get('quality',90))
        self.chk_gray = QCheckBox('Grayscale'); self.chk_gray.setChecked(cconf.get('grayscale',False))
        self.sl_bright = QSlider(Qt.Horizontal); self.sl_bright.setRange(5,30); self.sl_bright.setValue(int(cconf['enhancement'].get('brightness',1.0)*10))
        self.sl_contrast = QSlider(Qt.Horizontal); self.sl_contrast.setRange(5,30); self.sl_contrast.setValue(int(cconf['enhancement'].get('contrast',1.0)*10))
        self.sl_sharp = QSlider(Qt.Horizontal); self.sl_sharp.setRange(5,40); self.sl_sharp.setValue(int(cconf['enhancement'].get('sharpness',1.0)*10))
        self.spin_resize = QSpinBox(); self.spin_resize.setRange(0,8000); self.spin_resize.setValue(cconf['enhancement'].get('resize_w',0)); self.spin_resize.setSuffix(' px')

        # DPI control: default off -> preserve library default
        self.chk_custom_dpi = QCheckBox('Use custom DPI (override PDF render)')
        self.chk_custom_dpi.setChecked(cconf.get('use_custom_dpi', False))
        self.spin_dpi = QSpinBox(); self.spin_dpi.setRange(72,1200); self.spin_dpi.setValue(cconf.get('dpi',300));
        self.spin_dpi.setEnabled(self.chk_custom_dpi.isChecked())
        self.chk_custom_dpi.toggled.connect(self.spin_dpi.setEnabled)

        form.addRow('Quality (JPEG):', self.spin_qual)
        form.addRow(self.chk_gray)
        form.addRow('Brightness:', self.sl_bright)
        form.addRow('Contrast:', self.sl_contrast)
        form.addRow('Sharpness:', self.sl_sharp)
        form.addRow('Resize Width:', self.spin_resize)
        form.addRow(self.chk_custom_dpi)
        form.addRow('DPI:', self.spin_dpi)

        s_l.addLayout(form)
        # Save settings button
        btn_save_set = QPushButton('Save Settings'); btn_save_set.clicked.connect(self._save_settings)
        s_l.addWidget(btn_save_set)
        right.addTab(tab_set, 'Settings')

        splitter.addWidget(right)
        main.addWidget(splitter)

        # Footer: Progress + Cancel
        footer = QHBoxLayout()
        self.progress = QProgressBar(); self.progress.setValue(0); self.progress.setVisible(False)
        self.btn_cancel = QPushButton('Cancel'); self.btn_cancel.setEnabled(False); self.btn_cancel.clicked.connect(self._cancel)
        footer.addWidget(self.progress); footer.addWidget(self.btn_cancel)
        main.addLayout(footer)

    # --- Session & Sources ---
    def _add_src(self, path):
        if any(s['path'] == path for s in self.sources): return
        p = Path(path)
        ptype = 'other'
        if p.is_dir(): ptype = 'folder'
        elif p.suffix.lower() in CONTAINER_EXTS: ptype = 'archive'
        elif p.suffix.lower() == '.pdf': ptype = 'pdf'
        elif p.suffix.lower() in IMAGE_EXTS: ptype = 'image'
        self.sources.append({'path': str(path), 'type': ptype, 'label': p.name})

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, 'Add', '', 'Files (*.zip *.cbz *.rar *.pdf *.jpg *.png *.webp *.bmp *.tiff)')
        for f in files: self._add_src(f)
        self._refresh_list(); self._save_session()

    def delete_items(self):
        rows = sorted([i.row() for i in self.list_widget.selectedIndexes()], reverse=True)
        for r in rows:
            if r < len(self.sources): del self.sources[r]
        self._refresh_list(); self._save_session()

    def clear_all(self):
        self.sources.clear(); self._refresh_list(); self._save_session()

    def _refresh_list(self):
        self.list_widget.clear()
        for idx, s in enumerate(self.sources):
            icon = '📁' if s['type']=='folder' else ('📦' if s['type']=='archive' else ('🖼️' if s['type']=='image' else '📄'))
            self.list_widget.addItem(f"{idx+1}. {icon} {s['label']}")

    def _on_select(self):
        sel = self.list_widget.currentRow()
        if sel < 0 or sel >= len(self.sources): return
        s = self.sources[sel]
        # preview first image or PDF cover
        preview_path = None
        if s['type'] == 'image': preview_path = s['path']
        elif s['type'] == 'pdf' and HAS_FITZ:
            try:
                doc = fitz.open(s['path'])
                pix = doc.load_page(0).get_pixmap()
                tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.png')
                pix.save(tmp.name); tmp.close()
                preview_path = tmp.name
                doc.close()
            except Exception:
                preview_path = None
        elif s['type'] in ('archive','folder'):
            # show folder icon
            preview_path = None

        if preview_path and os.path.exists(preview_path):
            pix = QPixmap(preview_path)
            pix = pix.scaled(self.preview_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.preview_label.setPixmap(pix)
            # cleanup if was temp
            if preview_path.startswith(tempfile.gettempdir()):
                try: os.remove(preview_path)
                except: pass
        else:
            self.preview_label.setText('No preview available')

    # --- Processing ---
    def start_processing(self):
        if not self.sources:
            QMessageBox.warning(self, '!', 'No sources added')
            return
        is_merge = (self.combo_mode.currentIndex() == 1)
        fmt = self.combo_fmt.currentText()
        final_name = 'Output'
        if is_merge:
            txt, ok = QInputDialog.getText(self, 'Output Name', 'Enter output filename:')
            if ok and txt.strip(): final_name = txt.strip()
            else: return
        out_dir = QFileDialog.getExistingDirectory(self, 'Select Output Folder')
        if not out_dir: return

        # build runtime config
        run_conf = {
            'quality': self.spin_qual.value(),
            'grayscale': self.chk_gray.isChecked(),
            'enhancement': {
                'brightness': self.sl_bright.value()/10.0,
                'contrast': self.sl_contrast.value()/10.0,
                'sharpness': self.sl_sharp.value()/10.0,
                'resize_w': self.spin_resize.value()
            },
            'use_custom_dpi': self.chk_custom_dpi.isChecked(),
            'dpi': self.spin_dpi.value()
        }

        # disable UI
        self.btn_start.setEnabled(False); self.progress.setVisible(True); self.progress.setValue(0); self.log_list.clear()
        self.btn_cancel.setEnabled(True)

        self.worker = SuperWorker(self.sources, out_dir, is_merge, fmt, final_name, run_conf)
        self.worker.progress.connect(self.progress.setValue)
        self.worker.log.connect(self._log)
        self.worker.finished.connect(self._finished)
        self.worker.start()

    def _log(self, msg):
        self.log_list.addItem(msg); self.log_list.scrollToBottom()

    def _finished(self, outputs):
        self.btn_start.setEnabled(True); self.progress.setVisible(False); self.btn_cancel.setEnabled(False)
        if outputs:
            QMessageBox.information(self, APP_NAME, f"Done. Generated {len(outputs)} file(s)")
            for o in outputs: self._log(f"Output: {o}")
        else:
            QMessageBox.information(self, APP_NAME, 'No outputs (maybe cancelled)')
        self._log('--- END ---')

    def _cancel(self):
        try:
            if hasattr(self, 'worker'):
                self.worker.stop(); self._log('Cancel requested...')
                self.btn_cancel.setEnabled(False)
        except Exception:
            pass

    # --- Settings persistence ---
    def _save_settings(self):
        cconf['quality'] = self.spin_qual.value()
        cconf['grayscale'] = self.chk_gray.isChecked()
        cconf['enhancement']['brightness'] = self.sl_bright.value()/10.0
        cconf['enhancement']['contrast'] = self.sl_contrast.value()/10.0
        cconf['enhancement']['sharpness'] = self.sl_sharp.value()/10.0
        cconf['enhancement']['resize_w'] = self.spin_resize.value()
        cconf['use_custom_dpi'] = self.chk_custom_dpi.isChecked()
        cconf['dpi'] = self.spin_dpi.value()
        save_config()
        QMessageBox.information(self, 'Settings', 'Settings saved')

    # --- Session load/save ---
    def _save_session(self):
        try:
            with open(SESSION_PATH, 'w', encoding='utf-8') as f:
                json.dump({'sources': self.sources}, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _load_session(self):
        if SESSION_PATH.exists():
            try:
                with open(SESSION_PATH, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.sources = data.get('sources', [])
            except Exception:
                pass

    def closeEvent(self, e):
        self._save_settings(); self._save_session(); e.accept()

    # Drag & drop
    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls(): e.accept()
    def dropEvent(self, e):
        for u in e.mimeData().urls(): self._add_src(u.toLocalFile())
        self._refresh_list(); self._save_session()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    app.setFont(QFont('Segoe UI', 10))
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
