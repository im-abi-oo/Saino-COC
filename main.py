import sys
import os
import zipfile
import tempfile
import shutil
import subprocess
import gc
import time
import re
from datetime import datetime, timedelta
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog,
    QTreeWidget, QTreeWidgetItem, QLabel, QMessageBox, QHBoxLayout,
    QAbstractItemView, QSpinBox, QFormLayout, QProgressBar,
    QSizePolicy, QComboBox
)
try:
    from PySide6.QtWidgets import QShortcut
except Exception:
    from PySide6.QtGui import QShortcut

from PySide6.QtGui import QPixmap, QImage, QDragEnterEvent, QDropEvent, QKeySequence, QFont
from PySide6.QtCore import Qt, QThread, Signal, QEvent, QTimer
from PIL import Image

# ---------------- i18n ----------------
LANG_FA = 'fa'
LANG_EN = 'en'

STRINGS = {
    'en': {
        'title': "Saino COC",
        'preview': "Preview",
        'load_folder': "📁 Load Folder",
        'load_zip': "🗜 Load ZIP(s)",
        'move_up': "⬆ Move Up",
        'move_down': "⬇ Move Down",
        'convert': "📄 Convert to PDF",
        'clear': "🧹 Clear List",
        'remove': "✖ Remove Selected",
        'res_scale': "Resolution Scale (%):",
        'jpeg_quality': "JPEG Quality (%):",
        'no_images': "No images selected.",
        'zip_error': "Failed to open ZIP:",
        'created': "PDF created:",
        'skipped': "Some images were skipped:",
        'canceled': "Conversion canceled by user.",
        'open_file': "Open File",
        'open_folder': "Open Folder",
        'close': "Close",
        'open_options_title': "Conversion finished",
        'no_valid': "No valid images found to convert.",
        'processing': "Processing images...",
        'separate_or_combine': "Multiple groups detected. Create separate PDFs per group?",
        'yes': "Separate",
        'no': "Combine into one",
        'dropped_images': "Dropped Images",
        'lang_toggle': "FA/EN",
        'sort_label': "Sort:",
        'sort_default': "Default",
        'sort_name': "By Name",
        'sort_number': "By Number",
        'tool_window_hint': "Mini mode",
    },
    'fa': {
        'title': "Saino COC",
        'preview': "پیش‌نمایش",
        'load_folder': "📁 بارگذاری پوشه",
        'load_zip': "🗜 بارگذاری ZIP(ها)",
        'move_up': "⬆ بالا بردن",
        'move_down': "⬇ پایین بردن",
        'convert': "📄 تبدیل به PDF",
        'clear': "🧹 پاک‌کردن لیست",
        'remove': "✖ حذف انتخاب‌شده",
        'res_scale': "مقیاس رزولوشن (%):",
        'jpeg_quality': "کیفیت JPEG (%):",
        'no_images': "تصویری انتخاب نشده.",
        'zip_error': "باز کردن ZIP با خطا مواجه شد:",
        'created': "PDF ساخته شد:",
        'skipped': "بعضی تصاویر رد شدند:",
        'canceled': "کاربر عملیات را کنسل کرد.",
        'open_file': "باز کردن فایل",
        'open_folder': "باز کردن پوشه",
        'close': "بستن",
        'open_options_title': "فرآیند پایان یافت",
        'no_valid': "هیچ تصویر معتبری برای تبدیل یافت نشد.",
        'processing': "در حال پردازش تصاویر...",
        'separate_or_combine': "چند گروه شناسایی شد. برای هر گروه PDF جدا بسازم؟",
        'yes': "جدا",
        'no': "ترکیب کن",
        'dropped_images': "تصاویر کشیده شده",
        'lang_toggle': "FA/EN",
        'sort_label': "مرتب‌سازی:",
        'sort_default': "پیش‌فرض",
        'sort_name': "بر اساس نام",
        'sort_number': "بر اساس عدد",
        'tool_window_hint': "حالت مینی",
    }
}

# ---------------- utilities ----------------

IMAGE_EXTS = ('.png', '.jpg', '.jpeg', '.tif', '.tiff', '.bmp')


def natural_key(s: str):
    parts = re.split(r"(\d+)", s)
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def ensure_dir(path):
    try:
        os.makedirs(path, exist_ok=True)
    except Exception:
        pass


# ---------------- background cleanup worker ----------------
class TempCleanupWorker(QThread):
    """Scan system temp for old temp dirs matching our prefixes and remove them.
    Runs in background so startup is not blocked.
    """
    status = Signal(str)

    def __init__(self, prefixes=None, older_than_hours=24):
        super().__init__()
        self._stop = False
        self.prefixes = prefixes or ['saino_temp_root_', 'saino_zip_', 'saino_combined_', 'saino_proc_']
        self.older_than = timedelta(hours=older_than_hours)

    def run(self):
        tmp = tempfile.gettempdir()
        now = datetime.now()
        try:
            for name in os.listdir(tmp):
                if self._stop:
                    break
                for p in self.prefixes:
                    if name.startswith(p):
                        path = os.path.join(tmp, name)
                        try:
                            mtime = datetime.fromtimestamp(os.path.getmtime(path))
                        except Exception:
                            mtime = now
                        if now - mtime > self.older_than:
                            try:
                                if os.path.isdir(path):
                                    shutil.rmtree(path, ignore_errors=True)
                                else:
                                    try:
                                        os.remove(path)
                                    except Exception:
                                        pass
                            except Exception:
                                pass
                        break
        except Exception:
            pass

    def stop(self):
        self._stop = True


# ---------------- Conversion Worker ----------------
class ConversionWorker(QThread):
    progress = Signal(int)         # overall processed images count
    finished_signal = Signal(list, list)  # (created_pdfs, skipped_list)
    error = Signal(str)

    def __init__(self, groups_ordered, output_dir, scale, jpeg_quality, clean_after_group=True):
        super().__init__()
        self.groups = groups_ordered
        self.output_dir = output_dir
        self.scale = float(scale)
        self.jpeg_quality = int(jpeg_quality)
        self._is_canceled = False
        self.clean_after_group = clean_after_group

    def run(self):
        try:
            created = []
            skipped_all = []
            processed_count = 0
            total_images = sum(len(g['paths']) for g in self.groups)
            for group in self.groups:
                if self._is_canceled:
                    break
                name = group['name']
                paths = group['paths']
                if name == '__COMBINED__':
                    base = f"combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                else:
                    base = os.path.splitext(name)[0]
                out_pdf = os.path.join(self.output_dir, f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
                skipped = self._process_and_save(paths, out_pdf, processed_count, total_images)
                skipped_all.extend(skipped)
                if os.path.exists(out_pdf):
                    created.append(out_pdf)
                processed_count += max(1, len(paths))
                tdir = group.get('tempdir')
                if self.clean_after_group and tdir:
                    try:
                        shutil.rmtree(tdir, ignore_errors=True)
                    except Exception:
                        pass
                if self._is_canceled:
                    break
            self.finished_signal.emit(created, skipped_all)
        except Exception as e:
            self.error.emit(str(e))

    def cancel(self):
        self._is_canceled = True

    def _process_and_save(self, paths, out_pdf, processed_offset, total_images):
        processed_tmp = []
        skipped = []
        local_count = 0
        for idx, p in enumerate(paths):
            if self._is_canceled:
                break
            local_count += 1
            try:
                img = Image.open(p)
            except Exception as e:
                skipped.append((p, str(e)))
                processed_offset += 1
                self.progress.emit(processed_offset)
                continue
            try:
                if img.mode != 'RGB':
                    img = img.convert('RGB')
            except Exception:
                skipped.append((p, 'convert to RGB failed'))
                try:
                    img.close()
                except Exception:
                    pass
                processed_offset += 1
                self.progress.emit(processed_offset)
                continue
            try:
                if self.scale < 1.0:
                    w, h = img.size
                    img = img.resize((max(1, int(w * self.scale)), max(1, int(h * self.scale))), Image.LANCZOS)
                fd, tmp_path = tempfile.mkstemp(prefix='saino_proc_', suffix='.jpg')
                os.close(fd)
                img.save(tmp_path, format='JPEG', quality=self.jpeg_quality, optimize=True)
                processed_tmp.append(tmp_path)
            except Exception as e:
                skipped.append((p, f'processing failed: {e}'))
            finally:
                try:
                    img.close()
                except Exception:
                    pass
            gc.collect()
            processed_offset += 1
            self.progress.emit(processed_offset)

        pil_list = []
        for f in processed_tmp:
            try:
                im = Image.open(f)
                if im.mode != 'RGB':
                    im = im.convert('RGB')
                pil_list.append(im)
            except Exception as e:
                skipped.append((f, f'open processed failed: {e}'))

        if not pil_list:
            for f in processed_tmp:
                try:
                    os.remove(f)
                except Exception:
                    pass
            return skipped

        first, *others = pil_list
        try:
            first.save(out_pdf, 'PDF', save_all=True, append_images=others, optimize=True)
        except Exception as e:
            skipped.append((out_pdf, f'save failed: {e}'))
        finally:
            for im in pil_list:
                try:
                    im.close()
                except Exception:
                    pass
            for f in processed_tmp:
                try:
                    os.remove(f)
                except Exception:
                    pass
            gc.collect()
        return skipped


# ---------------- Main Window ----------------
class ImageToPDF(QWidget):
    def __init__(self):
        super().__init__()
        self.lang = LANG_FA
        self.t = lambda k: STRINGS[self.lang].get(k, k)

        self.setWindowTitle(self.t('title'))
        # Use normal window flags (so it appears in taskbar)
        self.setWindowFlags(Qt.Window | Qt.WindowCloseButtonHint | Qt.WindowMinimizeButtonHint)
        # prevent maximize
        self.setWindowFlag(Qt.WindowMaximizeButtonHint, False)
        self.setMinimumSize(980, 560)
        self.setMaximumWidth(1400)

        # Tree (left)
        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setSelectionMode(QAbstractItemView.SingleSelection)
        # enable internal drag & drop move
        self.tree.setDragEnabled(True)
        self.tree.setAcceptDrops(True)
        self.tree.setDropIndicatorShown(True)
        self.tree.setDragDropMode(QAbstractItemView.InternalMove)
        self.tree.setDefaultDropAction(Qt.MoveAction)

        # Preview (right)
        self.preview_label = QLabel(self.t('preview'))
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setFixedHeight(420)
        self.preview_label.setStyleSheet("background:#0f0f0f; border:1px solid #2b2b2b; color:#ddd; padding:8px;")
        preview_font = QFont(); preview_font.setPointSize(11); self.preview_label.setFont(preview_font)

        # Buttons & controls
        # Top buttons above preview
        self.load_folder_btn = QPushButton(self.t('load_folder'))
        self.load_folder_btn.clicked.connect(self.load_folder)
        self.load_zip_btn = QPushButton(self.t('load_zip'))
        self.load_zip_btn.clicked.connect(self.load_zip)

        # Bottom buttons below preview
        self.up_btn = QPushButton(self.t('move_up')); self.up_btn.clicked.connect(self.move_up)
        self.down_btn = QPushButton(self.t('move_down')); self.down_btn.clicked.connect(self.move_down)
        self.remove_btn = QPushButton(self.t('remove')); self.remove_btn.clicked.connect(self.remove_selected)
        self.clear_btn = QPushButton(self.t('clear')); self.clear_btn.clicked.connect(self.clear_all)
        self.convert_btn = QPushButton(self.t('convert')); self.convert_btn.clicked.connect(self.convert_to_pdf)

        # Language toggle button
        self.lang_btn = QPushButton(self.t('lang_toggle'))
        self.lang_btn.clicked.connect(self.toggle_language)

        # Sort combobox
        self.sort_combo = QComboBox()
        self.sort_combo.addItems([self.t('sort_default'), self.t('sort_name'), self.t('sort_number')])
        self.sort_combo.currentIndexChanged.connect(self.apply_sorting)

        # Controls: resolution scale and jpeg quality
        self.scale_spin = QSpinBox(); self.scale_spin.setRange(10, 100); self.scale_spin.setValue(100); self.scale_spin.setSuffix('%')
        self.jpeg_spin = QSpinBox(); self.jpeg_spin.setRange(10, 100); self.jpeg_spin.setValue(95); self.jpeg_spin.setSuffix('%')
        form = QFormLayout()
        form.addRow(self.t('res_scale'), self.scale_spin)
        form.addRow(self.t('jpeg_quality'), self.jpeg_spin)

        # Layout assembly
        top_btn_layout = QHBoxLayout()
        top_btn_layout.addWidget(self.load_folder_btn)
        top_btn_layout.addWidget(self.load_zip_btn)
        top_btn_layout.addStretch()
        top_btn_layout.addWidget(self.lang_btn)
        top_btn_layout.addWidget(QLabel(self.t('sort_label')))
        top_btn_layout.addWidget(self.sort_combo)

        bottom_btn_layout = QHBoxLayout()
        for b in [self.up_btn, self.down_btn, self.remove_btn, self.clear_btn, self.convert_btn]:
            bottom_btn_layout.addWidget(b)
        bottom_btn_layout.addStretch()

        right_layout = QVBoxLayout()
        right_layout.addLayout(top_btn_layout)
        right_layout.addWidget(self.preview_label)
        right_layout.addLayout(bottom_btn_layout)
        right_layout.addLayout(form)

        left_layout = QVBoxLayout()
        left_layout.addWidget(self.tree)

        main_layout = QHBoxLayout()
        main_layout.addLayout(left_layout, 36)
        main_layout.addLayout(right_layout, 64)
        self.setLayout(main_layout)

        # internal data
        self.base_temp_root = tempfile.mkdtemp(prefix='saino_temp_root_')
        self.group_tempdirs = {}            # group_key -> tempdir (or None)
        self.source_map = {}                # child_path -> group_key
        self.loaded_zip_order = []          # list of group keys in insertion order
        self.group_order = {}               # group_key -> list of child paths (in insertion order)

        # Drag/drop
        self.setAcceptDrops(True)
        self.tree.itemDoubleClicked.connect(self.on_item_double_click)

        # shortcut
        try:
            self.lang_shortcut = QShortcut(QKeySequence('Ctrl+L'), self)
            self.lang_shortcut.activated.connect(self.toggle_language)
        except Exception:
            pass

        # progress
        self.progress_bar = QProgressBar(self)
        self.progress_bar.setVisible(False)
        left_layout.addWidget(self.progress_bar)

        self.worker = None

        # start background cleanup (non-blocking) to remove old orphan temp dirs
        self._start_background_cleanup()

    # prevent maximizing/fullscreen attempts
    def changeEvent(self, ev):
        if ev.type() == QEvent.WindowStateChange:
            st = self.windowState()
            if st & (Qt.WindowFullScreen | Qt.WindowMaximized):
                self.setWindowState(Qt.WindowNoState)
        super().changeEvent(ev)

    def _start_background_cleanup(self):
        QTimer.singleShot(50, self._launch_cleanup_thread)

    def _launch_cleanup_thread(self):
        self.cleanup_worker = TempCleanupWorker(older_than_hours=24)
        self.cleanup_worker.start()

    # drag/drop override to resync mappings after InternalMove
    def dropEvent(self, e: QDropEvent):
        super().dropEvent(e)
        QTimer.singleShot(20, self._rebuild_all_mappings)

    def _rebuild_all_mappings(self):
        new_group_order = {}
        new_loaded_zip_order = []
        new_source_map = {}
        for i in range(self.tree.topLevelItemCount()):
            root = self.tree.topLevelItem(i)
            d = root.data(0, Qt.UserRole)
            if not d or d.get('type') != 'group':
                continue
            key = d.get('key')
            new_loaded_zip_order.append(key)
            lst = []
            for j in range(root.childCount()):
                cp = root.child(j).data(0, Qt.UserRole).get('path')
                lst.append(cp)
                new_source_map[cp] = key
            new_group_order[key] = lst
        self.group_order = new_group_order
        self.loaded_zip_order = new_loaded_zip_order
        self.source_map = new_source_map

    # ------------ loading groups ------------
    def load_folder(self):
        folder = QFileDialog.getExistingDirectory(self, self.t('load_folder'))
        if not folder: return
        basename = os.path.basename(folder)
        self._ensure_group(basename, basename)
        entries = [f for f in os.listdir(folder) if f.lower().endswith(IMAGE_EXTS)]
        entries.sort(key=natural_key)
        for ent in entries:
            full = os.path.join(folder, ent)
            self._add_child(basename, full)
        self.group_tempdirs[basename] = None

    def load_zip(self):
        files, _ = QFileDialog.getOpenFileNames(self, self.t('load_zip'), '', 'ZIP Files (*.zip)')
        if not files: return
        files.sort(key=lambda p: natural_key(os.path.basename(p)))
        for p in files:
            self._add_zip_group(p)

    def _add_zip_group(self, zip_path, clear_first=False):
        basename = os.path.basename(zip_path)
        # create a unique tempdir for this zip inside base_temp_root
        safe_name = re.sub(r'[^A-Za-z0-9_.-]', '_', basename)
        group_temp = tempfile.mkdtemp(prefix=f'saino_zip_{safe_name}_', dir=self.base_temp_root)
        try:
            with zipfile.ZipFile(zip_path, 'r') as z:
                # collect image entries
                names = [n for n in z.namelist() if n.lower().endswith(IMAGE_EXTS)]
                names.sort(key=natural_key)
                self._ensure_group(basename, basename)
                for name in names:
                    try:
                        # extract preserving internal path but under group_temp
                        extracted = z.extract(name, group_temp)
                    except Exception:
                        continue
                    # normalize path
                    extracted = os.path.normpath(extracted)
                    # if extracted is a directory (entry had trailing slash), scan inside
                    if os.path.isdir(extracted):
                        for root, _, files in os.walk(extracted):
                            for f in files:
                                if f.lower().endswith(IMAGE_EXTS):
                                    self._add_child(basename, os.path.join(root, f))
                    else:
                        if os.path.splitext(extracted)[1].lower() in IMAGE_EXTS:
                            self._add_child(basename, extracted)
                self.group_tempdirs[basename] = group_temp
                if basename not in self.loaded_zip_order:
                    self.loaded_zip_order.append(basename)
        except Exception as e:
            try:
                shutil.rmtree(group_temp, ignore_errors=True)
            except Exception:
                pass
            QMessageBox.warning(self, self.t('title'), f"{self.t('zip_error')} {e}")

    # ------------ tree helpers ------------
    def _ensure_group(self, key, display_name):
        root = self._find_group_item(key)
        if root is None:
            root = QTreeWidgetItem(self.tree)
            root.setText(0, display_name)
            root.setData(0, Qt.UserRole, {'type': 'group', 'key': key})
            root.setExpanded(True)
            root.setFirstColumnSpanned(True)
            font = root.font(0); font.setBold(True); root.setFont(0, font)
            # allow drag/drop on group nodes
            root.setFlags(root.flags() | Qt.ItemIsDragEnabled | Qt.ItemIsDropEnabled)
            self.group_order[key] = []
            if key not in self.loaded_zip_order:
                self.loaded_zip_order.append(key)
        return root

    def _find_group_item(self, key):
        for i in range(self.tree.topLevelItemCount()):
            it = self.tree.topLevelItem(i)
            d = it.data(0, Qt.UserRole)
            if d and d.get('type') == 'group' and d.get('key') == key:
                return it
        return None

    def _add_child(self, group_key, child_path):
        root = self._ensure_group(group_key, group_key)
        child_path = os.path.normpath(child_path)
        # avoid duplicates
        for i in range(root.childCount()):
            if root.child(i).data(0, Qt.UserRole).get('path') == child_path:
                return
        child = QTreeWidgetItem(root)
        child.setText(0, os.path.basename(child_path))
        child.setData(0, Qt.UserRole, {'type': 'image', 'path': child_path})
        # make draggable
        child.setFlags(child.flags() | Qt.ItemIsDragEnabled | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
        root.addChild(child)
        self.source_map[child_path] = group_key
        # maintain insertion order
        if group_key not in self.group_order:
            self.group_order[group_key] = []
        self.group_order[group_key].append(child_path)

    # ------------ UI behavior ------------
    def on_item_double_click(self, item, col):
        d = item.data(0, Qt.UserRole)
        if not d: return
        if d.get('type') == 'image':
            path = d.get('path')
            self._show_preview(path)

    def _show_preview(self, path):
        img = QImage(path)
        if img.isNull():
            self.preview_label.setText(self.t('preview'))
            self.preview_label.setPixmap(QPixmap())
            return
        scaled = img.scaled(self.preview_label.width(), self.preview_label.height(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        self.preview_label.setPixmap(QPixmap.fromImage(scaled))

    # ------------ move up/down (group or child) ------------
    def move_up(self):
        it = self.tree.currentItem()
        if it is None: return
        d = it.data(0, Qt.UserRole)
        if not d: return
        if d.get('type') == 'group':
            idx = self.tree.indexOfTopLevelItem(it)
            if idx > 0:
                self.tree.takeTopLevelItem(idx)
                self.tree.insertTopLevelItem(idx - 1, it)
                self._rebuild_all_mappings()
        elif d.get('type') == 'image':
            parent = it.parent()
            idx = parent.indexOfChild(it)
            if idx > 0:
                parent.removeChild(it)
                parent.insertChild(idx - 1, it)
                self._rebuild_all_mappings()

    def move_down(self):
        it = self.tree.currentItem()
        if it is None: return
        d = it.data(0, Qt.UserRole)
        if not d: return
        if d.get('type') == 'group':
            idx = self.tree.indexOfTopLevelItem(it)
            if idx < self.tree.topLevelItemCount() - 1:
                self.tree.takeTopLevelItem(idx)
                self.tree.insertTopLevelItem(idx + 1, it)
                self._rebuild_all_mappings()
        elif d.get('type') == 'image':
            parent = it.parent()
            idx = parent.indexOfChild(it)
            if idx < parent.childCount() - 1:
                parent.removeChild(it)
                parent.insertChild(idx + 1, it)
                self._rebuild_all_mappings()

    def remove_selected(self):
        it = self.tree.currentItem()
        if it is None: return
        d = it.data(0, Qt.UserRole)
        if not d: return
        if d.get('type') == 'group':
            idx = self.tree.indexOfTopLevelItem(it)
            key = d.get('key')
            tdir = self.group_tempdirs.get(key)
            if tdir:
                try: shutil.rmtree(tdir, ignore_errors=True)
                except Exception: pass
                self.group_tempdirs.pop(key, None)
            self.tree.takeTopLevelItem(idx)
            for i in range(it.childCount()):
                child = it.child(i)
                cp = child.data(0, Qt.UserRole).get('path')
                try: del self.source_map[cp]
                except Exception: pass
            try: del self.group_order[key]
            except Exception: pass
            try: self.loaded_zip_order.remove(key)
            except Exception: pass
        else:
            parent = it.parent()
            cp = it.data(0, Qt.UserRole).get('path')
            parent.removeChild(it)
            try: del self.source_map[cp]
            except Exception: pass
            self._rebuild_all_mappings()

    def clear_all(self):
        for t in list(self.group_tempdirs.values()):
            if t:
                try: shutil.rmtree(t, ignore_errors=True)
                except Exception: pass
        self.group_tempdirs.clear()
        self.tree.clear()
        self.source_map.clear()
        self.loaded_zip_order = []
        self.group_order.clear()
        self.preview_label.setText(self.t('preview'))
        self.preview_label.setPixmap(QPixmap())

    # ------------ sorting ------------
    def apply_sorting(self):
        idx = self.sort_combo.currentIndex()
        mode = 'default' if idx == 0 else ('name' if idx == 1 else 'number')
        for i in range(self.tree.topLevelItemCount()):
            root = self.tree.topLevelItem(i)
            d = root.data(0, Qt.UserRole)
            if not d or d.get('type') != 'group':
                continue
            key = d.get('key')
            children = [root.child(j).data(0, Qt.UserRole).get('path') for j in range(root.childCount())]
            if mode == 'default':
                order = self.group_order.get(key, children)
            elif mode == 'name':
                order = sorted(children, key=lambda p: os.path.basename(p).lower())
            else:
                order = sorted(children, key=lambda p: natural_key(os.path.basename(p)))
            mapping = {root.child(j).data(0, Qt.UserRole).get('path'): root.child(j) for j in range(root.childCount())}
            while root.childCount():
                root.takeChild(0)
            for p in order:
                if p in mapping:
                    root.addChild(mapping[p])
            if mode == 'default':
                self.group_order[key] = order

    # ------------ conversion UI flow ------------
    def convert_to_pdf(self):
        # gather groups in tree order
        groups = []
        for i in range(self.tree.topLevelItemCount()):
            root = self.tree.topLevelItem(i)
            d = root.data(0, Qt.UserRole)
            if not d or d.get('type') != 'group':
                continue
            group_name = d.get('key')
            paths = []
            for j in range(root.childCount()):
                child = root.child(j)
                pd = child.data(0, Qt.UserRole)
                if pd and pd.get('type') == 'image':
                    paths.append(pd.get('path'))
            groups.append((group_name, paths))

        if not groups:
            QMessageBox.warning(self, self.t('title'), self.t('no_images'))
            return

        multiple_groups = len(groups) > 1
        combine_mode = False
        if multiple_groups:
            msg = QMessageBox(self)
            msg.setWindowTitle(self.t('title'))
            msg.setText(self.t('separate_or_combine'))
            btn_sep = msg.addButton(self.t('yes'), QMessageBox.YesRole)
            btn_comb = msg.addButton(self.t('no'), QMessageBox.NoRole)
            msg.exec()
            clicked = msg.clickedButton()
            if clicked == btn_sep:
                combine_mode = False
            else:
                combine_mode = True

        groups_ordered = []
        if multiple_groups and not combine_mode:
            for gname, paths in groups:
                groups_ordered.append({
                    'name': gname,
                    'paths': paths,
                    'tempdir': self.group_tempdirs.get(gname),
                    'is_combined_group': False
                })
        else:
            # build combined temp and copy files with zero-padded sequence prefixes
            combined_temp = tempfile.mkdtemp(prefix='saino_combined_', dir=self.base_temp_root)
            all_paths = []
            total_images = 0
            for _, paths in groups:
                total_images += len(paths)
            pad = max(6, len(str(total_images)))
            seq = 1
            for i in range(self.tree.topLevelItemCount()):
                root = self.tree.topLevelItem(i)
                d = root.data(0, Qt.UserRole)
                if not d or d.get('type') != 'group':
                    continue
                for j in range(root.childCount()):
                    child = root.child(j)
                    pd = child.data(0, Qt.UserRole)
                    if pd and pd.get('type') == 'image':
                        src = pd.get('path')
                        ext = os.path.splitext(src)[1].lower()
                        newname = f"{seq:0{pad}d}_{os.path.basename(src)}"
                        dst = os.path.join(combined_temp, newname)
                        try:
                            shutil.copy2(src, dst)
                            all_paths.append(dst)
                        except Exception:
                            all_paths.append(src)
                        seq += 1
            groups_ordered.append({
                'name': '__COMBINED__',
                'paths': all_paths,
                'tempdir': combined_temp,
                'is_combined_group': True
            })

        out_dir = os.path.join(os.getcwd(), 'output_pdfs'); os.makedirs(out_dir, exist_ok=True)
        total_images = sum(len(g['paths']) for g in groups_ordered)
        if total_images <= 0:
            QMessageBox.warning(self, self.t('title'), self.t('no_valid'))
            for g in groups_ordered:
                tdir = g.get('tempdir')
                if tdir and g.get('is_combined_group'):
                    try: shutil.rmtree(tdir, ignore_errors=True)
                    except Exception: pass
            return

        scale = self.scale_spin.value() / 100.0
        jpeg_q = self.jpeg_spin.value()

        self.progress_bar.setVisible(True)
        self.progress_bar.setMaximum(total_images)
        self.progress_bar.setValue(0)
        self.convert_btn.setEnabled(False)

        self.worker = ConversionWorker(groups_ordered=groups_ordered, output_dir=out_dir, scale=scale, jpeg_quality=jpeg_q, clean_after_group=True)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_progress(self, v):
        try:
            self.progress_bar.setValue(min(self.progress_bar.maximum(), v))
        except Exception:
            pass

    def _on_error(self, m):
        QMessageBox.critical(self, self.t('title'), m)
        self._cleanup_after_worker()

    def _on_finished(self, created_pdfs, skipped_all):
        self._cleanup_after_worker()
        if not created_pdfs:
            QMessageBox.warning(self, self.t('title'), self.t('no_valid'))
            return

        msg = QMessageBox(self)
        msg.setWindowTitle(self.t('open_options_title'))
        if len(created_pdfs) == 1:
            msg.setText(f"{self.t('created')}\n{created_pdfs[0]}")
            btn_open = msg.addButton(self.t('open_file'), QMessageBox.AcceptRole)
        else:
            msg.setText(f"{len(created_pdfs)} PDFs created.")
            btn_open = None
        btn_open_folder = msg.addButton(self.t('open_folder'), QMessageBox.AcceptRole)
        btn_close = msg.addButton(self.t('close'), QMessageBox.RejectRole)
        msg.exec()
        clicked = msg.clickedButton()
        if clicked == btn_open and btn_open is not None:
            try:
                if sys.platform.startswith('win'):
                    os.startfile(created_pdfs[0])
                elif sys.platform == 'darwin':
                    subprocess.Popen(['open', created_pdfs[0]])
                else:
                    subprocess.Popen(['xdg-open', created_pdfs[0]])
            except Exception:
                pass
        elif clicked == btn_open_folder:
            out_dir = os.path.dirname(created_pdfs[0]) if created_pdfs else os.path.join(os.getcwd(), 'output_pdfs')
            try:
                if sys.platform.startswith('win'):
                    os.startfile(out_dir)
                elif sys.platform == 'darwin':
                    subprocess.Popen(['open', out_dir])
                else:
                    subprocess.Popen(['xdg-open', out_dir])
            except Exception:
                pass

        if skipped_all:
            warn = self.t('skipped') + '\n' + '\n'.join([f"{os.path.basename(p)}: {r}" for p, r in skipped_all[:10]])
            if len(skipped_all) > 10:
                warn += f"\n...and {len(skipped_all)-10} more"
            QMessageBox.warning(self, self.t('title'), warn)

    def _cleanup_after_worker(self):
        try:
            if self.worker and self.worker.isRunning():
                self.worker.cancel(); self.worker.wait(1000)
        except Exception:
            pass
        self.worker = None
        self.convert_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)
        gc.collect()

    def toggle_language(self):
        self.lang = LANG_EN if self.lang == LANG_FA else LANG_FA
        self.t = lambda k: STRINGS[self.lang].get(k, k)
        self.setWindowTitle(self.t('title'))
        self.load_folder_btn.setText(self.t('load_folder'))
        self.load_zip_btn.setText(self.t('load_zip'))
        self.up_btn.setText(self.t('move_up'))
        self.down_btn.setText(self.t('move_down'))
        self.remove_btn.setText(self.t('remove'))
        self.clear_btn.setText(self.t('clear'))
        self.convert_btn.setText(self.t('convert'))
        self.lang_btn.setText(self.t('lang_toggle'))
        # rebuild sort combo entries
        self.sort_combo.blockSignals(True)
        self.sort_combo.clear()
        self.sort_combo.addItems([self.t('sort_default'), self.t('sort_name'), self.t('sort_number')])
        self.sort_combo.blockSignals(False)
        # update preview
        self.preview_label.setText(self.t('preview'))
        # update dropped group name if exists
        dropped_item = self._find_group_item('__DROPPED__')
        if dropped_item:
            dropped_item.setText(0, self.t('dropped_images'))

    def closeEvent(self, ev):
        # Cancel conversion worker and wait
        try:
            if self.worker and self.worker.isRunning():
                self.worker.cancel()
                self.worker.wait(1000)
        except Exception:
            pass
        # Stop cleanup worker and wait
        try:
            if hasattr(self, 'cleanup_worker') and self.cleanup_worker.isRunning():
                self.cleanup_worker.stop()
                self.cleanup_worker.wait(1000)
        except Exception:
            pass
        # Remove temp dirs
        try:
            if os.path.exists(self.base_temp_root):
                shutil.rmtree(self.base_temp_root, ignore_errors=True)
            for t in list(self.group_tempdirs.values()):
                if t:
                    shutil.rmtree(t, ignore_errors=True)
        except Exception:
            pass
        super().closeEvent(ev)
        # Ensure QApplication exits when window closed
        QApplication.quit()


# ---------- run ----------
if __name__ == '__main__':
    # Note: keep console disabled at compile-time (e.g. with Nuitka flag --windows-console-mode=disable)
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(True)

    w = ImageToPDF()
    w.show()
    sys.exit(app.exec())
