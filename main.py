#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Saino COC — Final
Features in this final build:
- Tree view (QTreeWidget): top-level nodes = groups (ZIP or folder), children = images. Clear which image belongs to which ZIP.
- Natural numeric sorting for ZIPs and images inside them.
- Ability to reorder groups (select group root + Move Up/Move Down) and reorder images within a group (select child + Move Up/Move Down).
- Supports multi-ZIP selection; ZIPs are sorted numerically before extracting.
- On Convert: if multiple ZIP groups detected, a bilingual dialog asks whether to create separate PDFs per group or combine into a single PDF. Buttons and dialogs are fully bilingual (EN/FA). Toggle language with Ctrl+L.
- Conversion runs in a background QThread to avoid UI freeze.
- Memory optimization: scaled JPEGs are written to temp, then assembled into PDFs; temp cleaned up and GC run.
- Estimated final size shown before conversion.

Usage:
- Load Folder: adds a group (folder name) with its images (numeric-sorted).
- Load ZIP(s): select multiple ZIP files; they will be natural-sorted (CH1, CH2, CH10) and each becomes a group.
- Drag & drop ZIP(s) or folders or individual images — multiple ZIPs dropped together are sorted automatically.
- Select a group (click group name) to move the whole group up/down using the same Move Up/Move Down buttons. Select an image (child) to move it within its group.
- Double-click an image to preview it.
- Press Ctrl+L to toggle English / Persian UI.

Save as: saino_coc_final.py
Requires: PySide6, Pillow
"""

import sys
import os
import zipfile
import tempfile
import shutil
import subprocess
import gc
from datetime import datetime
import re

from PySide6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QFileDialog,
    QTreeWidget, QTreeWidgetItem, QLabel, QMessageBox, QHBoxLayout,
    QAbstractItemView, QSpinBox, QFormLayout, QGroupBox, QProgressBar
)
# QShortcut location differs across PySide6 versions
try:
    from PySide6.QtWidgets import QShortcut
except Exception:
    from PySide6.QtGui import QShortcut

from PySide6.QtGui import QPixmap, QImage, QDragEnterEvent, QDropEvent, QKeySequence
from PySide6.QtCore import Qt, QThread, Signal
from PIL import Image

# ---------------- i18n ----------------
LANG_EN = 'en'
LANG_FA = 'fa'

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
        'image_scale': "Image Scale:",
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
        'yes': "Yes",
        'no': "No",
        'estimate_title': "Estimated size",
        'estimate_proceed': "Estimated total: {size}\nProceed?",
        'combined_name_info': "Combined filename will be: {name}.pdf",
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
        'image_scale': "مقیاس تصویر:",
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
        'yes': "بله",
        'no': "خیر",
        'estimate_title': "حجم تقریبی",
        'estimate_proceed': "حجم تقریبی: {size}\nادامه می‌دهی؟",
        'combined_name_info': "نام فایل ترکیبی خواهد بود: {name}.pdf",
    }
}

# ---------------- utilities ----------------

def natural_key(s: str):
    parts = re.split(r"(\d+)", s)
    return [int(p) if p.isdigit() else p.lower() for p in parts]


def human_size(nbytes):
    if nbytes < 1024:
        return f"{nbytes} B"
    for unit in ["KB", "MB", "GB", "TB"]:
        nbytes /= 1024.0
        if abs(nbytes) < 1024.0:
            return f"{nbytes:.2f} {unit}"
    return f"{nbytes:.2f} PB"


def estimate_pdf_size(paths, scale):
    total = 0
    for p in paths:
        try:
            total += os.path.getsize(p)
        except Exception:
            pass
    compression_factor = 0.6
    return int(total * (scale ** 2) * compression_factor)


def longest_common_substring(strings):
    if not strings:
        return ""
    s0 = min(strings, key=len)
    best = ""
    L = len(s0)
    for i in range(L):
        for j in range(i + 1, L + 1):
            sub = s0[i:j]
            if len(sub) <= len(best):
                continue
            if all(sub in s for s in strings):
                best = sub
    return best.strip(" _-.")

# ---------------- Worker ----------------
class ConversionWorker(QThread):
    progress = Signal(int)  # per-image index (1-based), 0 = new group
    finished_signal = Signal(list, list)
    error = Signal(str)

    def __init__(self, groups_ordered, output_dir, scale, temp_dir):
        super().__init__()
        self.groups = groups_ordered
        self.output_dir = output_dir
        self.scale = scale
        self.temp_dir = temp_dir
        self._is_canceled = False

    def run(self):
        try:
            created = []
            skipped_all = []
            for group_name, paths in self.groups:
                if self._is_canceled:
                    break
                self.progress.emit(0)
                if not paths:
                    continue
                if group_name == '__COMBINED__':
                    base = f"combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
                else:
                    base = os.path.splitext(group_name)[0]
                out_pdf = os.path.join(self.output_dir, f"{base}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
                skipped = self._process_and_save(paths, out_pdf)
                skipped_all.extend(skipped)
                if os.path.exists(out_pdf):
                    created.append(out_pdf)
                if self._is_canceled:
                    break
            self.finished_signal.emit(created, skipped_all)
        except Exception as e:
            self.error.emit(str(e))

    def cancel(self):
        self._is_canceled = True

    def _process_and_save(self, paths, out_pdf):
        processed_tmp = []
        skipped = []
        for idx, p in enumerate(paths):
            if self._is_canceled:
                break
            self.progress.emit(idx + 1)
            try:
                img = Image.open(p)
            except Exception as e:
                skipped.append((p, str(e)))
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
                continue
            try:
                if self.scale < 1.0:
                    w, h = img.size
                    img = img.resize((max(1, int(w * self.scale)), max(1, int(h * self.scale))), Image.LANCZOS)
                fd, tmp_path = tempfile.mkstemp(prefix='saino_proc_', suffix='.jpg', dir=self.temp_dir)
                os.close(fd)
                img.save(tmp_path, format='JPEG', quality=85, optimize=True)
                processed_tmp.append(tmp_path)
            except Exception as e:
                skipped.append((p, f'processing failed: {e}'))
            finally:
                try:
                    img.close()
                except Exception:
                    pass
            gc.collect()

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
        env_lang = os.environ.get('LANG', '')
        self.lang = LANG_FA if 'fa' in env_lang.lower() or 'fa_IR' in env_lang else LANG_EN
        self.t = lambda k: STRINGS[self.lang].get(k, k)

        self.setWindowTitle(self.t('title'))
        self.setMinimumSize(1000, 650)

        self.tree = QTreeWidget()
        self.tree.setHeaderHidden(True)
        self.tree.setSelectionMode(QAbstractItemView.SingleSelection)

        self.preview_label = QLabel(self.t('preview'))
        self.preview_label.setAlignment(Qt.AlignCenter)
        self.preview_label.setFixedHeight(360)
        self.preview_label.setStyleSheet("background:#222; border:1px solid #444; color:#ddd;")

        # buttons
        load_folder_btn = QPushButton(self.t('load_folder'))
        load_folder_btn.clicked.connect(self.load_folder)
        load_zip_btn = QPushButton(self.t('load_zip'))
        load_zip_btn.clicked.connect(self.load_zip)
        up_btn = QPushButton(self.t('move_up'))
        up_btn.clicked.connect(self.move_up)
        down_btn = QPushButton(self.t('move_down'))
        down_btn.clicked.connect(self.move_down)
        convert_btn = QPushButton(self.t('convert'))
        convert_btn.clicked.connect(self.convert_to_pdf)
        self.convert_btn = convert_btn
        clear_btn = QPushButton(self.t('clear'))
        clear_btn.clicked.connect(self.clear_all)
        remove_btn = QPushButton(self.t('remove'))
        remove_btn.clicked.connect(self.remove_selected)

        self.quality_spin = QSpinBox(); self.quality_spin.setRange(10, 100); self.quality_spin.setValue(100); self.quality_spin.setSuffix('%')
        quality_layout = QFormLayout(); quality_layout.addRow(self.t('image_scale'), self.quality_spin)

        control_group = QGroupBox('Controls')
        control_layout = QVBoxLayout()
        for b in [load_folder_btn, load_zip_btn, up_btn, down_btn, remove_btn, clear_btn, convert_btn]:
            b.setStyleSheet('padding:8px; font-size:14px;')
            control_layout.addWidget(b)
        control_layout.addLayout(quality_layout)
        control_group.setLayout(control_layout)

        left_layout = QVBoxLayout(); left_layout.addWidget(control_group); left_layout.addWidget(self.tree)
        right_layout = QVBoxLayout(); right_layout.addWidget(self.preview_label)
        main_layout = QHBoxLayout(); main_layout.addLayout(left_layout, 40); main_layout.addLayout(right_layout, 60)
        self.setLayout(main_layout)

        # internal data
        self.temp_dir = tempfile.mkdtemp(prefix='saino_temp_')
        self.source_map = {}  # child_path -> group_basename
        self.loaded_zip_order = []  # ordered list of zip basenames when load_zip used

        # drag/drop
        self.setAcceptDrops(True)
        self.tree.itemDoubleClicked.connect(self.on_item_double_click)

        # keyboard lang toggle
        try:
            self.lang_shortcut = QShortcut(QKeySequence('Ctrl+L'), self)
            self.lang_shortcut.activated.connect(self.toggle_language)
        except Exception:
            pass

        # progress
        self.progress_bar = QProgressBar(self); self.progress_bar.setVisible(False); left_layout.addWidget(self.progress_bar)
        self.worker = None

    # ------------ drag/drop ------------
    def dragEnterEvent(self, e: QDragEnterEvent):
        if e.mimeData().hasUrls(): e.acceptProposedAction()
    def dropEvent(self, e: QDropEvent):
        urls = e.mimeData().urls()
        zip_paths = []
        for url in urls:
            path = url.toLocalFile()
            if os.path.isdir(path):
                self._add_folder_group(path)
            elif zipfile.is_zipfile(path):
                zip_paths.append(path)
            else:
                if path.lower().endswith(('.png', '.jpg', '.jpeg')):
                    # add to a special "Dropped images" group
                    self._ensure_group('__DROPPED__', 'Dropped Images')
                    self._add_child('__DROPPED__', path)
        if zip_paths:
            zip_paths.sort(key=lambda p: natural_key(os.path.basename(p)))
            for p in zip_paths:
                self._add_zip_group(p)
        e.acceptProposedAction()

    # ------------ loading groups ------------
    def load_folder(self):
        folder = QFileDialog.getExistingDirectory(self, 'Choose Folder')
        if not folder: return
        self._add_folder_group(folder, clear_first=True)

    def _add_folder_group(self, folder, clear_first=False):
        basename = os.path.basename(folder)
        if clear_first:
            self.clear_all()
            self.loaded_zip_order = []
        self._ensure_group(basename, basename)
        entries = [f for f in os.listdir(folder) if f.lower().endswith(('.png', '.jpg', '.jpeg'))]
        entries.sort(key=natural_key)
        for ent in entries:
            full = os.path.join(folder, ent)
            self._add_child(basename, full)

    def load_zip(self):
        files, _ = QFileDialog.getOpenFileNames(self, 'Choose ZIP File(s)', '', 'ZIP Files (*.zip)')
        if not files: return
        files.sort(key=lambda p: natural_key(os.path.basename(p)))
        self.clear_all()
        self.loaded_zip_order = [os.path.basename(p) for p in files]
        for p in files:
            self._add_zip_group(p)

    def _add_zip_group(self, zip_path, clear_first=False):
        basename = os.path.basename(zip_path)
        if clear_first:
            self.clear_all(); self.loaded_zip_order = [basename]
        try:
            with zipfile.ZipFile(zip_path, 'r') as z:
                names = [n for n in z.namelist() if n.lower().endswith(('.png', '.jpg', '.jpeg'))]
                names.sort(key=natural_key)
                self._ensure_group(basename, basename)
                for name in names:
                    extracted = z.extract(name, self.temp_dir)
                    self._add_child(basename, extracted)
        except Exception as e:
            QMessageBox.warning(self, self.t('title'), f"{self.t('zip_error')} {e}")

    # ------------ tree helpers ------------
    def _ensure_group(self, key, display_name):
        # if group with basename exists, return; else create top-level item
        root = self._find_group_item(key)
        if root is None:
            root = QTreeWidgetItem(self.tree)
            root.setText(0, display_name)
            root.setData(0, Qt.UserRole, {'type': 'group', 'key': key})
            root.setExpanded(True)
            # visual style
            root.setFirstColumnSpanned(True)
            font = root.font(0); font.setBold(True); root.setFont(0, font)
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
        # avoid duplicates
        for i in range(root.childCount()):
            if root.child(i).data(0, Qt.UserRole).get('path') == child_path:
                return
        child = QTreeWidgetItem(root)
        child.setText(0, os.path.basename(child_path))
        child.setData(0, Qt.UserRole, {'type': 'image', 'path': child_path})
        root.addChild(child)
        self.source_map[child_path] = group_key

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
        elif d.get('type') == 'image':
            parent = it.parent()
            idx = parent.indexOfChild(it)
            if idx > 0:
                parent.removeChild(it)
                parent.insertChild(idx - 1, it)

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
        elif d.get('type') == 'image':
            parent = it.parent()
            idx = parent.indexOfChild(it)
            if idx < parent.childCount() - 1:
                parent.removeChild(it)
                parent.insertChild(idx + 1, it)

    def remove_selected(self):
        it = self.tree.currentItem()
        if it is None: return
        d = it.data(0, Qt.UserRole)
        if not d: return
        if d.get('type') == 'group':
            # remove whole group
            idx = self.tree.indexOfTopLevelItem(it)
            self.tree.takeTopLevelItem(idx)
            # remove associated source_map entries
            # iterate children paths
            for i in range(it.childCount()):
                child = it.child(i)
                cp = child.data(0, Qt.UserRole).get('path')
                try: del self.source_map[cp]
                except: pass
        else:
            parent = it.parent()
            cp = it.data(0, Qt.UserRole).get('path')
            parent.removeChild(it)
            try: del self.source_map[cp]
            except: pass

    def clear_all(self):
        self.tree.clear(); self.source_map.clear(); self.loaded_zip_order = []
        self.preview_label.setText(self.t('preview'))

    # ------------ conversion UI flow ------------
    def convert_to_pdf(self):
        # gather groups in tree order
        groups = []  # list of (group_name, [paths])
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

        # detect multiple groups
        multiple_groups = len(groups) > 1
        combine_mode = False
        if multiple_groups:
            # bilingual dialog: Yes = separate, No = combine
            msg = QMessageBox(self)
            msg.setWindowTitle(self.t('title'))
            msg.setText(self.t('separate_or_combine'))
            btn_yes = msg.addButton(self.t('yes'), QMessageBox.YesRole)
            btn_no = msg.addButton(self.t('no'), QMessageBox.NoRole)
            msg.exec()
            clicked = msg.clickedButton()
            if clicked == btn_yes:
                combine_mode = False
            else:
                combine_mode = True

        groups_ordered = []
        if multiple_groups and not combine_mode:
            # keep groups in current tree order
            for gname, paths in groups:
                groups_ordered.append((gname, paths))
        else:
            # single combined group in tree order
            all_paths = []
            for _, paths in groups:
                all_paths.extend(paths)
            groups_ordered.append(('__COMBINED__', all_paths))

        # estimate
        scale = self.quality_spin.value() / 100.0
        total_est = sum(estimate_pdf_size(pths, scale) for _, pths in groups_ordered)

        # If combined and there are multiple group names, try to create meaningful filename
        combined_name = None
        if len(groups_ordered) == 1 and groups_ordered[0][0] == '__COMBINED__':
            group_keys = [g.data(0, Qt.UserRole).get('key') for g in [self.tree.topLevelItem(i) for i in range(self.tree.topLevelItemCount())]]
            lcs = longest_common_substring(group_keys)
            if lcs and len(lcs) >= 3:
                combined_name = lcs
            else:
                combined_name = f"combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        est_text = self.t('estimate_proceed').format(size=human_size(total_est))
        if combined_name:
            est_text += "\n" + self.t('combined_name_info').format(name=combined_name)

        reply = QMessageBox.question(self, self.t('estimate_title'), est_text, QMessageBox.Yes | QMessageBox.No)
        if reply != QMessageBox.Yes:
            return

        # prepare worker
        out_dir = os.path.join(os.getcwd(), 'output_pdfs'); os.makedirs(out_dir, exist_ok=True)
        self.progress_bar.setVisible(True)
        max_images = max((len(p) for _, p in groups_ordered), default=1)
        self.progress_bar.setMaximum(max_images)
        self.progress_bar.setValue(0)
        self.convert_btn.setEnabled(False)

        self.worker = ConversionWorker(groups_ordered=groups_ordered, output_dir=out_dir, scale=scale, temp_dir=self.temp_dir)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

        # store combined_name to rename later
        self._combined_name = combined_name

    def _on_progress(self, v):
        if v == 0:
            self.progress_bar.setValue(0)
        else:
            self.progress_bar.setValue(v)

    def _on_error(self, m):
        QMessageBox.critical(self, self.t('title'), m)
        self._cleanup_after_worker()

    def _on_finished(self, created_pdfs, skipped_all):
        self._cleanup_after_worker()
        if not created_pdfs:
            QMessageBox.warning(self, self.t('title'), self.t('no_valid'))
            return

        # if combined and single created name, rename to combined_name if available
        if len(created_pdfs) == 1 and getattr(self, '_combined_name', None):
            cur = created_pdfs[0]
            dirp = os.path.dirname(cur)
            new = os.path.join(dirp, f"{self._combined_name}.pdf")
            try:
                if os.path.exists(new):
                    new = os.path.join(dirp, f"{self._combined_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf")
                os.rename(cur, new)
                created_pdfs[0] = new
            except Exception:
                pass

        # dialog (bilingual)
        msg = QMessageBox(self); msg.setWindowTitle(self.t('open_options_title'))
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
                if sys.platform.startswith('win'): os.startfile(created_pdfs[0])
                elif sys.platform == 'darwin': subprocess.Popen(['open', created_pdfs[0]])
                else: subprocess.Popen(['xdg-open', created_pdfs[0]])
            except Exception:
                pass
        elif clicked == btn_open_folder:
            out_dir = os.path.dirname(created_pdfs[0]) if created_pdfs else os.path.join(os.getcwd(), 'output_pdfs')
            try:
                if sys.platform.startswith('win'): os.startfile(out_dir)
                elif sys.platform == 'darwin': subprocess.Popen(['open', out_dir])
                else: subprocess.Popen(['xdg-open', out_dir])
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
                self.worker.cancel(); self.worker.wait(100)
        except Exception:
            pass
        self.worker = None
        self.convert_btn.setEnabled(True)
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)
        gc.collect()

    def _on_error(self, m):
        QMessageBox.critical(self, self.t('title'), m)
        self._cleanup_after_worker()

    def toggle_language(self):
        self.lang = LANG_EN if self.lang == LANG_FA else LANG_FA
        self.t = lambda k: STRINGS[self.lang].get(k, k)
        self.setWindowTitle(self.t('title'))
        # update buttons and preview label
        for w in self.findChildren(QPushButton):
            txt = w.text()
            if 'Load' in txt or 'بارگذاری' in txt or '📁' in txt:
                w.setText(self.t('load_folder'))
            elif 'ZIP' in txt or 'ZIP' in txt:
                w.setText(self.t('load_zip'))
            elif 'Move Up' in txt or 'بالا' in txt or '⬆' in txt:
                w.setText(self.t('move_up'))
            elif 'Move Down' in txt or 'پایین' in txt or '⬇' in txt:
                w.setText(self.t('move_down'))
            elif 'Convert' in txt or 'تبدیل' in txt or '📄' in txt:
                w.setText(self.t('convert'))
            elif 'Clear' in txt or 'پاک' in txt or '🧹' in txt:
                w.setText(self.t('clear'))
            elif 'Remove' in txt or 'حذف' in txt or '✖' in txt:
                w.setText(self.t('remove'))
        self.preview_label.setText(self.t('preview'))

    def closeEvent(self, ev):
        try:
            if self.worker and self.worker.isRunning():
                self.worker.cancel(); self.worker.wait(200)
        except Exception:
            pass
        try:
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        except Exception:
            pass
        super().closeEvent(ev)

# ---------- run ----------
if __name__ == '__main__':
    app = QApplication(sys.argv)
    w = ImageToPDF()
    w.show()
    sys.exit(app.exec())
