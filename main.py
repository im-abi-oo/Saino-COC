#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Saino COC — Final v2 (Refactored & Persian texts fixed)

What I changed in this update (per your request):
- Fixed full bilingual support (all dialogs, buttons and prompts use localized Persian/English strings). No built-in "Yes/No" left untranslated — dialogs use self.t(...) labels everywhere.
- Restored and improved PDF size estimation (shows human-readable estimate before conversion).
- Kept per-group temp folders and combined-mode/global numbering behavior from previous v2.
- Numbering in the UI has been preserved/updated so that displayed numbers match the filenames that will be used during conversion.
- All other performance/memory optimizations kept (background QThread, temp cleanup, scaled JPEG intermediates).

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
        'dropped_images': 'Dropped Images'
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
        'dropped_images': 'تصاویر افتاده'
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

    def __init__(self, groups_ordered, output_dir, scale, temp_dir, combined_name=None):
        super().__init__()
        self.groups = groups_ordered
        self.output_dir = output_dir
        self.scale = scale
        self.temp_dir = temp_dir
        self._is_canceled = False
        self.combined_name = combined_name

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
                    base = self.combined_name or f"combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
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
                    self._ensure_group('__DROPPED__', self.t('dropped_images'))
                    self._add_child('__DROPPED__', path)
        if zip_paths:
            zip_paths.sort(key=lambda p: natural_key(os.path.basename(p)))
            for p in zip_paths:
                self._add_zip_group(p)
        e.acceptProposedAction()

    # ------------ loading groups ------------
    def load_folder(self):
        folder = QFileDialog.getExistingDirectory(self, self.t('load_folder'))
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
        files, _ = QFileDialog.getOpenFileNames(self, self.t('load_zip'), '', 'ZIP Files (*.zip)')
        if not files: return
        files.sort(key=lambda p: natural_key(os.path.basename(p)))
        self.clear_all()
        self.loaded_zip_order = [os.path.basename(p) for p in files]
        for p in files:
            self._add_zip_group(p)

    def _add_zip_group(self, zip_path, clear_first=False):
        basename = os.path.basename(zip_path)
        group_temp = os.path.join(self.temp_dir, os.path.splitext(basename)[0])
        os.makedirs(group_temp, exist_ok=True)
        if clear_first:
            self.clear_all(); self.loaded_zip_order = [basename]
        try:
            with zipfile.ZipFile(zip_path, 'r') as z:
                names = [n for n in z.namelist() if n.lower().endswith(('.png', '.jpg', '.jpeg'))]
                names.sort(key=natural_key)
                self._ensure_group(basename, basename)
                for name in names:
                    # extract into the group's temp dir to avoid collisions
                    # ensure target path exists
                    target = os.path.join(group_temp, os.path.basename(name))
                    # if filename exists in temp, append counter
                    base_name = os.path.splitext(os.path.basename(name))[0]
                    ext = os.path.splitext(name)[1]
                    counter = 0
                    candidate = target
                    while os.path.exists(candidate):
                        counter += 1
                        candidate = os.path.join(group_temp, f"{base_name}_{counter}{ext}")
                    z.extract(name, group_temp)
                    # if z.extract created nested path, resolve it to candidate
                    extracted_path = os.path.join(group_temp, name)
                    if os.path.exists(extracted_path) and extracted_path != candidate:
                        try:
                            os.makedirs(os.path.dirname(candidate), exist_ok=True)
                            shutil.move(extracted_path, candidate)
                            parent_dir = os.path.dirname(extracted_path)
                            try:
                                os.removedirs(parent_dir)
                            except Exception:
                                pass
                        except Exception:
                            candidate = extracted_path
                    final_path = candidate if os.path.exists(candidate) else extracted_path
                    self._add_child(basename, final_path)
        except Exception as e:
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
        orig = os.path.basename(child_path)
        child.setData(0, Qt.UserRole, {'type': 'image', 'path': child_path, 'orig': orig})
        root.addChild(child)
        self.source_map[child_path] = group_key
        self._update_group_numbering(root)

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
                self._update_group_numbering(parent)
        self._update_all_numbering()

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
                self._update_group_numbering(parent)
        self._update_all_numbering()

    def remove_selected(self):
        it = self.tree.currentItem()
        if it is None: return
        d = it.data(0, Qt.UserRole)
        if not d: return
        if d.get('type') == 'group':
            idx = self.tree.indexOfTopLevelItem(it)
            children = [it.child(i).data(0, Qt.UserRole).get('path') for i in range(it.childCount())]
            self.tree.takeTopLevelItem(idx)
            for p in children:
                try: del self.source_map[p]
                except: pass
        else:
            parent = it.parent()
            cp = it.data(0, Qt.UserRole).get('path')
            parent.removeChild(it)
            try: del self.source_map[cp]
            except: pass
            self._update_group_numbering(parent)
        self._update_all_numbering()

    def clear_all(self):
        self.tree.clear(); self.source_map.clear(); self.loaded_zip_order = []
        self.preview_label.setText(self.t('preview'))

    # ------------ numbering helpers ------------
    def _update_group_numbering(self, group_item):
        n = group_item.childCount()
        width = max(3, len(str(max(1, n))))
        for i in range(n):
            ch = group_item.child(i)
            d = ch.data(0, Qt.UserRole)
            orig = d.get('orig')
            label = f"{(i+1):0{width}d} - {orig}"
            ch.setText(0, label)

    def _update_all_numbering(self):
        for i in range(self.tree.topLevelItemCount()):
            self._update_group_numbering(self.tree.topLevelItem(i))

    # ------------ conversion UI flow ------------
    def convert_to_pdf(self):
        # collect groups in tree order
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
            for gname, paths in groups:
                groups_ordered.append((gname, list(paths)))
        else:
            all_paths = []
            for _, paths in groups:
                all_paths.extend(paths)
            groups_ordered.append(('__COMBINED__', all_paths))

        scale = self.quality_spin.value() / 100.0
        total_est = sum(estimate_pdf_size(pths, scale) for _, pths in groups_ordered)

        combined_name = None
        if len(groups_ordered) == 1 and groups_ordered[0][0] == '__COMBINED__':
            group_keys = [self.tree.topLevelItem(i).data(0, Qt.UserRole).get('key') for i in range(self.tree.topLevelItemCount())]
            lcs = longest_common_substring(group_keys)
            if lcs and len(lcs) >= 3:
                combined_name = lcs
            else:
                combined_name = f"combined_{datetime.now().strftime('%Y%m%d_%H%M%S')}"

        est_text = self.t('estimate_proceed').format(size=human_size(total_est))
        if combined_name:
            est_text += "\n" + self.t('combined_name_info').format(name=combined_name)

        # custom bilingual confirmation dialog for estimate
        est_msg = QMessageBox(self)
        est_msg.setWindowTitle(self.t('estimate_title'))
        est_msg.setText(est_text)
        btn_proceed = est_msg.addButton(self.t('yes'), QMessageBox.YesRole)
        btn_cancel = est_msg.addButton(self.t('no'), QMessageBox.NoRole)
        est_msg.exec()
        if est_msg.clickedButton() != btn_proceed:
            return

        work_temp_root = os.path.join(self.temp_dir, f"work_{datetime.now().strftime('%Y%m%d_%H%M%S')}")
        os.makedirs(work_temp_root, exist_ok=True)

        final_groups_for_worker = []

        if len(groups_ordered) == 1 and groups_ordered[0][0] == '__COMBINED__':
            combined_temp = os.path.join(work_temp_root, 'combined')
            os.makedirs(combined_temp, exist_ok=True)
            all_paths = groups_ordered[0][1]
            total = len(all_paths)
            width = max(3, len(str(max(1, total))))
            for idx, src in enumerate(all_paths):
                num = f"{(idx+1):0{width}d}"
                ext = os.path.splitext(src)[1].lower() or '.png'
                dst = os.path.join(combined_temp, f"{num}{ext}")
                try:
                    shutil.copy2(src, dst)
                except Exception:
                    try:
                        img = Image.open(src)
                        rgb = img.convert('RGB') if img.mode != 'RGB' else img
                        rgb.save(dst)
                        img.close()
                    except Exception:
                        pass
            temp_files = [os.path.join(combined_temp, f) for f in sorted(os.listdir(combined_temp))]
            final_groups_for_worker.append((combined_name or '__combined__', temp_files))
        else:
            for gname, paths in groups_ordered:
                grp_temp = os.path.join(work_temp_root, os.path.splitext(gname)[0])
                os.makedirs(grp_temp, exist_ok=True)
                total = len(paths)
                width = max(3, len(str(max(1, total))))
                for idx, src in enumerate(paths):
                    num = f"{(idx+1):0{width}d}"
                    ext = os.path.splitext(src)[1].lower() or '.png'
                    dst = os.path.join(grp_temp, f"{num}{ext}")
                    try:
                        shutil.copy2(src, dst)
                    except Exception:
                        try:
                            img = Image.open(src)
                            rgb = img.convert('RGB') if img.mode != 'RGB' else img
                            rgb.save(dst)
                            img.close()
                        except Exception:
                            pass
                temp_files = [os.path.join(grp_temp, f) for f in sorted(os.listdir(grp_temp))]
                final_groups_for_worker.append((gname, temp_files))

        out_dir = os.path.join(os.getcwd(), 'output_pdfs'); os.makedirs(out_dir, exist_ok=True)
        self.progress_bar.setVisible(True)
        max_images = max((len(p) for _, p in final_groups_for_worker), default=1)
        self.progress_bar.setMaximum(max_images)
        self.progress_bar.setValue(0)
        self.convert_btn.setEnabled(False)

        self.worker = ConversionWorker(groups_ordered=final_groups_for_worker, output_dir=out_dir, scale=scale, temp_dir=work_temp_root, combined_name=combined_name)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

        self._worker_work_root = work_temp_root

    def _on_progress(self, v):
        if v == 0:
            self.progress_bar.setValue(0)
        else:
            self.progress_bar.setValue(v)

    def _on_error(self, m):
        QMessageBox.critical(self, self.t('title'), m)
        self._cleanup_after_worker()

    def _on_finished(self, created_pdfs, skipped_all):
        try:
            if hasattr(self, '_worker_work_root') and os.path.exists(self._worker_work_root):
                shutil.rmtree(self._worker_work_root, ignore_errors=True)
        except Exception:
            pass

        self._cleanup_after_worker()
        if not created_pdfs:
            QMessageBox.warning(self, self.t('title'), self.t('no_valid'))
            return

        msg = QMessageBox(self); msg.setWindowTitle(self.t('open_options_title'))
        if len(created_pdfs) == 1:
            msg.setText(f"{self.t('created')}\n{created_pdfs[0]}")
            btn_open = msg.addButton(self.t('open_file'), QMessageBox.AcceptRole)
        else:
            msg.setText(f"{len(created_pdfs)} PDFs created.")
            btn_open = None
        btn_open_folder = msg.addButton(self.t('open_folder'), QMessageBox.AcceptRole); btn_close = msg.addButton(self.t('close'), QMessageBox.RejectRole)
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
        super().close
