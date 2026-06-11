"""
DOCX文件合并工具 - 将多个docx/doc文件合并为一个docx文件
支持文件拖放、文件夹导入、图片压缩，合并后文件不超过10MB
"""

import sys
import os
import io
import copy
import zipfile
import tempfile
import shutil
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QLabel, QFileDialog, QMessageBox,
    QProgressBar, QGroupBox, QCheckBox, QSpinBox, QListWidgetItem,
    QAbstractItemView, QStatusBar, QFrame, QSplitter, QComboBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QMimeData, QUrl
from PyQt5.QtGui import QFont, QIcon, QDragEnterEvent, QDropEvent, QPalette, QColor

from docx import Document
from docx.shared import Inches, Pt, Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from PIL import Image


MAX_SIZE_MB = 10
MAX_SIZE_BYTES = MAX_SIZE_MB * 1024 * 1024

SUPPORTED_EXTENSIONS = {'.docx', '.doc'}


def compress_image_in_memory(image_bytes, quality=60, max_dimension=1200):
    """压缩图片字节数据，返回压缩后的字节"""
    try:
        img = Image.open(io.BytesIO(image_bytes))

        if img.mode in ('RGBA', 'P'):
            img = img.convert('RGB')

        w, h = img.size
        if w > max_dimension or h > max_dimension:
            ratio = min(max_dimension / w, max_dimension / h)
            new_size = (int(w * ratio), int(h * ratio))
            img = img.resize(new_size, Image.LANCZOS)

        output = io.BytesIO()
        img.save(output, format='JPEG', quality=quality, optimize=True)
        compressed = output.getvalue()

        if len(compressed) < len(image_bytes):
            return compressed
        return image_bytes
    except Exception:
        return image_bytes


def compress_docx_images(docx_path, output_path, quality=60, max_dimension=1200):
    """压缩docx文件中的所有图片"""
    temp_dir = tempfile.mkdtemp()
    try:
        with zipfile.ZipFile(docx_path, 'r') as zin:
            zin.extractall(temp_dir)

        media_dir = os.path.join(temp_dir, 'word', 'media')
        if os.path.exists(media_dir):
            for fname in os.listdir(media_dir):
                fpath = os.path.join(media_dir, fname)
                ext = os.path.splitext(fname)[1].lower()
                if ext in ('.png', '.jpg', '.jpeg', '.bmp', '.tiff'):
                    with open(fpath, 'rb') as f:
                        original = f.read()
                    compressed = compress_image_in_memory(
                        original, quality, max_dimension
                    )
                    with open(fpath, 'wb') as f:
                        f.write(compressed)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as zout:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zout.write(file_path, arcname)
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def add_page_break(doc):
    """在文档末尾添加分页符"""
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    run._element.append(br)


def merge_documents(file_list, output_path, add_breaks=True,
                    compress=True, quality=60, max_dim=1200,
                    progress_callback=None):
    """
    合并多个docx文件为一个文件

    Args:
        file_list: docx文件路径列表
        output_path: 输出文件路径
        add_breaks: 是否在文件之间添加分页符
        compress: 是否压缩图片
        quality: 图片压缩质量 (1-100)
        max_dim: 图片最大尺寸
        progress_callback: 进度回调函数(current, total, message)
    """
    if not file_list:
        raise ValueError("没有可合并的文件")

    total = len(file_list)

    temp_files = []
    try:
        processed_files = []
        for i, fpath in enumerate(file_list):
            if progress_callback:
                progress_callback(i, total, f"处理中: {os.path.basename(fpath)}")

            if compress and fpath.lower().endswith('.docx'):
                temp_out = tempfile.mktemp(suffix='.docx')
                temp_files.append(temp_out)
                compress_docx_images(fpath, temp_out, quality, max_dim)
                processed_files.append(temp_out)
            else:
                processed_files.append(fpath)

        if progress_callback:
            progress_callback(total // 2, total, "正在合并文档...")

        base_doc = Document(processed_files[0])

        for i, fpath in enumerate(processed_files[1:], 1):
            if progress_callback:
                progress_callback(total // 2 + i, total,
                                  f"合并中: {os.path.basename(file_list[i])}")

            if add_breaks:
                add_page_break(base_doc)

            sub_doc = Document(fpath)
            _append_document(base_doc, sub_doc)

        if progress_callback:
            progress_callback(total - 1, total, "正在保存文件...")

        base_doc.save(output_path)

        if compress:
            temp_compressed = tempfile.mktemp(suffix='.docx')
            temp_files.append(temp_compressed)
            compress_docx_images(output_path, temp_compressed, quality, max_dim)
            shutil.move(temp_compressed, output_path)
            if temp_compressed in temp_files:
                temp_files.remove(temp_compressed)

        file_size = os.path.getsize(output_path)
        if progress_callback:
            progress_callback(total, total, "完成!")

        return file_size

    finally:
        for tf in temp_files:
            if os.path.exists(tf):
                os.remove(tf)


def _append_document(base_doc, sub_doc):
    """将sub_doc的内容追加到base_doc"""
    for element in sub_doc.element.body:
        tag = element.tag.split('}')[-1] if '}' in element.tag else element.tag
        if tag == 'sectPr':
            continue
        new_element = copy.deepcopy(element)
        base_doc.element.body.append(new_element)

    _copy_images(base_doc, sub_doc)


def _copy_images(base_doc, sub_doc):
    """将sub_doc中的图片资源复制到base_doc"""
    for rel in sub_doc.part.rels.values():
        if "image" in rel.reltype:
            try:
                image_part = rel.target_part
                new_rel = base_doc.part.relate_to(image_part, rel.reltype)

                for paragraph in base_doc.paragraphs:
                    for run in paragraph.runs:
                        for drawing in run._element.findall(
                            './/' + qn('a:blip')
                        ):
                            embed = drawing.get(qn('r:embed'))
                            if embed == rel.rId:
                                drawing.set(qn('r:embed'), new_rel)
            except Exception:
                pass


class MergeWorker(QThread):
    """后台合并线程"""
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(str, int)  # output_path, file_size
    error = pyqtSignal(str)

    def __init__(self, file_list, output_path, add_breaks, compress, quality, max_dim):
        super().__init__()
        self.file_list = file_list
        self.output_path = output_path
        self.add_breaks = add_breaks
        self.compress = compress
        self.quality = quality
        self.max_dim = max_dim

    def run(self):
        try:
            size = merge_documents(
                self.file_list,
                self.output_path,
                self.add_breaks,
                self.compress,
                self.quality,
                self.max_dim,
                progress_callback=self._on_progress
            )
            self.finished.emit(self.output_path, size)
        except Exception as e:
            self.error.emit(str(e))

    def _on_progress(self, current, total, message):
        self.progress.emit(current, total, message)


class FileListWidget(QListWidget):
    """支持拖放的文件列表"""
    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QAbstractItemView.DragDropMode.InternalMove)
        self.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.setAlternatingRowColors(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            paths = []
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if os.path.isfile(path):
                    ext = os.path.splitext(path)[1].lower()
                    if ext in SUPPORTED_EXTENSIONS:
                        paths.append(path)
                elif os.path.isdir(path):
                    for root, dirs, files in os.walk(path):
                        for f in sorted(files):
                            ext = os.path.splitext(f)[1].lower()
                            if ext in SUPPORTED_EXTENSIONS:
                                paths.append(os.path.join(root, f))
            if paths:
                self.files_dropped.emit(paths)
            event.acceptProposedAction()
        else:
            super().dropEvent(event)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.file_paths = []
        self.worker = None
        self._init_ui()

    def _init_ui(self):
        self.setWindowTitle("DOCX文件合并工具")
        self.setMinimumSize(750, 580)
        self.resize(800, 620)

        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setSpacing(10)
        main_layout.setContentsMargins(16, 16, 16, 16)

        # ── 标题 ──
        title = QLabel("📄 DOCX 文件合并工具")
        title.setFont(QFont("Microsoft YaHei", 18, QFont.Weight.Bold))
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("color: #2c3e50; margin-bottom: 4px;")
        main_layout.addWidget(title)

        subtitle = QLabel("将多个Word文档合并为一个，支持拖放文件和文件夹")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)
        subtitle.setStyleSheet("color: #7f8c8d; font-size: 12px; margin-bottom: 8px;")
        main_layout.addWidget(subtitle)

        # ── 文件列表区域 ──
        file_group = QGroupBox("待合并文件列表（支持拖放文件/文件夹到此处）")
        file_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 13px;
                border: 2px solid #bdc3c7;
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 16px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }
        """)
        file_layout = QVBoxLayout(file_group)

        self.file_list = FileListWidget()
        self.file_list.files_dropped.connect(self._on_files_dropped)
        self.file_list.setMinimumHeight(180)
        self.file_list.setStyleSheet("""
            QListWidget {
                border: 1px solid #ddd;
                border-radius: 4px;
                background: #fafafa;
                font-size: 12px;
                padding: 4px;
            }
            QListWidget::item {
                padding: 6px 8px;
                border-bottom: 1px solid #eee;
            }
            QListWidget::item:selected {
                background: #3498db;
                color: white;
            }
            QListWidget::item:alternate {
                background: #f5f5f5;
            }
        """)
        file_layout.addWidget(self.file_list)

        btn_layout = QHBoxLayout()
        btn_style = """
            QPushButton {
                background: #3498db;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 18px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover { background: #2980b9; }
            QPushButton:pressed { background: #2471a3; }
            QPushButton:disabled { background: #bdc3c7; }
        """
        danger_style = """
            QPushButton {
                background: #e74c3c;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 18px;
                font-size: 13px;
                font-weight: bold;
            }
            QPushButton:hover { background: #c0392b; }
            QPushButton:pressed { background: #a93226; }
            QPushButton:disabled { background: #bdc3c7; }
        """

        self.btn_add_files = QPushButton("➕ 添加文件")
        self.btn_add_files.setStyleSheet(btn_style)
        self.btn_add_files.clicked.connect(self._add_files)
        btn_layout.addWidget(self.btn_add_files)

        self.btn_add_folder = QPushButton("📁 添加文件夹")
        self.btn_add_folder.setStyleSheet(btn_style)
        self.btn_add_folder.clicked.connect(self._add_folder)
        btn_layout.addWidget(self.btn_add_folder)

        self.btn_move_up = QPushButton("⬆ 上移")
        self.btn_move_up.setStyleSheet(btn_style)
        self.btn_move_up.clicked.connect(self._move_up)
        btn_layout.addWidget(self.btn_move_up)

        self.btn_move_down = QPushButton("⬇ 下移")
        self.btn_move_down.setStyleSheet(btn_style)
        self.btn_move_down.clicked.connect(self._move_down)
        btn_layout.addWidget(self.btn_move_down)

        self.btn_remove = QPushButton("🗑 移除选中")
        self.btn_remove.setStyleSheet(danger_style)
        self.btn_remove.clicked.connect(self._remove_selected)
        btn_layout.addWidget(self.btn_remove)

        self.btn_clear = QPushButton("清空列表")
        self.btn_clear.setStyleSheet(danger_style)
        self.btn_clear.clicked.connect(self._clear_list)
        btn_layout.addWidget(self.btn_clear)

        file_layout.addLayout(btn_layout)
        main_layout.addWidget(file_group)

        # ── 设置区域 ──
        settings_group = QGroupBox("合并设置")
        settings_group.setStyleSheet("""
            QGroupBox {
                font-weight: bold;
                font-size: 13px;
                border: 2px solid #bdc3c7;
                border-radius: 8px;
                margin-top: 8px;
                padding-top: 16px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 12px;
                padding: 0 6px;
            }
        """)
        settings_layout = QHBoxLayout(settings_group)

        self.chk_page_break = QCheckBox("文件之间添加分页符")
        self.chk_page_break.setChecked(True)
        self.chk_page_break.setStyleSheet("font-size: 12px;")
        settings_layout.addWidget(self.chk_page_break)

        settings_layout.addSpacing(16)

        self.chk_compress = QCheckBox("压缩图片（减小文件体积）")
        self.chk_compress.setChecked(True)
        self.chk_compress.setStyleSheet("font-size: 12px;")
        self.chk_compress.stateChanged.connect(self._toggle_compress)
        settings_layout.addWidget(self.chk_compress)

        settings_layout.addSpacing(16)

        lbl_quality = QLabel("图片质量:")
        lbl_quality.setStyleSheet("font-size: 12px; font-weight: normal;")
        settings_layout.addWidget(lbl_quality)

        self.spin_quality = QSpinBox()
        self.spin_quality.setRange(10, 95)
        self.spin_quality.setValue(60)
        self.spin_quality.setSuffix("%")
        self.spin_quality.setStyleSheet("font-size: 12px;")
        settings_layout.addWidget(self.spin_quality)

        settings_layout.addSpacing(16)

        lbl_max = QLabel("图片最大尺寸:")
        lbl_max.setStyleSheet("font-size: 12px; font-weight: normal;")
        settings_layout.addWidget(lbl_max)

        self.spin_max_dim = QSpinBox()
        self.spin_max_dim.setRange(400, 3000)
        self.spin_max_dim.setValue(1200)
        self.spin_max_dim.setSuffix("px")
        self.spin_max_dim.setSingleStep(100)
        self.spin_max_dim.setStyleSheet("font-size: 12px;")
        settings_layout.addWidget(self.spin_max_dim)

        settings_layout.addStretch()
        main_layout.addWidget(settings_group)

        # ── 进度条 ──
        self.progress_bar = QProgressBar()
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                text-align: center;
                height: 24px;
                font-size: 12px;
            }
            QProgressBar::chunk {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #2ecc71, stop:1 #27ae60
                );
                border-radius: 5px;
            }
        """)
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        self.lbl_status = QLabel("")
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_status.setStyleSheet("color: #555; font-size: 12px;")
        main_layout.addWidget(self.lbl_status)

        # ── 合并按钮 ──
        self.btn_merge = QPushButton("🔗  开始合并")
        self.btn_merge.setFont(QFont("Microsoft YaHei", 14, QFont.Weight.Bold))
        self.btn_merge.setMinimumHeight(50)
        self.btn_merge.setStyleSheet("""
            QPushButton {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #2ecc71, stop:1 #27ae60
                );
                color: white;
                border: none;
                border-radius: 10px;
                font-size: 16px;
            }
            QPushButton:hover {
                background: qlineargradient(
                    x1:0, y1:0, x2:1, y2:0,
                    stop:0 #27ae60, stop:1 #1e8449
                );
            }
            QPushButton:pressed { background: #1e8449; }
            QPushButton:disabled { background: #bdc3c7; }
        """)
        self.btn_merge.clicked.connect(self._start_merge)
        main_layout.addWidget(self.btn_merge)

        # ── 状态栏 ──
        self.statusBar().showMessage("就绪 — 请添加要合并的文件")
        self.statusBar().setStyleSheet("font-size: 11px; color: #555;")

        self._update_count()

    def _toggle_compress(self, state):
        enabled = state == Qt.CheckState.Checked.value
        self.spin_quality.setEnabled(enabled)
        self.spin_max_dim.setEnabled(enabled)

    def _add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择Word文档", "",
            "Word文档 (*.docx *.doc);;所有文件 (*)"
        )
        if files:
            self._append_files(files)

    def _add_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            files = []
            for root, dirs, fnames in os.walk(folder):
                for f in sorted(fnames):
                    ext = os.path.splitext(f)[1].lower()
                    if ext in SUPPORTED_EXTENSIONS:
                        files.append(os.path.join(root, f))
            if files:
                self._append_files(files)
            else:
                QMessageBox.information(
                    self, "提示", "所选文件夹中未找到docx/doc文件"
                )

    def _on_files_dropped(self, paths):
        self._append_files(paths)

    def _append_files(self, paths):
        existing = set(self.file_paths)
        for p in paths:
            if p not in existing:
                self.file_paths.append(p)
                existing.add(p)
        self._refresh_list()

    def _refresh_list(self):
        self.file_list.clear()
        for i, p in enumerate(self.file_paths):
            fname = os.path.basename(p)
            size_kb = os.path.getsize(p) / 1024
            text = f"{i+1}. {fname}  ({size_kb:.0f} KB)"
            item = QListWidgetItem(text)
            item.setToolTip(p)
            self.file_list.addItem(item)
        self._update_count()

    def _update_count(self):
        n = len(self.file_paths)
        total_size = sum(os.path.getsize(p) for p in self.file_paths if os.path.exists(p))
        size_mb = total_size / (1024 * 1024)
        self.statusBar().showMessage(
            f"共 {n} 个文件 | 原始大小合计: {size_mb:.2f} MB | "
            f"限制: {MAX_SIZE_MB} MB"
        )

    def _move_up(self):
        rows = sorted(set(idx.row() for idx in self.file_list.selectedIndexes()))
        if not rows or rows[0] == 0:
            return
        for row in rows:
            self.file_paths[row - 1], self.file_paths[row] = (
                self.file_paths[row], self.file_paths[row - 1]
            )
        self._refresh_list()
        for row in rows:
            self.file_list.item(row - 1).setSelected(True)

    def _move_down(self):
        rows = sorted(
            set(idx.row() for idx in self.file_list.selectedIndexes()),
            reverse=True
        )
        if not rows or rows[0] >= len(self.file_paths) - 1:
            return
        for row in rows:
            self.file_paths[row + 1], self.file_paths[row] = (
                self.file_paths[row], self.file_paths[row + 1]
            )
        self._refresh_list()
        for row in rows:
            self.file_list.item(row + 1).setSelected(True)

    def _remove_selected(self):
        rows = sorted(
            set(idx.row() for idx in self.file_list.selectedIndexes()),
            reverse=True
        )
        for row in rows:
            if 0 <= row < len(self.file_paths):
                self.file_paths.pop(row)
        self._refresh_list()

    def _clear_list(self):
        if self.file_paths:
            reply = QMessageBox.question(
                self, "确认", "确定要清空文件列表吗？",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.file_paths.clear()
                self._refresh_list()

    def _start_merge(self):
        if not self.file_paths:
            QMessageBox.warning(self, "提示", "请先添加要合并的文件！")
            return

        doc_files = [f for f in self.file_paths if f.lower().endswith('.docx')]
        doc_legacy = [f for f in self.file_paths if f.lower().endswith('.doc')]

        if doc_legacy:
            QMessageBox.warning(
                self, "格式提示",
                f"检测到 {len(doc_legacy)} 个旧版 .doc 文件。\n"
                "本工具仅支持 .docx 格式，.doc 文件将被跳过。\n"
                "请先将 .doc 文件转换为 .docx 格式后再合并。"
            )
            if not doc_files:
                return
            self.file_paths = doc_files

        output_path, _ = QFileDialog.getSaveFileName(
            self, "保存合并后的文件", "merged_output.docx",
            "Word文档 (*.docx)"
        )
        if not output_path:
            return

        self._set_ui_busy(True)
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)

        self.worker = MergeWorker(
            list(self.file_paths),
            output_path,
            self.chk_page_break.isChecked(),
            self.chk_compress.isChecked(),
            self.spin_quality.value(),
            self.spin_max_dim.value()
        )
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_progress(self, current, total, message):
        if total > 0:
            pct = int(current / total * 100)
            self.progress_bar.setValue(pct)
        self.lbl_status.setText(message)

    def _on_finished(self, output_path, file_size):
        self._set_ui_busy(False)
        self.progress_bar.setValue(100)

        size_mb = file_size / (1024 * 1024)

        if file_size > MAX_SIZE_BYTES:
            QMessageBox.warning(
                self, "⚠️ 文件过大",
                f"合并后的文件大小为 {size_mb:.2f} MB，超过了 {MAX_SIZE_MB} MB 限制！\n\n"
                f"文件已保存到:\n{output_path}\n\n"
                "建议:\n"
                "1. 删减部分文件内容（尤其是图片较多的文件）\n"
                "2. 降低图片质量设置后重新合并\n"
                "3. 减少合并的文件数量\n"
                "4. 减小图片最大尺寸设置"
            )
        else:
            QMessageBox.information(
                self, "✅ 合并完成",
                f"合并成功！\n\n"
                f"文件大小: {size_mb:.2f} MB\n"
                f"保存位置: {output_path}"
            )

        self.lbl_status.setText(
            f"合并完成 — {size_mb:.2f} MB — {output_path}"
        )

    def _on_error(self, error_msg):
        self._set_ui_busy(False)
        self.progress_bar.setVisible(False)
        QMessageBox.critical(self, "错误", f"合并过程中出错:\n{error_msg}")
        self.lbl_status.setText("合并失败")

    def _set_ui_busy(self, busy):
        self.btn_merge.setEnabled(not busy)
        self.btn_add_files.setEnabled(not busy)
        self.btn_add_folder.setEnabled(not busy)
        self.btn_remove.setEnabled(not busy)
        self.btn_clear.setEnabled(not busy)
        if busy:
            self.btn_merge.setText("⏳ 合并中...")
        else:
            self.btn_merge.setText("🔗  开始合并")


def main():
    app = QApplication(sys.argv)

    app.setStyle("Fusion")
    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Window, QColor(245, 245, 245))
    palette.setColor(QPalette.ColorRole.WindowText, QColor(44, 62, 80))
    palette.setColor(QPalette.ColorRole.Base, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.AlternateBase, QColor(240, 240, 240))
    palette.setColor(QPalette.ColorRole.Button, QColor(52, 152, 219))
    palette.setColor(QPalette.ColorRole.ButtonText, QColor(255, 255, 255))
    palette.setColor(QPalette.ColorRole.Highlight, QColor(52, 152, 219))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor(255, 255, 255))
    app.setPalette(palette)

    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
