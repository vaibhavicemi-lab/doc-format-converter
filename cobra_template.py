import sys
import subprocess
import os
import re
import fitz  # PyMuPDF
import pandas as pd
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QEvent, QRect, QPoint, QSize, QMarginsF
from PyQt6.QtPrintSupport import QPrinter

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QTextEdit, QLineEdit,
    QFileDialog, QMessageBox, QInputDialog,
    QDialog, QMenu,
    QScrollArea, QSplitter,
    QProgressBar, QSizePolicy,
    QCheckBox, QGroupBox, QComboBox,
    QFrame, QRubberBand,
    QStyleFactory, QTableWidget, QTableWidgetItem
)
from PyQt6.QtGui import (
    QPixmap, QImage,
    QCursor, QColor, QPalette,
    QTextCursor,
    QStandardItemModel, QStandardItem,
    QAction, QIcon, QFont,
    QKeySequence
)


# ==========================================
# --- CHECKABLE COMBO BOX ---
# ==========================================
class CheckableComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.view().pressed.connect(self.handleItemPressed)
        self._model = QStandardItemModel(self)
        self.setModel(self._model)
        self._ignore_hide = False
        self.setModelColumn(0)
        self.setStyleSheet(
            "QComboBox QAbstractItemView { background-color: #ffffff; color: #000000; }"
        )

    def handleItemPressed(self, index):
        item = self._model.itemFromIndex(index)
        if item is None:
            return
        if item.checkState() == Qt.CheckState.Checked:
            item.setCheckState(Qt.CheckState.Unchecked)
        else:
            item.setCheckState(Qt.CheckState.Checked)
        self._ignore_hide = True
        self._model.dataChanged.emit(index, index)

    def addItem(self, text, checked=False):
        item = QStandardItem(text)
        item.setFlags(
            Qt.ItemFlag.ItemIsSelectable |
            Qt.ItemFlag.ItemIsEnabled |
            Qt.ItemFlag.ItemIsUserCheckable
        )
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        self._model.appendRow(item)

    def addItems(self, texts):
        for text in texts:
            self.addItem(text)

    def checkedItems(self):
        checked_items = []
        for i in range(self._model.rowCount()):
            item = self._model.item(i)
            if item.checkState() == Qt.CheckState.Checked and item.text() != "Enter Other...":
                checked_items.append(item.text())
        return checked_items

    def hidePopup(self):
        if self._ignore_hide:
            self._ignore_hide = False
            return
        super().hidePopup()

    def currentText(self):
        checked = self.checkedItems()
        return ", ".join(checked) if checked else ""

# ==========================================
# --- PURE PYTHON SUMMARIZER ---
# ==========================================
def simple_summarize(text, num_sentences=4):
    if not text:
        return ""
    stop_words = {"the", "is", "in", "and", "to", "of", "a", "for", "on", "with", "as", "by", "this", "that", "it", "are", "be", "or", "an", "at", "from", "which", "will"}
    sentences = re.split(r'(?<=[.!?]) +|\n+', text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 5]
    if len(sentences) <= num_sentences:
        return " ".join(sentences)

    words = re.findall(r'\b\w+\b', text.lower())
    freq = {}
    for w in words:
        if w not in stop_words and not w.isnumeric():
            freq[w] = freq.get(w, 0) + 1

    scores = {}
    for i, s in enumerate(sentences):
        score = 0
        s_words = re.findall(r'\b\w+\b', s.lower())
        for w in s_words:
            if w in freq:
                score += freq[w]
        scores[i] = score / max(len(s_words), 1)

    top_indices = sorted(sorted(scores, key=scores.get, reverse=True)[:num_sentences])
    return " ".join([sentences[i] for i in top_indices])

# ==========================================
# --- SMART TEXT EDITOR (SUPPORTS IMAGES)---
# ==========================================
class NotesEditor(QTextEdit):
    def canInsertFromMimeData(self, source):
        if source.hasImage() or source.hasText():
            return True
        return super().canInsertFromMimeData(source)

    def insertFromMimeData(self, source):
        if source.hasImage():
            clipboard = QApplication.clipboard()
            image = clipboard.image()
            if not image.isNull():
                cursor = self.textCursor()
                if image.width() > 500:
                    image = image.scaledToWidth(500, Qt.TransformationMode.SmoothTransformation)
                cursor.insertImage(image)
                self.insertPlainText("\n")
                return
        if source.hasText():
            self.insertPlainText(source.text())
            return
        super().insertFromMimeData(source)

# ==========================================
# --- PURE CODE ANALYSIS WORKER ---
# ==========================================
class AnalysisWorker(QThread):
    finished = pyqtSignal(dict)
    error = pyqtSignal(str)

    def __init__(self, text, topics, headings):
        super().__init__()
        self.text = text
        self.topics = topics
        self.headings = headings

    def run(self):
        try:
            results = {
                "project_name": "",
                "introduction": "",
                "found_topics": []
            }

            match = re.search(r'Project\s*Name[\s:]*(.+)', self.text, re.IGNORECASE)
            if match:
                clean_name = match.group(1).strip()
                if not clean_name.startswith("___"):
                    results["project_name"] = clean_name[:60]

            results["introduction"] = simple_summarize(self.text, num_sentences=4)

            # Use ONLY the headings passed in (from PDF detection), don't do additional text-based detection
            # This ensures we match topics based on actual PDF structure, not content
            combined_headings = self.headings if self.headings else []

            # Match topics only if they appear in the detected headings
            for topic in self.topics:
                tlow = topic.lower()
                for heading in combined_headings:
                    if not heading:
                        continue
                    hlow = heading.lower()
                    # exact word match or substring match within heading
                    if re.search(r'\b' + re.escape(tlow) + r'\b', hlow) or tlow in hlow or hlow in tlow:
                        results["found_topics"].append(topic)
                        break

            if results["introduction"] and "Introduction" not in results["found_topics"]:
                results["found_topics"].append("Introduction")

            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))

# ==========================================
# --- INTERACTIVE PDF SNIPPING TOOL ---
# ==========================================
class PdfPageLabel(QLabel):
    text_extracted = pyqtSignal(str)
    image_extracted = pyqtSignal(QPixmap)
    summary_requested = pyqtSignal(str) 

    def __init__(self, page, scroll_area, parent=None):
        super().__init__(parent)
        self.page = page  
        self.scroll_area = scroll_area 
        self.rubber_band = QRubberBand(QRubberBand.Shape.Rectangle, self)
        self.origin = QPoint()
        self.pan_start_pos = None

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.origin = event.pos()
            self.rubber_band.setGeometry(QRect(self.origin, QSize()))
            self.rubber_band.show()
        elif event.button() in [Qt.MouseButton.RightButton, Qt.MouseButton.MiddleButton]:
            self.pan_start_pos = event.globalPosition().toPoint()
            self.setCursor(Qt.CursorShape.ClosedHandCursor)

    def mouseMoveEvent(self, event):
        if not self.origin.isNull() and (event.buttons() & Qt.MouseButton.LeftButton):
            self.rubber_band.setGeometry(QRect(self.origin, event.pos()).normalized())
            
        elif self.pan_start_pos is not None and (event.buttons() & (Qt.MouseButton.RightButton | Qt.MouseButton.MiddleButton)):
            current_pos = event.globalPosition().toPoint()
            delta = current_pos - self.pan_start_pos
            
            h_bar = self.scroll_area.horizontalScrollBar()
            v_bar = self.scroll_area.verticalScrollBar()
            
            h_bar.setValue(h_bar.value() - delta.x())
            v_bar.setValue(v_bar.value() - delta.y())
            
            self.pan_start_pos = current_pos

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            rect = self.rubber_band.geometry()
            if rect.width() > 10 and rect.height() > 10:
                self.show_context_menu(event.pos(), rect)
            else:
                self.rubber_band.hide()
            self.origin = QPoint()
            
        elif event.button() in [Qt.MouseButton.RightButton, Qt.MouseButton.MiddleButton]:
            self.unsetCursor()
            self.pan_start_pos = None
            
    def get_text_from_rect(self, rect):
        if self.pixmap() and self.page:
            scale_x = self.page.rect.width / self.pixmap().width()
            scale_y = self.page.rect.height / self.pixmap().height()
            pdf_rect = fitz.Rect(rect.left() * scale_x, rect.top() * scale_y, rect.right() * scale_x, rect.bottom() * scale_y)
            pdf_rect = pdf_rect + (-3, -3, 3, 3)
            return self.page.get_textbox(pdf_rect).strip()
        return ""

    def show_context_menu(self, pos, rect):
        menu = QMenu(self)
        copy_img_action = menu.addAction("Extract as Image")
        copy_text_action = menu.addAction("Extract Text to Notes")
        menu.addSeparator()
        summarize_action = menu.addAction("Summarize Selection (Set as Intro)")

        action = menu.exec(self.mapToGlobal(pos))
        
        if action == copy_img_action:
            if self.pixmap(): self.image_extracted.emit(self.pixmap().copy(rect)) 
        elif action == copy_text_action:
            text = self.get_text_from_rect(rect)
            if text: self.text_extracted.emit(text) 
        elif action == summarize_action:
            text = self.get_text_from_rect(rect)
            if text: self.summary_requested.emit(text)
            
        self.rubber_band.hide()

# ==========================================
# --- TABLE EDITOR DIALOG ---
# ==========================================
class TableEditorDialog(QDialog):
    def __init__(self, table_name, headers, rows_data=None, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Edit Table: {table_name}")
        self.setGeometry(100, 100, 800, 500)
        self.table_name = table_name
        self.headers = headers
        self.rows_data = rows_data or [["" for _ in headers] for _ in range(3)]
        
        layout = QVBoxLayout(self)
        
        # Title
        title = QLabel(f"<b>Table: {table_name}</b>")
        layout.addWidget(title)
        
        # Table widget
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(len(headers))
        self.table_widget.setHorizontalHeaderLabels(headers)
        self.table_widget.setRowCount(len(self.rows_data))
        
        for row_idx, row_data in enumerate(self.rows_data):
            for col_idx, cell_data in enumerate(row_data):
                item = QTableWidgetItem(str(cell_data))
                self.table_widget.setItem(row_idx, col_idx, item)
        
        layout.addWidget(self.table_widget)
        
        # Button layout
        btn_layout = QHBoxLayout()
        
        add_row_btn = QPushButton("Add Row")
        add_row_btn.clicked.connect(self.add_row)
        btn_layout.addWidget(add_row_btn)
        
        delete_row_btn = QPushButton("Delete Row")
        delete_row_btn.clicked.connect(self.delete_row)
        btn_layout.addWidget(delete_row_btn)
        
        btn_layout.addStretch()
        
        save_btn = QPushButton("Save & Close")
        save_btn.setStyleSheet("background-color: #059669; color: white;")
        save_btn.clicked.connect(self.accept)
        btn_layout.addWidget(save_btn)
        
        cancel_btn = QPushButton("Cancel")
        cancel_btn.clicked.connect(self.reject)
        btn_layout.addWidget(cancel_btn)
        
        layout.addLayout(btn_layout)
    
    def add_row(self):
        row_idx = self.table_widget.rowCount()
        self.table_widget.insertRow(row_idx)
        for col_idx in range(len(self.headers)):
            item = QTableWidgetItem("")
            self.table_widget.setItem(row_idx, col_idx, item)
    
    def delete_row(self):
        current_row = self.table_widget.currentRow()
        if current_row >= 0:
            self.table_widget.removeRow(current_row)
    
    def get_table_data(self):
        data = []
        for row_idx in range(self.table_widget.rowCount()):
            row_data = []
            for col_idx in range(self.table_widget.columnCount()):
                item = self.table_widget.item(row_idx, col_idx)
                row_data.append(item.text() if item else "")
            data.append(row_data)
        return data

# ==========================================
# --- MULTI SELECT COMPONENT ---
# ==========================================
class MultiSelectWidget(QWidget):
    def __init__(self, options, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(5)
        
        self.display_box = QTextEdit()
        self.display_box.setReadOnly(True)
        self.display_box.setFixedHeight(65)
        self.display_box.setPlaceholderText("Selected options will appear here...")
        self.display_box.setStyleSheet("""
            background-color: white;
            color: #0f172a;
            border: 1px solid #cbd5e1;
            border-radius: 8px;
            padding: 8px;
        """)        
        self.combo = CheckableComboBox()
        clean_opts = [o for o in options if o != "Enter Other..."]
        self.combo.addItems(clean_opts)
        
        input_layout = QHBoxLayout()
        self.custom_input = QLineEdit()
        self.custom_input.setPlaceholderText("Or type other option here...")
        self.add_btn = QPushButton("Add")
        self.add_btn.setObjectName("ToolBtn")
        self.add_btn.setFixedWidth(60)
        self.add_btn.clicked.connect(self.add_custom_option)
        self.custom_input.returnPressed.connect(self.add_custom_option)
        
        input_layout.addWidget(self.custom_input)
        input_layout.addWidget(self.add_btn)
        
        layout.addWidget(self.combo)
        layout.addWidget(self.display_box)
        layout.addLayout(input_layout)
        
        self.combo._model.dataChanged.connect(self.update_display)
        
    def add_custom_option(self):
        text = self.custom_input.text().strip()
        if text:
            existing = [self.combo._model.item(i).text() for i in range(self.combo._model.rowCount())]
            if text not in existing:
                self.combo.addItem(text, checked=True)
            else:
                for i in range(self.combo._model.rowCount()):
                    if self.combo._model.item(i).text() == text:
                        self.combo._model.item(i).setCheckState(Qt.CheckState.Checked)
            self.custom_input.clear()
            self.update_display()
            
    def update_display(self):
        self.display_box.setText("\\n".join(self.combo.checkedItems()))
        
    def checkedItems(self):
        return self.combo.checkedItems()


# ==========================================
# --- MAIN APPLICATION ---
# ==========================================
class OfflineApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Offline Document Assistant - Fast Analyzer")
        self.resize(1450, 950) 
        
        self.file_path = ""
        self.doc = None
        self.page_data = [] 
        self.current_zoom = 600 
        self.is_maximized = False
        self.original_file_path = ""
        
        self.intro_text_data = "" 
        self.section_notes = {}
        self.manual_collapsed = False
        self.manual_section_maximized = False

        # Table data storage: table_name -> list of rows (each row is list of cells)
        self.table_data = {
            "5. Stakeholders": [["", "", "", ""]]*3,
            "7.3 Certification Task Allocation": [["", "", "", ""]]*3,
            "Annexure-1: Work Assignment List": [["", "", "", "", "", ""]]*3,
            "Annexure-1: Integration Checks": [["", "", "", "", ""]]*2,
            "Annexure-2: Contact Details": [["", "", "", "", ""]]*2,
        }
        
        # Image storage: heading -> list of QPixmap objects
        self.section_images = {}

        self.apply_stylesheet()
        self.init_ui()

    def apply_stylesheet(self):
        self.setStyleSheet("""
            QWidget { font-family: 'Segoe UI', sans-serif; font-size: 11pt; }
            QMainWindow, #mainWidget, QDialog, QMessageBox { background-color: #f8fafc; color: #0f172a; }
            #leftContainer { background: #ffffff; border-right: 1px solid #e2e8f0; }
            #rightContainer { background: #f1f5f9; }
            QFrame[class="Card"] { background: white; border-radius: 12px; border: 1px solid #e2e8f0; color: #0f172a; }
            QPushButton { background: #4f46e5; color: white; border: none; padding: 12px; border-radius: 8px; font-weight: bold; }
            QPushButton:hover { background: #4338ca; }
            QPushButton#ActionBtn { height: 40px; font-weight: bold; background-color: #f1f5f9; color: #334155; padding: 0 15px; border: 1px solid #cbd5e1; }
            QPushButton#ActionBtn:hover { background-color: #e2e8f0; }
            QPushButton#ToolBtn { background: white; color: #334155; border: 1px solid #cbd5e1; font-size: 16px; font-weight: bold; border-radius: 4px; padding: 5px; }
            QPushButton#ToolBtn:hover { background: #f1f5f9; }
            QPushButton#IconBtn { font-size: 24px; border-radius: 25px; background-color: #f8f9fa; border: 2px solid #ced4da; color: black; padding: 0; }
            QPushButton#IconBtn:hover { background-color: #e2e6ea; }
            QPushButton#IconBtnActive { font-size: 24px; border-radius: 25px; background-color: #d4edda; border: 2px solid #28a745; color: black; padding: 0; }
            QGroupBox { background: transparent; border: none; color: #0f172a; }
            QComboBox { padding: 8px; border-radius: 6px; border: 1px solid #cbd5e1; background-color: #ffffff; color: #000000; }
            QComboBox::drop-down { border: 0px; }
            QComboBox QAbstractItemView { border: 1px solid #cbd5e1; background-color: #ffffff; color: #000000; selection-background-color: #e2e8f0; selection-color: #000000; outline: 0; }
            QComboBox QAbstractItemView::item { background-color: #ffffff; color: #000000; padding: 8px; }
            QComboBox QAbstractItemView::item:selected { background-color: #e2e8f0; color: #000000; }
            QSplitter::handle { background-color: #94a3b8; border-radius: 2px; margin: 2px; }
            QLabel { color: #0f172a; background: transparent; }
            QLabel[class="Heading"] { font-size: 18px; font-weight: bold; color: #4f46e5; margin: 0; background: transparent; }
            QLineEdit, QTextEdit { background-color: #ffffff; color: #000000; border: 1px solid #cbd5e1; border-radius: 6px; padding: 8px; }
            QCheckBox { color: #0f172a; background: transparent; }
            QProgressBar { border: 1px solid #cbd5e1; background: #e2e8f0; border-radius: 5px; text-align: center; color: #0f172a; }
            QProgressBar::chunk { background-color: #4f46e5; }
            QScrollArea { border: none; background: transparent; }
            QScrollArea > QWidget > QWidget { background: transparent; }
            QMenu { background-color: white; color: #0f172a; border: 1px solid #cbd5e1; }
            QMenu::item:selected { background-color: #f1f5f9; }
            QToolTip { background: white; color: #0f172a; border: 1px solid #cbd5e1; }
        """)

    def init_ui(self):
        main_widget = QWidget()
        main_widget.setObjectName("mainWidget")
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setHandleWidth(6)

        # --- LEFT SIDE: Document Viewer ---
        left_container = QWidget()
        left_container.setObjectName("leftContainer")
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(20, 20, 20, 20)
        left_layout.setSpacing(15)
        
        left_card = QFrame()
        left_card.setProperty("class", "Card")
        left_card_layout = QVBoxLayout(left_card)
        left_card_layout.setContentsMargins(15, 15, 15, 15)
        left_card_layout.setSpacing(10)

        left_controls_layout = QHBoxLayout()
        
        self.upload_btn = QPushButton("Upload PDF Document")
        self.upload_btn.setObjectName("ActionBtn")
        self.upload_btn.clicked.connect(self.upload_file)
        
        self.close_preview_btn = QPushButton("Close Preview")
        self.close_preview_btn.setObjectName("ActionBtn")
        self.close_preview_btn.setStyleSheet("color: #dc2626;")
        self.close_preview_btn.clicked.connect(self.close_preview)
        self.close_preview_btn.hide()
        
        self.zoom_out_btn = QPushButton("-")
        self.zoom_out_btn.setObjectName("ToolBtn")
        self.zoom_in_btn = QPushButton("+")
        self.zoom_in_btn.setObjectName("ToolBtn")
        self.maximize_btn = QPushButton("🗖") 
        self.maximize_btn.setObjectName("ToolBtn")
        for btn in [self.zoom_out_btn, self.zoom_in_btn, self.maximize_btn]:
            btn.setFixedSize(40, 40)
            
        self.zoom_out_btn.clicked.connect(lambda: self.zoom_out(150))
        self.zoom_in_btn.clicked.connect(lambda: self.zoom_in(150))
        self.maximize_btn.clicked.connect(self.toggle_maximize)

        left_controls_layout.addWidget(self.upload_btn)
        left_controls_layout.addWidget(self.close_preview_btn)
        left_controls_layout.addStretch()
        z_label = QLabel("Zoom:")
        
        left_controls_layout.addWidget(z_label)
        left_controls_layout.addWidget(self.zoom_out_btn)
        left_controls_layout.addWidget(self.zoom_in_btn)
        left_controls_layout.addWidget(self.maximize_btn)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("QScrollArea { border: none; background: #f1f3f8; border-radius: 10px; }") 
        self.page_container = QWidget()
        self.page_container.setStyleSheet("background: #f1f3f8;")
        self.page_layout = QVBoxLayout(self.page_container)
        self.page_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter) 
        self.scroll_area.setWidget(self.page_container)
        self.scroll_area.viewport().installEventFilter(self)
        self.scroll_area.verticalScrollBar().setSingleStep(30)
        self.scroll_area.horizontalScrollBar().setSingleStep(30)

        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        self.text_preview.hide()
        self.text_preview.setStyleSheet("border: none; background: #f1f3f8; border-radius: 10px; padding: 10px; color: #333333;")

        left_card_layout.addLayout(left_controls_layout)
        tip_label = QLabel("Click and drag on the PDF to select and extract text, flowcharts, or generate your Introduction Summary!")
        tip_label.setStyleSheet("font-style: italic;")
        left_card_layout.addWidget(tip_label)
        left_card_layout.addWidget(self.scroll_area)
        left_card_layout.addWidget(self.text_preview)

        left_layout.addWidget(left_card)

        # --- RIGHT SIDE: Tools & Extraction ---
        self.right_container = QScrollArea()
        self.right_container.setWidgetResizable(True)
        self.right_container.setObjectName("rightContainer")
        self.right_container.setStyleSheet("QScrollArea#rightContainer { border: none; background: transparent; }")
        self.right_container.verticalScrollBar().setSingleStep(30)
        
        right_widget = QWidget()
        right_widget.setStyleSheet("background: transparent;")
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(20, 20, 20, 20)
        right_layout.setSpacing(20)
        self.right_container.setWidget(right_widget)

        # Card 1: Document Status & Checklist
        status_card = QFrame()
        status_card.setProperty("class", "Card")
        status_main_layout = QVBoxLayout(status_card)
        status_main_layout.setContentsMargins(15, 15, 15, 15)
        
        top_bar_layout = QHBoxLayout()
        title_lbl = QLabel("<h3>Task Directive Extraction Checklist</h3>")
        title_lbl.setProperty("class", "Heading")
        top_bar_layout.addWidget(title_lbl)
        top_bar_layout.addStretch()
        
        # Theme toggle removed entirely
        
        self.summary_icon_btn = QPushButton("AI")
        self.summary_icon_btn.setFixedSize(50, 50)
        self.summary_icon_btn.setObjectName("IconBtn")
        self.summary_icon_btn.setToolTip("View/Edit Introduction Summary")
        self.summary_icon_btn.clicked.connect(self.show_summary_popup)
        top_bar_layout.addWidget(self.summary_icon_btn)
        status_main_layout.addLayout(top_bar_layout)
        
        chk_layout = QHBoxLayout()
        self.check_upload = QCheckBox("File Uploaded")
        self.check_upload.setEnabled(False)
        self.check_upload.setStyleSheet("font-weight: bold;")
        chk_layout.addWidget(self.check_upload)
        
        self.topics = [
            "Introduction",
            "Scope Of Task Directive", "Product Break Down Structure" , "System LRU and Module detail" , " Check validation", "Test Rig Simulators ground equipements","Work Breakdown Structure"
                ]        
        self.topic_checkboxes = {}
        
        topics_scroll = QScrollArea()
        topics_scroll.setWidgetResizable(True)
        topics_scroll.setFixedHeight(120) 
        topics_scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")
        
        topics_widget = QWidget()
        topics_layout = QGridLayout(topics_widget) 
        topics_layout.setContentsMargins(5, 5, 5, 5)
        
        row, col = 0, 0
        for topic in self.topics:
            cb = QCheckBox(topic)
            
            self.topic_checkboxes[topic] = cb
            topics_layout.addWidget(cb, row, col)
            col += 1
            if col > 1: 
                col = 0
                row += 1

        topics_scroll.setWidget(topics_widget)
        status_main_layout.addLayout(chk_layout)
        status_main_layout.addWidget(topics_scroll)

        self.status_card = status_card
        right_layout.addWidget(status_card)

        # Card 2: Manual Data
        manual_card = QFrame()
        manual_card.setProperty("class", "Card")
        manual_layout = QVBoxLayout(manual_card)
        manual_layout.setContentsMargins(15, 15, 15, 15)

        manual_title = QLabel("<h3>Manual Data</h3>")
        manual_title.setProperty("class", "Heading")

        manual_header_layout = QHBoxLayout()
        manual_header_layout.addWidget(manual_title)
        manual_header_layout.addStretch()

        self.manual_collapse_btn = QPushButton("–")
        self.manual_collapse_btn.setObjectName("ToolBtn")
        self.manual_collapse_btn.setFixedSize(32, 32)
        self.manual_collapse_btn.setToolTip("Collapse/Expand Manual Input")
        self.manual_collapse_btn.clicked.connect(self.toggle_manual_collapse)

        self.manual_section_maximize_btn = QPushButton("🗖")
        self.manual_section_maximize_btn.setObjectName("ToolBtn")
        self.manual_section_maximize_btn.setFixedSize(32, 32)
        self.manual_section_maximize_btn.setToolTip("Maximize/Restore Manual Input")
        self.manual_section_maximize_btn.clicked.connect(self.toggle_manual_section_maximize)

        manual_header_layout.addWidget(self.manual_collapse_btn)
        manual_header_layout.addWidget(self.manual_section_maximize_btn)
        manual_layout.addLayout(manual_header_layout)

        manual_body = QWidget()
        manual_body_layout = QVBoxLayout(manual_body)
        manual_body_layout.setContentsMargins(0, 0, 0, 0)
        manual_body_layout.setSpacing(12)

        proj_label = QLabel("Project Name:")
        proj_label.setStyleSheet("font-weight: bold;")
        manual_body_layout.addWidget(proj_label)

        self.project_name_input = QLineEdit()
        self.project_name_input.setPlaceholderText("Type or paste Project Name here...")
        manual_body_layout.addWidget(self.project_name_input)

        dropdowns = [
            ("Scope Of Task Directive", ["Development of air system in accordance with IMTAR-21 Subpart- B", "Development of air system in accordance with IMTAR-21 Subpart- B", "Development of Airborne Software and CEH in accordance with IMTAR-21 Subpart- C6.", "Development of Ground support system in accordance with IMTAR-21 Subpart- T.", " Continued Airworthiness Coverage in accordance with IMTAR-21 Subpart-L. " , "Bought Out Item clearance in accordance with IMTAR-21 Subpart N and ACR 002/2025" , "Enter Other..."]),
            ("Certification Work Breakdown", ["Carry out SSA and classify the criticality of each System and LRU", "Requirements Analysis of each System and LRU ", "Finalization and approval of LRU Specification and build standard ", "Test Requirement Traceability Matrix (Means of Compliance)" , "Finalization of Type Approval Basis (TAB) and Airworthiness Certification Plan " , "Software Certification Plan (PSAC)" , "Hardware and CEH certification Plan" , "Preliminary Design Review & Critical Design Review " , "Test Adequacy Review " , "LRU Test Rig Specification Approval & Rig Acceptance by DGAQA" ,"Functional Test Plan, Integration Test Plan " ," Realization of Hardware and inspection by DGAQA " ,"Safety of Flight Testing/Qualification Testing " , "IV & V of Software and CEH, and Software certification activities ", "System Integration Testing" ,"Clearance for flight trials" ,"Satisfactory Flight trial feedback from Flight Ops/ Users" ,"Compliance to TAB, ACP, QTP, PSAC, TRTM" ,"Verification of Flight Trial feedback", "Production clearance & Service use clearance for the system", "RMTC/ MTC of the air system", "Continued Airworthiness Activities", "Enter Other..."]),
            ("Internal Distribution", ["Director (System’s)", "Director (A/C)", "Director (Propulsion)", "mDirector (Mat & PI)", "Group Director (MS)" , "RD, RCMA(APS)", "RD, RCMA (A/C)", "RD, RCMA(0.)", "RD, RCMA()", "RD, RCMA()", "RD, RCMA()", "E-Certification ", "C-Cat lab", "SRM CenterHead" ,"Enter Other..."]),
            
        ]

        dropdown_card = QFrame()
        dropdown_card.setStyleSheet("""
            background: white;
            border-radius: 12px;
            border: 1px solid #e2e8f0;
            padding: 15px;
        """)

        dropdown_layout = QVBoxLayout(dropdown_card)

        self.combos = []

        for label, items in dropdowns:
            row = QHBoxLayout()
            row.setSpacing(15)
            row.setContentsMargins(5, 5, 5, 5)

            lbl = QLabel(label + ":")
            lbl.setFixedWidth(220)
            lbl.setStyleSheet("""
                QLabel {
                    font-weight: 600;
                    font-size: 14px;
                    color: #1e293b;
                }
            """)

            combo = MultiSelectWidget(items)
            combo.setStyleSheet("""
                QTextEdit {
                    background-color: white;
                    border: 1px solid #cbd5e1;
                    border-radius: 8px;
                    padding: 8px;
                    font-size: 13px;
                }

                QComboBox {
                    background-color: white;
                    border: 1px solid #cbd5e1;
                    border-radius: 8px;
                    padding: 8px;
                }

                QLineEdit {
                    border-radius: 8px;
                    padding: 6px;
                }
            """)

            row.addWidget(lbl)
            row.addWidget(combo)
            dropdown_layout.addLayout(row)
            self.combos.append(combo)
            # Map common dropdown labels to named attributes used in export
            low = label.lower()
            if "scope" in low:
                self.scope_dropdown = combo
            elif "certification" in low or "certification work" in low:
                self.cert_dropdown = combo
            elif "internal" in low or "Internal distribution" in low:
                self.InternalDistribution_dropdown = combo

        manual_body_layout.addWidget(dropdown_card)

        self.manual_heading_widget = QWidget()
        manual_select_layout = QHBoxLayout()
        self.manual_heading_widget.setLayout(manual_select_layout)

        manual_heading_label = QLabel("Select Heading:")
        manual_heading_label.setStyleSheet("font-weight: bold;")

        self.manual_heading_dropdown = QComboBox()
        self.manual_heading_dropdown.addItems([
            "Introduction",
            "Reference",
            "Basis Of Task Directive",
            "Scope Of Task Directive",
            "Stakeholders",
            "Certification Work Breakdown",
            "Task Allocation",
            "Communication",
            "Annexure – 3"
        ])
        self.manual_heading_dropdown.currentTextChanged.connect(self.load_section_note)
        manual_select_layout.addWidget(manual_heading_label)
        manual_select_layout.addWidget(self.manual_heading_dropdown)
        manual_body_layout.addWidget(self.manual_heading_widget)

        # Table selector dropdown
        table_selector_layout = QHBoxLayout()
        table_selector_label = QLabel("Edit Table:")
        table_selector_label.setStyleSheet("font-weight: bold;")
        
        self.table_selector_dropdown = QComboBox()
        self.table_selector_dropdown.addItem("-- Select a Table --")
        self.table_selector_dropdown.addItems([
            "5. Stakeholders",
            "7.3 Certification Task Allocation",
            "Annexure-1: Work Assignment List",
            "Annexure-1: Integration Checks",
            "Annexure-2: Contact Details"
        ])
        self.table_selector_dropdown.currentTextChanged.connect(self.on_table_selected)
        
        table_selector_layout.addWidget(table_selector_label)
        table_selector_layout.addWidget(self.table_selector_dropdown)
        table_selector_layout.addStretch()
        manual_body_layout.addLayout(table_selector_layout)

        self.manual_notes_label = QLabel("Manual Notes:")
        self.manual_notes_label.setStyleSheet("font-weight: bold; margin-top: 10px;")
        manual_body_layout.addWidget(self.manual_notes_label)
        self.manual_input = NotesEditor()
        self.manual_input.setMinimumHeight(260)
        self.manual_input.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.manual_input.setPlaceholderText("Extracted contents here...")
        manual_body_layout.addWidget(self.manual_input)
        self.manual_input.textChanged.connect(self.save_current_section_note)

        self.manual_project_label = proj_label
        self.manual_dropdown_card = dropdown_card

        self.manual_body = manual_body
        manual_layout.addWidget(manual_body)
        right_layout.addWidget(manual_card, stretch=1)

        # Card 3: Actions
        actions_card = QFrame()
        actions_card.setProperty("class", "Card")
        actions_layout = QVBoxLayout(actions_card)
        actions_layout.setContentsMargins(15, 15, 15, 15)
        
        actions_title = QLabel("<h3>Actions</h3>")
        actions_title.setProperty("class", "Heading")
        actions_layout.addWidget(actions_title)

        self.status_label = QLabel("Ready")
        self.status_label.setStyleSheet("font-style: italic;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0) 
        self.progress_bar.hide()
        
        actions_layout.addWidget(self.status_label)
        actions_layout.addWidget(self.progress_bar)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        
        self.save_docx_btn = QPushButton("Generate Task Directive (Word)")
        self.save_docx_btn.setFixedHeight(45)
        self.save_docx_btn.setStyleSheet("background-color: #059669; color: white;")
        self.save_docx_btn.clicked.connect(self.export_docx)
        
        self.preview_pdf_btn = QPushButton("Preview PDF")
        self.preview_pdf_btn.setFixedHeight(45)
        self.preview_pdf_btn.setStyleSheet("background-color: #059669; color: white;")
        self.preview_pdf_btn.clicked.connect(self.preview_pdf)

        btn_layout.addWidget(self.save_docx_btn)
        btn_layout.addWidget(self.preview_pdf_btn)
        actions_layout.addLayout(btn_layout)

        self.actions_card = actions_card

        right_layout.addWidget(actions_card)

        self.splitter.addWidget(left_container)
        self.splitter.addWidget(self.right_container)
        self.splitter.setStretchFactor(0, 5) 
        self.splitter.setStretchFactor(1, 5) 
        layout.addWidget(self.splitter)


    def detect_pdf_headings(self, pdf_doc):
        detected_headings = []

        for page in pdf_doc:
            blocks = page.get_text("dict")["blocks"]

            for block in blocks:
                if "lines" not in block:
                    continue

                for line in block["lines"]:
                    for span in line["spans"]:

                        text = span["text"].strip()

                        if not text:
                            continue

                        font_size = span["size"]
                        font_name = span["font"].lower()

                        is_bold = (
                            "bold" in font_name
                            or "black" in font_name
                            or "heavy" in font_name
                        )

                        # HEADING RULES
                        if (
                            font_size >= 13      # Larger than body text
                            or is_bold           # Bold text
                        ) and len(text) < 80:    # Avoid paragraphs

                            detected_headings.append(text)
        
        # Normalize and deduplicate
        norm = []
        seen = set()
        for h in detected_headings:
            s = h.strip()
            if not s or len(s) > 120:
                continue
            # collapse whitespace
            s = re.sub(r"\s+", " ", s)
            s_lower = s.lower()
            if s_lower not in seen:
                norm.append(s)
                seen.add(s_lower)

        print("Detected Headings:", norm)
        return norm

    # ------------------------------------------
    # --- EDITABLE SUMMARY POPUP ---
    # ------------------------------------------
    def show_summary_popup(self):
        """Opens a popup window to view and edit the generated summary."""
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit Introduction Summary")
        dialog.resize(600, 400)
        layout = QVBoxLayout(dialog)

        text_edit = QTextEdit()
        text_edit.setStyleSheet("background-color: white; color: black; font-size: 11pt;")
        text_edit.setText(self.intro_text_data)
        
        text_edit.textChanged.connect(lambda: setattr(self, 'intro_text_data', text_edit.toPlainText()))

        close_btn = QPushButton("Save & Close")
        close_btn.setFixedHeight(40)
        close_btn.clicked.connect(dialog.accept)

        layout.addWidget(text_edit)
        layout.addWidget(close_btn)
        dialog.exec()

#--------------Add Save/Load Methods------------------#
    def load_section_note(self):
        heading = self.manual_heading_dropdown.currentText()

        # Special handling for Annexure – 3: image-only mode
        if heading == "Annexure – 3":
            self.manual_input.blockSignals(True)
            images = self.section_images.get(heading, [])
            image_count = len(images)
            self.manual_input.setPlainText(f"Annexure – 3 (Images only)\n\nNumber of images stored: {image_count}\n\n[Extract images from PDF and they will be added to this section]")
            self.manual_input.setReadOnly(True)
            self.manual_input.blockSignals(False)
        else:
            self.manual_input.setReadOnly(False)
            self.manual_input.blockSignals(True)
            self.manual_input.setPlainText(
                self.section_notes.get(heading, "")
            )
            self.manual_input.blockSignals(False)
    
    def save_current_section_note(self):
        heading = self.manual_heading_dropdown.currentText()
        
        # Don't save text for Annexure – 3 (images only)
        if heading != "Annexure – 3":
            self.section_notes[heading] = self.manual_input.toPlainText()

    def on_table_selected(self, table_name):
        """Handle table selection and open editor if a valid table is selected"""
        if table_name == "-- Select a Table --":
            return
        
        # Define table structures
        table_configs = {
            "5. Stakeholders": {
                "headers": ["Sl No.", "Organisation", "Role", "Activities"],
                "rows": 3
            },
            "7.3 Certification Task Allocation": {
                "headers": ["Sl No.", "Certification Activity", "Certification Work Centre", "Responsible Head"],
                "rows": 3
            },
            "Annexure-1: Work Assignment List": {
                "headers": ["Sl No", "Certifiable item", "Design agency", "Applicable Subpart", "Designated Directorate/ RCMA", "Dealing Officer(s)"],
                "rows": 3
            },
            "Annexure-1: Integration Checks": {
                "headers": ["Sl no.", "System Name", "Design agency", "Designated RCMA", "Dealing Officer"],
                "rows": 2
            },
            "Annexure-2: Contact Details": {
                "headers": ["Sl no.", "Name & Designation", "Certification Centre", "E-mail id", "Phone number"],
                "rows": 2
            }
        }
        
        if table_name not in table_configs:
            return
        
        config = table_configs[table_name]
        headers = config["headers"]
        
        # Get stored data or create new
        current_data = self.table_data.get(table_name, [["" for _ in headers] for _ in range(config["rows"])])
        
        # Open editor dialog
        dialog = TableEditorDialog(table_name, headers, current_data, self)
        
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # Save the edited data
            self.table_data[table_name] = dialog.get_table_data()
            
            # Reset dropdown
            self.table_selector_dropdown.blockSignals(True)
            self.table_selector_dropdown.setCurrentIndex(0)
            self.table_selector_dropdown.blockSignals(False)
            
            QMessageBox.information(self, "Success", f"Table '{table_name}' saved successfully!")
        else:
            # Reset dropdown if cancelled
            self.table_selector_dropdown.blockSignals(True)
            self.table_selector_dropdown.setCurrentIndex(0)
            self.table_selector_dropdown.blockSignals(False)


    # ------------------------------------------
    # --- AUTO-PASTING FUNCTIONS ---
    # ------------------------------------------
    def add_extracted_text(self, text):
        QApplication.clipboard().setText(text) 
        self.manual_input.append(text + "\n")
        
    def add_extracted_image(self, pixmap):
        if pixmap.width() > 500:
            pixmap = pixmap.scaledToWidth(500, Qt.TransformationMode.SmoothTransformation)
        QApplication.clipboard().setPixmap(pixmap)
        cursor = self.manual_input.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End) 
        self.manual_input.setTextCursor(cursor)
        cursor.insertImage(pixmap.toImage()) 
        self.manual_input.append("\n")
        
        # Store image with current heading
        current_heading = self.manual_heading_dropdown.currentText()
        if current_heading not in self.section_images:
            self.section_images[current_heading] = []
        self.section_images[current_heading].append(pixmap) 
        
    def generate_summary_from_selection(self, text):
        self.status_label.setText("Summarizing selection...")
        QApplication.processEvents()
        summary = simple_summarize(text, num_sentences=4)
        if summary:
            self.intro_text_data = summary 
            
            self.summary_icon_btn.setObjectName("IconBtnActive"); self.summary_icon_btn.style().unpolish(self.summary_icon_btn); self.summary_icon_btn.style().polish(self.summary_icon_btn)
            
            self.status_label.setText("Ready")
            self.topic_checkboxes["Introduction"].setChecked(True)
            QMessageBox.information(self, "Success", "Selection summarized! Click the AI icon to view or edit it.")
        else:
            self.status_label.setText("Ready")
            QMessageBox.warning(self, "Empty", "Could not generate a summary. Try selecting a larger paragraph.")

    # ------------------------------------------
    # --- UI INTERACTION FUNCTIONS ---
    # ------------------------------------------
    def eventFilter(self, source, event):
        if source == self.scroll_area.viewport() and event.type() == QEvent.Type.Wheel:
            if QApplication.keyboardModifiers() == Qt.KeyboardModifier.ControlModifier:
                delta = event.angleDelta().y()
                if delta > 0: self.zoom_in(step=50)
                elif delta < 0: self.zoom_out(step=50)
                return True 
        return super().eventFilter(source, event)

    def toggle_maximize(self):
        if not self.is_maximized:
            self.right_container.hide() 
            self.maximize_btn.setText("🗗") 
            self.is_maximized = True
        else:
            self.right_container.show() 
            self.maximize_btn.setText("🗖") 
            self.is_maximized = False

    def toggle_manual_collapse(self):
        self.manual_collapsed = not self.manual_collapsed
        self.manual_input.setVisible(not self.manual_collapsed)
        self.manual_notes_label.setVisible(not self.manual_collapsed)
        self.manual_collapse_btn.setText("+" if self.manual_collapsed else "–")

    def toggle_manual_section_maximize(self):
        if not self.manual_section_maximized:
            self.manual_input.setMaximumHeight(1000)
            self.manual_section_maximized = True
            self.manual_section_maximize_btn.setText("🗗")
        else:
            self.manual_input.setMaximumHeight(16777215)
            self.manual_section_maximized = False
            self.manual_section_maximize_btn.setText("🗖")

    def zoom_in(self, step=150):
        self.current_zoom += step
        self.refresh_pdf_view()

    def zoom_out(self, step=150):
        if self.current_zoom - step > 200: 
            self.current_zoom -= step
            self.refresh_pdf_view()

    def refresh_pdf_view(self):
        if not self.page_data: return

        if self.page_layout.count() != len(self.page_data):
            while self.page_layout.count():
                child = self.page_layout.takeAt(0)
                if child.widget(): child.widget().deleteLater()
            for pixmap, page_idx in self.page_data:
                page = self.doc.load_page(page_idx)
                lbl = PdfPageLabel(page, self.scroll_area)
                lbl.text_extracted.connect(self.add_extracted_text)
                lbl.image_extracted.connect(self.add_extracted_image)
                lbl.summary_requested.connect(self.generate_summary_from_selection)
                lbl.setStyleSheet("background-color: white; border: 1px solid #e0e0e0; margin-bottom: 10px; border-radius: 4px;")
                self.page_layout.addWidget(lbl)

        for i in range(self.page_layout.count()):
            lbl = self.page_layout.itemAt(i).widget()
            if isinstance(lbl, PdfPageLabel) and i < len(self.page_data):
                scaled_pixmap = self.page_data[i][0].scaledToWidth(self.current_zoom, Qt.TransformationMode.SmoothTransformation)
                lbl.setPixmap(scaled_pixmap)
                lbl.setFixedSize(scaled_pixmap.size())

    # ------------------------------------------
    # --- CORE LOGIC & FUNCTIONS ---
    # ------------------------------------------
    def close_preview(self):
        if not hasattr(self, 'original_file_path') or not self.original_file_path:
            return
            
        self.close_preview_btn.hide()
        path = self.original_file_path
        self.file_path = path
        
        try:
            self.page_data.clear()
            if path.lower().endswith('.pdf'):
                self.text_preview.hide()
                self.scroll_area.show()
                self.doc = fitz.open(path)
                
                self.status_label.setText("Restoring original PDF...")
                QApplication.processEvents()

                for i, page in enumerate(self.doc):
                    if i < 15: 
                        pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                        img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()
                        self.page_data.append((QPixmap.fromImage(img), i))
                
                self.current_zoom = 600 
                self.refresh_pdf_view() 
            else:
                self.scroll_area.hide()
                self.text_preview.show()
                with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                    self.text_preview.setText(f.read())
                    
            self.status_label.setText("Restored original uploaded document.")
        except Exception as e:
            QMessageBox.critical(self, "Restore Error", str(e))

    def upload_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Document", "", "Documents (*.pdf *.txt)")
        if path:
            self.file_path = path
            self.original_file_path = path
            self.close_preview_btn.hide()
            
            self.check_upload.setChecked(False)
            for cb in self.topic_checkboxes.values():
                cb.setChecked(False)
                
            self.project_name_input.clear()
            self.intro_text_data = "" 
            self.summary_icon_btn.setObjectName("IconBtn"); self.summary_icon_btn.style().unpolish(self.summary_icon_btn); self.summary_icon_btn.style().polish(self.summary_icon_btn)
            
            self.manual_input.clear() 
            self.page_data.clear() 
            
            try:
                if path.lower().endswith('.pdf'):
                    self.text_preview.hide()
                    self.scroll_area.show()
                    self.doc = fitz.open(path)
                    
                    self.status_label.setText("Loading PDF Pages...")
                    QApplication.processEvents()

                    for i, page in enumerate(self.doc):
                        if i < 15: 
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()
                            self.page_data.append((QPixmap.fromImage(img), i))
                    
                    self.current_zoom = 600 
                    self.refresh_pdf_view() 
                else:
                    self.scroll_area.hide()
                    self.text_preview.show()
                    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                        self.text_preview.setText(f.read())

                self.check_upload.setChecked(True)
                
                full_text = self.text_preview.toPlainText() if not self.doc else "".join(page.get_text() for page in self.doc)

                if self.doc:
                    detected_headings = self.detect_pdf_headings(self.doc)
                else:
                    detected_headings = []

                self.start_analysis_thread(full_text, detected_headings)
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not load file: {str(e)}")

    def start_analysis_thread(self, text, headings):
        self.status_label.setText("Scanning document with Python NLP...")
        self.progress_bar.show()
        self.worker = AnalysisWorker(text, self.topics, headings)
        self.worker.finished.connect(self.handle_analysis_done)
        self.worker.error.connect(lambda e: self.status_label.setText(f"Analysis Error: {e}"))
        self.worker.start()

    def handle_analysis_done(self, results):
        self.progress_bar.hide()
        self.status_label.setText("Analysis Complete")
        
        if results["project_name"]:
            self.project_name_input.setText(results["project_name"])
            
        if results["introduction"]:
            self.intro_text_data = results["introduction"]
            self.summary_icon_btn.setObjectName("IconBtnActive"); self.summary_icon_btn.style().unpolish(self.summary_icon_btn); self.summary_icon_btn.style().polish(self.summary_icon_btn)
            
        for topic in results["found_topics"]:
            if topic in self.topic_checkboxes:
                self.topic_checkboxes[topic].setChecked(True)

    # ------------------------------------------
    # ------ DOCUMENT GENERATION SECTION  ------
    # ------------------------------------------


    """
            ★
           ★★★
          ★★★★★
         ★★★★★★★
        ★★★★★★★★★
       ★★★★★★★★★★★
      ★★★★★★★★★★★★★
     ★★★★★★★★★★★★★★★
    ★★★★★★★★★★★★★★★★★
   ★★★★★★★★★★★★★★★★★★★
  ★★★★★★★★★★★★★★★★★★★★★
 ★★★★★★★★★★★★★★★★★★★★★★★
★★★★★★★★★★★★★★★★★★★★★★★★★
 ★★★★★★★★★★★★★★★★★★★★★★★
  ★★★★★★★★★★★★★★★★★★★★★
   ★★★★★★★★★★★★★★★★★★★
    ★★★★★★★★★★★★★★★★★
     ★★★★★★★★★★★★★★★
      ★★★★★★★★★★★★★
       ★★★★★★★★★★★
        ★★★★★★★★★
         ★★★★★★★
          ★★★★★
           ★★★
            ★
    """






    def preview_pdf(self):

        if not hasattr(self, "generated_preview_path") or not self.generated_preview_path:
            QMessageBox.warning(
                self,
                "No Preview Available",
                "Please generate/export the Task Directive first."
            )
            return

        try:
            self.file_path = self.generated_preview_path

            self.text_preview.hide()
            self.scroll_area.show()

            self.doc = fitz.open(self.generated_preview_path)

            self.status_label.setText("Loading Saved Task Directive Preview...")
            QApplication.processEvents()

            self.page_data.clear()

            for i, page in enumerate(self.doc):

                if i < 20:

                    pix = page.get_pixmap(
                        matrix=fitz.Matrix(2.0, 2.0)
                    )

                    img = QImage(
                        pix.samples,
                        pix.width,
                        pix.height,
                        pix.stride,
                        QImage.Format.Format_RGB888
                    ).copy()

                    self.page_data.append(
                        (QPixmap.fromImage(img), i)
                    )

            self.current_zoom = 600

            self.refresh_pdf_view()

            self.close_preview_btn.show()

            self.status_label.setText("Saved Task Directive Preview Loaded!")

        except Exception as e:

            QMessageBox.critical(
                self,
                "Preview Error",
                f"Failed to open saved preview:\n{str(e)}"
            )

    def inject_images_for_heading(self, doc, heading_name):
        """Add all images associated with a heading to the document"""
        images = self.section_images.get(heading_name, [])
        for pixmap in images:
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Save pixmap temporarily and insert into doc
            import tempfile
            with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                pixmap.save(tmp.name, "PNG")
                p.add_run().add_picture(tmp.name, width=Inches(4.5))
            doc.add_paragraph()  # Add blank line after image

    def export_docx(self):

        self.save_current_section_note()

        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Word",
            "Task_Directive_Final.docx",
            "Word (*.docx)"
        )

        if not save_path:
            return

        try:

            def add_page_border(section):
                sectPr = section._sectPr
                pgBorders = OxmlElement('w:pgBorders')
                pgBorders.set(qn('w:offsetFrom'), 'page')

                for side in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{side}')
                    border.set(qn('w:val'), 'single')
                    border.set(qn('w:sz'), '12')
                    border.set(qn('w:space'), '24')
                    border.set(qn('w:color'), '000000')
                    pgBorders.append(border)

                sectPr.append(pgBorders)

            def remove_page_border(section):
                sectPr = section._sectPr
                for el in sectPr.findall(qn('w:pgBorders')):
                    sectPr.remove(el)

            def add_heading(doc, text):
                p = doc.add_paragraph()
                run = p.add_run(text)
                run.bold = True
                run.font.size = Pt(12)

            doc = Document()

            style = doc.styles['Normal']
            style.font.name = 'Calibri'
            style.font.size = Pt(11)

            # ---------------- COVER PAGE ----------------
            section1 = doc.sections[0]

            section1.top_margin = Inches(1.5)
            section1.bottom_margin = Inches(1.5)
            section1.left_margin = Inches(1.5)
            section1.right_margin = Inches(1.5)

            add_page_border(section1)

            doc.add_paragraph("File No. ________________________")

            p = doc.add_paragraph()
            run = p.add_run("\nTASK DIRECTIVE\n")
            run.bold = True
            run.font.size = Pt(20)

            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            p = doc.add_paragraph("________ /2026")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            table = doc.add_table(rows=1, cols=2)
            table.cell(0, 0).text = "Issue No."
            table.cell(0, 1).text = "Date of Issue:"

            p_name = self.project_name_input.text().strip()

            if not p_name or p_name.lower() == "not found":
                p_name = "__________________________________________"

            doc.add_paragraph(f"\nPROJECT NAME: {p_name}")

            p = doc.add_paragraph("\n\n[ MEOWFISHHHHHHHHH HERE ]")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            doc.add_paragraph("CENTRE FOR MILITARY AIRWORTHINESS AND CERTIFICATION \n (CEMILAC)").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("DRDO, MoD").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("Marathahalli Colony Post, \n Bengaluru – 560037").alignment = WD_ALIGN_PARAGRAPH.CENTER

            # ---------------- MAIN CONTENT ----------------
            section2 = doc.add_section(WD_SECTION.NEW_PAGE)

            section2.header.is_linked_to_previous = False
            section2.footer.is_linked_to_previous = False

            section2.top_margin = Inches(0.75)
            section2.bottom_margin = Inches(0.75)
            section2.left_margin = Inches(0.75)
            section2.right_margin = Inches(0.75)

            remove_page_border(section2)

            # 1 INTRODUCTION
            add_heading(doc, "1. Introduction [Automated]")

            intro_text = self.intro_text_data.strip()
            if not intro_text:
                intro_text = "__________________________________________________"

            doc.add_paragraph(intro_text)

            note = self.section_notes.get("Introduction", "")
            if note.strip():
                doc.add_paragraph(note)
            
            # Inject images for Introduction
            self.inject_images_for_heading(doc, "Introduction")

            # 2 REFERENCE
            add_heading(doc, "2. Reference [Default]")
            doc.add_paragraph("  1. PBS-WBS \n 2. Project brief ")

            note = self.section_notes.get("Reference", "")
            if note.strip():
                doc.add_paragraph(note)
            
            # Inject images for Reference
            self.inject_images_for_heading(doc, "Reference")

            # 3 BASIS
            add_heading(doc, "3. Basis Of Task Directive")
            doc.add_paragraph("  1. IMAP-23, Indian Military Airworthiness Procedure - 23\n 2. IMTAR-21, Version 2.0 Indian Military Technical Airworthiness Requirements\n 3. Applicable Airworthiness Directives and CEMILAC Directives ")
            note = self.section_notes.get("Basis Of Task Directive", "")
            if note.strip():
                doc.add_paragraph(note)
            
            # Inject images for Basis Of Task Directive
            self.inject_images_for_heading(doc, "Basis Of Task Directive")

            # 4 SCOPE
            scope_val = getattr(self, 'scope_dropdown', None)

            scope_text = "\n".join(scope_val.checkedItems()) if scope_val and scope_val.checkedItems() else "____"

            add_heading(doc, "4. Scope Of Task Directive")

            doc.add_paragraph(
                f"The scope of the Task Directive is to assign the certification responsibilities to CEMILAC and RCMAs to enable smooth communication and transactions between the design agencies and the certification agency. The designated RCMAs provide concurrent certification coverage for:"
            )
            doc.add_paragraph(scope_text)
            note = self.section_notes.get("Scope Of Task Directive", "")
            if note.strip():
                doc.add_paragraph(note)
            
            # Inject images for Scope Of Task Directive
            self.inject_images_for_heading(doc, "Scope Of Task Directive")

            # 5 STAKEHOLDERS
            sth_val = getattr(self, 'stakeholder_dropdown', None)

            sth_text = "\n".join(sth_val.checkedItems()) if sth_val and sth_val.checkedItems() else "Both"

            add_heading(doc, f" Stakeholders ({sth_text})")

            doc.add_paragraph("The following are the major stakeholders from certification perspective for the development, integration of the system on the various Rigs.")

            note = self.section_notes.get("Stakeholders", "")
            if note.strip():
                doc.add_paragraph(note)
            
            # Inject images for Stakeholders
            self.inject_images_for_heading(doc, "Stakeholders")

            table = doc.add_table(rows=4, cols=4)
            table.style = 'Table Grid'

            headers = ["Sl No.", "Organisation", "Role", "Activities"]
            for i, h in enumerate(headers):
                table.cell(0, i).text = h
            
            # Add stored table data
            table_data = self.table_data.get("5. Stakeholders", [["" for _ in headers] for _ in range(3)])
            for row_idx, row_data in enumerate(table_data):
                for col_idx, cell_data in enumerate(row_data):
                    if row_idx + 1 < table.rows.__len__():  # +1 because row 0 is header
                        table.cell(row_idx + 1, col_idx).text = str(cell_data)

            # 6 CERTIFICATION
            cert_val = getattr(self, 'cert_dropdown', None)

            cert_text = "\n".join(cert_val.checkedItems()) if cert_val and cert_val.checkedItems() else "______________________________________________"

            add_heading(doc, "6. Certification Work Breakdown ")

            doc.add_paragraph(cert_text)

            note = self.section_notes.get("Certification Work Breakdown", "")
            if note.strip():
                doc.add_paragraph(note)
            
            # Inject images for Certification Work Breakdown
            self.inject_images_for_heading(doc, "Certification Work Breakdown")

            # 7 TASK ALLOCATION
            task_val = getattr(self, 'task_dropdown', None)

            task_text = "\n".join(task_val.checkedItems()) if task_val and task_val.checkedItems() else ""

            add_heading(doc, "7. Task Allocation")
            doc.add_paragraph(task_text)

            note = self.section_notes.get("Task Allocation", "")
            if note.strip():
                doc.add_paragraph(note)
            
            # Inject images for Task Allocation
            self.inject_images_for_heading(doc, "Task Allocation")

            add_heading(doc, "7.1 Coordinating Directorate ")
            doc.add_paragraph("The overall certification of the project would be managed by: \n Director () / RCMA ()\n The responsibilities of the Coordinating Directorate would be as follows:\n\n\n i. Plan, review and monitor the progress of the certification.\n\n ii. Participate in System level and software meetings \n \niii. Propose Constitution of SCRB and TARB as and when required.")
            add_heading(doc, "7.2 Single Point of Contact (SPoC) [Default]")
            doc.add_paragraph("The following officer would be responsible for overall certification coordination with all the stakeholders for smooth certification of the project:\n Sri …..., Sc-... \n RCMA ()/ Dte ()\n \n \n The responsibilities of the SPoC would be as follows: \n i. Ensure prompt assignment of Project Leader/ Dealing Officer by all RCMAs. \n ii. Ensure correct uplinking and downlinking of each project registered by the Design Agency, in line with the PBS\n iii. Work in coordination with the D&D agency and ensure timely execution of certification activities.\n iv. Liaise with concerned RCMAs, CEMILAC directorates and other stakeholders\n v. Ensure finalisation of TAB and ACP after taking inputs from all stakeholders \n vi. Keep a repository on incoming and outgoing artefacts/documents and ensure timely dissemination of information to all certification centres/stakeholders. \n vii. Apprise Coordinating Directorate the progression of certification activities.\n viii. Plan and coordinate Certification Review Meetings.\n ix. Member Secretary for SCRB and TARB as and when required. \n x. Coordination for installation/integration/flight clearance.")

            add_heading(doc, "7.3 Certification Task Allocation")
            doc.add_paragraph("The detailed certification task allocation is given below. The detailed list of allocation of certification responsibility to RCMAs and dealing officers, in respect of each System/LRU, is given in Annexure – 1 and contact details are given in Annexure – 2 respectively.")
            table = doc.add_table(rows=4, cols=4)
            table.style = 'Table Grid'

            headers = [
                "Sl No.",
                "Certification Activity",
                "Certification Work Centre",
                "Responsible Head"
            ]
            for i, h in enumerate(headers):
                table.cell(0, i).text = h
            
            # Add stored table data
            table_data = self.table_data.get("7.3 Certification Task Allocation", [["" for _ in headers] for _ in range(3)])
            for row_idx, row_data in enumerate(table_data):
                for col_idx, cell_data in enumerate(row_data):
                    if row_idx + 1 < len(table.rows):
                        table.cell(row_idx + 1, col_idx).text = str(cell_data)

            add_heading(doc, "7.4 Issue of Clearance [Default]")
            doc.add_paragraph("""a. On satisfactory completion of Software IV&V, clearance for Software would be issued by RD, RCMA (…) 
       \n b. On satisfactory completion of SOFT/QT tests on the LRU, clearance for Hardware and CEH would be issued by RDs of respective system RCMAs as per Annexure-1. 
        \n c. On completion of Software, Hardware and CEH clearances of LRUs, and compliance of Systems to TARB, the installation clearances for the systems would be issued by. 
       \n  d. On satisfactory completion of ground testing/ground adaptation, the system would be cleared by platform RD for ground integration and flight trials. 
       \n  e. On satisfactory flight trial completion, System RCMAs shall issue Provisional Clearance and service use for the LRUs. 
        \n f. Service use clearance for the Systems will be issued by.
       \n  g. Based on the recommendation of Director(Aircraft), the platform will be issued with RMTC/ MTC by CE(A), CEMILAC.""")

            # 8 SCRB
            add_heading(doc, "8. SCRB And TARB ")
            doc.add_paragraph("Based on the criticality, complexity and maturity level, the Chief Executive (A) would constitute System Certification Review Board as and when required. The Main contractor shall constitute Test Adequacy Review Board before major clearances, as requested by CEMILAC.")
            # 9 COMMUNICATION

            add_heading(doc, "9. Communication ")
            doc.add_paragraph(""" \n • Respective RCMAs would be responsible for on-boarding the Design Agency and the projects in e-Certification System (e-CS).
    \n • All project related technical communication between the design agency and CEMILAC would be through e-CS only.
    \n • Design agencies have to register the projects on e-CS as per the Annexure -1. Each line item is instantiated as a project on e-CS. These will be assigned internally by e-CS administrator to the designated RCMA and dealing officers. Design agencies are to use the ISO boot USB/ Crypto dongle to upload the documents project-wise on the e-CS. No documents will be accepted via email/ external storage devices/ hard copy
    \n • The inherent security features of the e-certification portal would ensure that the document is protected and is accessible only to the assigned personnel.
    \n • All the documents and certificates, and project development history would be automatically stored in CEMILAC data centre for retention and future reference. """)

            # 10 PROGRESS
            add_heading(doc, "10. Certification Progress Review [Default]")
            doc.add_paragraph("In order to ensure smooth progress of certification and provide mid-course corrections, review meetings chaired by Coordinating Director shall be conducted.\n\n(CE CEMILAC)\n OS & Chief Executive (Airworthiness)")

            # 11 DISTRIBUTION
            add_heading(doc, "11. Distribution List")
            add_heading(doc, "11.1 External Organizations")
            doc.add_paragraph(
                """\n1.Main contractor – Request to circulate to all the work centres of the project.\n   2. DG, DGAQA, New Delhi 	- Request to assign field establishments for 
  QA Coverage.\n3. AirHQ, IAF\n"""
            )  
            internal_val = getattr(self, 'InternalDistribution_dropdown', None)

            internal_text = "\n".join(internal_val.checkedItems()) if internal_val and internal_val.checkedItems() else ""
            add_heading(doc, "11.2 Internal Distribution")

            doc.add_paragraph(internal_text)
            note = self.section_notes.get("Internal Distribution", "")
            if note.strip():
                doc.add_paragraph(note)
            
            # Inject images for Scope Of Task Directive
            self.inject_images_for_heading(doc, "Internal Distribution")

            # ---------------- ANNEXURES ----------------
            doc.add_page_break()

            add_heading(doc, "Annexure-1")
            doc.add_paragraph("Work Assignment list of LRUs")

            table = doc.add_table(rows=4, cols=6)
            table.style = 'Table Grid'

            headers = ["Sl No", "Certifiable item", "Design agency", "Applicable Subpart", "Designated Directorate/ RCMA", "Dealing Officer(s)"]
            for i, h in enumerate(headers):
                table.cell(0, i).text = h
            
            # Add stored table data
            table_data = self.table_data.get("Annexure-1: Work Assignment List", [["" for _ in headers] for _ in range(3)])
            for row_idx, row_data in enumerate(table_data):
                for col_idx, cell_data in enumerate(row_data):
                    if row_idx + 1 < len(table.rows):
                        table.cell(row_idx + 1, col_idx).text = str(cell_data)

            doc.add_page_break()

            add_heading(doc, " Aircraft Integration Checks and Flight Clearance (Project Name)")

            table = doc.add_table(rows=3, cols=5)
            table.style = 'Table Grid'

            headers = ["Sl no.", "System Name", "Design agency", "Designated RCMA", "Dealing Officer"]
            for i, h in enumerate(headers):
                table.cell(0, i).text = h
            
            # Add stored table data
            table_data = self.table_data.get("Annexure-1: Integration Checks", [["" for _ in headers] for _ in range(2)])
            for row_idx, row_data in enumerate(table_data):
                for col_idx, cell_data in enumerate(row_data):
                    if row_idx + 1 < len(table.rows):
                        table.cell(row_idx + 1, col_idx).text = str(cell_data)

            doc.add_page_break()

            add_heading(doc, "Annexure-2")
            doc.add_paragraph("Contact details of dealing officers and RDs")

            table = doc.add_table(rows=3, cols=5)
            table.style = 'Table Grid'

            headers = ["Sl no.", "Name & Designation", "Certification Centre", "E-mail id", "Phone number"]
            for i, h in enumerate(headers):
                table.cell(0, i).text = h
            
            # Add stored table data
            table_data = self.table_data.get("Annexure-2: Contact Details", [["" for _ in headers] for _ in range(2)])
            for row_idx, row_data in enumerate(table_data):
                for col_idx, cell_data in enumerate(row_data):
                    if row_idx + 1 < len(table.rows):
                        table.cell(row_idx + 1, col_idx).text = str(cell_data)

            doc.add_page_break()

            add_heading(doc, "Annexure-3")
            doc.add_paragraph("Product Break Down Structure")
            
            # Inject images for Annexure – 3 (images only, no text)
            self.inject_images_for_heading(doc, "Annexure – 3")

            doc.save(save_path)
        
            pdf_path = save_path.replace(".docx", ".pdf")

            subprocess.run([
                "libreoffice",
                "--headless",
                "--convert-to",
                "pdf",
                save_path,
                "--outdir",
                os.path.dirname(save_path)
            ])

            self.generated_preview_path = pdf_path
            
            
            QMessageBox.information(
                self,
                "Success",
                "Task Directive Template created successfully!"
            )
            self.preview_pdf()

        except Exception as e:
            QMessageBox.critical(
                self,
                "Export Error",
                f"Failed to build template: {str(e)}"
            )

if __name__ == "__main__":
    app = QApplication(sys.argv)

    app.setStyle(QStyleFactory.create("Fusion"))

    palette = QPalette()
    palette.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
    palette.setColor(QPalette.ColorRole.Text, QColor("#000000"))
    palette.setColor(QPalette.ColorRole.Highlight, QColor("#e2e8f0"))
    palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#000000"))
    app.setPalette(palette)

    window = OfflineApp()
    window.show()

    sys.exit(app.exec())
