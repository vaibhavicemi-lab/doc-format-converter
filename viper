import sys
import re
import fitz  # PyMuPDF
import tempfile 
import os       

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QPushButton, QTextEdit, QFileDialog,
    QLabel, QSplitter, QMessageBox, QScrollArea, QProgressBar,
    QCheckBox, QGroupBox, QDialog, QRubberBand, QMenu, QLineEdit,
    QGridLayout, QFrame, QListWidget, QListWidgetItem, QTableWidget,
    QTableWidgetItem, QHeaderView, QAbstractItemView, QComboBox, QTabWidget
)
from PyQt6.QtGui import (
    QPixmap, QImage, QFont, QTextDocument, QTextCursor, 
    QDesktopServices, QStandardItemModel, QStandardItem, QCursor
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QEvent, QRect, QPoint, QSize, QUrl 
from PyQt6.QtPrintSupport import QPrinter

# --- NATIVE PDF VIEWER IMPORTS ---
try:
    from PyQt6.QtWebEngineWidgets import QWebEngineView
    from PyQt6.QtWebEngineCore import QWebEngineSettings, QWebEnginePage
    HAS_WEBENGINE = True

    # --- CUSTOM WEB ENGINE VIEW FOR RIGHT-CLICK COPY ---
    class PDFWebEngineView(QWebEngineView):
        def contextMenuEvent(self, event):
            menu = QMenu(self)
            menu.setStyleSheet("""
                QMenu { background-color: #ffffff; border: 1px solid #ced4da; border-radius: 4px; padding: 4px; font-size: 13px; color: #212529; }
                QMenu::item { padding: 6px 25px 6px 20px; background-color: transparent; }
                QMenu::item:selected { background-color: #f8f9fa; border-radius: 3px; color: #0d6efd; font-weight: bold;}
            """)
            
            # Force the Copy Action
            copy_action = menu.addAction("📋 Copy Selected Text")
            copy_action.triggered.connect(lambda: self.triggerPageAction(QWebEnginePage.WebAction.Copy))
            
            menu.addSeparator()
            
            reload_action = menu.addAction("🔄 Reload PDF")
            reload_action.triggered.connect(self.reload)

            menu.exec(event.globalPos())

except ImportError:
    HAS_WEBENGINE = False

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# --- INTERNAL PDF VIEWER DIALOG ---
# ==========================================
class PDFViewerDialog(QDialog):
    def __init__(self, pdf_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📄 Generated PDF Preview")
        self.resize(1000, 800)
        self.setStyleSheet("""
            QDialog { background-color: #f8f9fa; }
            QPushButton#CloseBtn { background-color: #6c757d; color: white; font-size: 14px; font-weight: bold; border-radius: 6px; padding: 10px 30px; border: none; }
            QPushButton#CloseBtn:hover { background-color: #5a6268; }
        """)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        layout.setSpacing(10)
        
        if HAS_WEBENGINE:
            self.viewer = QWebEngineView()
            self.viewer.settings().setAttribute(QWebEngineSettings.WebAttribute.PluginsEnabled, True)
            self.viewer.settings().setAttribute(QWebEngineSettings.WebAttribute.PdfViewerEnabled, True)
            self.viewer.setUrl(QUrl.fromLocalFile(os.path.abspath(pdf_path)))
            self.viewer.setStyleSheet("border: 1px solid #dee2e6; border-radius: 6px;")
            layout.addWidget(self.viewer)
        else:
            self.scroll = QScrollArea()
            self.scroll.setWidgetResizable(True)
            self.scroll.setStyleSheet("background: #e9ecef; border: 1px solid #ced4da; border-radius: 6px;")
            self.container = QWidget()
            self.container.setStyleSheet("background: transparent;")
            self.vbox = QVBoxLayout(self.container)
            self.vbox.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
            self.scroll.setWidget(self.container)
            layout.addWidget(self.scroll)
            
            try:
                doc = fitz.open(pdf_path)
                for page in doc:
                    pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                    img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()
                    lbl = QLabel()
                    lbl.setPixmap(QPixmap.fromImage(img))
                    lbl.setStyleSheet("border: 1px solid #adb5bd; margin: 10px; background: white;")
                    self.vbox.addWidget(lbl)
            except Exception as e:
                err_lbl = QLabel(f"Failed to load PDF preview:\n{e}")
                err_lbl.setStyleSheet("color: red; font-weight: bold; font-size: 14px;")
                self.vbox.addWidget(err_lbl)

        close_btn = QPushButton("Close Preview")
        close_btn.setObjectName("CloseBtn")
        close_btn.clicked.connect(self.accept)
        
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        btn_layout.addStretch()
        layout.addLayout(btn_layout)


# ==========================================
# --- UNIFIED PROJECT TABLES BUILDER ---
# ==========================================
class ProjectTablesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.parent_app = parent
        self.setWindowTitle("🗂️ Project Tables Builder")
        self.resize(850, 500)
        self.setStyleSheet("""
            QDialog { background: #f8f9fa; }
            QLabel { font-weight: bold; color: #212529; font-size: 13px;}
            QLineEdit, QComboBox { background: #ffffff; color: #212529; border: 1px solid #ced4da; border-radius: 4px; padding: 6px; }
            QLineEdit:focus, QComboBox:focus { border: 1px solid #0d6efd; }
            QTableWidget { background: #ffffff; alternate-background-color: #f8f9fa; border: 1px solid #ced4da; border-radius: 4px; }
            QHeaderView::section { background-color: #f1f3f5; font-weight: bold; border: 1px solid #ced4da; padding: 4px; color: #212529; }
            
            /* Tab Styling */
            QTabWidget::pane { border: 1px solid #dee2e6; border-radius: 4px; background: white; top: -1px; }
            QTabBar::tab { background: #f8f9fa; color: #6c757d; padding: 10px 20px; border: 1px solid #dee2e6; border-bottom: none; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 2px; font-weight: bold;}
            QTabBar::tab:selected { background: white; color: #0d6efd; border-top: 3px solid #0d6efd; border-bottom-color: white; }
            
            /* Button Styling */
            QPushButton#BlueBtn { background-color: #0d6efd; color: #ffffff; border: none; padding: 8px 15px; border-radius: 4px; font-weight: bold; font-size: 13px;}
            QPushButton#BlueBtn:hover { background-color: #0b5ed7; }
            QPushButton#RedBtn { background-color: #dc3545; color: #ffffff; border: none; padding: 8px 15px; border-radius: 4px; font-weight: bold; font-size: 13px;}
            QPushButton#RedBtn:hover { background-color: #bb2d3b; }
            QPushButton#SaveCloseBtn { background-color: #0d6efd; color: white; font-size: 14px; font-weight: bold; border-radius: 6px; border: none; }
            QPushButton#SaveCloseBtn:hover { background-color: #0b5ed7; }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(15, 15, 15, 15)
        self.tabs = QTabWidget()

        # ---------------------------------------------------------
        # TAB 1: STAKEHOLDERS TABLE
        # ---------------------------------------------------------
        tab_sh = QWidget()
        tab_sh.setStyleSheet("background: white;")
        layout_sh = QVBoxLayout(tab_sh)

        input_layout_sh = QHBoxLayout()
        self.org_combo = QComboBox()
        self.org_combo.setEditable(True) 
        self.org_combo.addItems(["CEMILAC", "RCMA", "DGAQA", "DRDO", "User", "Production Agency"])
        
        self.role_input = QLineEdit()
        self.role_input.setPlaceholderText("e.g. Certification Authority")
        
        self.act_input = QLineEdit()
        self.act_input.setPlaceholderText("e.g. Review & Approval")

        self.add_btn_sh = QPushButton("Add Row")
        self.add_btn_sh.setObjectName("BlueBtn")
        self.add_btn_sh.clicked.connect(self.add_stakeholder)

        input_layout_sh.addWidget(QLabel("Organisation:"))
        input_layout_sh.addWidget(self.org_combo, stretch=2)
        input_layout_sh.addWidget(QLabel("Role:"))
        input_layout_sh.addWidget(self.role_input, stretch=2)
        input_layout_sh.addWidget(QLabel("Activities:"))
        input_layout_sh.addWidget(self.act_input, stretch=3)
        input_layout_sh.addWidget(self.add_btn_sh)
        layout_sh.addLayout(input_layout_sh)

        self.table_sh = QTableWidget(0, 4)
        self.table_sh.setHorizontalHeaderLabels(["Sl No.", "Organisation", "Role", "Activities"])
        self.table_sh.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_sh.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents) 
        self.table_sh.setAlternatingRowColors(True)
        self.table_sh.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_sh.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout_sh.addWidget(self.table_sh)

        btn_layout_sh = QHBoxLayout()
        self.del_btn_sh = QPushButton("Delete Selected Row")
        self.del_btn_sh.setObjectName("RedBtn")
        self.del_btn_sh.clicked.connect(self.delete_selected_sh)
        
        self.clear_btn_sh = QPushButton("Clear All")
        self.clear_btn_sh.setObjectName("RedBtn")
        self.clear_btn_sh.clicked.connect(self.clear_all_sh)
        
        btn_layout_sh.addWidget(self.del_btn_sh)
        btn_layout_sh.addWidget(self.clear_btn_sh)
        btn_layout_sh.addStretch()
        layout_sh.addLayout(btn_layout_sh)

        self.tabs.addTab(tab_sh, "👥 Stakeholders Table")

        # ---------------------------------------------------------
        # TAB 2: CERTIFICATION TASK ALLOCATION TABLE
        # ---------------------------------------------------------
        tab_ct = QWidget()
        tab_ct.setStyleSheet("background: white;")
        layout_ct = QVBoxLayout(tab_ct)

        input_layout_ct = QHBoxLayout()
        
        self.act_input_ct = QLineEdit()
        self.act_input_ct.setPlaceholderText("e.g. Ground Testing")
        
        self.centre_combo = QComboBox()
        self.centre_combo.setEditable(True) 
        self.centre_combo.addItems(["CEMILAC", "RCMA", "DGAQA", "DRDO", "User", "Production Agency"])
        
        self.head_input = QLineEdit()
        self.head_input.setPlaceholderText("e.g. Director, RCMA")

        self.add_btn_ct = QPushButton("Add Row")
        self.add_btn_ct.setObjectName("BlueBtn")
        self.add_btn_ct.clicked.connect(self.add_task)

        input_layout_ct.addWidget(QLabel("Activity:"))
        input_layout_ct.addWidget(self.act_input_ct, stretch=2)
        input_layout_ct.addWidget(QLabel("Work Centre:"))
        input_layout_ct.addWidget(self.centre_combo, stretch=2)
        input_layout_ct.addWidget(QLabel("Resp. Head:"))
        input_layout_ct.addWidget(self.head_input, stretch=2)
        input_layout_ct.addWidget(self.add_btn_ct)
        layout_ct.addLayout(input_layout_ct)

        self.table_ct = QTableWidget(0, 4)
        self.table_ct.setHorizontalHeaderLabels(["Sl No.", "Certification Activity", "Certification Work Centre", "Responsible Head"])
        self.table_ct.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_ct.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents) 
        self.table_ct.setAlternatingRowColors(True)
        self.table_ct.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table_ct.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        layout_ct.addWidget(self.table_ct)

        btn_layout_ct = QHBoxLayout()
        self.del_btn_ct = QPushButton("Delete Selected Row")
        self.del_btn_ct.setObjectName("RedBtn")
        self.del_btn_ct.clicked.connect(self.delete_selected_ct)
        
        self.clear_btn_ct = QPushButton("Clear All")
        self.clear_btn_ct.setObjectName("RedBtn")
        self.clear_btn_ct.clicked.connect(self.clear_all_ct)

        btn_layout_ct.addWidget(self.del_btn_ct)
        btn_layout_ct.addWidget(self.clear_btn_ct)
        btn_layout_ct.addStretch()
        layout_ct.addLayout(btn_layout_ct)

        self.tabs.addTab(tab_ct, "🛠️ Cert Task Allocation Table")

        # ---------------------------------------------------------
        # MAIN DIALOG BUTTONS
        # ---------------------------------------------------------
        layout.addWidget(self.tabs)
        
        self.close_btn = QPushButton("Save & Close")
        self.close_btn.setObjectName("SaveCloseBtn")
        self.close_btn.setFixedHeight(40)
        self.close_btn.clicked.connect(self.accept)
        layout.addWidget(self.close_btn)

        self.refresh_sh_table()
        self.refresh_ct_table()

    # --- Stakeholder Methods ---
    def add_stakeholder(self):
        org = self.org_combo.currentText().strip()
        role = self.role_input.text().strip()
        act = self.act_input.text().strip()
        if org:
            self.parent_app.stakeholders_data.append({"org": org, "role": role, "activities": act})
            self.role_input.clear()
            self.act_input.clear()
            self.refresh_sh_table()

    def delete_selected_sh(self):
        selected = self.table_sh.currentRow()
        if selected >= 0:
            self.parent_app.stakeholders_data.pop(selected)
            self.refresh_sh_table()

    def clear_all_sh(self):
        self.parent_app.stakeholders_data.clear()
        self.refresh_sh_table()

    def refresh_sh_table(self):
        self.table_sh.setRowCount(0)
        for i, data in enumerate(self.parent_app.stakeholders_data):
            self.table_sh.insertRow(i)
            self.table_sh.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.table_sh.setItem(i, 1, QTableWidgetItem(data['org']))
            self.table_sh.setItem(i, 2, QTableWidgetItem(data['role']))
            self.table_sh.setItem(i, 3, QTableWidgetItem(data['activities']))

    # --- Cert Task Methods ---
    def add_task(self):
        act = self.act_input_ct.text().strip()
        centre = self.centre_combo.currentText().strip()
        head = self.head_input.text().strip()
        if act or centre:
            self.parent_app.cert_task_data.append({"activity": act, "centre": centre, "head": head})
            self.act_input_ct.clear()
            self.head_input.clear()
            self.refresh_ct_table()

    def delete_selected_ct(self):
        selected = self.table_ct.currentRow()
        if selected >= 0:
            self.parent_app.cert_task_data.pop(selected)
            self.refresh_ct_table()

    def clear_all_ct(self):
        self.parent_app.cert_task_data.clear()
        self.refresh_ct_table()

    def refresh_ct_table(self):
        self.table_ct.setRowCount(0)
        for i, data in enumerate(self.parent_app.cert_task_data):
            self.table_ct.insertRow(i)
            self.table_ct.setItem(i, 0, QTableWidgetItem(str(i + 1)))
            self.table_ct.setItem(i, 1, QTableWidgetItem(data['activity']))
            self.table_ct.setItem(i, 2, QTableWidgetItem(data['centre']))
            self.table_ct.setItem(i, 3, QTableWidgetItem(data['head']))


# ==========================================
# --- CUSTOM UI: INLINE DROPDOWN WITH [+] ---
# ==========================================
class DropdownPopup(QDialog):
    def __init__(self, parent_btn, options):
        super().__init__(parent_btn, Qt.WindowType.Popup | Qt.WindowType.FramelessWindowHint)
        self.parent_btn = parent_btn
        self.setStyleSheet("""
            QDialog { background: #ffffff; border: 1px solid #ced4da; border-radius: 6px; }
            QLabel { color: #6c757d; font-size: 11px; font-weight: bold; }
            QListWidget { background: transparent; border: none; color: #212529; outline: none; font-size: 13px; }
            QListWidget::item { padding: 6px; border-radius: 4px; color: #212529; }
            QListWidget::item:hover { background: #f8f9fa; color: #212529; }
            QListWidget::item:selected { background: #e9ecef; color: #212529; }
            QListWidget::indicator { width: 16px; height: 16px; border: 1px solid #ced4da; border-radius: 4px; background: #ffffff; }
            QListWidget::indicator:checked { background: #0d6efd; border: 1px solid #0d6efd; }
            QLineEdit { background: #ffffff; color: #212529; border: 1px solid #0d6efd; border-radius: 4px; padding: 4px; }
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)

        header_layout = QHBoxLayout()
        header_layout.addWidget(QLabel("SELECT OPTIONS"))
        header_layout.addStretch()
        
        self.add_btn = QPushButton("Add")
        self.add_btn.setFixedSize(40, 24)
        self.add_btn.setStyleSheet("""
            QPushButton { background: #f8f9fa; color: #212529; border: 1px solid #ced4da; border-radius: 4px; font-weight: bold; font-size: 12px; }
            QPushButton:hover { background: #e9ecef; }
        """)
        self.add_btn.clicked.connect(self.show_new_input)
        header_layout.addWidget(self.add_btn)
        layout.addLayout(header_layout)

        self.new_input = QLineEdit()
        self.new_input.setPlaceholderText("") 
        self.new_input.hide()
        self.new_input.returnPressed.connect(self.save_new_option)
        layout.addWidget(self.new_input)

        self.list_widget = QListWidget()
        for opt in options:
            if opt != "Enter Other...":
                self.add_item(opt)
                
        self.list_widget.itemPressed.connect(self.on_item_pressed)
        self.list_widget.itemChanged.connect(self.parent_btn.update_display)
        layout.addWidget(self.list_widget)

    def on_item_pressed(self, item):
        pos = self.list_widget.viewport().mapFromGlobal(QCursor.pos())
        if pos.x() > 25:
            state = item.checkState()
            item.setCheckState(Qt.CheckState.Unchecked if state == Qt.CheckState.Checked else Qt.CheckState.Checked)

    def add_item(self, text, checked=False):
        item = QListWidgetItem(text)
        item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        self.list_widget.addItem(item)

    def show_new_input(self):
        self.new_input.show()
        self.new_input.setFocus()

    def save_new_option(self):
        text = self.new_input.text().strip()
        if text:
            existing = [self.list_widget.item(i).text() for i in range(self.list_widget.count())]
            if text not in existing:
                self.add_item(text, checked=True)
            else:
                for i in range(self.list_widget.count()):
                    if self.list_widget.item(i).text() == text:
                        self.list_widget.item(i).setCheckState(Qt.CheckState.Checked)
        self.new_input.clear()
        self.new_input.hide()
        self.parent_btn.update_display()

class PopupMultiSelect(QPushButton):
    def __init__(self, options, parent=None):
        super().__init__("Select options...", parent)
        self.setStyleSheet("""
            QPushButton {
                text-align: left; padding: 10px; background: #ffffff; 
                border: 1px solid #ced4da; border-radius: 6px; color: #212529; font-size: 13px;
            }
            QPushButton:hover { border: 1px solid #0d6efd; }
        """)
        self.options = options
        self.popup = None

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            if not self.popup:
                self.popup = DropdownPopup(self, self.options)
            
            pos = self.mapToGlobal(self.rect().bottomLeft())
            self.popup.setFixedWidth(self.width())
            self.popup.move(pos.x(), pos.y() + 2)
            self.popup.show()

    def update_display(self):
        checked = self.checkedItems()
        if checked:
            self.setText(", ".join(checked))
        else:
            self.setText("Select options...")

    def checkedItems(self):
        if not self.popup: return []
        items = []
        for i in range(self.popup.list_widget.count()):
            item = self.popup.list_widget.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                items.append(item.text())
        return items

# ==========================================
# --- CUSTOM UI: COLLAPSIBLE PANEL ---
# ==========================================
class CollapsibleBox(QWidget):
    def __init__(self, title="", parent=None):
        super().__init__(parent)
        self.layout = QVBoxLayout(self)
        self.layout.setContentsMargins(0, 0, 0, 0)
        self.layout.setSpacing(0)

        self.toggle_btn = QPushButton(f"▼  {title}")
        self.toggle_btn.setStyleSheet("""
            QPushButton {
                text-align: left; font-weight: bold; background: #f8f9fa; 
                color: #0d6efd; padding: 12px; font-size: 14px; 
                border-top-left-radius: 8px; border-top-right-radius: 8px;
                border: 1px solid #dee2e6; border-bottom: none;
            }
            QPushButton:hover { background: #e9ecef; }
        """)
        self.toggle_btn.clicked.connect(self.on_press)
        self.layout.addWidget(self.toggle_btn)

        self.content_area = QFrame()
        self.content_area.setStyleSheet("""
            QFrame { background: #ffffff; border: 1px solid #dee2e6; 
            border-bottom-left-radius: 8px; border-bottom-right-radius: 8px; }
        """)
        self.content_layout = QVBoxLayout(self.content_area)
        self.content_layout.setContentsMargins(15, 15, 15, 15)
        self.content_layout.setSpacing(15)
        self.layout.addWidget(self.content_area)
        
        self.is_expanded = True

    def on_press(self):
        self.is_expanded = not self.is_expanded
        if self.is_expanded:
            self.toggle_btn.setText(self.toggle_btn.text().replace("▶", "▼"))
            self.toggle_btn.setStyleSheet(self.toggle_btn.styleSheet().replace("border-radius: 8px;", "border-top-left-radius: 8px; border-top-right-radius: 8px;"))
            self.content_area.show()
        else:
            self.toggle_btn.setText(self.toggle_btn.text().replace("▼", "▶"))
            self.toggle_btn.setStyleSheet(self.toggle_btn.styleSheet().replace("border-top-left-radius: 8px; border-top-right-radius: 8px;", "border-radius: 8px; border-bottom: 1px solid #dee2e6;"))
            self.content_area.hide()

# ==========================================
# --- PURE PYTHON SUMMARIZER ---
# ==========================================
def simple_summarize(text, target_ratio=0.4, min_sentences=3, max_sentences=10):
    if not text or len(text.strip()) < 20: return text.strip()
    clean_text = re.sub(r'(?<!\n)\n(?!\n)', ' ', text)
    sentences = re.split(r'(?<=[.!?])\s+', clean_text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 10]
    if not sentences: return ""
    num_sentences = int(len(sentences) * target_ratio)
    num_sentences = max(min_sentences, min(num_sentences, max_sentences))
    if len(sentences) <= num_sentences: return " ".join(sentences)
    stop_words = {"the", "is", "in", "and", "to", "of", "a", "for", "on", "with", "as", "by", "this", "that", "it", "are", "be", "or", "an", "at", "from", "which", "will"}
    words = re.findall(r'\b[a-zA-Z]{2,}\b', clean_text.lower())
    freq = {}
    for w in words:
        if w not in stop_words: freq[w] = freq.get(w, 0) + 1
    max_freq = max(freq.values()) if freq else 1
    for w in freq: freq[w] = freq[w] / max_freq
    scores = {}
    for i, s in enumerate(sentences):
        score = 0
        s_words = re.findall(r'\b[a-zA-Z]{2,}\b', s.lower())
        for w in s_words:
            if w in freq: score += freq[w]
        score = score / max(len(s_words), 1)
        if i < 2: score += 0.5
        scores[i] = score
    top_indices = sorted(sorted(scores, key=scores.get, reverse=True)[:num_sentences])
    return " ".join([sentences[i] for i in top_indices])

# ==========================================
# --- SMART TEXT EDITOR (SUPPORTS IMAGES)---
# ==========================================
class NotesEditor(QTextEdit):
    def canInsertFromMimeData(self, source):
        if source.hasImage() or source.hasText(): return True
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
# --- SELECTABLE TEXT PREVIEW TAB ---
# ==========================================
class SelectableTextPreview(QTextEdit):
    text_extracted = pyqtSignal(str)
    summary_requested = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setStyleSheet("""
            QTextEdit { background: #ffffff; color: #212529; font-size: 14px; 
            line-height: 1.6; padding: 15px; border: none; }
        """)

    def contextMenuEvent(self, event):
        menu = self.createStandardContextMenu()
        if self.textCursor().hasSelection():
            menu.addSeparator()
            extract_action = menu.addAction("📝 Extract Selected Text to Notes")
            summarize_action = menu.addAction("✨ Summarize Selection (Set as Intro)")
            
            action = menu.exec(event.globalPos())
            
            if action == extract_action:
                self.text_extracted.emit(self.textCursor().selectedText())
            elif action == summarize_action:
                self.summary_requested.emit(self.textCursor().selectedText())
        else:
            menu.exec(event.globalPos())

# ==========================================
# --- PURE PYTHON ANALYSIS WORKER ---
# ==========================================
class AnalysisWorker(QThread):
    finished = pyqtSignal(dict) 
    error = pyqtSignal(str)

    def __init__(self, text, topics):
        super().__init__()
        self.text = text
        self.topics = topics

    def run(self):
        try:
            results = {"project_name": "", "introduction": "", "found_topics": []}
            
            explicit_match = re.search(r'(?:Project\s*Name|Project\s*Title|Title|Subject)[\s:]*(.+)', self.text, re.IGNORECASE)
            if explicit_match and not explicit_match.group(1).strip().startswith("___"):
                results["project_name"] = explicit_match.group(1).strip()[:100]
                
            if not results["project_name"]:
                semantic_match = re.search(r'(?:project|report|proposal) (?:aims to|focuses on|proposes|is to|investigates) ([^\.]+)', self.text, re.IGNORECASE)
                if semantic_match:
                    results["project_name"] = semantic_match.group(1).strip().title()[:100]

            if not results["project_name"]:
                lines = [line.strip() for line in self.text.split('\n') if line.strip()]
                for line in lines[:10]:
                    if re.search(r'(Document Analysis Report|Summary|Page|Date|Author)', line, re.IGNORECASE): continue
                    if 10 < len(line) < 100:
                        results["project_name"] = line
                        break

            results["introduction"] = simple_summarize(self.text)
            
            for topic in self.topics:
                if topic.lower() in self.text.lower(): results["found_topics"].append(topic)
            if results["introduction"] and "Introduction" not in results["found_topics"]:
                results["found_topics"].append("Introduction")
                
            self.finished.emit(results)
        except Exception as e:
            self.error.emit(str(e))

# ==========================================
# --- INTERACTIVE PDF SNIPPING & PANNING ---
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
        self.rubber_band.hide() 
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu { background-color: #ffffff; border: 1px solid #ced4da; border-radius: 4px; padding: 4px; font-size: 13px; color: #212529; }
            QMenu::item { padding: 6px 25px 6px 20px; background-color: transparent; }
            QMenu::item:selected { background-color: #f8f9fa; border-radius: 3px; }
        """)
        copy_img_action = menu.addAction("🖼️ Extract as Image")
        copy_text_action = menu.addAction("📝 Extract Text to Notes")
        menu.addSeparator()
        summarize_action = menu.addAction("✨ Summarize Selection (Set as Intro)")
        action = menu.exec(self.mapToGlobal(pos))
        
        if action == copy_img_action:
            if self.pixmap(): self.image_extracted.emit(self.pixmap().copy(rect)) 
        elif action == copy_text_action:
            text = self.get_text_from_rect(rect)
            if text: self.text_extracted.emit(text) 
        elif action == summarize_action:
            text = self.get_text_from_rect(rect)
            if text: self.summary_requested.emit(text)

# ==========================================
# --- MAIN APPLICATION ---
# ==========================================
class OfflineApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Document Extraction Assistant")
        self.resize(1450, 950) 
        
        self.file_path = ""
        self.doc = None
        self.page_data = [] 
        self.current_zoom = 600 
        self.is_maximized = False
        
        # Data Variables
        self.extracted_project_name = "" 
        self.intro_text_data = "" 
        self.stakeholders_data = []
        self.cert_task_data = [] 
        
        self.apply_light_theme()
        
        # Initialize Popups
        self.init_data_dialog()
        self.init_tables_dialog()
        self.init_ui()

    def apply_light_theme(self):
        self.setStyleSheet("""
            QMainWindow, QWidget#mainWidget, QMessageBox { 
                background-color: #f8f9fa; color: #212529; font-family: 'Segoe UI', sans-serif; font-size: 13px;
            }
            QFrame[class="Card"] { background: #ffffff; border-radius: 8px; border: 1px solid #dee2e6; }
            QGroupBox { border: 1px solid #dee2e6; border-radius: 8px; margin-top: 15px; font-weight: bold; background: #ffffff;}
            QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; padding: 0 5px; color: #0d6efd; }
            QPushButton#PrimaryBtn { background-color: #0d6efd; color: #ffffff; border: none; padding: 8px 15px; border-radius: 4px; font-weight: bold; }
            QPushButton#PrimaryBtn:hover { background-color: #0b5ed7; }
            QPushButton#SecondaryBtn { background-color: #ffffff; color: #212529; border: 1px solid #ced4da; padding: 8px 15px; border-radius: 4px; font-weight: bold; }
            QPushButton#SecondaryBtn:hover { background-color: #f8f9fa; }
            QPushButton#ToolbarBtn { background-color: #ffffff; border: 1px solid #ced4da; border-radius: 22px; font-size: 20px; }
            QPushButton#ToolbarBtn:hover { background-color: #e9ecef; }
            QLineEdit, QTextEdit { background: #ffffff; color: #212529; border: 1px solid #ced4da; border-radius: 6px; padding: 8px; }
            QLineEdit:focus, QTextEdit:focus { border: 1px solid #0d6efd; }
            QLabel { color: #343a40; }
            QCheckBox { color: #212529; spacing: 8px; }
            QCheckBox::indicator { width: 16px; height: 16px; border: 1px solid #ced4da; border-radius: 4px; background: #ffffff; }
            QCheckBox::indicator:checked { background: #0d6efd; border: 1px solid #0d6efd; }
            QScrollBar:vertical { border: none; background: #f8f9fa; width: 10px; margin: 0px; }
            QScrollBar::handle:vertical { background: #ced4da; min-height: 20px; border-radius: 5px; }
            QScrollBar::handle:vertical:hover { background: #adb5bd; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { border: none; background: none; }
            QSplitter::handle { background: #dee2e6; width: 4px; }
        """)

    def init_ui(self):
        main_widget = QWidget()
        main_widget.setObjectName("mainWidget")
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(10, 10, 10, 10)
        
        self.splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- LEFT SIDE: Document Viewer ---
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(10)

        controls_layout = QHBoxLayout()
        self.upload_btn = QPushButton("📂 Upload PDF")
        self.upload_btn.setObjectName("PrimaryBtn")
        self.upload_btn.clicked.connect(self.upload_file)
        controls_layout.addWidget(self.upload_btn)
        
        controls_layout.addStretch()
        
        self.paste_notes_btn = QPushButton("📋 Paste to Notes")
        self.paste_notes_btn.setObjectName("SecondaryBtn")
        self.paste_notes_btn.clicked.connect(self.paste_to_notes)
        controls_layout.addWidget(self.paste_notes_btn)
        
        self.summarize_btn = QPushButton("✨ Summarize Copied Text")
        self.summarize_btn.setObjectName("SecondaryBtn")
        self.summarize_btn.clicked.connect(self.summarize_clipboard)
        controls_layout.addWidget(self.summarize_btn)

        left_layout.addLayout(controls_layout)

        self.left_tabs = QTabWidget()
        self.left_tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #dee2e6; border-radius: 6px; background: white; top:-1px; }
            QTabBar::tab { background: #f8f9fa; color: #6c757d; padding: 10px 20px; border: 1px solid #dee2e6; border-bottom: none; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 2px; font-weight: bold;}
            QTabBar::tab:selected { background: white; color: #0d6efd; border-top: 3px solid #0d6efd; border-bottom-color: white;}
        """)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background: #f8f9fa; border: none;") 
        self.page_container = QWidget()
        self.page_container.setStyleSheet("background: transparent;")
        self.page_layout = QVBoxLayout(self.page_container)
        self.page_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter) 
        self.scroll_area.setWidget(self.page_container)
        self.scroll_area.viewport().installEventFilter(self)

        if HAS_WEBENGINE:
            self.pdf_viewer = PDFWebEngineView(self)
            self.pdf_viewer.settings().setAttribute(QWebEngineSettings.WebAttribute.PluginsEnabled, True)
            self.pdf_viewer.settings().setAttribute(QWebEngineSettings.WebAttribute.PdfViewerEnabled, True)
            self.pdf_viewer.setStyleSheet("border: none;")
            self.left_tabs.addTab(self.pdf_viewer, "📝 Native PDF Viewer (Text Selection)")
            self.left_tabs.addTab(self.scroll_area, "🖼️ Visual Image Extractor")
        else:
            self.text_preview = SelectableTextPreview()
            self.text_preview.text_extracted.connect(self.add_extracted_text)
            self.text_preview.summary_requested.connect(self.generate_summary_from_selection)
            self.left_tabs.addTab(self.scroll_area, "🖼️ Visual Image Extractor")
            self.left_tabs.addTab(self.text_preview, "📝 Raw Text Viewer")
        
        left_layout.addWidget(self.left_tabs)

        # --- RIGHT SIDE: IDE Tools ---
        self.right_container = QWidget() 
        right_layout = QVBoxLayout(self.right_container)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(10)

        icon_layout = QHBoxLayout()
        icon_layout.addStretch()
        
        icon_lbl = QLabel("📄 Document Assistant  | ")
        icon_lbl.setStyleSheet("font-size: 16px; font-weight: bold; color: #6c757d; margin-right: 10px;")
        icon_layout.addWidget(icon_lbl)
        
        self.tables_icon_btn = QPushButton("🗂️")
        self.tables_icon_btn.setFixedSize(45, 45)
        self.tables_icon_btn.setToolTip("Manage Project Tables")
        self.tables_icon_btn.setObjectName("ToolbarBtn")
        self.tables_icon_btn.clicked.connect(self.show_tables_popup)
        icon_layout.addWidget(self.tables_icon_btn)
        
        self.data_icon_btn = QPushButton("📝")
        self.data_icon_btn.setFixedSize(45, 45)
        self.data_icon_btn.setToolTip("View/Edit Project Data & Notes")
        self.data_icon_btn.setObjectName("ToolbarBtn")
        self.data_icon_btn.clicked.connect(lambda: self.show_data_popup(0))
        icon_layout.addWidget(self.data_icon_btn)
        
        right_layout.addLayout(icon_layout)

        self.status_group = QGroupBox("Extraction Checklist")
        status_main_layout = QVBoxLayout(self.status_group)
        status_main_layout.setContentsMargins(10, 15, 10, 5) 
        
        status_header = QHBoxLayout()
        self.status_label = QLabel("Waiting for document...")
        self.status_label.setStyleSheet("color: #6c757d; font-style: italic;")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0) 
        self.progress_bar.hide()
        status_header.addWidget(self.status_label)
        status_header.addWidget(self.progress_bar) 
        status_main_layout.addLayout(status_header)
        
        self.topics = [
            "Introduction", "Reference", "Basis Of Task Directive",
            "Scope", "Stakeholders", "Cert Breakdown",
            "Task Allocation", "Coordinating Dir.", "SPoC",
            "Cert Allocation", "Clearance Issue", "SCRB & TARB",
            "Communication", "Review", "Distribution List"
        ]
        self.topic_checkboxes = {}
        
        topics_widget = QWidget()
        topics_layout = QGridLayout(topics_widget) 
        topics_layout.setContentsMargins(0, 5, 0, 0)
        topics_layout.setSpacing(4)
        
        row, col = 0, 0
        for topic in self.topics:
            cb = QCheckBox(topic)
            cb.setEnabled(False) 
            self.topic_checkboxes[topic] = cb
            topics_layout.addWidget(cb, row, col)
            col += 1
            if col > 2:  
                col = 0
                row += 1

        status_main_layout.addWidget(topics_widget)
        right_layout.addWidget(self.status_group)

        self.collapsible_params = CollapsibleBox("ADDITIONAL TASK PARAMETERS")
        
        dropdowns_data = [
            ("Scope Of Task Directive", ["Option 1", "Option 2", "Option 3"]),
            ("Certification Work Breakdown", ["Type A", "Type B", "Type C"]),
            ("Task Allocation", ["Auto", "Manual", "Hybrid"]),
            ("Communication Type", ["Email", "Meeting", "Report"])
        ]
        
        self.combos = {}
        
        param_scroll = QScrollArea()
        param_scroll.setWidgetResizable(True)
        param_scroll.setStyleSheet("border: none; background: transparent;")
        param_widget = QWidget()
        param_layout = QVBoxLayout(param_widget)
        param_layout.setContentsMargins(0, 0, 0, 0)
        param_layout.setSpacing(10)
        
        for label_text, items in dropdowns_data:
            row = QVBoxLayout()
            row.setSpacing(2)
            lbl = QLabel(label_text.upper())
            lbl.setStyleSheet("font-size: 11px; font-weight: bold; color: #6c757d;")
            combo = PopupMultiSelect(items)
            self.combos[label_text] = combo
            row.addWidget(lbl)
            row.addWidget(combo)
            param_layout.addLayout(row)
            
        param_layout.addStretch()
        param_scroll.setWidget(param_widget)
        self.collapsible_params.content_layout.addWidget(param_scroll)
        
        right_layout.addWidget(self.collapsible_params, stretch=1)

        actions_layout = QHBoxLayout()

        self.save_pdf_btn = QPushButton("💾 View Basic PDF")
        self.save_pdf_btn.setObjectName("SecondaryBtn")
        self.save_pdf_btn.setFixedHeight(45)
        self.save_pdf_btn.clicked.connect(self.export_pdf)
        actions_layout.addWidget(self.save_pdf_btn)

        self.save_docx_btn = QPushButton("Generate Word Doc")
        self.save_docx_btn.setObjectName("PrimaryBtn")
        self.save_docx_btn.setFixedHeight(45)
        self.save_docx_btn.clicked.connect(self.export_docx)
        actions_layout.addWidget(self.save_docx_btn, stretch=1)

        right_layout.addLayout(actions_layout)

        self.splitter.addWidget(left_container)
        self.splitter.addWidget(self.right_container)
        self.splitter.setStretchFactor(0, 5) 
        self.splitter.setStretchFactor(1, 4) 
        layout.addWidget(self.splitter)

    # ------------------------------------------
    # --- EDITABLE POPUPS ---
    # ------------------------------------------
    def init_data_dialog(self):
        self.data_dialog = QDialog(self)
        self.data_dialog.setWindowTitle("🗂️ Project Data & Notes")
        self.data_dialog.resize(750, 650)
        
        dialog_layout = QVBoxLayout(self.data_dialog)
        dialog_layout.setContentsMargins(15, 15, 15, 15)
        dialog_layout.setSpacing(15)

        self.data_tabs = QTabWidget()
        self.data_dialog.setStyleSheet("""
            QDialog { background-color: #f8f9fa; }
            QLabel { font-weight: bold; color: #343a40; font-size: 13px; margin-top: 5px; }
            QLineEdit, QTextEdit { 
                background-color: #ffffff; 
                border: 1px solid #ced4da; 
                border-radius: 6px; 
                padding: 10px; 
                font-size: 14px;
                color: #212529;
            }
            QLineEdit:focus, QTextEdit:focus { border: 1px solid #0d6efd; }
            QTabWidget::pane { border: 1px solid #dee2e6; background: #ffffff; border-radius: 4px; top: -1px; }
            QTabBar::tab { background: #f8f9fa; color: #6c757d; padding: 10px 20px; border: 1px solid #dee2e6; border-bottom: none; border-top-left-radius: 6px; border-top-right-radius: 6px; margin-right: 2px; font-weight: bold; }
            QTabBar::tab:selected { background: #ffffff; color: #0d6efd; border-top: 3px solid #0d6efd; border-bottom-color: white; }
            QPushButton#SaveCloseBtn { background-color: #0d6efd; color: white; font-size: 15px; font-weight: bold; border-radius: 6px; border: none; }
            QPushButton#SaveCloseBtn:hover { background-color: #0b5ed7; }
        """)
        
        summary_tab = QWidget()
        summary_tab.setStyleSheet("background: white;")
        summary_layout = QVBoxLayout(summary_tab)
        summary_layout.setContentsMargins(20, 20, 20, 20)
        
        name_lbl = QLabel("Project Name:")
        summary_layout.addWidget(name_lbl)

        self.popup_name_edit = QLineEdit()
        self.popup_name_edit.textChanged.connect(lambda text: setattr(self, 'extracted_project_name', text))
        summary_layout.addWidget(self.popup_name_edit)

        sum_lbl = QLabel("Introduction Summary:")
        summary_layout.addWidget(sum_lbl)

        self.popup_summary_edit = QTextEdit()
        self.popup_summary_edit.textChanged.connect(lambda: setattr(self, 'intro_text_data', self.popup_summary_edit.toPlainText()))
        summary_layout.addWidget(self.popup_summary_edit)
        
        notes_tab = QWidget()
        notes_tab.setStyleSheet("background: white;")
        notes_layout = QVBoxLayout(notes_tab)
        notes_layout.setContentsMargins(20, 20, 20, 20)
        
        self.manual_input = NotesEditor()
        self.manual_input.setPlaceholderText("Paste text, images, and flowcharts here...")
        notes_layout.addWidget(self.manual_input)
        
        self.data_tabs.addTab(summary_tab, "📑 Summary")
        self.data_tabs.addTab(notes_tab, "📝 Manual Notes")
        
        dialog_layout.addWidget(self.data_tabs)
        
        close_btn = QPushButton("Save & Close")
        close_btn.setObjectName("SaveCloseBtn")
        close_btn.setFixedHeight(45)
        close_btn.clicked.connect(self.data_dialog.hide)
        dialog_layout.addWidget(close_btn)

    def show_data_popup(self, tab_index=0):
        self.popup_name_edit.setText(self.extracted_project_name)
        self.popup_summary_edit.setText(self.intro_text_data)
        self.data_tabs.setCurrentIndex(tab_index)
        self.data_dialog.show()
        self.data_dialog.raise_()
        self.data_dialog.activateWindow()

    def init_tables_dialog(self):
        self.tables_dialog = ProjectTablesDialog(self)

    def show_tables_popup(self):
        self.tables_dialog.show()
        self.tables_dialog.raise_()
        self.tables_dialog.activateWindow()

    # ------------------------------------------
    # --- AUTO-PASTING FUNCTIONS ---
    # ------------------------------------------
    def flash_data_icon(self):
        self.data_icon_btn.setStyleSheet("""
            QPushButton { background-color: #198754; color: #FFFFFF; border: none; border-radius: 22px; font-weight: bold; font-size: 20px;}
        """)

    def paste_to_notes(self):
        clipboard_text = QApplication.clipboard().text()
        if clipboard_text.strip():
            self.manual_input.append(clipboard_text + "\n")
            self.flash_data_icon()
            QMessageBox.information(self, "Success", "Copied text added to Manual Notes!")
        else:
            QMessageBox.warning(self, "Empty", "Please copy some text from the PDF first (Ctrl+C).")
            
    def summarize_clipboard(self):
        clipboard_text = QApplication.clipboard().text()
        if clipboard_text.strip():
            self.generate_summary_from_selection(clipboard_text)
        else:
            QMessageBox.warning(self, "Empty", "Please copy some text from the PDF first (Ctrl+C).")

    def generate_summary_from_selection(self, text):
        self.status_label.setText("Summarizing selection...")
        QApplication.processEvents()
        
        summary = simple_summarize(text)
        if summary:
            self.intro_text_data = summary 
            self.popup_summary_edit.setText(summary)
            self.flash_data_icon()
            self.status_label.setText("Ready")
            self.topic_checkboxes["Introduction"].setChecked(True)
            QMessageBox.information(self, "Success", "Selection summarized! Click the 📝 icon to view or edit it.")
        else:
            self.status_label.setText("Ready")
            QMessageBox.warning(self, "Empty", "Could not generate a summary. Try copying a larger paragraph.")
            
    def add_extracted_text(self, text):
        QApplication.clipboard().setText(text) 
        self.manual_input.append(text + "\n")
        self.flash_data_icon()
        
    def add_extracted_image(self, pixmap):
        if pixmap.width() > 500:
            pixmap = pixmap.scaledToWidth(500, Qt.TransformationMode.SmoothTransformation)
        QApplication.clipboard().setPixmap(pixmap)
        cursor = self.manual_input.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End) 
        self.manual_input.setTextCursor(cursor)
        cursor.insertImage(pixmap.toImage()) 
        self.manual_input.append("\n") 
        self.flash_data_icon()

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
                lbl.setStyleSheet("background-color: white; border: 1px solid #ced4da; margin-bottom: 10px; border-radius: 4px;")
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
    def upload_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Document", "", "Documents (*.pdf *.txt)")
        if path:
            self.file_path = path
            
            for cb in self.topic_checkboxes.values(): cb.setChecked(False)
                
            self.extracted_project_name = ""
            self.intro_text_data = "" 
            
            default_icon_style = """
                QPushButton { font-size: 20px; border-radius: 22px; background-color: #ffffff; border: 1px solid #ced4da; }
                QPushButton:hover { background-color: #e9ecef; }
            """
            self.data_icon_btn.setStyleSheet(default_icon_style)
            
            self.manual_input.clear() 
            
            try:
                full_text = ""
                if path.lower().endswith('.pdf'):
                    if HAS_WEBENGINE:
                        self.pdf_viewer.setUrl(QUrl.fromLocalFile(os.path.abspath(path)))
                        self.left_tabs.setCurrentIndex(0)
                    else:
                        self.left_tabs.setCurrentIndex(1)
                        
                    self.status_label.setText("Extracting backend data...")
                    QApplication.processEvents()
                    self.doc = fitz.open(path)
                    full_text = "".join(page.get_text() for page in self.doc)
                    
                    self.page_data.clear()
                    for i, page in enumerate(self.doc):
                        if i < 15: 
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()
                            self.page_data.append((QPixmap.fromImage(img), i))
                    
                    self.current_zoom = 600 
                    self.refresh_pdf_view()
                    
                else:
                    if HAS_WEBENGINE: self.left_tabs.setCurrentIndex(2)
                    else: self.left_tabs.setCurrentIndex(1)
                    
                    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                        full_text = f.read()
                        if hasattr(self, 'text_preview'):
                            self.text_preview.setText(full_text)

                self.start_analysis_thread(full_text)
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not load file: {str(e)}")

    def start_analysis_thread(self, text):
        self.status_label.setText("Scanning document with Python NLP...")
        self.progress_bar.show()
        self.worker = AnalysisWorker(text, self.topics)
        self.worker.finished.connect(self.handle_analysis_done)
        self.worker.error.connect(lambda e: self.status_label.setText(f"Analysis Error: {e}"))
        self.worker.start()

    def handle_analysis_done(self, results):
        self.progress_bar.hide()
        self.status_label.setText("Analysis Complete")
        
        if results["project_name"]:
            self.extracted_project_name = results["project_name"]
            
        if results["introduction"]:
            self.intro_text_data = results["introduction"]
            self.flash_data_icon()
            
        for topic in results["found_topics"]:
            if topic in self.topic_checkboxes:
                self.topic_checkboxes[topic].setChecked(True)

    # ------------------------------------------
    # --- TEMPLATE INJECTION & EXPORT ---
    # ------------------------------------------
    def export_pdf(self):
        try:
            temp_file = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
            temp_path = temp_file.name
            temp_file.close() 

            printer = QPrinter(QPrinter.PrinterMode.HighResolution)
            printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
            printer.setOutputFileName(temp_path)
            
            ai_html = self.intro_text_data.replace('\n', '<br>')
            manual_html = self.manual_input.toHtml()
            
            scope_combo = self.combos.get("Scope Of Task Directive")
            scope_text = "<br>".join(scope_combo.checkedItems()) if scope_combo and scope_combo.checkedItems() else "____"
            
            sth_html = "<table border='1' cellspacing='0' cellpadding='5' width='100%'><tr><th>Sl No.</th><th>Organisation</th><th>Role</th><th>Activities</th></tr>"
            if hasattr(self, 'stakeholders_data') and self.stakeholders_data:
                for i, sh in enumerate(self.stakeholders_data):
                    sth_html += f"<tr><td>{i+1}</td><td>{sh.get('org', '')}</td><td>{sh.get('role', '')}</td><td>{sh.get('activities', '')}</td></tr>"
            else:
                sth_html += "<tr><td>&nbsp;</td><td></td><td></td><td></td></tr>"
            sth_html += "</table>"
            
            ct_html = "<table border='1' cellspacing='0' cellpadding='5' width='100%'><tr><th>Sl No.</th><th>Certification Activity</th><th>Certification Work Centre</th><th>Responsible Head</th></tr>"
            if hasattr(self, 'cert_task_data') and self.cert_task_data:
                for i, ct in enumerate(self.cert_task_data):
                    ct_html += f"<tr><td>{i+1}</td><td>{ct.get('activity', '')}</td><td>{ct.get('centre', '')}</td><td>{ct.get('head', '')}</td></tr>"
            else:
                ct_html += "<tr><td>&nbsp;</td><td></td><td></td><td></td></tr>"
            ct_html += "</table>"
            
            final_html = f"<h1>Document Analysis Report</h1><hr><h2>1. Introduction Summary</h2><p>{ai_html}</p><br><h2>2. Manual Notes & Flowcharts</h2>{manual_html}<br><h2>Scope of Directive</h2><p>{scope_text}</p><br><h2>Stakeholders</h2>{sth_html}<br><h2>Certification Task Allocation</h2>{ct_html}"
            
            doc = QTextDocument()
            doc.setHtml(final_html)
            doc.print(printer)
            
            self.preview_dialog = PDFViewerDialog(temp_path, self)
            self.preview_dialog.exec()
            
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Could not generate PDF: {str(e)}")

    def export_docx(self):
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Word", "Task_Directive_Final.docx", "Word (*.docx)")
        if not save_path: return
            
        try:
            def add_page_border(section):
                sectPr = section._sectPr
                pgBorders = OxmlElement('w:pgBorders')
                pgBorders.set(qn('w:offsetFrom'), 'page')
                for side in ['top', 'left', 'bottom', 'right']:
                    border = OxmlElement(f'w:{side}')
                    border.set(qn('w:val'), 'single'); border.set(qn('w:sz'), '12'); border.set(qn('w:space'), '24'); border.set(qn('w:color'), '000000')
                    pgBorders.append(border)
                sectPr.append(pgBorders)

            def remove_page_border(section):
                sectPr = section._sectPr
                for el in sectPr.findall(qn('w:pgBorders')): sectPr.remove(el)

            def add_heading(doc, text):
                p = doc.add_paragraph()
                run = p.add_run(text)
                run.bold = True; run.font.size = Pt(12)

            doc = Document()
            style = doc.styles['Normal']
            style.font.name = 'Calibri'; style.font.size = Pt(11)

            section1 = doc.sections[0]
            section1.top_margin = Inches(1.5); section1.bottom_margin = Inches(1.5)
            section1.left_margin = Inches(1.5); section1.right_margin = Inches(1.5)
            add_page_border(section1)

            doc.add_paragraph("File No. ________________________")
            p = doc.add_paragraph()
            run = p.add_run("\nTASK DIRECTIVE\n"); run.bold = True; run.font.size = Pt(20)
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p = doc.add_paragraph("________ /2026"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER

            table = doc.add_table(rows=1, cols=2)
            table.cell(0, 0).text = "Issue No."; table.cell(0, 1).text = "Date of Issue:"

            p_name = getattr(self, 'extracted_project_name', '').strip()
            if not p_name or p_name.lower() == "not found":
                p_name = "__________________________________________"
            doc.add_paragraph(f"\nPROJECT NAME: {p_name}")

            p = doc.add_paragraph("\n\n[ LOGO HERE ]"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("ABC Organization").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("Address of organization").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("Organization details").alignment = WD_ALIGN_PARAGRAPH.CENTER

            section2 = doc.add_section(WD_SECTION.NEW_PAGE)
            section2.header.is_linked_to_previous = False; section2.footer.is_linked_to_previous = False
            section2.top_margin = Inches(0.75); section2.bottom_margin = Inches(0.75)
            section2.left_margin = Inches(0.75); section2.right_margin = Inches(0.75)
            remove_page_border(section2)

            add_heading(doc, "1. Introduction [Automated]")
            intro_text = self.intro_text_data.strip()
            if not intro_text: intro_text = "__________________________________________________"
            doc.add_paragraph(intro_text)

            add_heading(doc, "2. Reference [Default]")
            doc.add_paragraph("__________________________________________________")
            add_heading(doc, "3. Basis Of Task Directive [Default]")
            doc.add_paragraph("__________________________________________________")
            
            scope_combo = self.combos.get("Scope Of Task Directive")
            scope_text = "\n".join(scope_combo.checkedItems()) if scope_combo and scope_combo.checkedItems() else "____"
            add_heading(doc, "4. Scope Of Task Directive [Drop down menu]")
            doc.add_paragraph(f"To assign the certification respectively:\n{scope_text}\n\n1) ____\n2) ____\n3) ____")

            add_heading(doc, f"5. Stakeholders [Automated + Manual Notes]")
            doc.add_paragraph("The following are the major stakeholders and manual notes extracted:")
            doc.add_paragraph(self.manual_input.toPlainText()) 

            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(["Sl No.", "Organisation", "Role", "Activities"]): 
                hdr_cells[i].text = h

            if hasattr(self, 'stakeholders_data') and self.stakeholders_data:
                for i, sh in enumerate(self.stakeholders_data):
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(i + 1)
                    row_cells[1].text = sh.get('org', '')
                    row_cells[2].text = sh.get('role', '')
                    row_cells[3].text = sh.get('activities', '')
            else:
                for _ in range(3): table.add_row() 

            cert_combo = self.combos.get("Certification Work Breakdown")
            cert_text = "\n".join(cert_combo.checkedItems()) if cert_combo and cert_combo.checkedItems() else "______________________________________________"
            add_heading(doc, "6. Certification Work Breakdown [Drop down menu]")
            doc.add_paragraph(cert_text)
            
            task_combo = self.combos.get("Task Allocation")
            task_text = "\n".join(task_combo.checkedItems()) if task_combo and task_combo.checkedItems() else "______________________________________________"
            add_heading(doc, "7. Task Allocation [Default]")
            doc.add_paragraph(task_text)
            add_heading(doc, "7.1 Coordinating Directorate [Default]")
            doc.add_paragraph("______________________________________________")
            add_heading(doc, "7.2 Single Point of Contact (SPoC) [Default]")
            doc.add_paragraph("______________________________________________")

            add_heading(doc, "7.3 Certification Task Allocation")
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(["Sl No.", "Certification Activity", "Certification Work Centre", "Responsible Head"]): 
                hdr_cells[i].text = h

            if hasattr(self, 'cert_task_data') and self.cert_task_data:
                for i, ct in enumerate(self.cert_task_data):
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(i + 1)
                    row_cells[1].text = ct.get('activity', '')
                    row_cells[2].text = ct.get('centre', '')
                    row_cells[3].text = ct.get('head', '')
            else:
                for _ in range(3): table.add_row()

            add_heading(doc, "7.4 Issue of Clearance [Default]")
            doc.add_paragraph("a. ___\nb. ___\nc. ___\nd. ___\ne. ___\nf. ___\ng. ___")

            add_heading(doc, "8. SCRB And TARB [Default]")
            doc.add_paragraph("______________________________________________")
            
            comm_combo = self.combos.get("Communication Type")
            comm_text = "\n".join(comm_combo.checkedItems()) if comm_combo and comm_combo.checkedItems() else "______________________________________________"
            add_heading(doc, "9. Communication [Default]")
            doc.add_paragraph(comm_text)
            add_heading(doc, "10. Certification Progress Review [Default]")
            doc.add_paragraph("______________________________________________\n\n(__________)")

            add_heading(doc, "11. Distribution List")
            doc.add_paragraph("11.1 External Organization\n1. ___\n2. ___\n3. ___\n11.2 Internal Distribution")

            doc.add_page_break()
            add_heading(doc, "Annexure-1")
            doc.add_paragraph("Work Assignment List of LRUs [Automate + Manual]")
            table = doc.add_table(rows=4, cols=6); table.style = 'Table Grid'
            for i, h in enumerate(["Sl No", "ABC", "Abc2", "Abc3", "Abc4", "Abc5"]): table.cell(0, i).text = h

            doc.add_page_break()
            add_heading(doc, "Integration Checks and Clearance [Manual]")
            table = doc.add_table(rows=3, cols=5); table.style = 'Table Grid'
            for i, h in enumerate(["Sl no.", "Abc1", "Abc2", "Abc3", "Abc4"]): table.cell(0, i).text = h

            doc.add_page_break()
            add_heading(doc, "Annexure-2")
            doc.add_paragraph("Contact details of dealing officers and RDs [Manual]")
            table = doc.add_table(rows=3, cols=5); table.style = 'Table Grid'
            for i, h in enumerate(["Sl no.", "Abc1", "Abc2", "Abc3", "Abc4"]): table.cell(0, i).text = h

            doc.add_page_break()
            add_heading(doc, "Annexure-3")
            doc.add_paragraph("Product Break Down Structure [Automate]")

            doc.save(save_path)
            QMessageBox.information(self, "Success", "Task Directive Template created successfully!")

        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to build template: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OfflineApp()
    window.show()
    sys.exit(app.exec())
