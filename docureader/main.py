import sys
import os
import ollama
import fitz  # PyMuPDF
import pdfplumber
import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QPushButton, QTextEdit, QFileDialog,
    QLabel, QSplitter, QMessageBox, QScrollArea, QProgressBar,
    QCheckBox, QGroupBox, QDialog, QComboBox, QFrame
)
from PyQt6.QtGui import QPixmap, QImage
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from fpdf import FPDF
from docx import Document
import subprocess
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from template_gen_2 import generate_template


# --- WORKER THREAD FOR AI ---
class AIWorker(QThread):
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, text):
        super().__init__()
        self.text = text
        

    def run(self):
        try:
            prompt = (
                f"Analyze this document. Provide a summary of key objectives, "
                f"technical findings, and conclusions:\n\n{self.text[:8000]}"
            )
            response = ollama.generate(model='llama3', prompt=prompt)
            self.finished.emit(response['response'])
        except Exception as e:
            self.error.emit(str(e))

# --- MAIN APPLICATION ---
class OfflineApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Offline AI Document Assistant - Pro Edition")
        self.resize(1400, 900)
        
        self.file_path = ""
        self.doc = None
        self.current_pdf_text = ""
        self.current_ai_summary = "No summary generated yet."
        self.zoom_factor = 1.2  # default zoom

        self.apply_stylesheet()
        self.init_ui()

    def apply_stylesheet(self):
        self.setStyleSheet("""
            * {
                font-family: 'Segoe UI', sans-serif;
            }
            QMainWindow, #mainWidget {
                background: #f5f7fb;
            }
            #leftContainer {
                background: #ffffff;
                border-right: 1px solid #e0e0e0;
            }
            #rightContainer {
                background: #f9fafc;
            }
            QFrame[class="Card"] {
                background: white;
                border-radius: 12px;
                border: 1px solid #e0e0e0;
            }
            QPushButton {
                background: #4f46e5;
                color: white;
                border: none;
                padding: 12px;
                border-radius: 8px;
                font-weight: 500;
                font-size: 14px;
            }
            QPushButton:hover {
                background: #4338ca;
            }
            QGroupBox {
                background: transparent;
                border: none;
            }
            QComboBox {
                padding: 8px;
                border-radius: 6px;
                border: 1px solid #ccc;
                background: white;
            }
            QSplitter::handle {
                background-color: #cbd5e1;
                border-radius: 2px;
                margin: 2px;
            }
        """)
    
    def render_pdf(self):
        if not self.doc:
            return

        self.clear_preview()

        for i, page in enumerate(self.doc):
            if i < 15:
                pix = page.get_pixmap(matrix=fitz.Matrix(self.zoom_factor, self.zoom_factor))
                img = QImage(
                    pix.samples, pix.width, pix.height,
                    pix.stride, QImage.Format.Format_RGB888
                ).copy()

                lbl = QLabel()
                lbl.setPixmap(QPixmap.fromImage(img))
                self.page_layout.addWidget(lbl)

    def zoom_in(self):
        self.zoom_factor += 0.2
        self.render_pdf()

    def zoom_out(self):
        if self.zoom_factor > 0.4:
            self.zoom_factor -= 0.2
            self.render_pdf()

    def init_ui(self):
        main_widget = QWidget()
        main_widget.setObjectName("mainWidget")
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)
        self.splitter.setHandleWidth(6)

        # ==========================================
        # --- LEFT SIDE: Document Viewer ---
        # ==========================================
        left_container = QWidget()
        left_container.setObjectName("leftContainer")
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(20, 20, 20, 20)
        left_layout.setSpacing(15)

        # Card for left side
        left_card = QFrame()
        left_card.setProperty("class", "Card")
        left_card_layout = QVBoxLayout(left_card)
        left_card_layout.setContentsMargins(15, 15, 15, 15)
        left_card_layout.setSpacing(10)

        self.upload_btn = QPushButton("Upload a File")
        self.upload_btn.setStyleSheet("height: 40px; font-weight: bold; background-color: #4f46e5; border-radius: 8px;")
        self.upload_btn.clicked.connect(self.upload_file)
        left_card_layout.addWidget(self.upload_btn)

        zoom_layout = QHBoxLayout()
        self.zoom_in_btn = QPushButton("➕")
        self.zoom_out_btn = QPushButton("➖")
        self.zoom_in_btn.setFixedSize(40, 40)
        self.zoom_out_btn.setFixedSize(40, 40)
        self.zoom_in_btn.setStyleSheet("font-size: 16px; padding: 0;")
        self.zoom_out_btn.setStyleSheet("font-size: 16px; padding: 0;")
        zoom_layout.addWidget(self.zoom_in_btn)
        zoom_layout.addWidget(self.zoom_out_btn)
        zoom_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)
        left_card_layout.addLayout(zoom_layout)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("QScrollArea { border: none; background: #f1f3f8; border-radius: 10px; }")

        self.page_container = QWidget()
        self.page_container.setStyleSheet("background: #f1f3f8;")
        self.page_layout = QVBoxLayout(self.page_container)
        self.page_layout.setContentsMargins(10, 10, 10, 10)
        self.scroll_area.setWidget(self.page_container)

        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        self.text_preview.hide()
        self.text_preview.setStyleSheet("border: none; background: #f1f3f8; border-radius: 10px; padding: 10px; color: #333;")

        left_card_layout.addWidget(self.scroll_area)
        left_card_layout.addWidget(self.text_preview)

        left_layout.addWidget(left_card)
        self.splitter.addWidget(left_container)

        # ==========================================
        # --- RIGHT SIDE: Tools & AI ---
        # ==========================================
        right_container = QWidget()
        right_container.setObjectName("rightContainer")
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(20, 20, 20, 20)
        right_layout.setSpacing(20)

        # Card 1: Document Status
        status_card = QFrame()
        status_card.setProperty("class", "Card")
        status_layout = QVBoxLayout(status_card)
        
        # Document Status title and AI Icon 
        top_bar_layout = QHBoxLayout()
        title_lbl = QLabel("<h3>Document Status</h3>")
        title_lbl.setStyleSheet("margin: 0; color: #333; font-size: 18px;")
        top_bar_layout.addWidget(title_lbl)
        top_bar_layout.addStretch()

        self.ai_icon_btn = QPushButton("AI")
        self.ai_icon_btn.setFixedSize(50, 50)
        self.ai_icon_btn.setToolTip("Click to view AI Summary")
        self.ai_icon_btn.setStyleSheet("""
            QPushButton {
                font-size: 24px; 
                border-radius: 25px; 
                background-color: #f8f9fa;
                border: 2px solid #ced4da;
                color: black;
                padding: 0;
            }
            QPushButton:hover { background-color: #e2e6ea; }
        """)
        self.ai_icon_btn.clicked.connect(self.show_ai_summary_popup)
        top_bar_layout.addWidget(self.ai_icon_btn)
        status_layout.addLayout(top_bar_layout)

        chk_layout = QHBoxLayout()
        self.check_upload = QCheckBox("File Uploaded")
        self.check_tables = QCheckBox("Tables Scanned")
        self.check_ai = QCheckBox("AI Analysis Done")
        for cb in [self.check_upload, self.check_tables, self.check_ai]:
            cb.setEnabled(False)
            cb.setStyleSheet("color: #3730a3; font-weight: bold;")
            chk_layout.addWidget(cb)
        status_layout.addLayout(chk_layout)
        right_layout.addWidget(status_card)

        # Card 2: Actions
        actions_card = QFrame()
        actions_card.setProperty("class", "Card")
        actions_layout = QVBoxLayout(actions_card)
        actions_title = QLabel("<h3>Actions</h3>")
        actions_title.setStyleSheet("margin: 0; color: #333; font-size: 18px;")
        actions_layout.addWidget(actions_title)

        self.status_label = QLabel("Ready")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)
        self.progress_bar.hide()
        actions_layout.addWidget(self.status_label)
        actions_layout.addWidget(self.progress_bar)

        self.scan_tables_btn = QPushButton("Scan for Tables")
        self.scan_tables_btn.setFixedHeight(40)
        self.scan_tables_btn.clicked.connect(self.scan_tables)
        actions_layout.addWidget(self.scan_tables_btn)

        self.generate_btn = QPushButton("Generate Document")
        self.generate_btn.clicked.connect(self.generate_doc)
        actions_layout.addWidget(self.generate_btn)

        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(10)
        self.save_docx_btn = QPushButton("Save as DOCX")
        self.download_pdf_btn = QPushButton("Download PDF")
        
        self.save_docx_btn.clicked.connect(self.save_docx)
        self.download_pdf_btn.clicked.connect(self.download_pdf)
        btn_layout.addWidget(self.save_docx_btn)
        btn_layout.addWidget(self.download_pdf_btn)
        actions_layout.addLayout(btn_layout)
        right_layout.addWidget(actions_card)

        # Card 3: Manual Entry
        manual_card = QFrame()
        manual_card.setProperty("class", "Card")
        manual_layout = QVBoxLayout(manual_card)
        manual_title = QLabel("<h3>Manual Data (Template Form)</h3>")
        manual_title.setStyleSheet("margin: 0; color: #333; font-size: 18px;")
        manual_layout.addWidget(manual_title)

        self.manual_form = QWidget()
        form_layout = QVBoxLayout(self.manual_form)
        form_layout.setContentsMargins(0, 0, 0, 0)
        form_layout.setSpacing(10)

        dropdowns = [
            ("Scope Of Task Directive", ["Option 1", "Option 2", "Option 3"]),
            ("Stakeholders", ["Internal", "External", "Both"]),
            ("Certification Work Breakdown", ["Type A", "Type B", "Type C"]),
            ("Task Allocation", ["Auto", "Manual", "Hybrid"]),
            ("Communication Type", ["Email", "Meeting", "Report"])
        ]
        
        self.combos = []
        for label, items in dropdowns:
            row = QHBoxLayout()
            lbl = QLabel(label)
            lbl.setFixedWidth(200)
            lbl.setStyleSheet("color: #333;")
            combo = QComboBox()
            combo.addItems(items)
            combo.setStyleSheet("color: #333;")
            row.addWidget(lbl)
            row.addWidget(combo)
            form_layout.addLayout(row)
            self.combos.append(combo)
            
        self.scope_dropdown, self.stakeholder_dropdown, self.cert_dropdown, self.task_dropdown, self.comm_dropdown = self.combos

        manual_layout.addWidget(self.manual_form)
        right_layout.addWidget(manual_card)

        # Card 4: Extracted Tables
        tables_card = QFrame()
        tables_card.setProperty("class", "Card")
        tables_layout = QVBoxLayout(tables_card)
        tables_title = QLabel("<h3>Extracted Tables</h3>")
        tables_title.setStyleSheet("margin: 0; color: black; font-size: 18px;")
        tables_layout.addWidget(tables_title)
        
        self.table_output = QTextEdit()
        self.table_output.setReadOnly(True)
        self.table_output.setStyleSheet("border: 1px solid black; border-radius: 6px; background: white; color: black;")
        tables_layout.addWidget(self.table_output)
        right_layout.addWidget(tables_card)

        # Add everything to splitter
        self.splitter.addWidget(left_container)
        self.splitter.addWidget(right_container)
        self.splitter.setStretchFactor(0, 5)
        self.splitter.setStretchFactor(1, 5)
        layout.addWidget(self.splitter)

    # ==========================================
    # --- LOGIC & FUNCTIONS ---
    # ==========================================
    def reset_status(self):
        self.check_upload.setChecked(False)
        self.check_tables.setChecked(False)
        self.check_ai.setChecked(False)
        self.current_ai_summary = "No summary generated yet."
        self.ai_icon_btn.setStyleSheet("""
            QPushButton { font-size: 24px; border-radius: 25px; background-color: #f8f9fa; border: 2px solid #ced4da; color: black; padding: 0; }
            QPushButton:hover { background-color: #e2e6ea; }
        """)

    def clear_preview(self):
        while self.page_layout.count():
            child = self.page_layout.takeAt(0)
            if child.widget(): child.widget().deleteLater()

    def upload_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Document", "", "Documents (*.pdf *.txt)")
        if path:
            self.file_path = path
            self.reset_status()
            self.clear_preview()
            
            try:
                if path.lower().endswith('.pdf'):
                    self.text_preview.hide()
                    self.scroll_area.show()
                    self.doc = fitz.open(path)
                    text = ""
                    self.current_pdf_text = text
                    self.render_pdf()
                else:
                    self.scroll_area.hide()
                    self.text_preview.show()
                    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                        self.current_pdf_text = f.read()
                    self.text_preview.setText(self.current_pdf_text)

                self.check_upload.setChecked(True)
                # Auto-analysis disabled per user request
                # self.start_ai_thread(self.current_pdf_text) 
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not load file: {str(e)}")

    def start_ai_thread(self, text):
        self.status_label.setText("AI is analyzing document...")
        self.progress_bar.show()
        self.worker = AIWorker(text)
        self.worker.finished.connect(self.handle_ai_done)
        self.worker.error.connect(lambda e: self.status_label.setText(f"AI Error: {e}"))
        self.worker.start()

    def handle_ai_done(self, summary):
        self.current_ai_summary = summary
        self.progress_bar.hide()
        self.status_label.setText("Analysis Complete")
        self.check_ai.setChecked(True)
        self.ai_icon_bgtn.setStyleSheet("""
            QPushButton { font-size: 24px; border-radius: 25px; background-color: #d4edda; border: 2px solid #28a745; color: black; padding: 0; }
            QPushButton:hover { background-color: #c3e6cb; }
        """)

    def show_ai_summary_popup(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("AI Automated Summary")
        dialog.resize(700, 500)
        layout = QVBoxLayout(dialog)
        
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setText(self.current_ai_summary)
        text_edit.setStyleSheet("font-size: 14px; line-height: 1.6; padding: 10px;")
        
        close_btn = QPushButton("Close")
        close_btn.setFixedHeight(40)
        close_btn.clicked.connect(dialog.accept)
        
        layout.addWidget(text_edit)
        layout.addWidget(close_btn)
        dialog.exec()

    def scan_tables(self):
        if not self.file_path or not self.file_path.lower().endswith('.pdf'):
            QMessageBox.warning(self, "Notice", "Table scanning is only available for PDF files.")
            return
            
        self.scan_tables_btn.setText("Scanning... Please wait.")
        QApplication.processEvents()
        
        try:
            table_content = "\n--- EXTRACTED TABLES ---\n"
            found_tables = False
            
            with pdfplumber.open(self.file_path) as pdf:
                for page in pdf.pages:
                    tables = page.extract_tables()
                    for table in tables:
                        if table and len(table) > 1:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            table_content += df.to_string(index=False) + "\n\n"
                            found_tables = True
            
            if found_tables:
                self.table_output.setText(table_content)
                self.check_tables.setChecked(True)
            else:
                QMessageBox.information(self, "No Tables", "Could not find any clear data tables in this document.")
                
        except Exception as e:
            QMessageBox.critical(self, "Table Scan Error", str(e))
            
        self.scan_tables_btn.setText("Scan for Tables")
 
    def generate_doc(self):
        try:
            path = generate_template(
                self.scope_dropdown.currentText(),
                self.stakeholder_dropdown.currentText(),
                self.cert_dropdown.currentText(),
                self.task_dropdown.currentText(),
                self.comm_dropdown.currentText()
            )
            self.generated_docx_path = path
            QMessageBox.information(self, "Success", "Document successfully generated!")
            self.status_label.setText("Template Generated + Filled")
        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    def save_docx(self):
        if not hasattr(self, "generated_docx_path"):
            QMessageBox.warning(self, "Notice", "Please generate the document first.")
            return
        
        save_path, _ = QFileDialog.getSaveFileName(self, "Save DOCX", "Document.docx", "Word Document (*.docx)")
        if save_path:
            import shutil
            shutil.copy(self.generated_docx_path, save_path)
            QMessageBox.information(self, "Success", f"Document saved as:\n{save_path}")

    def download_pdf(self):
        if not hasattr(self, "generated_docx_path"):
            QMessageBox.warning(self, "Notice", "Please generate the document first.")
            return

        save_path, _ = QFileDialog.getSaveFileName(self, "Download PDF", "Document.pdf", "PDF file (*.pdf)")
        if not save_path:
            return

        try:
            ps_script = f'''
$word = New-Object -ComObject Word.Application
$doc = $word.Documents.Open('{os.path.abspath(self.generated_docx_path)}')
$doc.SaveAs([ref]'{os.path.abspath(save_path)}', [ref]17)
$doc.Close()
$word.Quit()
'''
            flags = 0x08000000 if sys.platform == "win32" else 0
            subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], check=True, creationflags=flags)
            QMessageBox.information(self, "Success", f"PDF downloaded as:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to convert PDF via Word:\n{e}\nEnsure Microsoft Word is installed.")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OfflineApp()
    window.show()
    sys.exit(app.exec())
