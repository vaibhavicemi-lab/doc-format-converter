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
    QCheckBox, QGroupBox, QDialog, QComboBox
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

        self.init_ui()
    
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
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)

        # ==========================================
        # --- LEFT SIDE: Document Viewer ---
        # ==========================================
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)

        # Upload button
        self.upload_btn = QPushButton(" 📂 Upload a File")
        self.upload_btn.setStyleSheet("height: 50px; font-weight: bold; color: black; background-color: #f8f9fa;")
        self.upload_btn.clicked.connect(self.upload_file)

        left_layout.addWidget(self.upload_btn)

        # ✅ Zoom buttons (FIXED POSITION)
        zoom_layout = QHBoxLayout()

        self.zoom_in_btn = QPushButton("➕")
        self.zoom_out_btn = QPushButton("➖")

        self.zoom_in_btn.clicked.connect(self.zoom_in)
        self.zoom_out_btn.clicked.connect(self.zoom_out)

        zoom_layout.addWidget(self.zoom_in_btn)
        zoom_layout.addWidget(self.zoom_out_btn)

        left_layout.addLayout(zoom_layout)

        # Scroll area
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)

        self.page_container = QWidget()
        self.page_layout = QVBoxLayout(self.page_container)
        self.scroll_area.setWidget(self.page_container)

        # Text preview
        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        self.text_preview.hide()

        left_layout.addWidget(self.scroll_area)
        left_layout.addWidget(self.text_preview)

        # Add to splitter
        self.splitter.addWidget(left_container)
        layout.addWidget(self.splitter)


        # ==========================================
        # --- RIGHT SIDE: Tools & AI ---
        # ==========================================
        right_container = QWidget()
        right_layout = QVBoxLayout(right_container)
        right_layout.setSpacing(12) # Adds a nice gap between each row

        # 1. TOP BAR: Just the AI Icon on the far right
        top_bar_layout = QHBoxLayout()
        top_bar_layout.addStretch() # Pushes the icon to the right
        
        self.ai_icon_btn = QPushButton("🤖")
        self.ai_icon_btn.setFixedSize(65, 65)
        self.ai_icon_btn.setToolTip("Click to view AI Summary")
        self.ai_icon_btn.setStyleSheet("""
            QPushButton {
                font-size: 30px; 
                border-radius: 32px; 
                background-color: #f8f9fa;
                border: 2px solid #ced4da;
            }
            QPushButton:hover { background-color: #e2e6ea; }
        """)
        self.ai_icon_btn.clicked.connect(self.show_ai_summary_popup)
        top_bar_layout.addWidget(self.ai_icon_btn)
        
        right_layout.addLayout(top_bar_layout)

        # 2. STATUS BAR (Full width)
        self.status_group = QGroupBox("Status Bar")
        status_layout = QHBoxLayout() # Keeps checkboxes side-by-side
        self.check_upload = QCheckBox("File Uploaded")
        self.check_tables = QCheckBox("Tables Scanned")
        self.check_ai = QCheckBox("AI Analysis Done")
        
        for cb in [self.check_upload, self.check_tables, self.check_ai]:
            cb.setEnabled(False) 
            status_layout.addWidget(cb)
            
        self.status_group.setLayout(status_layout)
        right_layout.addWidget(self.status_group)
        
        # 3. PROGRESS & TOOLS
        self.status_label = QLabel("Ready")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0) 
        self.progress_bar.hide()
        
        self.scan_tables_btn = QPushButton(" 📊 Scan for Tables")
        self.scan_tables_btn.setFixedHeight(40)
        self.scan_tables_btn.clicked.connect(self.scan_tables)

        right_layout.addWidget(self.status_label)
        right_layout.addWidget(self.progress_bar)
        right_layout.addWidget(self.scan_tables_btn)

        # 4. MANUAL ENTRY (Directly underneath the tools)
        right_layout.addWidget(QLabel(" ✍️ Manual Data (Template Form):"))

        self.manual_form = QWidget()
        form_layout = QVBoxLayout(self.manual_form)

        # Scope
        form_layout.addWidget(QLabel("Scope Of Task Directive"))
        self.scope_dropdown = QComboBox()
        self.scope_dropdown.addItems(["Option 1", "Option 2", "Option 3"])
        form_layout.addWidget(self.scope_dropdown)

        # Stakeholders
        form_layout.addWidget(QLabel("Stakeholders"))
        self.stakeholder_dropdown = QComboBox()
        self.stakeholder_dropdown.addItems(["Internal", "External", "Both"])
        form_layout.addWidget(self.stakeholder_dropdown)

        # Certification
        form_layout.addWidget(QLabel("Certification Work Breakdown"))
        self.cert_dropdown = QComboBox()
        self.cert_dropdown.addItems(["Type A", "Type B", "Type C"])
        form_layout.addWidget(self.cert_dropdown)

        # Task Allocation
        form_layout.addWidget(QLabel("Task Allocation"))
        self.task_dropdown = QComboBox()
        self.task_dropdown.addItems(["Auto", "Manual", "Hybrid"])
        form_layout.addWidget(self.task_dropdown)

        # Communication
        form_layout.addWidget(QLabel("Communication Type"))
        self.comm_dropdown = QComboBox()
        self.comm_dropdown.addItems(["Email", "Meeting", "Report"])
        form_layout.addWidget(self.comm_dropdown)

        right_layout.addWidget(self.manual_form)

        # 5. BOTTOM SECTION: Export Buttons
        btn_layout = QHBoxLayout()
        self.save_pdf_btn = QPushButton(" 💾 View PDF")
        self.save_docx_btn = QPushButton(" 📝 Save as Word")
        self.save_pdf_btn.setFixedHeight(45)
        self.save_docx_btn.setFixedHeight(45)
        self.save_pdf_btn.clicked.connect(self.export_pdf)
        self.save_docx_btn.clicked.connect(self.export_docx)
        
        btn_layout.addWidget(self.save_pdf_btn)
        btn_layout.addWidget(self.save_docx_btn)
        
        right_layout.addLayout(btn_layout)

        # Add everything to the main Splitter
        self.splitter.addWidget(left_container)
        self.splitter.addWidget(right_container)
        self.splitter.setStretchFactor(0, 5) # Viewer gets half the screen
        self.splitter.setStretchFactor(1, 5) # Right side gets the other half
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
            QPushButton { font-size: 30px; border-radius: 32px; background-color: #f8f9fa; border: 2px solid #ced4da; }
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
                # ERROR FIX: Uses the thread instead of fetch_ai_summary!
                self.start_ai_thread(self.current_pdf_text) 
                
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
        self.ai_icon_btn.setStyleSheet("""
            QPushButton { font-size: 30px; border-radius: 32px; background-color: #d4edda; border: 2px solid #28a745; }
            QPushButton:hover { background-color: #c3e6cb; }
        """)

    def show_ai_summary_popup(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("🤖 AI Automated Summary")
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
                #self.manual_input.append(table_content)
                print(table_content)
                self.check_tables.setChecked(True)
                QMessageBox.information(self, "Success", "Tables extracted and added to manual notes!")
            else:
                QMessageBox.information(self, "No Tables", "Could not find any clear data tables in this document.")
                
        except Exception as e:
            QMessageBox.critical(self, "Table Scan Error", str(e))
            
        self.scan_tables_btn.setText(" 📊 Scan for Tables")
 
    def export_pdf(self):
        try:
            if not hasattr(self, "generated_docx_path"):
                QMessageBox.warning(self, "No File", "Generate template first!")
                return

            file_path = self.generated_docx_path

            # FORCE OPEN IN MICROSOFT WORD (Windows)
            try:
                subprocess.Popen(['start', 'winword', file_path], shell=True)
            except Exception:
                # fallback if Word not found
                os.startfile(file_path)

            self.status_label.setText("Opened in Word ")

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

    
    def export_docx(self):
        try:
            path = generate_template(
                self.scope_dropdown.currentText(),
                self.stakeholder_dropdown.currentText(),
                self.cert_dropdown.currentText(),
                self.task_dropdown.currentText(),
                self.comm_dropdown.currentText()
            )

            self.generated_docx_path = path

            QMessageBox.information(self, "Success", "Template generated!")
            self.status_label.setText("Template Generated + Filled ✅")

        except Exception as e:
            QMessageBox.critical(self, "Error", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OfflineApp()
    window.show()
    sys.exit(app.exec())