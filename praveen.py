import sys
import ollama
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QPushButton, QTextEdit, QFileDialog,
    QLabel, QSplitter, QMessageBox, QScrollArea, QProgressBar,
    QCheckBox, QGroupBox, QDialog, QRubberBand, QMenu
)
from PyQt6.QtGui import QPixmap, QImage, QFont, QTextDocument
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QEvent, QRect, QPoint
from PyQt6.QtPrintSupport import QPrinter
from docx import Document

# ==========================================
# --- WORKER THREAD FOR AI ---
# ==========================================
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

# ==========================================
# --- INTERACTIVE PDF SNIPPING TOOL ---
# ==========================================
class PdfPageLabel(QLabel):
    """A custom label that allows users to click and drag to copy text/images."""
    def __init__(self, page, parent=None):
        super().__init__(parent)
        self.page = page  # The fitz.Page object used for text extraction
        self.rubber_band = QRubberBand(QRubberBand.Shape.Rectangle, self)
        self.origin = QPoint()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.origin = event.pos()
            self.rubber_band.setGeometry(QRect(self.origin, self.origin))
            self.rubber_band.show()

    def mouseMoveEvent(self, event):
        if not self.origin.isNull():
            self.rubber_band.setGeometry(QRect(self.origin, event.pos()).normalized())

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            rect = self.rubber_band.geometry()
            # Only trigger if the user actually dragged a box (not just a click)
            if rect.width() > 10 and rect.height() > 10:
                self.show_context_menu(event.pos(), rect)
            else:
                self.rubber_band.hide()
            self.origin = QPoint()

    def show_context_menu(self, pos, rect):
        menu = QMenu(self)
        copy_img_action = menu.addAction("🖼️ Copy as Image (Flowchart/Diagram)")
        copy_text_action = menu.addAction("📝 Extract Text")

        action = menu.exec(self.mapToGlobal(pos))
        
        if action == copy_img_action:
            pixmap = self.pixmap()
            if pixmap:
                cropped = pixmap.copy(rect)
                QApplication.clipboard().setPixmap(cropped)
                
        elif action == copy_text_action:
            if self.pixmap() and self.page:
                # Math to convert screen pixels back to PDF coordinates
                scale = self.page.rect.width / self.pixmap().width()
                pdf_rect = fitz.Rect(
                    rect.left() * scale,
                    rect.top() * scale,
                    rect.right() * scale,
                    rect.bottom() * scale
                )
                text = self.page.get_textbox(pdf_rect)
                QApplication.clipboard().setText(text)
                
        self.rubber_band.hide()

# ==========================================
# --- MAIN APPLICATION ---
# ==========================================
class OfflineApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Offline AI Document Assistant - Pro Edition")
        self.resize(1400, 900)
        
        self.file_path = ""
        self.doc = None
        self.current_pdf_text = ""
        self.current_ai_summary = "No summary generated yet."
        
        # UI State Variables
        self.page_data = [] # Stores tuples of (QPixmap, page_index)
        self.current_zoom = 600 
        self.is_maximized = False
        
        self.init_ui()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)

        # ------------------------------------------
        # --- LEFT SIDE: Document Viewer ---
        # ------------------------------------------
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        
        left_controls_layout = QHBoxLayout()
        
        self.upload_btn = QPushButton(" 📂 Upload PDF or TXT")
        self.upload_btn.setStyleSheet("height: 40px; font-weight: bold; background-color: #f8f9fa; padding: 0 15px;")
        self.upload_btn.clicked.connect(self.upload_file)
        
        self.zoom_out_btn = QPushButton("➖")
        self.zoom_in_btn = QPushButton("➕")
        self.maximize_btn = QPushButton("🗖") 
        
        for btn in [self.zoom_out_btn, self.zoom_in_btn, self.maximize_btn]:
            btn.setFixedSize(40, 40)
            btn.setStyleSheet("font-size: 16px; font-weight: bold; background-color: #ffffff; border: 1px solid #ced4da; border-radius: 4px;")
            
        self.zoom_out_btn.clicked.connect(lambda: self.zoom_out(150))
        self.zoom_in_btn.clicked.connect(lambda: self.zoom_in(150))
        self.maximize_btn.clicked.connect(self.toggle_maximize)

        left_controls_layout.addWidget(self.upload_btn)
        left_controls_layout.addStretch()
        left_controls_layout.addWidget(QLabel("Zoom:"))
        left_controls_layout.addWidget(self.zoom_out_btn)
        left_controls_layout.addWidget(self.zoom_in_btn)
        left_controls_layout.addWidget(self.maximize_btn)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background-color: #525659;") 
        self.page_container = QWidget()
        self.page_layout = QVBoxLayout(self.page_container)
        self.page_layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self.scroll_area.setWidget(self.page_container)

        # Enable CTRL+Scroll Zooming
        self.scroll_area.viewport().installEventFilter(self)

        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        self.text_preview.hide()

        left_layout.addLayout(left_controls_layout)
        # Added a helpful tip for the user
        tip_label = QLabel("💡 Tip: Click and drag on the PDF to copy text, flowcharts, or diagrams!")
        tip_label.setStyleSheet("color: #0056b3; font-style: italic;")
        left_layout.addWidget(tip_label)
        left_layout.addWidget(self.scroll_area)
        left_layout.addWidget(self.text_preview)

        # ------------------------------------------
        # --- RIGHT SIDE: Tools & AI ---
        # ------------------------------------------
        self.right_container = QWidget() 
        right_layout = QVBoxLayout(self.right_container)

        top_bar_layout = QHBoxLayout()
        
        self.status_group = QGroupBox("Status Bar")
        status_layout = QHBoxLayout() 
        self.check_upload = QCheckBox("File Uploaded")
        self.check_ai = QCheckBox("AI Analysis Done")
        
        for cb in [self.check_upload, self.check_ai]:
            cb.setEnabled(False) 
            status_layout.addWidget(cb)
            
        self.status_group.setLayout(status_layout)
        top_bar_layout.addWidget(self.status_group)
        
        self.ai_icon_btn = QPushButton("🤖")
        self.ai_icon_btn.setFixedSize(65, 65)
        self.ai_icon_btn.setToolTip("Click to view AI Summary")
        self.ai_icon_btn.setStyleSheet("""
            QPushButton { font-size: 30px; border-radius: 32px; background-color: #f8f9fa; border: 2px solid #ced4da; }
            QPushButton:hover { background-color: #e2e6ea; }
        """)
        self.ai_icon_btn.clicked.connect(self.show_ai_summary_popup)
        top_bar_layout.addWidget(self.ai_icon_btn, alignment=Qt.AlignmentFlag.AlignTop)

        self.status_label = QLabel("Ready")
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0) 
        self.progress_bar.hide()

        self.manual_input = QTextEdit()
        self.manual_input.setPlaceholderText("Paste extracted text and flowcharts here! (Ctrl+V)")
        self.manual_input.setStyleSheet("background-color: white; font-size: 13px;")

        btn_layout = QHBoxLayout()
        self.save_pdf_btn = QPushButton(" 💾 View PDF")
        self.save_docx_btn = QPushButton(" 📝 Save as Word")
        self.save_pdf_btn.setFixedHeight(45)
        self.save_docx_btn.setFixedHeight(45)
        self.save_pdf_btn.clicked.connect(self.export_pdf)
        self.save_docx_btn.clicked.connect(self.export_docx)
        btn_layout.addWidget(self.save_pdf_btn)
        btn_layout.addWidget(self.save_docx_btn)

        right_layout.addLayout(top_bar_layout)
        right_layout.addWidget(self.status_label)
        right_layout.addWidget(self.progress_bar)
        right_layout.addWidget(QLabel(" ✍️ Manual Notes & Flowcharts:"))
        right_layout.addWidget(self.manual_input)
        right_layout.addLayout(btn_layout)

        self.splitter.addWidget(left_container)
        self.splitter.addWidget(self.right_container)
        self.splitter.setStretchFactor(0, 5) 
        self.splitter.setStretchFactor(1, 5) 
        layout.addWidget(self.splitter)

    # ------------------------------------------
    # --- MOUSE EVENT FILTER FOR CTRL+SCROLL ---
    # ------------------------------------------
    def eventFilter(self, source, event):
        if source == self.scroll_area.viewport() and event.type() == QEvent.Type.Wheel:
            if QApplication.keyboardModifiers() == Qt.KeyboardModifier.ControlModifier:
                delta = event.angleDelta().y()
                if delta > 0:
                    self.zoom_in(step=50)
                elif delta < 0:
                    self.zoom_out(step=50)
                return True 
        return super().eventFilter(source, event)

    # ------------------------------------------
    # --- UI INTERACTION FUNCTIONS ---
    # ------------------------------------------
    def toggle_maximize(self):
        if not self.is_maximized:
            self.right_container.hide() 
            self.maximize_btn.setText("🗗") 
            self.is_maximized = True
        else:
            self.right_container.show() 
            self.maximize_btn.setText("🗖") 
            self.is_maximized = False

    def zoom_in(self, step=150):
        self.current_zoom += step
        self.refresh_pdf_view()

    def zoom_out(self, step=150):
        if self.current_zoom - step > 200: 
            self.current_zoom -= step
            self.refresh_pdf_view()

    def refresh_pdf_view(self):
        if not self.page_data: 
            font = self.text_preview.font()
            font.setPointSize(max(8, int(self.current_zoom / 50)))
            self.text_preview.setFont(font)
            return

        if self.page_layout.count() != len(self.page_data):
            self.clear_preview()
            for pixmap, page_idx in self.page_data:
                page = self.doc.load_page(page_idx)
                # Use our new Interactive snipping tool label!
                lbl = PdfPageLabel(page)
                lbl.setStyleSheet("background-color: white; border: 1px solid black; margin-bottom: 10px;")
                self.page_layout.addWidget(lbl)

        for i in range(self.page_layout.count()):
            lbl = self.page_layout.itemAt(i).widget()
            if isinstance(lbl, PdfPageLabel) and i < len(self.page_data):
                original_pixmap = self.page_data[i][0]
                scaled_pixmap = original_pixmap.scaledToWidth(self.current_zoom, Qt.TransformationMode.SmoothTransformation)
                lbl.setPixmap(scaled_pixmap)

    def clear_preview(self):
        while self.page_layout.count():
            child = self.page_layout.takeAt(0)
            if child.widget(): child.widget().deleteLater()

    # ------------------------------------------
    # --- CORE LOGIC & FUNCTIONS ---
    # ------------------------------------------
    def reset_status(self):
        self.check_upload.setChecked(False)
        self.check_ai.setChecked(False)
        self.current_ai_summary = "No summary generated yet."
        self.ai_icon_btn.setStyleSheet("""
            QPushButton { font-size: 30px; border-radius: 32px; background-color: #f8f9fa; border: 2px solid #ced4da; }
            QPushButton:hover { background-color: #e2e6ea; }
        """)

    def upload_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Document", "", "Documents (*.pdf *.txt)")
        if path:
            self.file_path = path
            self.reset_status()
            self.page_data.clear() 
            
            try:
                if path.lower().endswith('.pdf'):
                    self.text_preview.hide()
                    self.scroll_area.show()
                    self.doc = fitz.open(path)
                    text = ""
                    
                    self.status_label.setText("Loading PDF Pages...")
                    QApplication.processEvents()

                    for i, page in enumerate(self.doc):
                        text += page.get_text()
                        if i < 15: 
                            pix = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0))
                            img = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()
                            # We now store the image AND the page index for the snipping tool
                            self.page_data.append((QPixmap.fromImage(img), i))
                    
                    self.current_pdf_text = text
                    self.current_zoom = 600 
                    self.clear_preview()
                    self.refresh_pdf_view() 
                else:
                    self.scroll_area.hide()
                    self.text_preview.show()
                    with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                        self.current_pdf_text = f.read()
                    self.text_preview.setText(self.current_pdf_text)

                self.check_upload.setChecked(True)
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

    # ------------------------------------------
    # --- UPGRADED PDF EXPORTER (SUPPORTS IMAGES) ---
    # ------------------------------------------
    def export_pdf(self):
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "Final_Report.pdf", "PDF (*.pdf)")
        if save_path:
            try:
                # Use QPrinter to natively render text, HTML, and pasted images!
                printer = QPrinter(QPrinter.PrinterMode.HighResolution)
                printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
                printer.setOutputFileName(save_path)
                
                # Combine everything into a beautiful HTML format for the printer
                ai_html = self.current_ai_summary.replace('\n', '<br>')
                manual_html = self.manual_input.toHtml()
                
                final_html = f"""
                <h1 style="color: #003366; text-align: center;">Document Analysis Report</h1>
                <hr>
                <h2 style="color: #003366;">1. AI Automated Summary</h2>
                <p style="font-size: 14px; line-height: 1.6;">{ai_html}</p>
                <br>
                <h2 style="color: #003366;">2. Manual Notes & Flowcharts</h2>
                {manual_html}
                """
                
                doc = QTextDocument()
                doc.setHtml(final_html)
                doc.print(printer)
                
                QMessageBox.information(self, "Success", "PDF with flowcharts saved successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", str(e))

    def export_docx(self):
        save_path, _ = QFileDialog.getSaveFileName(self, "Save Word Document", "Task_Directive_Report.docx", "Word (*.docx)")
        if not save_path:
            return

        try:
            from docx import Document
            from docx.shared import Pt, Inches
            from docx.enum.text import WD_ALIGN_PARAGRAPH
            from docx.enum.section import WD_SECTION
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn

            # --- INTERNAL HELPERS ---
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

            # --- DOCUMENT START ---
            doc = Document()
            style = doc.styles['Normal']
            style.font.name = 'Calibri'
            style.font.size = Pt(11)

            # --- SECTION 1: COVER PAGE ---
            section1 = doc.sections[0]
            section1.top_margin = Inches(1.5)
            section1.bottom_margin = Inches(1.5)
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

            doc.add_paragraph("\nPROJECT NAME: __________________________________________")
            p = doc.add_paragraph("\n\n[ LOGO HERE ]")
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("ABC Organization").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("Address of organization").alignment = WD_ALIGN_PARAGRAPH.CENTER

            # --- SECTION 2: MAIN CONTENT ---
            section2 = doc.add_section(WD_SECTION.NEW_PAGE)
            section2.header.is_linked_to_previous = False
            section2.footer.is_linked_to_previous = False
            remove_page_border(section2)

            # --- DATA INJECTION ---
            # 1. Introduction (From AI Summary)
            add_heading(doc, "1. Introduction [Automated]")
            doc.add_paragraph(self.current_ai_summary)

            add_heading(doc, "2. Reference [Default]")
            doc.add_paragraph("__________________________________________________")

            add_heading(doc, "3. Basis Of Task Directive [Default]")
            doc.add_paragraph("__________________________________________________")

            add_heading(doc, "4. Scope Of Task Directive")
            doc.add_paragraph("To assign the certification respectively")

            # 5. Stakeholders (From Manual Input)
            add_heading(doc, "5. Stakeholders [Extracted Data]")
            doc.add_paragraph(self.manual_input.toPlainText())

            # Building the tables as per your template
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            for i, h in enumerate(["Sl No.", "Organisation", "Role", "Activities"]):
                hdr_cells[i].text = h
            
            # --- THE REST OF YOUR TEMPLATE ---
            add_heading(doc, "6. Certification Work Breakdown")
            doc.add_paragraph("______________________________________________")

            add_heading(doc, "7. Task Allocation")
            doc.add_paragraph("______________________________________________")

            add_heading(doc, "11. Distribution List")
            doc.add_paragraph("11.1 External Organization\n11.2 Internal Distribution")

            # --- ANNEXURES ---
            doc.add_page_break()
            add_heading(doc, "Annexure-1")
            doc.add_paragraph("Work Assignment List of LRUs")
            
            # Save the final file
            doc.save(save_path)
            QMessageBox.information(self, "Success", f"Professional Task Directive template saved to:\n{save_path}")

        except Exception as e:
            QMessageBox.critical(self, "Word Export Error", f"Failed to generate template: {str(e)}")
            QMessageBox.critical(self, "Export Error", str(e))

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = OfflineApp()
    window.show()
    sys.exit(app.exec())
