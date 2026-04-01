import sys
import re
import fitz  # PyMuPDF
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout,
    QHBoxLayout, QPushButton, QTextEdit, QFileDialog,
    QLabel, QSplitter, QMessageBox, QScrollArea, QProgressBar,
    QCheckBox, QGroupBox, QDialog, QRubberBand, QMenu
)
from PyQt6.QtGui import QPixmap, QImage, QFont, QTextDocument, QTextCursor
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QEvent, QRect, QPoint, QSize
from PyQt6.QtPrintSupport import QPrinter

# --- WORD TEMPLATE IMPORTS ---
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ==========================================
# --- PURE PYTHON SUMMARIZER (NO AI NEEDED) ---
# ==========================================
def simple_summarize(text, num_sentences=4):
    """A standard NLP extractive summarizer using word frequencies."""
    if not text: return ""
    
    stop_words = {"the", "is", "in", "and", "to", "of", "a", "for", "on", "with", "as", "by", "this", "that", "it", "are", "be", "or", "an", "at", "from", "which", "will"}
    
    # Split text into sentences
    sentences = re.split(r'(?<=[.!?]) +|\n+', text)
    sentences = [s.strip() for s in sentences if len(s.strip()) > 5]
    
    if len(sentences) <= num_sentences:
        return " ".join(sentences)
        
    # Calculate word frequencies
    words = re.findall(r'\b\w+\b', text.lower())
    freq = {}
    for w in words:
        if w not in stop_words and not w.isnumeric():
            freq[w] = freq.get(w, 0) + 1
            
    # Score sentences based on high-frequency words
    scores = {}
    for i, s in enumerate(sentences):
        score = 0
        s_words = re.findall(r'\b\w+\b', s.lower())
        for w in s_words:
            if w in freq:
                score += freq[w]
        # Normalize score by sentence length to balance long/short sentences
        scores[i] = score / max(len(s_words), 1)
        
    # Get the top sentences in their original order
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
# --- INTERACTIVE PDF SNIPPING TOOL ---
# ==========================================
class PdfPageLabel(QLabel):
    text_extracted = pyqtSignal(str)
    image_extracted = pyqtSignal(QPixmap)
    summary_requested = pyqtSignal(str) # New signal for the summarizer!

    def __init__(self, page, parent=None):
        super().__init__(parent)
        self.page = page  
        self.rubber_band = QRubberBand(QRubberBand.Shape.Rectangle, self)
        self.origin = QPoint()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.origin = event.pos()
            self.rubber_band.setGeometry(QRect(self.origin, QSize()))
            self.rubber_band.show()

    def mouseMoveEvent(self, event):
        if not self.origin.isNull():
            self.rubber_band.setGeometry(QRect(self.origin, event.pos()).normalized())

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            rect = self.rubber_band.geometry()
            if rect.width() > 10 and rect.height() > 10:
                self.show_context_menu(event.pos(), rect)
            else:
                self.rubber_band.hide()
            self.origin = QPoint()
            
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
        copy_img_action = menu.addAction("🖼️ Extract as Image")
        copy_text_action = menu.addAction("📝 Extract Text to Notes")
        menu.addSeparator()
        # The new Summarization button!
        summarize_action = menu.addAction("✨ Summarize Selection (Set as Intro)")

        action = menu.exec(self.mapToGlobal(pos))
        
        if action == copy_img_action:
            if self.pixmap():
                self.image_extracted.emit(self.pixmap().copy(rect)) 
        elif action == copy_text_action:
            text = self.get_text_from_rect(rect)
            if text: self.text_extracted.emit(text) 
        elif action == summarize_action:
            text = self.get_text_from_rect(rect)
            if text: self.summary_requested.emit(text)
            
        self.rubber_band.hide()

# ==========================================
# --- MAIN APPLICATION ---
# ==========================================
class OfflineApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Offline Document Assistant - Template Edition")
        self.resize(1400, 900)
        
        self.file_path = ""
        self.doc = None
        self.current_pdf_text = ""
        self.current_summary = "No introduction summary set yet. Select text on the PDF and right-click to summarize it!"
        self.page_data = [] 
        self.current_zoom = 600 
        self.is_maximized = False
        
        self.init_ui()

    def init_ui(self):
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        layout = QVBoxLayout(main_widget)
        self.splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- LEFT SIDE: Document Viewer ---
        left_container = QWidget()
        left_layout = QVBoxLayout(left_container)
        left_controls_layout = QHBoxLayout()
        
        self.upload_btn = QPushButton(" 📂 Upload PDF Document")
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
        self.page_layout.setAlignment(Qt.AlignmentFlag.AlignTop) 
        self.scroll_area.setWidget(self.page_container)
        self.scroll_area.setAlignment(Qt.AlignmentFlag.AlignCenter) 
        self.scroll_area.viewport().installEventFilter(self)

        self.text_preview = QTextEdit()
        self.text_preview.setReadOnly(True)
        self.text_preview.hide()

        left_layout.addLayout(left_controls_layout)
        tip_label = QLabel("💡 Tip: Click and drag on the PDF to extract text, flowcharts, or generate your Introduction Summary!")
        tip_label.setStyleSheet("color: #0056b3; font-style: italic;")
        left_layout.addWidget(tip_label)
        left_layout.addWidget(self.scroll_area)
        left_layout.addWidget(self.text_preview)

        # --- RIGHT SIDE: Tools & Extraction ---
        self.right_container = QWidget() 
        right_layout = QVBoxLayout(self.right_container)

        top_bar_layout = QHBoxLayout()
        self.status_group = QGroupBox("Status Bar")
        status_layout = QHBoxLayout() 
        self.check_upload = QCheckBox("File Uploaded")
        self.check_summary = QCheckBox("Intro Set")
        for cb in [self.check_upload, self.check_summary]:
            cb.setEnabled(False) 
            status_layout.addWidget(cb)
        self.status_group.setLayout(status_layout)
        top_bar_layout.addWidget(self.status_group)
        
        self.summary_icon_btn = QPushButton("📑")
        self.summary_icon_btn.setFixedSize(65, 65)
        self.summary_icon_btn.setToolTip("View Current Introduction Summary")
        self.summary_icon_btn.setStyleSheet("""
            QPushButton { font-size: 30px; border-radius: 32px; background-color: #f8f9fa; border: 2px solid #ced4da; }
            QPushButton:hover { background-color: #e2e6ea; }
        """)
        self.summary_icon_btn.clicked.connect(self.show_summary_popup)
        top_bar_layout.addWidget(self.summary_icon_btn, alignment=Qt.AlignmentFlag.AlignTop)

        self.status_label = QLabel("Ready")

        self.manual_input = NotesEditor()
        self.manual_input.setPlaceholderText("Extracted text, Stakeholders, and flowcharts will appear here...")
        self.manual_input.setStyleSheet("background-color: white; font-size: 13px;")

        btn_layout = QHBoxLayout()
        self.save_pdf_btn = QPushButton(" 💾 View Basic PDF")
        self.save_docx_btn = QPushButton(" 📝 Generate Task Directive (Word)")
        self.save_docx_btn.setStyleSheet("background-color: #2b579a; color: white; font-weight: bold;")
        self.save_pdf_btn.setFixedHeight(45)
        self.save_docx_btn.setFixedHeight(45)
        self.save_pdf_btn.clicked.connect(self.export_pdf)
        self.save_docx_btn.clicked.connect(self.export_docx)
        btn_layout.addWidget(self.save_pdf_btn)
        btn_layout.addWidget(self.save_docx_btn)

        right_layout.addLayout(top_bar_layout)
        right_layout.addWidget(self.status_label)
        right_layout.addWidget(QLabel(" ✍️ Manual Notes & Extracted Data:"))
        right_layout.addWidget(self.manual_input)
        right_layout.addLayout(btn_layout)

        self.splitter.addWidget(left_container)
        self.splitter.addWidget(self.right_container)
        self.splitter.setStretchFactor(0, 5) 
        self.splitter.setStretchFactor(1, 5) 
        layout.addWidget(self.splitter)

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
        
    def generate_summary_from_selection(self, text):
        """Runs the offline Python summarizer on selected text."""
        self.status_label.setText("Summarizing selection...")
        QApplication.processEvents()
        
        # Call our standard code summarizer logic
        summary = simple_summarize(text, num_sentences=4)
        
        if summary:
            self.current_summary = summary
            self.status_label.setText("Summary Set")
            self.check_summary.setChecked(True)
            self.summary_icon_btn.setStyleSheet("""
                QPushButton { font-size: 30px; border-radius: 32px; background-color: #d4edda; border: 2px solid #28a745; }
                QPushButton:hover { background-color: #c3e6cb; }
            """)
            QMessageBox.information(self, "Success", "Selection summarized and set as Introduction for your template!")
        else:
            self.status_label.setText("Ready")

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
                lbl = PdfPageLabel(page)
                
                # Connect all our snipping signals
                lbl.text_extracted.connect(self.add_extracted_text)
                lbl.image_extracted.connect(self.add_extracted_image)
                lbl.summary_requested.connect(self.generate_summary_from_selection)
                
                lbl.setStyleSheet("background-color: white; border: 1px solid black; margin-bottom: 10px;")
                self.page_layout.addWidget(lbl)

        for i in range(self.page_layout.count()):
            lbl = self.page_layout.itemAt(i).widget()
            if isinstance(lbl, PdfPageLabel) and i < len(self.page_data):
                scaled_pixmap = self.page_data[i][0].scaledToWidth(self.current_zoom, Qt.TransformationMode.SmoothTransformation)
                lbl.setPixmap(scaled_pixmap)

    # ------------------------------------------
    # --- CORE LOGIC & FUNCTIONS ---
    # ------------------------------------------
    def upload_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Open Document", "", "Documents (*.pdf *.txt)")
        if path:
            self.file_path = path
            self.check_upload.setChecked(False)
            self.check_summary.setChecked(False)
            self.current_summary = "No introduction summary set yet. Select text on the PDF and right-click to summarize it!"
            self.summary_icon_btn.setStyleSheet("""
                QPushButton { font-size: 30px; border-radius: 32px; background-color: #f8f9fa; border: 2px solid #ced4da; }
                QPushButton:hover { background-color: #e2e6ea; }
            """)
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
                self.status_label.setText("Ready. Draw a box over text to summarize!")
                
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Could not load file: {str(e)}")

    def show_summary_popup(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("📑 Document Summary")
        dialog.resize(700, 500)
        layout = QVBoxLayout(dialog)
        text_edit = QTextEdit()
        text_edit.setReadOnly(True)
        text_edit.setText(self.current_summary)
        text_edit.setStyleSheet("font-size: 14px; line-height: 1.6; padding: 10px;")
        close_btn = QPushButton("Close")
        close_btn.setFixedHeight(40)
        close_btn.clicked.connect(dialog.accept)
        layout.addWidget(text_edit)
        layout.addWidget(close_btn)
        dialog.exec()

    # ------------------------------------------
    # --- TEMPLATE INJECTION & EXPORT ---
    # ------------------------------------------
    def export_pdf(self):
        save_path, _ = QFileDialog.getSaveFileName(self, "Save PDF", "Basic_Report.pdf", "PDF (*.pdf)")
        if save_path:
            try:
                printer = QPrinter(QPrinter.PrinterMode.HighResolution)
                printer.setOutputFormat(QPrinter.OutputFormat.PdfFormat)
                printer.setOutputFileName(save_path)
                ai_html = self.current_summary.replace('\n', '<br>')
                manual_html = self.manual_input.toHtml()
                final_html = f"<h1>Document Analysis Report</h1><hr><h2>1. Introduction Summary</h2><p>{ai_html}</p><br><h2>2. Manual Notes & Flowcharts</h2>{manual_html}"
                doc = QTextDocument()
                doc.setHtml(final_html)
                doc.print(printer)
                QMessageBox.information(self, "Success", "Basic PDF saved successfully!")
            except Exception as e:
                QMessageBox.critical(self, "Export Error", str(e))

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

            # --- SECTION 1: COVER PAGE ---
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

            doc.add_paragraph("\nPROJECT NAME: __________________________________________")
            p = doc.add_paragraph("\n\n[ LOGO HERE ]"); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("ABC Organization").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("Address of organization").alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph("Organization details").alignment = WD_ALIGN_PARAGRAPH.CENTER

            # --- SECTION 2: MAIN CONTENT ---
            section2 = doc.add_section(WD_SECTION.NEW_PAGE)
            section2.header.is_linked_to_previous = False; section2.footer.is_linked_to_previous = False
            section2.top_margin = Inches(0.75); section2.bottom_margin = Inches(0.75)
            section2.left_margin = Inches(0.75); section2.right_margin = Inches(0.75)
            remove_page_border(section2)

            # --- INJECTING YOUR DATA ---
            add_heading(doc, "1. Introduction [Automated]")
            # Injects your standard code summarized introduction
            doc.add_paragraph(self.current_summary)

            add_heading(doc, "2. Reference [Default]")
            doc.add_paragraph("__________________________________________________")

            add_heading(doc, "3. Basis Of Task Directive [Default]")
            doc.add_paragraph("__________________________________________________")

            add_heading(doc, "4. Scope Of Task Directive [Drop down menu]")
            doc.add_paragraph("To assign the certification respectively")
            doc.add_paragraph("1) ____\n2) ____\n3) ____")

            add_heading(doc, "5. Stakeholders [Automated + Manual Notes]")
            doc.add_paragraph("The following are the major stakeholders and manual notes extracted:")
            # Injects your manual snips here
            doc.add_paragraph(self.manual_input.toPlainText()) 

            table = doc.add_table(rows=4, cols=4); table.style = 'Table Grid'
            for i, h in enumerate(["Sl No.", "Organisation", "Role", "Activities"]): table.cell(0, i).text = h

            add_heading(doc, "6. Certification Work Breakdown [Drop down menu]")
            doc.add_paragraph("______________________________________________")

            add_heading(doc, "7. Task Allocation [Default]")
            doc.add_paragraph("______________________________________________")
            add_heading(doc, "7.1 Coordinating Directorate [Default]")
            doc.add_paragraph("______________________________________________")
            add_heading(doc, "7.2 Single Point of Contact (SPoC) [Default]")
            doc.add_paragraph("______________________________________________")

            add_heading(doc, "7.3 Certification Task Allocation")
            table = doc.add_table(rows=4, cols=4); table.style = 'Table Grid'
            for i, h in enumerate(["Sl No.", "Certification Activity", "Certification Work Centre", "Responsible Head"]): table.cell(0, i).text = h

            add_heading(doc, "7.4 Issue of Clearance [Default]")
            doc.add_paragraph("a. ___\nb. ___\nc. ___\nd. ___\ne. ___\nf. ___\ng. ___")

            add_heading(doc, "8. SCRB And TARB [Default]")
            doc.add_paragraph("______________________________________________")
            add_heading(doc, "9. Communication [Default]")
            doc.add_paragraph("______________________________________________")
            add_heading(doc, "10. Certification Progress Review [Default]")
            doc.add_paragraph("______________________________________________\n\n(__________)")

            add_heading(doc, "11. Distribution List")
            doc.add_paragraph("11.1 External Organization\n1. ___\n2. ___\n3. ___")
            doc.add_paragraph("11.2 Internal Distribution")

            # --- ANNEXURES ---
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
