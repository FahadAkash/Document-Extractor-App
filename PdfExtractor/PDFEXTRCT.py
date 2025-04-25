import sys
import os
import fitz  # PyMuPDF
from PIL import Image
from PyQt5.QtCore import Qt, QSize, QEvent, QPoint
from PyQt5.QtGui import QPixmap, QImage, QIcon, QFont, QColor, QPainter
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QListWidget, QVBoxLayout, QWidget,
    QLineEdit, QPushButton, QFileDialog, QLabel, QListWidgetItem,
    QMessageBox, QStatusBar, QHBoxLayout, QSizePolicy
)

# ======================
# CONSTANTS & STYLING
# ======================
STYLESHEET = """
    QMainWindow {
        background-color: #f5f7fa;
        font-family: 'Segoe UI';
    }
    QListWidget {
        background-color: white;
        border-radius: 10px;
        border: 1px solid #e0e0e0;
        padding: 10px;
    }
    QPushButton {
        background-color: #4a6baf;
        color: white;
        border: none;
        padding: 8px 16px;
        border-radius: 6px;
        font-size: 14px;
        min-width: 120px;
    }
    QPushButton:hover {
        background-color: #3a5a9f;
    }
    QPushButton:pressed {
        background-color: #2a4a8f;
    }
    QLineEdit {
        padding: 8px 12px;
        border: 1px solid #ddd;
        border-radius: 6px;
        font-size: 14px;
    }
    QLabel {
        font-size: 14px;
        color: #333;
    }
    QStatusBar {
        background-color: #e9ecef;
        color: #495057;
    }
    #placeholderLabel {
        color: #666;
        font-size: 24px;
        qproperty-alignment: AlignCenter;
    }
    QListWidget::item {
        color: #333;
        border: 2px solid transparent;
        border-radius: 8px;
        padding: 5px;
    }
    QListWidget::item:hover {
        border-color: #4a90e2;
        background-color: #f8f9fa;
        color: #333;
    }
    QListWidget::item:selected {
        background-color: #e3f2fd;
        border-color: #1a73e8;
        color: #222;
    }
    QListWidget::item:selected:!active {
        color: #222;
    }
"""

EMOJIS = {
    "folder": "📁",
    "pdf": "📄",
    "extract": "✂️",
    "select": "🔖",
    "image": "🖼️",
    "error": "❌",
    "success": "✅",
    "warning": "⚠️",
    "pdf_export": "📑",
    "doc": "📝"
}

# ======================
# UTILITY FUNCTIONS
# ======================
def show_error(message, parent=None):
    msg = QMessageBox(parent)
    msg.setIcon(QMessageBox.Critical)
    msg.setText(f"{EMOJIS['error']} Error")
    msg.setInformativeText(message)
    msg.setWindowTitle("Error")
    msg.exec_()

def show_success(message, parent=None):
    msg = QMessageBox(parent)
    msg.setIcon(QMessageBox.Information)
    msg.setText(f"{EMOJIS['success']} Success")
    msg.setInformativeText(message)
    msg.setWindowTitle("Success")
    msg.exec_()

def parse_page_range(range_text, total_pages):
    """Parse page range input (e.g., '1,3-5,7') into a list of page numbers."""
    pages = set()
    parts = range_text.split(',')
    for part in parts:
        part = part.strip()
        if not part:
            continue
        if '-' in part:
            start, end = part.split('-')
            try:
                start = int(start.strip())
                end = int(end.strip())
                if start > end or start < 1 or end > total_pages:
                    continue
                pages.update(range(start, end + 1))
            except ValueError:
                continue
        else:
            try:
                page = int(part)
                if 1 <= page <= total_pages:
                    pages.add(page)
            except ValueError:
                continue
    return sorted(pages)

# ======================
# CUSTOM WIDGETS
# ======================
class ThumbnailListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setViewMode(QListWidget.IconMode)
        self.setIconSize(QSize(150, 150))
        self.setResizeMode(QListWidget.Adjust)
        self.setSelectionMode(QListWidget.MultiSelection)
        self.setSpacing(15)
        self.setStyleSheet("""
            QListWidget::item {
                color: #333;
                border: 2px solid transparent;
                border-radius: 8px;
                padding: 5px;
            }
            QListWidget::item:hover {
                border-color: #4a90e2;
                background-color: #f8f9fa;
                color: #333;
            }
            QListWidget::item:selected {
                background-color: #e3f2fd;
                border-color: #1a73e8;
                color: #222;
            }
            QListWidget::item:selected:!active {
                color: #222;
            }
        """)
        
        # Hover zoom label
        self.zoom_label = QLabel(self)
        self.zoom_label.setWindowFlags(Qt.ToolTip)
        self.zoom_label.setStyleSheet("""
            background-color: white;
            border: 2px solid #4a90e2;
            border-radius: 4px;
            padding: 5px;
        """)
        self.zoom_label.hide()
        
        # Placeholder label
        self.placeholder_label = QLabel(f"{EMOJIS['pdf']} Drag PDF or DOC/DOCX here", self)
        self.placeholder_label.setObjectName("placeholderLabel")
        self.placeholder_label.setAlignment(Qt.AlignCenter)
        self.placeholder_label.show()
        
        # Enable mouse tracking
        self.viewport().installEventFilter(self)
        self.viewport().setMouseTracking(True)

    def resizeEvent(self, event):
        self.placeholder_label.setGeometry(0, 0, self.width(), self.height())
        super().resizeEvent(event)

    def load_pages(self, pdf_doc):
        """Load PDF pages as thumbnails with visible text."""
        self.clear()
        if pdf_doc:
            for page_num in range(len(pdf_doc)):
                try:
                    page = pdf_doc[page_num]
                    pix = page.get_pixmap(matrix=fitz.Matrix(0.25, 0.25))
                    img_data = {
                        "samples": pix.samples,
                        "width": pix.width,
                        "height": pix.height,
                        "stride": pix.stride,
                    }
                    qimg = QImage(
                        pix.samples, 
                        pix.width, 
                        pix.height, 
                        pix.stride, 
                        QImage.Format_RGB888
                    )
                    pixmap = QPixmap.fromImage(qimg)

                    item = QListWidgetItem(f"Page {page_num + 1}")
                    item.setIcon(QIcon(pixmap))
                    item.setTextAlignment(Qt.AlignCenter)  # Center-align text
                    item.setFlags(item.flags() | Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    item.setData(Qt.UserRole, (page_num, img_data))
                    self.addItem(item)
                except Exception as e:
                    print(f"Error loading page {page_num + 1}: {str(e)}")

        self.placeholder_label.setVisible(self.count() == 0)

    # Update the eventFilter to recreate QPixmap from raw data:
    def eventFilter(self, obj, event):
        if obj == self.viewport():
            if event.type() == QEvent.MouseMove:
                pos = event.pos()
                item = self.itemAt(pos)
                if item:
                    page_num, img_data = item.data(Qt.UserRole)
                    # Recreate QImage from raw data
                    qimg = QImage(
                        img_data['samples'],
                        img_data['width'],
                        img_data['height'],
                        img_data['stride'],
                        QImage.Format_RGB888
                    )
                    original_pixmap = QPixmap.fromImage(qimg)
                
                    # Rest of the zoom code remains the same
                    zoomed_pixmap = original_pixmap.scaled(
                        original_pixmap.width() * 2,
                        original_pixmap.height() * 2,
                        Qt.KeepAspectRatio,
                        Qt.SmoothTransformation
                    )
                    self.zoom_label.setPixmap(zoomed_pixmap)
                    self.zoom_label.adjustSize()
                    self.zoom_label.move(self.mapToGlobal(pos) + QPoint(20, 20))
                    self.zoom_label.show()
                else:
                    self.zoom_label.hide()
        elif event.type() == QEvent.Leave:
            self.zoom_label.hide()
        return super().eventFilter(obj, event)

# ======================
# MAIN APPLICATION
# ======================
class DocumentExtractorApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf_doc = None
        self.current_file = None
        self.init_ui()
        self.init_styles()

    def init_ui(self):
        """Initialize the main UI components."""
        self.setWindowTitle(f"{EMOJIS['pdf']} Document Extractor")
        self.setGeometry(100, 100, 1100, 800)
        self.setAcceptDrops(True)

        # Central Widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # Main Layout
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)

        # Document Thumbnail List
        self.page_list = ThumbnailListWidget(self)
        main_layout.addWidget(self.page_list)

        # Controls Layout
        controls_layout = QHBoxLayout()

        # Page Range Input
        self.range_input = QLineEdit()
        self.range_input.setPlaceholderText(f"{EMOJIS['select']} Example: 1,3-5,7")
        controls_layout.addWidget(self.range_input, 3)

        # Select Button
        self.select_btn = QPushButton(f"{EMOJIS['select']} Select Pages")
        self.select_btn.clicked.connect(self.select_pages)
        controls_layout.addWidget(self.select_btn)

        main_layout.addLayout(controls_layout)

        # Extraction Buttons
        btn_layout = QHBoxLayout()
        self.extract_btn = QPushButton(f"{EMOJIS['image']} Extract as Single Image")
        self.extract_btn.clicked.connect(self.extract_pages)
        btn_layout.addWidget(self.extract_btn)

        self.extract_separate_btn = QPushButton(f"{EMOJIS['image']} Extract as Separate Images")
        self.extract_separate_btn.clicked.connect(self.extract_separate_pages)
        btn_layout.addWidget(self.extract_separate_btn)

        self.extract_pdf_btn = QPushButton(f"{EMOJIS['pdf_export']} Save as PDF")
        self.extract_pdf_btn.clicked.connect(self.extract_as_pdf)
        btn_layout.addWidget(self.extract_pdf_btn)

        main_layout.addLayout(btn_layout)

        # Output Folder Section
        folder_layout = QHBoxLayout()
        self.folder_path = QLineEdit()
        self.folder_path.setPlaceholderText(f"{EMOJIS['folder']} Select output folder...")
        folder_layout.addWidget(self.folder_path, 3)

        self.browse_btn = QPushButton(f"{EMOJIS['folder']} Browse")
        self.browse_btn.clicked.connect(self.browse_folder)
        folder_layout.addWidget(self.browse_btn)

        main_layout.addLayout(folder_layout)

        # Status Bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Ready")

    def init_styles(self):
        """Apply stylesheet and fonts."""
        self.setStyleSheet(STYLESHEET)
        font = QFont("Segoe UI", 10)
        self.setFont(font)

    # ======================
    # CORE FUNCTIONALITY
    # ======================
    def dragEnterEvent(self, event):
        """Handle drag enter event for document files."""
        if event.mimeData().hasUrls():
            urls = event.mimeData().urls()
            if any(url.toLocalFile().lower().endswith(('.pdf', '.doc', '.docx')) for url in urls):
                event.acceptProposedAction()
            else:
                event.ignore()
        else:
            event.ignore()

    def dropEvent(self, event):
        """Handle drop event to load document."""
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith('.pdf'):
                self.load_pdf(file_path)
            elif file_path.lower().endswith(('.doc', '.docx')):
                self.convert_doc_to_pdf(file_path)

    def load_pdf(self, file_path):
        """Load and display PDF pages as thumbnails."""
        try:
            self.pdf_doc = fitz.open(file_path)
            self.current_file = file_path
            self.page_list.load_pages(self.pdf_doc)
            self.status_bar.showMessage(f"{EMOJIS['pdf']} Loaded: {os.path.basename(file_path)}")
        except Exception as e:
            show_error(f"Failed to load PDF:\n{str(e)}", self)
            self.pdf_doc = None
            self.current_file = None

    def convert_doc_to_pdf(self, doc_path):
        """Convert DOC/DOCX to PDF using LibreOffice."""
        try:
            output_dir = os.path.dirname(doc_path)
            cmd = f'soffice --headless --convert-to pdf --outdir "{output_dir}" "{doc_path}"'
            os.system(cmd)
            
            pdf_path = os.path.splitext(doc_path)[0] + '.pdf'
            if os.path.exists(pdf_path):
                self.load_pdf(pdf_path)
                show_success(f"Converted to PDF: {os.path.basename(pdf_path)}", self)
            else:
                show_error("Conversion failed. Please install LibreOffice.", self)
        except Exception as e:
            show_error(f"Conversion error: {str(e)}", self)

    def select_pages(self):
        """Select pages based on range input."""
        if not self.pdf_doc:
            show_error("No document loaded!", self)
            return

        range_text = self.range_input.text().strip()
        if not range_text:
            show_error("Please enter a page range", self)
            return

        try:
            total_pages = len(self.pdf_doc)
            pages = parse_page_range(range_text, total_pages)
            if not pages:
                show_error("No valid pages in the specified range", self)
                return

            self.page_list.clearSelection()
            for page_num in pages:
                if 0 <= page_num - 1 < self.page_list.count():
                    item = self.page_list.item(page_num - 1)
                    item.setSelected(True)

            self.status_bar.showMessage(f"{EMOJIS['select']} Selected {len(pages)} pages")
        except Exception as e:
            show_error(f"Invalid page range:\n{str(e)}", self)

    def browse_folder(self):
        """Open folder dialog to select output directory."""
        folder = QFileDialog.getExistingDirectory(self, "Select Output Folder")
        if folder:
            self.folder_path.setText(folder)

    def extract_pages(self):
        """Extract selected pages as a single combined image."""
        if not self.pdf_doc:
            show_error("No document loaded!", self)
            return

        selected_items = self.page_list.selectedItems()
        if not selected_items:
            show_error("No pages selected!", self)
            return

        output_folder = self.folder_path.text().strip()
        if not output_folder:
            output_folder = os.path.join(os.getcwd(), "extracted_images")

        try:
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            page_nums = [self.page_list.row(item) for item in selected_items]
            images = []

            for page_num in page_nums:
                try:
                    page = self.pdf_doc[page_num]
                    pix = page.get_pixmap()
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    images.append(img)
                except Exception as e:
                    print(f"Error processing page {page_num + 1}: {str(e)}")

            if not images:
                show_error("No valid pages to extract", self)
                return

            # Calculate dimensions for combined image
            max_width = max(img.width for img in images)
            total_height = sum(img.height for img in images)

            # Create and save combined image
            combined = Image.new("RGB", (max_width, total_height))
            y_offset = 0
            for img in images:
                combined.paste(img, (0, y_offset))
                y_offset += img.height

            output_path = os.path.join(output_folder, "merged_pages.png")
            combined.save(output_path, "PNG")
            show_success(f"Saved combined image to:\n{output_path}", self)
            self.status_bar.showMessage(f"{EMOJIS['success']} Extraction complete!")
        except Exception as e:
            show_error(f"Error during extraction:\n{str(e)}", self)

    def extract_separate_pages(self):
        """Extract selected pages as separate image files."""
        if not self.pdf_doc:
            show_error("No document loaded!", self)
            return

        selected_items = self.page_list.selectedItems()
        if not selected_items:
            show_error("No pages selected!", self)
            return

        output_folder = self.folder_path.text().strip()
        if not output_folder:
            output_folder = os.path.join(os.getcwd(), "extracted_images")

        try:
            if not os.path.exists(output_folder):
                os.makedirs(output_folder)

            page_nums = [self.page_list.row(item) for item in selected_items]
            success_count = 0

            for page_num in page_nums:
                try:
                    page = self.pdf_doc[page_num]
                    pix = page.get_pixmap()
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    filename = f"page_{page_num + 1}.png"
                    img.save(os.path.join(output_folder, filename), "PNG")
                    success_count += 1
                except Exception as e:
                    print(f"Error extracting page {page_num + 1}: {str(e)}")

            if success_count > 0:
                show_success(f"Extracted {success_count} pages to:\n{output_folder}", self)
                self.status_bar.showMessage(f"{EMOJIS['success']} Extracted {success_count} pages")
            else:
                show_error("Failed to extract any pages", self)
        except Exception as e:
            show_error(f"Error during extraction:\n{str(e)}", self)

    def extract_as_pdf(self):
        """Export selected pages as new PDF file."""
        if not self.pdf_doc:
            show_error("No document loaded!", self)
            return

        selected_items = self.page_list.selectedItems()
        if not selected_items:
            show_error("No pages selected!", self)
            return

        path, _ = QFileDialog.getSaveFileName(
            self, "Save PDF", "", "PDF Files (*.pdf)")
        if not path:
            return

        try:
            new_pdf = fitz.open()
            for item in selected_items:
                page_num = self.page_list.row(item)
                new_pdf.insert_pdf(self.pdf_doc, from_page=page_num, to_page=page_num)
            new_pdf.save(path)
            show_success(f"PDF saved to:\n{path}", self)
            self.status_bar.showMessage(f"{EMOJIS['success']} PDF saved successfully")
        except Exception as e:
            show_error(f"PDF save failed: {str(e)}", self)

# ======================
# APPLICATION ENTRY
# ======================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = DocumentExtractorApp()
    window.show()
    sys.exit(app.exec_())