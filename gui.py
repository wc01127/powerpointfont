from PyQt5.QtWidgets import QMainWindow, QComboBox, QLabel, QVBoxLayout, QWidget, QTextEdit
from PyQt5.QtCore import Qt
from powerpoint_processor import process_powerpoint

class PowerPointFontChanger(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('PowerPoint Font Changer')
        self.setGeometry(100, 100, 600, 400)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        layout = QVBoxLayout()

        # Font selection dropdown
        self.font_combo = QComboBox()
        self.font_combo.addItems(['Arial', 'Calibri', 'Times New Roman', 'Helvetica', 'Verdana'])
        layout.addWidget(QLabel('Select Font:'))
        layout.addWidget(self.font_combo)

        # Drag and drop area
        self.drop_label = QLabel('Drag and drop PowerPoint file here')
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setStyleSheet('border: 2px dashed #aaa')
        layout.addWidget(self.drop_label)

        # Output area
        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        layout.addWidget(QLabel('Output:'))
        layout.addWidget(self.output_text)

        central_widget.setLayout(layout)

        self.setAcceptDrops(True)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for file_path in files:
            if file_path.lower().endswith('.pptx'):
                self.process_file(file_path)

    def process_file(self, file_path):
        output_file = file_path.rsplit('.', 1)[0] + '_updated.pptx'
        target_font = self.font_combo.currentText()
        
        percent_changed, slides_changed = process_powerpoint(file_path, output_file, target_font)
        
        # Capture the printed output
        import io
        import sys
        captured_output = io.StringIO()
        sys.stdout = captured_output
        process_powerpoint(file_path, output_file, target_font)
        sys.stdout = sys.__stdout__
        debug_info = captured_output.getvalue()
        
        output_text = f"File processed: {file_path}\n"
        output_text += f"Output file: {output_file}\n"
        output_text += f"Percentage of text changed: {percent_changed:.2f}%\n"
        if slides_changed:
            output_text += f"Slides changed: {', '.join(map(str, sorted(slides_changed)))}\n"
        else:
            output_text += "No slides were changed.\n"
        
        output_text += "\nDetailed Debug Information:\n"
        output_text += debug_info
        
        self.output_text.setText(output_text)