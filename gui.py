from PyQt5.QtWidgets import QMainWindow, QComboBox, QLabel, QVBoxLayout, QWidget, QTextEdit
from PyQt5.QtCore import Qt
from powerpointfont import set_uniform_font, analyze_fonts


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
        self.font_combo.addItems(
            ['Arial', 'Calibri', 'Times New Roman', 'Helvetica', 'Verdana'])
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

        # Analyze fonts
        font_analysis = analyze_fonts(file_path)

        # Process the file
        changes_made, fonts_changed = set_uniform_font(
            file_path, output_file, target_font)

        # Prepare output text
        output_text = f"File processed: {file_path}\n"
        output_text += f"Target font: {target_font}\n\n"

        output_text += "Font Analysis:\n"
        for font, count in font_analysis.items():
            output_text += f"  - {font}: {count} occurrence{'s' if count > 1 else ''}\n"

        output_text += "\nFont Uniformization:\n"
        if changes_made:
            output_text += f"Fonts updated to '{target_font}'. Saved as '{output_file}'\n"
            output_text += "Font changes summary:\n"
            for font, count in fonts_changed.items():
                output_text += f"  - {font} to {target_font}: {count} occurrence{'s' if count > 1 else ''}\n"
            if "Default font" in fonts_changed:
                output_text += "  Note: Some 'Default font' text was changed to the target font.\n"
        else:
            output_text += f"No font changes were necessary. All fonts are already '{target_font}'.\n"

        self.output_text.setText(output_text)
