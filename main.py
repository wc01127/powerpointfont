import sys
from PyQt5.QtWidgets import QApplication
from gui import PowerPointFontChanger

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = PowerPointFontChanger()
    ex.show()
    sys.exit(app.exec_())