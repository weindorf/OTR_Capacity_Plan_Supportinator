import sys
import os
from PyQt6.QtWidgets import QApplication
from .main_window import MainWindow
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())

if __name__ == '__main__':
    main()
