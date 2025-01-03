import sys
import os
from PyQt6.QtWidgets import QApplication
from .main_window import MainWindow
sys.path.insert(0, os.path.abspath(os.path.dirname(__file__)))

def main():
    print("Starting application...")
    app = QApplication(sys.argv)
    print("Created QApplication")
    
    try:
        main_window = MainWindow()
        print("Created MainWindow")
        main_window.show()
        print("Showed MainWindow")
        sys.exit(app.exec())
    except Exception as e:
        print(f"Error occurred: {e}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
