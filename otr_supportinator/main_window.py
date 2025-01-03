import os
import tempfile
from PyQt6.QtWidgets import QMainWindow, QTabWidget, QVBoxLayout, QWidget, QStatusBar, QMenuBar, QMenu, QMessageBox, QApplication, QSizePolicy
from PyQt6.QtGui import QAction
from .tabs.summary_file_generator_tab import SummaryFileGeneratorTab
from .tabs.pop_tab import PopTab
from .tabs.summary_file_combiner_tab import SummaryFileCombinerTab

class MainWindow(QMainWindow):
    def __init__(self):
        print("Initializing MainWindow...")
        super().__init__()
        self.temp_dir = tempfile.mkdtemp()
        print("Created temp directory")
        self.setWindowTitle("OTR Capacity Plan Upload Supportinator")
        self.setGeometry(100, 100, 900, 1000)
        self.setMinimumWidth(900)
        print("Set window properties")

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        print("Set up central widget and layout")

        self.tab_widget = QTabWidget()
        self.main_layout.addWidget(self.tab_widget)
        print("Added tab widget")

        try:
            print("Creating SummaryFileGeneratorTab...")
            self.summary_file_generator_tab = SummaryFileGeneratorTab(self)
            print("Created SummaryFileGeneratorTab")
        except Exception as e:
            print(f"Error creating SummaryFileGeneratorTab: {e}")
            import traceback
            traceback.print_exc()

        try:
            print("Creating PopTab...")
            self.pop_tab = PopTab(self)
            print("Created PopTab")
        except Exception as e:
            print(f"Error creating PopTab: {e}")
            import traceback
            traceback.print_exc()

        try:
            print("Creating SummaryFileCombinerTab...")
            self.summary_file_combiner_tab = SummaryFileCombinerTab(self)
            print("Created SummaryFileCombinerTab")
        except Exception as e:
            print(f"Error creating SummaryFileCombinerTab: {e}")
            import traceback
            traceback.print_exc()

        self.tab_widget.addTab(self.summary_file_generator_tab, "Summary File Generator")
        self.tab_widget.addTab(self.pop_tab, "PoP")
        self.tab_widget.addTab(self.summary_file_combiner_tab, "Summary File Combiner")
        print("Added tabs to tab widget")

        # Ensure each tab takes up all available space
        for i in range(self.tab_widget.count()):
            self.tab_widget.widget(i).setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)

        self.create_menu_bar()

        self.temp_dir = tempfile.mkdtemp()

        print("MainWindow initialization complete")

    def closeEvent(self, event):
        # Clean up the temporary directory when the application closes
        for filename in os.listdir(self.temp_dir):
            file_path = os.path.join(self.temp_dir, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    os.rmdir(file_path)
            except Exception as e:
                print(f'Failed to delete {file_path}. Reason: {e}')
        os.rmdir(self.temp_dir)
        super().closeEvent(event)

    def create_menu_bar(self):
        menu_bar = QMenuBar(self)
        self.setMenuBar(menu_bar)
        
        file_menu = QMenu("File", self)
        menu_bar.addMenu(file_menu)
        
        restart_action = QAction("Restart", self)
        restart_action.triggered.connect(self.restart)
        file_menu.addAction(restart_action)
        
        quit_action = QAction("Quit", self)
        quit_action.triggered.connect(self.quit)
        file_menu.addAction(quit_action)

    def restart(self):
        self.summary_file_generator_tab.restart()
        self.pop_tab.restart()
        self.summary_file_combiner_tab.restart()
        self.statusBar.showMessage("Application restarted", 5000)

    def quit(self):
        reply = QMessageBox.question(self, 'Quit', 'Do you want to quit?',
                                     QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                     QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            self.clean_up_temp_files()
            QApplication.instance().quit()

    def clean_up_temp_files(self):
        try:
            if os.path.exists(self.temp_dir):
                for file in os.listdir(self.temp_dir):
                    file_path = os.path.join(self.temp_dir, file)
                    if os.path.isfile(file_path):
                        os.unlink(file_path)
                os.rmdir(self.temp_dir)
        except Exception as e:
            print(f"Error while cleaning up temporary files: {str(e)}")
