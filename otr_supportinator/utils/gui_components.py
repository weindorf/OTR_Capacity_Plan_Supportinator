# gui_components.py
import os
from PyQt6.QtWidgets import (QLabel, QListWidget, QVBoxLayout, QPushButton, 
                             QWidget, QListWidgetItem, QFileDialog, QMessageBox,
                             QProgressDialog, QDialog, QProgressBar)
from PyQt6.QtCore import Qt, pyqtSignal, QTimer
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QMouseEvent

class DropLabel(QLabel):
    file_dropped = pyqtSignal(str)

    def __init__(self, text, parent=None):
        super().__init__(text, parent)
        self.setAcceptDrops(True)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 5px;
                padding: 20px;
                background-color: #f0f0f0;
            }
            QLabel:hover {
                background-color: #e0e0e0;
                border-color: #999;
            }
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("""
                QLabel {
                    border: 2px solid #66a3ff;
                    border-radius: 5px;
                    padding: 20px;
                    background-color: #e6f2ff;
                }
            """)

    def dragLeaveEvent(self, event):
        self.reset_style()

    def dropEvent(self, event: QDropEvent):
        self.reset_style()
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        if files:
            self.file_dropped.emit(files[0])

    def mousePressEvent(self, event: QMouseEvent):
        if event.button() == Qt.MouseButton.LeftButton:
            self.setStyleSheet("""
                QLabel {
                    border: 2px solid #66a3ff;
                    border-radius: 5px;
                    padding: 20px;
                    background-color: #d0d0d0;
                }
            """)
            QTimer.singleShot(100, self.reset_style)  # Reset style after 100ms
            file_path, _ = QFileDialog.getOpenFileName(self, "Select Excel file", "", "Excel Files (*.xlsx *.xls)")
            if file_path:
                self.file_dropped.emit(file_path)

    def reset_style(self):
        self.setStyleSheet("""
            QLabel {
                border: 2px dashed #aaa;
                border-radius: 5px;
                padding: 20px;
                background-color: #f0f0f0;
            }
            QLabel:hover {
                background-color: #e0e0e0;
                border-color: #999;
            }
        """)


class FileDropArea(QWidget):
    files_added = pyqtSignal()
    files_cleared = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.layout = QVBoxLayout(self)
        
        self.label = QLabel("Drop Excel files here or click to browse")
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.layout.addWidget(self.label)
        
        self.file_list = QListWidget()
        self.file_list.setMinimumWidth(400)  # Adjust as needed
        self.layout.addWidget(self.file_list)
        
        self.setStyleSheet("""
            FileDropArea {
                border: 2px dashed #aaa;
                border-radius: 5px;
                padding: 20px;
                background-color: #f0f0f0;
            }
            FileDropArea:hover {
                background-color: #e0e0e0;
                border-color: #999;
            }
        """)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("""
                FileDropArea {
                    border: 2px solid #66a3ff;
                    border-radius: 5px;
                    padding: 20px;
                    background-color: #e6f2ff;
                }
            """)

    def dragLeaveEvent(self, event):
        self.reset_style()

    def dropEvent(self, event: QDropEvent):
        self.reset_style()
        files = [u.toLocalFile() for u in event.mimeData().urls() if u.toLocalFile().endswith(('.xlsx', '.xls'))]
        self.add_files(files)

    def reset_style(self):
        self.setStyleSheet("""
            FileDropArea {
                border: 2px dashed #aaa;
                border-radius: 5px;
                padding: 20px;
                background-color: #f0f0f0;
            }
            FileDropArea:hover {
                background-color: #e0e0e0;
                border-color: #999;
            }
        """)

    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Excel Files", "", "Excel Files (*.xlsx *.xls)")
        self.add_files(files)

    def add_files(self, files):
        for file in files:
            if file and os.path.isfile(file):  # Check if file exists
                item = QListWidgetItem(os.path.basename(file))
                item.setData(Qt.ItemDataRole.UserRole, file)
                self.file_list.addItem(item)
            else:
                print(f"Invalid file: {file}")  # Debug print
        
        self.update_label()
        
        if files:
            self.files_added.emit()

    def update_label(self):
        count = self.file_list.count()
        self.label.setText(f"{count} file{'s' if count != 1 else ''} selected")

    def clear_all_files(self):
        self.file_list.clear()
        self.update_label()
        self.files_cleared.emit()


class CustomProgressDialog(QProgressDialog):
    def __init__(self, title, message, parent=None):
        super().__init__(message, None, 0, 100, parent)
        self.setWindowTitle(title)
        self.setWindowModality(Qt.WindowModality.WindowModal)
        self.setMinimumDuration(0)
        self.setCancelButton(None)
        self.setMinimumWidth(300)
        self.setRange(0, 0)  # Set to indeterminate mode
        
        self.progress_value = 0
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_progress)
        self.timer.start(100)  # Update every 100 ms
        
    def update_progress(self):
        self.setValue(self.progress_value)
        
    def set_progress(self, value):
        self.progress_value = value
        
    def close(self):
        self.timer.stop()
        super().close()


class FileProcessingDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Processing File")
        self.setFixedSize(300, 100)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)

        layout = QVBoxLayout(self)

        self.label = QLabel("Processing file...")
        layout.addWidget(self.label)

        self.progress_bar = QProgressBar(self)
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        layout.addWidget(self.progress_bar)

    def set_message(self, message):
        self.label.setText(message)


def show_error_message(parent, title, message):
    QMessageBox.critical(parent, title, message)

def show_info_message(parent, title, message):
    QMessageBox.information(parent, title, message)

def show_question_dialog(parent, title, message):
    return QMessageBox.question(parent, title, message,
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                QMessageBox.StandardButton.No)

def get_open_file_names(parent, caption, directory, file_filter):
    return QFileDialog.getOpenFileNames(parent, caption, directory, file_filter)

def get_save_file_name(parent, caption, directory, file_filter):
    return QFileDialog.getSaveFileName(parent, caption, directory, file_filter)
