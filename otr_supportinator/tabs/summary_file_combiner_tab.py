from PyQt6.QtWidgets import (QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
                             QListWidget, QFileDialog, QScrollArea, QWidget,
                             QTableWidget, QTableWidgetItem, QHeaderView, QDialog, 
                             QRadioButton, QDialogButtonBox, QComboBox, QSpinBox,
                             QCheckBox, QGroupBox, QSizePolicy, QMessageBox)
from PyQt6.QtCore import Qt, pyqtSignal, QSize, QRect
from PyQt6.QtGui import QColor, QResizeEvent, QDropEvent, QDragEnterEvent, QFontMetrics, QPainter
from .base_tab import BaseTab
import os
import re

class FileListWidget(QListWidget):
    files_changed = pyqtSignal()
    file_added = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setDragDropMode(QListWidget.DragDropMode.InternalMove)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            event.setDropAction(Qt.DropAction.CopyAction)
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith(('.xlsx', '.xls')):
                    self.addItem(file_path)
                    self.file_added.emit(file_path)
            event.accept()
            self.files_changed.emit()
        else:
            super().dropEvent(event)

class WrappingLabel(QLabel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.setWordWrap(True)
        self.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Minimum)

    def paintEvent(self, event):
        painter = QPainter(self)
        metrics = QFontMetrics(self.font())
        leading = metrics.leading()
        height = 0
        line_count = 0
        available_width = self.width()

        for line in self.text().split('\n'):
            height += leading
            if not line:
                height += metrics.height()
                continue

            line_width = metrics.horizontalAdvance(line)
            if line_width <= available_width:
                line_rect = QRect(0, height, line_width, metrics.height())
                painter.drawText(line_rect, self.alignment(), line)
                height += metrics.height()
            else:
                # Wrap the line character by character
                current_line = ""
                for char in line:
                    new_line = current_line + char
                    new_width = metrics.horizontalAdvance(new_line)
                    if new_width <= available_width:
                        current_line = new_line
                    else:
                        line_rect = QRect(0, height, metrics.horizontalAdvance(current_line), metrics.height())
                        painter.drawText(line_rect, self.alignment(), current_line)
                        height += metrics.height()
                        current_line = char
                    line_count += 1
                if current_line:
                    line_rect = QRect(0, height, metrics.horizontalAdvance(current_line), metrics.height())
                    painter.drawText(line_rect, self.alignment(), current_line)
                    height += metrics.height()

        self.setMinimumHeight(height)

class SummaryFileCombinerTab(BaseTab):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.file_data = {}
        self.planning_week = None
        self.combinations = []
        self.init_ui()

    def init_ui(self):
        main_layout = self.layout

        # File drop area
        self.file_list = FileListWidget()
        self.file_list.setMinimumHeight(50)
        self.file_list.files_changed.connect(self.update_ui_state)
        self.file_list.file_added.connect(self.process_file)
        self.file_list.itemSelectionChanged.connect(self.update_ui_state)
        main_layout.addWidget(self.file_list)

        # File buttons layout
        file_buttons_layout = QHBoxLayout()

        self.browse_button = QPushButton("Browse Files")
        self.browse_button.clicked.connect(self.browse_files)
        file_buttons_layout.addWidget(self.browse_button)

        self.remove_button = QPushButton("Remove Selected")
        self.remove_button.clicked.connect(self.remove_selected_files)
        file_buttons_layout.addWidget(self.remove_button)

        main_layout.addLayout(file_buttons_layout)

        # Planning week layout
        self.planning_week_widget = QWidget()
        planning_week_layout = QHBoxLayout(self.planning_week_widget)
        planning_week_layout.setContentsMargins(0, 0, 0, 0)
        self.planning_week_label = QLabel("Planning Week: Not Set")
        planning_week_layout.addWidget(self.planning_week_label)
        main_layout.addWidget(self.planning_week_widget)

        # Planned Weeks Available table
        self.planned_weeks_table = QTableWidget(10, 3)
        self.planned_weeks_table.setHorizontalHeaderLabels(["Amazon Week", "Available Planned Weeks", "Source Control"])
        self.setup_table()
        main_layout.addWidget(self.planned_weeks_table)

        # Combinations area
        self.combinations_widget = QWidget()
        combinations_layout = QHBoxLayout(self.combinations_widget)
        combinations_layout.setSpacing(0)
        combinations_layout.setContentsMargins(0, 0, 0, 0)
        default_presets = ["All", "A", "B", "C", "D"]
        for i in range(5):
            combination = self.create_combination_widget(i, default_presets[i])
            combinations_layout.addWidget(combination)
            self.combinations.append(combination)
        main_layout.addWidget(self.combinations_widget)

        # Add some vertical spacing
        main_layout.addSpacing(20)

        # Generate Combined Files button
        self.generate_button = QPushButton("Generate Combined Files")
        self.generate_button.clicked.connect(self.generate_combined_files)
        main_layout.addWidget(self.generate_button)

        self.update_ui_state()

    def setup_table(self):
        self.planned_weeks_table.setRowCount(10)
        self.planned_weeks_table.setColumnCount(4)
        self.planned_weeks_table.setHorizontalHeaderLabels(["Planning Horizon", "Amazon Week", "Available Planned Weeks", "Source Control"])

        planning_horizons = ["W-1", "W-2", "W-3", "W-4", "W-5", "W-6", "W-7", "W-8", "W-9", "W-10"]
        for i, horizon in enumerate(planning_horizons):
            self.planned_weeks_table.setItem(i, 0, QTableWidgetItem(horizon))
            self.planned_weeks_table.setItem(i, 1, QTableWidgetItem(""))
            self.planned_weeks_table.setItem(i, 2, QTableWidgetItem(""))
            button = QPushButton("Select Source")
            button.setEnabled(False)
            self.planned_weeks_table.setCellWidget(i, 3, button)

        self.planned_weeks_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.planned_weeks_table.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)

    def create_combination_widget(self, index, default_preset):
        group = QGroupBox(f"Combination {index + 1}")
        layout = QVBoxLayout()

        # Dropdown for preset selections
        preset_dropdown = QComboBox()
        preset_dropdown.addItems(["A", "B", "C", "D", "All", "Custom"])
        preset_dropdown.setCurrentText(default_preset)
        layout.addWidget(preset_dropdown)

        # Week range selection
        range_layout = QHBoxLayout()
        start_week = QSpinBox()
        start_week.setRange(1, 10)
        end_week = QSpinBox()
        end_week.setRange(1, 10)
        range_layout.addWidget(QLabel("From:"))
        range_layout.addWidget(start_week)
        range_layout.addWidget(QLabel("To:"))
        range_layout.addWidget(end_week)
        layout.addLayout(range_layout)

        # Title label
        title_label = WrappingLabel()
        title_label.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignLeft)
        layout.addWidget(title_label)

        # Generate checkbox
        generate_checkbox = QCheckBox("Generate")
        generate_checkbox.setChecked(True)  # Set to True by default
        layout.addWidget(generate_checkbox)

        group.setLayout(layout)

        # Connect signals
        preset_dropdown.currentIndexChanged.connect(lambda: self.update_combination_range(group, index))
        start_week.valueChanged.connect(lambda: self.update_combination_title(group, index))
        end_week.valueChanged.connect(lambda: self.update_combination_title(group, index))
        generate_checkbox.stateChanged.connect(lambda: self.update_combination_title(group, index))

        self.update_combination_range(group, index)
        return group
    
    def resizeEvent(self, event: QResizeEvent):
        super().resizeEvent(event)
        self.update_combination_widths()

    def update_combination_widths(self):
        total_width = self.combinations_widget.width()
        combination_width = total_width // 5
        for combination in self.combinations:
            combination.setFixedWidth(combination_width)
    
    def update_combination_range(self, group, index):
        preset_dropdown = group.layout().itemAt(0).widget()
        start_week = group.layout().itemAt(1).itemAt(1).widget()
        end_week = group.layout().itemAt(1).itemAt(3).widget()
        
        preset = preset_dropdown.currentText()
        ranges = {
            "A": (2, 2), "B": (3, 5), "C": (6, 7), "D": (8, 10),
            "All": (2, 10), "Custom": (1, 10)
        }

        start, end = ranges[preset]
        start_week.setValue(start)
        end_week.setValue(end)

        self.check_range_validity(group, index)
        self.update_combination_title(group, index)

    def update_combinations_validity(self):
        for index, combination in enumerate(self.combinations):
            self.check_range_validity(combination, index)

    def check_range_validity(self, group, index):
        start_week = group.layout().itemAt(1).itemAt(1).widget()
        end_week = group.layout().itemAt(1).itemAt(3).widget()
        generate_checkbox = group.layout().itemAt(3).widget()

        if start_week.value() > end_week.value():
            generate_checkbox.setEnabled(False)
            return

        # Check if all weeks in the range are available
        all_weeks_available = all(
            self.planned_weeks_table.item(w-1, 2).text().startswith("w-")
            for w in range(start_week.value(), end_week.value() + 1)
        )

        generate_checkbox.setEnabled(all_weeks_available)
        if all_weeks_available and not generate_checkbox.isEnabled():
            generate_checkbox.setChecked(True)  # Re-enable and check if it becomes valid

    def update_all_combination_titles(self):
        for index, combination in enumerate(self.combinations):
            self.update_combination_title(combination, index)

    def update_combination_title(self, group, index):
        start_week = group.layout().itemAt(1).itemAt(1).widget().value()
        end_week = group.layout().itemAt(1).itemAt(3).widget().value()
        title_label = group.layout().itemAt(2).widget()
        generate_checkbox = group.layout().itemAt(3).widget()

        self.check_range_validity(group, index)

        weeks = ".".join(str(w) for w in range(start_week, end_week + 1))
        if self.planning_week is not None:
            title = f"summary_file_plwk{self.planning_week}_w-{weeks}"
        else:
            title = f"summary_file_plwk[Not Set]_w-{weeks}"
        
        if not generate_checkbox.isChecked():
            title += " (Disabled)"
        title_label.setText(title)

    def update_ui_state(self):
        has_files = self.file_list.count() > 0
        has_error = "Multiple Planning Weeks Detected" in self.planning_week_label.text()
        
        self.remove_button.setEnabled(has_files and len(self.file_list.selectedItems()) > 0)
        self.browse_button.setEnabled(not has_error)
        self.planned_weeks_table.setEnabled(not has_error)
        self.generate_button.setEnabled(has_files and not has_error)
        
        for combination in self.combinations:
            combination.setEnabled(not has_error)
        
        if has_files and self.planning_week is not None and not has_error:
            self.update_table()

    def remove_selected_files(self):
        for item in self.file_list.selectedItems():
            file_path = item.text()
            row = self.file_list.row(item)
            self.file_list.takeItem(row)
            if file_path in self.file_data:
                del self.file_data[file_path]
        
        # Reset planning week and hide warning if no files left
        if self.file_list.count() == 0:
            self.planning_week = None
            self.planning_week_label.setText("Planning Week: Not Set")
            self.planning_week_widget.setStyleSheet("")
        else:
            # Check if the error state needs to be cleared
            first_file = self.file_list.item(0).text()
            first_planning_week = self.extract_planning_week(os.path.basename(first_file))
            if all(self.extract_planning_week(os.path.basename(self.file_list.item(i).text())) == first_planning_week 
                for i in range(self.file_list.count())):
                self.planning_week = first_planning_week
                self.planning_week_label.setText(f"Planning Week: {self.planning_week}")
                self.planning_week_widget.setStyleSheet("")

        self.update_planned_weeks()
        self.update_all_combination_titles()  # Add this line
        self.update_ui_state()

    def update_table(self):
        if self.planning_week is None:
            return
        
        for i in range(10):
            amazon_week = (self.planning_week + i + 1) % 52
            if amazon_week == 0:
                amazon_week = 52
            self.planned_weeks_table.item(i, 1).setText(str(amazon_week))
        
        self.update_planned_weeks()

    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Excel Files", "", "Excel Files (*.xlsx *.xls)")
        for file in files:
            self.file_list.addItem(file)
            self.process_file(file)

    def process_file(self, file_path):
        file_name = os.path.basename(file_path)
        extracted_planning_week = self.extract_planning_week(file_name)
        if extracted_planning_week:
            if self.planning_week is None:
                self.planning_week = extracted_planning_week
                self.planning_week_label.setText(f"Planning Week: {self.planning_week}")
                self.update_all_combination_titles()  # Add this line
            elif self.planning_week != extracted_planning_week:
                self.planning_week_label.setText("Planning Week: Multiple Planning Weeks Detected - Remove Summary File From Incorrect Planning Week")
                self.planning_week_widget.setStyleSheet("background-color: red; color: white; padding: 5px;")
        
        planned_horizons = self.extract_planned_horizons(file_name)
        self.file_data[file_path] = planned_horizons
        self.update_planned_weeks()
        self.update_combinations_validity()
        self.update_ui_state()

    def extract_planning_week(self, file_name):
        match = re.search(r"summary_file_plwk(\d+)", file_name)
        return int(match.group(1)) if match else None

    def extract_planned_horizons(self, file_name):
        match = re.search(r"w-([\d\.]+)", file_name)
        return [int(horizon) for horizon in match.group(1).split('.')] if match else []

    def update_planned_weeks(self):
        for i in range(10):
            self.planned_weeks_table.item(i, 2).setText("")
            self.planned_weeks_table.item(i, 2).setBackground(QColor("white"))
            self.planned_weeks_table.cellWidget(i, 3).setEnabled(False)

        for file_path, horizons in self.file_data.items():
            for horizon in horizons:
                if 1 <= horizon <= 10:
                    row = horizon - 1
                    current_text = self.planned_weeks_table.item(row, 2).text()
                    if current_text:
                        self.planned_weeks_table.item(row, 2).setText("Duplicate")
                        self.planned_weeks_table.item(row, 2).setBackground(QColor("red"))
                        button = self.planned_weeks_table.cellWidget(row, 3)
                        button.setEnabled(True)
                        button.clicked.connect(lambda checked, r=row: self.select_source(r))
                    else:
                        self.planned_weeks_table.item(row, 2).setText(f"w-{horizon}")
                        self.planned_weeks_table.item(row, 2).setBackground(QColor("green"))

    def select_source(self, row):
        dialog = QDialog(self)
        dialog.setWindowTitle("Select Source File")
        layout = QVBoxLayout()

        horizon = int(self.planned_weeks_table.item(row, 0).text().split('-')[1])
        files_with_horizon = [file for file, horizons in self.file_data.items() if horizon in horizons]

        for file in files_with_horizon:
            radio = QRadioButton(os.path.basename(file))
            layout.addWidget(radio)

        button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        dialog.setLayout(layout)

        if dialog.exec() == QDialog.DialogCode.Accepted:
            for i in range(layout.count() - 1):  # -1 to exclude button box
                radio = layout.itemAt(i).widget()
                if radio.isChecked():
                    selected_file = [file for file in files_with_horizon if os.path.basename(file) == radio.text()][0]
                    horizon = int(self.planned_weeks_table.item(row, 0).text().split('-')[1])
                    self.planned_weeks_table.item(row, 2).setText(f"w-{horizon}")
                    self.planned_weeks_table.item(row, 2).setBackground(QColor(173, 216, 230))  # Light blue
                    self.planned_weeks_table.cellWidget(row, 3).setEnabled(False)
                    # Disconnect the button to prevent multiple connections
                    self.planned_weeks_table.cellWidget(row, 3).clicked.disconnect()
                    break

    def generate_combined_files(self):
        # Collect enabled combinations
        enabled_combinations = []
        for combination in self.combinations:
            generate_checkbox = combination.layout().itemAt(3).widget()
            if generate_checkbox.isChecked():
                start_week = combination.layout().itemAt(1).itemAt(1).widget().value()
                end_week = combination.layout().itemAt(1).itemAt(3).widget().value()
                title_label = combination.layout().itemAt(2).widget()
                enabled_combinations.append({
                    'start_week': start_week,
                    'end_week': end_week,
                    'title': title_label.text()
                })

        if not enabled_combinations:
            QMessageBox.warning(self, "No Combinations", "Please enable at least one combination before generating files.")
            return

        # TODO: Implement the file combination logic here
        # For now, we'll just show a message with the enabled combinations
        combination_details = "\n".join([f"Combination: {c['title']} (Weeks {c['start_week']}-{c['end_week']})" for c in enabled_combinations])
        QMessageBox.information(self, "Combinations to Generate", f"The following combinations will be generated:\n\n{combination_details}")


    def restart(self):
        self.file_list.clear()
        self.file_data.clear()
        self.planning_week = None
        self.planning_week_label.setText("Planning Week: Not Set")
        self.planning_week_widget.setStyleSheet("")
        self.setup_table()
        for combination in self.combinations:
            self.update_combination_range(combination, self.combinations.index(combination))
        self.update_ui_state()  # Call this to update the UI state after restarting

    def process(self):
        # Implement the process logic here
        pass