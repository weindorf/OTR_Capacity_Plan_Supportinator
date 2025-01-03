import io
import sys
import os
import re
import pandas as pd
from datetime import datetime
from PyQt6.QtWidgets import (QVBoxLayout, QHBoxLayout, QLabel, QPushButton, 
                             QSpinBox, QLineEdit, QTextEdit, QProgressDialog,
                             QMessageBox, QFileDialog, QComboBox, QGroupBox, 
                             QFormLayout, QMainWindow)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from .base_tab import BaseTab
from ..utils.gui_components import FileDropArea
from ..utils.file_utils import process_file
from ..utils.date_utils import get_amazon_week

class SummaryFileGeneratorWorker(QThread):
    progress_update = pyqtSignal(int, str)
    error_occurred = pyqtSignal(str)
    finished = pyqtSignal(object, object, object, str, list)
    request_save_file = pyqtSignal(str, str)
    operation_cancelled = pyqtSignal()
    file_saved = pyqtSignal()

    def __init__(self, files, planning_type, temp_dir, suggested_filename, parent=None):
        super().__init__(parent)
        self.files = files
        self.planning_type = planning_type
        self.temp_dir = temp_dir
        self.suggested_filename = suggested_filename
        self.parent = parent
        self.is_cancelled = False
        self.save_file_path = None

    def run(self):
        try:
            self.progress_update.emit(0, "Starting file processing...")
            pivot_table, weekly_summary, region_weekly_summary = self.process_files()
            
            if self.is_cancelled:
                self.operation_cancelled.emit()
                return

            self.progress_update.emit(95, "Preparing to save file...")
            self.request_save_file.emit(self.suggested_filename, os.path.dirname(self.files[0]))
            
            while self.save_file_path is None:
                QThread.msleep(100)
                if self.is_cancelled:
                    self.operation_cancelled.emit()
                    return

            if self.save_file_path:
                self.progress_update.emit(97, "Saving summary file...")
                pivot_table.to_excel(self.save_file_path, index=False)
                self.progress_update.emit(100, "File saved successfully.")
                self.file_saved.emit()
                self.finished.emit(pivot_table, weekly_summary, region_weekly_summary, self.save_file_path, self.warnings)
            else:
                self.error_occurred.emit("File save cancelled.")
        except Exception as e:
            self.error_occurred.emit(str(e))

    def cancel(self):
        self.is_cancelled = True

    def process_files(self):
        results = []
        total_files = len(self.files)

        for i, file_path in enumerate(self.files, 1):
            self.progress_update.emit(int(90 * i / total_files), f"Processing file {i} of {total_files}")
            
            if file_path is None:
                continue
            
            try:
                result = process_file(file_path)
                if result is not None:
                    results.append(result)
            except Exception as e:
                self.error_occurred.emit(f"Error processing file {file_path}: {str(e)}")
            
            if self.is_cancelled:
                raise Exception("Operation cancelled by user")

        if not results:
            raise ValueError("No valid data found in any of the input files.")

        self.progress_update.emit(90, "Combining results...")
        combined_df = pd.concat(results, ignore_index=True)

        self.progress_update.emit(92, "Creating pivot table...")

        try:
            # Pivot the combined dataframe
            pivot_table = pd.pivot_table(combined_df,
                                        values='value',
                                        index=['region', 'node', 'cycle', 'forecast_period_start'],
                                        columns=['metric', 'sub_metric'],
                                        aggfunc='first',
                                        fill_value=None)

            # Reset the index to make 'region', 'node', etc. regular columns
            pivot_table = pivot_table.reset_index()

            # Flatten the multi-level column names
            pivot_table.columns = [' '.join(col).strip() if isinstance(col, tuple) else col for col in pivot_table.columns]

            # Convert 'forecast_period_start' to datetime
            pivot_table['forecast_period_start'] = pd.to_datetime(pivot_table['forecast_period_start'])

            # Calculate and insert 'amazon_week'
            pivot_table['forecast_period_start'] = pd.to_datetime(pivot_table['forecast_period_start'])
            pivot_table.insert(1, 'amazon_week', pivot_table['forecast_period_start'].apply(get_amazon_week))

            # Calculate CVP
            if '1 - FO volume' in pivot_table.columns and '2 - otr_capa calculated_total' in pivot_table.columns:
                pivot_table.insert(5, 'CVP', pivot_table[['1 - FO volume', '2 - otr_capa calculated_total']].min(axis=1))
            else:
                pivot_table.insert(5, 'CVP', None)
                self.log_message.emit("Warning: Unable to calculate CVP due to missing columns.")

            # Add generated_at column
            pivot_table.insert(6, 'generated_at', datetime.now().strftime('%Y-%m-%d %H:%M:%S'))

            # Convert forecast_period_start back to string
            pivot_table['forecast_period_start'] = pivot_table['forecast_period_start'].dt.strftime('%Y-%m-%d')

            # List of expected columns (based on the original output)
            expected_columns = [
                'region', 'amazon_week', 'node', 'cycle', 'forecast_period_start', '1 - FO volume',
                '2 - otr_capa calculated_total', 'CVP', 'generated_at', '2 - otr_capa optimizer_total',
                '3 - hdp capacity', '4 - amflex alloted_capacity', '4 - amflex bau_avg_capa',
                '4 - amflex capacity_ask', '4 - amflex commitment_capacity', '4 - amflex max_block',
                '4 - amflex mde_max_capa', '4 - amflex spr', '4 - amflex vans_alloted',
                '4 - amflex vans_ask', '4 - amflex vans_committed', '4 - amflex_keicar capacity',
                '4 - amflex_keicar spr', '4 - amflex_keicar vans', '4.1 - amflex_total capacity',
                '4.1 - amflex_total spr', '4.1 - amflex_total vans', '5 - dsp2.0_keivan capacity',
                '5 - dsp2.0_keivan spr', '5 - dsp2.0_keivan vans', '5 - dsp2.0_largevan capacity',
                '5 - dsp2.0_largevan spr', '5 - dsp2.0_largevan vans', '5 - dsp_1t_walker capacity',
                '5 - dsp_1t_walker spr', '5 - dsp_1t_walker vans', '5 - dsp_biker capacity',
                '5 - dsp_biker spr', '5 - dsp_biker vans', '5 - dsp_keivan capacity',
                '5 - dsp_keivan spr', '5 - dsp_keivan vans', '5 - dsp_keivan vans_rescue',
                '5 - dsp_keivan_walker capacity', '5 - dsp_keivan_walker spr',
                '5 - dsp_keivan_walker vans', '5 - dsp_largevan capacity', '5 - dsp_largevan spr',
                '5 - dsp_largevan vans', '5 - dsp_walker capacity', '5 - dsp_walker spr',
                '5 - dsp_walker vans', '5.1 - dsp_total capacity', '5.1 - dsp_total spr',
                '5.1 - dsp_total vans', '6 - excess/shortage capacity'
                ]


            # Check for missing columns
            missing_columns = [col for col in expected_columns if col not in pivot_table.columns]
            if missing_columns:
                self.warnings.append(f"Warning: The following expected columns are missing: {missing_columns}")
                self.warnings.append("This may indicate issues with the input data or data processing.")

            # Add missing columns with None values
            for col in missing_columns:
                pivot_table[col] = None

            # Reorder columns
            pivot_table = pivot_table[expected_columns]

            # Calculate total van ask
            if '4 - amflex vans_ask' in pivot_table.columns:
                total_van_ask = pivot_table['4 - amflex vans_ask'].sum()
                self.progress_update.emit(94, f"Total Van ask (all weeks): {round(int(total_van_ask),0)}")
            else:
                self.warnings.append("Warning: Unable to calculate total van ask. '4 - amflex vans_ask' column is missing.")

            # Create summaries
            if '4 - amflex vans_ask' in pivot_table.columns:
                weekly_summary = pivot_table.groupby('amazon_week')['4 - amflex vans_ask'].sum().reset_index()
                weekly_summary['4 - amflex vans_ask'] = weekly_summary['4 - amflex vans_ask'].astype(int)

                region_weekly_summary = pivot_table.groupby(['amazon_week', 'region'])['4 - amflex vans_ask'].sum().reset_index()
                region_weekly_summary['4 - amflex vans_ask'] = region_weekly_summary['4 - amflex vans_ask'].astype(int)
            else:
                weekly_summary = pd.DataFrame(columns=['amazon_week', '4 - amflex vans_ask'])
                region_weekly_summary = pd.DataFrame(columns=['amazon_week', 'region', '4 - amflex vans_ask'])
                self.warnings.append("Warning: Unable to create summaries. '4 - amflex vans_ask' column is missing.")

            self.progress_update.emit(95, "Pivot table and summaries created")

            return pivot_table, weekly_summary, region_weekly_summary

        except Exception as e:
            self.error_occurred.emit(f"Error creating pivot table: {str(e)}")
            raise

    def progress_callback(self, value, message):
        if self.is_cancelled:
            raise Exception("Operation cancelled by user")
        self.progress_update.emit(value, message)

class SummaryFileGeneratorTab(BaseTab):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.main_window = self.get_main_window()
        self.init_ui()

    def init_ui(self):
        self.file_drop_area = FileDropArea()
        self.layout.addWidget(self.file_drop_area)

        # File buttons layout
        file_buttons_layout = QHBoxLayout()
        self.browse_button = QPushButton("Browse Files")
        self.browse_button.clicked.connect(self.browse_files)
        self.remove_selected_button = QPushButton("Remove Selected")
        self.remove_selected_button.clicked.connect(self.remove_selected_files)
        self.remove_selected_button.setEnabled(False)
        self.clear_all_button = QPushButton("Clear All")
        self.clear_all_button.clicked.connect(self.file_drop_area.clear_all_files)
        self.clear_all_button.setEnabled(False)
        file_buttons_layout.addWidget(self.browse_button)
        file_buttons_layout.addWidget(self.remove_selected_button)
        file_buttons_layout.addWidget(self.clear_all_button)
        self.layout.addLayout(file_buttons_layout)

        self.settings_group = QGroupBox("Settings")
        settings_layout = QFormLayout()
        self.settings_group.setLayout(settings_layout)


        self.plan_type_combo = QComboBox()
        self.plan_type_combo.addItems(["ask", "DS Callout", "final", "custom"])
        self.plan_type_combo.setEditable(True)
        self.plan_type_combo.setEnabled(False)
        self.plan_type_combo.setCurrentText("")  # Set to empty string initially
        settings_layout.addRow("Plan Type:", self.plan_type_combo)

        self.planning_week_spin = QSpinBox()
        self.planning_week_spin.setRange(1, 52)
        self.planning_week_spin.setValue(get_amazon_week(datetime.now()))
        self.planning_week_spin.setEnabled(False)
        settings_layout.addRow("Planning Week:", self.planning_week_spin)

        self.file_name_preview = QLineEdit()
        self.file_name_preview.setReadOnly(False)
        settings_layout.addRow("Output File Name:", self.file_name_preview)

        self.layout.addWidget(self.settings_group)

        self.generate_button = QPushButton("Generate Summary File")
        self.generate_button.clicked.connect(self.process)
        self.layout.addWidget(self.generate_button)

        self.output_text = QTextEdit()
        self.output_text.setReadOnly(True)
        self.layout.addWidget(self.output_text)

        self.file_drop_area.files_added.connect(self.update_ui_state)
        self.file_drop_area.files_cleared.connect(self.update_ui_state)
        self.file_drop_area.file_list.itemSelectionChanged.connect(self.update_remove_button_state)
        self.plan_type_combo.currentTextChanged.connect(self.update_filename_preview)
        self.planning_week_spin.valueChanged.connect(self.update_filename_preview)
        self.file_name_preview.textEdited.connect(self.on_filename_edited)

        self.update_ui_state()

    def browse_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Files", "", "Excel Files (*.xlsx *.xls)")
        for file in files:
            self.file_drop_area.add_file(file)

    def remove_selected_files(self):
        for item in self.file_drop_area.file_list.selectedItems():
            self.file_drop_area.file_list.takeItem(self.file_drop_area.file_list.row(item))
        self.update_ui_state()

    def update_remove_button_state(self):
        self.remove_selected_button.setEnabled(len(self.file_drop_area.file_list.selectedItems()) > 0)

    def update_ui_state(self):
        has_files = self.file_drop_area.file_list.count() > 0
        self.clear_all_button.setEnabled(has_files)
        self.plan_type_combo.setEnabled(has_files)
        self.planning_week_spin.setEnabled(has_files)
        
        if not has_files:
            self.plan_type_combo.setCurrentText("")
            self.planning_week_spin.setValue(0)
        else:
            self.update_plan_type_and_week()
        
        self.update_filename_preview()

    def update_plan_type_and_week(self):
        if self.file_drop_area.file_list.count() > 0:
            first_file_name = self.file_drop_area.file_list.item(0).text()
            
            # Extract plan type
            plan_type_match = re.search(r'_(ask|DS Callout|final|custom)\.', first_file_name)
            if plan_type_match:
                self.plan_type_combo.setCurrentText(plan_type_match.group(1))
            else:
                self.plan_type_combo.setCurrentText("ask")
            
            # Extract planning week
            week_match = re.search(r'plwk(\d+)', first_file_name)
            if week_match:
                self.planning_week_spin.setValue(int(week_match.group(1)))
            else:
                self.planning_week_spin.setValue(get_amazon_week(datetime.now()))

    def update_filename_preview(self):
        plan_type = self.plan_type_combo.currentText()
        planning_week = self.planning_week_spin.value()
        
        if self.file_drop_area.file_list.count() > 0:
            y_string = ''
            first_file_name = self.file_drop_area.file_list.item(0).text()
            match = re.search(r'W-(\d+(?:\.\d+)*)(?:\D|$)', first_file_name)
            if match:
                y_string = match.group(1)
            
            filename = f"summary_file_plwk{planning_week}_w-{y_string}_{plan_type}.xlsx"
            self.file_name_preview.setText(filename)
            self.file_name_preview.setEnabled(True)
        else:
            self.file_name_preview.clear()
            self.file_name_preview.setEnabled(False)

    def get_main_window(self):
        parent = self.parent()
        while parent is not None:
            if isinstance(parent, QMainWindow):
                return parent
            parent = parent.parent()
        return None

    def process(self):
        files = [self.file_drop_area.file_list.item(i).data(Qt.ItemDataRole.UserRole) 
                 for i in range(self.file_drop_area.file_list.count())]
        planning_type = self.plan_type_combo.currentText()
        suggested_filename = self.file_name_preview.text()

        if not files:
            QMessageBox.warning(self, "No Files", "Please add files before generating summary.")
            return
        
        if any(file is None for file in files):
            QMessageBox.warning(self, "Invalid Files", "Some files are invalid. Please check your selection.")
            return

        if not planning_type:
            QMessageBox.warning(self, "Missing Input", "Please enter the Planning Type.")
            return

        self.progress_dialog = QProgressDialog("Generating Summary File", "Cancel", 0, 100, self)
        self.progress_dialog.setWindowTitle("Generating Summary File")
        self.progress_dialog.setWindowModality(Qt.WindowModality.WindowModal)
        self.progress_dialog.setAutoReset(False)
        self.progress_dialog.setAutoClose(False)
        self.progress_dialog.setValue(0)
        self.progress_dialog.canceled.connect(self.cancel_summary_generation)
        self.progress_dialog.show()

        self.worker = SummaryFileGeneratorWorker(files, planning_type, self.main_window.temp_dir, suggested_filename, self)
        self.worker.progress_update.connect(self.update_progress)
        self.worker.error_occurred.connect(self.handle_error)
        self.worker.finished.connect(self.handle_finished)
        self.worker.request_save_file.connect(self.get_save_file_name)
        self.worker.operation_cancelled.connect(self.handle_cancellation)
        self.worker.file_saved.connect(self.progress_dialog.close)
        
        self.worker.start()

    def on_filename_edited(self):
        self.custom_output_name = True

    def cancel_summary_generation(self):
        if self.worker:
            self.worker.cancel()
        self.progress_dialog.close()

    def update_progress(self, value, message):
        self.progress_dialog.setValue(value)
        self.progress_dialog.setLabelText(message)

    def handle_error(self, error_message):
        self.progress_dialog.close()
        QMessageBox.critical(self, "Error", f"An error occurred: {error_message}")

    def handle_finished(self, pivot_table, weekly_summary, region_weekly_summary, output_file, warnings):
        self.progress_dialog.close()
        self.display_results(pivot_table, weekly_summary, region_weekly_summary, output_file, warnings)

    def display_results(self, pivot_table, weekly_summary, region_weekly_summary, output_file, warnings):
        total_van_ask = pivot_table['4 - amflex vans_ask'].sum() if '4 - amflex vans_ask' in pivot_table.columns else 0
        
        self.output_text.clear()
        self.output_text.append(f"<h2>Summary Report</h2>")
        self.output_text.append(f"<p><b>Total Van ask (all weeks):</b> {round(int(total_van_ask),0)}</p>")
        self.output_text.append(f"<p><b>Output file:</b> {output_file}</p>")
        
        self.output_text.append("<h3>Weekly Breakdown</h3>")
        self.add_table_to_text_edit(weekly_summary)
        
        self.output_text.append("<h3>Region-wise Breakdown</h3>")
        self.add_table_to_text_edit(region_weekly_summary)
        
        if warnings:
            self.output_text.append("<h3 style='color: red;'>Warnings</h3>")
            for warning in warnings:
                self.output_text.append(f"<p style='color: red;'>{warning}</p>")
        
        if warnings:
            QMessageBox.warning(self, "Successful with Errors", f"Summary file saved successfully as:\n{output_file}\n\nThere were warnings during processing. Please check the summary report for details.")
        else:
            QMessageBox.information(self, "Success", f"Summary file saved successfully as:\n{output_file}")

    def add_table_to_text_edit(self, df):
        html = "<table border='1' cellpadding='3' cellspacing='0'>"
        html += "<tr>" + "".join(f"<th>{col}</th>" for col in df.columns) + "</tr>"
        for _, row in df.iterrows():
            html += "<tr>" + "".join(f"<td>{value}</td>" for value in row) + "</tr>"
        html += "</table>"
        self.output_text.append(html)

    def get_save_file_name(self, suggested_filename, default_dir):
        file_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Summary File",
            os.path.join(default_dir, suggested_filename),
            "Excel Files (*.xlsx)"
        )
        if self.worker:
            self.worker.save_file_path = file_path if file_path else None

    def update_output_display(self, message):
        self.output_text.append(message)

    def handle_cancellation(self):
        QMessageBox.information(self, "Cancelled", "Operation was cancelled by the user.")

    def restart(self):
        self.file_drop_area.clear_all_files()
        self.plan_type_combo.setCurrentText("")
        self.planning_week_spin.setValue(0)
        self.plan_type_combo.setEnabled(False)
        self.planning_week_spin.setEnabled(False)
        self.file_name_preview.clear()
        self.output_text.clear()
        self.custom_output_name = False
        self.update_ui_state()
