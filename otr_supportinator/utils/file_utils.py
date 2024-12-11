import os
import re
import pandas as pd
import numpy as np
import openpyxl
from datetime import datetime
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from openpyxl import load_workbook

def process_file(file_path):
    if not file_path or not os.path.isfile(file_path):
        raise ValueError(f"Invalid file path: {file_path}")
    
    print(f"Processing file: {file_path}")  # Debug print
    try:
        print(f"\nProcessing file: {file_path}")
        
        # Read the Excel file using openpyxl
        wb = openpyxl.load_workbook(file_path, data_only=True)
        sheet = wb.active
        
        # Convert openpyxl worksheet to a list of lists, preserving original values
        data = []
        for row in sheet.iter_rows():
            row_data = []
            for cell in row:
                if isinstance(cell.value, datetime):
                    row_data.append(cell.value.strftime('%Y-%m-%d'))
                elif cell.data_type == 'e':  # Error cell
                    row_data.append(None)
                elif cell.data_type == 'f':  # Formula cell
                    row_data.append(cell.value)
                else:
                    row_data.append(cell.value)
            data.append(row_data)
        
        # Create DataFrame from the data
        df = pd.DataFrame(data[1:], columns=data[0])
        
        print(f"Original DataFrame shape: {df.shape}")
        print(f"Original DataFrame columns: {df.columns.tolist()}")
        
        # Identify index columns and date columns
        index_columns = ['region', 'channel_type', 'parent_node', 'prefecture', 'carrier', 'node', 'cycle', 'metric', 'sub_metric']
        date_columns = [col for col in df.columns if col not in index_columns]
        print(f"Identified date columns: {date_columns}")

        # Melt the dataframe to long format
        df_melted = df.melt(id_vars=index_columns, 
                            var_name='forecast_period_start', 
                            value_name='value')
        
        # Ensure forecast_period_start is datetime
        df_melted['forecast_period_start'] = pd.to_datetime(df_melted['forecast_period_start'], format='%Y-%m-%d', errors='coerce')
        
        # Drop rows with invalid dates
        df_melted = df_melted.dropna(subset=['forecast_period_start'])
        
        print(f"Melted DataFrame shape: {df_melted.shape}")
        print(f"Melted DataFrame columns: {df_melted.columns.tolist()}")
        print(f"Sample of final DataFrame:\n{df_melted.head().to_string()}\n")

        # Print unique combinations of metric and sub_metric
        print("Unique metric and sub_metric combinations:")
        print(df_melted.groupby(['metric', 'sub_metric']).size().reset_index().to_string())

        return df_melted
    except Exception as e:
        print(f"Error processing file {file_path}: {str(e)}")
        return None

def validate_file(file_path):
    file_name = os.path.basename(file_path)
    last_modified = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime("%Y-%m-%d %H:%M:%S")
    
    is_valid_format = file_path.lower().endswith('.xlsx')
    is_valid_name = file_name.startswith('summary_file_w')
    
    return {
        'is_valid_format': is_valid_format,
        'is_valid_name': is_valid_name,
        'file_name': file_name,
        'last_modified': last_modified
    }

def clean_data(value):
    if isinstance(value, (int, float)):
        if np.isnan(value) or np.isinf(value):
            return ''  # Convert NaN and Inf to empty string
    return value

def merge_excel_files(file_paths, temp_dir, progress_callback):
    total_files = len(file_paths)
    merged_output_file = os.path.join(temp_dir, 'merged_data.xlsx')
    
    MAX_ROWS = 1000000  # Slightly less than 1,048,576 to leave room for header

    with pd.ExcelWriter(merged_output_file, engine='openpyxl') as writer:
        header = ["region", "channel_type", "parent_node", "prefecture", "carrier", "node", "cycle", "metric", "sub_metric", "forecast_period_start", "value"]
        
        sheet_index = 1
        row_offset = 0
        current_sheet = None

        # Create the first sheet
        sheet_name = f'Sheet{sheet_index}'
        writer.book.create_sheet(sheet_name)
        current_sheet = writer.book[sheet_name]
        
        # Write header for the first sheet
        for col_num, column_title in enumerate(header, 1):
            current_sheet.cell(row=1, column=col_num, value=column_title)
        row_offset = 1

        for i, file_path in enumerate(file_paths, 1):
            progress_callback(int(30 * i / total_files), f"Processing file {i} of {total_files}")
            
            try:
                df = process_file(file_path, progress_callback)
                
                for _, row in df.iterrows():
                    if row_offset % MAX_ROWS == 0:
                        # Create a new sheet
                        sheet_index += 1
                        sheet_name = f'Sheet{sheet_index}'
                        writer.book.create_sheet(sheet_name)
                        current_sheet = writer.book[sheet_name]
                        row_offset = 0

                        # Write header for the new sheet
                        for col_num, column_title in enumerate(header, 1):
                            current_sheet.cell(row=1, column=col_num, value=column_title)
                        row_offset = 1

                    row_offset += 1
                    for col_num, value in enumerate(row, 1):
                        current_sheet.cell(row=row_offset, column=col_num, value=value)
            except Exception as e:
                progress_callback(0, f"Error processing file {file_path}: {str(e)}")
                raise

            progress_callback(int(60 * i / total_files), f"Merged {i} of {total_files} files")

        # At the end of the function, check if the file exists
        if not os.path.exists(merged_output_file):
            raise FileNotFoundError(f"Failed to create merged file: {merged_output_file}")

        return merged_output_file

def get_planning_week(file_name):
    match = re.search(r'summary_file_w(?:ee)?k(\d+)', file_name)
    return int(match.group(1)) if match else None

def save_file_with_retry(self, pivot_table, suggested_filename):
    while True:
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Save Summary File", suggested_filename, "Excel Files (*.xlsx)"
        )
        
        if not file_path:  # User cancelled the dialog
            return None

        if self.try_save_file(file_path, pivot_table):
            return file_path
        else:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Icon.Warning)
            msg_box.setText("The selected file is currently open or inaccessible.")
            msg_box.setInformativeText("Would you like to try saving with a different name?")
            msg_box.setStandardButtons(QMessageBox.StandardButton.Retry | QMessageBox.StandardButton.Cancel)
            msg_box.setDefaultButton(QMessageBox.StandardButton.Retry)
            
            if msg_box.exec() == QMessageBox.StandardButton.Cancel:
                return None

def get_default_directory():
    return r"W:\Shared With Me\11. OTR\01_ShareFolder\output to bigpush"