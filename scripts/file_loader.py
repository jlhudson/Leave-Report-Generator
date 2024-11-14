import os
import pandas as pd
from tkinter import Tk, messagebox
from scripts.employee_manager import EmployeeManager

import warnings
# Suppress the specific warning about header/footer parsing
warnings.filterwarnings("ignore", message="Cannot parse header or footer so it will be ignored")

def show_error_popup(message):
    root = Tk()
    root.withdraw()
    messagebox.showerror("Error", message)
    root.destroy()

def find_excel_files(directory):
    excel_files = [f for f in os.listdir(directory) if f.endswith(('.xlsx', '.xls'))]
    if len(excel_files) != 2:
        show_error_popup(f"Error: Expected 2 Excel files, found {len(excel_files)}")
        return None, None
    return os.path.join(directory, excel_files[0]), os.path.join(directory, excel_files[1])

def validate_headers(df, headers):
    return all(header in df.columns for header in headers)

def load_and_validate_data(file_path, leave_headers, work_areas_headers):
    try:
        df = pd.read_excel(file_path)
        print(f"Loaded: {os.path.basename(file_path)} ({len(df)} rows)")
    except Exception as e:
        show_error_popup(f"Error loading {file_path}: {str(e)}")
        return None, ""

    if validate_headers(df, leave_headers):
        return df, "Employee Leave"
    elif validate_headers(df, work_areas_headers):
        return df, "Employee Work Areas"
    else:
        show_error_popup(f"Error: {file_path} has invalid headers")
        return None, ""

def load_employee_data(directory):
    leave_headers = ['Employee_Code', 'Employee_Name', 'Shift_Type', 'Start_Time', 'End_Time', 'Status']
    work_areas_headers = ['Employee_Code', 'Employee_Name', 'Employment_Type_Name', 'Location', 'Department', 'Role']

    print("\nProcessing Excel files...")
    file1, file2 = find_excel_files(directory)
    if not file1 or not file2:
        return None

    df1, type1 = load_and_validate_data(file1, leave_headers, work_areas_headers)
    df2, type2 = load_and_validate_data(file2, leave_headers, work_areas_headers)

    if df1 is None or df2 is None:
        return None

    if type1 == "Employee Leave" and type2 == "Employee Work Areas":
        leave_data, work_areas_data = df1, df2
    elif type1 == "Employee Work Areas" and type2 == "Employee Leave":
        leave_data, work_areas_data = df2, df1
    else:
        show_error_popup("Error: Unable to determine file types")
        return None

    print("\nInitializing employee manager...")
    employee_manager = EmployeeManager(leave_data=leave_data, work_areas_data=work_areas_data)
    employee_manager.process_employees()

    return employee_manager.employees