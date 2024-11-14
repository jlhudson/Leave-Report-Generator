import os
import pandas as pd
import tkinter as tk
from tkinter import messagebox

def show_error_popup(message: str):
    """Display error message in a popup window."""
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Error", message)
    root.destroy()

def find_excel_files(directory: str):
    """Find and validate the presence of exactly 2 Excel files."""
    if not os.path.exists(directory):
        show_error_popup(f"Error: The 'Reports' folder does not exist at {directory}")
        return None, None

    excel_files = [f for f in os.listdir(directory) if f.endswith(('.xlsx', '.xls'))]

    if len(excel_files) != 2:
        show_error_popup(f"Error: Expected 2 Excel files, found {len(excel_files)}")
        return None, None

    return tuple(os.path.join(directory, f) for f in excel_files)

def load_excel_file(file_path: str):
    """Load an Excel file into a pandas DataFrame."""
    try:
        engine = 'openpyxl' if file_path.endswith('.xlsx') else 'xlrd'
        return pd.read_excel(file_path, engine=engine)
    except PermissionError:
        show_error_popup(
            f"Error: Unable to open '{os.path.basename(file_path)}'. The file may be open in another program.")
    except Exception as e:
        show_error_popup(f"Error reading {os.path.basename(file_path)}: {str(e)}")
    return None

def validate_headers(df: pd.DataFrame, expected_headers: list):
    """Check if DataFrame contains all expected headers."""
    return all(header in df.columns for header in expected_headers)

def load_and_validate_data(file_path: str, leave_headers: list, work_areas_headers: list):
    """Load and validate Excel file against expected formats."""
    df = load_excel_file(file_path)

    if df is not None:
        if validate_headers(df, leave_headers):
            print(f"Successfully loaded and validated {os.path.basename(file_path)} as Employee Leave")
            return df, "Employee Leave"
        elif validate_headers(df, work_areas_headers):
            print(f"Successfully loaded and validated {os.path.basename(file_path)} as Employee Work Areas")
            return df, "Employee Work Areas"
        else:
            show_error_popup(
                f"Error: {os.path.basename(file_path)} does not have the expected headers for either dataset")
    return None, ""

def validate_excel_files(reports_folder: str):
    """Main function to validate and load Excel files."""
    # Define expected headers for both file types
    leave_headers = ['Employee_Code', 'Employee_Name', 'Shift_Type', 'Start_Time', 'End_Time', 'Status']
    work_areas_headers = ['Employee_Code', 'Employee_Name', 'Employment_Type_Name', 'Location', 'Department', 'Role']

    # Find Excel files
    file1, file2 = find_excel_files(reports_folder)
    if not file1 or not file2:
        return None, None

    # Load and validate both files
    df1, type1 = load_and_validate_data(file1, leave_headers, work_areas_headers)
    df2, type2 = load_and_validate_data(file2, leave_headers, work_areas_headers)

    if df1 is None or df2 is None:
        return None, None

    # Determine which file is which type
    if type1 == "Employee Leave" and type2 == "Employee Work Areas":
        return df1, df2
    elif type1 == "Employee Work Areas" and type2 == "Employee Leave":
        return df2, df1
    else:
        show_error_popup("Error: Unable to determine which file is which")
        return None, None