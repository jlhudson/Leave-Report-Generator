import os
import shutil
from datetime import datetime
from scripts.file_loader import load_employee_data
from scripts.employee_leave_report_generator import generate_leave_report
from scripts.departmental_leave_report_generator import generate_departmental_leave_report


def safe_remove_directory(directory):
    """Safely remove a directory and its contents with error handling."""
    if not os.path.exists(directory):
        return

    try:
        shutil.rmtree(directory)
    except PermissionError:
        try:
            for root, dirs, files in os.walk(directory):
                for file in files:
                    try:
                        file_path = os.path.join(root, file)
                        os.unlink(file_path)
                    except Exception:
                        pass
            
            for root, dirs, files in os.walk(directory, topdown=False):
                for dir_name in dirs:
                    try:
                        dir_path = os.path.join(root, dir_name)
                        os.rmdir(dir_path)
                    except Exception:
                        pass
                        
            try:
                os.rmdir(directory)
            except Exception:
                pass
        except Exception:
            print("Warning: Directory cleanup incomplete - continuing with existing directory...")


def get_relative_path(full_path, base_dir):
    """Convert full path to relative path from base directory"""
    try:
        return os.path.relpath(full_path, base_dir)
    except:
        return full_path


# In report.py, update the main function:

def main():
    base_dir = os.path.dirname(__file__)
    reports_dir = os.path.join(base_dir, 'Humanforce Reports')
    
    # Create date-specific folder name
    current_date = datetime.now()
    folder_name = current_date.strftime("Leave Report %d %b %Y")
    output_dir = os.path.join(base_dir, folder_name)

    # Safely remove existing directory if it exists
    safe_remove_directory(output_dir)
    
    # Create fresh output directory
    try:
        os.makedirs(output_dir, exist_ok=True)
    except Exception:
        print("Warning: Using existing directory...")

    # Load employee data
    employee_data = load_employee_data(reports_dir)  # This returns a tuple of (employees, employee_manager)

    if employee_data is None:
        print("Error: Failed to load employee data.")
        return

    employees, employee_manager = employee_data  # Unpack the tuple
    print(f"Processed {len(employees)} employees")

    # Generate reports - pass both employees and employee_manager to both generators
    generate_leave_report(employees, output_dir, employee_manager)
    generate_departmental_leave_report(employees, output_dir, employee_manager)

if __name__ == "__main__":
    main()
    print("Report Generation Finished,  please close this window.")
