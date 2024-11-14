import os
from scripts.file_loader import load_employee_data
from scripts.leave_report_generator import generate_leave_report
from scripts.departmental_leave_report_generator import generate_departmental_leave_report


def main():
    base_dir = os.path.dirname(__file__)
    reports_dir = os.path.join(base_dir, 'reports')
    output_dir = os.path.join(base_dir, 'output')

    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)

    # Load employee data
    employees = load_employee_data(reports_dir)

    if employees is None:
        print("Error: Failed to load employee data.")
        return

    print(f"\nProcessed {len(employees)} employees:")
    for emp_code, employee in employees.items():
        print(f"{employee.name}: {len(employee.leave_dates)} leave dates, {len(employee.work_areas)} work areas")

    # Generate original leave report
    print("\nGenerating employee leave report...")
    generate_leave_report(employees, output_dir)
    print("Employee leave report generation complete.")

    # Generate new departmental leave report
    print("\nGenerating departmental leave report...")
    generate_departmental_leave_report(employees, output_dir)
    print("Departmental leave report generation complete.")


if __name__ == "__main__":
    main()