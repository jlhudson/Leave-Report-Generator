import os
from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter


class DepartmentalLeaveReportGenerator:
    def __init__(self, employees):
        self.employees = {code: emp for code, emp in employees.items() if len(emp.leave_dates) > 0}
        self.wb = Workbook()
        self.ws = None
        self.colors = ['FF0000', '00FF00', '0000FF', 'FFFF00', '00FFFF', 'FF00FF']
        self.status_colors = {}
        self.COMBINED_LOCATIONS = [
            ("ADELAIDE HILLS & STRATHALBYN", ["ADELAIDE HILLS", "STRATHALBYN"]),
        ]

    def _get_all_leave_dates(self):
        """Collect all unique leave dates from all employees."""
        all_dates = set()
        for employee in self.employees.values():
            all_dates.update(employee.leave_dates.keys())
        return sorted(list(all_dates))

    def _format_date_header(self, date):
        """Format date as 'Mon 25/11'."""
        return date.strftime('%a %d/%m')

    def _format_employee_name(self, full_name):
        """Format employee name as 'FirstName L.'"""
        parts = full_name.strip().split()
        if len(parts) > 1:
            return f"{parts[0]} {parts[-1][0]}."
        return full_name

    def _assign_status_colors(self):
        """Assign colors to unique statuses."""
        unique_statuses = set()
        for employee in self.employees.values():
            for leave_date in employee.leave_dates.values():
                unique_statuses.add(leave_date.status)

        for i, status in enumerate(sorted(unique_statuses)):
            color = self.colors[i % len(self.colors)]
            self.status_colors[status] = color

    def _get_location_departments(self, location_employees):
        """Get all departments for a location and their employees."""
        dept_employees = {}
        for emp in location_employees.values():
            for work_area in emp.work_areas:
                if work_area.department not in dept_employees:
                    dept_employees[work_area.department] = []
                dept_employees[work_area.department].append(emp)

        # Remove duplicates and sort by name within each department
        for dept in dept_employees:
            dept_employees[dept] = sorted(
                list(set(dept_employees[dept])),
                key=lambda x: x.name
            )

        return dict(sorted(dept_employees.items()))

    def _generate_worksheet(self, ws_name, location_employees, all_dates):
        """Generate a worksheet for the given location."""
        ws = self.wb.create_sheet(ws_name)

        # Define styles
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        thick_border = Border(
            left=Side(style='medium'),
            right=Side(style='medium'),
            top=Side(style='medium'),
            bottom=Side(style='medium')
        )

        header_fill = PatternFill(start_color='CCCCCC', end_color='CCCCCC', fill_type='solid')

        # Get departments and their employees
        dept_employees = self._get_location_departments(location_employees)

        current_row = 1

        # Write date headers
        date_headers = [''] + [self._format_date_header(date) for date in all_dates]
        for col, header in enumerate(date_headers, 1):
            cell = ws.cell(row=current_row, column=col)
            cell.value = header
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center')
            cell.border = thick_border

        # Process each department
        for dept_name, dept_emps in dept_employees.items():
            current_row += 2  # Add spacing between departments

            # Write department header
            cell = ws.cell(row=current_row, column=1)
            cell.value = dept_name
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.border = thick_border

            # Write employee rows for this department
            for employee in dept_emps:
                current_row += 1

                # Write employee name in first column
                cell = ws.cell(row=current_row, column=1)
                cell.value = self._format_employee_name(employee.name)
                cell.border = thin_border

                # Write leave status for each date
                for col, date in enumerate(all_dates, 2):
                    cell = ws.cell(row=current_row, column=col)
                    cell.border = thin_border

                    if date in employee.leave_dates:
                        leave_date = employee.leave_dates[date]
                        cell.value = self._format_employee_name(employee.name)
                        cell.fill = PatternFill(
                            start_color=self.status_colors[leave_date.status],
                            end_color=self.status_colors[leave_date.status],
                            fill_type='solid'
                        )
                        cell.alignment = Alignment(horizontal='center')

        # Add legend
        current_row += 3
        ws.cell(row=current_row, column=1, value="Legend:").font = Font(bold=True)
        current_row += 1
        for status, color in self.status_colors.items():
            cell = ws.cell(row=current_row, column=1)
            cell.value = status
            cell.alignment = Alignment(horizontal='left')

            color_cell = ws.cell(row=current_row, column=2)
            color_cell.fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
            current_row += 1

        # Adjust column widths
        ws.column_dimensions['A'].width = 30  # Names column
        for col in range(2, len(all_dates) + 2):
            ws.column_dimensions[get_column_letter(col)].width = 12

    def _get_employees_for_location(self, location):
        """Get employees for a specific location or combined locations."""
        if isinstance(location, tuple):
            # Handle combined locations
            combined_name, location_list = location
            return {
                code: emp for code, emp in self.employees.items()
                if any(area.location in location_list for area in emp.work_areas)
            }
        else:
            # Handle single location
            return {
                code: emp for code, emp in self.employees.items()
                if any(area.location == location for area in emp.work_areas)
            }

    def _get_all_locations(self):
        """Get all unique locations including combined ones."""
        locations = set()
        for employee in self.employees.values():
            for work_area in employee.work_areas:
                locations.add(work_area.location)
        return sorted(list(locations))

    def generate_report(self):
        """Generate the departmental leave report in Excel format."""
        # Assign colors to statuses
        self._assign_status_colors()

        # Get all unique dates
        all_dates = self._get_all_leave_dates()

        # Remove default sheet if it exists
        if 'Sheet' in self.wb.sheetnames:
            self.wb.remove(self.wb['Sheet'])

        # Generate location-specific worksheets
        locations = self._get_all_locations()
        for location in locations:
            filtered_employees = self._get_employees_for_location(location)
            if filtered_employees:  # Only create sheet if there are employees
                safe_name = str(location)[:31]  # Excel worksheet names limited to 31 chars
                self._generate_worksheet(safe_name, filtered_employees, all_dates)

        # Generate combined location worksheets
        for combined_name, location_list in self.COMBINED_LOCATIONS:
            filtered_employees = self._get_employees_for_location((combined_name, location_list))
            if filtered_employees:
                safe_name = combined_name[:31]
                self._generate_worksheet(safe_name, filtered_employees, all_dates)

    def save_report(self, directory):
        """Save the report with the specified naming convention."""
        current_date = datetime.now()
        month_year = current_date.strftime("%b %Y").title()
        filename = f"Departmental Leave Report {month_year}.xlsx"
        filepath = os.path.join(directory, filename)
        self.wb.save(filepath)
        print(f"Report saved: {filepath}")


def generate_departmental_leave_report(employees, output_directory):
    """Main function to generate and save the departmental leave report."""
    report_generator = DepartmentalLeaveReportGenerator(employees)
    report_generator.generate_report()
    report_generator.save_report(output_directory)