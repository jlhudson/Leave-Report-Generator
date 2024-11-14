from datetime import timedelta

import pandas as pd


class LeaveStatusManager:
    def __init__(self):
        self.statuses = set()
        self.colors = ['FF0000', '00FF00', '0000FF', 'FFFF00', '00FFFF', 'FF00FF']
        self.status_colors = {}

    def add_status(self, status):
        """Add a new leave status to the set of known statuses."""
        self.statuses.add(status)

    def assign_colors(self):
        """Assign colors to statuses (sorted alphabetically)."""
        sorted_statuses = sorted(self.statuses)
        self.status_colors = {
            status: self.colors[i % len(self.colors)]
            for i, status in enumerate(sorted_statuses)
        }
        return self.status_colors

    def get_all_statuses(self):
        """Return all known leave statuses."""
        return sorted(list(self.statuses))


class LeaveDate:
    def __init__(self, date, status, shift_type):
        self.date = date
        self.status = status
        self.shift_type = shift_type

    def __repr__(self):
        return f"{self.date} ({self.status} - {self.shift_type})"


class WorkArea:
    def __init__(self, location, department, role):
        self.location = location
        self.department = department
        self.role = role

    def __repr__(self):
        return f"{self.location} -> {self.department} -> {self.role}"

    def __eq__(self, other):
        return (self.location, self.department, self.role) == (other.location, other.department, other.role)

    def __hash__(self):
        return hash((self.location, self.department, self.role))


class Employee:
    def __init__(self, emp_code, name, employment_type):
        self.emp_code = emp_code
        self.name = name
        self.employment_type = employment_type
        self.leave_dates = {}
        self.work_areas = set()

    def add_leave_date(self, leave_date, status, shift_type):
        if leave_date not in self.leave_dates or self.leave_dates[leave_date].status != status:
            self.leave_dates[leave_date] = LeaveDate(leave_date, status, shift_type)

    def add_work_area(self, work_area):
        self.work_areas.add(work_area)


class EmployeeManager:
    def __init__(self, leave_data, work_areas_data):
        self.leave_data = leave_data
        self.work_areas_data = work_areas_data
        self.employees = {}
        self.leave_status_manager = LeaveStatusManager()
        self.departments = set()
        # Define keywords for filtering out employees
        self.EXCLUDED_NAME_KEYWORDS = {'DNR', 'CANCELLED', 'XXX'}

    def should_exclude_employee(self, name):
        """Check if employee should be excluded based on name keywords."""
        upper_name = name.upper()
        return any(keyword in upper_name for keyword in self.EXCLUDED_NAME_KEYWORDS)

    def get_status_colors(self):
        """Get the color mapping for all leave statuses."""
        return self.leave_status_manager.status_colors

    def process_employees(self):
        self._process_work_areas()
        self._process_leave_dates()
        self.leave_status_manager.assign_colors()

    def _process_work_areas(self):
        skipped_employees = []
        for _, row in self.work_areas_data.iterrows():
            emp_code = str(row['Employee_Code'])
            name = row['Employee_Name']

            # Track skipped employees
            if self.should_exclude_employee(name):
                skipped_employees.append(f"{name} ({emp_code})")
                continue

            employment_type = row['Employment_Type_Name']
            work_area = WorkArea(row['Location'], row['Department'], row['Role'])

            # Track departments
            self.departments.add(row['Department'])

            if emp_code not in self.employees:
                self.employees[emp_code] = Employee(emp_code, name, employment_type)

            self.employees[emp_code].add_work_area(work_area)

        if skipped_employees:
            print("\nSkipped employees due to name filtering:")
            for emp in sorted(skipped_employees):
                print(f"  - {emp}")
            print(f"Total skipped: {len(skipped_employees)}")

    def _process_leave_dates(self):
        skipped_entries = []
        for _, row in self.leave_data.iterrows():
            emp_code = str(row['Employee_Code'])

            # Track skipped leave entries
            if emp_code not in self.employees:
                name = row['Employee_Name']
                skipped_entries.append(f"{name} ({emp_code})")
                continue

            start_time = pd.to_datetime(row['Start_Time'])
            end_time = pd.to_datetime(row['End_Time'])
            status = row['Status']
            shift_type = row['Shift_Type']

            self.leave_status_manager.add_status(status)

            if end_time.time() == pd.Timestamp('00:00:00').time():
                end_time -= timedelta(minutes=1)

            if start_time.date() == end_time.date():
                self.employees[emp_code].add_leave_date(start_time.date(), status, shift_type)
            else:
                for date in pd.date_range(start=start_time, end=end_time, freq='D'):
                    self.employees[emp_code].add_leave_date(date.date(), status, shift_type)

        if skipped_entries:
            print("\nSkipped leave entries due to unknown employee codes:")
            for entry in sorted(set(skipped_entries)):  # Use set to remove duplicates
                print(f"  - {entry}")
            print(f"Total skipped: {len(set(skipped_entries))}")

    def get_department_leave_counts(self, date):
        """Get leave counts per department and status for a specific date."""
        department_counts = {dept: {status: 0 for status in self.leave_status_manager.get_all_statuses()}
                             for dept in self.departments}

        for employee in self.employees.values():
            if date in employee.leave_dates:
                leave_status = employee.leave_dates[date].status
                for work_area in employee.work_areas:
                    department_counts[work_area.department][leave_status] += 1

        return department_counts

    def get_all_statuses(self):
        """Get all possible leave statuses."""
        return self.leave_status_manager.get_all_statuses()
