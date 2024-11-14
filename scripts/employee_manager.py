import pandas as pd
from datetime import timedelta

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

    def process_employees(self):
        self._process_work_areas()
        self._process_leave_dates()

    def _process_work_areas(self):
        for _, row in self.work_areas_data.iterrows():
            emp_code = str(row['Employee_Code'])
            name = row['Employee_Name']
            employment_type = row['Employment_Type_Name']
            work_area = WorkArea(row['Location'], row['Department'], row['Role'])

            if emp_code not in self.employees:
                self.employees[emp_code] = Employee(emp_code, name, employment_type)

            self.employees[emp_code].add_work_area(work_area)

    def _process_leave_dates(self):
        for _, row in self.leave_data.iterrows():
            emp_code = str(row['Employee_Code'])
            start_time = pd.to_datetime(row['Start_Time'])
            end_time = pd.to_datetime(row['End_Time'])
            status = row['Status']
            shift_type = row['Shift_Type']

            if emp_code not in self.employees:
                continue

            if end_time.time() == pd.Timestamp('00:00:00').time():
                end_time -= timedelta(minutes=1)

            if start_time.date() == end_time.date():
                self.employees[emp_code].add_leave_date(start_time.date(), status, shift_type)
            else:
                for date in pd.date_range(start=start_time, end=end_time, freq='D'):
                    self.employees[emp_code].add_leave_date(date.date(), status, shift_type)