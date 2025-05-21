import pandas as pd
import io
from datetime import datetime
import copy

# ==== Sample Entry Template ====
SAMPLE_ENTRY = {
    "timesheet_for": "2025-04-17",
    "timesheet_user": {
        "id": 10,
        "full_name": "Sushmita Gautam",
        "email": "sushmita@aayulogic.com",
        "organization": {
            "name": "Aayu Bank Pvt. Ltd.",
            "abbreviation": "ABPL",
            "slug": "aayu-bank-pvt-ltd"
        },
        "is_online": "false",
        "last_online": "2024-12-10T12:25:31.205924+05:45",
        "is_audit_user": "false",
        "is_current": "true",
        "step": 1,
        "employee_code": "EMP7"
    },
    "timesheet_entries": [],
    "day": "Thursday",
    "punch_in": "N/A",
    "punch_out": "N/A",
    "expected_punch_in": "2025-04-17 10:00:00 AM",
    "expected_punch_out": "2025-04-17 05:00:00 PM",
    "worked_hours": "00:00:00",
    "expected_work_hours": "07:00:00",
    "overtime": "00:00:00",
    "punctuality": "N/A",
    "coefficient": "Workday",
    "leave_coefficient": "No Leave",
    "logs": "N/A",
    "late_in": "00:00",
    "early_out": "00:00",
    "total_lost_hours": "00:00"
}

# ==== Generate Many Sample Records ====
def generate_sample_data(n=100000):
    data = []
    for i in range(n):
        entry = copy.deepcopy(SAMPLE_ENTRY)
        entry["timesheet_user"]["full_name"] = f"Employee {i+1}"
        entry["timesheet_user"]["email"] = f"employee{i+1}@example.com"
        entry["timesheet_user"]["employee_code"] = f"EMP{i+1:05d}"
        data.append(entry)
    return data

# ==== Export Config ====
export_fields = {
    "timesheet_user.full_name": "Full Name",
    "timesheet_user.step": "Step/Grade",
    "timesheet_for": "Timesheet For",
    "day": "Day",
    "timesheet_user.email": "Email",
    "timesheet_user.employee_code": "Employee Code",
    "punch_in": "Punch In",
    "punch_out": "Punch Out",
    "expected_punch_out": "Expected Out Time",
    "expected_punch_in": "Expected In Time",
    "worked_hours": "Worked Hours",
    "expected_work_hours": "Expected Work Hours",
    "late_in": "Late In",
    "early_out": "Early Out",
    "overtime": "Overtime",
    "logs": "Logs",
    "timesheet_entries": "Timesheet Entries",
    "total_lost_hours": "Total Lost Hours",
    "punctuality": "Punctuality",
    "coefficient": "Shift Remarks"
}

export_title = "Daily Attendance"
export_filename = f"attendance_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

# ==== Helper Function ====
def get_nested_value(data, dotted_key):
    keys = dotted_key.split(".")
    for key in keys:
        data = data.get(key, {})
    return data if data != {} else ""

# ==== Export Function ====
def export_to_excel(data, fields, title, filename):
    processed_data = [
        {header: get_nested_value(item, key) for key, header in fields.items()}
        for item in data
    ]
    df = pd.DataFrame(processed_data)
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=title)
    print(f"Excel file '{filename}' created successfully!")

# ==== Run Export ====
if __name__ == "__main__":
    sample_data = generate_sample_data(1000)  # Generate 100,000 records
    export_to_excel(sample_data, export_fields, export_title, export_filename)
