from datetime import datetime, timedelta
import pandas as pd
import unittest
import os
import glob

# -----------------------------
# Helper
# -----------------------------

# Generate global day slots at 15-minute intervals (0 to 95 representing 00:00 to 23:45)
def generate_day_slots(interval=15):
    fmt = "%H:%M"
    slots = []
    start_dt = datetime.strptime("00:00", fmt)
    for i in range(24 * 60 // interval):  # 96 slots
        slots.append(start_dt.strftime(fmt))
        start_dt += timedelta(minutes=interval)
    return slots

day_slots = generate_day_slots()
# day_slots[0] == "00:00", day_slots[95] == "23:45"

def time_to_slot(time_str, interval=15):
    h, m = map(int, time_str.split(":"))
    return (h * 60 + m) // interval


def timespan_to_slot(time_span: tuple):
    """
    Take in a tuple of (start_time, end_time) in "HH:MM" format.
    Return a range of the corresponding 15-minute slots in 24 hours.

    Example:
        ("09:30", "18:00") -> range(38, 72)
    """
    start, end = time_span
    return range(time_to_slot(start), time_to_slot(end))


def select_excel_file():
    """Prompt user to select an Excel file from the current directory"""
    # Get all Excel files in the current directory
    excel_files = glob.glob("*.xlsx")
    excel_files = [f for f in excel_files if not f.startswith("~$")]  # Exclude temp files
    
    if not excel_files:
        print("No Excel files found in the current directory.")
        return None
    
    print("Available Excel files:")
    for i, file in enumerate(excel_files, 1):
        print(f"{i}. {file}")
    
    while True:
        try:
            choice = input(f"\nSelect a file (1-{len(excel_files)}): ")
            choice_idx = int(choice) - 1
            if 0 <= choice_idx < len(excel_files):
                selected_file = excel_files[choice_idx]
                print(f"Selected: {selected_file}")
                return selected_file
            else:
                print(f"Please enter a number between 1 and {len(excel_files)}")
        except ValueError:
            print("Please enter a valid number")

def read_from_excel_new_format(file_path, day_of_week):
    """Read employee data from the new Excel format with Weekly sheet"""
    try:
        # Read the Weekly sheet
        df = pd.read_excel(file_path, sheet_name="Weekly", header=None)
        
        # Read the Team sheet to get department information
        team_df = pd.read_excel(file_path, sheet_name="Team", header=2)
        
        # Create employee name to department mapping
        employee_dept_mapping = {}
        for _, row in team_df.iterrows():
            if pd.notna(row['Employee Id']) and pd.notna(row['First Name']) and pd.notna(row['Last Name']):
                # Create employee name (same format as in Weekly sheet)
                employee_name = f"{row['First Name']} {row['Last Name']}".strip()
                department = row['Department'] if pd.notna(row['Department']) else ""
                employee_dept_mapping[employee_name] = department
        
        # Find the header row (row 7 based on our analysis)
        header_row = 7
        headers = df.iloc[header_row].tolist()
        
        # Map day names to their Start column indices (based on actual Excel structure)
        day_start_columns = {
            "M": 6,   # Monday Start column
            "T": 9,   # Tuesday Start column
            "W": 12,  # Wednesday Start column
            "Th": 15, # Thursday Start column
            "F": 18,  # Friday Start column
            "Sa": 21, # Saturday Start column
            "Su": 24  # Sunday Start column
        }
        
        if day_of_week not in day_start_columns:
            raise ValueError(f"Invalid day: {day_of_week}. Must be one of: {list(day_start_columns.keys())}")
        
        start_col = day_start_columns[day_of_week]
        end_col = start_col + 1
        hours_col = start_col + 2
        
        working_employees = {}
        
        # Process data starting from row 8 (after header)
        for idx in range(header_row + 1, len(df)):
            row = df.iloc[idx]
            
            # Skip empty rows
            if pd.isna(row[1]):  # ID column
                continue
                
            # Extract employee info
            employee_id = str(row[1]) if not pd.isna(row[1]) else ""
            first_name = str(row[2]) if not pd.isna(row[2]) else ""
            last_name = str(row[3]) if not pd.isna(row[3]) else ""
            contract = str(row[4]) if not pd.isna(row[4]) else ""
            
            # Create employee name
            employee_name = f"{first_name} {last_name}".strip()
            if not employee_name:
                continue
            
            # Extract shift information for the selected day
            start_time = row[start_col] if not pd.isna(row[start_col]) else None
            end_time = row[end_col] if not pd.isna(row[end_col]) else None
            hours = row[hours_col] if not pd.isna(row[hours_col]) else 0
            
            # Only include employees who have a shift on the selected day
            if start_time is not None and end_time is not None:
                # Convert hours to float for comparison
                try:
                    hours_float = float(hours) if hours != 0 else 0
                    if hours_float > 0:
                        # Convert time objects to string format
                        if hasattr(start_time, 'strftime'):
                            start_str = start_time.strftime("%H:%M")
                        else:
                            start_str = str(start_time)
                            
                        if hasattr(end_time, 'strftime'):
                            end_str = end_time.strftime("%H:%M")
                        else:
                            end_str = str(end_time)
                        
                        # Get department from Team sheet mapping
                        department = employee_dept_mapping.get(employee_name, contract)
                        
                        working_employees[employee_name] = {
                            "shift": (start_str, end_str),
                            "department": department,  # Now using actual department from Team sheet
                            "employee_id": employee_id,
                            "hours": hours_float
                        }
                except (ValueError, TypeError):
                    # Skip if hours can't be converted to float
                    continue
        
        return working_employees
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return {}

def read_from_excel(day_of_week, file_path=None):
    """Read employee data from Excel file - supports both old and new formats"""
    if file_path is None:
        # Try new format first
        excel_files = glob.glob("*.xlsx")
        excel_files = [f for f in excel_files if not f.startswith("~$")]
        
        if excel_files:
            file_path = excel_files[0]  # Use first available file
        else:
            print("No Excel files found in the current directory.")
            return {}
    
    # Try new format first
    try:
        # Check if Weekly sheet exists
        df_sheets = pd.read_excel(file_path, sheet_name=None)
        if "Weekly" in df_sheets:
            return read_from_excel_new_format(file_path, day_of_week)
    except:
        pass
    
    # Fall back to old format
    try:
        # Read both sheets
        roster_df = pd.read_excel(file_path, sheet_name="Roster")
        dept_df = pd.read_excel(file_path, sheet_name="Department")

        # Replace NaN with empty string for blanks
        roster_df = roster_df.fillna("")
        dept_df = dept_df.fillna("")

        dept_lookup = {}
        for _, row in dept_df.iterrows():
            dept_lookup[row["Employee"]] = row["Department"]

        working_employees = {}

        for _, row in roster_df.iterrows():
            name = row["Employee"]
            shift = row[day_of_week]

            if shift:
                start, end = shift.split("-")
                working_employees[name] = {
                    "shift": (start, end),
                    "department": dept_lookup.get(name, "")
                }

        return working_employees
    except Exception as e:
        print(f"Error reading Excel file in old format: {e}")
        return {}

# -----------------------------
# Tester
# -----------------------------

class TestTimeSpanToSlot(unittest.TestCase):
    def test_full_hour(self):
        self.assertEqual(timespan_to_slot(("09:00", "17:00")), (36, 68))
        # 9:00 = 9*60 = 540 → 540/15 = 36
        # 17:00 = 1020 → 1020/15 = 68

    def test_half_hour(self):
        self.assertEqual(timespan_to_slot(("09:30", "18:00")), (38, 72))

    def test_midnight(self):
        self.assertEqual(timespan_to_slot(("00:00", "00:15")), (0, 1))

    def test_end_of_day(self):
        self.assertEqual(timespan_to_slot(("23:30", "23:45")), (94, 95))


if __name__ == "__main__":
    unittest.main()
