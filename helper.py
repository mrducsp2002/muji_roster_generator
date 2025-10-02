from datetime import datetime, timedelta
import pandas as pd
import unittest

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


def read_from_excel(day_of_week):
    # Read both sheets
    roster_df = pd.read_excel("modify_this.xlsx", sheet_name="Roster")
    dept_df = pd.read_excel("modify_this.xlsx", sheet_name="Department")

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
