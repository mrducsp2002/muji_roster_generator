from datetime import datetime, timedelta
import os
from helper import timespan_to_slot
from collections import defaultdict
import pandas as pd
from openpyxl.styles import PatternFill

def slot_to_time(slot):
    """Convert slot number back to time string - kept for internal use only"""
    # Assuming slots start at 08:00 and each slot is 15 minutes
    start_time = datetime.strptime("08:00", "%H:%M")
    slot_time = start_time + timedelta(minutes=slot * 15)
    return slot_time.strftime("%H:%M")


def add_15_minutes(time_str):
    """Add 15 minutes to a time string - kept for potential future use"""
    time_obj = datetime.strptime(time_str, "%H:%M")
    new_time = time_obj + timedelta(minutes=15)
    return new_time.strftime("%H:%M")


def format_task_name(task):
    """Format task/department names for display"""
    task_names = {
        "FR": "Fitting Room",
        "GR": "Greeter",
        "R": "Register",
        "Break": "Break",
        "H": "H",
        "HH": "Home & Hardware",
        "L": "Ladies",
        "M": "Mens",
        "H&B": "Health & Beauty",
        "40": "Break (40min)",
        "10": "Break (10min)"
    }
    return task_names.get(task, task)


def print_roster_header(current_day, store_hours):
    """Print formatted header for the roster"""
    day_names = {
        "M": "Monday", "T": "Tuesday", "W": "Wednesday",
        "Th": "Thursday", "F": "Friday", "Sa": "Saturday", "Su": "Sunday"
    }

    print("=" * 80)
    print(f"STORE ROSTER - {day_names[current_day].upper()}")
    print(
        f"Store Hours: {store_hours[current_day][0]} - {store_hours[current_day][1]}")
    print("=" * 80)


def print_coverage_summary(roster, current_day, store_hours):
    """Print coverage summary for customer service tasks"""
    print("\n📋 CUSTOMER SERVICE COVERAGE SUMMARY")
    print("-" * 50)

    store_opening_slots = timespan_to_slot(store_hours[current_day])
    uncovered_slots = {"FR": [], "GR": [], "R": []}

    for slot in store_opening_slots:
        for task in ["FR", "GR", "R"]:
            if len(roster[slot][task]) == 0:
                uncovered_slots[task].append(slot)

    task_names = {"FR": "Fitting Room", "GR": "Greeter", "R": "Register"}

    for task, task_name in task_names.items():
        if uncovered_slots[task]:
            uncovered_slots_str = [
                f"Slot {slot}" for slot in uncovered_slots[task]]
            print(
                f"⚠️  {task_name}: UNCOVERED at {', '.join(uncovered_slots_str)}")
        else:
            print(f"✅ {task_name}: Fully covered")


def print_employee_schedule(roster, current_day, working_employees):
    """Print individual employee schedules"""
    print("\n👥 INDIVIDUAL EMPLOYEE SCHEDULES")
    print("-" * 50)

    # Get all employees and their shifts
    employee_schedules = defaultdict(list)

    for slot in roster:
        for task_or_dept in roster[slot]:
            for employee in roster[slot][task_or_dept]:
                employee_schedules[employee].append((slot, task_or_dept))

    # Sort employees by name
    for employee in sorted(employee_schedules.keys()):
        print(
            f"{employee} ({working_employees[employee]['department']} - Shift {working_employees[employee]['shift']}):")

        # Group consecutive same tasks
        schedule = employee_schedules[employee]
        if schedule:
            # Sort by slot number
            schedule.sort(key=lambda x: x[0])

            current_task = schedule[0][1]
            start_slot = schedule[0][0]
            end_slot = schedule[0][0]

            for i, (slot, task) in enumerate(schedule[1:], 1):
                if task == current_task and slot == end_slot + 1:
                    end_slot = slot
                else:
                    # Print previous task block
                    if start_slot == end_slot:
                        print(
                            f"  Slot {start_slot}: {format_task_name(current_task)}")
                    else:
                        print(
                            f"  Slots {start_slot}-{end_slot}: {format_task_name(current_task)}")
                    current_task = task
                    start_slot = slot
                    end_slot = slot

            # Print the last task block
            if start_slot == end_slot:
                print(f"  Slot {start_slot}: {format_task_name(current_task)}")
            else:
                print(
                    f"  Slots {start_slot}-{end_slot}: {format_task_name(current_task)}")
        else:
            print("  No assignments")


def print_hourly_breakdown(roster, current_day, store_hours):
    """Print slot-by-slot breakdown showing who's doing what"""
    print("\n🕒 SLOT-BY-SLOT BREAKDOWN")
    print("-" * 80)

    store_opening_slots = timespan_to_slot(store_hours[current_day])

    for slot in roster:
        if slot in store_opening_slots:
            print(f"Slot {slot}:")

            # Customer Service Tasks
            cs_tasks = []
            for task in ["FR", "GR", "R"]:
                if roster[slot][task]:
                    cs_tasks.append(f"{task}: {', '.join(roster[slot][task])}")
                else:
                    cs_tasks.append(f"{task}: EMPTY")

            print(f"  CS: {' | '.join(cs_tasks)}")

            # Other activities
            other_activities = []
            for activity in ["Break", "H"]:
                if roster[slot][activity]:
                    other_activities.append(
                        f"{activity}: {', '.join(roster[slot][activity])}")

            if other_activities:
                print(f"  Other: {' | '.join(other_activities)}")

            # Department work
            dept_work = []
            for dept in ["HH", "L", "M", "H&B"]:
                if roster[slot][dept]:
                    dept_work.append(
                        f"{dept}: {', '.join(roster[slot][dept])}")

            if dept_work:
                print(f"  Dept: {' | '.join(dept_work)}")

            print()  # Add space between slots


def print_register_coverage(roster, current_day, store_hours):
    """Print register staffing levels for each slot"""
    print("\n🧾 REGISTER STAFFING LEVELS")
    print("-" * 50)

    store_opening_slots = timespan_to_slot(store_hours[current_day])

    for slot in store_opening_slots:
        register_staff = roster[slot]["R"]
        staff_count = len(register_staff)

        if staff_count == 0:
            status = "⚠️  NO COVERAGE"
        elif staff_count == 1:
            status = "✅ MINIMAL"
        elif staff_count <= 3:
            status = "✅ GOOD"
        else:
            status = "🔥 HEAVY"

        time_str = slot_to_time(slot)
        print(f"Slot {slot} ({time_str}): {staff_count} people - {status}")
        if register_staff:
            print(f"  Staff: {', '.join(register_staff)}")
        print()


def print_statistics(roster, current_day, store_hours, working_employees):
    """Print useful statistics"""
    print("\n📊 ROSTER STATISTICS")
    print("-" * 40)

    store_opening_slots = timespan_to_slot(store_hours[current_day])

    # Count task assignments per employee
    employee_task_counts = defaultdict(lambda: defaultdict(int))
    total_cs_slots = 0

    for slot in store_opening_slots:
        for task in ["FR", "GR", "R"]:
            for employee in roster[slot][task]:
                employee_task_counts[employee][task] += 1
                total_cs_slots += 1

    print("Customer Service Task Distribution:")
    for employee in sorted(working_employees.keys()):
        fr_count = employee_task_counts[employee]["FR"]
        gr_count = employee_task_counts[employee]["GR"]
        r_count = employee_task_counts[employee]["R"]
        total = fr_count + gr_count + r_count

        if total > 0:
            print(
                f"  {employee}: FR({fr_count}) GR({gr_count}) R({r_count}) = {total} slots")

    print(f"\nTotal CS slots filled: {total_cs_slots}")
    print(f"Required CS slots: {len(store_opening_slots) * 3}")
    coverage_percentage = (
        total_cs_slots / (len(store_opening_slots) * 3)) * 100
    print(f"Coverage: {coverage_percentage:.1f}%")

    # Register staffing summary
    register_counts = []
    for slot in store_opening_slots:
        register_counts.append(len(roster[slot]["R"]))

    if register_counts:
        print(f"\nRegister Staffing Summary:")
        print(
            f"  Average staff per slot: {sum(register_counts) / len(register_counts):.1f}")
        print(f"  Minimum staff: {min(register_counts)}")
        print(f"  Maximum staff: {max(register_counts)}")

        # Count slots by staffing level
        no_coverage = sum(1 for count in register_counts if count == 0)
        minimal = sum(1 for count in register_counts if count == 1)
        good = sum(1 for count in register_counts if 2 <= count <= 3)
        heavy = sum(1 for count in register_counts if count > 3)

        print(f"  Slots with no coverage: {no_coverage}")
        print(f"  Slots with minimal coverage (1): {minimal}")
        print(f"  Slots with good coverage (2-3): {good}")
        print(f"  Slots with heavy coverage (4+): {heavy}")


def print_cs_coverage_totals(roster, current_day, store_hours):
    """Print total coverage count for each CS task per slot"""
    print("\n CS TASK COVERAGE TOTALS")
    print("-" * 50)

    store_opening_slots = timespan_to_slot(store_hours[current_day])

    print(f"{'Slot':<6} {'FR':<4} {'GR':<4} {'R':<4}")
    print("-" * 20)

    for slot in store_opening_slots:
        fr_count = len(roster[slot]["FR"])
        gr_count = len(roster[slot]["GR"])
        r_count = len(roster[slot]["R"])

        print(f"{slot:<6} {fr_count:<4} {gr_count:<4} {r_count:<4}")


def export_roster_to_excel(roster, current_day, working_employees, filename=None):
    output_dir = "roster_output"
    os.makedirs(output_dir, exist_ok=True)  # create folder if not exists

    if filename is None:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"roster_{current_day}_{timestamp}.xlsx"

    # Ensure path points into roster_output
    filepath = os.path.join(output_dir, filename)

    # Prepare employee schedules
    slots = sorted(roster.keys())
    employee_schedules = {emp: {slot: "" for slot in slots}
                          for emp in working_employees.keys()}

    for slot in slots:
        for task_or_dept, employees in roster[slot].items():
            for emp in employees:
                if emp in employee_schedules:
                    task_name = task_or_dept
                    if employee_schedules[emp][slot]:
                        employee_schedules[emp][slot] += f" + {task_name}"
                    else:
                        employee_schedules[emp][slot] = task_name

    # Sort employees by shift then name
    sorted_employees = sorted(employee_schedules.keys(),
                              key=lambda emp: (working_employees.get(emp, {}).get("shift", 999), emp))
    sorted_employee_schedules = {
        emp: employee_schedules[emp] for emp in sorted_employees}

    df = pd.DataFrame.from_dict(sorted_employee_schedules, orient='index')

    # Column headers with slot info
    time_columns = {slot: f"{slot}" for slot in slots}
    df = df.rename(columns=time_columns)

    # Add department and shift
    dept_info, shift_info = [], []
    for emp in df.index:
        if emp in working_employees:
            dept_info.append(working_employees[emp]["department"])
            shift_info.append(f"Shift {working_employees[emp]['shift']}")
        else:
            dept_info.append("")
            shift_info.append("")
    df.insert(0, "Department", dept_info)
    df.insert(1, "Shift", shift_info)

    # Define colors for tasks and departments
    color_map = {
        # Tasks
        "FR": "FFC7CE",      # Fitting Room - light red
        "GR": "C6EFCE",      # Greeter - light green
        "R": "FFEB9C",       # Register - light yellow
        "40": "D9D9D9",      # Break - gray
        "10": "D9D9D9",
        "H": "BDD7EE",  # Hurdle/Setup - light blue
        # Departments
        "CS": "FFE4B5",      # Customer Service - light orange
        "FR Dept": "D8BFD8",  # Example, can add more
        "GR Dept": "ADD8E6",
        # Shift
        "Shift 0": "FFFFFF",  # optional: white
        "Shift 1": "F0F8FF"
    }

    try:
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            df.to_excel(
                writer, sheet_name=f'{current_day}_Schedule', index=True, index_label='Employee')
            worksheet = writer.sheets[f'{current_day}_Schedule']

            # Auto-adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                worksheet.column_dimensions[column_letter].width = min(
                    max_length + 2, 30)

            # Apply colors to all cells
            # start from column 2 (Department)
            for row in worksheet.iter_rows(min_row=2, min_col=2):
                for cell in row:
                    if cell.value:
                        # If multiple tasks, pick first for coloring
                        first_value = str(cell.value).split(" + ")[0]
                        fill_color = color_map.get(
                            first_value, "FFFFFF")  # default white
                        cell.fill = PatternFill(start_color=fill_color,
                                                end_color=fill_color,
                                                fill_type="solid")

            # Also color the Employee index column (optional)
            for row_idx, emp in enumerate(df.index, start=2):
                cell = worksheet.cell(row=row_idx, column=1)
                dept = working_employees.get(emp, {}).get("department", "")
                fill_color = color_map.get(dept, "FFFFFF")
                cell.fill = PatternFill(start_color=fill_color,
                                        end_color=fill_color,
                                        fill_type="solid")

            # Create summary sheet if needed
            create_summary_sheet(writer, roster, current_day, slots)

        print(f"✅ Roster exported to: {filepath}")
        return filepath

    except Exception as e:
        print(f"❌ Error exporting to Excel: {str(e)}")
        return None

def create_summary_sheet(writer, roster, current_day, slots):
    """Create a summary sheet with coverage statistics"""

    summary_data = []

    # Coverage summary for each slot
    for slot in slots:
        slot_summary = {"Slot": slot, "Time": slot_to_time(slot)}

        # Count coverage for each task
        slot_summary["Fitting Room"] = len(roster[slot].get("FR", []))
        slot_summary["Greeter"] = len(roster[slot].get("GR", []))
        slot_summary["Register"] = len(roster[slot].get("R", []))
        slot_summary["On Break"] = len(roster[slot].get(
            "40", [])) + len(roster[slot].get("10", []))
        slot_summary["H"] = len(roster[slot].get("H", []))

        # Department coverage
        slot_summary["Home & Hardware"] = len(roster[slot].get("HH", []))
        slot_summary["Ladies"] = len(roster[slot].get("L", []))
        slot_summary["Mens"] = len(roster[slot].get("M", []))
        slot_summary["Health & Beauty"] = len(roster[slot].get("H&B", []))

        # Total employees working
        total_working = sum([
            len(employees) for task, employees in roster[slot].items()
            if task not in ["40", "10"]  # Exclude breaks from total
        ])
        slot_summary["Total Working"] = total_working

        summary_data.append(slot_summary)

    # Convert to DataFrame
    summary_df = pd.DataFrame(summary_data)

    # Write to Excel
    summary_df.to_excel(
        writer, sheet_name=f'{current_day}_Summary', index=False)

    # Format the summary sheet
    summary_worksheet = writer.sheets[f'{current_day}_Summary']

    # Auto-adjust column widths
    for column in summary_worksheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = max_length + 2
        summary_worksheet.column_dimensions[column_letter].width = adjusted_width
