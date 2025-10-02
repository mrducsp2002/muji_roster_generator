from datetime import datetime, timedelta
from helper import timespan_to_slot, read_from_excel
from roster_printer import (
    print_roster_header,
    print_coverage_summary,
    print_cs_coverage_totals,
    print_employee_schedule,
    export_roster_to_excel 
)
import random
from collections import defaultdict

tasks = ["FR", "GR", "R"]  # Fitting Room, Greeter, Register
store_hours = {
    "M": ("09:30", "18:00"),
    "T": ("09:30", "18:00"),
    "W": ("09:30", "18:00"),
    "Th": ("09:30", "21:00"),
    "F": ("09:30", "18:00"),
    "Sa": ("10:00", "19:00"),
    "Su": ("10:00", "19:00")
}

# Modify this later to read from Excel
register_coverage = {
    "M": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2,2,2],
    "T": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2,2,2],
    "W": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2,2,2],
    "Th": [1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,4,4,4,4,4,4,4,4,3,3,3,3,2,2,2,2, 2,2,2,2],
    "F": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2,2,2],
    "Su": [1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,4,4,4,4,3,3,3,3,3,3,3,3,2,2,2,2, 2,2,2,2],
    "Sa": [1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,4,4,4,4,3,3,3,3,3,3,3,3,2,2,2,2, 2,2,2,2]
}

def generate_roster(current_day):
    """Generate roster for the given day"""
    working_employees = read_from_excel(current_day)

    # Get store and employee working slots
    store_opening_slots = timespan_to_slot(store_hours[current_day])

    working_employee_slots = {
        employee: timespan_to_slot(
            working_employees[employee]["shift"])
        for employee in working_employees
    }

    working_employee_departments = {
        employee: info["department"] for employee, info in working_employees.items()
    }

    print_roster_header(current_day, store_hours)
    print(f"Available employees: {len(working_employees)}")
    print(f"Store operating slots: {len(store_opening_slots)}")

    # Initialize roster structure

    roster = {}
    for slot in store_opening_slots:
        roster[slot] = {
            "FR": [],    # Fitting Room
            "GR": [],    # Greeter
            "R": [],     # Register
            "40": [],  # Break slot
            "HH": [],
            "L": [],
            "M": [],
            "H&B": [],
            "Stat": [],
            "Acc": [],
            "Hurdle": [],
            "10": [],
        }

    # Track task assignments for fairness
    employee_task_counts = {emp: {"FR": 0, "GR": 0, "R": 0}
                            for emp in working_employees}

    # Generate roster
    roster = fill_roster(roster, current_day, working_employee_slots,
                         working_employee_departments, working_employees)

    # Print formatted output
    print_employee_schedule(roster, current_day, working_employees)
    print_cs_coverage_totals(roster, current_day, store_hours)
    print_coverage_summary(roster, current_day, store_hours)

    return roster


def fill_roster(roster, current_day, working_employee_slots, working_employee_departments, working_employees):
    """Assign tasks to employees, iterate every 15 minutes until the end of the day"""
    customer_service_tasks = ["FR", "GR", "R"]

    # Track which employees has done task today
    employee_CS_task_done_tracker = {
        emp: {"FR": False, "GR": False, "R": False} for emp in working_employee_slots}

    morning_shift_employees = [
        emp for emp in working_employee_slots if working_employee_slots[emp][0] <= 40]
    afternoon_shift_employees = [
        emp for emp in working_employee_slots if 40 < working_employee_slots[emp][0] < 50]
    if current_day == "Th":
        late_shift_employees = [
            emp for emp in working_employee_slots if working_employee_slots[emp][0] >= 50]
        # 19:00 - 20:30 - 22:00
        roster = assign_breaks(roster, late_shift_employees, 68, 71, 74)

    # Assign breaks first for all employees
    roster = assign_breaks(roster, morning_shift_employees,
                           49, 52, 55)  # 12:15 - 13:00 - 13:45
    roster = assign_breaks(roster, afternoon_shift_employees,
                           56, 59, 62)  # 15:00 - 16:30 - 18:00
    
    CS_employees = [emp for emp in working_employee_slots if working_employee_departments[emp] == "CS"]
    print(f"CS Employees: {CS_employees}")
    for emp in CS_employees:
        print(working_employee_slots[emp])
        roster = assign_CS_employee_blocks(
            roster, emp, working_employee_slots[emp], customer_service_tasks, block_size=3)

    for idx, slot in enumerate(roster):
        # Assing CS employee first
                
        # Employees available for the slot (exclude those already assigned to CS tasks, breaks, or hurdles)
        employees_available_for_slot = []
        for emp in working_employee_slots:
            if emp not in CS_employees and slot in working_employee_slots[emp]:
                # Check if employee is already assigned to any task in this slot
                already_assigned = False
                for task_or_dept in roster[slot]:
                    if emp in roster[slot][task_or_dept]:
                        already_assigned = True
                        break

                if not already_assigned:
                    employees_available_for_slot.append(emp)

        # Assign hurdle task if it's the starting slot of any employee working hours
        for emp in working_employee_slots:
            if working_employee_slots[emp] and working_employee_slots[emp][0] == slot:
                roster[slot]["Hurdle"].append(emp)
                if emp in employees_available_for_slot:
                    employees_available_for_slot.remove(emp)

        # Assign customer service tasks first
        for task in customer_service_tasks:
            if task == "R":
                num_required_before = register_coverage[current_day][idx]
                if slot_covered(roster, slot, task) >= num_required_before:
                    continue
                else:
                    num_required = num_required_before - \
                        slot_covered(roster, slot, task)
                    print(f"Task R at slot {slot} needs {num_required} more")
            else:
                num_required = 1
                num_required_before = 1
                if slot_covered(roster, slot, task) >= num_required:
                    continue

            candidates = find_employees_for_task(
                roster, slot, task, employees_available_for_slot, employee_CS_task_done_tracker)
            if not candidates: 
                for emp in employee_CS_task_done_tracker:
                    employee_CS_task_done_tracker[emp][task] = False
                candidates = find_employees_for_task(
                    roster, slot, task, employees_available_for_slot, employee_CS_task_done_tracker)
            selected_employee = []

            if candidates:
                if task == "R":
                    print(
                        f"Task R bf: {candidates}, {len(candidates)}, {num_required}")
                    if len(candidates) < num_required:
                        # Reset tracker and recalc
                        for emp in employee_CS_task_done_tracker:
                            employee_CS_task_done_tracker[emp][task] = False
                        print(employees_available_for_slot)
                        candidates = find_employees_for_task(
                            roster, slot, task, employees_available_for_slot, employee_CS_task_done_tracker)
                        print(
                            f"Task R af: {candidates}, {len(candidates)}, {num_required}")
                    try:
                        selected_employee = random.sample(
                            candidates, k=num_required)
                    except ValueError:
                        selected_employee = candidates[:] + random.sample(roster[slot-1][task], k = num_required - len(candidates))
                        print(
                            f"⚠️ Not enough candidates for {task} at slot {slot}. Assigned {len(selected_employee)} instead of {num_required}.")
                else:
                    selected_employee = [random.choice(candidates)]
                    
                print(f"Selected for {task} at slot {slot}: {selected_employee}")

            # Apply assignment
            for e in selected_employee:
                block_size = 4
                # Extended block if near end of shift
                if max(roster.keys()) - (slot + block_size) < 3:
                    block_size = max(roster.keys()) - slot + 1
                for i in range(block_size):
                    next_slot = slot + i
                    if next_slot in roster and next_slot in working_employee_slots[e]:
                        assigned_emps = {
                            emp: task_name for task_name, emps in roster[next_slot].items() for emp in emps}
                        if e not in assigned_emps:
                            if len(roster[next_slot][task]) >= num_required_before: 
                                continue
                            roster[next_slot][task].append(e)
                            if e in employees_available_for_slot:
                                employees_available_for_slot.remove(e)
                        else:
                            print(
                                f"⚠️ Employee {e} already assigned to {assigned_emps[e]} at slot {next_slot}")
            
                print(f"Task: {task}: {roster[slot][task]}")

            employee_CS_task_done_tracker = mark_task_done(
                selected_employee, task, employee_CS_task_done_tracker)

        # Assign remaining employees to departments
        for employee in employees_available_for_slot:
            dept = working_employee_departments[employee]
            roster[slot][dept].append(employee)

        print(f"Slot {slot} assignments:")

        assigned_employees = []
        for task_name, emps in roster[slot].items():
            assigned_employees.extend(emps)
        duplicates = [emp for emp in set(
            assigned_employees) if assigned_employees.count(emp) > 1]
        if duplicates:
            print(f"⚠️ Duplicate assignment in slot {slot}: {duplicates}")
            
    for employee in working_employee_slots:
        if working_employee_slots[employee]:
            roster = assign_10_minute_break(roster, employee, working_employee_slots[employee], working_employees)

    return roster

def assign_10_minute_break(roster: dict, employee: str, emp_slots: list, working_employees: dict):
    if emp_slots[0] < 40:  # Afternoon shift
        lunch_break_end = 55
    elif emp_slots[0] < 50:  # Evening shift
        lunch_break_end = 62
    else:  # Late night shift
        lunch_break_end = 71
        
    possible_slots = [slot for slot, tasks in roster.items() if employee in tasks.get(working_employees[employee]['department'], []) and slot > lunch_break_end + 4]
    if possible_slots:
        selected_slot = random.choice(possible_slots)
        roster[selected_slot]["10"].append(employee)
        # Remove employee from other tasks in that slot
        for task in roster[selected_slot]:
            if task != "10" and employee in roster[selected_slot][task]:
                roster[selected_slot][task].remove(employee)
    
    return roster
    

def assign_breaks(roster: dict, current_batch: list, start_slot, mid_slot, end_slot):
    first_half = random.sample(current_batch, len(current_batch)// 2)  
    second_half = [emp for emp in current_batch if emp not in first_half]

    # First group
    for slot in range(start_slot, mid_slot):
        for emp in first_half:
            roster[slot]["40"].append(emp)

    # Second group
    for slot in range(mid_slot, end_slot):
        for emp in second_half:
            roster[slot]["40"].append(emp)
    
    return roster

def find_employees_for_task(roster, slot, task, employees_pool, task_done):
    candidates = []

    for employee in employees_pool:
        # Employee has not done this task today
        if task_done[employee][task]:
            continue
        
        # do_not_assign = False
        # for i in range(2):
        #     slot_to_check = slot + i
        #     if slot_to_check in roster and employee in roster[slot_to_check]["40"]:
        #         do_not_assign = True
        #         break
        # if do_not_assign:
        #     continue

        candidates.append(employee)

    return candidates


def slot_covered(roster, slot, task):
    return len(roster[slot][task])


def assign_CS_employee_blocks(roster, emp, emp_slots, customer_service_tasks, block_size=4):
    emp_slots = [slot for slot in emp_slots if slot in roster]

    def is_slot_free(slot):
        return all(emp not in roster[slot][t] for t in ["40", "10", "Hurdle", "R", "GR", "FR"])

    def get_available_tasks(slot):
        # Return tasks that are currently empty in this slot
        available = []
        for task in customer_service_tasks:
            if task == "R":  # If R can have multiple people, keep this condition
                available.append(task)
            elif len(roster[slot][task]) == 0:  # Task is empty in this slot
                available.append(task)
        return available

    slot_assigned = 0
    task = None

    for i, slot in enumerate(emp_slots):
        if not is_slot_free(slot):
            continue

        # If no task assigned yet, or need to switch task
        if task is None or slot_assigned >= block_size:
            available_tasks = get_available_tasks(slot)
            if available_tasks:
                task = random.choice(available_tasks)
                slot_assigned = 0
            else:
                # No empty tasks available, skip this slot or handle as needed
                continue

        roster[slot][task].append(emp)
        slot_assigned += 1

        # Check if a break is coming up in the next 1 or 2 slots
        break_upcoming = False
        for lookahead in range(1, 4):
            if i + lookahead < len(emp_slots) and (emp in roster[emp_slots[i + lookahead]]["40"] or emp in roster[emp_slots[i + lookahead]]["10"]):
                break_upcoming = True
                break

        if break_upcoming:
            slot_assigned = 0   # Reset the block counter
            continue

        # Assign remaining slots at end of shift
        if len(emp_slots) - i <= 3:
            for remaining_slot in emp_slots[i + 1:]:
                if is_slot_free(remaining_slot):
                    available_tasks = get_available_tasks(remaining_slot)
                    if available_tasks:
                        final_task = task if task in available_tasks else random.choice(
                            available_tasks)
                        roster[remaining_slot][final_task].append(emp)
            break

    return roster

def mark_task_done(employees: list, task, tracker):
    """Mark that an employee has done a specific task today"""
    for emp in employees:
        tracker[emp][task] = True

    # Count how many employees have NOT done this task
    not_done_count = sum(1 for emp in tracker if not tracker[emp][task])
    if not_done_count <= 2 :
        for emp in tracker:
            tracker[emp][task] = False
    return tracker


def main():
    current_day = ""
    while current_day not in store_hours:
        current_day = str(input("Enter the day (M, T, W, Th, F, Sa, Su): "))
        if current_day not in store_hours:
            print("Invalid day. Please enter one of M, T, W, Th, F, Sa, Su.")

    roster = generate_roster(current_day)
    
    working_employees = read_from_excel(current_day)

    # Ask if user wants to export to Excel
    export_choice = input(
        "\nWould you like to export the roster to Excel? (y/n): ").lower().strip()

    if export_choice in ['y', 'yes']:
        filename=None
        exported_file = export_roster_to_excel(
            roster, current_day, working_employees, filename)
        if exported_file:
            print(f"\n📊 Excel file created successfully: {exported_file}")
        else:
            print("\n❌ Excel export failed. Please check the error messages above.")


if __name__ == "__main__":
    main()
