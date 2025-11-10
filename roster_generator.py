from datetime import datetime, timedelta
from typing import Dict, List, Tuple, Set, Optional
from collections import defaultdict
import random

from helper import timespan_to_slot, read_from_excel, select_excel_file
from roster_printer import (
    print_roster_header,
    print_coverage_summary,
    print_cs_coverage_totals,
    export_roster_to_excel 
)

# Normalize department keys to match roster structure
def normalize_department_key(department: str) -> str:
    mapping = {
        "M": "M's",
        "L": "L's",
        "Acc": "Acc.",
        "Stat": "Stat.",
    }
    return mapping.get(department, department)

# Configuration Constants
class RosterConfig:
    """Configuration constants for roster generation"""
    
    # Task types
    CUSTOMER_SERVICE_TASKS = ["FR", "GR", "R"]  # Fitting Room, Greeter, Register
    
    # Store operating hours by day
    STORE_HOURS = {
        "M": ("09:30", "18:00"),
        "T": ("09:30", "18:00"),
        "W": ("09:30", "18:00"),
        "Th": ("09:30", "21:00"),
        "F": ("09:30", "18:00"),
        "Sa": ("10:00", "19:00"),
        "Su": ("10:00", "19:00")
    }
    
    # Register coverage requirements by day (slots per 15-min interval)
    REGISTER_COVERAGE = {
        "M": [1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2],
        "T": [1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2],
        "W": [1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2],
        "Th": [1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,1,1,1,1],
        "F": [1,1,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2,2],
        "Su": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2],
        "Sa": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2]
    }
    
    # Task assignment settings
    DEFAULT_BLOCK_SIZE = 4
    CS_BLOCK_SIZE = 4
    MIN_EMPLOYEES_FOR_TASK_RESET = 2
    
    # Break assignment slots (15-min intervals)
    MORNING_BREAK_SLOTS = (48, 51, 52, 55)  
    AFTERNOON_BREAK_SLOTS = (56, 59, 60, 63)  
    LATE_BREAK_SLOTS = (64, 67, 68, 71)  
    # Shift time boundaries (in 15-min slots)
    MORNING_SHIFT_MAX = 40
    AFTERNOON_SHIFT_MIN = 40
    AFTERNOON_SHIFT_MAX = 50
    LATE_SHIFT_MIN = 50

def generate_roster(current_day: str, file_path: str = None) -> Dict[int, Dict[str, List[str]]]:
    """Generate roster for the given day
    
    Args:
        current_day: Day of the week (M, T, W, Th, F, Sa, Su)
        file_path: Path to Excel file (optional, will prompt user if not provided)
        
    Returns:
        Dictionary mapping time slots to task assignments
    """
    working_employees = read_from_excel(current_day, file_path)

    # Get store and employee working slots
    store_opening_slots = timespan_to_slot(RosterConfig.STORE_HOURS[current_day])

    working_employee_slots = {
        employee: timespan_to_slot(working_employees[employee]["shift"])
        for employee in working_employees
    }

    working_employee_departments = {
        employee: info["department"] 
        for employee, info in working_employees.items()
    }

    print_roster_header(current_day, RosterConfig.STORE_HOURS)
    print(f"Available employees: {len(working_employees)}")
    print(f"Store operating slots: {len(store_opening_slots)}")

    # Initialize roster structure
    roster = _initialize_roster_structure(store_opening_slots)

    # Generate roster
    roster = fill_roster(roster, current_day, working_employee_slots,
                         working_employee_departments, working_employees)

    # Print formatted output
    print_cs_coverage_totals(roster, current_day, RosterConfig.STORE_HOURS)
    print_coverage_summary(roster, current_day, RosterConfig.STORE_HOURS)

    return roster


def _initialize_roster_structure(store_opening_slots: range) -> Dict[int, Dict[str, List[str]]]:
    """Initialize the roster data structure with empty task lists
    
    Args:
        store_opening_slots: Range of time slots when store is open
        
    Returns:
        Dictionary with empty task lists for each slot
    """
    roster = {}
    for slot in store_opening_slots:
        roster[slot] = {
            "FR": [],      # Fitting Room
            "GR": [],      # Greeter
            "R": [],       # Register
            "40": [],      # 40-minute break
            "HH": [],      # Home & Hardware
            "L's": [],       # Ladies
            "M's": [],       # Mens
            "H&B": [],     # Health & Beauty
            "Stat.": [],    # Stationary
            "Acc.": [],     # Accessories
            "H": [],       # H
            "10": [],      # 10-minute break
            "F": [],
            "SPV": [],
            "ASM": [],
            "SM": [], 
            "ADM": []
        }
    return roster


def fill_roster(roster: Dict[int, Dict[str, List[str]]], 
                current_day: str, 
                working_employee_slots: Dict[str, range], 
                working_employee_departments: Dict[str, str], 
                working_employees: Dict[str, Dict[str, str]]) -> Dict[int, Dict[str, List[str]]]:
    """Assign tasks to employees, iterate every 15 minutes until the end of the day
    
    Args:
        roster: Dictionary mapping time slots to task assignments
        current_day: Day of the week
        working_employee_slots: Employee working time slots
        working_employee_departments: Employee department assignments
        working_employees: Full employee information
        
    Returns:
        Updated roster with task assignments
    """
    # Track which employees have done each task today
    employee_CS_task_done_tracker = _initialize_task_tracker(working_employee_slots)
    
    # Categorize employees by shift
    shift_employees = _categorize_employees_by_shift(working_employee_slots)
    
    # Assign breaks for all shift groups
    roster = _assign_all_breaks(roster, shift_employees, current_day, working_employees)

    # Process each time slot
    for idx, slot in enumerate(roster):
        roster = _process_slot(roster, slot, idx, current_day, working_employee_slots, 
                              working_employee_departments, employee_CS_task_done_tracker)

    return roster


def _initialize_task_tracker(working_employee_slots: Dict[str, range]) -> Dict[str, Dict[str, bool]]:
    """Initialize task completion tracker for CS employees"""
    return {
        emp: {"FR": False, "GR": False, "R": False} 
        for emp in working_employee_slots
    }


def _categorize_employees_by_shift(working_employee_slots: Dict[str, range]) -> Dict[str, List[str]]:
    """Categorize employees by their shift times"""
    morning_shift = [
        emp for emp in working_employee_slots 
        if working_employee_slots[emp][0] <= RosterConfig.MORNING_SHIFT_MAX
    ]
    afternoon_shift = [
        emp for emp in working_employee_slots 
        if RosterConfig.AFTERNOON_SHIFT_MIN < working_employee_slots[emp][0] < RosterConfig.AFTERNOON_SHIFT_MAX
    ]
    late_shift = [
        emp for emp in working_employee_slots 
        if working_employee_slots[emp][0] >= RosterConfig.LATE_SHIFT_MIN
    ]
    
    return {
        "morning": morning_shift,
        "afternoon": afternoon_shift,
        "late": late_shift
    }


def _assign_all_breaks(roster: Dict[int, Dict[str, List[str]]], 
                      shift_employees: Dict[str, List[str]], 
                      current_day: str,
                      working_employees: Dict[str, Dict[str, str]]) -> Dict[int, Dict[str, List[str]]]:
    """Assign breaks for all shift groups - only 40-min breaks for shifts > 6 hours"""
    # Filter employees who have shifts longer than 6 hours for 40-minute breaks
    long_shift_morning = [
        emp for emp in shift_employees["morning"] 
        if working_employees[emp]["hours"] > 6.0
    ]
    long_shift_afternoon = [
        emp for emp in shift_employees["afternoon"] 
        if working_employees[emp]["hours"] > 6.0
    ]
    long_shift_late = [
        emp for emp in shift_employees["late"] 
        if working_employees[emp]["hours"] > 6.0
    ]
    
    # Assign 40-minute breaks only to long shifts
    if long_shift_morning:
        roster = assign_breaks(roster, long_shift_morning, *RosterConfig.MORNING_BREAK_SLOTS)
    
    if long_shift_afternoon:
        roster = assign_breaks(roster, long_shift_afternoon, *RosterConfig.AFTERNOON_BREAK_SLOTS)
    
    # Late shift breaks (only for Thursday and only for long shifts)
    if current_day == "Th" and long_shift_late:
        roster = assign_breaks(roster, long_shift_late, *RosterConfig.LATE_BREAK_SLOTS)
    
    return roster


def _process_slot(roster: Dict[int, Dict[str, List[str]]], 
                  slot: int, 
                  idx: int, 
                  current_day: str,
                  working_employee_slots: Dict[str, range],
                  working_employee_departments: Dict[str, str],
                  employee_CS_task_done_tracker: Dict[str, Dict[str, bool]]) -> Dict[int, Dict[str, List[str]]]:
    """Process a single time slot and assign tasks"""
    # Get available employees for this slot
    employees_available = _get_available_employees_for_slot(
        roster, slot, working_employee_slots
    )
    
    # Assign hurdle tasks for employees starting their shift
    roster = _assign_hurdle_tasks(roster, slot, working_employee_slots, employees_available)
    
    # Assign customer service tasks
    roster = _assign_customer_service_tasks(
        roster, slot, idx, current_day, employees_available, employee_CS_task_done_tracker, working_employee_slots, working_employee_departments
    )
    
    # Assign remaining employees to their departments
    roster = _assign_department_tasks(roster, slot, employees_available, working_employee_departments)
    
    # Check for duplicate assignments
    _check_duplicate_assignments(roster, slot)
    
    return roster


def _get_available_employees_for_slot(roster: Dict[int, Dict[str, List[str]]], 
                                     slot: int, 
                                     working_employee_slots: Dict[str, range]) -> List[str]:
    """Get employees available for assignment in this slot"""
    available = []
    for emp in working_employee_slots:
        if slot in working_employee_slots[emp]:
            # Check if employee is already assigned to any task in this slot
            already_assigned = any(
                emp in roster[slot][task_or_dept] 
                for task_or_dept in roster[slot]
            )
            if not already_assigned:
                available.append(emp)
    return available


def _assign_hurdle_tasks(roster: Dict[int, Dict[str, List[str]]], 
                        slot: int, 
                        working_employee_slots: Dict[str, range],
                        employees_available: List[str]) -> Dict[int, Dict[str, List[str]]]:
    """Assign hurdle tasks to employees starting their shift"""
    for emp in working_employee_slots:
        if working_employee_slots[emp] and working_employee_slots[emp][0] == slot and slot == 48:
            roster[slot]["H"].append(emp)
            if emp in employees_available:
                employees_available.remove(emp)
    return roster


def _assign_customer_service_tasks(roster: Dict[int, Dict[str, List[str]]], 
                                  slot: int, 
                                  idx: int, 
                                  current_day: str,
                                  employees_available: List[str],
                                  task_tracker: Dict[str, Dict[str, bool]],
                                  working_employee_slots: Dict[str, range],
                                  working_employee_departments: Dict[str, str]) -> Dict[int, Dict[str, List[str]]]:
    """Assign customer service tasks for this slot"""
    for task in RosterConfig.CUSTOMER_SERVICE_TASKS:
        num_required, num_required_before = _get_task_requirements(task, slot, idx, current_day, roster)
        
        if slot_covered(roster, slot, task) >= num_required_before:
            continue
            
        if task == "R" and slot_covered(roster, slot, task) < num_required_before:
            num_required = num_required_before - slot_covered(roster, slot, task)
            print(f"Task R at slot {slot} needs {num_required} more")
        
        selected_employees = _select_employees_for_task(
            roster, slot, task, employees_available, task_tracker, num_required, working_employee_departments, working_employee_slots
        )
        
        if selected_employees:
            roster = _apply_task_assignment(
                roster, slot, task, selected_employees, employees_available, 
                working_employee_slots, num_required_before, current_day
            )
            task_tracker = mark_task_done(selected_employees, task, task_tracker)
    
    return roster


def _get_task_requirements(task: str, slot: int, idx: int, current_day: str, 
                          roster: Dict[int, Dict[str, List[str]]]) -> Tuple[int, int]:
    """Get the number of employees required for a task"""
    if task == "R":
        num_required_before = RosterConfig.REGISTER_COVERAGE[current_day][idx]
        return num_required_before, num_required_before
    else:
        return 1, 1


def _select_employees_for_task(roster: Dict[int, Dict[str, List[str]]], 
                              slot: int, 
                              task: str, 
                              employees_available: List[str],
                              task_tracker: Dict[str, Dict[str, bool]], 
                              num_required: int,
                              working_employee_departments: Dict[str, str],
                              working_employee_slots: Dict[str, range]) -> List[str]:
    """Select employees for a specific task"""
    # Filter by department and role restrictions
    if task == "FR":
        # FR: only M's, L's, Acc. are eligible, exclude SPV, ASM, SM, ADM
        # Also exclude employees whose shift ends within next 2 slots (30 mins)
        filtered_pool = [
            emp for emp in employees_available 
            if (normalize_department_key(working_employee_departments.get(emp)) in ("M's", "L's", "Acc.") and
                working_employee_departments.get(emp) not in ("SPV", "ASM", "SM", "ADM") and
                max(working_employee_slots[emp]) - slot > 2)
        ]
        # If no M's or L's available, allow all departments except management
        if not filtered_pool: 
            filtered_pool = [
                emp for emp in employees_available 
                if (working_employee_departments.get(emp) not in ("SPV", "ASM", "SM", "ADM") and
                    max(working_employee_slots[emp]) - slot > 2)
            ]
            
    elif task == "GR":
        # GR: exclude SPV, ASM, SM, ADM and employees ending shift within 2 slots
        filtered_pool = [
            emp for emp in employees_available 
            if (working_employee_departments.get(emp) not in ("SPV", "ASM", "SM", "ADM") and
                max(working_employee_slots[emp]) - slot > 2)
        ]
    elif task == "R":
        # R: exclude ADM and employees ending shift within 2 slots
        filtered_pool = [
            emp for emp in employees_available 
            if (working_employee_departments.get(emp) not in ("ADM",) and
                max(working_employee_slots[emp]) - slot > 2)
        ]
    else:
        filtered_pool = employees_available

    candidates = find_employees_for_task(roster, slot, task, filtered_pool, task_tracker)
    
    if not candidates:
        # Reset tracker and try again
        for emp in task_tracker:
            task_tracker[emp][task] = False
        candidates = find_employees_for_task(roster, slot, task, filtered_pool, task_tracker)
    
    if not candidates:
        return []
    
    if task == "R":
        print(f"Task R bf: {candidates}, {len(candidates)}, {num_required}")
        if len(candidates) < num_required:
            # Reset tracker and recalculate
            for emp in task_tracker:
                task_tracker[emp][task] = False
            print(employees_available)
            candidates = find_employees_for_task(roster, slot, task, filtered_pool, task_tracker)
            print(f"Task R af: {candidates}, {len(candidates)}, {num_required}")
        
        try:
            selected = random.sample(candidates, k=num_required)
        except ValueError:
            # Not enough candidates, use what we have plus some from previous slot
            selected = candidates[:] + random.sample(
                roster[slot-1][task], 
                k=num_required - len(candidates)
            )
            print(f"⚠️ Not enough candidates for {task} at slot {slot}. "
                  f"Assigned {len(selected)} instead of {num_required}.")
    else:
        selected = [random.choice(candidates)]
    
    print(f"Selected for {task} at slot {slot}: {selected}")
    return selected


def _apply_task_assignment(roster: Dict[int, Dict[str, List[str]]], 
                          slot: int, 
                          task: str, 
                          selected_employees: List[str],
                          employees_available: List[str],
                          working_employee_slots: Dict[str, range],
                          num_required_before: int,
                          current_day: str) -> Dict[int, Dict[str, List[str]]]:
    """Apply task assignment to selected employees"""
    # Get store opening slot
    store_opening_slot = min(roster.keys())
    
    for emp in selected_employees:
        block_size = RosterConfig.DEFAULT_BLOCK_SIZE
        
        # If it's the first slot and store opens at 9:30 AM, use block size of 6 (1.5 hours)
        store_opening_time = RosterConfig.STORE_HOURS[current_day][0]
        if slot == store_opening_slot and store_opening_time.endswith(":30"):
            block_size = 6
        
        # Get employee's shift end and store closing
        emp_end_slot = max(working_employee_slots[emp]) if working_employee_slots[emp] else slot
        store_closing_slot = max(roster.keys())
        
        # Look ahead up to 2 blocks (8 slots) for a 40-min break
        lookahead_slots = 8
        break_slot = None
        for i in range(1, lookahead_slots):
            next_slot = slot + i
            if next_slot in roster and emp in roster[next_slot]["40"]:
                break_slot = next_slot
                break
        
        # Extended block if near end of employee's shift OR store closing OR 40-min break
        if emp_end_slot - (slot + block_size) < 3:
            block_size = emp_end_slot - slot + 1
        elif store_closing_slot - (slot + block_size) < 3:
            block_size = store_closing_slot - slot + 1
        elif break_slot is not None and break_slot > slot:
            # If break found within 3 slots, extend task to reach the break
            block_size = break_slot - slot + 1
        
        if block_size <= 0:
            continue
            
        for i in range(block_size):
            next_slot = slot + i
            if next_slot in roster and next_slot in working_employee_slots[emp]:
                assigned_emps = {
                    emp: task_name for task_name, emps in roster[next_slot].items() for emp in emps
                }
                if emp not in assigned_emps:
                    if len(roster[next_slot][task]) >= num_required_before:
                        continue
                    roster[next_slot][task].append(emp)
                    if emp in employees_available:
                        employees_available.remove(emp)
                else:
                    print(f"⚠️ Employee {emp} already assigned to {assigned_emps[emp]} at slot {next_slot}")

    
    return roster


def _assign_department_tasks(roster: Dict[int, Dict[str, List[str]]], 
                            slot: int, 
                            employees_available: List[str],
                            working_employee_departments: Dict[str, str]) -> Dict[int, Dict[str, List[str]]]:
    """Assign remaining employees to their departments"""
    for employee in employees_available:
        raw_dept = working_employee_departments[employee]
        dept = normalize_department_key(raw_dept)
        target_key = dept if dept in roster[slot] else raw_dept
        if target_key not in roster[slot]:
            continue
        roster[slot][target_key].append(employee)
    return roster


def _check_duplicate_assignments(roster: Dict[int, Dict[str, List[str]]], slot: int) -> None:
    """Check for and report duplicate assignments in a slot"""
    print(f"Slot {slot} assignments:")
    
    assigned_employees = []
    for task_name, emps in roster[slot].items():
        assigned_employees.extend(emps)
    
    duplicates = [emp for emp in set(assigned_employees) 
                 if assigned_employees.count(emp) > 1]
    if duplicates:
        print(f"⚠️ Duplicate assignment in slot {slot}: {duplicates}")


def assign_breaks(roster: Dict[int, Dict[str, List[str]]], 
                 current_batch: List[str], 
                 start_slot: int, 
                 mid_slot_1: int,
                 mid_slot_2: int,
                 end_slot: int) -> Dict[int, Dict[str, List[str]]]:
    """Assign 40-minute breaks and 10-minute breaks to a batch of employees
    
    Args:
        roster: Dictionary mapping time slots to task assignments
        current_batch: List of employees to assign breaks to
        start_slot: Start slot for first group 40-min breaks
        mid_slot_1: End slot for first group 40-min breaks (10-min break follows)
        mid_slot_2: Start slot for second group 40-min breaks
        end_slot: End slot for second group 40-min breaks (10-min break follows)
        
    Returns:
        Updated roster with breaks assigned
    """
    # Split employees into two groups
    first_half = random.sample(current_batch, len(current_batch) // 2)  
    second_half = [emp for emp in current_batch if emp not in first_half]

    # Assign 40-minute breaks to first group
    for slot in range(start_slot, mid_slot_1):
        for emp in first_half:
            roster[slot]["40"].append(emp)
    
    # Assign 10-minute break to first group (right after their 40-min break)
    for emp in first_half:
        if mid_slot_1 in roster:
            roster[mid_slot_1]["10"].append(emp)

    # Assign 40-minute breaks to second group
    for slot in range(mid_slot_2, end_slot):
        for emp in second_half:
            roster[slot]["40"].append(emp)
    
    # Assign 10-minute break to second group (right after their 40-min break)
    for emp in second_half:
        if end_slot in roster:
            roster[end_slot]["10"].append(emp)
    
    return roster


def find_employees_for_task(roster: Dict[int, Dict[str, List[str]]], 
                           slot: int, 
                           task: str, 
                           employees_pool: List[str], 
                           task_done: Dict[str, Dict[str, bool]]) -> List[str]:
    """Find employees available for a specific task
    
    Args:
        roster: Dictionary mapping time slots to task assignments
        slot: Current time slot
        task: Task to assign
        employees_pool: Available employees
        task_done: Task completion tracker
        
    Returns:
        List of candidate employees for the task
    """
    candidates = []

    for employee in employees_pool:
        # Skip if employee has already done this task today
        if task_done[employee][task]:
            continue
        
        # Note: Commented out break checking logic - could be re-enabled if needed
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


def slot_covered(roster: Dict[int, Dict[str, List[str]]], slot: int, task: str) -> int:
    """Get the number of employees currently assigned to a task in a slot
    
    Args:
        roster: Dictionary mapping time slots to task assignments
        slot: Time slot to check
        task: Task to count
        
    Returns:
        Number of employees assigned to the task
    """
    return len(roster[slot][task])


def mark_task_done(employees: List[str], 
                  task: str, 
                  tracker: Dict[str, Dict[str, bool]]) -> Dict[str, Dict[str, bool]]:
    """Mark that employees have done a specific task today
    
    Args:
        employees: List of employee names
        task: Task that was completed
        tracker: Task completion tracker
        
    Returns:
        Updated task tracker
    """
    for emp in employees:
        tracker[emp][task] = True

    # Count how many employees have NOT done this task
    not_done_count = sum(1 for emp in tracker if not tracker[emp][task])
    if not_done_count <= RosterConfig.MIN_EMPLOYEES_FOR_TASK_RESET:
        for emp in tracker:
            tracker[emp][task] = False
    return tracker


def main() -> None:
    """Main function to run the roster generator"""
    # Prompt user to select Excel file
    selected_file = select_excel_file()
    if not selected_file:
        print("No file selected. Exiting.")
        return
    
    current_day = ""
    while current_day not in RosterConfig.STORE_HOURS:
        current_day = str(input("Enter the day (M, T, W, Th, F, Sa, Su): "))
        if current_day not in RosterConfig.STORE_HOURS:
            print("Invalid day. Please enter one of M, T, W, Th, F, Sa, Su.")

    roster = generate_roster(current_day, selected_file)
    working_employees = read_from_excel(current_day, selected_file)

    # Ask if user wants to export to Excel
    export_choice = input(
        "\nWould you like to export the roster to Excel? (y/n): ").lower().strip()

    if export_choice in ['y', 'yes']:
        exported_file = export_roster_to_excel(
            roster, current_day, working_employees, filename=None)
        if exported_file:
            print(f"\n📊 Excel file created successfully: {exported_file}")
        else:
            print("\n❌ Excel export failed. Please check the error messages above.")


if __name__ == "__main__":
    main()
