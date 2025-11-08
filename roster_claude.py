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
        "M": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2,2,2],
        "T": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2,2,2],
        "W": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2,2,2],
        "Th": [1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,4,4,4,4,4,4,4,4,3,3,3,3,2,2,2,2,2,2,2,2],
        "F": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2,2,2],
        "Su": [1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,4,4,4,4,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2],
        "Sa": [1,1,1,1,2,2,2,2,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,3,2,2,2,2,2,2,2,2]
    }
    
    # Task assignment settings
    DEFAULT_BLOCK_SIZE = 4
    CS_BLOCK_SIZE = 3
    MIN_EMPLOYEES_FOR_TASK_RESET = 2
    
    # Break assignment slots (15-min intervals)
    MORNING_BREAK_SLOTS = (49, 52, 55)  # 12:15 - 13:00 - 13:45
    AFTERNOON_BREAK_SLOTS = (56, 59, 62)  # 15:00 - 16:30 - 18:00
    LATE_BREAK_SLOTS = (68, 71, 74)  # 19:00 - 20:30 - 22:00
    
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
    
    # Assign CS employees to their blocks
    CS_employees = _get_cs_employees(working_employee_slots, working_employee_departments)
    roster = _assign_cs_employee_blocks(roster, CS_employees, working_employee_slots, working_employee_departments)

    # Process each time slot
    for idx, slot in enumerate(roster):
        roster = _process_slot(roster, slot, idx, current_day, working_employee_slots, 
                              working_employee_departments, CS_employees, employee_CS_task_done_tracker)

    # Assign 10-minute breaks
    roster = _assign_10_minute_breaks(roster, working_employee_slots, working_employees)

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


def _get_cs_employees(working_employee_slots: Dict[str, range], 
                     working_employee_departments: Dict[str, str]) -> List[str]:
    """Get list of Customer Service employees"""
    return [
        emp for emp in working_employee_slots 
        if working_employee_departments[emp] == "CS"
    ]


def _assign_cs_employee_blocks(roster: Dict[int, Dict[str, List[str]]], 
                              CS_employees: List[str], 
                              working_employee_slots: Dict[str, range],
                              working_employee_departments: Dict[str, str]) -> Dict[int, Dict[str, List[str]]]:
    """Assign CS employees to their task blocks"""
    print(f"CS Employees: {CS_employees}")
    for emp in CS_employees:
        print(working_employee_slots[emp])
        roster = assign_CS_employee_blocks(
            roster, emp, working_employee_slots[emp], 
            RosterConfig.CUSTOMER_SERVICE_TASKS, 
            block_size=RosterConfig.CS_BLOCK_SIZE,
            working_employee_departments=working_employee_departments
        )
    return roster


def _process_slot(roster: Dict[int, Dict[str, List[str]]], 
                  slot: int, 
                  idx: int, 
                  current_day: str,
                  working_employee_slots: Dict[str, range],
                  working_employee_departments: Dict[str, str],
                  CS_employees: List[str],
                  employee_CS_task_done_tracker: Dict[str, Dict[str, bool]]) -> Dict[int, Dict[str, List[str]]]:
    """Process a single time slot and assign tasks"""
    # Get available employees for this slot
    employees_available = _get_available_employees_for_slot(
        roster, slot, working_employee_slots, CS_employees
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
                                     working_employee_slots: Dict[str, range],
                                     CS_employees: List[str]) -> List[str]:
    """Get employees available for assignment in this slot"""
    available = []
    for emp in working_employee_slots:
        if emp not in CS_employees and slot in working_employee_slots[emp]:
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
        if working_employee_slots[emp] and working_employee_slots[emp][0] == slot:
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
            roster, slot, task, employees_available, task_tracker, num_required, working_employee_departments
        )
        
        if selected_employees:
            roster = _apply_task_assignment(
                roster, slot, task, selected_employees, employees_available, 
                working_employee_slots, num_required_before
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
                              working_employee_departments: Dict[str, str]) -> List[str]:
    """Select employees for a specific task"""
    # Filter by department and role restrictions
    if task == "FR":
        # FR: only M's and L's are eligible, exclude SPV, ASM, SM, ADM
        filtered_pool = [
            emp for emp in employees_available 
            if (normalize_department_key(working_employee_departments.get(emp)) in ("M's", "L's") and
                working_employee_departments.get(emp) not in ("SPV", "ASM", "SM", "ADM"))
        ]
    elif task == "GR":
        # GR: exclude SPV, ASM, SM, ADM
        filtered_pool = [
            emp for emp in employees_available 
            if working_employee_departments.get(emp) not in ("SPV", "ASM", "SM", "ADM")
        ]
    elif task == "R":
        # R: exclude ADM
        filtered_pool = [
            emp for emp in employees_available 
            if working_employee_departments.get(emp) not in ("ADM",)
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
                          num_required_before: int) -> Dict[int, Dict[str, List[str]]]:
    """Apply task assignment to selected employees"""
    for emp in selected_employees:
        block_size = RosterConfig.DEFAULT_BLOCK_SIZE
        # Extended block if near end of shift
        if max(roster.keys()) - (slot + block_size) < 3:
            block_size = max(roster.keys()) - slot + 1

        # Look ahead up to 2 blocks (8 slots) and stop at a 40-min break if present
        lookahead_slots = 8
        break_slot = None
        for i in range(lookahead_slots):
            next_slot = slot + i
            if next_slot in roster and emp in roster[next_slot]["40"]:
                break_slot = next_slot
                break
        if break_slot is not None and break_slot > slot:
            block_size = min(block_size, break_slot - slot)
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
        
        print(f"Task: {task}: {roster[slot][task]}")
    
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


def _assign_10_minute_breaks(roster: Dict[int, Dict[str, List[str]]], 
                            working_employee_slots: Dict[str, range],
                            working_employees: Dict[str, Dict[str, str]]) -> Dict[int, Dict[str, List[str]]]:
    """Assign 10-minute breaks to all employees"""
    for employee in working_employee_slots:
        if working_employee_slots[employee]:
            roster = assign_10_minute_break(roster, employee, working_employee_slots[employee], working_employees)
    return roster

def assign_10_minute_break(roster: Dict[int, Dict[str, List[str]]], 
                          employee: str, 
                          emp_slots: range, 
                          working_employees: Dict[str, Dict[str, str]]) -> Dict[int, Dict[str, List[str]]]:
    """Assign a 10-minute break to an employee
    
    Args:
        roster: Dictionary mapping time slots to task assignments
        employee: Employee name
        emp_slots: Employee's working time slots
        working_employees: Full employee information
        
    Returns:
        Updated roster with 10-minute break assigned
    """
    # Determine lunch break end time based on shift
    if emp_slots[0] < RosterConfig.MORNING_SHIFT_MAX:  # Morning shift
        lunch_break_end = 55
    elif emp_slots[0] < RosterConfig.AFTERNOON_SHIFT_MAX:  # Afternoon shift
        lunch_break_end = 62
    else:  # Late night shift
        lunch_break_end = 71
        
    # Find possible slots for 10-minute break (after lunch break + 4 slots, avoid last block)
    emp_end_slot = max(emp_slots) if emp_slots else 0
    avoid_last_block = emp_end_slot   # Avoid last 4 slots (1 hour)
    possible_slots = [
        slot for slot, tasks in roster.items() 
        if employee in tasks.get(normalize_department_key(working_employees[employee]['department']), []) 
        and slot > lunch_break_end + 3
        and slot < avoid_last_block
    ]
    
    if possible_slots:
        selected_slot = random.choice(possible_slots)
        roster[selected_slot]["10"].append(employee)
        
        # Remove employee from other tasks in that slot
        for task in roster[selected_slot]:
            if task != "10" and employee in roster[selected_slot][task]:
                roster[selected_slot][task].remove(employee)
    
    return roster
    

def assign_breaks(roster: Dict[int, Dict[str, List[str]]], 
                 current_batch: List[str], 
                 start_slot: int, 
                 mid_slot: int, 
                 end_slot: int) -> Dict[int, Dict[str, List[str]]]:
    """Assign 40-minute breaks to a batch of employees
    
    Args:
        roster: Dictionary mapping time slots to task assignments
        current_batch: List of employees to assign breaks to
        start_slot: Start slot for first group breaks
        mid_slot: Start slot for second group breaks
        end_slot: End slot for second group breaks
        
    Returns:
        Updated roster with breaks assigned
    """
    # Split employees into two groups
    first_half = random.sample(current_batch, len(current_batch) // 2)  
    second_half = [emp for emp in current_batch if emp not in first_half]

    # Assign breaks to first group
    for slot in range(start_slot, mid_slot):
        for emp in first_half:
            roster[slot]["40"].append(emp)

    # Assign breaks to second group
    for slot in range(mid_slot, end_slot):
        for emp in second_half:
            roster[slot]["40"].append(emp)
    
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


def assign_CS_employee_blocks(roster: Dict[int, Dict[str, List[str]]], 
                             emp: str, 
                             emp_slots: range, 
                             customer_service_tasks: List[str], 
                             block_size: int = 4,
                             working_employee_departments: Dict[str, str] = None) -> Dict[int, Dict[str, List[str]]]:
    """Assign a CS employee to task blocks throughout their shift
    
    Args:
        roster: Dictionary mapping time slots to task assignments
        emp: Employee name
        emp_slots: Employee's working time slots
        customer_service_tasks: List of CS tasks
        block_size: Size of task blocks
        
    Returns:
        Updated roster with CS employee assignments
    """
    emp_slots = [slot for slot in emp_slots if slot in roster]

    def is_slot_free(slot: int) -> bool:
        """Check if employee is free in this slot"""
        return all(emp not in roster[slot][t] for t in ["40", "10", "H", "R", "GR", "FR"])

    def get_available_tasks(slot: int) -> List[str]:
        """Get tasks that are available in this slot"""
        available = []
        for task in customer_service_tasks:
            # Skip customer service tasks for SPV, ASM, SM, ADM employees
            if working_employee_departments and working_employee_departments.get(emp) in ("SPV", "ASM", "SM", "ADM"):
                if task in ("FR", "GR", "R"):
                    continue
            
            if task == "R":  # Register can have multiple people
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
                # No empty tasks available, skip this slot
                continue

        roster[slot][task].append(emp)
        slot_assigned += 1

        # Check if a break is coming up in the next 1 or 2 slots
        break_upcoming = False
        for lookahead in range(1, 4):
            if (i + lookahead < len(emp_slots) and 
                (emp in roster[emp_slots[i + lookahead]]["40"] or 
                 emp in roster[emp_slots[i + lookahead]]["10"])):
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
                        final_task = (task if task in available_tasks 
                                    else random.choice(available_tasks))
                        roster[remaining_slot][final_task].append(emp)
            break

    return roster


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
