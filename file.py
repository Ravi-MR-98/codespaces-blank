from datetime import datetime, timedelta
from collections import defaultdict
import pandas as pd

# Inputs
month_year = '03/2023'
employees = ['Alice', 'Bob', 'Charlie', 'Dave']
max_hours_per_day = 8
max_hours_per_week = 40
max_shifts_per_day = 1
rest_time_between_shifts = 12
rest_days_per_week = 1
employee_availability = {'Alice': ['03/05/2023', '03/06/2023'], 'Bob': [], 'Charlie': ['03/10/2023'], 'Dave': []}
employee_shift_preference = {'Alice': [0, 1], 'Bob': [0, 2], 'Charlie': [1], 'Dave': [0, 1, 2]}

# Constants
SHIFT_TIMES = [(8, 16), (16, 24), (0, 8)]
HOURS_PER_SHIFT = 8


def generate_schedule(month_year: str,
                      employees: list,
                      max_hours_per_day: int,
                      max_hours_per_week: int,
                      max_shifts_per_day: int,
                      rest_time_between_shifts: int,
                      rest_days_per_week: int,
                      employee_availability: dict,
                      employee_shift_preference: dict):
    # Convert month_year to datetime object and get number of days in month
    month_year_dt = datetime.strptime(month_year + '/01', '%m/%Y/%d')
    days_in_month = (month_year_dt.replace(month=month_year_dt.month % 12 + 1) - timedelta(days=1)).day

    # Initialize schedule and weekly hours for each employee
    schedule = defaultdict(lambda: defaultdict(list))
    weekly_hours = defaultdict(int)

    # Initialize current shift for each employee
    current_shift = defaultdict(lambda: -1)

    # Iterate through each day in month
    for day in range(1, days_in_month + 1):
        date_str = f"{month_year}/{str(day).zfill(2)}"
        date_dt = datetime.strptime(date_str, '%m/%Y/%d')

        # Reset weekly hours if it's a new week
        if date_dt.weekday() == 0:
            weekly_hours = defaultdict(int)

        # Check if any employees need to switch shifts due to rest day
        for employee, hours in weekly_hours.items():
            if hours >= max_hours_per_week - rest_days_per_week * HOURS_PER_SHIFT:
                current_shift[employee] = -1

        # Iterate through each shift
        for shift_index, (start_hour, end_hour) in enumerate(SHIFT_TIMES):
            workers_needed = 2 if shift_index != 2 else 1

            # Assign workers to shift
            for _ in range(workers_needed):
                assigned = False

                # Check if any employees are available and prefer this shift
                for employee in employees:
                    if date_str not in employee_availability.get(employee, []) and (current_shift[employee] == shift_index or current_shift[employee] == -1) and shift_index in employee_shiftpreference.get(employee, [0, 1, 2]) and weekly_hours[employee] + HOURS_PER_SHIFT <= max_hours_ per_worker per week:
                        schedule[date_str][shift_index].append(employee)
                        weekly_hours[employee] += HOURS_PER_SHIFT

                        current_shift[employee] = shift_index

                        assigned = True

                        break

                # If no employees are available or prefer this shift,assign external worker
                if not assigned:
                    schedule[date_str][shift_index].append('(External)')

    return schedule


schedule = generate_schedule(
    month_year=month_year,
    employees=employees,
    max_hours_per_day=max_hours_per_day,
    max_hours_per_week=max_hours_per_week,
    max_shifts_per_day=max_shifts_per_day,
    rest_time_between_shifts=rest_time_between_shifts,
    rest_days_per_week=rest_days_per_week,
    employee_availability=employee_availability,
    employee_shift_preference=employee_shift_preference
)

for day, schedule_for_day in schedule.items():
    print(f"{day}:")

    for shift_index, (start_hour, end_hour) in enumerate(SHIFT_TIMES):
        start_time = f"{str(start_hour).zfill(2)}h"
        end_time = f"{str(end_hour).zfill(2)}h" if end_hour != 24 else "00h"

        print(f" {start_time}-{end_time}: {','.join(schedule_for_day[shift_index])}")

# Create a DataFrame from the schedule
df = pd.DataFrame.from_dict(schedule, orient='index')

# Rename columns
df.columns = [f"{str(start_hour).zfill(2)}h-{str(end_hour).zfill(2)}h" if end_hour != 24 else "00h-00h" for start_hour, end_hour in SHIFT_TIMES]

# Fill NaN values with empty strings
df.fillna('', inplace=True)

# Combine workers in each cell into a single string separated by commas
df = df.applymap(lambda x: ', '.join(x) if isinstance(x, list) else x)

# Export DataFrame to Excel file
df.to_excel('schedule.xlsx')
