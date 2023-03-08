from datetime import datetime, timedelta
from collections import defaultdict
import xlsxwriter

def schedule(month_year: str,
             employees: list,
             max_hours_per_day: int,
             max_hours_per_week: int,
             max_shifts_per_day: int,
             rest_time_between_shifts: int,
             rest_days_per_week: int,
             employee_availability: dict,
             employee_shift_preference: dict):
    # parse month and year
    month_year = datetime.strptime(month_year, '%m/%Y')
    days_in_month = (month_year.replace(month=month_year.month % 12 + 1) - timedelta(days=1)).day

    # initialize schedule
    schedule = defaultdict(lambda: defaultdict(list))

    # initialize last shift worked
    last_shift = defaultdict(str)
    last_rest_day = defaultdict(lambda: month_year.replace(day=1) - timedelta(days=1))

    # initialize employee availability
    for employee in employees:
        if employee not in employee_availability:
            employee_availability[employee] = []

    # initialize employee shift preference
    for employee in employees:
        if employee not in employee_shift_preference:
            employee_shift_preference[employee] = []

    # initialize hours worked per week
    hours_worked_per_week = defaultdict(int)

    # iterate over days in month
    for day in range(1, days_in_month + 1):
        date = month_year.replace(day=day)
        weekday = date.weekday()

        # reset hours worked per week on Monday
        if weekday == 0:
            hours_worked_per_week = defaultdict(int)

        # check if employees have reached maximum rest days per week
        for employee in employees:
            shifts_worked_this_week = sum(1 for d in range(max(1, day - weekday), day) if schedule[d][employee])
            if shifts_worked_this_week == 7 - rest_days_per_week:
                employee_availability[employee].append(date)

        # assign shifts to employees
        for shift in ['08h-16h', '16h-00h', '00h-08h']:
            required_workers = 2 if shift != '00h-08h' else 1

            available_employees = [e for e in employees if date not in employee_availability[e] and shift not in employee_shift_preference[e]]

            sorted_employees = sorted(available_employees, key=lambda e: hours_worked_per_week[e])

            for i in range(required_workers):
                worker_assigned_to_shift = False

                for worker_index, worker_name in enumerate(sorted_employees):
                    worker_schedule = schedule[day][worker_name]
                    worker_hours_today = sum([8 for shift_time_range_str in worker_schedule])

                    if len(worker_schedule) < max_shifts_per_day and worker_hours_today < max_hours_per_day and \
                            hours_worked_per_week[worker_name] < max_hours_per_week:
                        last_shift_end_time = None

                        if len(worker_schedule) > 0:
                            last_shift_end_time = datetime.strptime(worker_schedule[-1].split("-")[1], "%Hh")

                        current_shift_start_time = datetime.strptime(shift.split("-")[0], "%Hh")

                        enough_rest_between_shifts = True

                        if last_shift_end_time != None and current_shift_start_time < last_shift_end_time:
                            current_shift_start_time += timedelta(days=1)

                        if last_shift_end_time != None and (
                                current_shift_start_time - last_shift_end_time).seconds / 3600 < rest_time_between_shifts:
                            enough_rest_between_shifts = False

                        if enough_rest_between_shifts == True:
                            # check if worker has taken a rest day since their last shift
                            if last_shift[worker_name] != '' and last_shift[worker_name] != shift and date - \
                                    last_rest_day[worker_name] > timedelta(days=1):
                                continue

                            schedule[day][worker_name].append(shift)
                            hours_worked_per_week[worker_name] += 8
                            last_shift[worker_name] = shift
                            worker_assigned_to_shift = True

                        if worker_assigned_to_shift == False:
                            schedule[day]['(External)'].append(shift)

                    sorted_employees.pop(worker_index)

                    if worker_assigned_to_shift == True:
                        break

        # update last rest day for workers that didn't work today
        for employee in employees:
            if not schedule[day][employee]:
                last_rest_day[employee] = date
            else:
                # reset last rest day for workers that worked today
                last_rest_day[employee] = month_year.replace(day=1) - timedelta(days=1)

    return schedule

def export_schedule_to_excel(schedule: dict, filename: str):
    # create workbook and worksheet
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    # write headers
    worksheet.write(0, 0, 'Day')
    for i, shift in enumerate(['08h-16h', '16h-00h', '00h-08h']):
        worksheet.write(0, i + 1, shift)

    # write data
    for row, day in enumerate(sorted(schedule.keys())):
        worksheet.write(row + 1, 0, day)
        for col, shift in enumerate(['08h-16h', '16h-00h', '00h-08h']):
            employees = [e for e in schedule[day] if shift in schedule[day][e]]
            worksheet.write(row + 1, col + 1, ', '.join(employees))

    # close workbook
    workbook.close()

# example usage of function with sample data

month_year='03/2023'
employees=['Alice','Bob','Charlie','Dave','John','Michael','Dilan']
max_hours_per_day=8
max_hours_per_week=40
max_shifts_per_day=1
rest_time_between_shifts=12
rest_days_per_week=1
employee_availability={'Alice':[datetime(2023,3,5),datetime(2023,3,6)],'Bob':[datetime(2023,3,7)]}
employee_shift_preference={'Alice':['16h-00h', '00h-08h']}

result=schedule(month_year,
             employees,
             max_hours_per_day,
             max_hours_per_week,
             max_shifts_per_day,
             rest_time_between_shifts,
             rest_days_per_week,
             employee_availability,
             employee_shift_preference)

for day in result:
    print(f"Day {day}:")
    for employee in result[day]:
        print(f"\t{employee}: {result[day][employee]}")

# generate schedule
result = schedule(month_year,
                  employees,
                  max_hours_per_day,
                  max_hours_per_week,
                  max_shifts_per_day,
                  rest_time_between_shifts,
                  rest_days_per_week,
                  employee_availability,
                  employee_shift_preference)

# export schedule to Excel file
export_schedule_to_excel(result,'schedule.xlsx')