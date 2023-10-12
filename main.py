import pandas as pd

def extract_numeric_hours(value):
    if isinstance(value, str):
        # spliting the string by ':' to separate hours and minutes
        hours, minutes = value.split(':')
        # converting hours and minutes to float and calculating the total hours
        total_hours = float(hours) + float(minutes) / 60
        return total_hours
    return None

def analyze_employee_file(file_path):
    # reading the excel file into a pandas dataframe
    df = pd.read_excel(file_path, engine='openpyxl') # openpyxl is used as the engine because the default engine (xlrd) does not support the newer .xlsx format

    # renaming the columns for easier access
    df = df.rename(columns={
        'Position ID': 'PositionID',
        'Position Status': 'PositionStatus',
        'Time': 'TimeIn',
        'Time Out': 'TimeOut',
        'Timecard Hours (as Time)': 'TimecardHours',
        'Pay Cycle Start Date': 'PayCycleStartDate',
        'Pay Cycle End Date': 'PayCycleEndDate',
        'Employee Name': 'EmployeeName',
        'File Number': 'FileNumber',
    })

    # converting the 'TimeIn' and 'TimeOut' columns to datetime objects
    df['TimeIn'] = pd.to_datetime(df['TimeIn'], format='%Y-%m-%d %H:%M:%S', errors='coerce') # errors='coerce' to handle missing values
    df['TimeOut'] = pd.to_datetime(df['TimeOut'], format='%Y-%m-%d %H:%M:%S', errors='coerce')

    # extracting the numeric hours from the 'TimecardHours' column
    df['TimecardHours'] = df['TimecardHours'].apply(extract_numeric_hours)

    # dropping rows with missing values in the 'TimeIn' and 'TimecardHours' columns
    df = df.dropna(subset=['TimeIn', 'TimecardHours'])

    # sorting the dataframe by 'EmployeeName' and 'TimeIn'
    # df = df.sort_values(by=['EmployeeName', 'TimeIn'])

    # creating a dictionary to store the employee data
    employee_data = {}

    # iterating over the rows of the dataframe
    for i, row in df.iterrows():
        # extracting the employee name, shift date, and shift hours from the row
        employee_name = row['EmployeeName']
        shift_date = row['TimeIn'].date()
        shift_hours = row['TimecardHours']

        # if the employee name is not in the dictionary, add it
        if employee_name not in employee_data:
            employee_data[employee_name] = {}

        # if the shift date is not in the dictionary, add it
        if shift_date not in employee_data[employee_name]:
            employee_data[employee_name][shift_date] = {'total_hours': 0.0, '14HourShift': False}

        # add the shift hours to the total hours
        employee_data[employee_name][shift_date]['total_hours'] += shift_hours

        # if the shift hours are greater than or equal to 14, set the 14HourShift flag to True
        if shift_hours >= 14.0:
            employee_data[employee_name][shift_date]['14HourShift'] = True
    
    employees_1_to_10_hours = {}
    employess_consecutive_7_days = set()
    employees_14_hours_shift = set()

    for employee_name, shifts in employee_data.items():
        worked_days = list(shifts.keys())
        worked_days.sort()

        for day in worked_days:
            total_hours = shifts[day]['total_hours']

            if 1.0 <= total_hours <= 10.0:
                if employee_name not in employees_1_to_10_hours:
                    employees_1_to_10_hours[employee_name] = []

                employees_1_to_10_hours[employee_name].append(day)

            if len(worked_days) >= 7:
                consecutive = True
                for i in range(6, len(worked_days)):
                    if (worked_days[i] - worked_days[i - 6]).days > 6:
                        consecutive = False
                        break
                if consecutive:
                    employess_consecutive_7_days.add(employee_name)

            if shifts[day]['14HourShift']:
                employees_14_hours_shift.add((employee_name, day))

    print("Summary: Employees who worked for 1 to 10 hours")
    for employee, dates in employees_1_to_10_hours.items():
        position_id = df[df['EmployeeName'] == employee].iloc[0]['PositionID']
        date_str = ', '.join(map(str, dates))
        print(f"Name: {employee}, Position ID: {position_id}, Dates: {date_str}")

    print("\nSummary: Employees who worked 7 consecutive days")
    for employee in employess_consecutive_7_days:
        position_id = df[df['EmployeeName'] == employee].iloc[0]['PositionID']
        print(f"Name: {employee}, Position ID: {position_id}")

    print("\nSummary: Employees who worked 14 hours in a day")
    for employee, date in employees_14_hours_shift:
        position_id = df[df['EmployeeName'] == employee].iloc[0]['PositionID']
        print(f"Name: {employee}, Position ID: {position_id}, Date: {date}")

if __name__ == '__main__':
    file_path = 'Assignment_Timecard.xlsx'
    analyze_employee_file(file_path)