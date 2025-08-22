import pandas as pd
import calendar
from datetime import date, timedelta

def read_config_from_excel(file_path, sheet_name="Employees"):
    """Reads employee data and config data from a single Excel sheet."""
    try:
        # Read the entire sheet to locate both tables
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=0)

        # Find the row that separates the two tables (where 'Keys' is in the first column)
        first_col_name = df.columns[0]
        separator_indices = df[df[first_col_name] == 'Keys'].index.tolist()

        if not separator_indices:
            raise ValueError("Configuration table with header 'Keys' not found in the sheet.")
        
        separator_index = separator_indices[0]

        # Extract the employee data (everything before the separator)
        employees_df = df.iloc[:separator_index].dropna(how='all').reset_index(drop=True)
        
        # Extract the config data (everything from the separator onwards)
        config_df_raw = df.iloc[separator_index:].reset_index(drop=True)
        config_df_raw.columns = config_df_raw.iloc[0] # Set header to 'Keys', 'Values'
        config_df = config_df_raw[1:].set_index('Keys')
        config = config_df['Values'].to_dict()

        # --- Parse Employee Data ---
        employees = employees_df['Employee'].tolist()
        leave_dates = {}
        employee_shifts = {}
        weekend_workers = set()

        for _, row in employees_df.iterrows():
            employee = row.get('Employee')
            if not employee: continue
            
            if pd.notna(row['Leave_Days (comma-separated)']):
                try:
                    leaves = [int(d.strip()) for d in str(row['Leave_Days (comma-separated)']).split(',') if d.strip()]
                    leave_dates[employee] = leaves
                except (ValueError, AttributeError):
                    print(f"Warning: Could not parse leave days for {employee}. Skipping.")

            # Check the new 'Weekend_Shift' column to identify weekend workers
            if pd.notna(row.get('Weekend_Shift')) and str(row.get('Weekend_Shift')).strip() != '':
                weekend_workers.add(employee)

            # Get the assigned shift directly from the 'Assigned_Shift' column
            shift_str = str(row.get('Assigned_Shift', 'Off')).strip().lower()
            shift_name = shift_str.capitalize()
            employee_shifts[employee] = shift_name if shift_name else 'Off'


        # --- Parse Config Data ---
        year = int(config['year'])
        month = int(config['Month'])

        return year, month, employees, leave_dates, employee_shifts, weekend_workers

    except FileNotFoundError:
        print(f"FATAL ERROR: The configuration file was not found at '{file_path}'")
        return None, None, None, None, None, None
    except Exception as e:
        print(f"FATAL ERROR: Could not read the Excel configuration file. Reason: {e}")
        return None, None, None, None, None, None


def create_monthly_roster(year, month, employees, leave_dates, employee_shifts, weekend_workers):
    """Generates a monthly work roster based on employee roles (weekday/weekend)."""
    month_name = calendar.month_name[month]
    _, num_days = calendar.monthrange(year, month)
    dates = [f"{day:02d}-{month_name[:3]}" for day in range(1, num_days + 1)]
    roster = pd.DataFrame(index=employees, columns=dates, data='') # Start with empty strings

    # 1. Apply Paid Leaves first (highest priority)
    for employee, leaves in leave_dates.items():
        for day in leaves:
            if 1 <= day <= num_days:
                roster.loc[employee, f"{day:02d}-{month_name[:3]}"] = 'Paid Leave'

    # 2. Apply fixed schedules (Week-Offs for regular staff, Comp-Offs for weekend staff)
    for day in range(1, num_days + 1):
        date_col = f"{day:02d}-{month_name[:3]}"
        current_date = date(year, month, day)
        day_of_week = current_date.weekday() # Mon=0, Sun=6

        for employee in employees:
            if roster.loc[employee, date_col] == '': # Only fill if no leave is present
                if employee in weekend_workers:
                    # Assign Comp-Offs on alternating Mon/Tue or Thu/Fri
                    week_start = current_date - timedelta(days=day_of_week)
                    week_number = week_start.isocalendar()[1]
                    is_comp_off_day = (week_number % 2 == 0 and day_of_week in [0, 1]) or \
                                      (week_number % 2 != 0 and day_of_week in [3, 4])
                    if is_comp_off_day:
                        roster.loc[employee, date_col] = 'Comp-Off'
                else: # Regular worker
                    if day_of_week >= 5: # Saturday or Sunday
                        roster.loc[employee, date_col] = 'Week-Off'

    # 3. Fill in the remaining empty slots with the employee's assigned shift
    for employee in employees:
        for day in range(1, num_days + 1):
            date_col = f"{day:02d}-{month_name[:3]}"
            if roster.loc[employee, date_col] == '':
                roster.loc[employee, date_col] = employee_shifts.get(employee, 'Off')

    return roster

if __name__ == "__main__":
    # --- CONFIGURE HERE ---
    INPUT_FILE = 'roster_input.xlsx'
    SHEET_NAME = 'Employees' # The name of the sheet with the employee data
    # --- END OF CONFIGURATION ---

    config_data = read_config_from_excel(INPUT_FILE, sheet_name=SHEET_NAME)

    if config_data[0] is not None:
        year, month, employees, leave_dates, employee_shifts, weekend_workers = config_data

        final_roster = create_monthly_roster(year, month, employees, leave_dates, employee_shifts, weekend_workers)
        
        # Save the roster to Excel, grouped by shift
        month_name = calendar.month_name[month]
        output_filename = f'Roster_{month_name}_{year}.xlsx'
        output_sheet_name = 'Shift_Roster'
        
        try:
            with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
                current_row = 0
                # Dynamically get all unique shifts from the config to create groups
                shift_order = sorted([s for s in set(employee_shifts.values()) if s not in ['Off', '']])
                
                for shift_name in shift_order:
                    employees_in_shift = [emp for emp, shift in employee_shifts.items() if shift == shift_name]
                    
                    if not employees_in_shift:
                        continue

                    shift_roster = final_roster.loc[employees_in_shift]

                    # Write a header for the shift group
                    pd.DataFrame([f'{shift_name} Shift Roster']).to_excel(writer, sheet_name=output_sheet_name, startrow=current_row, index=False, header=False)
                    current_row += 2

                    # Write the actual roster data for this shift
                    shift_roster.to_excel(writer, sheet_name=output_sheet_name, startrow=current_row)
                    
                    # Update the row position for the next group, adding extra space
                    current_row += len(shift_roster.index) + 3

            print("\n" + "="*40)
            print("Successfully generated roster!")
            print(f"File saved as: '{output_filename}'")
            print(f"Sheet name: '{output_sheet_name}'")
            print("="*40)
            
            print("\n--- Roster Preview ---")
            print(final_roster)

        except Exception as e:
            print(f"FATAL ERROR: Could not save the roster to Excel. Reason: {e}")
