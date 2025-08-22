**Intelligent Roster Automator**
A powerful and user-friendly Python script that automates the creation of monthly employee shift schedules. It eliminates hours of manual work, reduces human error, and handles complex scheduling rules with ease, all configured through a single, simple Excel file.

**Overview**
Manual employee scheduling is a time-consuming, repetitive, and error-prone task. The Intelligent Roster Automator solves this by ingesting a simple Excel file with employee details, 
leave requests, and role assignments, and instantly generating a complete, accurate, and clearly formatted monthly roster.
This tool is perfect for managers in retail, hospitality, healthcare, call centers, or any business with a shift-based workforce.

**Key Features**
Simple Configuration: All settings, employees, and leaves are managed in one user-friendly Excel file. No coding knowledge is required to use it.
Role-Based Scheduling: Intelligently distinguishes between regular weekday staff and dedicated weekend staff, applying different rules to each.
Complex Rule Automation: Automatically handles alternating compensatory off schedules for weekend staff (e.g., Mon/Tue on even weeks, Thu/Fri on odd weeks).
Leave Priority: Correctly prioritizes planned paid leaves over any other scheduled shift or day off.
Error Reduction: Eliminates the risk of manual errors like forgetting leave requests or miscalculating off-days.
Formatted Output: Generates a clean, human-readable Excel roster, neatly grouped by shift for maximum clarity.
Getting Started

**Follow these steps to get the automator up and running.**

**Prerequisites**
Python 3.6 or newer.

**pip (Python's package installer).**

**Installation**
Clone this repository or download the source code.
Open a terminal or command prompt in the project directory.
Install the required Python libraries by running:
bash

**pip install pandas openpyxl**
 
 **Configuration**
The entire scheduling process is controlled by the roster_input.xlsx file.

Create a file named roster_input.xlsx in the same directory as the script.
Inside the file, create a sheet (e.g., named Employees).
Set up the sheet with two tables as shown below, separated by at least one blank row.

Employee Data Table
Employee	Leave_Days (comma-separated)	Assigned_Shift	Weekend_Shift
Mani	3,4	Morning	
Sandeep	1	noon	
Kavi		Morning	Weekend shift
Priya	11	noon	Weekend shift
RC	20,24	Night	Weekend shift

Employee: The name of the employee.
Leave_Days: A comma-separated list of dates the employee is on leave.
Assigned_Shift: The shift the employee works (e.g., Morning, Noon, Night).
Weekend_Shift: Enter any text here (e.g., "Weekend shift") to designate this employee as a weekend worker. Leave it blank for regular staff.
Configuration Table
Keys	Values
year	2025
Month	9
year: The year for which to generate the roster.
Month: The month number (1-12).

**Running the Automator**
With the roster_input.xlsx file configured, simply run the script from your terminal:

bash

**python roster_automator.py**

The script will execute and generate the final roster file in the same directory.

**Sample Output**
The script will produce a file named Roster_{MonthName}_{Year}.xlsx (e.g., Roster_September_2025.xlsx). The output is neatly organized with employees grouped by their assigned shift.

**Contributing**
Contributions are welcome! If you have ideas for new features, find a bug, or want to improve the code, please feel free to:
