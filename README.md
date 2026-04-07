#Automated Dashboard Report (Python + VBA)

This project automates the process of updating and sending dashboard reports using Python and VBA (Excel Macros).
It reduces manual work by automatically picking the latest file, updating the report, refreshing the dashboard, and sending it within a fixed time.

#How It Works

Get Latest File
The Python script checks the Downloads folder and picks the most recent file automatically.
Update Data Sheet
The data from the latest file is copied and pasted into the Data sheet of the main Excel report.
Run via Button (VBA)
A button will be added in Excel. When click, it triggers VBA, which calls the Python script.
Refresh Dashboard
After updating data, the Excel dashboard is refreshed automatically.
Scheduled Run
Once started, the automation keeps running until 12:00.
Auto Send Report
The updated report is generated and sent within 1 hour.

#Feature

-Fully automated reporting
-Python + VBA integration
-One-click execution
-Automatic dashboard refresh
-Runs on schedule
-Timely report delivery

