# Attendance Tracker

## Overview
The Attendance Tracker is a Python-based employee attendance system that utilizes QR code scanning to log check-in and check-out times. It features a Tkinter-based user interface and stores attendance records in JSON and Excel formats.
 
## Features
- **QR Code Scanning**: Uses OpenCV to detect and decode QR codes containing employee information.
- **Automated Check-in/Check-out**: Employees scan their unique QR codes to register attendance.
- **Tkinter GUI**: A user-friendly interface for managing attendance records.
- **Real-time Log Display**: Displays daily attendance records with employee names and timestamps.
- **Excel and JSON Data Export**: Automatically saves attendance records and allows manual download in Excel and JSON formats.
- **Daily Data Reset**: Resets the interface daily while preserving accumulated attendance data.
- **Overtime and Worktime Calculation**: Automatically calculates work hours and overtime based on predefined rules.
- **Auto-Save Feature**: Automatically saves attendance records daily at a scheduled time.
- **Sound Notifications**: Beep sound feedback for successful scans.
- **Auto-Save Attendance Data**: Saves attendance records daily at 10:01 PM.
- **Employee Count Display**: Shows the number of workers scanned for the day.
- **Customizable Notification System**: Pop-up notifications for user actions.

## System Requirements
- Python 3.x
- Required Python Libraries:
  ```bash
  pip install opencv-python openpyxl
  ```
- Additional built-in libraries used:
  - `tkinter`
  - `json`
  - `re`
  - `datetime`
  - `time`
  - `os`
  - `winsound`

## Installation
1. Clone the repository or download the source code.
2. Install the required dependencies:
   ```bash
   pip install opencv-python openpyxl
   ```
3. Run the application:
   ```bash
   python attendance_tracker.py
   ```

## Usage
1. **Start Scanner**: Click the "Start Scanner" button to begin scanning QR codes.
2. **Scan QR Code**: Employees scan their QR codes for check-in and check-out.
3. **View Attendance Log**: The scanned records appear in the log panel.
4. **Download Data**:
   - Click "Download Excel" to save attendance records as an Excel file.
   - Click "Download JSON" to save attendance records in JSON format.
5. **Stop Scanner**: Click "Stop Scanner" to turn off the camera.
6. **Automatic Saving**: The system auto-saves attendance data daily.

## Troubleshooting
- If the application does not start, use `A-Tracker no icon.exe` instead.
- Icons for the application can be downloaded from the provided Google Drive link.

## Data Storage
- **JSON Files**: Stores attendance data in a JSON file (`C:\AttendanceFiles\data_<date>.json`).
- **Excel Files**: Generates an Excel file for daily attendance (`C:\AttendanceFiles\data.xlsx`).

## License
This project is licensed under the MIT License.

## Author
**Ian Tolentino**

## Contact
For inquiries, suggestions, or bug reports, please contact: iantolentino0110@gmail.com

