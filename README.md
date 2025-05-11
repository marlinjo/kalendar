# Raspored -> Google Cal
Za brzo prebacivanje rasporeda iz ZZR excelice u Google Calendar

Trenutna verzija ne radi za specijaliste zbog "T/T5j" "V/dež", "x" i sličnih smjena, u slučaju interesa će se riješiti nešto

[Tutorial video](https://www.youtube.com/watch?v=E5Vtaitaquc)


///////////////////////////////////////////////////////////
# Schedule Calendar Converter

## Overview
This tool converts work schedules from Excel files to iCalendar (.ics) format for easy import into calendar applications like Google Calendar. It's specifically designed for medical professionals at the Klinički bolnički centar Zagreb (Zagreb Clinical Hospital Center) to convert their monthly work schedules into calendar events.

## Features
- Automatically extracts schedule information from Excel spreadsheets
- Converts different shift types (morning, afternoon, night shifts, days off) into appropriate calendar events
- Creates properly formatted iCalendar (.ics) files
- Handles different schedule codes with appropriate time ranges:
  - Default schedule (morning shift): 07:30 - 15:00
  - Afternoon shift (ending with "p"): 14:00 - 20:00
  - Night shift (H16): 15:30 - 07:30 (next day)
  - Full-day events: SD, PD, GO, bol, sist, postdipl, klaićeva
- Automatically opens Google Calendar import page after generating the .ics file
- Supports Croatian month naming in file exports

## Requirements
- Python 3.6 or higher
- Required Python packages:
  - tkinter (usually included with Python)
  - openpyxl
- Microsoft Excel or compatible software for viewing schedule spreadsheets
- Web browser (preferably Google Chrome)

## Installation
1. Ensure Python 3.6+ is installed on your system
2. Install required packages:
   ```
   pip install openpyxl
   ```

## Usage
1. Run the script:
   ```
   python schedule_converter.py
   ```
2. When prompted, select the Excel file containing your schedule
3. Enter your surname as it appears in the schedule (this is used to find your row in the Excel file)
4. The script will:
   - Process your schedule data
   - Create an .ics file named "[Your Surname] [Month in Croatian].ics" in the same directory as the Excel file
   - Automatically open Google Calendar's import page in your web browser

5. In Google Calendar:
   - Click "Import" 
   - Select the generated .ics file
   - Select the calendar where you want to import your schedule
   - Click "Import" to complete the process

## Excel File Format Requirements
For the script to work correctly, your Excel file should:
- Have dates in row 1 starting from column B (format: D.M.YYYY)
- Have surnames in column A
- Contain schedule codes in the cells where rows (employees) and columns (dates) intersect

## Schedule Codes
The script recognizes the following schedule codes:
- Regular shift codes (default: 07:30 - 15:00)
- Codes ending with "p" (afternoon: 14:00 - 20:00)
- "H16" (night shift: 15:30 - 07:30 next day)
- Full-day events: "SD", "PD", "GO", "bol", "sist", "postdipl", "klaićeva"

## Troubleshooting
- **"Neispravan unos"**: Ensure your surname is entered exactly as it appears in the Excel file
- **Date parsing errors**: Make sure dates in the Excel file are in the correct format (D.M.YYYY)
- **File not found errors**: Check file paths and permissions

## License
This software is provided as-is with no warranty. Free for personal and professional use by medical staff at KBC Zagreb.

## Author
Ideja KT, kod ugl. ChatGPT, Claude za Readme
