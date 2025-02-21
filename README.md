# Shift Schedule Generator

This Python script generates a shift schedule for employees, considering both weekday and weekend shifts. The schedule follows a 3-shift rotation system and adapts based on the type of day.

## Features
- Supports separate shift rotations for weekday and weekend workers.
- Allows users to specify the first worker for the third shift of the first day.
- If the first day is a weekend, the user also selects the first worker for the first shift of the first weekday.
- Automatically rotates shifts while maintaining the 3-day rule.
- Saves the schedule in an Excel file with color-coded worker assignments.

## Requirements
- Python 3.x
- `openpyxl` library (install with `pip install openpyxl`)

## Usage
1. Run the script:
   ```sh
   python sched.py
