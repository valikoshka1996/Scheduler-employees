# User Schedule Generator

This script generates a user schedule based on the following logic:

- Each day has 2 or 3 shifts.
- Every worker rotates through 3 shifts for 3 consecutive days (Shift 1 - Shift 2 - Shift 3).
- Separate schedules for weekend and weekday shifts.
- You can add custom weekends (in case of holidays).

## How it Works?

You can run the script with the following command:
`python import\ calendar\ v0.0.19.py`

However, you must have the following libraries installed:

- `calendar`
- `random`
- `os`
- `datetime`
- `openpyxl`
- `tkinter`

Alternatively, you can navigate to the `build/dist` folder and run the compiled `.exe` file.

### If running the program via the script:
- The schedule will be generated next to the script file.

### If running the compiled `.exe` file:
- The schedule will be saved in the `build/dist` folder.

## Main Window of the Program:

<img width="576" height="335" alt="image" src="https://github.com/user-attachments/assets/288fda3d-7320-4f93-9a80-6bf6f148cca1" />


## Example Generation - Input Data:

<img width="573" height="337" alt="image" src="https://github.com/user-attachments/assets/cc13cdd7-cdf1-4e9a-991a-0acfcdb48c74" />

## Result:

<img width="1590" height="277" alt="image" src="https://github.com/user-attachments/assets/009e216f-0549-4f84-a487-29ec92d3ebae" />
