# User Schedule Generator

This script generates a user schedule based on the following logic:

- Each day has 3 shifts.
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

![{041EDF9B-6562-4559-B0DD-1A6E71870E18}](https://github.com/user-attachments/assets/795c1bb3-adb5-4c27-9b7e-7c80ada21b12)


## Example Generation - Input Data:

![{22C4F8AA-8B5D-4C5E-85EA-250D2DDF8F28}](https://github.com/user-attachments/assets/e9f3ae26-5443-4dd8-a92c-e3eaff76576c)

## Result:

![{18A62CD1-9DA9-456D-8D51-DCA9E18DBC7D}](https://github.com/user-attachments/assets/4bc9a0f1-fa80-4ee6-bd3c-ed6593558163)
