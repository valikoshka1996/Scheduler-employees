# User Schedule Generator

This script generates a user schedule based on the following logic:

- Each day has 3 shifts.
- Every worker rotates through 3 shifts for 3 consecutive days (Shift 1 - Shift 2 - Shift 3).
- Separate schedules for weekend and weekday shifts.
- You can add custom weekends (in case of holidays).

## How it Works?

You can run the script with the following command:
python import\ calendar\ v0.0.19.py

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
![{4B173C25-F2F5-4C3B-B006-0ACE87709CF8}](https://github.com/user-attachments/assets/59d3a7fc-bd00-49b4-947b-de6714e483fb)



## Example Generation - Input Data:

![{8B4ED9B8-0B1D-44C5-851D-27D2A24E3A63}](https://github.com/user-attachments/assets/9276dfed-b485-4738-ae7c-3807c56472b8)


## Result:

_(Screenshot of the result window)_
