from openpyxl import *
from datetime import datetime
from dictionary import *
from tkinter import Tk
from tkinter.filedialog import askopenfilename

ts = load_workbook('Weekly_Timesheet.xlsx')
ts1 = ts.active
Tk().withdraw()

now = datetime.now()

#vars
start = 0
timeStart = 0
end = 0
timeEnd = 0
timeWorked = 0
firstWeek = True
col = 'H'
raw = 0

#user query
print("Is this the first timecard of this paycycle? (y/n)")
answer = input()

if 'y' in answer:
    firstWeek = True
elif 'n' in answer:
    print("Please select a file.")
    filename = askopenfilename()
    print(filename)

print("Lets get the date for a MM/DD/YY format.")

while True:
    try:
        month = int(input("Enter month (1-12):"))
        day = int(input("Enter day (1-31):"))
        year = int(input("Enter year (00-99): "))
        weekday = datetime(2000+year, month, day).weekday()
    except ValueError:
        print("Sorry, I don't understand that. Please enter a numeric value.")
        continue
    else:
        break

if year > 99:
    year % 100

while True:
    try:
        start = int(input('Enter start time (9-5): '))
        end = int(input('Enter end time (9-5): '))
    except ValueError:
        print("I'm sorry, I couldn't really understand that. Please enter a valid time between 1 - 12.")
        continue
    else:
        if start > 12:
            timeStart = start
            start = start - 12
        timeWorked = end - start
        break

if weekday == 1:
    row = '11'
    ts1[slot["month"] + row] = month
    ts1[slot["day"] + row] = day
    ts1[slot["year"] + row] = year
    ts1[slot["start"] + row] = timeStart
    ts1[slot["end"] + row] = end
    ts1[slot["total"] + row] = timeWorked

if weekday == 2:
    row = '12'
    ts1[slot["month"] + row] = month
    ts1[slot["day"] + row] = day
    ts1[slot["year"] + row] = year
    ts1[slot["start"] + row] = timeStart
    ts1[slot["end"] + row] = end
    ts1[slot["total"] + row] = timeWorked

if weekday == 3:
    row = '13'
    ts1[slot["month"] + row] = month
    ts1[slot["day"] + row] = day
    ts1[slot["year"] + row] = year
    ts1[slot["start"] + row] = timeStart
    ts1[slot["end"] + row] = end
    ts1[slot["total"] + row] = timeWorked


ts1['B11'] = timeStart
ts.save('Document2.xlsx')