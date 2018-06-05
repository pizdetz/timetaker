from openpyxl import *
from datetime import datetime

ts = load_workbook('Weekly_Timesheet.xlsx')

now = datetime.now()

print('Enter start time (9-5): ')
start = input()
print('Enter end time (9-5): ')
end = input()
ts1 = ts.active


ts1['B11'] = start
ts.save('Document2.xlsx')