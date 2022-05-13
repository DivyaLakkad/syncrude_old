from openpyxl import load_workbook
import xlwings as xw
import os
import sys
import win32com.client
import time
import glob


def Remove_password_xlsx(filename, pw_str):
    xcl = win32com.client.gencache.EnsureDispatch('Excel.Application')
    xcl.Visible = False
    xcl.DisplayAlerts = False
    wb = xcl.Workbooks.Open(filename)
    # wb.Visible = False
    # wb.DisplayAlerts = False
    sheet1 = wb.Worksheets("Timesheet On-Site")
    sheet2 = wb.Worksheets("Timesheet Off-Site")
    sheet1.Unprotect(pw_str)
    sheet2.Unprotect(pw_str)

    wb.Unprotect(pw_str)
    # wb.UnprotectSharing(pw_str)

    wb.Save()
    xcl.Quit()

def Set_password_xlsx(filename, pw_str):
    xcl = win32com.client.gencache.EnsureDispatch('Excel.Application')
    xcl.Visible = False
    xcl.DisplayAlerts = False
    wb = xcl.Workbooks.Open(filename)
    sheet1 = wb.Worksheets("Timesheet On-Site")
    sheet2 = wb.Worksheets("Timesheet Off-Site")
    sheet1.Protect(pw_str)
    sheet2.Protect(pw_str)

    wb.Protect(pw_str)
    # wb.ProtectSharing(pw_str)

    wb.Save()
    xcl.Quit()


employee_list = []
ST_hours = []
OT_hours = []
DOT_hours = []

try:
    # Total Hours Details in THD Folder
    list_of_files = glob.glob(r".\THD\*.xlsx")
    path_THD = max(list_of_files, key=os.path.getctime)

    # TimeSheet Template
    list_of_files = glob.glob(r".\TS\*.xlsx")
    path_TS = max(list_of_files, key=os.path.getctime)

    # #Make day folder if it does not exist
    # list_of_files = glob.glob(r".\TS\Day\*.xlsx")
    # path_TS_Days = max(list_of_files, key=os.path.getctime)
    path_TS_Days = r'.\TS\Day'
    if not os.path.exists(path_TS_Days):
        os.makedirs(path_TS_Days)
    #
    # #Make night folder if it does not exist
    # list_of_files = glob.glob(r".\TS\Night\*.xlsx")
    # path_TS_Nights = max(list_of_files, key=os.path.getctime)

    path_TS_Nights = r'.\TS\Night'
    if not os.path.exists(path_TS_Nights):
        os.makedirs(path_TS_Nights)
except:
    pass

Remove_password_xlsx(r'C:\Users\divyal\Desktop\projects\Syncrude\TS\Timesheet Syncrude_2021_Template - Copy.xlsx',
                     "Syncrude")

wb = xw.Book(path_TS)
time.sleep(5)
path_TS_Days = os.path.join(path_TS_Days,"Day.xlsx")
path_TS_Nights = os.path.join(path_TS_Nights,"Nights.Xlsx")
wb.save(path_TS_Days)
wb.save(path_TS_Nights)
wb.close()

Set_password_xlsx(r'C:\Users\divyal\Desktop\projects\Syncrude\TS\Timesheet Syncrude_2021_Template - Copy.xlsx',
                  "Syncrude")

THD = load_workbook(path_THD)
ws_THD = THD.active

for cell in ws_THD['A']:
    if cell.value == "Graveyard Shift":
        day_length = len(employee_list)
    if cell.value == "Name":
        pass
    else:
        employee_full_name = cell.value
        employee_full_name = employee_full_name.strip()
        xx = employee_full_name.split(",")
        Last_Name = xx[0]
        Last_Name = Last_Name.strip()
        First_Name = xx[1]
        First_Name = First_Name.strip()
        First_Last = f"{First_Name} {Last_Name}"
        employee_list.append(First_Last)

for cell in ws_THD['F']:

    if cell.value == "ST":
        pass
    else:
        ST_hours.append(cell.value)

for cell in ws_THD['G']:

    if cell.value == "OT":
        pass
    else:
        OT_hours.append(cell.value)

for cell in ws_THD['H']:

    if cell.value == "DOT":
        pass
    else:
        DOT_hours.append(cell.value)

print(employee_list)
print(ST_hours)
print(OT_hours)
print(DOT_hours)

try:
    x = day_length

except:
    day_length = len(employee_list)

Total_len = len(employee_list)

app = xw.App()

TS_Days = xw.Book(path_TS_Days)
ws_TS_Days = TS_Days.sheets.active

ws_TS_Days.range('A17').value = "Day"
start = 21
for i in range(0, day_length):
    ws_TS_Days.range((start, 1)).value = employee_list[i]
    ws_TS_Days.range((start, 5)).value = ST_hours[i]
    ws_TS_Days.range((start, 6)).value = OT_hours[i]
    ws_TS_Days.range((start, 7)).value = DOT_hours[i]
    start = start + 1

TS_Days.save()
TS_Days.close()
app.quit()

time.sleep(5)

if day_length != Total_len:
    app = xw.App()

    TS_Nights = xw.Book(path_TS_Nights)
    ws_TS_Nights = TS_Nights.sheets.active

    ws_TS_Nights.range('A17').value = "Nights"
    start = 21

    for j in range(day_length + 1, Total_len):
        ws_TS_Nights.range((start, 1)).value = employee_list[j]
        ws_TS_Nights.range((start, 5)).value = ST_hours[j]
        ws_TS_Nights.range((start, 6)).value = OT_hours[j]
        ws_TS_Nights.range((start, 7)).value = DOT_hours[j]
        start = start + 1

    time.sleep(10)
    TS_Nights.save()
    TS_Nights.close()
    app.quit()

else:
    pass