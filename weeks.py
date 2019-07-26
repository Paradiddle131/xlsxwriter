import xlsxwriter

workbook = xlsxwriter.Workbook('weeks.xlsx')
worksheet = workbook.add_worksheet()

wrap = workbook.add_format({'text_wrap': True})
bold = workbook.add_format({'bold': True})
cell_format = workbook.add_format({'text_wrap': False, 'bold': True})
cell_format.set_font_size(14)
cell_format.set_align('center')
cell_format.set_align('vcenter')

row = 0
column = 0
# 33. Hafta (12-18/08.2019) # 25. Hafta (17-23/06.2019) # 1.Hafta (30/12/2019-05/06.2019)
weekNo = 1
day1 = 30
day2 = 5
month = 12
months31 = [1, 3, 5, 7, 8, 10, 12]
months30 = [2, 4, 6, 9, 11]
secondMonth = 0
day1Str, day2Str = [0]*2
while weekNo <= 52:
    if len(list(str(month))) == 1:
        monthStr = '0' + str(month)
    else:
        monthStr = str(month)
    if len(list(str(day1))) == 1:
        day1Str = '0' + str(day1)
    if len(list(str(day2))) == 1:
        day2Str = '0' + str(day2)
    if abs(day2 - day1) > 6:
        if len(list(str(month))) == 1 and month != 9:
            secondMonth = '0' + str(month + 1)
        elif month == 12:
            secondMonth = "01"
        else:
            secondMonth = str(month + 1)
        worksheet.write(row, column, "{}. Hafta ({}/{}-{}/{}.2019)".format(weekNo, day1Str, monthStr, day2Str, secondMonth), cell_format)
    else:
        worksheet.write(row, column, "{}. Hafta ({}-{}/{}.2019)".format(weekNo, day1Str, day2Str, monthStr), cell_format)
    day1 += 7
    day2 += 7
    if any(month == x for x in months31):
        if day1 > 31:
            month += 1
            day1 += 1
        if day2 > 31:
            day2 += 1
        day1 %= 32
        day2 %= 32
    elif any(month == x for x in months30):
        if day1 > 30:
            month += 1
            day1 += 1
        if day2 > 30:
            day2 += 1
        day1 %= 31
        day2 %= 31
    if month > 12:
        month += 1
    month %= 13
    column += 1
    # row += 1
    weekNo += 1
workbook.close()