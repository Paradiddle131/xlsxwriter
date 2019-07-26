month = 8
print(list(str(month)))
print(len(list(str(month))))

if len(list(str(month))) == 1:
    monthStr = '0' + str(month)

print(monthStr)