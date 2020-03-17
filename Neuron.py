# Importing Excel file
import sys
import xlrd
path = "audit_risk.xlsx"


inputWorkbook = xlrd.open_workbook(path)
inputSheet = inputWorkbook.sheet_by_index(0)


attribute = []  # for keep Attribute
IDXfromAttr = []  # for keep Index Attribute

for colhead in range(inputSheet.ncols):  # import to list only pre-process attribute
    if(colhead == 0 or colhead == 4 or colhead == 4 or colhead == 5 or colhead == 8 or colhead == 11 or colhead == 12 or colhead == 15 or colhead == 18 or colhead == 19 or colhead == 20 or colhead == 21 or colhead == 24 or colhead == 25 or colhead == 28 or colhead == 29 or colhead == 30 or colhead == 31 or colhead == 33 or colhead == 34 or colhead == 36 or colhead == 37 or colhead == 40 or colhead == 41 or colhead == 42 or colhead == 45 or colhead == 46):
        IDXfromAttr.append(colhead)
        attribute.append(inputSheet.cell_value(0, colhead))

x0 = []
x1 = []
x2 = []
x3 = []
x4 = []
x5 = []
x6 = []
x7 = []
x8 = []
x9 = []
x10 = []
x11 = []
x12 = []
x13 = []
x14 = []
x15 = []
x16 = []
x17 = []
x18 = []
x19 = []
x20 = []
x21 = []
x22 = []
x23 = []
x24 = []
xResult = []


for idx, IDXValue in enumerate(IDXfromAttr):
    for colrow in range(inputSheet.nrows):
        if colrow != 0 and colrow <= 620:

            if(idx == 0):
                x0.append(
                    inputSheet.cell_value(colrow, IDXValue))
            if(idx == 1):
                x1.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 2):
                x2.append(
                    inputSheet.cell_value(colrow, IDXValue))
            if(idx == 3):
                x3.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 4):
                x4.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 5):
                x5.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 6):
                x6.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 7):
                x7.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 8):
                x8.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 9):
                x9.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 10):
                x10.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 11):
                x11.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 12):
                x12.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 13):
                x13.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 14):
                x14.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 15):
                x15.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 16):
                x16.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 17):
                x17.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 18):
                x18.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 19):
                x19.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 20):
                x20.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 21):
                x21.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 22):
                x22.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 23):
                x23.append(inputSheet.cell_value(colrow, IDXValue))
            if(idx == 24):
                x24.append(round(inputSheet.cell_value(colrow, IDXValue), 3))
            if(idx == 25):
                xResult.append(inputSheet.cell_value(colrow, IDXValue))

# print('******************************',
#       attribute[0], '******************************')
# print(x0)
# print(len(x0))

# print('******************************',
#       attribute[1], '******************************')
# print(x1)
# print(len(x1))

# print('******************************',
#       attribute[2], '******************************')
# print(x2)
# print(len(x2))

# print('******************************',
#       attribute[3], '******************************')
# print(x3)
# print(len(x3))

# print('******************************',
#       attribute[4], '******************************')
# print(x4)
# print(len(x4))

# print('******************************',
#       attribute[5], '******************************')
# print(x5)
# print(len(x5))

# print('******************************',
#       attribute[6], '******************************')
# print(x6)
# print(len(x6))


# print('******************************',
#       attribute[7], '******************************')
# print(x7)
# print(len(x7))


# print('******************************',
#       attribute[8], '******************************')
# print(x8)
# print(len(x8))


# print('******************************',
#       attribute[9], '******************************')
# print(x9)
# print(len(x9))


# print('******************************',
#       attribute[10], '******************************')
# print(x10)
# print(len(x10))

# print('******************************',
#       attribute[11], '******************************')
# print(x11)
# print(len(x11))

# print('******************************',
#       attribute[12], '******************************')
# print(x12)
# print(len(x12))

# print('******************************',
#       attribute[13], '******************************')
# print(x13)
# print(len(x13))

# print('******************************',
#       attribute[14], '******************************')
# print(x14)
# print(len(x14))

# print('******************************',
#       attribute[15], '******************************')
# print(x15)
# print(len(x15))

# print('******************************',
#       attribute[16], '******************************')
# print(x16)
# print(len(x16))

# print('******************************',
#       attribute[17], '******************************')
# print(x17)
# print(len(x17))

# print('******************************',
#       attribute[18], '******************************')
# print(x18)
# print(len(x18))

# print('******************************',
#       attribute[19], '******************************')
# print(x19)
# print(len(x19))

# print('******************************',
#       attribute[20], '******************************')
# print(x20)
# print(len(x20))

# print('******************************',
#       attribute[21], '******************************')
# print(x21)
# print(len(x21))

# print('******************************',
#       attribute[22], '******************************')
# print(x22)
# print(len(x22))

# print('******************************',
#       attribute[23], '******************************')
# print(x23)
# print(len(x23))

# print('******************************',
#       attribute[24], '******************************')
# print(x24)
# print(len(x24))

# print('******************************',
#       attribute[25], '******************************')
# print(xResult)
# print(len(xResult))
