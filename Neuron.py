# Importing Excel file
import numpy
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


def signmoidFN(v):
    e = 2.71828
    return round(1/(1+pow(e, -v)), 3)


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
err = 1.0

# import xlsx each cell dataset to list


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

hiddenLayerCount = 1
InputNodeCount = len(IDXfromAttr)-1
weightCount = InputNodeCount*(InputNodeCount-1)

w = []  # keep every weight

print('hidden layer = ', hiddenLayerCount)
print('InputNodeCount = ', InputNodeCount)
print('weightCount = ', weightCount)

for i in range(weightCount*weightCount):
    # Generates Weight to list
    w.append(round(numpy.random.uniform(0.1, 0.5), 1))

bias = -5.0
v = []
trainRow = 620
vi = 0.0
y = []

newW = [0.0]*25

while (err > 0.02):
    
    for rowIDX in range(trainRow):
        for kIDX in range(InputNodeCount-1):
            for jIDX in range(InputNodeCount):
                if(jIDX == 0):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x0[rowIDX], 3)
                if(jIDX == 1):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x1[rowIDX], 3)
                if(jIDX == 2):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x2[rowIDX], 3)
                if(jIDX == 3):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x3[rowIDX], 3)
                if(jIDX == 4):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x4[rowIDX], 3)
                if(jIDX == 5):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x5[rowIDX], 3)
                if(jIDX == 6):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x6[rowIDX], 3)
                if(jIDX == 7):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x7[rowIDX], 3)
                if(jIDX == 8):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x8[rowIDX], 3)
                if(jIDX == 9):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x9[rowIDX], 3)
                if(jIDX == 10):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x10[rowIDX], 3)
                if(jIDX == 11):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x11[rowIDX], 3)
                if(jIDX == 12):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x12[rowIDX], 3)
                if(jIDX == 13):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x13[rowIDX], 3)
                if(jIDX == 14):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x14[rowIDX], 3)
                if(jIDX == 15):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x15[rowIDX], 3)
                if(jIDX == 16):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x16[rowIDX], 3)
                if(jIDX == 17):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x17[rowIDX], 3)
                if(jIDX == 18):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x18[rowIDX], 3)
                if(jIDX == 19):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x19[rowIDX], 3)
                if(jIDX == 20):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x20[rowIDX], 3)
                if(jIDX == 21):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x21[rowIDX], 3)
                if(jIDX == 22):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x22[rowIDX], 3)
                if(jIDX == 23):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x23[rowIDX], 3)
                if(jIDX == 24):
                    vi += round(w[(jIDX*InputNodeCount-1)+kIDX] +
                                newW[kIDX]*x24[rowIDX], 3)

            v.append(round(vi+bias, 3))
            vi = 0.0
        break


    print('v = ', v)
    print(len(v))


    for vItem in v:
        y.append(round(signmoidFN(vItem), 3))
    print('y = ', y)


    outputLayer = []


    for i in range(len(y)):
        # Generates Weight to list
        w.append(round(numpy.random.uniform(0.1, 0.5), 1))

    vi = 0

    for i in range(len(y)):
        vi = y[i]*w[weightCount*weightCount+i]
        vi += bias
        outputLayer.append(signmoidFN(vi))

    print('outputLayer = ', outputLayer)


    err = xResult[0] - outputLayer[0]

    print('err = ', err)


    def grad(err, y):
        return err * y * (1-y)


    newW.clear()

    wii = 0.0
    learnR = -0.925


    for i in range(len(y)):
        wii = (-1*learnR)*grad(err, y[i])*w[weightCount*weightCount+i]
        newW.append(round(wii, 5))
        wii = 0.0

print(newW)
print(len(newW))
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
