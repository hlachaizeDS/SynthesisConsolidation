
def wellNbToStr(wellNb):

    col = ((wellNb - 1) // 8) + 1
    row = ((wellNb - 1) % 8) + 1
    row = chr(ord("A") + row - 1)

    return row + str(col)

def excelColumn(columnNumber):

    columnName=""

    if columnNumber>26:
        columnName=columnName + chr(ord("A")+((columnNumber-1)//26)-1)

    columnName= columnName + chr(ord("A")+((columnNumber-1)%26))
    return columnName

def xlCell(row,column):

    return excelColumn(column) + str(row)