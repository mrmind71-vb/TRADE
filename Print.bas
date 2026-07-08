Attribute VB_Name = "mPrintCard"
Public Const cUpMargin = 0
Public Const cLeftMargin = 1
Public Const cCardWidth = 2
Public Const cCardHeight = 3
Public Const cRows = 4
Public Const cBeginRow = 5
Public Const cBeginCol = 6
Public Const cPageWidth = 7
Public Const cCols = 8
Public SettingArray(8)
Function MyMeasure(nValue)
    MyMeasure = nValue * 567
End Function
