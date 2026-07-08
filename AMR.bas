Attribute VB_Name = "AMR"
Function LastPay(cCode, pTable)
    pTable.FindFirst " CODE = " & MyParn(cCode) & " AND TYPE = '7' "
    If pTable.NoMatch Then GoTo 3000
    Do While Not pTable.NoMatch And "7" = pTable.Type And pTable.code = cCode
        If "7" <> pTable.vtype Or pTable.code <> cCode Then GoTo 3000
        dDate = pTable.vDate
        pTable.MoveNext
        If pTable.EOF Then GoTo 3000
    Loop
   LastPay = dDate
3000
End Function
Function TMove_Client(cType, cCode, dDate1, dDate2, pTable)
    pTable.Seek ">=", cCode, cType, dDate1
    If pTable.NoMatch Then GoTo 1000
    Do While Not pTable.NoMatch And pTable.vDate <= dDate2 And cType = pTable.vtype And pTable.code = cCode
        NRETURN = NRETURN + pTable.sal + pTable.PAY
        pTable.MoveNext
        If pTable.EOF Then GoTo 1000
    Loop
   
1000   TMove_Client = NRETURN
End Function
Function Bal_Client(cCode, pTable)
    pTable.FindFirst " CODE = " & MyParn(cCode)
    NRETURN = 0
    If pTable.NoMatch Then GoTo 2000
    Do While Not pTable.NoMatch And pTable.code = cCode
        NRETURN = NRETURN + TurnValue(pTable.sal, Null, 0) - TurnValue(pTable.PAY, Null, 0)
        pTable.MoveNext
        If pTable.EOF Then GoTo 2000
    Loop
2000  Bal_Client = NRETURN
End Function

Function BalItem2(cItem, pTable)
    pTable.Seek ">=", cItem
    If Not pTable.NoMatch Then
    Do While Not pTable.EOF And pTable.Item = cItem
        BalItem2 = BalItem2 + TurnValue(pTable.In, Null, 0) - TurnValue(pTable.OUT, Null, 0)
        pTable.MoveNext
        If pTable.EOF Then Exit Do
    Loop
    End If
End Function
Function BalSubItem(cItem, cNo, pTable1, pTable2, cPack, nUnique)
BalSubItem = 0
If nUnique <> 0 Then
    pTable1.Seek ">=", cItem, cNo
    If Not pTable1.NoMatch Then
        Do While pTable1.Item = cItem And pTable1.cNo = cNo
            BalSubItem = BalSubItem + TurnValue(pTable1.In, Null, 0) - TurnValue(pTable1.OUT, Null, 0)
            pTable1.MoveNext
            If pTable1.EOF Then Exit Do
        Loop
    End If
    pTable2.Seek "=", cItem, cNo
    If Not pTable2.NoMatch Then
        BalSubItem = TurnValue(BalSubItem, Null, 0) + QuantToUnits(pTable2.f_stock1, pTable2.f_stock2, cPack)
    End If
Else
    pTable1.Seek ">=", cItem
    If Not pTable1.NoMatch Then
        Do While Mid(pTable1.Item, 1, 6) = cItem
        If pTable1.cNo = cNo Then
                BalSubItem = BalSubItem + TurnValue(pTable1.In, Null, 0) - TurnValue(pTable1.OUT, Null, 0)
        End If
        pTable1.MoveNext
        If pTable1.EOF Then Exit Do
        Loop
    End If
    
    pTable2.Seek ">=", cItem
    If Not pTable2.NoMatch Then
        Do While Mid(pTable2.Item, 1, 6) = cItem
            If pTable2.Store = cNo Then
                BalSubItem = TurnValue(BalSubItem, Null, 0) + QuantToUnits(pTable2.f_stock1, pTable2.f_stock2, cPack)
            End If
            pTable2.MoveNext
            If pTable2.EOF Then Exit Do
        Loop
    End If
End If
End Function
Function BalBank(cPass, pTable)
    Dim cBalStr  As String
    BalBank = 0
    pTable.FindFirst " id = " & MyParn(cPass)
    If Not pTable.NoMatch Then
        Do While pTable.bank = cPass
            BalBank = BalBank + TurnValue(pTable.In, Null, 0) - TurnValue(pTable.OUT, Null, 0)
            pTable.MoveNext
            If pTable.EOF Then Exit Do
        Loop
    End If
End Function
Function Bal_Supp(cCode, pTable)
    pTable.Index = "nFILE4_11_3"
    pTable.Seek ">=", cCode
    NRETURN = 0
    If pTable.NoMatch Then GoTo 2000
    Do While Not pTable.NoMatch And pTable.code = cCode
        NRETURN = NRETURN + TurnValue(pTable.PLUS, Null, 0) - TurnValue(pTable.minus, Null, 0)
        pTable.MoveNext
        If pTable.EOF Then GoTo 2000
    Loop
2000  Bal_Supp = NRETURN
pTable.Index = "nFILE4_11_2"
End Function
Function Say2Code(pTable, nFlag, cCode)
Say2Code = " "
pTable.FindFirst " code = " & MyParn(cCode) & " and flag = " & nFlag
If Not pTable.NoMatch Then
    Say2Code = pTable.DESCA
End If
End Function

Function SayCode(pTable, nFlag, cCode)
SayCode = " "
pTable.FindFirst " flag = " & nFlag & " and code = " & MyParn(cCode)
If Not pTable.NoMatch Then
    SayCode = pTable.DESCA
End If
End Function

