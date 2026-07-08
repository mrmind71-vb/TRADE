Attribute VB_Name = "ACC"
Public nPurchCashAcc As String      '  1
Public nPurchAcc As String          '  2
Public nRPurchCashAcc As String     '  3
Public nRPurchAcc As String         '  4

Public nSalesCashAcc As String      '  11
Public nSalesAcc As String          '  12
Public nRSalesCashAcc As String     '  13
Public nRSalesAcc As String         '  14

Public nCustAcc As String           '  8
Public nSuppAcc As String           '  5

Public nBankAcc As String           '  6
Public nBoxAcc As String            '  7

Public nDiscAcc As String           '  9
Public nPonsAcc As String           '  10

Public nChq1Acc As String            ' 15
Public nChq2Acc As String            ' 17

Public nChargAcc As String          '  19
Public nIncomAcc As String          '  20
Public nStockAcc As String          '  21

Public Sub FillAccCode()
 nPurchCashAcc = GetAcc(1)      '  1
 nPurchAcc = GetAcc(2)          '  2
 
 nRPurchCashAcc = GetAcc(3)     '  3
 nRPurchAcc = GetAcc(4)         '  4

 nSalesCashAcc = GetAcc(11)     '  11
 nSalesAcc = GetAcc(12)         '  12
 nRSalesCashAcc = GetAcc(13)    '  13
 nRSalesAcc = GetAcc(14)        '  14

 nCustAcc = GetAcc(8)           '  8
 nSuppAcc = GetAcc(5)           '  5

 nBankAcc = GetAcc(6)           '  6
 nBoxAcc = GetAcc(7)            '  7

 nDiscAcc = GetAcc(9)           '  9
 nPonsAcc = GetAcc(10)          '  10

 nChq1Acc = GetAcc(15)           '  15
 nChq2Acc = GetAcc(16)           '  15
 
 nChargAcc = GetAcc(17)         '  17
 nIncomAcc = GetAcc(18)         '  18
 
 nStockAcc = GetAcc(21)         '  18

End Sub
Public Function GetAcc(pFlag)
    GetAcc = GetDesca("SELECT ACC FROM ACC0_30 WHERE CODE = " & pFlag)
End Function
Public Function TotCostStock(dDate)
'GetLastCost(!Item, xDate1.Text, !balance, !PACKAGE)
Dim sourcetable As ADODB.Recordset
Dim nTotal As Double
Dim nTCostItem As Double
cString = "SELECT Sum(FILE1_11.[IN] - FILE1_11.[out]) AS Balance, FILE1_10.ITEM,FILE1_10.PACKAGE , FILE1_10.COST " & _
          "FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM "

If IsDate(Date) Then
    cString = cString & turnFound(cString) & " date <= " & datesq(dDate)
End If

cString = cString & " GROUP BY FILE1_10.ITEM,file1_10.package  , FILE1_10.COST"
          
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, GetCon, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        nTCostItem = GetLastCost(!Item, dDate, !balance, !PACKAGE)
        If nTCostItem = 0 And !balance > 0 Then
            nTCostItem = !COST * !balance / !PACKAGE
        End If
        nTotal = nTotal + nTCostItem
                
        .MoveNext
    Loop
End With
TotCostStock = nTotal
End Function
Public Function GetLastCost(cItem, dDate, nBal, nPack)
    Dim nQ As Double
    Dim nTCost As Double
    Dim nRDisc As Double
    Dim nRRate As Double
    Dim ItemMovetable As New ADODB.Recordset
    ItemMovetable.Open "SELECT * FROM FILE1_11 WHERE ([TYPE] = '2' OR [TYPE] = 'z' )  AND ITEM = " & MyParn(cItem) & " AND DATE <= " & datesq(dDate) & " ORDER BY DATE DESC ", GetCon, adOpenStatic, adLockReadOnly, adCmdText
'    Me.Caption = cItem
    nQ = nBal
    With ItemMovetable
        Do While Not .EOF
            nRDisc = (100 - !Discount) / 100
            nRRate = (100 - Val(!Rate & "")) / 100
            If nQ > 0 Then
                If nQ >= !In Then
                    nTCost = nTCost + ((!In * !price * nRDisc * nRRate) / nPack)
                    nQ = nQ - !In
                Else
                    nTCost = nTCost + ((nQ * !price * nRDisc * nRRate) / nPack)
                    nQ = 0
                End If
            Else
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    GetLastCost = nTCost
End Function


