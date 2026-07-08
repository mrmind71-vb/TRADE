Attribute VB_Name = "data"
Declare Function GetComputerNameA Lib "KERNEL32" (ByVal lpBuffer As String, nSize As Long) As Long

Public strCon As String, cUserBox As String, lShowBranch As Boolean, lMainShow As Boolean, sCodeVisaBranch As String
Public conShop As New ADODB.Connection, nMaxDisc2 As Double, conPict As New ADODB.Connection, lMainServer As Boolean
Public pServerIp As String, lServerPict As Boolean, cServerNamePICT As String
Public pServerData As String, cUserStore As String
Public strConfact As String, strConPICT As String, nCountBranch, nCountBranch_fr
Public strConfact2 As String, strConfact3 As String
Public cComPort As String, cDocCopySales As String
Public aUser As Variant
Public strConfact2MO As String
Public strConShop As String
Public strConShop_Fr As String
Public GetCon As New ADODB.Connection
Public sMdfName As String, sCatalog As String, cExpress As String
Function openCon(ByRef pCon As ADODB.Connection, Optional ByVal pString As String = "", Optional pTimeOut As Integer = 10) As String
On Error GoTo myerror
Dim cString As String
If pString = "" Then cString = strCon Else cString = pString
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.ConnectionTimeout = pTimeOut
pCon.Open cString
openCon = "ok"
Exit Function
myerror:
MsgBox Err.Description
openCon = Err.Description
Err.Clear
End Function
Function openConFACT(ByRef pCon As ADODB.Connection) As String
On Error GoTo myerror
Dim cString As String
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.Open strConfact
openConFACT = "ok"
Exit Function
myerror:
openConFACT = Err.Description
'MsgBox Err.Description
Err.Clear
End Function
Function openConPICT(ByRef pCon As ADODB.Connection) As String
On Error GoTo myerror
Dim cString As String
If pCon.State = adStateOpen Then pCon.Close
pCon.CommandTimeout = 20
pCon.CursorLocation = adUseClient
pCon.Open strConPICT
openConPICT = "ok"
lServerPict = True
Exit Function
myerror:
openConPICT = Err.Description
'MsgBox strConfact
'MsgBox Err.Description
Err.Clear
End Function
Function openConFACT2(ByRef pCon As ADODB.Connection) As String
On Error GoTo myerror
Dim cString As String

If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.Open strConfact2
openConFACT2 = "ok"
Exit Function
myerror:
openConFACT2 = Err.Description
Err.Clear
End Function
Function openConFACT3(ByRef pCon As ADODB.Connection) As String
On Error GoTo myerror
Dim cString As String
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.Open strConfact3
openConFACT3 = "ok"
Exit Function
myerror:
openConFACT3 = Err.Description
Err.Clear
End Function
Function closeCon(ByRef pCon As ADODB.Connection) As Boolean
On Error GoTo myerror
If pCon.State = adStateOpen Then pCon.Close
Set pCon = Nothing
closeCon = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Function ReadFile(cFile) As String
Dim TextLine
On Error GoTo myerror
Open cFile For Input As #1   ' Open file.
Do While Not EOF(1)
    Line Input #1, TextLine  ' Read line into variable.
    ReadFile = ReadFile & turn(ReadFile, " ") & TextLine
Loop
Close #1   ' Close file.
Exit Function
myerror:
Err.Clear
ReadFile = ""
End Function
Public Sub ToFileExel(MyGrid)
    Dim irow As Integer
    Dim icol As Integer
    Dim objExcl As Excel.Application
    Dim objWk As Excel.Workbook
    Dim objSht As Excel.Worksheet
    Dim iHead As Integer
    Dim vHead As Variant
    On Error Resume Next
    Set objExcl = New Excel.Application
    objExcl.Application.Visible = False
    Set objWk = objExcl.Workbooks.Add
    Set objSht = objWk.Sheets(1)
    vHead = Split(MyGrid.FormatString, "|")
    For iHead = 1 To UBound(vHead)
        If Len(Trim(vHead(iHead))) > 0 Then objSht.Cells(1, iHead) = vHead(iHead)
    Next
    For irow = 0 To MyGrid.Rows - 1
        For icol = 1 To MyGrid.Cols - 1
            If Not MyGrid.ColHidden(icol) Then
                MyGrid.Row = irow
                MyGrid.Col = icol
                objSht.Cells(irow + 1, icol) = MyGrid.text
            End If
        Next icol
    Next irow
    '''''''''''''''''''''''''''''
    '›Ì Õ«·  —Ìœ Õ›Ÿ Ê—ﬁ… «·√ﬂ”·
    'objWk.SaveAs "c:\Book1.xls"
    'objWk.Close
    '«·”ÿ— «· «·Ì Ì€·ﬁ »—‰«„Ã «·√ﬂ”·
    'objExcl.Quit
    ''''''''''''''''''''''''''''
    objExcl.Application.Visible = True
    Set objSht = Nothing
    Set objWk = Nothing
    Set objExcl = Nothing

End Sub

Function aLastDrive(bFlash As Boolean) As Variant
Dim fs, d, DC, letter
Set fs = CreateObject("Scripting.FileSystemObject")
Set DC = fs.Drives
For Each d In DC
    If (d.DriveType = 2 And Not bFlash) Or d.DriveType = 1 Then
        On Error Resume Next
        aLastDrive = AddFlag(Empty, "LETTER", d.DriveLetter)
        aLastDrive = AddFlag(aLastDrive, "SERIAL", d.SerialNumber)
        aLastDrive = AddFlag(aLastDrive, "TYPE", d.DriveType)
    End If
Next
End Function
Function FixFiles(pDir As String) As Boolean
Dim fs As New FileSystemObject
Dim aRet As Variant, nDelete As Long
On Error Resume Next
aRet = retFArray(pDir, "bak")
nDelete = (UBound(aRet) + 1) - nMaxFiles
For i = 0 To (nDelete)
    fs.DeleteFile pDir & "\" & aRet(i)
Next
Err.Clear
End Function

Function retFArray(pFolder As String, sExt As String) As Variant
Dim fso As New FileSystemObject, FileCount As Long
Dim fNames()
ReDim fNames(0)
If Not fso.FolderExists(pFolder) Then
    retFArray = fNames
    Exit Function
End If
Set fold = fso.GetFolder(pFolder)
For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = "." & sExt And Len(File.Name) > 4 Then
        If IsNumeric(Mid(File.Name, Len(sCatalog) + 2, 8)) Then FileCount = FileCount + 1
    End If
Next


ReDim fNames(FileCount)
cFcount = 0

For Each File In fold.Files
    If LCase(Right(File.Name, 4)) = "." & sExt And Len(File.Name) > 4 Then
        If IsNumeric(Mid(File.Name, Len(sCatalog) + 2, 8)) Then
            cFcount = cFcount + 1
            fNames(cFcount) = LCase(File.Name)
        End If
    End If
Next

For tName = 1 To FileCount
    For nName = (tName + 1) To FileCount
        If StrComp(fNames(tName), fNames(nName), 0) = 1 Then
            buffer = fNames(nName)
            fNames(nName) = fNames(tName)
            fNames(tName) = buffer
        End If
    Next
Next
retFArray = fNames
End Function
Function NewflagBranch6(CTABLE, cField, pBranch, pCon As ADODB.Connection) As String
    NewflagBranch6 = IncRec(GetDesca("SELECT MAX(" & cField & ") From " & CTABLE & " WHERE BRANCH = " & MyParn(pBranch), pCon))
    If NewflagBranch6 = "" Then NewflagBranch6 = pBranch & "0001"
End Function
Function NewflagBranch(CTABLE, cField, pBranch, pCon As ADODB.Connection) As String
    NewflagBranch = IncRec(GetDesca("SELECT MAX(" & cField & ") From " & CTABLE & " WHERE BRANCH = " & MyParn(cBranch), pCon))
    If NewflagBranch = "" Then NewflagBranch = cBranch & "000001"
End Function

Function retFormatString(pString As Variant) As String
For i = 0 To UBound(pString)
    retFormatString = retFormatString & turn(retFormatString, "|") & pString(i)
Next
End Function
Function openConShop(ByRef pCon As ADODB.Connection, Optional ByVal pString As String = "", Optional ByVal lMsg As Boolean = True) As String
On Error GoTo myerror
Dim cString As String
If pString = "" Then cString = strConShop Else cString = pString
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.Open cString
openConShop = "ok"
Exit Function
myerror:
'If lMsg Then MsgBox cString
openConShop = Err.Description
Err.Clear
End Function
Function LoadConStringshop()
Dim cServerName As String, cUserId As String, cPassword As String

If lServerOnLineShop Then
    cUserId = decrypt("062F8C407C77", "dr")
    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
Else
'    cUserId = decrypt(RetSetting("userId", App.Path & "\conf.txt"), "dr")
'    cPassword = decrypt(RetSetting("Password", App.Path & "\conf.txt"), "dr")
    cUserId = "DATA_PRO"
    cPassword = "2015"
End If

LoadConStringshop = "provider=SQLOLEDB;data source=" & pServerIp & ";initial " _
            & "catalog=" & pServerData & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"

End Function
Function Newflag_PurchBr(CTABLE, cField, pstore, pCon As ADODB.Connection) As String
Dim loctable As New ADODB.Recordset
'If pcon Is Nothing Then
'    loctable.Open "Select Max(" & cField & ") as Maxof From " & CTABLE & " WHERE STORE = " & MyParn(pstore), GetCon, adOpenStatic, adLockReadOnly, adCmdText
'Else
    loctable.Open "Select Max(" & cField & ") as Maxof From " & CTABLE & " WHERE STORE = " & MyParn(pstore), pCon, adOpenStatic, adLockReadOnly, adCmdText
'End If
If Not (loctable.EOF And loctable.BOF) Then Newflag_PurchBr = IncRec(loctable!maxOf & "")
If Newflag_PurchBr = "" Then Newflag_PurchBr = pstore & "000001"
loctable.Close
Set loctable = Nothing
End Function
Public Sub ToFileExel2(MyGrid, Optional aIg As Variant, Optional nRowHead As Long = 0, Optional aRowMerge As Variant = Empty, Optional aCol As Variant = Empty, Optional nRate As Double = 0, Optional aWidth As Variant = Empty, Optional arowHeight As Variant = Empty, Optional aSetUp As Variant = Empty, Optional nSize As Integer = 12, Optional acolSplit As Variant = Empty, Optional myform As Form)
    Dim irow As Long, i As Long, i2 As Long, NCOLS As Long, nFixedCols As Long, nFixedRows As Long, n As Long
    Dim icol As Long
    Dim objExcl As Excel.Application
    Dim objWk As Excel.Workbook
    Dim objSht As Excel.Worksheet
    Dim iHead As Long
    Dim vHead As Variant
    On Error Resume Next
    Set objExcl = Excel.Application
    objExcl.Application.Visible = False
    Set objWk = objExcl.Workbooks.Add
    Set objSht = objWk.Sheets(1)
    objExcl.Application.DisplayAlerts = False
    Dim NROWS As Long
        
    objSht.PageSetup.TopMargin = 10
    objSht.PageSetup.LeftMargin = 10
    objSht.PageSetup.HeaderMargin = 20
    objSht.PageSetup.CenterHeader = "&B &14"
    
        
    For i = 0 To MyGrid.FixedRows - 1
        If Not MyGrid.RowHidden(i) Then
            nFixedRows = nFixedRows + 1
        End If
    Next
                        
    For i = 0 To MyGrid.FixedCols - 1
        If Not MyGrid.ColHidden(i) Then
            nFixedCols = nFixedCols + 1
        End If
    Next
            
    
    For icol = 0 To MyGrid.Cols - 1
        If Not MyGrid.ColHidden(icol) Then
            NCOLS = NCOLS + 1
            If nFixedRows > 0 Then objSht.Range(objSht.Cells(1, NCOLS), objSht.Cells(nFixedRows, NCOLS)).NumberFormat = "@"
            If MyGrid.Rows > nFixedRows Then
                If Not (MyGrid.ColDataType(icol) = flexDTDouble) Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, NCOLS), objSht.Cells(MyGrid.Rows, NCOLS)).NumberFormat = "@"
                Else
                    objSht.Range(objSht.Cells(nFixedRows + 1, NCOLS), objSht.Cells(MyGrid.Rows, NCOLS)).NumberFormat = ""
                End If
            End If
        End If
    Next icol
    
    If Not myform Is Nothing Then
        myform.prog1.Visible = True
        myform.prog1.Value = 0
    End If
    
    For irow = 0 To MyGrid.Rows - 1
        If (Not myform Is Nothing) And MyGrid.Rows > 1 Then
            myform.prog1.Value = IIf((irow / (MyGrid.Rows - 1)) * 100 > 100, 100, (irow / (MyGrid.Rows - 1)) * 100)
        End If
        If Not MyGrid.RowHidden(irow) Then
            NROWS = NROWS + 1
            NCOLS = 0
            For icol = 0 To MyGrid.Cols - 1
                If Not MyGrid.ColHidden(icol) Then
                    NCOLS = NCOLS + 1
                    If MyGrid.ColDataType(icol) = flexDTDate Then
                        objSht.Cells(NROWS, NCOLS) = myFormat_p(MyGrid.Cell(flexcpTextDisplay, irow, icol))
                    Else
                        objSht.Cells(NROWS, NCOLS) = MyGrid.Cell(flexcpTextDisplay, irow, icol)
                    End If
                End If
            Next icol
        End If
    Next irow
                                
    Dim nRow2 As Long
    If Not IsEmpty(aCol) Then
        For nCol = 0 To UBound(aCol)
            nValue = 0
            For nRow2 = 1 To NROWS
                If Trim(objSht.Cells(nRow2, aCol(nCol))) <> Trim(cValue & "") Then
                    If nValue <> 0 Then
                        objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
                    End If
                    cValue = Trim(objSht.Cells(nRow2, aCol(nCol)))
                    nValue = 0
                    nBegin = nRow2
                Else
                    nValue = nValue + 1
                End If
            Next
            If nValue <> 0 Then
                objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
            End If
        Next
    End If
  
    If Not IsEmpty(acolSplit) Then
        For i = 0 To UBound(acolSplit)
            objSht.Columns(retFlag(acolSplit(i), "col")).PageBreak = xlPageBreakManual
        Next
    End If
    If nFixedRows > 0 Then
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedRows, NCOLS)).HorizontalAlignment = xlCenter
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedRows, NCOLS)).Interior.ColorIndex = 40
    End If
    If NROWS > 0 Then
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(NROWS, NCOLS)).VerticalAlignment = xlCenter
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(NROWS, NCOLS)).Borders.ColorIndex = 0
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(NROWS, NCOLS)).Font.Size = nSize
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(NROWS, NCOLS)).Font.Bold = True
    End If
    
    nCol = 0
    For icol = 0 To MyGrid.Cols - 1
        If Not MyGrid.ColHidden(icol) Then
            nCol = nCol + 1
            If NROWS > 0 Then
                If MyGrid.ColFormat(icol) = "(##,##.##" Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).NumberFormat = "_(#,###.00_);[Red](#,###.00);0.00"
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).HorizontalAlignment = xlRight
                ElseIf MyGrid.ColAlignment(icol) = flexAlignLeftCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).HorizontalAlignment = xlLeft
                ElseIf MyGrid.ColAlignment(icol) = flexAlignRightCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).HorizontalAlignment = xlRight
                ElseIf MyGrid.ColAlignment(icol) = flexAlignCenterCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).HorizontalAlignment = xlCenter
                Else
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).HorizontalAlignment = xlLeft
                End If
            End If
            If nRate > 0 Then objSht.Columns(nCol).ColumnWidth = (MyGrid.ColWidth(icol) / 100) * nRate
        End If
    Next icol
               
    If Not IsEmpty(aWidth) Then
        For i = 0 To UBound(aWidth)
            If Val(retFlag(aWidth(i), "width")) = 0 Then
                objSht.Range(objSht.Cells(1, retFlag(aWidth(i), "col")), objSht.Cells(NROWS, retFlag(aWidth(i), "col"))).Columns.AutoFit
            Else
                objSht.Columns(retFlag(aWidth(i), "col")).ColumnWidth = Val(retFlag(aWidth(i), "width")) / 100
            End If
        Next
    End If
    
    If Not IsEmpty(aRowMerge) Then
        For i = 0 To UBound(aRowMerge)
            If Not IsEmpty(retFlag(aRowMerge(i), "cols")) Then
                If Not IsEmpty(retFlag(aRowMerge(i), "text")) Then
                    objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))) = retFlag(aRowMerge(i), "text")
                End If
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))).Merge
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).Interior.ColorIndex = 19
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).Borders.ColorIndex = 0
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).Font.Bold = True
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).Font.Size = nSize
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).HorizontalAlignment = xlCenter
            End If
            
            If retFlag(aRowMerge(i), "split") Then
                objSht.Rows(retFlag(aRowMerge(i), "row") + 1).PageBreak = xlPageBreakManual
            End If
            
            If retFlag(aRowMerge(i), "word_wrap") Then
                'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).WrapText = True
                objSht.Rows(retFlag(aRowMerge(i), "row") + 1).WrapText = True
            End If
            If Not IsEmpty(retFlag(aRowMerge(i), "height")) Then
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).RowHeight = retFlag(arowHeight(i), "height")
               objSht.Rows(retFlag(aRowMerge(i), "row") + 1).RowHeight = retFlag(aRowMerge(i), "height")
            End If
             If Not IsEmpty(retFlag(aRowMerge(i), "back_color")) Then
               objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).Interior.ColorIndex = retFlag(aRowMerge(i), "back_color")
            End If
            
            If retFlag(aRowMerge(i), "bold") Then
               objSht.Rows(retFlag(aRowMerge(i), "row") + 1).Font.Bold = True
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Bold = True
            End If
        Next
    End If
    
    objSht.PageSetup.Orientation = xlLandscape
    If Not IsEmpty(aSetUp) Then
        If Not IsEmpty(retFlag(aSetUp, "title_col")) Then
            objSht.PageSetup.PrintTitleColumns = objSht.Columns(retFlag(aSetUp, "title_col")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "title_row")) Then
            objSht.PageSetup.PrintTitleRows = objSht.Rows(retFlag(aSetUp, "title_row")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "freeze")) Then
            ActiveWindow.SplitColumn = retFlag(aSetUp, "freeze")
            ActiveWindow.SplitRow = 0
            ActiveWindow.FreezePanes = True
        End If
        
        If retFlag(aSetUp, "autoSize") Then
            objSht.PageSetup.FitToPagesWide = True
        End If
                
        If retFlag(aSetUp, "center_header") <> "" Then
            objSht.PageSetup.CenterHeader = "&14&B" & retFlag(aSetUp, "center_header")
        End If
    End If
    
    
    If Not IsMissing(aIg) Then
        For i = 0 To UBound(aIg)
            objSht.Rows(aIg(i)).Hidden = True
        Next
    End If
    
    
    If Not IsEmpty(pHeader) Then
        For i = 0 To UBound(pHeader)
            If Trim(pHeader(i)) <> "" Then
                objSht.Range("A1", "S1").Insert
                objSht.Range("A1", Chr(64 + NCOLS) & "1").Merge
                objSht.Range("A1", Chr(64 + NCOLS) & "1").Font.Size = nSize + 1
                objSht.Range("A1", Chr(64 + NCOLS) & "1").Font.Bold = True
                objSht.Range("A1", Chr(64 + NCOLS) & "1").VerticalAlignment = xlCenter
                objSht.Range("A1", Chr(64 + NCOLS) & "1").HorizontalAlignment = xlCenter
            End If
        Next
        
        For i = 0 To UBound(pHeader)
            If Trim(pHeader(i)) <> "" Then
                n = n + 1
                objSht.Cells(n, 1) = pHeader(i)
            End If
        Next
    
    End If
    If Not myform Is Nothing Then
        myform.prog1.Visible = True
        myform.prog1.Value = 0
    End If
    objExcl.Application.Visible = True
'    If Not IsEmpty(aRowMerge) Then
'        For i = 0 To UBound(aRowMerge)
'            objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols") + 1)).Merge
'        Next
'    End If

'''''''''''''''''''''''''''''
'›Ì Õ«·  —Ìœ Õ›Ÿ Ê—ﬁ… «·√ﬂ”·
'objWk.SaveAs "c:\Book1.xls"
'objWk.Close
'«·”ÿ— «· «·Ì Ì€·ﬁ »—‰«„Ã «·√ﬂ”·
'objExcl.Quit
''''''''''''''''''''''''''''
Set objSht = Nothing
Set objWk = Nothing
Set objExcl = Nothing
If Not myform Is Nothing Then myform.prog1.Visible = False
End Sub
'Function myFormat_p(sdate As Variant, Optional bLong As Boolean = False) As String
'myFormat_p = Format(sdate & "", "YYYY/M/D") & IIf(bLong, Format(sdate & "", " HH:NN"), "")
'End Function

Function myFormat(sDate As Variant) As String
    myFormat = Format(sDate, "YYYY-MM-DD")
End Function
Function myFormat_sp(sDate As Variant) As Variant
myFormat_sp = TurnValue(myFormat(sDate))
End Function

Function openCon_F(ByRef pCon As ADODB.Connection, Optional ByVal pString As String = "") As String
On Error GoTo myerror
Dim cString As String
If pString = "" Then cString = strCon Else cString = pString
If lMainShow Then cString = LoadConString_J
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.Open cString
openCon_F = "ok"
Exit Function
myerror:
MsgBox cString
openCon_F = Err.Description
Err.Clear
End Function

Public Function GetComputerName() As String
Dim sResult As String * 255
    GetComputerNameA sResult, 255
    GetComputerName = Left$(sResult, InStr(sResult, Chr$(0)) - 1)
End Function
Public Function LastDayOfMonth(dInput As Date) As Integer
    Dim iWeekDay As Integer
    Select Case Month(dInput)
        Case 4, 6, 9, 11            '30-day months.
            LastDayOfMonth = 30
        Case 1, 3, 5, 7, 8, 10, 12  '31-day months.
            LastDayOfMonth = 31
        Case 2                      'February
            LastDayOfMonth = 28
            On Error GoTo LastDayOfMonth_Done
            iWeekDay = Weekday("02/29/" & str$(Year(dInput)))
            LastDayOfMonth = 29
        Case Else                   'Bad date.
            LastDayOfMonth = 0
    End Select
LastDayOfMonth_Done:
    On Error GoTo 0
End Function
Public Sub ToFileExelNew(MyGrid, Optional aIg As Variant, Optional nRowHead As Long = 0, Optional aRowMerge As Variant = Empty, Optional aCol As Variant = Empty, Optional nRate As Double = 0, Optional aWidth As Variant = Empty, Optional arowHeight As Variant = Empty, Optional aSetUp As Variant = Empty, Optional nSize As Integer = 12, Optional acolSplit As Variant = Empty, Optional myform As Form, Optional pHeader As Variant = Empty)
    Dim irow As Long, i As Long, i2 As Long, NCOLS As Long, nFixedCols As Long, nFixedRows As Long, n As Long
    Dim icol As Long
    Dim objExcl As Excel.Application
    Dim objWk As Excel.Workbook
    Dim objSht As Excel.Worksheet
    Dim iHead As Long
    Dim vHead As Variant
    On Error Resume Next
    Set objExcl = Excel.Application
    objExcl.Application.Visible = False
    Set objWk = objExcl.Workbooks.Add
    Set objSht = objWk.Sheets(1)
    objExcl.Application.DisplayAlerts = False
    Dim NROWS As Long
        
    objSht.PageSetup.TopMargin = 10
    objSht.PageSetup.LeftMargin = 10
    objSht.PageSetup.HeaderMargin = 20
    objSht.PageSetup.CenterHeader = "&B &14"
    
        
    For i = 0 To MyGrid.FixedRows - 1
        If Not MyGrid.RowHidden(i) Then
            nFixedRows = nFixedRows + 1
        End If
    Next
                        
    For i = 0 To MyGrid.FixedCols - 1
        If Not MyGrid.ColHidden(i) Then
            nFixedCols = nFixedCols + 1
        End If
    Next
            
    
    For icol = 0 To MyGrid.Cols - 1
        If Not MyGrid.ColHidden(icol) Then
            NCOLS = NCOLS + 1
            If nFixedRows > 0 Then
                objSht.Range(objSht.Cells(1, NCOLS), objSht.Cells(nFixedRows, NCOLS)).NumberFormat = "@"
            End If
            If MyGrid.Rows > nFixedRows Then
                If Not (MyGrid.ColDataType(icol) = flexDTDouble) Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, NCOLS), objSht.Cells(MyGrid.Rows, NCOLS)).NumberFormat = "@"
                Else
                    objSht.Range(objSht.Cells(nFixedRows + 1, NCOLS), objSht.Cells(MyGrid.Rows, NCOLS)).NumberFormat = ""
                End If
            End If
        End If
    Next icol
    
    If Not myform Is Nothing Then
        myform.prog1.Visible = True
        myform.prog1.Value = 0
    End If
    
    For irow = 0 To MyGrid.Rows - 1
        If (Not myform Is Nothing) And MyGrid.Rows > 1 Then
            myform.prog1.Value = IIf((irow / (MyGrid.Rows - 1)) * 100 > 100, 100, (irow / (MyGrid.Rows - 1)) * 100)
        End If
        If Not MyGrid.RowHidden(irow) Then
            NROWS = NROWS + 1
            NCOLS = 0
            
            For icol = 0 To MyGrid.Cols - 1
                If Not MyGrid.ColHidden(icol) Then
                    NCOLS = NCOLS + 1
                    If MyGrid.ColDataType(icol) = flexDTDate Then
                        objSht.Cells(NROWS, NCOLS) = myFormat_p(MyGrid.Cell(flexcpTextDisplay, irow, icol))
                    Else
                        objSht.Cells(NROWS, NCOLS) = MyGrid.Cell(flexcpTextDisplay, irow, icol)
                    End If
                End If
            Next icol
        End If
    Next irow
                                
    Dim nRow2 As Long
    If Not IsEmpty(aCol) Then
        For nCol = 0 To UBound(aCol)
            nValue = 0
            For nRow2 = 1 To NROWS
                If Trim(objSht.Cells(nRow2, aCol(nCol))) <> Trim(cValue & "") Then
                    If nValue <> 0 Then
                        objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
                    End If
                    cValue = Trim(objSht.Cells(nRow2, aCol(nCol)))
                    nValue = 0
                    nBegin = nRow2
                Else
                    nValue = nValue + 1
                End If
            Next
            If nValue <> 0 Then
                objSht.Range(objSht.Cells(nBegin, aCol(nCol)), objSht.Cells(nBegin + nValue, aCol(nCol))).Merge
            End If
        Next
    End If
  
    If Not IsEmpty(acolSplit) Then
        For i = 0 To UBound(acolSplit)
            objSht.Columns(retFlag(acolSplit(i), "col")).PageBreak = xlPageBreakManual
        Next
    End If
    If nFixedRows > 0 Then
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedRows, NCOLS)).HorizontalAlignment = xlCenter
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(nFixedRows, NCOLS)).Interior.ColorIndex = 40
    End If
    If NROWS > 0 Then
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(NROWS, NCOLS)).VerticalAlignment = xlCenter
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(NROWS, NCOLS)).Borders.ColorIndex = 0
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(NROWS, NCOLS)).Font.Size = nSize
        objSht.Range(objSht.Cells(1, 1), objSht.Cells(NROWS, NCOLS)).Font.Bold = True
    End If
    
    nCol = 0
    For icol = 0 To MyGrid.Cols - 1
        If Not MyGrid.ColHidden(icol) Then
            nCol = nCol + 1
            If NROWS > 0 Then
                If MyGrid.ColFormat(icol) = "(##,##.##" Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).NumberFormat = "_(#,###.00_);[Red](#,###.00);0.00"
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).HorizontalAlignment = xlRight
                ElseIf MyGrid.ColAlignment(icol) = flexAlignLeftCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).HorizontalAlignment = xlLeft
                ElseIf MyGrid.ColAlignment(icol) = flexAlignRightCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).HorizontalAlignment = xlRight
                ElseIf MyGrid.ColAlignment(icol) = flexAlignCenterCenter Then
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).HorizontalAlignment = xlCenter
                Else
                    objSht.Range(objSht.Cells(nFixedRows + 1, nCol), objSht.Cells(NROWS, NCOLS)).HorizontalAlignment = xlLeft
                End If
            End If
            If nRate > 0 Then objSht.Columns(nCol).ColumnWidth = (MyGrid.ColWidth(icol) / 100) * nRate
        End If
    Next icol
               
    If Not IsEmpty(aWidth) Then
        For i = 0 To UBound(aWidth)
            If Val(retFlag(aWidth(i), "width")) = 0 Then
                objSht.Range(objSht.Cells(1, retFlag(aWidth(i), "col")), objSht.Cells(NROWS, retFlag(aWidth(i), "col"))).Columns.AutoFit
            Else
                objSht.Columns(retFlag(aWidth(i), "col")).ColumnWidth = Val(retFlag(aWidth(i), "width")) / 100
            End If
        Next
    End If
    
    If Not IsEmpty(aRowMerge) Then
        For i = 0 To UBound(aRowMerge)
            If Not IsEmpty(retFlag(aRowMerge(i), "cols")) Then
                If Not IsEmpty(retFlag(aRowMerge(i), "text")) Then
                    objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))) = retFlag(aRowMerge(i), "text")
                End If
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols"))).Merge
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).Interior.ColorIndex = 19
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).Borders.ColorIndex = 0
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).Font.Bold = True
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).Font.Size = nSize
                objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).HorizontalAlignment = xlCenter
            End If
            
            If retFlag(aRowMerge(i), "split") Then
                objSht.Rows(retFlag(aRowMerge(i), "row") + 1).PageBreak = xlPageBreakManual
            End If
            
            If retFlag(aRowMerge(i), "word_wrap") Then
                'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).WrapText = True
                objSht.Rows(retFlag(aRowMerge(i), "row") + 1).WrapText = True
            End If
            If Not IsEmpty(retFlag(aRowMerge(i), "height")) Then
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).RowHeight = retFlag(arowHeight(i), "height")
               objSht.Rows(retFlag(aRowMerge(i), "row") + 1).RowHeight = retFlag(aRowMerge(i), "height")
            End If
             If Not IsEmpty(retFlag(aRowMerge(i), "back_color")) Then
               objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, NCOLS)).Interior.ColorIndex = retFlag(aRowMerge(i), "back_color")
            End If
            
            If retFlag(aRowMerge(i), "bold") Then
               objSht.Rows(retFlag(aRowMerge(i), "row") + 1).Font.Bold = True
               'objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, nCols)).Font.Bold = True
            End If
        Next
    End If
    
    objSht.PageSetup.Orientation = xlLandscape
    If Not IsEmpty(aSetUp) Then
        If Not IsEmpty(retFlag(aSetUp, "title_col")) Then
            objSht.PageSetup.PrintTitleColumns = objSht.Columns(retFlag(aSetUp, "title_col")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "title_row")) Then
            objSht.PageSetup.PrintTitleRows = objSht.Rows(retFlag(aSetUp, "title_row")).Address
        End If
        
        If Not IsEmpty(retFlag(aSetUp, "freeze")) Then
            ActiveWindow.SplitColumn = retFlag(aSetUp, "freeze")
            ActiveWindow.SplitRow = 0
            ActiveWindow.FreezePanes = True
        End If
        
        If retFlag(aSetUp, "autoSize") Then
            objSht.PageSetup.FitToPagesWide = True
        End If
                
        If retFlag(aSetUp, "center_header") <> "" Then
            objSht.PageSetup.CenterHeader = "&14&B" & retFlag(aSetUp, "center_header")
        End If
    End If
    
    
    If Not IsMissing(aIg) Then
        For i = 0 To UBound(aIg)
            objSht.Rows(aIg(i)).Hidden = True
        Next
    End If
    
    
    If Not IsEmpty(pHeader) Then
        For i = 0 To UBound(pHeader)
            If Trim(pHeader(i)) <> "" Then
                objSht.Range("A1", "Z1").Insert
                objSht.Range("A1", Chr(64 + NCOLS) & "1").Merge
                objSht.Range("A1", Chr(64 + NCOLS) & "1").Font.Size = nSize + 1
                objSht.Range("A1", Chr(64 + NCOLS) & "1").Font.Bold = True
                objSht.Range("A1", Chr(64 + NCOLS) & "1").VerticalAlignment = xlCenter
                objSht.Range("A1", Chr(64 + NCOLS) & "1").HorizontalAlignment = xlCenter
            End If
        Next

        For i = 0 To UBound(pHeader)
            If Trim(pHeader(i)) <> "" Then
                n = n + 1
                objSht.Cells(n, 1) = pHeader(i)
            End If
        Next
    End If
    If Not myform Is Nothing Then
        myform.prog1.Visible = True
        myform.prog1.Value = 0
    End If
    objExcl.Application.Visible = True
'    If Not IsEmpty(aRowMerge) Then
'        For i = 0 To UBound(aRowMerge)
'            objSht.Range(objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + 1), objSht.Cells(retFlag(aRowMerge(i), "row") + 1, retFlag(aRowMerge(i), "col") + retFlag(aRowMerge(i), "cols") + 1)).Merge
'        Next
'    End If

'''''''''''''''''''''''''''''
'›Ì Õ«·  —Ìœ Õ›Ÿ Ê—ﬁ… «·√ﬂ”·
'objWk.SaveAs "c:\Book1.xls"
'objWk.Close
'«·”ÿ— «· «·Ì Ì€·ﬁ »—‰«„Ã «·√ﬂ”·
'objExcl.Quit
''''''''''''''''''''''''''''
Set objSht = Nothing
Set objWk = Nothing
Set objExcl = Nothing
If Not myform Is Nothing Then myform.prog1.Visible = False
End Sub

