Attribute VB_Name = "abd"
Public bEdit As Boolean
Public rdItem As New adodb.Recordset
Public contemp As New adodb.Connection, consec As New adodb.Connection, condef As New adodb.Connection
Public bSupermode As Boolean
Public aBranch As Variant
Public searchArray
Function Newflag(sTable, cField, pCon As adodb.Connection) As Long
Dim loctable As New adodb.Recordset
loctable.Open "Select Max(" & cField & ") as Maxof From " & sTable, pCon, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.EOF And loctable.BOF) Then Newflag = Val(loctable!maxOf & "") + 1
loctable.Close
Set loctable = Nothing
End Function
Function NewflagDoc(sDate As String, sStore As String, pCon As adodb.Connection) As String
Dim loctable As New adodb.Recordset, cString As String
If cBranch <> "00" Then
    If Len(sStore) = 2 Then
        cString = "Select MAX(SUBSTRING(DOC_NO2,3,3)) as MaxofDocNo FROM FILE6_20H "
    Else
        cString = "Select MAX(SUBSTRING(DOC_NO2,4,3)) as MaxofDocNo FROM FILE6_20H "
    End If
    cString = cString & turn(cString) & "[DATE] = " & DateSq(sDate)
    cString = cString & turn(cString) & "[BOX] = " & MyParn(cBranchBox)
    loctable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
    If Len(sStore) = 2 Then sStore = Val(sStore)
    If Len(sStore) = 3 Then sStore = Val(sStore)
    If Not loctable.EOF Then cString = sStore & RetZero(Val(loctable!maxOfDocNo & "") + 1, 3) Else cString = sStore & RetZero(1, 3)
    NewflagDoc = RetZero(Day(sDate), 2) & RetZero(Month(sDate), 2) & RetZero(Format(sDate, "YY"), 2) & cString
    loctable.Close
    Set loctable = Nothing
Else
    NewflagDoc = IncRec(GetDesca("select max(doc_no) from file6_20h where ISINVOICE = 0 AND  branch = '00' ", pCon))
    If NewflagDoc = "" Then NewflagDoc = RetZero("1", 11)
End If
End Function
Function DateSq(ByVal X As Variant, Optional X2 As String = "") As String
If Not IsDate(X) Then
    DateSq = X
    Exit Function
End If
X = Format(X, "dd-mm-yyyy")
If IsDate(X2) Then X2 = Format(TimeValue(X2), "hh:nn:ss")
DateSq = MyParn(Right(X, 4) & "-" & Mid(X, 4, 2) & "-" & Left(X, 2) & turn(X2, Space(1)) & X2)
'x = DateValue(Format(x, "dd-mm-yyyy"))
'DateSq = "Convert(DateTime," & MyParn(Month(x) & "/" & Day(x) & "/" & Year(x)) & ",102)"

End Function
Function DateConv(ByVal X As Variant)
If Not IsDate(X) Then
    DateConv = ""
    Exit Function
End If
X = Format(X, "dd-mm-yyyy")
DateConv = Right(X, 4) & "-" & Mid(X, 4, 2) & "-" & Left(X, 2)
End Function
Function DateFix(dDate) As Variant
If Not IsDate(dDate) Then
    DateFix = Null
    Exit Function
End If
DateFix = DateValue(Format(dDate, "dd-mm-yyyy"))
End Function
Function MyParnAnd(cSearch, cField) As String
Dim aString, cString2
aString = Split(Trim(cSearch), " ")
For i2 = 0 To UBound(aString)
    If Trim(aString(i2)) <> "" Then cString2 = cString2 & IIf(cString2 = "", "", " and ") & cField & " Like " & MyParn("%" & aString(i2) & "%")
Next
MyParnAnd = cString2
End Function
Function aGetDesca(pString, pCon As adodb.Connection) As Variant
Dim loctable As New adodb.Recordset
'If pcon Is Nothing Then
'    loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
'Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
'End If
ReDim aRet(0)
If Not (loctable.BOF And loctable.EOF) Then
    ReDim aRet(loctable.Fields.Count)
    For i = 0 To loctable.Fields.Count - 1
        aRet(i + 1) = loctable.Fields(i).Value
    Next
End If
aGetDesca = aRet
loctable.Close
Set loctable = Nothing
End Function
Function GetDesca(pString, pCon As adodb.Connection) As String
Dim loctable As New adodb.Recordset
loctable.CursorLocation = adUseClient
'If pcon Is Nothing Then
'     loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
'Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
'End If
If Not (loctable.BOF And loctable.EOF) Then GetDesca = loctable(0) & ""
loctable.Close
Set loctable = Nothing
End Function
Function GetBoolean(pString, pCon As adodb.Connection) As Integer
Dim loctable As New adodb.Recordset
loctable.CursorLocation = adUseClient
'If pcon Is Nothing Then
'    loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdTextElse
'Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
'End If

If Not (loctable.BOF And loctable.EOF) Then
    GetBoolean = IIf(loctable(0), 1, 0)
Else
    GetBoolean = -1
End If
loctable.Close
Set loctable = Nothing
End Function
Function addDate(pValue) As String
If Not IsDate(pValue) Then
    addDate = "NULL"
Else
    addDate = DateSq(pValue)
End If
End Function

Function retitem(cItem, cField)
'rdItem.Find "item = " & MyParn(cItem), , adSearchForward, adBookmarkFirst
'If Not rdItem.EOF Then retitem = rdItem(cField)
End Function
Function BetweenString(pValue1 As String, pValue2 As String, Optional pString1 = " „‰ ", Optional pString2 = " Õ Ì ")
If Trim(pValue1) & Trim(pValue2) = "" Then Exit Function
If Trim(pValue2) = "" Then
    BetweenString = pString1 & pValue1
ElseIf Trim(pValue1) = "" Then
    BetweenString = pString2 & pValue2
Else
    BetweenString = pString1 & pValue1 & pString2 & pValue2
End If
End Function

Function retHeader(aHeader, nBegin, pCount, Optional pSep As String = "  ") As String
Dim nFound As Integer, i As Integer, nCount As Integer
For i = 0 To UBound(aHeader)
    If aHeader(i) <> "" Then
        If nFound >= nBegin Then
            retHeader = retHeader + IIf(retHeader = "", "", pSep) & aHeader(i)
            nCount = nCount + 1
            If nCount = pCount Then Exit For
        End If
        nFound = nFound + 1
    End If
Next
End Function


Sub grdMake(pString As String, pFieldBound, pFieldList, pCon As adodb.Connection, Grid1 As VSFlexGrid, Optional pRows As Integer = 10)
Dim rstLocal As adodb.Recordset, cString As String
Set rstLocal = New adodb.Recordset
Grid1.Rows = 0
Grid1.Rows = pRows
rstLocal.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
cString = "#" & ";"
Do Until rstLocal.EOF
    cString = cString & "|#" & rstLocal(pFieldBound) & ";" & rstLocal(pFieldList)
    rstLocal.MoveNext
Loop
rstLocal.Close
Set rstLocal = Nothing
Grid1.ColComboList(0) = cString
End Sub
Function GrdQry(pGrid As VSFlexGrid, pField, Optional isString As Boolean) As String
Dim cString As String
For i = 0 To pGrid.Rows - 1
    If pGrid.TextMatrix(i, 0) <> "" Then cString = cString & IIf(cString = "", "(", " or ") & pField & " = " & IIf(isString, "'", "") & pGrid.TextMatrix(i, 0) & IIf(isString, "'", "")
Next
If cString <> "" Then cString = cString & ")"
GrdQry = Trim(cString)
End Function
Function GrdTitle(pGrid As VSFlexGrid) As String
Dim cString As String
For i = 0 To pGrid.Rows - 1
    If pGrid.TextMatrix(i, 0) <> "" Then cString = cString & IIf(cString = "", "", " - ") & pGrid.Cell(flexcpTextDisplay, i, 0, i, 0)
Next
GrdTitle = cString
End Function
Function retFilter(pTable As adodb.Recordset, pFilter)
Dim aFilter
ReDim aFilter(pTable.Fields.Count - 1)
pTable.Filter = pFilter
If Not (pTable.EOF And pTable.BOF) Then
    For i = 0 To pTable.Fields.Count - 1
        aFilter(i) = pTable.Fields(i).Value
    Next
Else
    For i = 0 To pTable.Fields.Count - 1
        aFilter(i) = Null
    Next
End If
retFilter = aFilter
End Function
Sub MyCreateFolder(pDir)
On Error Resume Next
Dim fs As FileSystemObject, i As Long
Set fs = CreateObject("Scripting.FileSystemObject")
aString = Split(pDir, "\")
cString = aString(0)
For i = 1 To UBound(aString)
    cString = cString & "\" & aString(i)
    If Not fs.FolderExists(cString) Then fs.CreateFolder (cString)
Next
Err.Clear
End Sub
Function LastDrive(Optional bLetter As Boolean = False)
Dim fs, d, DC, letter
Set fs = CreateObject("Scripting.FileSystemObject")
Set DC = fs.Drives
For Each d In DC
    If d.DriveType = 2 Then
        On Error Resume Next
        letter = IIf(bLetter, d.DriveLetter, d.SerialNumber)
    End If
Next
LastDrive = letter
End Function
Function CreateInsert(ByVal aInsert, ByVal CTABLE) As String
Dim cString1 As String, cString2 As String
For i = 0 To UBound(aInsert)
    If aInsert(i, 0) <> "" Then
        cString1 = cString1 & IIf(cString1 = "", "", ",") & aInsert(i, 0)
        cString2 = cString2 & IIf(cString2 = "", "", ",") & aInsert(i, 1)
    End If
Next
CreateInsert = "Insert into " & CTABLE & " (" & _
                cString1 & _
                ")"
CreateInsert = CreateInsert & " values(" & _
                cString2 & _
                ")"
End Function
Function CreateUpdate(ByVal aInsert, ByVal CTABLE, ByVal cCondition, Optional ByVal aIg) As String
Dim bUpDate As Boolean, cString As String
If IsMissing(aIg) Then aIg = Array(0)

For i = 0 To UBound(aIg)
    If aIg(i) >= 0 And aIg(i) < UBound(aInsert) Then
        aInsert(aIg(i), 0) = ""
    End If
Next

For i = 0 To UBound(aInsert)
    If aInsert(i, 0) <> "" Then
        CreateUpdate = CreateUpdate & IIf(CreateUpdate = "", "", ",") & aInsert(i, 0) & _
                       " = " & aInsert(i, 1)
    End If
Next
CreateUpdate = "UPDATE " & CTABLE & " SET " & _
               CreateUpdate & _
               cCondition
End Function
Function DefAdd(sFlag, sFlagDesca, sFlagValue)
On Error Resume Next
condef.Execute "insert into DEFTABLE(Flag,FlagDesca,FlagValue)" & _
               "Values(" & _
               addstring(sFlag) & "," & _
               addstring(sFlagDesca) & "," & _
               addstring(sFlagValue) & _
               ")"
If Err.Number = -2147467259 Then
    Err.Clear
    condef.Execute "update defTable Set " & _
                   " FlagValue  = " & addstring(sFlagValue) & _
                   " where Flag = " & MyParn(sFlag) & _
                   " and FlagDesca = " & MyParn(sFlagDesca)
End If
Err.Clear
End Function
Function DefGet(sFlag, sFlagDesca) As String
Dim loctable As New adodb.Recordset
cString = "Select * From defTable " & _
          " where Flag = " & MyParn(sFlag) & _
          " and FlagDesca = " & MyParn(sFlagDesca)
loctable.Open cString, condef, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.BOF And loctable.EOF) Then DefGet = loctable!FlagVAlue & ""
loctable.Close
Set loctable = Nothing
End Function
Function StrList(cString, pCon As adodb.Connection)
Dim listTable As New adodb.Recordset
listTable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
Do Until listTable.EOF
    StrList = StrList & "|#" & listTable.Fields(0) & ";" & listTable.Fields(1)
    listTable.MoveNext
Loop
End Function
Function StrListA(cString, pCon As adodb.Connection)
Dim listTable As New adodb.Recordset
StrListA = "|#" & "" & ";" & ""
listTable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
Do Until listTable.EOF
    StrListA = StrListA & "|#" & listTable.Fields(0) & ";" & listTable.Fields(1)
    listTable.MoveNext
Loop
End Function
Sub FilterGrd(pGrid, pString As String, Optional pCol As Integer = 1)
Dim aString
If Trim(pString) = "" Then
    For i = 1 To pGrid.Rows - 2:         pGrid.RowHidden(i) = False:    Next
End If
aString = Split(Trim(pString))
For nRow = 1 To pGrid.Rows - 1
    pGrid.RowHidden(nRow) = False
    For i = 0 To UBound(aString)
        If Trim(aString(i)) <> "" Then
            pGrid.RowHidden(nRow) = InStr(1, pGrid.TextMatrix(nRow, pCol), Trim(aString(i))) = 0
           If pGrid.RowHidden(nRow) = True Then Exit For
        End If
    Next
Next
End Sub
Function myNear(nNumber, nMode) As Double
Dim nNum1, nNum2, nNum3
If Val(nNumber) = 0 Then
    myMynear = 0
    Exit Function
ElseIf nMode = 0 Then
    myNear = nNumber
    Exit Function
End If
nNum1 = (nNumber - Fix(nNumber)) * 100
nNum1 = nNum1 Mod (nMode * 100)
myNear = (nNumber) - (nNum1 / 100) + IIf(nNum1 > 0, nMode, 0)
End Function
Function RetSetting(cSearch As String, Optional cFile As String) As String
Dim TextLine As String
On Error GoTo myerror
If cFile = "" Then cFile = App.Path & "\conf.txt"
Open cFile For Input As #1   ' Open file.
Do While Not EOF(1)   ' Loop until end of file.
   Line Input #1, TextLine   ' Read line into variable.
   If InStr(1, LCase(TextLine), LCase(cSearch) & "=") > 0 Then
       RetSetting = Mid(TextLine, Len(cSearch) + 2)
       Exit Do
   End If
Loop
Close #1   ' Close file.
Exit Function
myerror:
Err.Clear
RetSetting = ""
End Function
Function addSetting(cField, cValue, Optional cFile As String) As Boolean
Dim TextLine As String, cText As String, aLocal, nFoundTimes As Integer
Dim fs As New FileSystemObject
If cFile = "" Then cFile = App.Path & "\conf.txt"
On Error GoTo myerror
If fs.FileExists(cFile) Then
    Open cFile For Input As #1   ' Open file.
    Do While Not EOF(1)   ' Loop until end of file.
       Line Input #1, TextLine   ' Read line into variable.
       cText = cText & turn(cText, vbCrLf) & TextLine
    Loop
    Close #1   ' Close file.
End If

aLocal = Split(cText, vbCrLf)
Open cFile For Output As #2   ' Open file.
For i = 0 To UBound(aLocal)
    nFound = InStr(1, LCase(Trim(aLocal(i))), LCase(Trim(cField)) & "=")
    If nFound > 0 Then
        nFoundTimes = nFoundTimes + 1
        If nFoundTimes = 1 Then
            If Trim(cValue) <> "" Then
                Print #2, UCase(Trim(cField)) & "=" & Trim(cValue)
            End If
        End If
    Else
        Print #2, aLocal(i)
    End If
Next
If nFoundTimes = 0 Then
    If Trim(cValue) <> "" Then
        Print #2, Trim(UCase(cField)) & "=" & Trim(cValue)
    End If
End If
Close #2
addSetting = True
Exit Function
myerror:
Err.Clear
End Function
Function turn(ByVal cString, Optional ByVal strFind, Optional ByVal caseFound, Optional ByVal CaseNotfound) As String
If IsMissing(strFind) And IsMissing(caseFound) And IsMissing(CaseNotfound) Then
    If Trim(cString) <> "" Then
        turn = IIf(InStr(1, LCase(cString), " where ") > 0, " AND ", " WHERE ")
    End If
ElseIf (Not IsMissing(strFind)) And IsMissing(caseFound) And IsMissing(CaseNotfound) Then
    If Trim(cString) <> "" Then turn = strFind
ElseIf (Not IsMissing(strFind)) And (Not IsMissing(caseFound)) And IsMissing(CaseNotfound) Then
    If Trim(strFind) <> "" Then
        turn = IIf(InStr(1, LCase(cString), LCase(strFind)) > 0, caseFound, strFind)
    Else
        turn = IIf(Trim(cString) = "", caseFound, strFind)
    End If
ElseIf (Not IsMissing(strFind)) And (Not IsMissing(caseFound)) And (Not IsMissing(CaseNotfound)) Then
    If Trim(strFind) <> "" Then
        turn = IIf(InStr(1, UCase(cString), UCase(strFind)) > 0, caseFound, CaseNotfound)
    Else
        turn = IIf(Trim(cString) = "", caseFound, CaseNotfound)
    End If
End If
End Function
Function ArbString(ByVal pString)
Dim aLocal
pString = Trim(pString)
If pString = "" Then Exit Function
aLocal = Split(pString)
For i = 0 To UBound(aLocal)
    If Trim(aLocal(i)) <> "" Then
        ArbString = ArbString & turn(ArbString, Chr(254) & " ") & Trim(aLocal(i))
    End If
Next
ArbString = Chr(254) & ArbString & Chr(254)
End Function
Function Myvalue(ByVal pValue As Variant, Optional pFormat As String = "") As String
Myvalue = IIf(Val(pValue & "") = 0, "", pValue)
If pFormat <> "" Then Myvalue = Format(Myvalue)
End Function
Private Sub xDate_Validate(Cancel As Boolean)
With xDate
If (Not IsDate(.text)) And Trim(.text) <> "" Then .text = ""
.text = Format(.text, "dd-mm-yyyy")
End With
End Sub

Function NumSql(cField As String) As String
NumSql = "case when " & cField & " = 0 then Null else " & cField & " end "
End Function

Public Sub Inform(Mcaption As String, Optional mCaption2 As String, Optional nInterval As Integer = 900)
On Error Resume Next
'Informfrm.sLabel1 = Mcaption
'Informfrm.sLabel2 = mCaption2
Informfrm.sLabel1 = Mcaption
Informfrm.sLabel2 = mCaption2
Informfrm.nInterval = nInterval
Informfrm.Show 1
Err.Clear
End Sub
Public Sub InformOk(Mcaption As String)
'On Error Resume Next
'Load InformOkfrm
InformOkfrm.Label1.Caption = Mcaption
InformOkfrm.Show 1
Err.Clear
End Sub
Function myRecordSet(pString As String, pCon As adodb.Connection, Optional pType As Integer = adCmdText) As adodb.Recordset
Dim loctable As New adodb.Recordset
Dim cmd As New adodb.command
cmd.CommandType = pType
cmd.CommandTimeout = 600
'If pcon Is Nothing Then cmd.ActiveConnection = GetCon Else
cmd.ActiveConnection = pCon
cmd.CommandText = pString
Set loctable = cmd.Execute
Set myRecordSet = loctable
Set rdTable = Nothing
End Function
Sub myGotFocus(ByRef pControl As Variant, Optional bStart As Boolean = True)
On Error Resume Next
With pControl
If bStart Then
    .SelStart = 0
    .SelLength = Len(pControl.text)
End If
.BackColor = &HC0FFFF
End With
Err.Clear
End Sub
Sub myLostFocus(ByRef pControl As Variant)
    pControl.BackColor = &H80000005
End Sub
Sub myValidDate(ByRef pControl As Variant)
With pControl
If IsDate(.text) Then
    .text = myFormat_p(.text)
Else
    .text = ""
End If
End With
End Sub
Sub myValidDate_en(ByRef pControl As Variant)
With pControl
If IsDate(.text) Then
    .text = Format(.text, "d/m/yyyy")
Else
    .text = ""
End If
End With
End Sub
Sub myValidDate2(ByRef pControl As Variant)
With pControl
If IsDate(.text) Then
    .text = myFormat_p(.text)
Else
    .text = ""
End If
End With
End Sub
Function TempSave(pform As Variant, Optional sFlag As String = "") As String
TempSave = tempPath & turn(tempPath, "\") & pform.Name & ".txt"
End Function
Function LoadConString(Optional ByVal pFile As String = "")
Dim cServerName As String, cUserId As String, cPassword As String, cFile As String
If pFile = "" Then cFile = App.Path & "\conf.txt" Else cFile = pFile
If lServerOnLine Then
'    cServerName = "junior-sql.database.windows.net"
'    cUserId = decrypt("062F8C407C77", "dr")
'    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
Else
    cServerName = RetSetting("server", cFile)
    cUserId = decrypt(RetSetting("userid", cFile), "dr")
    cPassword = decrypt(RetSetting("Password", cFile), "dr")
End If
If (cUserId = "" Or cPassword = "") And pFile = "" Then
    If cServerName = "" Then cServerName = MyParn(".\" & cLocal)
    LoadConString = "provider=SQLOLEDB;data source = " & cServerName & " ;initial " _
            & "catalog=" & sCatalog & ";Trusted_Connection=yes"
Else
    cServerName = cServerName & turn(cLocal, "\") & cLocal
    LoadConString = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
            & "catalog=" & sCatalog & ";user id = " & cUserId & ";" & "password = " & cPassword
End If
End Function
Function ReplaceStr(TextIn, ByVal SearchStr As String, _
                        ByVal Replacement As String, _
                        ByVal CompMode As Integer)
   Dim WorkText As String, Pointer As Integer
     If IsNull(TextIn) Then
       ReplaceStr = Null
     Else
       WorkText = TextIn
       Pointer = InStr(1, WorkText, SearchStr, CompMode)
       Do While Pointer > 0
         WorkText = Left(WorkText, Pointer - 1) & Replacement & _
                    Mid(WorkText, Pointer + Len(SearchStr))
         Pointer = InStr(Pointer + Len(Replacement), WorkText, _
                         SearchStr, CompMode)
       Loop
       ReplaceStr = WorkText
     End If
   End Function
   Function SQLFixup(TextIn)
     SQLFixup = ReplaceStr(TextIn, "'", "''", 0)
   End Function

   Function JetSQLFixup(TextIn)
   Dim temp
     temp = ReplaceStr(TextIn, "'", "''", 0)
     JetSQLFixup = ReplaceStr(temp, "|", "' & chr(124) & '", 0)
   End Function
   Function FindFirstFixup(TextIn)
   Dim temp
     temp = ReplaceStr(TextIn, "'", "' & chr(39) & '", 0)
     FindFirstFixup = ReplaceStr(temp, "|", "' & chr(124) & '", 0)
   End Function
Function emptyRow(pGrid As Variant, Row As Long, Optional Col As Long = -1) As Boolean
For i = IIf(Col = -1, pGrid.FixedCols, Col) To pGrid.Cols - 1
    If Trim(pGrid.TextMatrix(Row, i)) <> "" Then Exit Function
Next
emptyRow = True
End Function
Function createCommand(pString As String, pCon As adodb.Connection) As Boolean
On Error GoTo myerror
Dim FS1 As New adodb.command
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = pCon
FS1.CommandText = pString
FS1.Execute
Set FS1 = Nothing
Exit Function
Err.Clear
Set FS1 = Nothing
myerror:
Set FS1 = Nothing

Err.Clear

End Function
Function RetDiscount(ByVal nValue1 As Double, nValue2 As Double) As Long
nValue1 = nValue1 - 100
nValue1 = (Int(nValue1 / 1000)) * 100
'RetDiscount = nValue1 - Val(xcard_discount & "")
RetDiscount = nValue1 - nValue2
End Function
Function validPhoto(cPhoto) As Boolean
On Error GoTo myerror
If Dir(cPhoto) = "" Or cPhoto = "" Then Exit Function
myPhoto = LoadPicture(cPhoto)
validPhoto = True
Exit Function
myerror:
Err.Clear
End Function
Sub MyLoadPhoto(sModelNo As String, pimgX As ImgXCtrl)
On Error GoTo myerror
pimgX.Images.Clear
pimgX.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub
Function AddFlag(ByVal aString As Variant, ByVal cFlag As Variant, Optional ByVal cFlagValue As Variant, Optional bEdit As Boolean = False) As Variant
If IsEmpty(aString) Then aString = Array()
If IsMissing(cFlagValue) Then
    ReDim Preserve aString(UBound(aString) + 1)
    aString(UBound(aString)) = cFlag
    AddFlag = aString
    Exit Function
End If

If bEdit Then
    If UBound(aString) > 0 Then
        For i = 0 To UBound(aString) Step 2
            If Trim(LCase(aString(i))) = Trim(LCase(cFlag)) Then
                aString(i + 1) = cFlagValue
                AddFlag = aString
                Exit Function
            End If
        Next
    End If
End If
ReDim Preserve aString(UBound(aString) + 2)
aString(UBound(aString) - 1) = cFlag
aString(UBound(aString)) = cFlagValue
AddFlag = aString
End Function
Function retFlag(aString As Variant, cFlag As String) As Variant
If IsEmpty(aString) Then Exit Function
Dim nPos As Long
If UBound(aString) > 0 Then
    For nPos = 0 To UBound(aString) Step 2
        If Trim(LCase(aString(nPos))) = Trim(LCase(cFlag)) Then
            retFlag = aString(nPos + 1)
            Exit For
        End If
    Next
End If
End Function
Function addInsert(ByVal aInsert, ByVal CTABLE) As String
Dim cString1 As String, cString2 As String
For i = 0 To UBound(aInsert) Step 2
    cString1 = cString1 & IIf(cString1 = "", "", ",") & aInsert(i)
    cString2 = cString2 & IIf(cString2 = "", "", ",") & aInsert(i + 1)
Next
addInsert = "Insert into " & CTABLE & " (" & _
                cString1 & _
                ")"
addInsert = addInsert & " values(" & _
                cString2 & _
                ")"
End Function
Function addUpdate(ByVal aInsert, ByVal CTABLE, ByVal cCondition) As String
For i = 0 To UBound(aInsert) Step 2
    addUpdate = addUpdate & turn(addUpdate, ",") & aInsert(i) & _
                   " = " & aInsert(i + 1)
Next
addUpdate = "UPDATE " & CTABLE & " SET " & _
               addUpdate
If cCondition <> "" Then addUpdate = addUpdate & turn(addUpdate) & cCondition
End Function
Function RetPrinter(pName) As Variant
Dim printer As printer, aRet As Variant
For Each printer In Printers
    If LCase(Trim(printer.DeviceName)) = LCase(Trim(pName)) Then
        aRet = AddFlag(aRet, "name", pName)
        aRet = AddFlag(aRet, "port", printer.Port)
        aRet = AddFlag(aRet, "driver", printer.DriverName)
        RetPrinter = aRet
        Exit For
    End If
Next
End Function
Sub FixPrinter(pReport As CrystalReport, Optional pType As String = "1")
Dim aRet As Variant, cPrinterName, cPort As String, cDriver As String
cPrinterName = RetSetting("printer" & pType, tempPath & turn(tempPath, "\") & "printers.txt")
If cPrinterName <> "" Then
    aRet = RetPrinter(cPrinterName)
    If Not IsEmpty(aRet) Then
        pReport.PrinterDriver = retFlag(aRet, "driver")
        pReport.PrinterPort = retFlag(aRet, "port")
        pReport.PrinterName = retFlag(aRet, "name")
    End If
End If
End Sub
Function RetPrinterByType(Optional sType As String = "1") As String
Dim sPrinter As String, aRet As Variant
sPrinter = RetSetting("printer" & sType, tempPath & turn(tempPath, "\") & "printers.txt")
aRet = RetPrinter(sPrinter)
If Not IsEmpty(aRet) Then RetPrinterByType = sPrinter
End Function
Function GetFields(pString, pCon As adodb.Connection) As Variant
Dim loctable As New adodb.Recordset
'If pcon Is Nothing Then
'    loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
'Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
'End If
If Not (loctable.BOF And loctable.EOF) Then
    For i = 0 To loctable.Fields.Count - 1
        GetFields = AddFlag(GetFields, LCase(loctable.Fields(i).Name), loctable.Fields(i).Value)
    Next
End If
loctable.Close
Set loctable = Nothing
End Function
Function GetField(pString, pCon As adodb.Connection) As Variant
Dim loctable As New adodb.Recordset
'If pcon Is Nothing Then
'    loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
'Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
'End If
If Not (loctable.BOF And loctable.EOF) Then
    GetField = loctable(0).Value
End If
loctable.Close
Set loctable = Nothing
End Function
Function GetRows(pString, pCon As adodb.Connection) As Variant
Dim loctable As New adodb.Recordset
'If pcon Is Nothing Then
'    loctable.Open pString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
'Else
    loctable.Open pString, pCon, adOpenStatic, adLockReadOnly, adCmdText
'End If

If Not (loctable.BOF And loctable.EOF) Then
    Dim aRet
    aRet = Array()
    Do Until loctable.EOF
        ReDim Preserve aRet(UBound(aRet) + 1)
    
        For i = 0 To loctable.Fields.Count - 1
            aRet(UBound(aRet)) = AddFlag(aRet(UBound(aRet)), LCase(loctable.Fields(i).Name), loctable.Fields(i).Value)
        Next
        loctable.MoveNext
    Loop
    GetRows = aRet
End If
loctable.Close
Set loctable = Nothing
End Function


Sub LoadText(myform As Form, Optional bForAll As Boolean = False, Optional sFlag As String)
Dim cFilesave As String, cText As String
If bForAll Then cFilesave = App.Path & "\" & myform.Name & turn(sFlag, "_") & sFlag & ".txt" Else cFilesave = tempPath & "\" & myform.Name & turn(sFlag, "_") & sFlag & ".txt"
For i = 0 To myform.Count - 1
    If TypeOf myform(i) Is TextBox Then
        myform(i).text = RetSetting(myform(i).Name, cFilesave)
    ElseIf TypeOf myform(i) Is DataCombo Then
        myform(i).BoundText = RetSetting(myform(i).Name, cFilesave)
    ElseIf TypeOf myform(i) Is CheckBox Then
        If RetSetting(myform(i).Name, cFilesave) <> "" Then myform(i).Value = RetSetting(myform(i).Name, cFilesave)
    End If
Next
End Sub

Sub SaveText(myform As Form, Optional bForAll As Boolean = False, Optional aText As Variant, Optional sFlag As String = "")
Dim cFilesave As String
If bForAll Then cFilesave = App.Path & "\" & myform.Name & turn(sFlag, "_") & sFlag & ".txt" Else cFilesave = tempPath & "\" & myform.Name & turn(sFlag, "_") & sFlag & ".txt"
If Not IsMissing(aText) Then
    For i = 0 To UBound(aText)
        cText = cText & turn(cText, "", "@") & LCase(aText(i)) & "@"
    Next
End If
For i = 0 To myform.Count - 1
    If InStr(1, cText, "@" & LCase(myform(i).Name) & "@") > 0 Or IsMissing(aText) Then
        If TypeOf myform(i) Is TextBox Then
            addSetting myform(i).Name, myform(i).text, cFilesave
        ElseIf TypeOf myform(i) Is DataCombo Then
            addSetting myform(i).Name, myform(i).BoundText, cFilesave
        ElseIf TypeOf myform(i) Is CheckBox Then
            addSetting myform(i).Name, myform(i).Value, cFilesave
        End If
    End If
Next
End Sub
Function LoadConString_J(Optional ByVal pFile As String = "")
Dim cServerName As String, cUserId As String, cPassword As String, cFile As String
If pFile = "" Then cFile = App.Path & "\conf.txt" Else cFile = pFile
cServerName = RetSetting("server", cFile)
cUserId = decrypt(RetSetting("userid", cFile), "dr")
cPassword = decrypt(RetSetting("Password", cFile), "dr")
cServerName = cServerName & turn(cLocal, "\") & cLocal
LoadConString_J = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
        & "catalog=JUNIOR_SHOP" & ";user id = " & cUserId & ";" & "password = " & cPassword
End Function
Public Function Tr(pString As Variant, Optional pReturn As String = " AND ") As String
Tr = IIf(Trim(pString & "") = "", "", pReturn)
End Function
Sub SaveImageToAccess(pPic As Picture, rs As adodb.Recordset, pColName As String)
Dim pb As PropertyBag
Set pb = New PropertyBag
pb.WriteProperty "MyImage", pPic
rs.Fields(pColName).AppendChunk pb.Contents
rs.Update
Set pb = Nothing
End Sub
Function GetPictureFromRecordset(rs As adodb.Recordset, _
                                         pColName As String) As Picture
    Dim pb As PropertyBag

    Set pb = New PropertyBag
    pb.Contents = rs.Fields(pColName).GetChunk(rs.Fields(pColName).ActualSize)
    Set GetPictureFromRecordset = pb.ReadProperty("MyImage")
    Set pb = Nothing
End Function
Function NextEmpty(pGrid As Object, Row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
Dim nLast
For i = IIf(nBegincol = -1, 0, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, IIf(nEndCol > pGrid.Cols - 1, pGrid.Cols - 1, nEndCol))
    If Trim(pGrid.TextMatrix(Row, i)) = "" And pGrid.ColHidden(i) = False Then
        NextEmpty = i
        Exit Function
    End If
Next
NextEmpty = IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
End Function
Function mySplit(pString, Optional nPos As Integer = 1, Optional pchr As String = ",") As String
Dim aString
aString = Split(pString, pchr)
If UBound(aString) >= nPos - 1 Then
    mySplit = aString(nPos - 1)
End If
End Function
Public Function DefUser() As Boolean
DefUser = RetSetting("DefUser", "C:\WORK\USERS\Users.txt") = "1"
End Function


