Attribute VB_Name = "special2"
Public rsBranch As ADODB.Recordset
Public rsMall As ADODB.Recordset
Public rsBranches As ADODB.Recordset
Public rsUser As ADODB.Recordset
Public rsAddress As ADODB.Recordset
Public Function openConDoc(ByRef con As ADODB.Connection, Optional pCatalog As String = "TRADE_DOCS") As String
Dim cServerName As String
Dim cUserId As String
Dim cPassword As String
Dim cString As String

cServerName = RetSetting("server", App.Path & "\CONF.TXT")
cUserId = decrypt(RetSetting("userId", App.Path & "\CONF.TXT"), "dr")
cPassword = decrypt(RetSetting("Password", App.Path & "\CONF.TXT"), "dr")

cString = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
                & "catalog=" & pCatalog & ";user id=" & cUserId & ";" & "password=" & cPassword & _
                ";Encrypt=True" & _
                ";TrustServerCertificate=False" & _
                ";Timeout=10"
openConDoc = openCon(con, cString)
End Function
Public Function ConDocString(Optional pCatalog As String = "TRADE_DOCS", Optional pTimeOut As Integer = 3) As String
Dim cServerName As String
Dim cUserId As String
Dim cPassword As String
Dim cString As String

cServerName = RetSetting("server", App.Path & "\CONF.TXT")
cUserId = decrypt(RetSetting("userId", App.Path & "\CONF.TXT"), "dr")
cPassword = decrypt(RetSetting("Password", App.Path & "\CONF.TXT"), "dr")

ConDocString = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
                & "catalog=" & pCatalog & ";user id=" & cUserId & ";" & "password=" & cPassword & _
                ";Encrypt=True" & _
                ";TrustServerCertificate=False" & _
                ";Timeout=" & pTimeOut

End Function

Sub accountLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT ACCOUNT.CODE," & _
          " ACCOUNT.desca," & _
          " ACCOUNT.WITH_BRANCH " & _
          " FROM  ACCOUNT"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by ACCOUNT.code"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%ACCOUNT.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·Õ”«»"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "≈”„ «·Õ”«»"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "»›—⁄"
GrdArray(2, 1) = 0

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "ACCOUNT.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·Õ”«»« "
oSearch.Show 1
End Sub
Sub BranchLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT Branch.CODE," & _
          " Branch.desca" & _
          " FROM  Branch"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by Branch.code"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·›—⁄"
listarray(0, 1) = "(%%Branch.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·›—⁄"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "≈”„ «·›—⁄"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "BRANCH.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·›—Ê⁄"
oSearch.Show 1
End Sub
Sub CLIENTLOOKUP(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT FILE3_10.CODE, FILE3_10.desca, FILE3_50.DESCA FROM  FILE3_10 LEFT JOIN FILE3_50 ON FILE3_10.[GROUP] = FILE3_50.CODE"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by FILE3_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·„Ã„Ê⁄…"
listarray(0, 1) = "(%%FILE3_10.DESCA%% OR %%FILE3_50.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE3_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.sCaption = "≈” ⁄·«„ «·⁄„·«¡"
oSearch.Show 1
End Sub
Sub suplookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT FILE4_10.CODE," & _
          " FILE4_10.desca," & _
          " FILE4_50.DESCA " & _
          " FROM  FILE4_10 LEFT JOIN FILE4_50 ON FILE4_10.[GROUP] = FILE4_50.CODE"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by FILE4_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·„Ã„Ê⁄…"
listarray(0, 1) = "(%%FILE4_10.DESCA%% OR %%FILE4_50.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·„Ê—œ"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE4_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·„Ê—œÌ‰"
oSearch.Show 1
End Sub
Sub BoxLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT FILE0_50.CODE," & _
          "FILE0_50.DESCA," & _
          " branch.desca AS BRANCH_DESCA" & _
          " FROM FILE0_50 INNER JOIN branch ON FILE0_50.BRANCH = branch.code"
    
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE0_50.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%FILE0_50.DESCA%% OR %%BRANCH.DESCA%%)"

listarray(1, 0) = "«·›—⁄"
listarray(1, 1) = "(FILE0_50.BRANCH = 'cFilter')"
listarray(1, 2) = "SELECT BRANCH.CODE,BRANCH.DESCA FROM BRANCH"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "ﬂÊœ «·Œ“‰…"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·›—⁄"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE0_50.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·Œ“‰"
oSearch.Show 1
End Sub
Sub BankLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm
                 
cString = "SELECT FILE5_10.CODE," & _
          "FILE5_10.desca," & _
          " FILE5_50.DESCA " & _
          " FROM  FILE5_10 " & _
          " LEFT JOIN FILE5_50 ON FILE5_10.[GROUP] = FILE5_50.CODE"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE5_10.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·„Ã„Ê⁄…"
listarray(0, 1) = "(%%FILE5_10.DESCA%% OR %%FILE5_50.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·»‰ﬂ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·»‰ﬂ"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE5_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·»‰Êﬂ"
oSearch.Show 1
End Sub
Sub partLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT FILE8_71.CODE," & _
          " FILE8_71.desca" & _
          " FROM  FILE8_71"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE8_71.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«”„"
listarray(0, 1) = "(%%FILE8_71.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·‘—Ìﬂ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·‘—Ìﬂ"
GrdArray(1, 1) = 4000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE4_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·‘—ﬂ«¡"
oSearch.Show 1
End Sub
Sub ChargeLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT FILE8_51.CODE," & _
          " FILE8_51.DESCA," & _
          " FILE8_51.desca AS GROUP_DESCA" & _
          " FROM FILE8_51 LEFT JOIN FILE8_52 ON FILE8_51.[MAINGROUP] = FILE8_52.CODE"
    
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE8_51.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·„’—Ê›"
listarray(0, 1) = "(%%FILE8_51.DESCA%% OR %%FILE8_52.DESCA%%)"

listarray(1, 0) = "„Ã„Ê⁄… —∆Ì”Ì…"
listarray(1, 1) = "(FILE8_51.[MAINGROUP] = 'cFilter')"
listarray(1, 2) = "SELECT FILE8_52.CODE,FILE8_52.DESCA FROM FILE8_52"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "ﬂÊœ «·„’—Ê›"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«·„’—Ê›"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE5_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.sCaption = "≈” ⁄·«„ «·„’«—Ì›"
oSearch.Show 1
End Sub
Sub IncomeLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT FILE8_61.CODE," & _
          "FILE8_61.DESCA," & _
          " FILE8_61.desca AS GROUP_DESCA" & _
          " FROM FILE8_61 LEFT JOIN FILE8_62 ON FILE8_61.[MAINGROUP] = FILE8_62.CODE"
    
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by FILE8_61.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·«Ì—«œ"
listarray(0, 1) = "(%%FILE8_61.DESCA%% OR %%FILE8_62.DESCA%%)"

listarray(1, 0) = "„Ã„Ê⁄… —∆Ì”Ì…"
listarray(1, 1) = "(FILE8_61.MAINGROUP = 'cFilter')"
listarray(1, 2) = "SELECT FILE8_62.CODE,FILE8_62.DESCA FROM FILE8_62"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "ﬂÊœ «·«Ì—«œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«·«Ì—«œ"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE5_10.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·«Ì—«œ« "
oSearch.Show 1
End Sub
Sub AccLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1
cString = "SELECT CODE," & _
          " DESCA" & _
          " FROM   ACC0_10"
If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by CODE"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·«”„-«·ﬂÊœ"
listarray(0, 1) = "(**CODE** OR %%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1500

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·Õ”«»« "
oSearch.Show 1
End Sub
Public Function accountRs(con As ADODB.Connection, Optional id_image As String, Optional id_cash As String) As ADODB.Recordset
Dim aPrm As Variant
If id_image <> "" Then
    aPrm = AddFlag(aPrm, "id_image", id_image)
ElseIf id_cash <> "" Then
    aPrm = AddFlag(aPrm, "id_cash", id_cash)
End If
Set accountRs = cmd("[dbo].[sp_get_account_image]", con, adStoredProc, aPrm).Execute
End Function

Function addText(cField, cValue, Optional cFile As String) As Boolean
Dim TextLine As String, cText As String, aLocal, nFoundTimes As Integer
Dim fs As New FileSystemObject
If cFile = "" Then cFile = cPathConf & "\CONF.TXT"
Dim FileNumber
On Error GoTo myerror
If fs.FileExists(cFile) Then
    FileNumber = FreeFile()
    Open cFile For Input As #FileNumber   ' Open file.
    Do While Not EOF(FileNumber)   ' Loop until end of file.
       Line Input #FileNumber, TextLine   ' Read line into variable.
       cText = cText & Tr(cText, vbCrLf) & TextLine
    Loop
    Close #FileNumber   ' Close file.
End If

aLocal = Split(cText, vbCrLf)

FileNumber = FreeFile()
Open cFile For Output As #FileNumber   ' Open file.
For i = 0 To UBound(aLocal)
    nFound = InStr(1, LCase(Trim(aLocal(i))), LCase(Trim(cField)) & "=")
    If nFound > 0 Then
        nFoundTimes = nFoundTimes + 1
        If nFoundTimes = 1 Then
            If Trim(cValue) <> "" Then
                Print #FileNumber, UCase(Trim(cField)) & "=" & Trim(cValue)
            End If
        End If
    Else
        Print #FileNumber, aLocal(i)
    End If
Next

If nFoundTimes = 0 Then
    If Trim(cValue) <> "" Then
        Print #FileNumber, Trim(UCase(cField)) & "=" & Trim(cValue)
    End If
End If
Close #FileNumber
addText = True
Exit Function
myerror:
Err.Clear
End Function
Function isNum(pNumber As Variant) As Boolean
If Round(Val(pNumber & ""), 0) & "" <> Trim(pNumber & "") Then Exit Function
isNum = True
End Function
Public Function YearsString(Optional pYearBegin As Integer = 0, Optional pYearEnd As Integer = 0)
Dim i As Long
YearsString = "SELECT '' AS CODE"
For i = IIf(pYearEnd = 0, Year(Date), pYearEnd) To IIf(pYearBegin = 0, Year(Date), pYearBegin) Step -1
    YearsString = YearsString & Tr(YearsString, " UNION All ") & "SELECT " & "CAST(" & i & " AS VARCHAR(4)) " & " AS [CODE]"
Next
End Function
Public Function MonthString(Optional pMonth As Integer = 12)
Dim i As Long
MonthString = "SELECT '' AS CODE"
For i = pMonth To 1 Step -1
    MonthString = MonthString & Tr(MonthString, " UNION All ") & "SELECT " & "CAST(" & i & " AS VARCHAR(2)) " & " AS [CODE]"
Next
End Function
Public Function get_Id(con As ADODB.Connection) As String
Dim loctable As New ADODB.Recordset
Set loctable = cmd("select SCOPE_IDENTITY() AS ID", con).Execute
If Not loctable.EOF Then
    get_Id = loctable!ID & ""
End If

loctable.Close
Set loctable = Nothing
End Function
Sub ItemsLook(oForm As Form, oSearch As Search_abd, Optional pOrderNo As String = "", Optional pOffer As String = "", Optional bUnLoad As Boolean = False)
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(9, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "SELECT  Top 1000 item," & _
                  "file1_10.mosm," & _
                  "fact.desca," & _
                  "file1_10.modelfact0," & _
                  "file1_10.supp," & _
                  "file1_10.desca," & _
                  "file1_10.scal," & _
                  "file1_10.color," & _
                  "PRICE," & _
                  "PRICE_2 " & _
                  "from file1_10 " & _
                  "inner join fact on fact.code = file1_10.fact" & _
                  " WHERE ISSTOP = 0 "
If pOrderNo <> "" Then
    Generalarray(1) = Generalarray(1) & " AND ITEM IN (SELECT ITEM FROM FILE6_50 WHERE DOC_NO = " & MyParn(pOrderNo) & ")"
End If

If pOffer <> "" Then
    Generalarray(1) = Generalarray(1) & " AND FILE1_10.MODELNO IN (SELECT MODELNO FROM FILE0_90 WHERE FILE0_90.DOC_NO = " & MyParn(pOffer) & ")"
End If


Generalarray(2) = "ORDER BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_SCAL "
Generalarray(3) = 12000
Generalarray(5) = False

listarray(0, 0) = "≈”„ «·’‰› ° «·„’‰⁄ ° —ﬁ„ „ÊœÌ· & „ﬁ«” "
listarray(0, 1) = "(%%FILE1_10.DESCA%% OR %%file1_10.MODELFACT%% OR %%FACT.DESCA%%)"

listarray(1, 0) = "„Ê”„ "
listarray(1, 1) = "(%%File1_10.MOSM%%)"

listarray(2, 0) = "„ﬁ«”"
listarray(2, 1) = "%%File1_10.SCAL%%"

listarray(3, 0) = "«··Ê‰"
listarray(3, 1) = "(%%File1_10.COLOR%%)"
 
GrdArray(0, 0) = "»«—ﬂÊœ"
GrdArray(0, 1) = 900

GrdArray(1, 0) = "„Ê”„"
GrdArray(1, 1) = 800

GrdArray(2, 0) = "«·„’‰⁄"
GrdArray(2, 1) = 1800

GrdArray(3, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "„ﬂ »"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·’‰›"
GrdArray(5, 1) = 4500

GrdArray(6, 0) = "«·„ﬁ«”"
GrdArray(6, 1) = 1300

GrdArray(7, 0) = "«··Ê‰"
GrdArray(7, 1) = 800

GrdArray(8, 0) = " «·”⁄—"
GrdArray(8, 1) = 1000

GrdArray(9, 0) = " ”⁄— «Êﬂ«“ÊÌ‰"
GrdArray(9, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.bUnLoad = bUnLoad
oSearch.bNoRef = True
oSearch.sCaption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Function IncData(sTable As String, sField As String, con As ADODB.Connection, Optional pwhere As String = "", Optional pDefault As String) As String
Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "Select Max(" & sField & ") as Maxof " & _
       " FROM " & sTable
If pwhere <> "" Then cString = cString & " where " & pwhere
Set loctable = cmd(cString, con).Execute
If Not (loctable.EOF) Then
    If IsNull(loctable!maxOf) Then
        IncData = pDefault
    Else
        IncData = IncrementString(loctable!maxOf & "")
    End If
End If
loctable.Close
Set loctable = Nothing
End Function
Public Function IncrementString(ByVal currentStr As String) As String
    Dim Length As Integer
    Dim numericPart As String
    Dim alphaPart As String
    Dim maxNumeric As Long
    Dim currentNum As Double
    Dim i As Integer
    Length = Len(currentStr)
    If IsNumeric(currentStr) Then
        Length = Len(currentStr)
        currentNum = Val(currentStr) + 1
        
        ' ??? ????? ????? ???? ?????? (??? 999)? ???? ?? A01
        If Len(CStr(currentNum)) > Length Then
            If Length = 1 Then
                IncrementString = "A"
            Else
                IncrementString = "A" & Format(1, String(Length - 1, "0"))
            End If
        Else
            ' ????? ????? ?? ?????? ??? ??????? (Replicate Zeros)
            IncrementString = Format(currentNum, String(Length, "0"))
        End If
        Exit Function
    End If
    
    
    
    ' ??? ????? ????? ?? ?????? ?? ??????
    i = Length
    Do While i > 0 And IsNumeric(Mid(currentStr, i, 1))
        i = i - 1
    Loop
    
    alphaPart = Left(currentStr, i)
    numericPart = Right(currentStr, Length - i)
    maxNumeric = (10 ^ Len(numericPart)) - 1
    
    ' ?????? ??????: ??? ??? ????? ?????? ??? ?? ???? ??????
    If Val(numericPart) < maxNumeric Then
        currentNum = Val(numericPart) + 1
        IncrementString = alphaPart & Right(String(Len(numericPart), "0") & currentNum, Len(numericPart))
    
    ' ?????? ???????: ?????? ???? ?????? (????? 999 ?? A99)
    Else
        ' ??? ??? ?????? ????? ??? (??? 999) ????? ??? A01
        If alphaPart = "" Then
            IncrementString = "A" & Right(String(Length - 1, "0") & "1", Length - 1)
        Else
            ' ????? ????? ????? (A -> B ?? Z -> AA)
            alphaPart = IncrementAlpha(alphaPart)
            If Len(alphaPart) > Length Then
                IncrementString = Left(currentStr, i)
            Else
                IncrementString = Left(alphaPart & Right(String(Length, "0") & "1", Length - Len(alphaPart)), Length)
            End If
        End If
    End If
End Function
' ???? ?????? ?????? ?????? (A ????? B ? ? Z ????? AA)
Private Function IncrementAlpha(ByVal alpha As String) As String
    Dim lastChar As Integer
    If alpha = "" Then
        IncrementAlpha = "A"
    Else
        lastChar = Asc(Right(alpha, 1))
        If lastChar < 90 Then ' ??? ?? Z
            IncrementAlpha = Left(alpha, Len(alpha) - 1) & Chr(lastChar + 1)
        Else
            IncrementAlpha = IncrementAlpha(Left(alpha, Len(alpha) - 1)) & "A"
        End If
    End If
End Function
Public Sub CLIENTLOOKUP2(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = oForm
cString = "SELECT FILE3_10.CODE," & _
           " FILE3_10.DESCA," & _
           " FILE3_50.DESCA" & _
           " FROM FILE3_10 " & _
           " LEFT JOIN FILE3_50 ON FILE3_10.[GROUP] = FILE3_50.CODE"

If pFilter <> "" Then cString = cString & " WHERE " & pFilter

Generalarray(1) = cString
Generalarray(2) = "order by FILE3_10.DESCA"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì·"
listarray(0, 1) = "%%FILE3_10.DESCA%%"

listarray(1, 0) = "«·›—⁄"
listarray(1, 1) = "(FILE3_10.[GROUP] = 'cFilter')"
listarray(1, 2) = "SELECT CODE,DESCA FROM FILE3_50"
listarray(1, 3) = "CODE"
listarray(1, 4) = "DESCA"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 900

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2500

Dim aFilter As Variant
aFilter = AddFlag(aFilter, "FILTER", True)
aFilter = AddFlag(aFilter, "FIELD", "CODE")

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.sCaption = "«” ⁄·«„ «·⁄„·«¡"
oSearch.aFilter = aFilter
oSearch.Show 1
End Sub
Sub StoreLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1                   2                   3                   4               5
cString = "SELECT file0_40.CODE," & _
          " file0_40.desca" & _
          " FROM  file0_40"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by file0_40.code"
Generalarray(3) = 4500
Generalarray(5) = True

listarray(0, 0) = "«·„Œ“‰"
listarray(0, 1) = "(%%file0_40.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ «·„Œ“‰"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "≈”„ «·„Œ“‰"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "file0_40.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·„Œ«“‰"
oSearch.Show 1
End Sub
Sub ColorLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

Generalarray(1) = "SELECT [C_COLOR]," & _
                   "[COLOR]" & _
                   "  FROM FILE1_10"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then Generalarray(1) = Generalarray(1) & " WHERE " & cWhere

Generalarray(2) = "Order by C_COLOR"
Generalarray(3) = 4500
Generalarray(5) = True

listarray(0, 0) = "COLOR"
listarray(0, 1) = "(%%COLOR%%)"

GrdArray(0, 0) = "ﬂÊœ «··Ê‰"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«··Ê‰"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «·«·Ê«‰"
oSearch.Show 1
End Sub
Public Sub SearchGrd(pGrid As Object, searchText As String, Col As Long, Optional bExct As Boolean = False)
Dim R As Long

' Stop redrawing for better performance
pGrid.Redraw = flexRDNone

For R = pGrid.FixedRows To pGrid.Rows - 1
    ' Check if search text exists in specific column (e.g., Column 1)
    If bExct Then
        If Trim(searchText) = "" Then
            pGrid.RowHeight(R) = -1
        ElseIf LCase$(pGrid.TextMatrix(R, Col)) = searchText Then
            ' Reset row height to default (-1) to show it
            pGrid.RowHeight(R) = -1
        Else
            ' Set row height to 0 to hide it
            'pGrid.RowHidden(r) = True
            pGrid.RowHeight(R) = 0
        End If
        
    Else
        If Trim(searchText) = "" Then
            pGrid.RowHeight(R) = -1
        ElseIf InStr(1, LCase$(pGrid.TextMatrix(R, Col)), searchText) > 0 Then
            ' Reset row height to default (-1) to show it
            pGrid.RowHeight(R) = -1
        Else
            ' Set row height to 0 to hide it
            'pGrid.RowHidden(r) = True
            pGrid.RowHeight(R) = 0
        End If
    End If
Next R
'salesfrm.grid1.Redraw = flexRDBuffered
' Resume redrawing
pGrid.Redraw = flexRDBuffered
End Sub
Public Function IsDgt(ByVal strData As String) As Boolean
    ' Check if string is empty
    If Len(Trim(strData)) = 0 Then
        IsDgt = False
        Exit Function
    End If
    
    ' Pattern checks if the entire string consists only of 0-9
    ' The String(Len(strData), "#") creates a pattern like "####"
    IsDgt = (Trim(strData) Like String(Len(Trim(strData)), "#"))
End Function
Sub SaveImageToRs(pPic As Picture, rs As ADODB.Recordset, pColName As String)
Dim pb As PropertyBag
Set pb = New PropertyBag
pb.WriteProperty "MyImage", pPic, 100
rs.Fields(pColName).AppendChunk pb.Contents
rs.Update
Set pb = Nothing
End Sub
Public Function SaveImageToDB(ByVal FilePath As String, ByVal ImgX1 As Variant, Optional ByVal TableName As String, Optional ByVal FieldName As String, Optional ByVal IDNAME As String, Optional ByVal IDValue As Long, Optional con As ADODB.Connection) As Boolean
    Dim rs As New ADODB.Recordset
    Dim mstream As New ADODB.Stream
    Dim sql As String
    
    ' 1. Open the stream and load the image file
    Dim bytes() As Byte
    If FilePath <> "" Then
        mstream.Type = adTypeBinary
        mstream.Open
        mstream.LoadFromFile FilePath
    Else
        ImgX1.Export.ToMemoryFile bytes, ixmfJPG
    End If
    ' 2. Fetch the record you want to update
    sql = "SELECT " & FieldName & " FROM " & TableName & " WHERE " & IDNAME & " = " & IDValue
    rs.Open sql, con, adOpenStatic, adLockOptimistic
    
    If Not rs.EOF Then
        ' 3. Assign the stream's contents to the varbinary field
        If FilePath <> "" Then
            rs.Fields(FieldName).Value = mstream.Read
        Else
            rs.Fields(FieldName).Value = bytes
        End If
        rs.Update
    End If
    
    ' Cleanup
    If FilePath <> "" Then
        mstream.Close
    End If
    rs.Close
    Set mstream = Nothing
    Set rs = Nothing
End Function









