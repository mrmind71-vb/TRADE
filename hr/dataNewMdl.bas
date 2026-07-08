Attribute VB_Name = "dataNewMdl"
Public Enum cmType
adStoredProc = adCmdStoredProc
adText = adCmdText
adTable = adCmdTable
End Enum

Public Enum tbMode
    tbFind = 1
    tbPrevious = 2
    tblast = 3
    tbNext = 4
    tbFirst = 5
End Enum

Public Enum tableMode
found = 1
NotFound = 2
NoRecords = 3
NotMatch = 3
End Enum
Public Function myCmd(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 300) As ADODB.Recordset
Dim loctable As New ADODB.Recordset
Dim cmd As New ADODB.Command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
Set myCmd = cmd.Execute
End Function
Public Function myPrcString(pProc As String, Optional aParam As Variant = Empty) As String
Dim i As Long
For i = 0 To UBound(aParam) Step 2
    myPrcString = myPrcString & IIf(myPrcString = "", "", ",") & "@" & aParam(i) & " = " & aParam(i + 1)
Next
myPrcString = "EXEC " & pProc & " " & myPrcString
End Function
Public Function mycmdEx(pString As String, Optional con As ADODB.Connection, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 100) As ADODB.Command
Dim loctable As New ADODB.Recordset
Dim cmd As New ADODB.Command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
cmd.Execute
Set mycmdEx = cmd
End Function
Public Function myField(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional pDef As Variant = Empty, Optional nTimeOut As Integer = 100) As Variant
On Error GoTo myerror:
Dim cmd As New ADODB.Command
Dim loctable As ADODB.Recordset
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
Set loctable = cmd.Execute
If Not loctable.EOF Then
    myField = loctable.Fields(0).Value
ElseIf Not IsEmpty(pDef) Then
    myField = pDef
End If
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
myField = pDef
End Function
Public Function myFields(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 100) As Variant
Dim cmd As New ADODB.Command
Dim loctable As ADODB.Recordset
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim i As Long
    For i = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
    Next
End If
Set loctable = cmd.Execute
If Not (loctable.BOF And loctable.EOF) Then
    For i = 0 To loctable.Fields.Count - 1
        myFields = AddFlag(myFields, LCase(loctable.Fields(i).Name), loctable.Fields(i).Value)
    Next
End If
loctable.Close
Set loctable = Nothing
End Function
Public Function openConHr(ByRef pCon As ADODB.Connection, Optional ByVal pString As String = "", Optional ByVal lMsg As Boolean = True) As String
On Error GoTo myerror
Dim cString As String
If pString = "" Then cString = LoadConStringHr Else cString = pString
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.Open cString
openConHr = "ok"
Exit Function
myerror:
openConHr = Err.Description
Err.Clear
End Function
Public Function LoadConStringHr()
Dim cServerName As String, cUserId As String, cPassword As String, sCatalogHr As String
'If lServerOnLine Or lServerOnLineShop Then
    cServerName = "junior-sql.database.windows.net"
    cUserId = decrypt("062F8C407C77", "dr")
    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
    sCatalogHr = "JUNIOR_SHOP"
'Else
'    cServerName = RetSetting("server", App.Path & "\hr\confHr.txt")
'    cUserId = decrypt(RetSetting("userId", App.Path & "\hr\confHr.txt"), "dr")
'    cPassword = decrypt(RetSetting("Password", App.Path & "\hr\confHr.txt"), "dr")
'    sCatalogHr = RetSetting("Catalog", App.Path & "\conf.txt")
'End If
If cServerName = "" Then cServerName = "."
LoadConStringHr = "provider=SQLOLEDB;data source=" & MyParn(cServerName) & ";initial " _
            & "catalog =" & sCatalogHr & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Function myFormat_p(sDate As Variant) As String
myFormat_p = Format(sDate, "YYYY/M/D")
End Function


