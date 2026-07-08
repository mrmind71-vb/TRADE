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
Found = 1
NotFound = 2
NoRecords = 3
NotMatch = 3
End Enum
Public Function mycmd(pString As String, Optional con As adodb.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 300) As adodb.Recordset
Dim loctable As New adodb.Recordset
Dim cmd As New adodb.command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim I As Long
    For I = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(I)).Value = aParam(I + 1)
    Next
End If
Set mycmd = cmd.Execute
End Function
Public Function myCmdRecord(pString As String, Optional con As adodb.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 300) As adodb.Record
Dim loctable As New adodb.Recordset
Dim cmd As New adodb.command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim I As Long
    For I = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(I)).Value = aParam(I + 1)
    Next
End If

Set myCmdRecord = cmd.Execute
End Function

Public Function myPrcString(pProc As String, Optional aParam As Variant = Empty) As String
Dim I As Long

If Not IsEmpty(aParam) Then
    For I = 0 To UBound(aParam) Step 2
        myPrcString = myPrcString & IIf(myPrcString = "", " ", ",") & "@" & aParam(I) & " = " & aParam(I + 1)
    Next
End If
myPrcString = "EXEC " & pProc & myPrcString
End Function
Public Function mycmdEx(pString As String, Optional con As adodb.Connection, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 600) As adodb.command
Dim loctable As New adodb.Recordset
Dim cmd As New adodb.command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = adCmdStoredProc
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim I As Long
    For I = 0 To UBound(aParam) Step 2
       cmd.Parameters("@" & aParam(I)).Value = aParam(I + 1)
    Next
End If
cmd.Execute
Set mycmdEx = cmd
End Function
Public Function myCommand(pString As String, Optional con As adodb.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 300) As adodb.command
Dim loctable As New adodb.Recordset
Dim cmd As New adodb.command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim I As Long
    For I = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(I)).Value = aParam(I + 1)
    Next
End If
cmd.Execute
Set myCommand = cmd
End Function
Public Function cmd(pString As String, Optional con As adodb.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 1000) As adodb.command
Set cmd = New adodb.command
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim I As Long
    For I = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(I)).Value = aParam(I + 1)
    Next
End If
End Function
Public Function myField(pString As String, Optional con As adodb.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional pDef As Variant = Empty, Optional nTimeOut As Integer = 100) As Variant
On Error GoTo myError:
Dim loctable As adodb.Recordset
Set loctable = cmd(pString, con, pType).Execute
'cmd.CommandTimeout = nTimeOut
'cmd.ActiveConnection = con
'cmd.CommandType = pType
'cmd.CommandText = pString
'If Not IsEmpty(aParam) Then
'    Dim i As Long
'    For i = 0 To UBound(aParam) Step 2
'        cmd.Parameters("@" & aParam(i)).Value = aParam(i + 1)
'    Next
'End If
If Not loctable.EOF Then
    myField = loctable.Fields(0).Value
ElseIf Not IsEmpty(pDef) Then
    myField = pDef
End If
loctable.Close
Finally:
Set loctable = Nothing
Exit Function
myError:
MsgBox Err.Description
'Err.Clear
myField = pDef
GoTo Finally
End Function
Public Function myFieldValue(pString As String, pField As String, Optional con As adodb.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional pDef As Variant = Empty, Optional nTimeOut As Integer = 100) As Variant
On Error GoTo myError:
Dim cmd As New adodb.command
Dim loctable As adodb.Recordset
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim I As Long
    For I = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(I)).Value = aParam(I + 1)
    Next
End If
Set loctable = cmd.Execute
If Not loctable.EOF Then
    myFieldValue = loctable.Fields(pField).Value
ElseIf Not IsEmpty(pDef) Then
    myFieldValue = pDef
End If
Exit Function
myError:
MsgBox Err.Description
Err.Clear
myFieldValue = pDef
End Function
Public Function myFields(pString As String, Optional con As adodb.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 100) As Variant
Dim cmd As New adodb.command
Dim loctable As adodb.Recordset
cmd.CommandTimeout = nTimeOut
cmd.ActiveConnection = con
cmd.CommandType = pType
cmd.CommandText = pString
If Not IsEmpty(aParam) Then
    Dim I As Long
    For I = 0 To UBound(aParam) Step 2
        cmd.Parameters("@" & aParam(I)).Value = aParam(I + 1)
    Next
End If
Set loctable = cmd.Execute
If Not (loctable.BOF And loctable.EOF) Then
    For I = 0 To loctable.Fields.Count - 1
        myFields = AddFlag(myFields, LCase(loctable.Fields(I).Name), loctable.Fields(I).Value)
    Next
End If
loctable.Close
Set loctable = Nothing
End Function
Public Function openConHr(ByRef pCon As adodb.Connection, Optional ByVal pString As String = "", Optional ByVal lMsg As Boolean = True) As String
On Error GoTo myError
Dim cString As String
If pString = "" Then cString = LoadConStringHr Else cString = pString
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.Open cString
openConHr = "ok"
Exit Function
myError:
openConHr = Err.Description
Err.Clear
End Function
Public Function LoadConStringHr()
Dim cServerName As String, cUserId As String, cPassword As String, sCatalogHr As String
'If lServerOnLine Or lServerOnLineShop Then

    If Not vpn Then
        cServerName = "junior-sql.database.windows.net"
        cUserId = decrypt("062F8C407C77", "dr")
        cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
        sCatalogHr = "JUNIOR_SHOP"
    Else
        cServerName = servername_vpn
        cUserId = userid_vpn
        cPassword = password_vpn
        sCatalogHr = "JUNIOR_SHOP"
    End If
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
Function myFormat_p(sDate As Variant, Optional bLong As Boolean = False) As String
If Not IsDate(sDate) Then
    myFormat_p = ""
Else
    myFormat_p = Format(sDate & "", "YYYY/M/D") & IIf(bLong, Format(sDate & "", " HH:NN"), "")
End If
End Function
Function TransCount(con As adodb.Connection) As Integer
Dim loctable As New adodb.Recordset
Set loctable = mycmd("select @@TRANCOUNT as myCount", con)
On Error GoTo myError
If Not loctable.EOF Then
    TransCount = loctable!myCount
End If
Finally:
loctable.Close
Set loctable = Nothing
Exit Function
On Error GoTo myError:
myError:
TransCount = -1
Err.Clear
GoTo Finally
End Function
Public Function clearTemp(con As adodb.Connection)
con.Execute " If(OBJECT_ID('tempdb..#temp') Is Not Null)" & _
           " Begin" & _
           " Drop Table #Temp" & _
           " End"
End Function
Function createRs(pTable As adodb.Recordset) As adodb.Recordset
Set createRs = New adodb.Recordset
createRs.ActiveConnection = Nothing
createRs.CursorLocation = adUseClient ' Or adUseServer if needed
createRs.LockType = adLockBatchOptimistic ' Or other suitable lock type
For I = 0 To pTable.Fields.Count - 1
    If pTable.Fields(I).Type = 131 Then
        createRs.Fields.Append pTable.Fields(I).Name, adDecimal, pTable.Fields(I).NumericScale, adFldIsNullable
        'createRs.Fields(i).NumericScale = pTable.Fields(i).NumericScale
    Else
        createRs.Fields.Append pTable.Fields(I).Name, pTable.Fields(I).Type, pTable.Fields(I).DefinedSize, adFldIsNullable + adFldNegativeScale
    End If
Next
createRs.Open
Do Until pTable.EOF
    createRs.AddNew ' Add a new record
    For I = 0 To pTable.Fields.Count - 1
        If createRs.Fields(I).Type = adBoolean Then
            createRs.Fields(I).Value = IIf(pTable.Fields(I).Value, 1, 0)
       Else
           createRs.Fields(I).Value = pTable.Fields(I).Value
        End If
    Next
    createRs.Update
    pTable.MoveNext
Loop
End Function
Public Function rsField(rsTable As adodb.Recordset, pValue As Variant, pReturn As String, Optional pField As String = "code") As Variant
If Trim(pValue & "") = "" Then
    rsField = Null
    Exit Function
End If

rsTable.Find pField & " = " & (pValue & ""), , adSearchForward, adBookmarkFirst
If Not rsTable.EOF Then
     rsField = rsTable(pReturn)
End If
End Function


