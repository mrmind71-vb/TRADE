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
Public Function mycmd(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 300) As ADODB.RecordSet
Dim locTable As New ADODB.RecordSet
Dim cmd As New ADODB.command
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
Set mycmd = cmd.Execute
End Function
Public Function myCmdRecord(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 300) As ADODB.Record
Dim locTable As New ADODB.RecordSet
Dim cmd As New ADODB.command
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

Set myCmdRecord = cmd.Execute
End Function

Public Function myPrcString(pProc As String, Optional aParam As Variant = Empty) As String
Dim i As Long

If Not IsEmpty(aParam) Then
    For i = 0 To UBound(aParam) Step 2
        myPrcString = myPrcString & IIf(myPrcString = "", " ", ",") & "@" & aParam(i) & " = " & aParam(i + 1)
    Next
End If
myPrcString = "EXEC " & pProc & myPrcString
End Function
Public Function mycmdEx(pString As String, Optional con As ADODB.Connection, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 600) As ADODB.command
Dim locTable As New ADODB.RecordSet
Dim cmd As New ADODB.command
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
Public Function myCommand(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 300) As ADODB.command
Dim locTable As New ADODB.RecordSet
Dim cmd As New ADODB.command
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
cmd.Execute
Set myCommand = cmd
End Function
Public Function cmd(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 1000) As ADODB.command
Set cmd = New ADODB.command
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
End Function
Public Function myField(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional pDef As Variant = Empty, Optional nTimeOut As Integer = 100) As Variant
On Error GoTo myError:
Dim locTable As ADODB.RecordSet
Set locTable = cmd(pString, con, pType).Execute
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
If Not locTable.EOF Then
    myField = locTable.Fields(0).Value
ElseIf Not IsEmpty(pDef) Then
    myField = pDef
End If
locTable.Close
Finally:
Set locTable = Nothing
Exit Function
myError:
MsgBox Err.Description
'Err.Clear
myField = pDef
GoTo Finally
End Function
Public Function myFieldValue(pString As String, pField As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional pDef As Variant = Empty, Optional nTimeOut As Integer = 100) As Variant
On Error GoTo myError:
Dim cmd As New ADODB.command
Dim locTable As ADODB.RecordSet
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
Set locTable = cmd.Execute
If Not locTable.EOF Then
    myFieldValue = locTable.Fields(pField).Value
ElseIf Not IsEmpty(pDef) Then
    myFieldValue = pDef
End If
Exit Function
myError:
MsgBox Err.Description
Err.Clear
myFieldValue = pDef
End Function
Public Function myFields(pString As String, Optional con As ADODB.Connection, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional nTimeOut As Integer = 100) As Variant
Dim cmd As New ADODB.command
Dim locTable As ADODB.RecordSet
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
Set locTable = cmd.Execute
If Not (locTable.BOF And locTable.EOF) Then
    For i = 0 To locTable.Fields.Count - 1
        myFields = AddFlag(myFields, LCase(locTable.Fields(i).Name), locTable.Fields(i).Value)
    Next
End If
locTable.Close
Set locTable = Nothing
End Function
Public Function openConHr(ByRef pCon As ADODB.Connection, Optional ByVal pString As String = "", Optional ByVal lMsg As Boolean = True) As String
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
Function TransCount(con As ADODB.Connection) As Integer
Dim locTable As New ADODB.RecordSet
Set locTable = mycmd("select @@TRANCOUNT as myCount", con)
On Error GoTo myError
If Not locTable.EOF Then
    TransCount = locTable!myCount
End If
Finally:
locTable.Close
Set locTable = Nothing
Exit Function
On Error GoTo myError:
myError:
TransCount = -1
Err.Clear
GoTo Finally
End Function
Public Function clearTemp(con As ADODB.Connection)
con.Execute " If(OBJECT_ID('tempdb..#temp') Is Not Null)" & _
           " Begin" & _
           " Drop Table #Temp" & _
           " End"
End Function
Function createRs(pTable As ADODB.RecordSet) As ADODB.RecordSet
Set createRs = New ADODB.RecordSet
createRs.ActiveConnection = Nothing
createRs.CursorLocation = adUseClient ' Or adUseServer if needed
createRs.LockType = adLockBatchOptimistic ' Or other suitable lock type
For i = 0 To pTable.Fields.Count - 1
    If pTable.Fields(i).Type = 131 Then
        createRs.Fields.Append pTable.Fields(i).Name, adDecimal, pTable.Fields(i).NumericScale, adFldIsNullable
        'createRs.Fields(i).NumericScale = pTable.Fields(i).NumericScale
    Else
        createRs.Fields.Append pTable.Fields(i).Name, pTable.Fields(i).Type, pTable.Fields(i).DefinedSize, adFldIsNullable + adFldNegativeScale
    End If
Next
createRs.Open
Do Until pTable.EOF
    createRs.AddNew ' Add a new record
    For i = 0 To pTable.Fields.Count - 1
        If createRs.Fields(i).Type = adBoolean Then
            createRs.Fields(i).Value = IIf(pTable.Fields(i).Value, 1, 0)
       Else
           createRs.Fields(i).Value = pTable.Fields(i).Value
        End If
    Next
    createRs.Update
    pTable.MoveNext
Loop
End Function
Public Function rsField(rsTable As ADODB.RecordSet, pValue As Variant, pReturn As String, Optional pField As String = "code") As Variant
If Trim(pValue & "") = "" Then
    rsField = Null
    Exit Function
End If

rsTable.Find pField & " = " & (pValue & ""), , adSearchForward, adBookmarkFirst
If Not rsTable.EOF Then
     rsField = rsTable(pReturn)
End If
End Function


