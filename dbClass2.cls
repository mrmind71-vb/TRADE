VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public connectionString As String
Private sqlStatements() As String
' Initialize the connection
Private Sub Class_Initialize()
End Sub
Public Function myRs(strSql, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional Timeout As Integer = 3, Optional QueryTimeout As Integer = 300) As ADODB.RecordSet
    On Error GoTo ErrorHandler
            
    Dim cn As New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.ConnectionTimeout = Timeout
    cn.CommandTimeout = QueryTimeout
        
    cn.Open IIf(connectionString = "", strCon, connectionString)
            
    Set myRs = New ADODB.RecordSet
        
    
    ' ?????? ??? ???????
        
    Dim cmd As New ADODB.command
    cmd.ActiveConnection = cn
    cmd.CommandType = pType
    cmd.CommandText = strSql ' Ensure strSql is declared globally or in your module
    
    If Not IsEmpty(aParam) Then
        Dim i As Long
        For i = 0 To UBound(aParam) Step 2
            cmd.Parameters("@" & aParam(i)).value = aParam(i + 1)
        Next
    End If

    myRs.CursorLocation = 3 ' adUseClient
    Set myRs = cmd.Execute

    Set myRs.ActiveConnection = Nothing
    Set cmd.ActiveConnection = Nothing
    
    Exit Function

CleanUp:
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    Exit Function

ErrorHandler:
    MsgBox "Database Error: " & Err.Description, vbCritical
    If Not myRs Is Nothing Then
        If myRs.State = 1 Then myRs.Close ' Fixed "rs.Close" typo
        Set myRs = Nothing
    End If
    Resume CleanUp
End Function
Public Function ExecuteTransaction(Optional Timeout As Integer = 3, Optional QueryTimeout As Integer = 300) As Boolean
    On Error GoTo ErrorHandler
    
    Dim cn As New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.ConnectionTimeout = Timeout
    cn.CommandTimeout = QueryTimeout
            
    cn.Open IIf(connectionString = "", strCon, connectionString)
    
    Dim IsTransactionStarted As Boolean
       
    ' «› —«÷ «·‰Ã«Õ ›Ï «·»œ«Ì…
    ExecuteTransaction = True
    
    ' ?????? ?? ?? ??????? ?????? ???????? ?????

    ' ??? ???????? (Transaction)
    cn.BeginTrans

    Dim i As Long
    ' ?????? ??? ???? ????? ?????? ?? ???????? ????????
    For i = 1 To UBound(sqlStatements)
        ' ?????? ?? ?? ?????? ???? ?????
        If Trim$(sqlStatements(i)) <> "" Then
            cn.Execute sqlStatements(i), , adCmdText + adExecuteNoRecords
             IsTransactionStarted = True
        End If
    Next i

    ' ????? ????????? ?? ??? ???? ???? ?????
    cn.CommitTrans
    Exit Function

CleanUp:
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
    Exit Function

ErrorHandler:
    MsgBox "›‘·  «·⁄„·Ì…  „ «· —«Ã⁄ ⁄‰ ﬂ«›… «· €ÌÌ—« " & vbCrLf & _
           "«·Œÿ√ : " & Err.Description & vbCrLf & _
           "—ﬁ„ «·Ã„·… «·„”»»… ··Œÿ«: " & CStr(i), vbCritical, "Œÿ√ ›Ï ﬁ«⁄œ… «·»Ì«‰« "
     
    On Error Resume Next
    If IsTransactionStarted Then
        cn.RollbackTrans
    End If
           
    ExecuteTransaction = False
    On Error GoTo 0
    Resume CleanUp
End Function
Public Function Execute(strSql As String, Optional Timeout As Integer = 3, Optional QueryTimeout As Integer = 300) As Boolean
    On Error GoTo ErrorHandler
    
    Dim cn As New ADODB.Connection
    cn.CursorLocation = adUseClient
    cn.ConnectionTimeout = Timeout
    cn.CommandTimeout = QueryTimeout
            
    cn.Open IIf(connectionString = "", strCon, connectionString)


    ' ?????? ??? ???? ????? ?????? ?? ???????? ????????
    cn.Execute strSql, , adCmdText + adExecuteNoRecords


    Execute = True

CleanUp:
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    
    ' ????? ????????? ?? ??? ???? ???? ?????
    Exit Function

ErrorHandler:
    MsgBox "›‘·  «·⁄„·Ì… " & vbCrLf & _
           "«·Œÿ√ : " & Err.Description _
           , vbCritical, "Œÿ√ ›Ï ﬁ«⁄œ… «·»Ì«‰« "
    
Resume CleanUp
End Function
Public Function addSql(sValue As String)
ReDim Preserve sqlStatements(UBound(sqlStatements) + 1)
sqlStatements(UBound(sqlStatements)) = sValue
End Function
Public Function clearSql()
ReDim sqlStatements(0)
End Function
Function IncData(sTable As String, ByRef value As Variant, sField As String, Optional pWhere As String = "", Optional bNum As Boolean = False, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional Timeout As Integer = 3, Optional QueryTimeout As Integer = 300) As Boolean
Dim loctable As New ADODB.RecordSet
Dim cString As String
cString = "Select Max(" & sField & ") as Maxof " & _
       " FROM " & sTable
If pWhere <> "" Then cString = cString & " where " & pWhere

Set loctable = myRs(cString, pType, aParam, Timeout, QueryTimeout)
If loctable Is Nothing Then Exit Function

If Not (loctable.EOF) Then
    If IsNull(loctable!maxOf) Then
        value = "1"
    ElseIf bNum Then
        value = Val(loctable!maxOf & "") + 1
    Else
        value = IncrementString(loctable!maxOf & "")
    End If
End If
loctable.Close
Set loctable = Nothing
IncData = True
End Function
Public Function rsValue(strSql As String, ByRef value As Variant, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional Timeout As Integer = 3, Optional QueryTimeout As Integer = 300) As Boolean
Dim loctable As ADODB.RecordSet
Set loctable = myRs(strSql, pType, aParam, Timeout, QueryTimeout)
If loctable Is Nothing Then Exit Function

If Not loctable.EOF Then
    value = loctable.Fields(0).value
Else
    value = Empty
End If
Set loctable = Nothing
rsValue = True
End Function
Public Function rsvalues(strSql As String, ByRef aValues As Variant, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional Timeout As Integer = 3, Optional QueryTimeout As Integer = 300) As Boolean
Dim loctable As ADODB.RecordSet
Set loctable = myRs(strSql, pType, aParam, Timeout, QueryTimeout)
If loctable Is Nothing Then Exit Function

Dim i As Long
If Not (loctable.BOF And loctable.EOF) Then
    For i = 0 To loctable.Fields.Count - 1
        aValues = AddFlag(rsvalues, LCase(loctable.Fields(i).Name), loctable.Fields(i).value)
    Next
Else
    aValues = Empty
End If
loctable.Close
Set loctable = Nothing
rsvalues = True
End Function
Public Function strList(strSql, Optional pType As cmType = adText, Optional aParam As Variant = Empty, Optional Timeout As Integer = 3, Optional QueryTimeout As Integer = 300) As String
Dim loctable As ADODB.RecordSet
Set loctable = myRs(strSql, pType, aParam, Timeout, QueryTimeout)
If loctable Is Nothing Then Exit Function
Do Until loctable.EOF
    strList = strList & "|#" & loctable.Fields(0) & ";" & loctable.Fields(1)
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
End Function

