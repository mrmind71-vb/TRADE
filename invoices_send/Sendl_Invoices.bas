Attribute VB_Name = "mdSend_Invoices"
Dim baseUrl As String
Dim ownerName As String
Dim tenantName As String
Dim agentToken As String
Public Function SendInvoices(con As ADODB.Connection, myForm As Form, pDate As String, jArray As ChilkatJsonArray, ByRef jsonResponseStr As String) As Boolean
Dim nCount As Long
If Not myDefine(con) Then Exit Function

Dim Succees As Integer
If bMissed Then
    myForm.AddNote "”Ì „  ÕœÌœ «·«Ì«„ »œÊ‰ „»Ì⁄« "
Else
    myForm.AddNote "”Ì „ ‰ﬁ· ›Ê« Ì— ÌÊ„ " & myFormat_p(pDate) & "  ÿ»Ìﬁ „Ê· «·›—⁄"
End If

Dim url As New ChilkatUrl
Dim rest As New ChilkatRest
If Not createRest(rest, url) Then Exit Function

If Not IsMissing(myForm) Then
    myForm.prog1.Visible = True
    myForm.prog1.Value = 0
End If

Dim aPrm As Variant
aPrm = AddFlag(Empty, "Branch", "019")
aPrm = AddFlag(aPrm, "DATE", myFormat(pDate))

Dim loctable As New ADODB.Recordset
Set loctable = myCmd("[dbo].[sp_invoices_branch_total]", con, adStoredProc, aPrm)

If loctable.EOF Then
    nCount = 0
Else
    nCount = loctable.RecordCount
End If

Dim json As New ChilkatJsonObject
If Not createHeader(pDate, nCount, json) Then Exit Function
If Not createLines(pDate, loctable, json, myForm) Then Exit Function


'Clipboard.Clear
'Clipboard.SetText json.Emit()

jsonResponseStr = rest.FullRequestString("POST", "/trackerMW/DSPushApi/SalesReport", json.Emit())
If (rest.LastMethodSuccess <> 1) Then
    cError = rest.LastErrorText
ElseIf rest.ResponseStatusCode >= 400 Then
    cError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")"
Else
    Set json = New ChilkatJsonObject
    json.Load jsonResponseStr
    
    Set jArray = json.ArrayOf("missingReportDates")
End If
    
If cError = "" Then
    SendInvoices = True
    If nCount = 0 Then
        myForm.AddNote "·« ÌÊÃœ »Ì«‰«  ··«—”«·", vbGreen
    Else
        myForm.AddNote " „ «—”«· »Ì«‰«  " & nCount & " ›« Ê—… «·Ì «· ÿ»Ìﬁ"
    End If
Else
    MsgBox Err.Description
    myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «· ÿ»Ìﬁ«  »‰Ã«Õ", vbRed
End If


If Not IsMissing(myForm) Then
    myForm.Caption = sCaption
    myForm.prog1.Visible = False
    myForm.prog1.Value = 0
End If
End Function
Private Function createHeader(pDate As String, nCount As Long, ByRef json As ChilkatJsonObject) As Boolean
Dim success As Integer
success = json.AddStringAt(-1, "reportType", "Daily")
success = json.AddStringAt(-1, "requestId", createUUID)
success = json.AddStringAt(-1, "ownerName", ownerName)
success = json.AddStringAt(-1, "tenantName", tenantName)
success = json.AddStringAt(-1, "reportDate", myFormat(pDate))
success = json.AddStringAt(-1, "requestDateTime", Format(Now, "yyyy-mm-dd hh:nn:ss"))
success = json.AddStringAt(-1, "transactionCount", nCount)
createHeader = success = 1
End Function
Private Function myDefine(con As ADODB.Connection) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select * From SettingInvoiceSend", con)

If loctable.EOF Then
    MsgBox Err.Description
    myForm.AddNote "„·› ÷»ÿ »Ì«‰«  «· ÿ»Ìﬁ »œÊ‰ »Ì«‰« ", vbRed
    Exit Function
End If

If IsNull(loctable!baseUrl) Or IsNull(loctable!agentToken) Or IsNull(loctable!tenantName) Or IsNull(loctable!ownerName) Then
    myForm.AddNote "„·› ÷»ÿ »Ì«‰«  «· ÿ»Ìﬁ €Ì— „ﬂ „·…", vbRed
    Exit Function
End If

baseUrl = loctable!baseUrl
ownerName = loctable!ownerName
tenantName = loctable!tenantName
agentToken = loctable!agentToken

myDefine = True
End Function
Private Function createRest(rest As ChilkatRest, url As ChilkatUrl) As Boolean
Dim sucess As Integer
success = url.ParseUrl(baseUrl)
success = rest.AddHeader("Content-Type", "application/json;charset=UTF-8")
success = rest.AddHeader("Accept", "application/json")
success = rest.AddHeader("Access-Control-Allow-Headers", "Authorization")
success = rest.AddHeader("AgentToken", agentToken)

Dim bAutoReconnect As Long
bAutoReconnect = 1

success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)
createRest = success = 1
End Function
Private Function createLines(pDate As String, loctable As ADODB.Recordset, ByRef json As ChilkatJsonObject, myForm As Form) As Boolean
Dim success As Integer
success = json.AddArrayAt(-1, "salesTransactions")

Dim aLines As ChilkatJsonArray
Set aLines = json.ArrayAt(json.Size - 1)

Dim nRecordCount As Long, i As Long
nRecordCount = loctable.RecordCount
sCaption = myForm.Caption

Do Until loctable.EOF
    i = i + 1
    If Not IsMissing(myForm) Then
        myForm.prog1.Value = Round((i) / (nRecordCount), 2) * 100
        myForm.Caption = sCaption & " - Record " & (i + 1) & " from " & nRecordCount
    End If
    
    success = aLines.AddObjectAt(-1)
    Set oLine = aLines.ObjectAt(aLines.Size - 1)
    
    success = oLine.AddStringAt(-1, "tenantTransId", loctable!doc_no)
    success = oLine.AddStringAt(-1, "currencyCode", "EGP")
    success = oLine.AddStringAt(-1, "gross", Format(mRound(loctable!total_net), "Fixed"))
    success = oLine.AddStringAt(-1, "net", Format(mRound(loctable!total_net), "Fixed"))
    success = oLine.AddStringAt(-1, "serviceCharge", "0.00")
    success = oLine.AddStringAt(-1, "discount", "0.00")
    success = oLine.AddStringAt(-1, "tax", "0.00")
    success = oLine.AddStringAt(-1, "transDate", Format(pDate, "YYYY-MM-DDThh:mm:ss"))
    loctable.MoveNext
Loop
createLines = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
myEnd:
myForm.Caption = sCaption
End Function
Public Function createUUID()
Dim crypt As New ChilkatCrypt2
createUUID = crypt.GenerateUuid()
End Function
Public Function myFunc(pFunction As String, Optional pParam1 As String = "", Optional pParam2 As String = "", Optional pParam3 As String = "", Optional pParam4 As String = "", Optional pParam5 As String = "", Optional pParam6 As String = "", Optional pParam7 As String = "", Optional pParam8 As String = "", Optional pParam9 As String = "", Optional pParam10 As String = "") As String
If pParam1 <> "" Then myFunc = pParam1
If pParam2 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam2
If pParam3 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam3
If pParam4 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam4
If pParam5 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam5
If pParam6 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam6
If pParam7 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam7
If pParam8 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam8
If pParam9 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam9
If pParam10 <> "" Then myFunc = myFunc & IIf(myFunc = "", "", ",") & pParam10
myFunc = pFunction & "(" & myFunc & ")"
End Function
Public Function MyFuncValue(pFunction As String, pCon As ADODB.Connection, Optional pParam1 As String = "", Optional pParam2 As String = "", Optional pParam3 As String = "", Optional pParam4 As String = "", Optional pParam5 As String = "", Optional pParam6 As String = "", Optional pParam7 As String = "", Optional pParam8 As String = "", Optional pParam9 As String = "", Optional pParam10 As String = "") As Variant
Dim cPrm As String, cString As String
If pParam1 <> "" Then cPrm = pParam1
If pParam2 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam2
If pParam3 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam3
If pParam4 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam4
If pParam5 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam5
If pParam6 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam6
If pParam7 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam7
If pParam8 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam8
If pParam9 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam9
If pParam10 <> "" Then cPrm = cPrm & IIf(cPrm = "", "", ",") & pParam10

cString = "Select " & pFunction & "(" & cPrm & ") as [value]"

Dim loctable As New ADODB.Recordset
Set loctable = myCmd(cString, pCon)
MyFuncValue = loctable!Value
loctable.Close
Set loctable = Nothing
End Function

