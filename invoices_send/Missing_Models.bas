Attribute VB_Name = "mdlGet_Missing_invoices"
Dim missingUrl As String
Dim ownerName As String
Dim tenantName As String
Dim agentToken As String
Public Function GetMissing(con As ADODB.Connection, myForm As InvoiceTransfrm) As String
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim sbJson As New ChilkatStringBuilder
Dim jArray As ChilkatJsonArray
Dim Succees As Integer

myForm.AddNote "”Ì „  ÕœÌœ «·«Ì«„ »œÊ‰ „»Ì⁄« "

If Not myDefine(con, myForm) Then Exit Function

If Not createRest(rest, url) Then
    myForm.AddNote "„‘ﬂ·… ›Ï «·œŒÊ· «·Ì —«»ÿ  ÿ»Ìﬁ «—”«· »Ì«‰«  «·›—⁄"
    Exit Function
End If

GetMissing = rest.FullRequestNoBody("GET", url.Path)

If (rest.LastMethodSuccess <> 1) Then
    myForm.AddNote rest.LastErrorText, vbRed
ElseIf rest.ResponseStatusCode >= 400 Then
    myForm.AddNote GetMissing & "-" & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")", vbRed
ElseIf rest.ResponseStatusCode <> 200 Then
    myForm.AddNote rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")", vbRed
End If
End Function
Private Function myDefine(con As ADODB.Connection, myForm As InvoiceTransfrm) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select * From SettingInvoiceSend", con)

If loctable.EOF Then
    MsgBox Err.Description
    myForm.AddNote "„·› ÷»ÿ »Ì«‰«  «· ÿ»Ìﬁ »œÊ‰ »Ì«‰« ", vbRed
    Exit Function
End If

If IsNull(loctable!missingUrl) Or IsNull(loctable!agentToken) Then
    myForm.AddNote "„·› ÷»ÿ »Ì«‰«  «·„·›«  «·„›ﬁÊœ…", vbRed
    Exit Function
End If

missingUrl = loctable!missingUrl
ownerName = loctable!ownerName
tenantName = loctable!tenantName
agentToken = loctable!agentToken

myDefine = True
End Function
Private Function createRest(rest As ChilkatRest, url As ChilkatUrl) As Boolean
Dim sucess As Integer
success = url.ParseUrl(missingUrl)
success = rest.AddHeader("Content-Type", "application/json;charset=UTF-8")
success = rest.AddHeader("Accept", "application/json")
success = rest.AddHeader("Access-Control-Allow-Headers", "Authorization")
success = rest.AddHeader("AgentToken", agentToken)

Dim bAutoReconnect As Long
bAutoReconnect = 1

success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)
createRest = success = 1
End Function

