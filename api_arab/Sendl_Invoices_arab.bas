Attribute VB_Name = "mdSend_Invoices_Arab"
Dim token As String
Dim auth As String
Dim apiKey As String
Dim baseUrl As String
Dim storeId As String
Public Function SendInvoicesArab(con As ADODB.Connection, pBranch As String, ByRef pError As String, pCount As Long) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("[dbo].[sp_invoices_branch_arab]", con, adStoredProc, AddFlag(Empty, "Branch", pBranch))

If loctable.EOF Then Exit Function

If Not myDefine(con, pError) Then Exit Function

Dim invTable As New ADODB.Recordset
token = getTokenArab(con, pError)

If token = "" Then
    pError = "No Token"
    Exit Function
End If

Dim url As New ChilkatUrl
Dim rest As New ChilkatRest

success = url.ParseUrl(baseUrl)

success = rest.AddHeader("Authorization", auth)
success = rest.AddHeader("Content-Type", "application/json")
success = rest.AddHeader("Accept", "application/json")
success = rest.AddHeader("x-api-key", apiKey)
success = rest.AddHeader("x-token", token)
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)

If rest.LastMethodSuccess <> 1 Then
    pError = rest.LastErrorText
    Exit Function
End If

Dim json As New ChilkatJsonObject
Dim json_resp As New ChilkatJsonObject
Dim aCount As Variant

Dim aSub As Variant
Dim aInv As Variant

On Error GoTo myerror

Do Until loctable.EOF
    If createLines(loctable, json) Then
        jsonResponseStr = rest.FullRequestString("POST", baseUrl & "/Transactions/Record", json.Emit())
        If (rest.LastMethodSuccess <> 1) Then
            pError = rest.LastErrorText & vbCrLf & _
                    jsonResponseStr
            
            Exit Function
        ElseIf rest.ResponseStatusCode >= 400 Then
            pError = rest.ResponseStatusCode & vbCrLf & _
                     rest.ResponseStatusText & vbCrLf & _
                     rest.LastErrorText & vbCrLf & _
                     jsonResponseStr
            'pError = jsonResponseStr
            'SendInvoicesArab = AddFlag(aInsert, "status", rest.ResponseStatusCode)
            'SendInvoicesArab = AddFlag(aInsert, "status", rest.ResponseStatusCode)
            'SendInvoicesArab = AddFlag(aInsert, "response", jsonResponseStr)
        ElseIf rest.ResponseStatusCode = 200 Then
            pCount = pCount + 1
            success = json_resp.Load(jsonResponseStr)
            If json_resp.StringOf("status") = "1" Then
                If Not json_resp.ObjectOf("result") Is Nothing Then
                    con.Execute "update file6_20h set " & _
                           " posted_arab = 1 ," & _
                           " id_arab = " & addvalue(json_resp.ObjectOf("result").StringOf("id")) & _
                           " Where doc_no = " & MyParn(loctable!doc_no)
                End If
            End If
        End If
    End If
    loctable.MoveNext
Loop

SendInvoicesArab = True
myerror:
End Function
Private Function myDefine(con As ADODB.Connection, ByRef sError As String) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select * From SettingInvoiceArab", con)

If loctable.EOF Then
    sError = "„·› ÷»ÿ »Ì«‰«  «· ÿ»Ìﬁ »œÊ‰ »Ì«‰« "
    Exit Function
End If

If IsNull(loctable!auth) Or IsNull(loctable!apiKey) Or IsNull(loctable!User) Or IsNull(loctable!PassWord) Or IsNull(loctable!storeId) Then
    sError = "„·› «· ÿ»Ìﬁ €Ì— „ﬂ „·…"
    Exit Function
End If

auth = loctable!auth
apiKey = loctable!apiKey
baseUrl = loctable!baseUrl
'baseUrl = "https://rsmsapi.marakez.net/webservice/RestApi/"
storeId = loctable!storeId
myDefine = True
End Function
Private Function createLines(loctable As ADODB.Recordset, ByRef json As ChilkatJsonObject) As Boolean
Dim success As Integer
success = json.AddStringAt(-1, "store_id", storeId)
success = json.AddStringAt(-1, "invoice_no", loctable!doc_no)
success = json.AddStringAt(-1, "invoice_date", Format(loctable!Date, "YYYY-MM-DD hh:mm:ss:000"))
success = json.AddStringAt(-1, "subtotal", Format(mRound(loctable!Subtotal), "Fixed"))
success = json.AddStringAt(-1, "tax", Format(mRound(loctable!tax), "Fixed"))
success = json.AddStringAt(-1, "service", "0.00")
success = json.AddStringAt(-1, "total", Format(mRound(loctable!TOTAL), "Fixed"))
success = json.AddStringAt(-1, "discount", "0.00")
createLines = True
End Function
