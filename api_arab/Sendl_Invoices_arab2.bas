Attribute VB_Name = "mdSend_Invoices_Arab2"
Dim token As String
Dim auth As String
Dim apiKey As String
Dim baseUrl As String
Dim storeId As String
Public Function SendInvoicesArab(con As ADODB.Connection, pBranch As String, ByRef pError As String, pCount As Long) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("[dbo].[sp_invoices_branch_arab]", con, adStoredProc, AddFlag(Empty, "Branch", pBranch))

If loctable.EOF Then
    SendInvoicesArab = True
    Exit Function
End If

If Not myDefine(con, pError) Then
    Exit Function
End If

token = getTokenArab(con, pError)

Dim httpB As New ChilkatHttp
Dim reqB As New ChilkatHttpRequest

httpB.SetRequestHeader "Authorization", auth
httpB.SetRequestHeader "Content-Type", "application/json"
httpB.SetRequestHeader "Accept", "Application/json"
httpB.SetRequestHeader "x-api-key", apiKey
httpB.SetRequestHeader "x-token", token

Dim json As New ChilkatJsonObject
Dim respB As ChilkatHttpResponse
Dim sId As String
Do Until loctable.EOF
    If createLines(loctable, json) Then
        Set respB = New ChilkatHttpResponse
        Set respB = httpB.PostJson2(baseUrl & "/Transactions/Record/", "application/json", json.Emit())
        If (respB.LastMethodSuccess <> 1) Then
            pError = respB.LastErrorText & vbCrLf & _
                    respB.BodyStr
            Exit Function
        ElseIf respB.StatusCode >= 400 Then
            pError = respB.StatusCode & vbCrLf & _
                     respB.StatusText & vbCrLf & _
                     respB.LastErrorText & vbCrLf & _
                     respB.BodyStr
        ElseIf respB.StatusCode = 200 Then
            pCount = pCount + 1
            
            json.Clear
            success = json.Load(respB.BodyStr)
            sId = json.ObjectOf("result").StringOf("id")
            If json.StringOf("status") = "1" Then
                If Not json.ObjectOf("result") Is Nothing Then sId = json.ObjectOf("result").StringOf("id")
            End If
            
            con.Execute "update file6_20h set " & _
                   " posted_arab = 1 ," & _
                   " id_arab = " & addvalue(sId) & _
                   " Where doc_no = " & MyParn(loctable!doc_no)
        End If
        
    End If
    loctable.MoveNext
Loop
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
storeId = loctable!storeId
myDefine = True
End Function
Private Function createLines(loctable As ADODB.Recordset, ByRef json As ChilkatJsonObject) As Boolean
Dim success As Integer
json.Clear
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
