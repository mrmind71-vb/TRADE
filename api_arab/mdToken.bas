Attribute VB_Name = "inv_Token_arab"
Dim rest As New ChilkatRest
Dim url As ChilkatUrl
Dim sbJson As New ChilkatStringBuilder
Dim jArray As ChilkatJsonArray
Dim Succees As Integer
Public Function getTokenArab(con As ADODB.Connection, ByRef pError As String) As String
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("SettingInvoiceArab", con, adTable)

If loctable.EOF Then
    pError = "„·› ÷»ÿ »Ì«‰«  «· ÊﬂÌ‰ €Ì— „ÊÃÊœ"
    Exit Function
End If

Dim httpB As New ChilkatHttp
Dim reqB As New ChilkatHttpRequest


reqB.httpVerb = "POST"
reqB.Path = "Users/Login"

reqB.AddHeader "Authorization", loctable!auth & ""
reqB.AddHeader "Content-Type", "application/x-www-form-urlencoded"
reqB.AddHeader "Accept", "Application/json"
reqB.AddHeader "contentType", "Application/json"
reqB.AddHeader "x-api-key", loctable!apiKey & ""
reqB.AddParam "userName", loctable!UserAccess & ""
reqB.AddParam "password", loctable!UserPassWord & "" '

Dim respB As ChilkatHttpResponse
Set respB = httpB.PostUrlEncoded(loctable!baseUrl & reqB.Path, reqB)
If (httpB.LastMethodSuccess = 0) Then
    pError = httpB.LastErrorText
    Exit Function
End If

Dim json As New ChilkatJsonObject
success = json.Load(respB.BodyStr)
If (success <> 1) Then
    pError = "Error : " & json.LastErrorText
    Exit Function
End If
If Not json.ObjectOf("result") Is Nothing Then
    getTokenArab = json.ObjectOf("result").StringOf("token")
End If
End Function

