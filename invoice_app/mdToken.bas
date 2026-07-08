Attribute VB_Name = "inv_Token"
Dim rest As New ChilkatRest
Dim url As ChilkatUrl
Dim sbJson As New ChilkatStringBuilder
Dim jArray As ChilkatJsonArray
Dim Succees As Integer
Public Function getToken(ByRef pError As String) As String
Dim httpB As New ChilkatHttp
Dim reqB As New ChilkatHttpRequest
reqB.httpVerb = "POST"
reqB.Path = TokenUrl

reqB.contentType = "application/x-www-form-urlencoded"
reqB.AddParam "grant_type", "client_credentials"
reqB.AddParam "scope", "InvoicingAPI"
reqB.AddParam "client_id", client_id
reqB.AddParam "client_secret", client_secret

Dim respB As ChilkatHttpResponse
Set respB = httpB.PostUrlEncoded(idSrvBaseUrl & TokenUrl, reqB)
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
getToken = json.StringOf("access_token")
End Function


