Attribute VB_Name = "Send_Models"
Dim vBaseUr As Variant
Dim vAgentToken As Variant
Dim vTenantName
Dim vOwnerName As Variant
Public Function SendInvoices(con As ADODB.Connection, myForm As Form, Optional bAll As Boolean = False) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select * From SettingInvoiceSend", con)

If loctable.EOF Then
    MsgBox Err.Description
    myForm.AddNote "„·› ÷»ÿ »Ì«‰«  «· ÿ»Ìﬁ »œÊ‰ »Ì«‰« ", vbRed
    Exit Function
End If

If IsNull(loctable!BaseUrl) Or IsNull(loctable!AgentToken) Or IsNull(loctable!TenantName) Or IsNull(OwnerName) Then
    myForm.AddNote "„·› ÷»ÿ »Ì«‰«  «· ÿ»Ìﬁ €Ì— „ﬂ „·…", vbRed
    Exit Function
End If


Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim Succees As Integer
Dim json As ChilkatJsonObject

Dim aJson As Variant


myForm.AddNote "”Ì „ ‰ﬁ· »Ì«‰«  «·„ÊœÌ·«  «·Ì «·„Êﬁ⁄"

success = url.ParseUrl(loctable!BaseUrl)
success = rest.AddHeader("Content-Type", "application/json;charset=UTF-8")
success = rest.AddHeader("Accept", "application/json")
success = rest.AddHeader("Access-Control-Allow-Headers", "Authorization")
success = rest.AddHeader("AgentToken", loctable!AgentToken)

Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)

Dim aParam(0) As Variant
aParam(0) = Array("ALL", adBoolean, adParamInput, -1, IIf(bAll, 1, 0))

'Set loctable = New ADODB.Recordset
'Set loctable = mycmd("GetModels", con, aParam)
If Not IsMissing(myForm) Then
    myForm.prog1.Visible = True
    myForm.prog1.Value = 0
End If

Dim jArray As New ChilkatJsonArray
Dim i As Long
Dim nRecordCount As Long
Dim sCaption As String
Dim CLIST As String

sCaption = myForm.Caption
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    If Not IsMissing(myForm) Then
        myForm.prog1.Value = Round((i + 1) / (nRecordCount), 2) * 100
        myForm.Caption = sCaption & " - Record " & (i + 1) & " from " & nRecordCount
    End If
    nLast = jArray.Size
    success = jArray.AddObjectAt(nLast)
    Set json = jArray.ObjectAt(nLast)
    success = json.UpdateString("code", loctable!CODE & "")
    success = json.UpdateString("name", loctable!Name & "")
    success = json.UpdateString("age", loctable!age & "")
    success = json.UpdateString("section", loctable!Section & "")
    success = json.UpdateString("barcode", loctable!BARCODE & "")
    success = json.UpdateString("unit", loctable!UNIT & "")
    success = json.UpdateInt("limit", mRound(loctable!Limit))
    success = json.UpdateNumber("repeat", mRound(loctable!Repeat))
    success = json.UpdateBool("show_in_app", IIf(Abs(loctable!App) = 0, 0, 1))
    
    CLIST = CLIST & IIf(CLIST = "", "", ",") & MyParn(loctable!CODE)
    
    If jArray.Size = mbSize Then
        myForm.Caption = sCaption & " - Sending Data"
        jsonResponseStr = rest.FullRequestString("POST", modelUrl, jArray.Emit)
        If (rest.LastMethodSuccess <> 1) Then
            cError = rest.LastErrorText
            Exit Do
        ElseIf rest.ResponseStatusCode >= 400 Then
            cError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")"
            Exit Do
        Else
            nCount = nCount + jArray.Size
        End If
        jArray.Clear
    End If
    loctable.MoveNext
    i = i + 1
Loop

If cError = "" Then
    If Trim(jArray.Emit()) <> "" Then
        jsonResponseStr = rest.FullRequestString("POST", modelUrl, jArray.Emit)
        If (rest.LastMethodSuccess <> 1) Then
            cError = rest.LastErrorText
        ElseIf rest.ResponseStatusCode >= 400 Then
            cError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")"
        Else
            nCount = nCount + jArray.Size
        End If
    End If
End If

If cError = "" Then
    SendModels = True
    If nCount = 0 Then
        myForm.AddNote IIf(bAll, "·« ÌÊÃœ »Ì«‰«  ··«—”«·", "·« ÌÊÃœ »Ì«‰«  ··«—”«·"), vbGreen
    Else
        myForm.AddNote " „ «—”«· »Ì«‰«  " & nCount & " „ÊœÌ· «·Ì «·„Êﬁ⁄"
    End If
Else
    MsgBox Err.Description
    myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·„ÊœÌ·«  »‰Ã«Õ", vbRed
End If

If Not IsMissing(myForm) Then
    myForm.Caption = sCaption
    myForm.prog1.Visible = False
    myForm.prog1.Value = 0
End If
End Function
Private Function sendItems(con As ADODB.Connection, myForm As Form, Optional bAll As Boolean = False) As Boolean
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim Succees As Integer
Dim json As ChilkatJsonObject
Dim cError As String

myForm.AddNote "”Ì „ ‰ﬁ· »Ì«‰«  «·«’‰«› «·Ì «·„Êﬁ⁄"


success = url.ParseUrl(mainUrl)
success = rest.AddHeader("Authorization", mbToken)
success = rest.AddHeader("Content-Type", "application/json")

Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)


Dim aParam(0) As Variant
aParam(0) = Array("ALL", adBoolean, adParamInput, -1, IIf(bAll, 1, 0))

Dim loctable As New ADODB.Recordset
Set loctable = myRecordProc("GetItems", con, aParam)

If Not IsMissing(myForm) Then
    myForm.prog1.Visible = True
End If


Dim jArray As New ChilkatJsonArray
Dim i As Long
Dim aJson As Variant
Dim nRecordCount As Long

nRecordCount = loctable.RecordCount


SetKbLayout Lang_AR

Dim sCaption As String
Dim CLIST As String
sCaption = myForm.Caption
Do Until loctable.EOF
    myForm.prog1.Value = Round((i + 1) / (nRecordCount), 2) * 100
    myForm.Caption = sCaption & " - Record " & (i + 1) & " from " & nRecordCount
    nLast = jArray.Size
        
    success = jArray.AddObjectAt(nLast)
    Set json = jArray.ObjectAt(nLast)
    
    success = json.UpdateString("model", loctable!MODEL & "")
    success = json.UpdateString("size", loctable!Size & "")
    success = json.UpdateString("code", loctable!CODE & "")
    success = json.UpdateString("model_barcode_with_size", loctable!BARCODE & "")
    success = json.UpdateNumber("price", mRound(loctable!price))
    success = json.UpdateBool("available_in_stock", loctable!in_stock)
    success = json.UpdateNumber("repeat", mRound(loctable!Repeat))
    
    CLIST = CLIST & IIf(CLIST = "", "", ",") & MyParn(loctable!CODE)
       
    If jArray.Size = mbSize Then
        myForm.Caption = sCaption & " - Sending Data"
        jsonResponseStr = rest.FullRequestString("POST", itemUrl, jArray.Emit)
         If (rest.LastMethodSuccess <> 1) Then
            cError = rest.LastErrorText
            Exit Do
        ElseIf rest.ResponseStatusCode >= 400 Then
            cError = "Error : " & rest.ResponseStatusCode & " (" & rest.ResponseStatusText & ")"
            Exit Do
        Else
            nCount = nCount + jArray.Size
            updateList CLIST, con
        End If
        jArray.Clear
    End If
    
    loctable.MoveNext
    i = i + 1
Loop

If cError = "" Then
    If Trim(jArray.Emit()) <> "" Then
        jsonResponseStr = rest.FullRequestString("POST", "/basic-data/items-bulk-create-update/", jArray.Emit)
        nCount = nCount + jArray.Size
        updateList CLIST, con
    End If
End If

If cError <> "" Then
    MsgBox Err.Description
    myForm.AddNote "·„   „ ⁄„·Ì… ‰ﬁ· «·„ÊœÌ·«  »‰Ã«Õ", vbRed
Else
    sendItems = True
    If nCount = 0 Then
        myForm.AddNote IIf(bAll, "·« ÌÊÃœ »Ì«‰«  ··«—”«·", "ﬂ· «·»Ì«‰«  „ÕœÀ…"), vbGreen
    Else
        myForm.AddNote (" „ «—”«· »Ì«‰« " & Space(1) & nCount & Space(1) & "’‰› «·Ì «·„Êﬁ⁄")
        CountSentError myForm, itemUrlShow, "«·«’‰«›"
    End If
End If

If Not IsMissing(myForm) Then
    myForm.Caption = sCaption
    myForm.prog1.Visible = False
    myForm.prog1.Value = 0
End If
End Function
End Function
Public Function Getlist(myFrom As Form, pUrl As String, ByRef cError As String, Optional pField As String = "code") As String
Dim jArray As New ChilkatJsonArray
Dim json As ChilkatJsonObject
Set jArray = GetJsonArray(mainUrl & pUrl, cError)
If cError <> "" Then Exit Function
For i = 0 To jArray.Size - 1
    Set json = jArray.ObjectAt(i)
    Getlist = Getlist & "[" & json.StringOf(pField) & "]"
Next
End Function
Public Function GetlistString(myFrom As Form, pUrl As String, ByRef cError As String, Optional pField As String = "code") As String
Dim jArray As New ChilkatJsonArray
Dim json As ChilkatJsonObject

Set jArray = GetJsonArray(mainUrl & pUrl, cError)
If cError <> "" Then Exit Function
For i = 0 To jArray.Size - 1
    Set json = jArray.ObjectAt(i)
    GetlistString = GetlistString & IIf(GetlistString = "", "", ",") & MyParn(json.StringOf(pField))
Next
End Function
Private Function updateList(ByRef CLIST As String, con As ADODB.Connection)
If Trim(CLIST) <> "" Then
    con.Execute "update file1_10 set changed = 0 where file1_10.item in(" & CLIST & ")", nCount
    CLIST = ""
End If
End Function
Private Function UpdateBalance(con As ADODB.Connection, myForm As Form, bAll As Boolean) As Boolean
Dim nCount As Long
Dim nAffect As Long
con.BeginTrans


myForm.AddNote ("”Ì „  ÕœÌÀ «—’œ… «·«’‰«›")

On Error GoTo myerror
con.Execute "exec [dbo].UpdateItemsNoBal " & IIf(bAll, "1", "0"), nAffect
nCount = nAffect

con.Execute "exec [dbo].UpdateItemsBal " & IIf(bAll, "1", "0"), nAffect
nCount = nCount + nAffect
con.CommitTrans

UpdateBalance = True

If nCount = 0 Then
    myForm.AddNote "«·«—’œ… „ÕœÀ… »«·ﬂ«„· !! ·„ Ì „ «· ⁄œÌ·", vbGreen
Else
    myForm.AddNote (" „  ⁄œÌ· «—’œ… " & nCount & " ’‰›")
End If
Exit Function
myerror:
MsgBox Err.Description
myForm.AddNote "Œÿ√ «À‰«¡ «· ÕœÌÀ ·„ Ì „ «· ÕœÌÀ »‰Ã«Õ", vbRed
Err.Clear
con.RollbackTrans
End Function


