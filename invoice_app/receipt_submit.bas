Attribute VB_Name = "inv_submit"
Public Function CreateNoSignDoc(pDoc_No As String, con As ADODB.Connection, pError As String, pType As String, bGroup As Boolean) As String
Dim json As New ChilkatJsonObject

SetKbLayout Lang_AR

If Not CreateIssuer(con, json, pError) Then Exit Function
   
If Not CreateReceiver(pDoc_No, json, con, pError) Then Exit Function

If Not CreateDocHeader(pDoc_No, json, con, pError, pType) Then Exit Function

If Not CreateDocPayment(json, pError) Then Exit Function

If Not CreateDocDelivery(json, pError) Then Exit Function

If Not CreateDocLines(pDoc_No, json, con, pError, bGroup) Then Exit Function

If Not CreateDocFooter(pDoc_No, json, con, pError, bGroup) Then Exit Function

json.EmitCompact = 0

CreateNoSignDoc = json.Emit()
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Function CreateIssuer(con As ADODB.Connection, ByRef json As ChilkatJsonObject, ByRef pError As String) As Boolean
Dim loctable As ADODB.Recordset

On Error GoTo myerror
Set loctable = myCmd("select * from SettingCompany", con)
If loctable.EOF Then
    pError = "Invoice data user not found"
    Exit Function
End If

success = json.AddObjectAt(-1, "issuer")

Dim Juser As ChilkatJsonObject
Set Juser = json.ObjectAt(json.Size - 1)
success = Juser.AddObjectAt(-1, "address")

Dim Detail As ChilkatJsonObject
Set Detail = Juser.ObjectAt(Juser.Size - 1)

success = Detail.AddStringAt(-1, "branchID", loctable!branchId & "")
success = Detail.AddStringAt(-1, "country", loctable!Country & "")
success = Detail.AddStringAt(-1, "governate", loctable!GOVERNATE & "")
success = Detail.AddStringAt(-1, "regionCity", loctable!RegionCity & "")
success = Detail.AddStringAt(-1, "street", loctable!Street & "")
success = Detail.AddStringAt(-1, "buildingNumber", loctable!building & "")
success = Detail.AddStringAt(-1, "postalCode", loctable!Postal & "")
success = Detail.AddStringAt(-1, "floor", loctable!Floor & "")
success = Detail.AddStringAt(-1, "landmark", loctable!LandMark & "")
success = Detail.AddStringAt(-1, "additionalInformation", loctable!Additional & "")

success = Juser.AddStringAt(-1, "type", "B" & "")
success = Juser.AddStringAt(-1, "id", loctable!user_id & "")
success = Juser.AddStringAt(-1, "name", loctable!Name & "")
CreateIssuer = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Function CreateReceiver(pDoc_No As String, ByRef json As ChilkatJsonObject, con As ADODB.Connection, ByRef pError As String) As Boolean
Dim loctable As ADODB.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(pDoc_No))
Set loctable = myCmd("einv.sp_receiver", con, adStoredProc, aPrm)

On Error GoTo myerror
If loctable.EOF Then
    pError = "Invoice data receiver not found"
    Exit Function
End If

success = json.AddObjectAt(-1, "receiver")

Dim JReceiver As ChilkatJsonObject
Set JReceiver = json.ObjectAt(json.Size - 1)
success = JReceiver.AddObjectAt(-1, "address")

Dim JAddress As ChilkatJsonObject
Set JAddress = JReceiver.ObjectAt(JReceiver.Size - 1)
success = JAddress.AddStringAt(-1, "country", loctable!Country & "")
success = JAddress.AddStringAt(-1, "governate", loctable!GOVERNATE & "")
success = JAddress.AddStringAt(-1, "regionCity", loctable!RegionCity & "")
success = JAddress.AddStringAt(-1, "Street", loctable!Street & "")
success = JAddress.AddStringAt(-1, "buildingNumber", loctable!buildingNumber & "")
success = JAddress.AddStringAt(-1, "postalCode", "")
success = JAddress.AddStringAt(-1, "floor", "")
success = JAddress.AddStringAt(-1, "landmark", "")
success = JAddress.AddStringAt(-1, "additionalInformation", "")

If Not IsNull(loctable!fileNo) Then
    success = JReceiver.AddStringAt(-1, "type", "B")
    success = JReceiver.AddStringAt(-1, "id", loctable!fileNo)
Else
    success = JReceiver.AddStringAt(-1, "type", "P")
    success = JReceiver.AddStringAt(-1, "id", loctable!ID_NO & "")
End If
success = JReceiver.AddStringAt(-1, "name", loctable!Name & "")
CreateReceiver = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Function CreateDocHeader(pDoc_No As String, ByRef json As ChilkatJsonObject, con As ADODB.Connection, ByRef pError As String, pType As String) As Boolean
Dim loctable As ADODB.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(pDoc_No))
Set loctable = myCmd("einv.sp_inv_header", con, adStoredProc, aPrm)
On Error GoTo myerror
success = json.AddStringAt(-1, "documentType", pType)
success = json.AddStringAt(-1, "documentTypeVersion", inv_version)
success = json.AddStringAt(-1, "dateTimeIssued", myFormat_Z(loctable!dateTax & ""))
success = json.AddStringAt(-1, "taxpayerActivityCode", taxpayerActivityCode)
success = json.AddStringAt(-1, "internalID", loctable!inv_tax & "")
success = json.AddStringAt(-1, "purchaseOrderReference", loctable!purchaseOrderReference & "")
success = json.AddStringAt(-1, "purchaseOrderDescription", "")
success = json.AddStringAt(-1, "salesOrderReference", "")
success = json.AddStringAt(-1, "proformaInvoiceNumber", "")

CreateDocHeader = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Function CreateDocPayment(ByRef json As ChilkatJsonObject, ByRef pError As String) As Boolean
Dim success As Integer
success = json.AddObjectAt(-1, "payment")

Dim JPayment As ChilkatJsonObject
Set JPayment = json.ObjectAt(json.Size - 1)
success = JPayment.AddStringAt(-1, "bankName", "")
success = JPayment.AddStringAt(-1, "bankAddress", "")
success = JPayment.AddStringAt(-1, "bankAccountNo", "")
success = JPayment.AddStringAt(-1, "bankAccountIBAN", "")
success = JPayment.AddStringAt(-1, "swiftCode", "")
success = JPayment.AddStringAt(-1, "terms", "")
CreateDocPayment = True
End Function
Private Function CreateDocDelivery(ByRef json As ChilkatJsonObject, ByRef pError As String) As Boolean
Dim success As Integer
success = json.AddObjectAt(-1, "delivery")

Dim Jdelivery As ChilkatJsonObject
Set Jdelivery = json.ObjectAt(json.Size - 1)

success = Jdelivery.AddStringAt(-1, "approach", "")
success = Jdelivery.AddStringAt(-1, "packaging", "")
success = Jdelivery.AddStringAt(-1, "dateValidity", "")
success = Jdelivery.AddStringAt(-1, "exportPort", "")
success = Jdelivery.AddNumberAt(-1, "grossWeight", 0)
success = Jdelivery.AddNumberAt(-1, "netWeight", 0)
success = Jdelivery.AddStringAt(-1, "terms", "")
CreateDocDelivery = True
End Function
Private Function CreateDocLines(pDoc_No As String, ByRef json As ChilkatJsonObject, con As ADODB.Connection, ByRef pError As String, bGroup As Boolean) As Boolean
Dim loctable As New ADODB.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(pDoc_No))
aPrm = AddFlag(aPrm, "TAX", mRound(taxRate / 100, 2))
aPrm = AddFlag(aPrm, "GROUP", IIf(bGroup, 1, 0))

Set loctable = myCmd("einv.sp_inv_lines", con, adStoredProc, aPrm)
If loctable.EOF Then
    pError = "»Ì«‰«  «·„” ‰œ €Ì— ”„Ã·…"
    Exit Function
End If

success = json.AddArrayAt(-1, "invoiceLines")
Dim aLines As ChilkatJsonArray
Set aLines = json.ArrayAt(json.Size - 1)

Dim oLine As ChilkatJsonObject
Dim osub As ChilkatJsonObject
Do Until loctable.EOF
    success = aLines.AddObjectAt(-1)
    Set oLine = aLines.ObjectAt(aLines.Size - 1)
    
    success = oLine.AddStringAt(-1, "description", loctable!Description & "")
    success = oLine.AddStringAt(-1, "itemType", loctable!itemType)
    success = oLine.AddStringAt(-1, "itemCode", loctable!itemCode & "")
    success = oLine.AddStringAt(-1, "unitType", loctable!unitType)
    success = oLine.AddNumberAt(-1, "quantity", loctable!quantity)
    success = oLine.AddStringAt(-1, "internalCode", loctable!internalCode)
    success = oLine.AddNumberAt(-1, "salesTotal", loctable!SalesTotal)
    success = oLine.AddNumberAt(-1, "total", loctable!TOTAL)
    success = oLine.AddNumberAt(-1, "valueDifference", 0)
    success = oLine.AddNumberAt(-1, "totalTaxableFees", 0)
    success = oLine.AddNumberAt(-1, "netTotal", loctable!netTotal)
    success = oLine.AddNumberAt(-1, "itemsDiscount", 0)
    
    success = oLine.AddObjectAt(-1, "unitValue")
    Set osub = oLine.ObjectAt(oLine.Size - 1)
    success = osub.AddStringAt(-1, "currencySold", "EGP")
    success = osub.AddNumberAt(-1, "amountEGP", loctable!amountEGP)
'
    success = oLine.AddObjectAt(-1, "discount")
    Set osub = oLine.ObjectAt(oLine.Size - 1)
    success = osub.AddNumberAt(-1, "rate", loctable!Rate)
    success = osub.AddNumberAt(-1, "amount", loctable!discount)
    
    Dim aTax As ChilkatJsonArray
    success = oLine.AddArrayAt(-1, "taxableItems")
    Set aTax = oLine.ArrayAt(oLine.Size - 1)
    
    Dim oTax As ChilkatJsonObject
    success = aTax.AddObjectAt(-1)
    Set oTax = aTax.ObjectAt(aTax.Size - 1)
    
    success = oTax.AddStringAt(-1, "taxType", "T1")
    success = oTax.AddNumberAt(-1, "amount", loctable!tax)
    success = oTax.AddStringAt(-1, "subType", subType)
    success = oTax.AddNumberAt(-1, "rate", taxRate)
                            
    loctable.MoveNext
Loop
CreateDocLines = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Public Function CreateDocFooter(pDoc_No As String, ByRef json As ChilkatJsonObject, con As ADODB.Connection, ByRef pError As String, bGroup As Boolean) As Boolean
Dim cString As String
Dim loctable As New ADODB.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(pDoc_No))
aPrm = AddFlag(aPrm, "TAX", mRound(taxRate / 100, 2))
aPrm = AddFlag(aPrm, "GROUP", IIf(bGroup, 1, 0))

Set loctable = myCmd("einv.sp_inv_totals", con, adStoredProc, aPrm)
If loctable.EOF Then
    pError = "Invoice data receiver not found"
    Exit Function
End If

success = json.AddNumberAt(-1, "totalDiscountAmount", loctable!totalDiscountAmount)
success = json.AddNumberAt(-1, "totalSalesAmount", loctable!totalSalesAmount)
success = json.AddNumberAt(-1, "netAmount", loctable!netAmount)

Dim aTaxes As ChilkatJsonArray
success = json.AddArrayAt(-1, "taxTotals")
Set aTaxes = json.ArrayAt(json.Size - 1)
success = aTaxes.AddObjectAt(-1)

Dim oTax As ChilkatJsonObject
Set oTax = aTaxes.ObjectAt(aTaxes.Size - 1)
success = oTax.AddStringAt(-1, "taxType", "T1")
success = oTax.AddNumberAt(-1, "amount", loctable!tax)

success = json.AddNumberAt(-1, "totalAmount", loctable!totalAmount)
success = json.AddNumberAt(-1, "extraDiscountAmount", 0)
success = json.AddNumberAt(-1, "totalItemsDiscountAmount", 0)

CreateDocFooter = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Function CreateSignDoc(pDoc_No As String, con As ADODB.Connection, ByRef cError As String, myForm As Form, pType As String, bGroup As Boolean) As String
Dim sDocument As String
sDocument = CreateNoSignDoc(pDoc_No, con, cError, pType, bGroup)

If cError <> "" Then Exit Function

If Trim(sDocument) = "" Then
    cError = "»Ì«‰«  €Ì— ’ÕÌÕ… «Ê €Ì— „ÊÃÊœ…"
    Exit Function
End If

Clipboard.Clear
Clipboard.SetText sDocument

Dim sPath As String
sPath = App.Path & "\" & RetSetting("SIGNER_TOKEN", App.Path & "\invoice_app\conf.txt")
ShellExWait sPath, vbNullString, myForm, showSign

sDocument = Clipboard.GetText()
If Trim(sDocument) = "" Then
    cError = "»Ì«‰«  €Ì— ’ÕÌÕ… «Ê €Ì— „ÊÃÊœ…"
    Exit Function
ElseIf Mid(Trim(sDocument), 1, Len(errorFlag)) = errorFlag Then
    cError = Mid(sDocument, Len(errorFlag) + 1)
    Exit Function
End If

Dim json As New ChilkatJsonObject

success = json.Load(sDocument)

'Dim aDocs As ChilkatJsonArray
'Set aDocs = json.ArrayOf("documents")

If json.ArrayOf("documents") Is Nothing Then
    cError = "·„ Ì „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
    Exit Function
ElseIf json.ArrayOf("documents").Size = 0 Then
    cError = "·„ Ì „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
    Exit Function
End If

If json.ArrayOf("documents").ObjectAt(0).ArrayOf("signatures") Is Nothing Then
    cError = "·„ Ì „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
    Exit Function
ElseIf json.ArrayOf("documents").ObjectAt(0).ArrayOf("signatures").Size = 0 Then
    cError = "·„ Ì „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
    Exit Function
Else
    sSign = json.ArrayOf("documents").ObjectAt(0).ArrayOf("signatures").ObjectAt(0).StringOf("value")
End If

'Dim asigns As ChilkatJsonArray
'Set asigns = jDoc.ArrayOf("signatures")
'If asigns Is Nothing Then
'    cError = "·„ Ì „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
'    Exit Function
'End If
'
'Dim oSign As ChilkatJsonObject
'Set oSign = asigns.ObjectAt(0)
'sSign = oSign.StringOf("value")

If sSign = "" Then
    cError = "·„ Ì „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
    Exit Function
Else
    eInform " „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
End If
CreateSignDoc = sDocument
End Function
Public Function SubmitDocument(pDoc_No As String, con As ADODB.Connection, ByRef pError As String, myForm As Form, pType As String, bGroup As Boolean) As Boolean
If Not validSend(pDoc_No, con, pType) Then Exit Function

If Not DefineUrl(con) Then Exit Function

Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

Dim sToken As String
sToken = getToken(pError)
If pError <> "" Then Exit Function
    
If sToken = "" Then
    pError = "Empty Token"
    Exit Function
End If
       
Clipboard.Clear
Clipboard.SetText sToken

Dim sDocument As String
sDocument = CreateSignDoc(pDoc_No, con, pError, myForm, pType, bGroup)

If pError <> "" Then
    Exit Function
End If

If sDocument = "" Then
    pError = "·„ Ì „ «—”«· »Ì«‰«  „‰ «·„Êﬁ⁄"
    Exit Function
End If


Dim sResponse  As String, sStatusCode As String
sResponse = SubmitDocumentString(sDocument, sToken, pError, sStatusCode)
   
If pError <> "" Then
    Exit Function
End If

Inform sStatusCode

If sStatusCode >= 400 Then
    MsgBox sResponse
    Exit Function
End If
    
Dim json As New ChilkatJsonObject
success = json.Load(sResponse)

Dim aDocs As New ChilkatJsonArray

If json.ArrayOf("rejectedDocuments").Size > 0 Then
     If Not json.ArrayOf("rejectedDocuments").ObjectAt(0).ObjectOf("error") Is Nothing Then
        pError = json.ArrayOf("rejectedDocuments").ObjectAt(0).ObjectOf("error").StringOf("message")
    End If
End If

Dim aInsert As Variant
con.BeginTrans
If Not json.ArrayOf("acceptedDocuments") Is Nothing Then
    If json.ArrayOf("acceptedDocuments").Size > 0 Then
        
        sUuid = json.ArrayOf("acceptedDocuments").ObjectAt(0).StringOf("uuid")
        aInsert = AddFlag(aInsert, "uuid", addstring(sUuid))
        aInsert = AddFlag(aInsert, "inv_tax", addstring(sUuid))
        aInsert = AddFlag(aInsert, "DATE_UUID", addDate(Date))
        If pType = "I" Then
            con.Execute addUpdate(aInsert, "FILE6_20H", "DOC_NO = " & MyParn(pDoc_No))
        ElseIf pType = "C" Then
            con.Execute addUpdate(aInsert, "FILE6_20H", "DOC_NO = " & MyParn(pDoc_No))
        End If
        SubmitDocument = True
    End If
End If

aInsert = AddFlag(Empty, "doc_no", addstring(pDoc_No))
aInsert = AddFlag(aInsert, "uuid", addstring(sUuid))
aInsert = AddFlag(aInsert, "document_body", addstring(sDocument))
aInsert = AddFlag(aInsert, "token", addstring(sToken))
aInsert = AddFlag(aInsert, "DATE", "GetDate()")
aInsert = AddFlag(aInsert, "RESPONSE", addstring(sResponse))
aInsert = AddFlag(aInsert, "StatusCode", addstring(sStatusCode))
aInsert = AddFlag(aInsert, "ERROR", addstring(pError))
If pType = "I" Then
    con.Execute addInsert(aInsert, "FILE6_20E")
ElseIf pType = "C" Then
    con.Execute addInsert(aInsert, "FILE6_20E")
End If
con.CommitTrans

Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function
Private Function SubmitDocumentString(pDocument As String, pToken As String, ByRef pError As String, ByRef pRespCode As String) As String
Dim rest As New ChilkatRest
Dim url As New ChilkatUrl
Dim Succees As Integer
Dim json As ChilkatJsonObject

success = url.ParseUrl(apiBaseUrl)
success = rest.AddHeader("Authorization", "Bearer " & pToken)
success = rest.AddHeader("Content-Type", "application/json")
'success = rest.AddHeader("Accept-Encoding", "gzip, deflate, br")
Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)
jsonResponseStr = rest.FullRequestString("POST", SubmitUrl, pDocument)

pRespCode = rest.ResponseStatusCode

If (rest.LastMethodSuccess <> 1) Then
    pError = rest.LastErrorText & vbCrLf & jsonResponseStr
End If
SubmitDocumentString = jsonResponseStr
End Function
Private Function validSend(pDoc_No As String, con As ADODB.Connection, pType As String) As Boolean
Dim loctable As ADODB.Recordset
aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(pDoc_No))
Set loctable = myCmd("einv.sp_valid_send", con, adStoredProc, aPrm)

If loctable.EOF Then
    MsgBox "·«  ÊÃœ „” ‰œ »Â–« «·—ﬁ„"
    Exit Function
End If

If IsNull(loctable!inv_tax) Then
    MsgBox "—ﬁ„ «·„” ‰œ «·÷—Ì»Ì €Ì— „”Ã·"
    Exit Function
End If

If IsNull(loctable!dateTax) Then
    MsgBox " «—ÌŒ «·„” ‰œ «·÷—Ì»Ì €Ì— „”Ã·"
    Exit Function
End If

If IsNull(loctable!Name) Then
    MsgBox "≈”„ «·⁄„Ì· «·÷—Ì»Ì €Ì— „”Ã·"
    Exit Function
End If

If IsNull(loctable!fileNo) And IsNull(loctable!ID_NO) Then
    MsgBox "·« ÌÊÃœ —ﬁ„ „·› ÷—Ì»Ì «Ê —ﬁ„ ÂÊÌ…"
    Exit Function
End If

If IsNull(loctable!GOVERNATE) Then
    MsgBox "„Õ«›Ÿ… «·⁄„Ì· €Ì— „”Ã·…"
    Exit Function
End If

validSend = True
End Function
