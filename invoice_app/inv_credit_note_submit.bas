Attribute VB_Name = "inv_credit_note_submit"
Public Function CreateNoSignDoc(pDoc_No As String, con As ADODB.Connection, pError As String) As String
Dim Json As New ChilkatJsonObject
SetKbLayout Lang_AR

If Not CreateIssuer(con, Json, pError) Then Exit Function
   
If Not CreateReceiver(pDoc_No, Json, con, pError) Then Exit Function

If Not CreateDocHeader(pDoc_No, Json, con, pError) Then Exit Function

If Not CreateDocPayment(Json, pError) Then Exit Function

If Not CreateDocDelivery(Json, pError) Then Exit Function

If Not CreateDocLines(pDoc_No, Json, con, pError) Then Exit Function

If Not CreateDocFooter(pDoc_No, Json, con, pError) Then Exit Function
Json.EmitCompact = 0
CreateNoSignDoc = Json.Emit()
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Function CreateIssuer(con As ADODB.Connection, ByRef Json As ChilkatJsonObject, ByRef pError As String) As Boolean
Dim loctable As ADODB.Recordset

On Error GoTo myerror
Set loctable = myCmd("select * from SettingCompany", con)
If loctable.EOF Then
    pError = "Invoice data user not found"
    Exit Function
End If

success = Json.AddObjectAt(-1, "issuer")

Dim Juser As ChilkatJsonObject
Set Juser = Json.ObjectAt(Json.Size - 1)
success = Juser.AddObjectAt(-1, "address")

Dim Detail As ChilkatJsonObject
Set Detail = Juser.ObjectAt(Juser.Size - 1)

success = Detail.AddStringAt(-1, "branchID", loctable!branchId & "")
success = Detail.AddStringAt(-1, "country", loctable!Country & "")
success = Detail.AddStringAt(-1, "governate", loctable!Governate & "")
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
Private Function CreateReceiver(pDoc_No As String, ByRef Json As ChilkatJsonObject, con As ADODB.Connection, ByRef pError As String) As Boolean
Dim loctable As ADODB.Recordset
Dim cString As String
cString = "SELECT  FILE3_10.Country" & _
          ",code0_20.DESCA AS governate" & _
        ",Case when CODE0_10.DESCA is null then 'No Region' else CODE0_10.DESCA END AS regionCity" & _
        ",CASE WHEN FILE3_10.[ADDRESS] IS NULL THEN 'No Address' ELSE FILE3_10.[ADDRESS] END AS street" & _
        ",CASE WHEN FILE3_10.BUILDING IS NULL THEN 'No Building Number' ELSE FILE3_10.BUILDING END AS buildingNumber" & _
        ",FILE3_10.[TypeTax]" & _
        ",FILE3_10.FileNo" & _
        " ,FILE3_10.id_no" & _
        ",FILE3_10.Desca as Name" & _
        " FROM  FILE3_10 INNER JOIN FILE6_10H ON FILE6_10H.CODE = FILE3_10.CODE" & _
        " LEFT JOIN code0_10 ON FILE3_10.area = code0_10.CODE " & _
        " LEFT OUTER JOIN code0_20 ON FILE3_10.gov = code0_20.CODE" & _
        " WHERE FILE6_10H.DOC_NO = " & MyParn(pDoc_No)

On Error GoTo myerror
Set loctable = myCmd(cString, con)
If loctable.EOF Then
    pError = "Invoice data receiver not found"
    Exit Function
End If

success = Json.AddObjectAt(-1, "receiver")

Dim JReceiver As ChilkatJsonObject
Set JReceiver = Json.ObjectAt(Json.Size - 1)
success = JReceiver.AddObjectAt(-1, "address")

Dim JAddress As ChilkatJsonObject
Set JAddress = JReceiver.ObjectAt(JReceiver.Size - 1)
success = JAddress.AddStringAt(-1, "country", loctable!Country & "")
success = JAddress.AddStringAt(-1, "governate", loctable!Governate & "")
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
Private Function CreateDocHeader(pDoc_No As String, ByRef Json As ChilkatJsonObject, con As ADODB.Connection, ByRef pError As String) As Boolean
Dim cString As String
cString = "SELECT DOC_NO " & _
          ",FILE6_10H.[time]" & _
          ",FILE6_10H.documentType" & _
          ",FILE6_10H.inv_tax" & _
          ",FILE6_10H.dateTax" & _
          " From FILE6_10H " & _
          " WHERE FILE6_10H.DOC_NO = " & MyParn(pDoc_No)
          
On Error GoTo myerror
Set loctable = myCmd(cString, con)
If loctable.EOF Then
    pError = "Invoice data header not found"
    Exit Function
End If

success = Json.AddStringAt(-1, "documentType", "I")
success = Json.AddStringAt(-1, "documentTypeVersion", inv_version)
success = Json.AddStringAt(-1, "dateTimeIssued", myFormat_Z(loctable!dateTax & ""))
success = Json.AddStringAt(-1, "taxpayerActivityCode", taxpayerActivityCode)
success = Json.AddStringAt(-1, "internalID", loctable!inv_tax & "")
success = Json.AddStringAt(-1, "purchaseOrderReference", "")
success = Json.AddStringAt(-1, "purchaseOrderDescription", "")
success = Json.AddStringAt(-1, "salesOrderReference", "")
success = Json.AddStringAt(-1, "proformaInvoiceNumber", "")

CreateDocHeader = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Function CreateDocPayment(ByRef Json As ChilkatJsonObject, ByRef pError As String) As Boolean
Dim success As Integer
success = Json.AddObjectAt(-1, "payment")

Dim JPayment As ChilkatJsonObject
Set JPayment = Json.ObjectAt(Json.Size - 1)
success = JPayment.AddStringAt(-1, "bankName", "")
success = JPayment.AddStringAt(-1, "bankAddress", "")
success = JPayment.AddStringAt(-1, "bankAccountNo", "")
success = JPayment.AddStringAt(-1, "bankAccountIBAN", "")
success = JPayment.AddStringAt(-1, "swiftCode", "")
success = JPayment.AddStringAt(-1, "terms", "")
CreateDocPayment = True
End Function
Private Function CreateDocDelivery(ByRef Json As ChilkatJsonObject, ByRef pError As String) As Boolean
Dim success As Integer
success = Json.AddObjectAt(-1, "delivery")

Dim Jdelivery As ChilkatJsonObject
Set Jdelivery = Json.ObjectAt(Json.Size - 1)

success = Jdelivery.AddStringAt(-1, "approach", "")
success = Jdelivery.AddStringAt(-1, "packaging", "")
success = Jdelivery.AddStringAt(-1, "dateValidity", "")
success = Jdelivery.AddStringAt(-1, "exportPort", "")
success = Jdelivery.AddNumberAt(-1, "grossWeight", 0)
success = Jdelivery.AddNumberAt(-1, "netWeight", 0)
success = Jdelivery.AddStringAt(-1, "terms", "")
CreateDocDelivery = True
End Function
Private Function CreateDocLines(pDoc_No As String, ByRef Json As ChilkatJsonObject, con As ADODB.Connection, ByRef pError As String) As Boolean
Dim loctable As New ADODB.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(pDoc_No))
aPrm = AddFlag(aPrm, "TAX", mRound(taxRate / 100, 2))
Set loctable = myCmd("dbo.sp_inv_lines", con, adStoredProc, aPrm)
If loctable.EOF Then
    pError = "Invoice data items not found"
    Exit Function
End If

success = Json.AddArrayAt(-1, "invoiceLines")
Dim aLines As ChilkatJsonArray
Set aLines = Json.ArrayAt(Json.Size - 1)

Dim oLine As ChilkatJsonObject
Dim osub As ChilkatJsonObject
Do Until loctable.EOF
    success = aLines.AddObjectAt(-1)
    Set oLine = aLines.ObjectAt(aLines.Size - 1)
    
    success = oLine.AddStringAt(-1, "description", loctable!Description & "")
    success = oLine.AddStringAt(-1, "itemType", loctable!itemType)
    success = oLine.AddStringAt(-1, "itemCode", loctable!itemCode)
    success = oLine.AddStringAt(-1, "unitType", loctable!unitType)
    success = oLine.AddNumberAt(-1, "quantity", loctable!quantity)
    success = oLine.AddStringAt(-1, "internalCode", loctable!internalCode)
    success = oLine.AddNumberAt(-1, "salesTotal", loctable!salestotal)
    success = oLine.AddNumberAt(-1, "total", loctable!total)
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
Public Function CreateDocFooter(pDoc_No As String, ByRef Json As ChilkatJsonObject, con As ADODB.Connection, ByRef pError As String) As Boolean
Dim cString As String
Dim loctable As New ADODB.Recordset
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(pDoc_No))
aPrm = AddFlag(aPrm, "TAX", mRound(taxRate / 100, 2))
Set loctable = myCmd("dbo.sp_inv_totals", con, adStoredProc, aPrm)
If loctable.EOF Then
    pError = "Invoice data receiver not found"
    Exit Function
End If

success = Json.AddNumberAt(-1, "totalDiscountAmount", loctable!totalDiscountAmount)
success = Json.AddNumberAt(-1, "totalSalesAmount", loctable!totalSalesAmount)
success = Json.AddNumberAt(-1, "netAmount", loctable!netAmount)

Dim aTaxes As ChilkatJsonArray
success = Json.AddArrayAt(-1, "taxTotals")
Set aTaxes = Json.ArrayAt(Json.Size - 1)
success = aTaxes.AddObjectAt(-1)

Dim oTax As ChilkatJsonObject
Set oTax = aTaxes.ObjectAt(aTaxes.Size - 1)
success = oTax.AddStringAt(-1, "taxType", "T1")
success = oTax.AddNumberAt(-1, "amount", loctable!tax)

success = Json.AddNumberAt(-1, "totalAmount", loctable!totalAmount)
success = Json.AddNumberAt(-1, "extraDiscountAmount", 0)
success = Json.AddNumberAt(-1, "totalItemsDiscountAmount", 0)

CreateDocFooter = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Public Function CreateSignDoc(pDoc_No As String, con As ADODB.Connection, ByRef cError As String, myForm As Form) As String
Dim sDocument As String
sDocument = CreateNoSignDoc(pDoc_No, con, cError)

If cError <> "" Then Exit Function

If Trim(sDocument) = "" Then
    cError = "»Ì«‰«  €Ì— ’ÕÌÕ… «Ê €Ì— „ÊÃÊœ…"
    Exit Function
End If

Clipboard.Clear
Clipboard.SetText sDocument

Dim sPath As String
sPath = App.Path & "\" & RetSetting("SIGNER_TOKEN", App.Path & "\invoice_app\conf.txt")
ShellExWait sPath, vbNullString, myForm

sDocument = Clipboard.GetText()
If Trim(sDocument) = "" Then
    cError = "»Ì«‰«  €Ì— ’ÕÌÕ… «Ê €Ì— „ÊÃÊœ…"
    Exit Function
End If

Dim Json As New ChilkatJsonObject

success = Json.Load(sDocument)

'Dim aDocs As ChilkatJsonArray
'Set aDocs = json.ArrayOf("documents")

If Json.ArrayOf("documents") Is Nothing Then
    cError = "·„ Ì „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
    Exit Function
ElseIf Json.ArrayOf("documents").Size = 0 Then
    cError = "·„ Ì „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
    Exit Function
End If

If Json.ArrayOf("documents").ObjectAt(0).ArrayOf("signatures") Is Nothing Then
    cError = "·„ Ì „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
    Exit Function
ElseIf Json.ArrayOf("documents").ObjectAt(0).ArrayOf("signatures").Size = 0 Then
    cError = "·„ Ì „ ⁄„· «· ÊﬁÌ⁄ »‰Ã«Õ"
    Exit Function
Else
    sSign = Json.ArrayOf("documents").ObjectAt(0).ArrayOf("signatures").ObjectAt(0).StringOf("value")
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
End If

CreateSignDoc = sDocument
End Function
Public Function SubmitDocument(pDoc_No As String, con As ADODB.Connection, ByRef pError As String, myForm As Form, Optional pType As String = 10) As Boolean
If Not validSend(pDoc_No, con) Then Exit Function

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
sDocument = CreateSignDoc(pDoc_No, con, pError, myForm)

If pError <> "" Then
    MsgBox pError
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

If sStatusCode >= 400 Then
    MsgBox sResponse
    Exit Function
End If
    
Dim Json As New ChilkatJsonObject
success = Json.Load(sResponse)

Dim aDocs As New ChilkatJsonArray

If Json.ArrayOf("rejectedDocuments").Size > 0 Then
     If Not Json.ArrayOf("rejectedDocuments").ObjectAt(0).ObjectOf("error") Is Nothing Then
        pError = Json.ArrayOf("rejectedDocuments").ObjectAt(0).ObjectOf("error").StringOf("message")
    End If
End If


'Set aDocs = Nothing
'Set aDocs = New ChilkatJsonArray
'
'Set aDocs = oJson.ArrayOf("acceptedDocuments")
'Dim ouuid As ChilkatJsonObject
'Dim sUuid As String




Dim aInsert As Variant
con.BeginTrans
If Not Json.ArrayOf("acceptedDocuments") Is Nothing Then
    If Json.ArrayOf("acceptedDocuments").Size > 0 Then
        sUuid = Json.ArrayOf("acceptedDocuments").ObjectAt(0).StringOf("uuid")
        aInsert = AddFlag(aInsert, "uuid", addstring(sUuid))
        aInsert = AddFlag(aInsert, "inv_tax2", addstring(sUuid))
        aInsert = AddFlag(aInsert, "DATE_UUID", addDate(Date))
        con.Execute addUpdate(aInsert, "FILE6_10H", "FILE6_10H.DOC_NO = " & MyParn(pDoc_No))
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
con.Execute addInsert(aInsert, "FILE6_20E")
con.CommitTrans
MsgBox sStatusCode
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
Dim Json As ChilkatJsonObject

success = url.ParseUrl(apiBaseUrl)
success = rest.AddHeader("Authorization", "Bearer " & pToken)
success = rest.AddHeader("Content-Type", "application/json")
'success = rest.AddHeader("Accept-Encoding", "gzip, deflate, br")

Dim bAutoReconnect As Long
bAutoReconnect = 1
success = rest.Connect(url.Host, url.Port, url.Ssl, bAutoReconnect)
jsonResponseStr = rest.FullRequestString("POST", SubmitUrl, pDocument)

pRespCode = rest.ResponseStatusCode

If (rest.LastMethodSuccess <> 1) Then
    pError = rest.LastErrorText
End If
SubmitDocumentString = jsonResponseStr
End Function
Private Function validSend(pDoc_No As String, con As ADODB.Connection) As Boolean
Dim loctable As ADODB.Recordset
Set loctable = myCmd("select file3_10.desca as Name,file3_10.fileNo,file3_10.id_no,FILE6_10H.inv_tax,FILE6_10H.DateTax,code0_20.desca as Governate " & _
                      " from FILE6_10H inner join file3_10 on FILE6_10H.code = file3_10.code " & _
                      " LEFT OUTER JOIN code0_20 ON FILE3_10.gov = code0_20.CODE" & _
                      " where FILE6_10H.doc_no = " & MyParn(pDoc_No), con)

If loctable.EOF Then
    MsgBox "·«  ÊÃœ ›« Ê—… »Â–« «·—ﬁ„"
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
    MsgBox "·« ÌÊÃœ —ﬁ„ „·› ÷—Ì»Ì «Ê —ﬁ„ "
    Exit Function
End If

If IsNull(loctable!Governate) Then
    MsgBox "„Õ«›Ÿ… «·⁄„Ì· €Ì— „”Ã·…"
    Exit Function
End If

validSend = True
End Function
