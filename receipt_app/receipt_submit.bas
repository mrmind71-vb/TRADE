Attribute VB_Name = "receipt_submit"
Public Function submitReceipt(pDoc_No As String, con As ADODB.Connection, pType As String, ByRef pError As String) As Boolean
If Not validReceipt(pDoc_No, pType, con, pError) Then Exit Function

Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

Dim sToken As String
sToken = getToken_rc(pError)

If pError <> "" Then Exit Function
           
Dim pUUID As String, pPreviousUUID As String, pDate As String
Dim sDocument As String

sDocument = getReceiptString(pDoc_No, pUUID, pPreviousUUID, pDate, con, pType, pError)


If pError <> "" Then Exit Function

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

Dim aInsert As Variant
If Not validJsonArray(json, "acceptedDocuments") Then
    con.BeginTrans
    aInsert = AddFlag(Empty, "doc_no", addstring(pDoc_No))
    aInsert = AddFlag(aInsert, "document_body", addstring(sDocument))
    aInsert = AddFlag(aInsert, "DATE", "GetDate()")
    aInsert = AddFlag(aInsert, "RESPONSE", addstring(sResponse))
    aInsert = AddFlag(aInsert, "StatusCode", addstring(sStatusCode))
    aInsert = AddFlag(aInsert, "ERROR", addstring(pError))
    con.Execute addInsert(aInsert, "FILE6_20R")
    con.CommitTrans
    pError = json.Emit()
    Exit Function
End If

con.BeginTrans
sUuid = json.ArrayOf("acceptedDocuments").ObjectAt(0).StringOf("uuid")
aInsert = AddFlag(aInsert, "UUID_RC", addstring(sUuid))
aInsert = AddFlag(aInsert, "PREVIOUS_UUID", addstring(pPreviousUUID))
aInsert = AddFlag(aInsert, "POS_SERIAL", addstring(GetContRc("setting|posSerial")))
aInsert = AddFlag(aInsert, "DATE_QR", addstring(pDate))
con.Execute addUpdate(aInsert, "FILE6_20H", "DOC_NO = " & MyParn(pDoc_No))
submitReceipt = True

aInsert = AddFlag(Empty, "doc_no", addstring(pDoc_No))
aInsert = AddFlag(aInsert, "uuid", addstring(sUuid))
aInsert = AddFlag(aInsert, "document_body", addstring(sDocument))
aInsert = AddFlag(aInsert, "PREVIOUS_UUID", addstring(pPreviousUUID))
aInsert = AddFlag(aInsert, "DATE", "GetDate()")
aInsert = AddFlag(aInsert, "RESPONSE", addstring(sResponse))
aInsert = AddFlag(aInsert, "StatusCode", addstring(sStatusCode))
aInsert = AddFlag(aInsert, "ERROR", addstring(pError))
con.Execute addInsert(aInsert, "FILE6_20R")
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

success = url.ParseUrl(GetContRc("setting|apiBaseUrl"))
success = rest.AddHeader("Authorization", "Bearer " & pToken)
success = rest.AddHeader("Content-Type", "application/json")

Dim bAutoReconnect As Long
'bAutoReconnect = 1
success = rest.connect(url.Host, url.Port, url.Ssl, bAutoReconnect)

jsonResponseStr = rest.FullRequestString("POST", GetContRc("setting|submitUrl"), pDocument)
pRespCode = rest.ResponseStatusCode

If (rest.LastMethodSuccess <> 1) Then
    pError = rest.LastErrorText & vbCrLf & jsonResponseStr
End If

SubmitDocumentString = jsonResponseStr
End Function
Private Function validSend(pDoc_No As String, con As ADODB.Connection, pType As String) As Boolean
Dim loctable As ADODB.Recordset
aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(pDoc_No))
Set loctable = myCmd("einv.sp_valid_receipt", con, adStoredProc, aPrm)

If loctable.EOF Then
    MsgBox "„” ‰œ €Ì— „ÊÃÊœ «Ê ·Ì” »Â «’‰«›"
    Exit Function
End If

validSend = True
End Function
Public Function getReceiptString(pDoc_No As String, pUUID As String, pPreviousUUID As String, pDate As String, con As ADODB.Connection, pType As String, pError As String) As String
Dim sLang As String
sLang = RetZero(LCase(GetKeyBoard(kbMode.ILANGUAGE)), 10)
If sLang <> Lang_AR Then SetKbLayout Lang_AR

Dim json As New ChilkatJsonObject
Dim aReceipt As New ChilkatJsonArray

success = json.AddArrayAt(-1, "receipts")
Set aReceipt = json.ArrayOf("receipts")
success = aReceipt.AddObjectAt(-1)
Dim receiptObj As New ChilkatJsonObject
Set receiptObj = aReceipt.ObjectAt(aReceipt.Size - 1)

success = receiptObj.AddObjectAt(-1, "header")
Dim headerObj As ChilkatJsonObject
Set headerObj = receiptObj.ObjectOf("header")
setHeaderObj headerObj, pDoc_No, pType, pPreviousUUID, pDate, con, pError

success = receiptObj.AddObjectAt(-1, "documentType")
setTypeObj receiptObj.ObjectAt(receiptObj.Size - 1), pType, pError

success = receiptObj.AddObjectAt(-1, "seller")
setSellerObj receiptObj.ObjectAt(receiptObj.Size - 1), pError

success = receiptObj.AddObjectAt(-1, "buyer")
setBuyerObj receiptObj.ObjectAt(receiptObj.Size - 1), pError

success = receiptObj.AddArrayAt(-1, "itemData")
setItemDataArray receiptObj.ArrayAt(receiptObj.Size - 1), pDoc_No, pType, con, pError

setReceiptObjTotal receiptObj, pDoc_No, pType, con, Error

Dim sDocument As String
sDocument = Serialize(receiptObj.Emit())
pUUID = cryptUUID(sDocument)
success = headerObj.SetStringOf("uuid", pUUID)

getReceiptString = json.Emit()
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Function setHeaderObj(ByRef json As ChilkatJsonObject, pDoc_No As String, pType As String, ByRef pPreviousUUID As String, pDate As String, con As ADODB.Connection, ByRef pError As String) As Boolean
Dim loctable As ADODB.Recordset
Set loctable = myCmd("select [time],file6_20h.doc_ret,[dbo].[GetRefUUID](FILE6_20H.DOC_NO) as referenceUUID from file6_20h where doc_no = " & MyParn(pDoc_No), con)
On Error GoTo myerror

If Not loctable.EOF Then
    pPreviousUUID = MyFuncValue("dbo.GetPreviousUUID", con, MyParn(pDoc_No)) & ""
    pDate = DateGreen(loctable!Time)
    'pDate = DateGreen(Now)
    success = json.AddStringAt(-1, "dateTimeIssued", myFormat_z(pDate))
    success = json.AddStringAt(-1, "receiptNumber", pDoc_No)
    success = json.AddStringAt(-1, "uuid", "")
    success = json.AddStringAt(-1, "previousUUID", pPreviousUUID)
    If pType = "r" Then
        success = json.AddStringAt(-1, "referenceUUID", loctable!referenceUUID & "")
    End If
    success = json.AddStringAt(-1, "currency", "EGP")
    success = json.AddNumberAt(-1, "exchangeRate", 0)
    If pType = "RWR" Then
        Dim pDateSales As String
        success = json.AddStringAt(-1, "documentUseReason", "B")
        Dim sDateSales As String
        sDateSales = invField(loctable!doc_ret, "date", con)
        If IsDate(sDateSales) Then
            success = json.AddStringAt(-1, "salesIssuedDateTime", myFormat_z(DateGreen(sDateSales)))
        End If
    End If
    setHeaderObj = True
End If
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Function setTypeObj(ByRef json As ChilkatJsonObject, pType As String, ByRef pError As String) As Boolean
success = json.AddStringAt(-1, "receiptType", pType)
success = json.AddStringAt(-1, "typeVersion", GetContRc("setting|version"))
setTypeObj = True
End Function
Private Function setSellerObj(ByRef json As ChilkatJsonObject, ByRef pError As String) As Boolean
success = json.AddStringAt(-1, "rin", GetContRc("company|rin"))
success = json.AddStringAt(-1, "companyTradeName", GetContRc("company|companyTradeName"))
success = json.AddStringAt(-1, "branchCode", GetContRc("company|branchCode"))

success = json.AddObjectAt(-1, "branchAddress")
Dim addressObj As New ChilkatJsonObject
Set addressObj = json.ObjectAt(json.Size - 1)
success = addressObj.AddStringAt(-1, "country", GetContRc("company|country"))
success = addressObj.AddStringAt(-1, "governate", GetContRc("company|governate"))
success = addressObj.AddStringAt(-1, "regionCity", GetContRc("company|regionCity"))
success = addressObj.AddStringAt(-1, "street", GetContRc("company|street"))
success = addressObj.AddStringAt(-1, "buildingNumber", GetContRc("company|buildingNumber"))

success = json.AddStringAt(-1, "deviceSerialNumber", GetContRc("setting|posSerial"))
success = json.AddStringAt(-1, "activityCode", GetContRc("company|activityCode"))
setSellerObj = True
End Function
Private Function setBuyerObj(json As ChilkatJsonObject, ByRef pError As String) As Boolean
success = json.AddStringAt(-1, "type", "P")
success = json.AddStringAt(-1, "id", "")
success = json.AddStringAt(-1, "name", "")
success = json.AddStringAt(-1, "mobileNumber", "")
success = json.AddStringAt(-1, "paymentNumber", "")
setBuyerObj = True
End Function
Private Function setItemDataArray(aJson As ChilkatJsonArray, pDoc_No As String, pType As String, con As ADODB.Connection, ByRef pError As String) As Boolean
Dim loctable As New ADODB.Recordset
Dim aPrm As Variant
Dim nTaxRate As Double
nTaxRate = Val(GetContRc("setting|taxRate"))
aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(pDoc_No))
aPrm = AddFlag(aPrm, "TAX", mRound(nTaxRate / 100, 2))

Set loctable = myCmd("einv.sp_receipt_line", con, adStoredProc, aPrm)
If loctable.EOF Then
    pError = "»Ì«‰«  «·„” ‰œ €Ì— „”Ã·…"
    Exit Function
End If

Dim itemObj As ChilkatJsonObject
Do Until loctable.EOF
    success = aJson.AddObjectAt(-1)
    Set itemObj = aJson.ObjectAt(aJson.Size - 1)
    
    success = itemObj.AddStringAt(-1, "internalCode", loctable!internalCode)
    success = itemObj.AddStringAt(-1, "description", myEsc(loctable!Description))
    success = itemObj.AddStringAt(-1, "itemType", loctable!itemType)
    success = itemObj.AddStringAt(-1, "itemCode", loctable!itemCode & "")
    success = itemObj.AddStringAt(-1, "unitType", loctable!unitType)
    success = itemObj.AddNumberAt(-1, "quantity", loctable!quantity * IIf(pType = "r" Or pType = "RWR", -1, 1))
    success = itemObj.AddNumberAt(-1, "unitPrice", loctable!unitPrice)
    
    success = itemObj.AddNumberAt(-1, "netSale", loctable!netSale * IIf(pType = "r" Or pType = "RWR", -1, 1))
    success = itemObj.AddNumberAt(-1, "totalSale", loctable!totalSale * IIf(pType = "r" Or pType = "RWR", -1, 1))
    success = itemObj.AddNumberAt(-1, "total", loctable!TOTAL * IIf(pType = "r" Or pType = "RWR", -1, 1))
    
    ' Discount Array
    success = itemObj.AddArrayAt(-1, "commercialDiscountData")
    Dim aDiscount As ChilkatJsonArray
    Set aDiscount = itemObj.ArrayAt(itemObj.Size - 1)
    
    success = aDiscount.AddObjectAt(-1)
    Dim discountObj As ChilkatJsonObject
    Set discountObj = aDiscount.ObjectAt(aDiscount.Size - 1)
           
    success = discountObj.AddNumberAt(-1, "amount", loctable!discount * IIf(pType = "r" Or pType = "RWR", -1, 1))
    success = discountObj.AddStringAt(-1, "description", "XYZ")
    success = discountObj.AddNumberAt(-1, "rate", 0)
    ' Tax Items
    
    success = itemObj.AddArrayAt(-1, "taxableItems")
    Dim aTax As ChilkatJsonArray
    Set aTax = itemObj.ArrayAt(itemObj.Size - 1)
        
    success = aTax.AddObjectAt(-1)
    Dim taxObj As ChilkatJsonObject
    Set taxObj = aTax.ObjectAt(aTax.Size - 1)
    
    success = taxObj.AddStringAt(-1, "taxType", "T1")
    success = taxObj.AddNumberAt(-1, "amount", loctable!tax * IIf(pType = "r" Or pType = "RWR", -1, 1))
    success = taxObj.AddStringAt(-1, "subType", GetContRc("setting|subType"))
    success = taxObj.AddNumberAt(-1, "rate", nTaxRate)
                                        
    loctable.MoveNext
Loop
setItemDataArray = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Private Function setReceiptObjTotal(json As ChilkatJsonObject, pDoc_No As String, pType As String, con As ADODB.Connection, ByRef pError As String) As Boolean
Dim loctable As New ADODB.Recordset
Dim aPrm As Variant
Dim nTaxRate As Double
nTaxRate = Val(GetContRc("setting|taxRate"))
aPrm = AddFlag(aPrm, "DOC_NO", TurnValue(pDoc_No))
aPrm = AddFlag(aPrm, "TAX", mRound(nTaxRate / 100, 2))

On Error GoTo myerror

Set loctable = myCmd("einv.sp_receipt_total", con, adStoredProc, aPrm)
If loctable.EOF Then
    pError = "»Ì«‰«  «·„” ‰œ €Ì— „”Ã·…"
    Exit Function
End If
If Not loctable.EOF Then
    success = json.AddNumberAt(-1, "totalSales", loctable!totalSales * IIf(pType = "r" Or pType = "RWR", -1, 1))
    success = json.AddNumberAt(-1, "totalCommercialDiscount", loctable!totalCommercialDiscount * IIf(pType = "r" Or pType = "RWR", -1, 1))
    success = json.AddNumberAt(-1, "netAmount", loctable!netAmount * IIf(pType = "r" Or pType = "RWR", -1, 1))
    success = json.AddNumberAt(-1, "totalAmount", loctable!totalAmount * IIf(pType = "r" Or pType = "RWR", -1, 1))

    
    success = json.AddArrayAt(-1, "taxTotals")
    Dim aTax As ChilkatJsonArray
    Set aTax = json.ArrayAt(json.Size - 1)
    
    success = aTax.AddObjectAt(aTax.Size - 1)
    Dim taxObj As New ChilkatJsonObject
    Set taxObj = aTax.ObjectAt(aTax.Size - 1)
    success = taxObj.AddStringAt(-1, "taxType", "T1")
    success = taxObj.AddNumberAt(-1, "amount", loctable!tax * IIf(pType = "r" Or pType = "RWR", -1, 1))
End If
success = json.AddStringAt(-1, "paymentMethod", "C")
setReceiptObjTotal = True
Exit Function
myerror:
pError = Err.Description
Err.Clear
End Function
Public Function getToken_rc(Optional ByRef pError As String) As String
Dim success As Integer
Dim httpB As New ChilkatHttp
Dim reqB As New ChilkatHttpRequest
reqB.httpVerb = "POST"

Dim xml As New ChilkatXml
success = xml.LoadXmlFile(tempPath & "\receipt.xml")
If success = 0 Then
    pError = xml.LastErrorXml
    Exit Function
End If

reqB.Path = xml.GetChildContent("setting|tokenUrl")
reqB.contentType = "application/x-www-form-urlencoded"
reqB.AddHeader "posserial", xml.GetChildContent("setting|posSerial")
reqB.AddHeader "pososversion", xml.GetChildContent("setting|pososVersion")

reqB.AddParam "grant_type", "client_credentials"
reqB.AddParam "client_id", xml.GetChildContent("setting|client_id")
reqB.AddParam "client_secret", xml.GetChildContent("setting|client_secret")

Dim respB As ChilkatHttpResponse
Set respB = httpB.PostUrlEncoded(xml.GetChildContent("setting|idSrvBaseUrl") & xml.GetChildContent("setting|tokenUrl"), reqB)
If (httpB.LastMethodSuccess = 0) Then
    pError = httpB.LastErrorText
    Exit Function
End If

Dim json As New ChilkatJsonObject
success = json.Load(respB.BodyStr)
If (success <> 1) Then
    pError = json.LastErrorText
    Exit Function
End If
getToken_rc = json.StringOf("access_token")
If getToken_rc = "" Then pError = "Empty Token"
End Function
Public Function GetContX(pFile As String, pPath As String, Optional ByRef pError As String)
Dim xml As New ChilkatXml
success = xml.LoadXmlFile(pFile)
If success = 0 Then
    pError = xml.LastErrorXml
    Exit Function
End If

GetContX = xml.GetChildContent(pPath)
End Function
Public Function GetContRc(pPath As String, Optional ByRef pError As String) As String
Dim xml As New ChilkatXml
success = xml.LoadXmlFile(tempPath & "\receipt.xml")
If success = 0 Then
    pError = xml.LastErrorXml
    Exit Function
End If

GetContRc = xml.GetChildContent(pPath)
End Function
Public Function cryptUUIDOld(pDocument As String) As String
Dim crypt As New ChilkatCrypt2

' Choose the hash algorithm.
' Can be  "sha1", "sha256", "sha384", "sha512", "md2", "md5", "haval", "ripemd128", "ripemd160","ripemd256", or "ripemd320".
crypt.Charset = "utf-8"
crypt.HashAlgorithm = "sha256"
' The byte representation of the string matters when hashing.
' Tell Chilkat that we want to use the utf-8 byte representation.
Dim hashBytes() As Byte
hashBytes = crypt.HashString(pDocument)

' Let's examine the hash as a hex string.
Dim sb As New ChilkatStringBuilder
success = sb.AppendEncoded(hashBytes, "hexlower")

cryptUUIDOld = sb.GetAsString()
End Function
Private Function cryptUUID(strTextToHash As String) As String
'Dim strTextToHash As String
'strTextToHash = "the text to be hashed..."

Dim crypt As New ChilkatCrypt2
crypt.Charset = "utf-8"
crypt.EncodingMode = "hexlower"
crypt.HashAlgorithm = "SHA256"

Dim sb As New ChilkatStringBuilder
success = sb.Append(strTextToHash)
success = sb.ToCRLF()

cryptUUID = crypt.HashStringENC(sb.GetAsString())
End Function
Public Function iFormat(pDate As String) As String
iFormat = Format(pDate, "YYYY-MM-DD HH:NN:SS")
End Function
Private Function myEsc(pString As String) As String
myEsc = Replace(pString, Chr(13), "")
myEsc = Replace(myEsc, Chr(10), "")
End Function
Public Function validJsonArray(json As ChilkatJsonObject, sName As String) As Boolean
If json.ArrayOf(sName) Is Nothing Then Exit Function
If json.ArrayOf(sName).Size = 0 Then Exit Function
validJsonArray = True
End Function
Private Function validReceipt(pDoc_No As String, pType As String, con As ADODB.Connection, ByRef pError As String) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("[dbo].[sp_invoice_data]", con, adStoredProc, AddFlag(Empty, "doc_no", pDoc_No))
If loctable.EOF And loctable.BOF Then
    pError = "›« Ê—… €Ì— „”Ã·…"
End If

If pType = "r" Then
    If IsNull(MyFuncValue("[dbo].[GetRefUUID]", con, MyParn(pDoc_No))) Then
        pError = "›« Ê—… „— Ã⁄ »œÊ‰ „—Ã⁄"
    End If
Else
    validReceipt = True
End If
loctable.Close
Set loctable = Nothing
End Function
Public Function SendReceipt(pDoc_No As String, con As ADODB.Connection, Optional pError As String = "") As Boolean
Dim pType As String
If (Not receiptType(pDoc_No, pType, con, pError)) Then Exit Function
If Not submitReceipt(pDoc_No, con, pType, pError) Then Exit Function
SendReceipt = True
End Function
Private Function receiptType(pDoc_No As String, ByRef pType As String, con As ADODB.Connection, pError As String) As Boolean
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select file6_20h.date,FILE6_20H.DOC_RET," & _
              "max(file6_20.quant) as maxQuant, " & _
              " min(file6_20.quant) as minQuant" & _
              " from file6_20 inner join file6_20h on file6_20.doc_no = file6_20h.doc_no " & _
              " where file6_20.doc_no = " & MyParn(pDoc_No) & _
              " group by file6_20h.date,FILE6_20H.DOC_RET", con)
              
If Not (loctable.EOF And loctable.BOF) Then
    If IsNull(loctable!maxQuant) And IsNull(loctable!minQuant) Then
        pError = "«·„” ‰œ ·« ÌÕ ÊÌ ⁄·Ì «’‰«›"
        Exit Function
    ElseIf loctable!maxQuant > 0 And loctable!minQuant < 0 Then
        pError = "«·„” ‰œ ÌÕ ÊÌ ⁄·Ì ﬂ„Ì«  „ÊÃ»… Êﬂ„Ì«  ”«·»…"
        Exit Function
    ElseIf loctable!maxQuant = 0 And loctable!minQuant = 0 Then
        pError = "«·„” ‰œ ·« ÌÕ ÊÌ ⁄·Ì ﬂ„Ì« "
        Exit Function
    ElseIf loctable!maxQuant > 0 Then
        pType = "s"
    ElseIf loctable!minQuant < 0 Then
        pType = "r"
    End If
Else
    pError = "·«  ÊÃœ „” ‰œ«  „»Ì⁄« "
    Exit Function
End If

If pType = "r" Then
    If IsNull(loctable!doc_ret) Then
        pError = "„” ‰œ „— Ã⁄ »œÊ‰ „” ‰œ »Ì⁄ ··—ÃÊ⁄ «·ÌÂ"
        Exit Function
    ElseIf IsEmpty(invField(loctable!doc_ret, "doc_no", con)) Then
        pError = "„” ‰œ „— Ã€ „”Ã· €Ì— „ÊÃÊœ"
        Exit Function
    ElseIf IsNull(invField(loctable!doc_ret, "UUID_RC", con)) Then
        If IsDate(GetContRc("setting|date_ref")) Then
            Dim sDate As String
            sDate = myFormat(invField(loctable!doc_ret, "DATE", con))
            If sDate <= myFormat(GetContRc("setting|date_ref")) Then
                pType = "RWR"
            Else
                pError = "„” ‰œ „— Ã€ »œÊ‰ —ﬁ„ „—Ã⁄Ì"
                Exit Function
            End If
        Else
            pType = "RWR"
        End If
    End If
End If
receiptType = True
End Function
Public Function DateGreen(pDate As String) As String
Dim dateTime As New CkDateTime
success = dateTime.SetFromTimestamp(Format(pDate, "YYYY-MM-DDThh:mm:ssTZD"))
success = dateTime.AddSeconds(-1 * dateTime.UtcOffset)
DateGreen = dateTime.GetAsIso8601("YYYY-MM-DD hh:mm:ss", 0)
End Function
Public Function GreenToLocal(pDate As String, Optional pFormat As String = "YYYY-MM-DDThh:mm:ssTZD") As String
Dim dateTime As New CkDateTime
success = dateTime.SetFromTimestamp(Format(pDate, "YYYY-MM-DDThh:mm:ssTZD"))
GreenToLocal = dateTime.GetAsIso8601("YYYY-MM-DDThh:mm:ssTZD", 1)
End Function
Public Function retQrCode(ByVal pDoc_No As String, con As ADODB.Connection, Optional pFormat As String = "YYYY-MM-DDThh:mmZ") As String
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select date_Qr,uuid_rc from file6_20h where doc_no = " & MyParn(pDoc_No), con)
Dim xml As New ChilkatXml
success = xml.LoadXmlFile(tempPath & "\receipt.xml")
If success = 0 Then Exit Function

If IsNull(loctable!UUID_RC) Then Exit Function
If Trim(loctable!UUID_RC) = "" Then Exit Function
If Not loctable.EOF Then
    retQrCode = xml.GetChildContent("setting|QRUrl") & "/" & loctable!UUID_RC & "/share/" & Format(loctable!date_qr, pFormat)
End If
loctable.Close
Set loctable = Nothing
End Function
Public Function invField(pDoc_No As String, pField As String, con As ADODB.Connection) As Variant
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("[dbo].[sp_invoice_data]", con, adStoredProc, AddFlag(aPrm, "doc_no", pDoc_No))
If loctable.EOF Then
    invField = Empty
Else
    invField = loctable(pField).Value
End If
loctable.Close
Set loctable = Nothing
End Function


