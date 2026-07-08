Attribute VB_Name = "Special"
Public nusercode As String
Public cusername As String
Public cSalesMan As String
Public dSalesDate As String, aPassword As Variant, aAddModel As Variant, dMaxDate As Date
Public aPrinterBar As Variant
Function itemCost(cItem, Optional pDate As String = "") As Single
Dim con As New adodb.Connection
If con.State = adStateOpen Then con.Close
openCon con

cString = "Select FILE7_20.PRICE FROM  " & _
          "FILE7_20 INNER JOIN FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO " & _
          " WHERE ITEM = " & MyParn(cItem)
          
If IsDate(pDate) Then
    cString = cString & turnFound(cString) & " FILE7_20H.Date <= " & DateSq(pDate)
End If
cString = cString & " " & " ORDER BY FILE7_20H.DATE DESC"
itemCost = Round(Val(GetDesca(cString, con) & ""), 2)
If itemCost = 0 Then itemCost = Val(GetDesca("select cost from file1_10 where item = " & MyParn(cItem), con))
End Function

Sub ModelLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "Select TOP 1000 FILE1_10h.MODEL ,FILE1_10h.DescA,File1_10h.MODELFACT0 ,FILE1_10h.MOSM ,FILE1_10h.SUPP , fact.DESCA   FROM FILE1_10h INNER JOIN fact ON FILE1_10h.FACT = fact.CODE  LEFT JOIN MOSM ON MOSM.MOSM = FILE1_10H.MOSM "
Generalarray(2) = "order by MOSM.DATE DESC  , FILE1_10H.MODEL "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ „ÊœÌ·"
listarray(0, 1) = "(%%File1_10h.MODELFACT0+FACT.DESCA%%)"

listarray(1, 0) = "«·„ÊœÌ·"
listarray(1, 1) = "(%%File1_10h.DESCA%%)"

'listarray(1, 0) = "„Ê”„ ° „Ê—œ "
'listarray(1, 1) = "(%%File1_10h.MOSM & FILE1_10h.SUPP%%)"

GrdArray(0, 0) = "„ÊœÌ·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·’‰› "
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·„Ê’„ "
GrdArray(3, 1) = 800

GrdArray(4, 0) = " „ﬂ » Ã„·…"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·„’‰⁄"
GrdArray(5, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub ModelNoLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
'                                        0               1                    2                   3           4            5               6
Generalarray(1) = "Select TOP 100 FILE1_10h.MODELNO,FILE1_10h.DescA,File1_10h.MODELFACT0,FILE1_10h.MOSM,fact.DESCA  FROM FILE1_10h INNER JOIN fact ON FILE1_10h.FACT = fact.CODE"
Generalarray(2) = " GROUP BY FILE1_10h.MODELNO,FILE1_10h.DescA,File1_10h.MODELFACT0,FILE1_10h.MOSM,fact.DESCA order by file1_10h.MOSM "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ „ÊœÌ·"
listarray(0, 1) = "(%%File1_10h.MODELFACT0 + FACT.DESCA%%)"

listarray(1, 0) = "«·„ÊœÌ·"
listarray(1, 1) = "(%%File1_10h.DESCA%%)"

'listarray(1, 0) = "„Ê”„ ° „Ê—œ "
'listarray(1, 1) = "(%%File1_10h.MOSM & FILE1_10h.SUPP%%)"

GrdArray(0, 0) = "„ÊœÌ·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·’‰› "
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(2, 1) = 1000

GrdArray(3, 0) = "«·„Ê’„ "
GrdArray(3, 1) = 800

GrdArray(4, 0) = " «·„’‰⁄"
GrdArray(4, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub ItemsLookupAll(oForm As Form, oSearch As Form, Optional pOrderNo As String = "")
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(9, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "SELECT   top 2000 item , file1_10.mosm , fact.desca , file1_10.modelfact0 , file1_10.supp,file1_10.desca , file1_10.scal , file1_10.color , PRICE , PRICE_2 from file1_10 inner join fact on fact.code = file1_10.fact  WHERE ISSTOP = 0 "
If pOrderNo <> "" Then
    Generalarray(1) = Generalarray(1) & " AND ITEM IN (SELECT ITEM FROM FILE6_50 WHERE DOC_NO = " & MyParn(pOrderNo) & ")"
End If
Generalarray(2) = "ORDER BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_SCAL "
Generalarray(3) = 12000
Generalarray(5) = False

listarray(0, 0) = "≈”„ «·’‰› ° «·„’‰⁄ ° —ﬁ„ „ÊœÌ· & „ﬁ«” "
listarray(0, 1) = "(%%FILE1_10.DESCA+file1_10.MODELFACT+FACT.DESCA%%)"

listarray(1, 0) = "„Ê”„ "
listarray(1, 1) = "(%%File1_10.MOSM%%)"

listarray(2, 0) = "„ﬁ«”"
listarray(2, 1) = "%%File1_10.SCAL%%"

listarray(3, 0) = "«··Ê‰"
listarray(3, 1) = "(%%File1_10.COLOR%%)"
 
GrdArray(0, 0) = "»«—ﬂÊœ"
GrdArray(0, 1) = 900

GrdArray(1, 0) = "„Ê”„"
GrdArray(1, 1) = 800

GrdArray(2, 0) = "«·„’‰⁄"
GrdArray(2, 1) = 1800

GrdArray(3, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "„ﬂ »"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·’‰›"
GrdArray(5, 1) = 4500

GrdArray(6, 0) = "«·„ﬁ«”"
GrdArray(6, 1) = 1300

GrdArray(7, 0) = "«··Ê‰"
GrdArray(7, 1) = 800

GrdArray(8, 0) = " «·”⁄—"
GrdArray(8, 1) = 1000

GrdArray(9, 0) = " ”⁄— «Êﬂ«“ÊÌ‰"
GrdArray(9, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub ItemsLookup(oForm As Form, oSearch As Form, Optional pOrderNo As String = "", Optional pMosm As String)
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(9, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "SELECT   top 2000 item , file1_10.mosm , fact.desca , file1_10.modelfact0 , file1_10.supp,file1_10.desca , file1_10.scal , file1_10.color , PRICE , PRICE_2 from file1_10 inner join fact on fact.code = file1_10.fact  WHERE ISSTOP = 0 "
If pOrderNo <> "" Then
    Generalarray(1) = Generalarray(1) & " AND ITEM IN (SELECT ITEM FROM FILE6_50 WHERE DOC_NO = " & MyParn(pOrderNo) & ")"
End If
Generalarray(2) = "ORDER BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_SCAL "
Generalarray(3) = 12000
Generalarray(5) = False

listarray(0, 0) = "≈”„ «·’‰› ° «·„’‰⁄ ° —ﬁ„ „ÊœÌ· & „ﬁ«” "
listarray(0, 1) = "(%%FILE1_10.DESCA%%  OR %%file1_10.MODELFACT%% OR %%FACT.DESCA%%)"

listarray(1, 0) = "«·„Ê”„"
listarray(1, 1) = "(FILE1_10.[MOSM] = 'cFilter')"
listarray(1, 2) = "SELECT MOSM,MOSM AS DESCA FROM MOSM"
listarray(1, 3) = "MOSM"
listarray(1, 4) = "MOSM"
listarray(1, 5) = pMosm


listarray(2, 0) = "„ﬁ«”"
listarray(2, 1) = "%%File1_10.SCAL%%"

listarray(3, 0) = "«··Ê‰"
listarray(3, 1) = "(%%File1_10.COLOR%%)"
 
GrdArray(0, 0) = "»«—ﬂÊœ"
GrdArray(0, 1) = 900

GrdArray(1, 0) = "„Ê”„"
GrdArray(1, 1) = 800

GrdArray(2, 0) = "«·„’‰⁄"
GrdArray(2, 1) = 1800

GrdArray(3, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "„ﬂ »"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·’‰›"
GrdArray(5, 1) = 4500

GrdArray(6, 0) = "«·„ﬁ«”"
GrdArray(6, 1) = 1300

GrdArray(7, 0) = "«··Ê‰"
GrdArray(7, 1) = 800

GrdArray(8, 0) = " «·”⁄—"
GrdArray(8, 1) = 1000

GrdArray(9, 0) = " ”⁄— «Êﬂ«“ÊÌ‰"
GrdArray(9, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub CustLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = oForm

Generalarray(1) = "Select FILE3_10.CODE , FILE3_10.DESCA , FILE3_50.DESCA  FROM FILE3_10  INNER JOIN FILE3_50  ON FILE3_10.[GROUP] = FILE3_50.CODE  "
Generalarray(2) = "order by file3_10.desca "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì· "
listarray(0, 1) = "(%%File3_10.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·⁄„Ì·"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ ⁄„·«¡ √Ã·"
oSearch.Show 1
End Sub
Sub FactLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm

Generalarray(1) = "Select FACT.CODE , FACT.DESCA FROM FACT"
Generalarray(2) = "order by FACT.CODE"
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ«·„Œ ’—-«·«”„"
listarray(0, 1) = "(%%FACT.CODE%% OR %%FACT.DESCA%%)"


GrdArray(0, 0) = "«·ﬂÊœ «·„Œ ’—"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·≈”„"
GrdArray(1, 1) = 7000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "«” ⁄·«„ «·„’«‰⁄"
oSearch.Show 1
End Sub
Sub SuppLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 1)

Set Generalarray(0) = oForm

Generalarray(1) = "Select FILE4_10.CODE , FILE4_10.DESCA , FILE4_50.DESCA  FROM FILE4_10  left JOIN FILE4_50  ON FILE4_10.[GROUP] = FILE4_50.CODE  "
Generalarray(2) = "order by file4_10.desca "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·„Ê—œ "
listarray(0, 1) = "(%%File4_10.DESCA%%    )"


GrdArray(0, 0) = "ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·„Ê—œ"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „Ê—œÌ‰  "
oSearch.Show 1
End Sub
Sub CustSUPPLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select FILE3_20.CODE , FILE3_20.DESCA , file3_20.Phone1 , IsOk FROM FILE3_20  "
Generalarray(2) = "order by file3_20.desca "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì· "
listarray(0, 1) = "(%%File3_20.DESCA%%   or %%File3_20.Phone1%%    )"


GrdArray(0, 0) = "ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·⁄„Ì·"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = " ·Ì›Ê‰"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = " "
GrdArray(3, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ ⁄„·«¡ ﬂ«—  Âœ«Ì« "
oSearch.Show 1
End Sub

Function ValidDate() As Boolean
ValidDate = Format("01-12-2000", "dd-mm-yyyy") = "01-12-2000"
If Not ValidDate Then
    cString = "‰Ÿ«„ «· «—ÌŒ €Ì— ’«·Õ" & vbCrLf & _
            "«Ã⁄· «·œÊ·… ›Ï ·ÊÕ… «· Õﬂ„" & vbCrLf & _
            "control Panel" & vbCrLf & _
            "Regional and language options" & vbCrLf & _
            "Egypt"
            
    MsgBox ArbString(cString), vbCritical
'    End

End If
End Function
Function LastBalance(cItem, cStore, pCon As adodb.Connection) As Double
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
If cItem = "" Or cStore = "" Then Exit Function
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Store", adVarWChar, adParamInput, 3, cStore)
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adDouble, adParamInput, 20, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)

cmdTable.CommandText = "LastBalance"
Set obj = cmdTable.Execute
LastBalance = Val(cmdTable.Parameters(2).Value & "")
Set obj = Nothing
End Function


Function lastColor(cItem As String, cDoc_No As String, pCon As adodb.Connection) As Variant
Dim obj As New adodb.Recordset, aRet(1) As Double
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adVarWChar, adParamInput, 15, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("doc_no", adVarWChar, adParamInput, 6, cDoc_No)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret1", adDouble, adParamOutput)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret2", adDouble, adParamOutput)

cmdTable.CommandText = "lastcolor"
Set obj = cmdTable.Execute
aRet(0) = Val(cmdTable.Parameters(2).Value & "")
aRet(1) = Val(cmdTable.Parameters(3).Value & "")
lastColor = aRet
Set obj = Nothing
End Function
Function LastCost(cItem, pCon As adodb.Connection) As Double
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adVarWChar, adParamInput, 15, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)

cmdTable.CommandText = "LastCost"
Set obj = cmdTable.Execute
LastCost = Val(cmdTable.Parameters(1).Value & "")
Set obj = Nothing
End Function
Function RetValChr(ByVal sValue) As String
sValue = Trim(sValue)
Dim str1 As String, str2 As String
For i = 1 To Len(sValue)
    If IsNumeric(Right(sValue, i)) Then
        str1 = Val(Right(sValue, i))
        Exit For
    End If
Next
If str1 <> "" Then
    RetValChr = Left(sValue, Len(str1)) & RetZero(str1, 10)
End If
End Function
Function retFile(pModel As String) As String
Dim cString As String
If Len(Mid(pModel, 4, 3) & turn(Mid(pModel, 4, 3), "\")) < 3 Then Exit Function
retFile = App.Path & "\PICT\" & Mid(pModel, 4, 3) & turn(Mid(pModel, 4, 3), "\") & pModel & ".jpg"
End Function
Function RetItemBalance(cItem, cStore, dDate, pCon As adodb.Connection, Optional pCode As String = "") As Double
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Store", adVarWChar, adParamInput, 3, cStore)
cmdTable.Parameters.Append cmdTable.CreateParameter("date", adVarWChar, adParamInput, 12, Format(dDate, "YYYY-MM-DD"))
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adDouble, adParamInput, 20, cItem)
If pCode <> "" Then cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adVarWChar, adParamInput, 6, pCode)

cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)
If pCode = "" Then
    cmdTable.CommandText = "retitemBalance"
    Set obj = cmdTable.Execute
    RetItemBalance = Val(cmdTable.Parameters(3).Value & "")
Else
    cmdTable.CommandText = "retitemBalance_CODE"
    Set obj = cmdTable.Execute
    RetItemBalance = Val(cmdTable.Parameters(4).Value & "")
End If
Set obj = Nothing
End Function
Function ItemBalanceNoStore(cItem, dDate, pCon As adodb.Connection) As Double
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc

cmdTable.Parameters.Append cmdTable.CreateParameter("date", adVarWChar, adParamInput, 12, Format(dDate, "YYYY-MM-DD"))
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adDouble, adParamInput, 20, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)

cmdTable.CommandText = "itemBalanceNoStore"
Set obj = cmdTable.Execute
ItemBalanceNoStore = Val(cmdTable.Parameters(2).Value & "")
Set obj = Nothing
End Function
Function ModelFind(pModel As String, pCon As adodb.Connection) As adodb.Recordset
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("MODELNO", adVarWChar, adParamInput, 30, pModel)
cmdTable.CommandText = "ModelFind"
Set obj = cmdTable.Execute
Set ModelFind = obj
Set obj = Nothing
End Function
Function card_last(pCode As String, pCon As adodb.Connection) As String
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adVarWChar, adParamInput, 6, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return1", adVarWChar, adParamOutput, 10)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return2", adVarWChar, adParamOutput, 10)
cmdTable.CommandText = "card_last"
Set rdTable = cmdTable.Execute
'pDoc_no = cmdTable.Parameters(1).Value & ""
card_last = cmdTable.Parameters(2).Value & ""
Set rdTable = Nothing
End Function
Function valid_card(pCode As String, pDate As String, pCon As adodb.Connection) As String
Dim sDate
sDate = Format(card_last(pCode, pCon), "yyyy-mm-dd")
If sDate = "" Then
    valid_card = "null"
    Exit Function
End If
If sDate < Format(DateAdd("m", -6, pDate), "yyyy-mm-dd") Then
    valid_card = "valid"
End If
valid_card = "ok"
End Function
Function card_Sales(pCode As String, pCon As adodb.Connection) As Double
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adVarWChar, adParamInput, 6, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adDecimal, adParamOutput)
cmdTable.Parameters(1).Precision = 18
cmdTable.Parameters(1).NumericScale = 2
cmdTable.CommandText = "card_sales"
Set rdTable = cmdTable.Execute
card_Sales = cmdTable.Parameters(1).Value
Set rdTable = Nothing
End Function
Function card_discount(pCode As String, pCon As adodb.Connection) As Double
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adVarWChar, adParamInput, 6, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adDecimal, adParamOutput)
cmdTable.Parameters(1).Precision = 18
cmdTable.Parameters(1).NumericScale = 2
cmdTable.CommandText = "card_discount"
Set rdTable = cmdTable.Execute
card_discount = cmdTable.Parameters(1).Value
Set rdTable = Nothing
End Function


Function isOffice(pCode As String, pCon As adodb.Connection) As Boolean
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adVarWChar, adParamInput, 3, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adBoolean, adParamOutput)
cmdTable.CommandText = "isOffice"
Set rdTable = cmdTable.Execute
isOffice = cmdTable.Parameters(1).Value
Set rdTable = Nothing
End Function
Function retCode(pCode As String, pCon As adodb.Connection) As String
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("code", adVarWChar, adParamInput, 3, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adVarWChar, adParamOutput, 3)
cmdTable.CommandText = "retCode"
Set rdTable = cmdTable.Execute
retCode = cmdTable.Parameters(1).Value
Set rdTable = Nothing
End Function
Function fact_Desc(pCode As String, pCon As adodb.Connection) As String
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adVarWChar, adParamInput, 3, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adVarWChar, adParamOutput)
cmdTable.CommandText = "fact_desc"
Set rdTable = cmdTable.Execute
fact_Desca = cmdTable.Parameters(1).Value & ""
Set rdTable = Nothing
End Function
Function Item_cost(pBARCODE As String, pCon As adodb.Connection) As Double
    Item_cost = Val(GetDesca(" SELECT COSTITEM FROM FILE1_10 WHERE ITEM = " & pBARCODE, pCon) & "")
End Function
Function retSysdate() As String
retSysdate = Format(IIf(Val(Format(Time, "hh")) > 4, Date, DateAdd("d", -1, Date)), "dd-mm-yyyy")
End Function
Function Model_Found(pModel, pCon As adodb.Connection) As Integer
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Model", adVarWChar, adParamInput, 30, pModel)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adDouble, adParamOutput)

cmdTable.CommandText = "Model_Found"
Set obj = cmdTable.Execute
Model_Found = Val(cmdTable.Parameters(1).Value & "")
Set obj = Nothing
End Function
Function BAL_SUPPL(cCode, cMosm, pCon As adodb.Connection) As Double
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
If cCode = "" Or cMosm = "" Then Exit Function
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE", adVarWChar, adParamInput, 6, cCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("MOSM", adVarWChar, adParamInput, 3, cMosm)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)

cmdTable.CommandText = "BAL_SUPPL"
Set obj = cmdTable.Execute
BAL_SUPPL = Val(cmdTable.Parameters(2).Value & "")
Set obj = Nothing
End Function
Function validNumber(ByVal pNumber As String) As Boolean
Dim cString As String
pNumber = Trim(pNumber)
If Not IsNumeric(pNumber) Then Exit Function
If Val(pNumber) <= 0 Then Exit Function
For i = 1 To Len(validNumber)
    cString = Mid(pNumber, i, 1)
    If i = 1 And cString = "0" Then Exit Function
    If Not IsNumeric(cString) Then Exit Function
Next
validNumber = True
'validNumber = Int(pNumber) & "" = pNumber & "" And Val(pNumber) > 0
End Function
Function ValidInt(pString As Variant)
ValidInt = Int(Val(pString & "")) & "" = Trim(pString)
End Function
Function ItemFind(pItem As String, pCon As adodb.Connection) As adodb.Recordset
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
If Len(pItem) > 7 Then pItem = 0

cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adInteger, adParamInput, 20, pItem)
cmdTable.CommandText = "ItemFind"
Set rdTable = cmdTable.Execute
Set ItemFind = rdTable
Set rdTable = Nothing
End Function
Function ItemFind_BARCODE(pBARCODE As String, pCon As adodb.Connection) As adodb.Recordset
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("BARCODE", adChar, adParamInput, 20, pBARCODE)
cmdTable.CommandText = "ItemFind_BARCODE"
Set rdTable = cmdTable.Execute
Set ItemFind_BARCODE = rdTable
Set rdTable = Nothing
End Function

Function validItem(pItem As String, pCon As adodb.Connection) As Boolean
On Error GoTo myerror
If Not ValidInt(pItem) Then Exit Function
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adInteger, adParamInput, 10, pItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adBoolean, adParamOutput)
cmdTable.CommandText = "validItem"
Set rdTable = cmdTable.Execute
validItem = cmdTable.Parameters(1).Value
Set rdTable = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Function ItemFields(ByVal pItem As String, pCon As adodb.Connection) As Variant
If Not ValidInt(pItem) Then Exit Function
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adInteger, adParamInput, 10, pItem)
cmdTable.CommandText = "ItemFind"
Set rdTable = cmdTable.Execute
If Not (rdTable.EOF And rdTable.BOF) Then
    For i = 0 To rdTable.Fields.Count - 1
        ItemFields = AddFlag(ItemFields, rdTable.Fields(i).Name, rdTable.Fields(i).Value)
    Next
End If
Set rdTable = Nothing
End Function
Function ItemField(pItem As String, pField As String, pCon As adodb.Connection) As Variant
If Not ValidInt(pItem) Then Exit Function
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc

cmdTable.Parameters.Append cmdTable.CreateParameter("item", adInteger, adParamInput, 10, pItem)
cmdTable.CommandText = "ItemFind"
Set rdTable = cmdTable.Execute
If Not (rdTable.EOF And rdTable.BOF) Then ItemField = rdTable(pField)
Set rdTable = Nothing
End Function
Function QItemIn(cItem, cStore, pCon As adodb.Connection) As Double
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
If cItem = "" Or cStore = "" Then Exit Function
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("Store", adVarWChar, adParamInput, 3, cStore)
cmdTable.Parameters.Append cmdTable.CreateParameter("item", adDouble, adParamInput, 20, cItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adDouble, adParamOutput)
cmdTable.CommandText = "RetItemIn"
Set obj = cmdTable.Execute
QItemIn = Val(cmdTable.Parameters(2).Value & "")
Set obj = Nothing
End Function
Sub GrModelLookupAll(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "Select GRMODEL ,GRDESCA FROM FILE1_10h WHERE GRMODEL IS NOT NULL "
Generalarray(2) = "GROUP BY GRMODEL ,GRDESCA order by GRMODEL "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "„Ã„Ê⁄…- »Ì«‰"
listarray(0, 1) = "(%%GRMODEL+GRDESCA%%)"


GrdArray(0, 0) = "„Ã„Ê⁄… "
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "»Ì«‰ «·„Ã„Ê⁄…"
GrdArray(1, 1) = 4000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „Ã„Ê⁄«  „ÊœÌ·« "
oSearch.Show 1
End Sub

Sub ModelLookupAll_bal(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "Select TOP 1000 FILE1_10h.MODEL ,FILE1_10h.DescA,File1_10h.MODELFACT0 ,FILE1_10h.MOSM ,FILE1_10h.SUPP , fact.DESCA   FROM FILE1_10h INNER JOIN fact ON FILE1_10h.FACT = fact.CODE  LEFT JOIN MOSM ON MOSM.MOSM = FILE1_10H.MOSM "
Generalarray(2) = "order by MOSM.DATE DESC  , FILE1_10H.MODEL "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ „ÊœÌ·"
listarray(0, 1) = "(%%File1_10h.MODELFACT0+FACT.DESCA%%)"

listarray(1, 0) = "«·„ÊœÌ·"
listarray(1, 1) = "(%%File1_10h.DESCA%%)"

'listarray(1, 0) = "„Ê”„ ° „Ê—œ "
'listarray(1, 1) = "(%%File1_10h.MOSM & FILE1_10h.SUPP%%)"

GrdArray(0, 0) = "„ÊœÌ·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·’‰› "
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·„Ê’„ "
GrdArray(3, 1) = 800

GrdArray(4, 0) = " „ﬂ » Ã„·…"
GrdArray(4, 1) = 800

GrdArray(5, 0) = " «·„’‰⁄"
GrdArray(5, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ „ÊœÌ·«  "
oSearch.Show 1
End Sub
Sub ModelLookupAll_Ser1(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "Select FILE1_10.MODEL , FILE1_10.ITEM , FILE1_10.DESCA , FILE8_51.desca  , FILE8_52.DESCA FROM FILE1_10  LEFT JOIN FILE8_51 ON FILE8_51.CODE = FILE1_10.CHARGE LEFT JOIN FILE8_52 ON FILE8_51.maingroup = FILE8_52.CODE WHERE FILE1_10.ISNOITEM = 1 "
Generalarray(2) = "order by FILE1_10.ITEM "
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%File1_10.DESCA%%)"

GrdArray(0, 0) = "„ÊœÌ·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "»«—ﬂÊœ"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "«·»Ì«‰"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "„’—Ê› "
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "„’—Ê› —∆Ì”Ì… "
GrdArray(4, 1) = 3000


searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «’‰«› Œœ„… "
oSearch.Show 1
End Sub
Sub ModelLookupAll_Ser2(oForm As Form, oSearch As Form)
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = oForm
'                           0               1               2                   3           4                   5               6
Generalarray(1) = "Select FILE1_10.MODEL , FILE1_10.ITEM , FILE1_10.DESCA , FILE8_51.desca  , FILE8_52.DESCA FROM FILE1_10  LEFT JOIN FILE8_51 ON FILE8_51.CODE = FILE1_10.CHARGE LEFT JOIN FILE8_52 ON FILE8_51.maingroup = FILE8_52.CODE WHERE FILE1_10.ISNOITEM2 = 1 "
Generalarray(2) = "order by FILE1_10.ITEM "
Generalarray(3) = 3000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%File1_10.DESCA%%)"

GrdArray(0, 0) = "„ÊœÌ·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "»«—ﬂÊœ"
GrdArray(1, 1) = 1000

GrdArray(2, 0) = "«·»Ì«‰"
GrdArray(2, 1) = 4000

GrdArray(3, 0) = "„’—Ê›"
GrdArray(3, 1) = 4000

GrdArray(4, 0) = "„’—Ê› —∆”ÌÏ"
GrdArray(4, 1) = 4000



searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ «’‰«› «’Ê· "
oSearch.Show 1
End Sub
Function ItemFind_Fact(pModelFact0 As String, pScal As String, pCon As adodb.Connection) As adodb.Recordset
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("MODELFACT0", adChar, adParamInput, 20, pModelFact0)
cmdTable.Parameters.Append cmdTable.CreateParameter("SCAL", adChar, adParamInput, 10, pScal)
cmdTable.CommandText = "ItemFind_Fact"
Set rdTable = cmdTable.Execute
Set ItemFind_Fact = rdTable
Set rdTable = Nothing
End Function
Function ItemFind_Model_S_C(pModel, pCode_Color, pCode_Scal, pCon As adodb.Connection) As Double
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("MODEL", adVarWChar, adParamInput, 20, pModel)
cmdTable.Parameters.Append cmdTable.CreateParameter("CODE_COLOR", adInteger, adParamInput, 3, pCode_Color)
cmdTable.Parameters.Append cmdTable.CreateParameter("Code_Scal", adInteger, adParamInput, 3, pCode_Scal)
cmdTable.Parameters.Append cmdTable.CreateParameter("Ret", adInteger, adParamOutput)

cmdTable.CommandText = "ItemFind_Model_S_C"
Set obj = cmdTable.Execute
ItemFind_Model_S_C = Val(cmdTable.Parameters(3).Value & "")
Set obj = Nothing
End Function
Function Item_BARCODE(pBARCODE, pCon As adodb.Connection) As Double
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("BARCODE", adVarWChar, adParamInput, 30, pBARCODE)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adDouble, adParamOutput)
cmdTable.CommandText = "Item_BARCODE"
Set obj = cmdTable.Execute
Item_BARCODE = Val(cmdTable.Parameters(1).Value & "")
Set obj = Nothing
End Function
Function retDOC_NO(pDoc_no As String, pCon As adodb.Connection) As String
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("DOC_NO", adVarWChar, adParamInput, 15, pDoc_no)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adVarWChar, adParamOutput, 15)
cmdTable.CommandText = "retDOC_NO"
Set rdTable = cmdTable.Execute
retDOC_NO = cmdTable.Parameters(1).Value & ""
Set rdTable = Nothing
End Function
Function SALES_DOC_FIND(ByVal pDoc_no As String, pCon As adodb.Connection) As Variant
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("DOC_NO", adChar, adParamInput, 20, pDoc_no)
cmdTable.CommandText = "SALES_DOC_FIND"
Set rdTable = cmdTable.Execute
Set SALES_DOC_FIND = rdTable
Set rdTable = Nothing
End Function
Function ItemFind_BARCODE2(pItem As String, pCon As adodb.Connection) As adodb.Recordset
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
If Len(pItem) > 8 Then pItem = ""
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("BARCODE2", adChar, adParamInput, 20, pItem)
cmdTable.CommandText = "ItemFind_BARCODE2"
Set rdTable = cmdTable.Execute
Set ItemFind_BARCODE2 = rdTable
Set rdTable = Nothing
End Function
Function RetDEmItem(pItem As String, pstore As String, pCon As adodb.Connection) As Double
Dim obj As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("ITEM", adVarWChar, adParamInput, 15, pItem)
cmdTable.Parameters.Append cmdTable.CreateParameter("STORE", adVarWChar, adParamInput, 5, pstore)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adDouble, adParamOutput)
cmdTable.CommandText = "RetDEmItem"
Set obj = cmdTable.Execute
RetDEmItem = Val(cmdTable.Parameters(2).Value & "")
Set obj = Nothing
End Function
Function ISPRINTED_sales(pDoc_no As String, pCon As adodb.Connection) As Boolean
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("DOC_NO", adVarWChar, adParamInput, 15, pCode)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adBoolean, adParamOutput)
cmdTable.CommandText = "ISPRINTED_sales"
'Set rdTable = cmdTable.Execute
ISPRINTED_Pr = cmdTable.Parameters(1).Value
Set rdTable = Nothing
End Function
Function Ret_Phone(pPhone As String, pCon As adodb.Connection) As String
Dim rdTable As New adodb.Recordset
Dim cmdTable As New adodb.command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("PHONE", adVarWChar, adParamInput, 15, pPhone)
cmdTable.Parameters.Append cmdTable.CreateParameter("Return", adVarWChar, adParamOutput, 200)
cmdTable.CommandText = "Ret_Phone"
Set rdTable = cmdTable.Execute
Ret_Phone = cmdTable.Parameters(1).Value & ""
Set rdTable = Nothing
End Function
Function TransStock(pDoc_no As String, pClosed As Integer, con As adodb.Connection, Optional ByRef pErrorNumber As Variant, Optional ByRef pErrorMsg As Variant, Optional ByRef pUpdate As Long = 0, Optional ByRef pDelete As Long = 0, Optional ByRef pInsert As Long = 0) As Boolean
Dim cmd As New adodb.command
aPrm = AddFlag(Empty, "DOC_NO", pDoc_no)
aPrm = AddFlag(aPrm, "CLOSED", pClosed)
Set cmd = mycmdEx("dbo.sp_stock_trans", con, aPrm)

If Not IsNull(cmd.Parameters("@ERROR_NUMBER").Value) Then
     pErrorNumber = cmd.Parameters("@ERROR_NUMBER").Value
     pErrorMsg = cmd.Parameters("@ERROR_MSG").Value
     Exit Function
End If

pUpdate = cmd.Parameters("@COUNT_UPDATE")
pDelete = cmd.Parameters("@COUNT_DELETE")
pInsert = cmd.Parameters("@COUNT_INSERT")
TransStock = True
End Function
Sub PriceLookup(oForm As Form, oSearch As Form, Optional sControl As String = "", Optional sAddRow As String = "", Optional pwhere As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = oForm
Generalarray(1) = "Select 'PRICE','”⁄— „” Â·ﬂ' AS DESCA " & _
                  " UNION ALL " & _
                  "SELECT 'COST2','”⁄— «·Ã„·…'"
                  
If pwhere <> "" Then Generalarray(1) = Generalarray(1) & " WHERE " & pwhere
Generalarray(2) = ""
Generalarray(3) = 5000
Generalarray(5) = True

listarray(0, 0) = "«·”⁄—"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·”⁄—"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·”⁄—"
GrdArray(1, 1) = 4000

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.aAddRow = aRow
oSearch.bNoSort = True
oSearch.sCaption = "≈” ⁄·«„ «·«”⁄«—"
oSearch.sControl = sControl
oSearch.Show 1
End Sub
Function RetTxt(cSearch As String, Optional cFile As String) As String
Dim TextLine As String
On Error GoTo myerror

If cFile = "" Then
    cFile = cPathConf & "\CONF.TXT"
End If

Dim FileNumber
FileNumber = FreeFile()

Open cFile For Input As #FileNumber   ' Open file.

Do While Not EOF(FileNumber)   ' Loop until end of file.
   Line Input #FileNumber, TextLine   ' Read line into variable.
   If InStr(FileNumber, LCase(TextLine), LCase(cSearch) & "=") > 0 Then
       RetTxt = Mid(TextLine, Len(cSearch) + 2)
       Exit Do
   End If
Loop

Close #FileNumber   ' Close file.
Exit Function
myerror:
Err.Clear
RetTxt = ""
End Function
Sub SaveImageToAccess(pPic As Picture, rs As adodb.Recordset, pColName As String)
Dim pb As PropertyBag
Set pb = New PropertyBag
pb.WriteProperty "MyImage", pPic, 100
rs.Fields(pColName).AppendChunk pb.Contents
rs.Update
Set pb = Nothing
End Sub
Public Sub GetNewSize(ByVal OrigWidth As Long, ByVal OrigHeight As Long, _
                      ByRef NewWidth As Long, ByRef NewHeight As Long)
    
    Dim Ratio As Double
    
    ' ≈–«  „  ÕœÌœ «·⁄—÷ «·ÃœÌœ° ”‰Õ”» «·«— ›«⁄ »‰«¡ ⁄·ÌÂ
    If NewWidth > 0 Then
        Ratio = NewWidth / OrigWidth
        NewHeight = OrigHeight * Ratio
    ' ≈–«  „  ÕœÌœ «·«— ›«⁄ «·ÃœÌœ° ”‰Õ”» «·⁄—÷ »‰«¡ ⁄·ÌÂ
    ElseIf NewHeight > 0 Then
        Ratio = NewHeight / OrigHeight
        NewWidth = OrigWidth * Ratio
    End If
End Sub


