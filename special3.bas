Attribute VB_Name = "speical3"
Public Sub StagesOnlineLook(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

cString = "SELECT STAGES_CODES.CODE," & _
          "STAGES_CODES.DESCA" & _
          " FROM STAGES_CODES"

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere

Generalarray(1) = cString

Generalarray(2) = "Order by STAGES_CODES.DESCA"
Generalarray(3) = 8000
Generalarray(5) = True

listarray(0, 0) = "«·„—Õ·…"
listarray(0, 1) = "(%%STAGES_CODES.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "««·„—Õ·…"
GrdArray(1, 1) = 6000


searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "STAGES_CODES.CODE")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ „—«Õ· «· ÃÂÌ“"
oSearch.Show 1
End Sub

Public Function retItemCode(pItem As String, con As ADODB.Connection) As String
If Trim(pItem) = "" Then Exit Function
retItemCode = myField("Select  [dbo].[fn_ret_item](" & MyParn(pItem) & ")", con) & ""
End Function
Public Function rsItemCode(pItem As String, Optional con As ADODB.Connection) As String
If Trim(pItem) = "" Then Exit Function
rsItemCode = rsValue("Select  [dbo].[fn_ret_item](" & MyParn(pItem) & ")", con) & ""
End Function
Public Function GetNumbersFromString(ByVal inputString As String) As String
    Dim i As Long
    Dim char As String
    Dim resultString As String

    For i = 1 To Len(inputString)
        char = Mid(inputString, i, 1)
        If IsNumeric(char) Then
            resultString = resultString & char
        End If
    Next i

    GetNumbersFromString = resultString
End Function
Public Function UpdateInvTotal(pDoc_no As String, con As ADODB.Connection) As Boolean
con.Execute "UPDATE FILE6_20H " & _
            " SET FILE6_20H.LATE = FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT" & _
            " FROM FILE6_20H" & _
            " WHERE FILE6_20H.DOC_NO = " & MyParn(pDoc_no)
End Function
Public Sub SalesOnlineLookup(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(4, 5)
Dim GrdArray(8, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

'                       0                   1
cString = "SELECT FILE6_20H.DOC_NO," & _
          "FILE6_20H.DOC_NO2," & _
          "CONVERT(VARCHAR(10),FILE6_20H.DATE,111)," & _
          "FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT," & _
          "FILE6_20H.[NAME]," & _
          "FILE6_20H.PHONE, " & _
          "FILE6_20H.SHIP_NO, " & _
          "SHIP.DESCA" & _
          " FROM  FILE6_20H " & _
          " INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE AND FILE0_40.ONLINE = 1" & _
          " LEFT JOIN SHIP ON FILE6_20H.SHIP = SHIP.DESCA" & _
          " WHERE FILE6_20H.ONLINE = 2"

'If cBranch <> "" Then
'    cString = cString & " AND FILE6_20H.BRANCH = " & MyParn(cBranch)
'End If

If cFilter <> "" Then cWhere = cWhere & Tr(cWhere) & cFilter
If cWhere <> "" Then cString = cString & " AND " & cWhere
Generalarray(1) = cString

Generalarray(2) = "Order by FILE6_20H.DATE DESC,FILE6_20H.DOC_NO"
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_20H.DOC_NO = 'cFilter' OR FILE6_20H.DOC_NO2 = 'cFilter')"

listarray(1, 0) = "«·«”„-«·⁄‰Ê«‰"
listarray(1, 1) = "(%%FILE6_20H.NAME%% OR %%FILE6_20H.ADDRESS%%)"

listarray(2, 0) = "—ﬁ„ «·»Ê·Ì’…"
listarray(2, 1) = "(SHIP_NO LIKE 'cFilter%')"

listarray(3, 0) = "«· ·Ì›Ê‰"
listarray(3, 1) = "(PHONE LIKE 'cFilter%')"

listarray(4, 0) = "«·‘—ﬂ…"
listarray(4, 1) = "(SHIP.CODE = 'cFilter')"
listarray(4, 2) = "SELECT  CODE,DESCA FROM SHIP"
listarray(4, 3) = "CODE"
listarray(4, 4) = "DESCA"


GrdArray(0, 0) = "ﬂÊœ «·„” ‰œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "—›„ «·›« Ê—…"
GrdArray(1, 1) = 1600

GrdArray(2, 0) = " «—ÌŒ «·›« Ê—…"
GrdArray(2, 1) = 1300

GrdArray(3, 0) = "«·≈Ã„«·Ì"
GrdArray(3, 1) = 1200

GrdArray(4, 0) = "«·«”„"
GrdArray(4, 1) = 2500

GrdArray(5, 0) = "«·⁄‰Ê«‰"
GrdArray(5, 1) = 3500

GrdArray(6, 0) = "«· ·Ì›Ê‰"
GrdArray(6, 1) = 1700

GrdArray(7, 0) = "—ﬁ„ «·»Ê·Ì’…"
GrdArray(7, 1) = 1400

GrdArray(8, 0) = "‘—ﬂ… «·‘Õ‰"
GrdArray(8, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE6_20H.DOC_NO")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ ›Ê« Ì— «Ê‰ ·«Ì‰"
oSearch.Show 1
End Sub
Public Sub SalesBalanceLook(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(3, 5)
Dim GrdArray(7, 1)
Dim cWhere As String
Set Generalarray(0) = oForm

Generalarray(1) = "SELECT " & _
                  " FILE6_20H.DOC_NO ," & _
                  " FORMAT(FILE6_20H.date,'yyyy/M/d')," & _
                  " FORMAT(FILE6_20H.date_Due,'yyyy/M/d')," & _
                  " FILE3_10.DESCA," & _
                  " FILE6_20H.TOTAL_ITEM - FILE6_20H.DISCOUNT," & _
                  " v.BALANCE," & _
                  " INV_TYPE_CODES.DESCA," & _
                  " FILE6_20H.NOTES" & _
                  " FROM FILE6_20H" & _
                  " INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE" & _
                  " INNER JOIN INV_TYPE_CODES ON FILE6_20H.INV_TYPE = INV_TYPE_CODES.CODE" & _
                  " INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE" & _
                  " LEFT JOIN FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE" & _
                  " LEFT JOIN vw_inv_Balance v ON FILE6_20H.DOC_NO = v.DOC_NO" & _
                  " WHERE COALESCE(V.BALANCE,0) <> 0 " & _
                  " AND FILE6_20H.DATE_DUE IS NOT NULL"
                  
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " AND " & cFilter
End If

Generalarray(2) = "Order by DATE DESC,DOC_NO2 DESC "
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·„” ‰œ"
listarray(0, 1) = "(FILE6_20H.DOC_NO = 'cFilter' OR FILE6_20H.DOC_NO2 = 'cFilter')"

listarray(1, 0) = "«· «—ÌŒ"
listarray(1, 1) = "(##FILE6_20H.DATE##)"

listarray(2, 0) = "«·⁄„Ì·"
listarray(2, 1) = "(%%FILE3_10.DESCA%%)"

listarray(3, 0) = "«·»Ì«‰"
listarray(3, 1) = "(%%FILE6_20H.NOTES%%)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1800

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1300

GrdArray(2, 0) = " «—ÌŒ «· Õ’Ì·"
GrdArray(2, 1) = 1300

GrdArray(3, 0) = "«·⁄„Ì·"
GrdArray(3, 1) = 3500

GrdArray(4, 0) = "«·ﬁÌ„…"
GrdArray(4, 1) = 1300

GrdArray(5, 0) = "«·—’Ìœ"
GrdArray(5, 1) = 1200

GrdArray(6, 0) = "‰Ê⁄ «·›« Ê—…"
GrdArray(6, 1) = 1500

GrdArray(7, 0) = "«·»Ì«‰"
GrdArray(7, 1) = 2500


searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE6_20H.DOC_NO")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ «—’œ… ›Ê« Ì—"
oSearch.Show 1
End Sub
Public Sub PriceListLook(oForm As Form, oSearch As Form, Optional cFilter As String = "", Optional bFilter As Boolean = False, Optional sAddRow As String = "")
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)
Set Generalarray(0) = oForm
Generalarray(1) = "SELECT  " & _
                  " FILE6_53H.DOC_NO ," & _
                  " FILE6_53H.NOTES," & _
                  " FILE3_10.DESCA" & _
                  " FROM FILE6_53H" & _
                  " INNER JOIN FILE3_10 ON FILE6_53H.CODE = FILE3_10.CODE"

If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & " WHERE " & cFilter
End If

Generalarray(2) = "Order by doc_no DESC"
Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "«·⁄„Ì·"
listarray(0, 1) = "(%%FILE3_10.DESCA%%)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "≈”„ «·ﬁ«∆„…"
GrdArray(1, 1) = 3500

GrdArray(2, 0) = "«·⁄„Ì·"
GrdArray(2, 1) = 6500

searchArray = Array(Generalarray, listarray, GrdArray)
If bFilter Then
    Dim aFilter As Variant
    aFilter = AddFlag(aFilter, "FILTER", True)
    aFilter = AddFlag(aFilter, "FIELD", "FILE6_53H.DOC_NO")
    oSearch.aFilter = aFilter
End If

Dim aRow As Variant
If sAddRow <> "" Then
    aRow = AddFlag(Empty, "text", sAddRow)
    aRow = AddFlag(aRow, "col", 1)
End If
oSearch.aAddRow = aRow

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.nMax_records = 1000
oSearch.Caption = "≈” ⁄·«„ ﬁÊ«∆„ «·«”⁄«—"
oSearch.Show 1
End Sub






