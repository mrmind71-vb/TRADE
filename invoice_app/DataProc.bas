Attribute VB_Name = "inv_DataProc"
Public Function myDateSql(pDate As String) As String
If IsDate(pDate) Then
    myDateSql = addstring(Format(pDate, "YYYY-MM-DD HH:NN"))
Else
    myDateSql = "null"
End If
End Function
Public Function mRound(ByVal nValue As Variant, Optional nRound As Integer = 2) As Double
Dim cString As String
cString = "##"
If nRound > 0 Then cString = cString & "." & String(nRound, "#")
mRound = Val(Format(nValue, cString))
End Function
Public Function ArString(aString As Variant, Optional pSep As String = ",", Optional nCount As Integer = -1) As String
Dim i As Integer
For i = 0 To IIf(nCount = -1, UBound(aString), nCount)
    ArString = ArString & IIf(ArString = "", "", pSep) & aString(i)
Next
End Function


