Attribute VB_Name = "abd2"
Function StrList2(cString, Optional pCon As ADODB.Connection)
Dim listTable As New ADODB.Recordset
If pCon Is Nothing Then
    listTable.Open cString, GetCon, adOpenStatic, adLockReadOnly, adCmdText
Else
    listTable.Open cString, pCon, adOpenStatic, adLockReadOnly, adCmdText
End If
StrList2 = "#" & ";"
Do Until listTable.EOF
    StrList2 = StrList2 & "|#" & listTable.Fields(0) & ";" & listTable.Fields(1)
    listTable.MoveNext
Loop
listTable.Close
Set listable = Nothing
End Function
Function NextVisible(pGrid As Object, Row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
Dim nLast
For i = IIf(nBegincol = -1, 0, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, IIf(nEndCol > pGrid.Cols - 1, pGrid.Cols - 1, nEndCol))
    If pGrid.ColHidden(i) = False Or pGrid.ColWidth(i) = 0 Then
        NextVisible = i
        Exit Function
    End If
Next
NextVisible = IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
End Function
Sub MyEditItem(pGrid As Variant, Row As Long, Col As Long, Optional isCode As Boolean = False)
With pGrid
Exit Sub
If Trim(.TextMatrix(Row, Col)) <> "" And Row > 0 Then
    .TextMatrix(.Rows - 1, Col) = .TextMatrix(Row, Col)
    If isCode Then .TextMatrix(.Rows - 1, Col + 1) = .TextMatrix(Row, Col + 1)
End If
End With
End Sub
Public Function getId(con As ADODB.Connection) As Variant
Dim cmd As New ADODB.command
Set cmd = mycmdEx("dbo.sp_GET_ID", con)
getId = cmd.Parameters("@ID").Value
End Function
Public Sub myInform(Mcaption As String, Optional mCaption2 As String, Optional nInterval As Integer = 900)
On Error Resume Next
myInformfrm.sLabel1 = Mcaption
myInformfrm.sLabel2 = mCaption2
myInformfrm.nInterval = nInterval
myInformfrm.Show 1
DoEvents
Err.Clear
End Sub
Public Function ValidNum(ByVal pCode As Variant, Optional nLen As Integer = 0, Optional pZero As Boolean = False, Optional nMax As Integer = 18, Optional nMin As Integer = 1) As Boolean
Dim sNumber As String
If Trim(pCode & "") = "" Then Exit Function
pCode = Trim(pCode & "")

If nLen <> 0 Then
    If Len(pCode) <> nLen Then Exit Function
Else
    If Len(pCode & "") < nMin Then Exit Function
    If Len(pCode & "") > nMax Then Exit Function
    If Mid(pCode, 1, 1) = "0" And (pCode <> "0" Or Not pZero) Then Exit Function
End If

For i = 1 To Len(pCode)
    If Not IsNumeric(Mid(pCode, i, 1)) Then
        Exit Function
    End If
Next
ValidNum = True
End Function
Public Function ParentF(sPath As String) As String
ParentF = Left(sPath, InStrRev(sPath, "\") - 1)
End Function
Function validYear(nYear As String) As Boolean
If Not ValidInt(nYear) Then Exit Function
validYear = Val(nYear) >= 1990 And Val(nYear) <= 2060
End Function

Public Function retDimImg(pWidthImg As Long, pHeightImg As Long, pMaxWidth As Long, pMaxHeight As Long) As Variant
    Dim nRate As Double
    Dim finalWidth As Long
    Dim finalHeight As Long
    
    On Error GoTo myerror

    ' ÍÓÇÈ ÇáäÓÈÉ áßá ãä ÇáÚÑÖ æÇáØæá ãÞÇÑäÉ ÈÇáÍÌã ÇáÃÞÕì ÇáãÊÇÍ
    Dim ratioW As Double: ratioW = pMaxWidth / pWidthImg
    Dim ratioH As Double: ratioH = pMaxHeight / pHeightImg
    
    ' ÇÎÊíÇÑ ÇáäÓÈÉ ÇáÃÕÛÑ áÖãÇä ÇÍÊæÇÁ ÇáÕæÑÉ ÏÇÎá ÇáÅØÇÑ Ïæä ÊÔæíå
    If ratioW < ratioH Then
        nRate = ratioW
    Else
        nRate = ratioH
    End If
    
    finalWidth = Int(pWidthImg * nRate)
    finalHeight = Int(pHeightImg * nRate)

    ' ÇÓÊÎÏÇã ÇáÏÇáÉ ÇáÎÇÕÉ Èß áÊÎÒíä ÇáÞíã
    Dim result As Variant
    result = AddFlag(Empty, "width", finalWidth)
    retDimImg = AddFlag(result, "height", finalHeight)

    Exit Function

myerror:
    MsgBox "Error: " & Err.Description
    retDimImg = Empty
End Function

