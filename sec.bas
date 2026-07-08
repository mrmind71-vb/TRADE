Attribute VB_Name = "sec"
Public bopt1 As Boolean, bopt2 As Boolean, bopt3 As Boolean, bOpt4 As Boolean, bOpt5 As Boolean, bOpt6 As Boolean, bOpt7 As Boolean, bOpt8 As Boolean, bOpt9 As Boolean, bOpt10 As Boolean, bOpt11 As Boolean, cPassDisc As String
Public bCrypt As Boolean
Public sectable As New ADODB.Recordset, MainPath As String
Dim Units(11), One(10), Tens(10), Hundred(10)
Public nPaidYear As Integer, nLateDays As Integer

Function MyOnly(pNumber)
Units(0) = ""
Units(1) = "«ÕœÌ"
Units(2) = "«À‰«"
Units(3) = "À·«À…"
Units(4) = "√—»⁄…"
Units(5) = "Œ„”…"
Units(6) = "” …"
Units(7) = "”»⁄…"
Units(8) = "À„«‰Ì…"
Units(9) = " ”⁄…"
Units(10) = "⁄‘—…"

One(1) = "Ê«Õœ"
One(2) = "√À‰Ì‰"
One(3) = "À·«À…"
One(4) = "√—»⁄…"
One(5) = "Œ„”…"
One(6) = "” …"
One(7) = "”»⁄…"
One(8) = "À„«‰Ì…"
One(9) = " ”⁄…"

Tens(0) = ""
Tens(1) = ""
Tens(2) = "⁄‘—Ê‰"
Tens(3) = "À·«ÀÊ‰"
Tens(4) = "√—»⁄Ê‰"
Tens(5) = "Œ„”Ê‰"
Tens(6) = "” Ê‰"
Tens(7) = "”»⁄Ê‰"
Tens(8) = "À„«‰Ê‰"
Tens(9) = " ”⁄Ê‰"

Hundred(0) = ""
Hundred(1) = "„«∆…"
Hundred(2) = "„«∆ Ì‰"
Hundred(3) = "À·À„«∆…"
Hundred(4) = "√—»⁄„«∆…"
Hundred(5) = "Œ„”„«∆…"
Hundred(6) = "” „«∆…"
Hundred(7) = "”»⁄„«∆…"
Hundred(8) = "À„‰„«∆…"
Hundred(9) = " ”⁄„«∆…"
If pNumber <> Int(pNumber) Then
    nDec = Format(pNumber, "00.00")
    nDec = Val(Right(nDec, 2))
End If
nNumber = Int(pNumber)
Select Case Len(nNumber)
Case 1
    MyOnly = Ret1(nNumber)
Case 2
    MyOnly = Ret2(nNumber)
Case 3
    MyOnly = Ret3(nNumber)
Case 4
    MyOnly = Ret4(nNumber)
Case 5
    MyOnly = Ret5(nNumber)
Case 6
    MyOnly = Ret6(nNumber)
Case 7
    MyOnly = Ret7(nNumber)
Case 8
    MyOnly = Ret8(nNumber)
End Select
MyOnly = "›ﬁÿ Êﬁœ—Â " & MyOnly & " Ã‰ÌÂ „’—Ì "
If Val(nDec) > 0 Then
    Select Case Len(nDec)
    Case 1
        MyOnly = MyOnly & "Ê" & Ret1(nDec) & IIf(nDec = 1, " ﬁ—‘", "ﬁ—Ê‘")
    Case 2
        MyOnly = MyOnly & " Ê" & Ret2(nDec) & IIf(nDec = 10, " ﬁ—Ê‘", " ﬁ—‘«")
    End Select
    MyOnly = MyOnly & " ·« €Ì— "
End If
'If (nNumber < 0) Then
'   bMinus = True
'End If
'nNumber = Str(nNumber)

'nLen = Len(nNumber)
'If Len(nNumber) = 1 Then cString = One(nNumber)
'
'If Len(nNumber) = 2 Then
'    If Val(nNumber) < 20 Then
'        cString = Units(Left(nNumber, 1)) & "⁄‘—… "
'    Else
'        cString = Units(Right(nNumber, 1)) & "Ê" & Tens(Left(nNumber, 1))
'     End If
'End If
'
'If Len(nNumber) = 4 Then
'    If Left(Number, 1) = 1 Then
'        cString = "√·›"
'    ElseIf Left(Number, 1) = 2 Then
'        cString = "«·›Ì‰"
'    Else
'        cString = Units(1) & "¬·«›"
'    End If
'    nNumberSub = Left()
'    If Val(nNumberSub) < 10 Then
'        cStringSub = "Ê" & One(nNumberSub)
'    ElseIf Val(nNumberSub) < 20 Then
'        cStringSub = Units(Left(Val(nNumberSub), 1)) & "⁄‘—… "
'    Else
'        cStringSub = Units(Right(nNumberSub, 1)) & "Ê" & Tens(Left(nNumberSub, 1))
'    End If
'End If

End Function
Function Ret1(nNumber)
nNumber = Left(nNumber, 1)
Ret1 = One(nNumber)
End Function
Function Ret2(nNumber)
nNumber = Val(nNumber)
If Val(nNumber) < 20 Then
    Ret2 = Units(Right(nNumber, 1)) & " ⁄‘—… "
Else
    Ret2 = Units(Right(nNumber, 1)) & _
    IIf(Units(Right(nNumber, 1)) = "", "", " Ê") & _
    Tens(Left(nNumber, 1))
End If
End Function
Function Ret3(nNumber)
If Len(nNumber) = 3 Then
    cString = Hundred(Val(Left(nNumber, 1)))
    nNumber = Val(Right(nNumber, 2))
    Select Case Len(nNumber)
    Case 1
        Ret3 = cString & addwaw(Ret1(nNumber))
    Case 2
        Ret3 = cString & addwaw(Ret2(nNumber))
    End Select
End If
End Function
Function Ret4(nNumber)
nNumber = Val(nNumber)
If Left(nNumber, 1) = 1 Then
    cString = "√·›"
ElseIf Left(nNumber, 1) = 2 Then
    cString = "«·›Ì‰"
Else
    cString = Ret1(Left(nNumber, 1)) & " ¬·«› "
End If
nNumber = Val(Right(nNumber, 3))
Select Case Len(nNumber)
Case 1
    Ret4 = cString & addwaw(Ret1(nNumber))
Case 2
    Ret4 = cString & addwaw(Ret2(nNumber))
Case 3
    Ret4 = cString & addwaw(Ret3(nNumber))
End Select
End Function
Function Ret5(nNumber)
If Left(nNumber, 2) = 10 Then
    cString = "⁄‘—… ¬·«›"
Else
    cString = Ret2(Left(nNumber, 2)) & " √·› "
End If
nNumber = Val(Right(nNumber, 3))
Select Case Len(nNumber)
Case 1
    Ret5 = cString & addwaw(Ret1(nNumber))
Case 2
    Ret5 = cString & addwaw(Ret2(nNumber))
Case 3
    Ret5 = cString & addwaw(Ret3(nNumber))
End Select
End Function
Function Ret6(nNumber)
cString = Ret3(Left(nNumber, 3)) & " √·› "
nNumber = Val(Right(nNumber, 3))
Select Case Len(nNumber)
Case 1
    Ret6 = cString & addwaw(Ret1(nNumber))
Case 2
    Ret6 = cString & addwaw(Ret2(nNumber))
Case 3
    Ret6 = cString & addwaw(Ret3(nNumber))
End Select
End Function
Function Ret7(nNumber)
If Left(Ret7, 1) = 1 Then
    cString = "„·ÌÊ‰"
Else
    cString = Ret1(Left(nNumber, 1)) & " „·ÌÊ‰ "
End If

nNumber = Val(Right(nNumber, 6))
Select Case Len(nNumber)
Case 1
    Ret7 = cString & addwaw(Ret1(nNumber))
Case 2
    Ret7 = cString & addwaw(Ret2(nNumber))
Case 3
    Ret7 = cString & addwaw(Ret3(nNumber))
Case 4
    Ret7 = cString & addwaw(Ret4(nNumber))
Case 5
    Ret7 = cString & addwaw(Ret5(nNumber))
Case 6
    Ret7 = cString & addwaw(Ret6(nNumber))
End Select
End Function
Function Ret8(nNumber)
If Left(nNumber, 2) = 10 Then
    cString = "⁄‘—… „·«ÌÌ‰"
Else
    cString = Ret2(Left(nNumber, 2)) & " „·ÌÊ‰"
End If

nNumber = Val(Right(nNumber, 6))
Select Case Len(nNumber)
Case 1
    Ret8 = cString & addwaw(Ret1(nNumber))
Case 2
    Ret8 = cString & addwaw(Ret2(nNumber))
Case 3
    Ret8 = cString & addwaw(Ret3(nNumber))
Case 4
    Ret8 = cString & addwaw(Ret4(nNumber))
Case 5
    Ret8 = cString & addwaw(Ret5(nNumber))
Case 6
    Ret8 = cString & addwaw(Ret6(nNumber))
End Select
End Function
Function addwaw(cString)
addwaw = IIf(cString = "", "", " Ê" & cString)
End Function
Function MyEmpty(cString) As Boolean
MyEmpty = (Trim(cString) = "")
End Function
Function arbDay(dDate)
Dim aWeek(6)
aWeek(0) = "«·Ã„⁄…"
aWeek(1) = "«·”» "
aWeek(2) = "«·√Õœ"
aWeek(3) = "«·≈À‰Ì‰"
aWeek(4) = "«·À·«À«¡"
aWeek(5) = "«·«—»⁄«¡"
aWeek(6) = "«·Œ„Ì”"
arbDay = aWeek(Weekday(dDate, vbFriday) - 1)
End Function
Function arbMonth(nMonth)
Dim aMonth(11)
aMonth(0) = "Ì‰«Ì—"
aMonth(1) = "›»—«Ì—"
aMonth(2) = "„«—”"
aMonth(3) = "≈»—Ì·"
aMonth(4) = "„«ÌÊ"
aMonth(5) = "ÌÊ‰ÌÊ"
aMonth(6) = "ÌÊ·ÌÊ"
aMonth(7) = "√€”ÿ”"
aMonth(8) = "” „»—"
aMonth(9) = "√ﬂ Ê»—"
aMonth(10) = "‰Ê›„»—"
aMonth(11) = "œÌ”„»—"
arbMonth = aMonth(nMonth - 1)
End Function
Function DateString(DDate1, DDate2) As String
nMonth = DateDiff("m", DDate1, DDate2)
nYears = nMonth / 12
nYears = Fix(nYears)
nMonth = nMonth Mod 12
If (nYears) > 0 Then
   DateString = nYears & " ”‰… "
End If
If nMonth > 0 Then
   DateString = DateString & "  " & nMonth & " ‘Â— "
End If
End Function
Function turnFound2(cString, Optional strFind As String = " WHERE ", Optional caseFound, Optional CaseNotfound)
If strFind <> "" Then
    If cString = "" Then
        turnFound2 = ""
        Exit Function
    End If
    If IsMissing(caseFound) Then caseFound = " AND "
    If IsMissing(CaseNotfound) Then CaseNotfound = " WHERE "
    turnFound2 = IIf(InStr(1, UCase(cString), UCase(strFind)) > 0, caseFound, CaseNotfound)
ElseIf IsMissing(caseFound) And IsMissing(CaseNotfound) Then
    turnFound2 = IIf(Trim(cString) = "", "", " and ")
ElseIf IsMissing(caseFound) And Not IsMissing(CaseNotfound) Then
    turnFound2 = IIf(Trim(cString) = "", "", CaseNotfound)
ElseIf (Not IsMissing(caseFound)) And IsMissing(CaseNotfound) Then
    turnFound2 = IIf(Trim(cString) = "", "", caseFound)
ElseIf (Not IsMissing(caseFound)) And (Not IsMissing(CaseNotfound)) Then
    turnFound2 = IIf(Trim(cString) = "", caseFound, CaseNotfound)
End If
End Function
Function turnFound(cString, Optional strFind As String = " WHERE ", Optional caseFound = " AND ")
If Trim(cString) <> "" Then
    If Trim(LCase(strFind)) = "where" Or Trim(LCase(strFind)) = "having" Then
        turnFound = IIf(InStr(1, UCase(cString), UCase(strFind)) > 0, caseFound, strFind)
    Else
        turnFound = strFind
    End If
End If
End Function

Function TurnAnd(cString) As String
TurnAnd = IIf(Trim(cString) = "", "", " and ")
End Function
Function TurnWhere(cString) As String
TurnWhere = IIf(Trim(cString) = "", "", " where ")
End Function
Function TurnOr(cString) As String
TurnOr = IIf(Trim(cString) = "", "", " or ")
End Function

Function ArabicDate(dDate) As String
ArabicDate = ArabicDay(dDate) & " " & Format(dDate, "yyyy/m/d")
End Function
Function ArabicDay(dDate) As String
Dim aDay(6)
If Not IsDate(dDate) Then Exit Function
aDay(0) = "«·Ã„⁄Â"
aDay(1) = "«·”» "
aDay(2) = "«·«Õœ"
aDay(3) = "«·«À‰Ì‰"
aDay(4) = "«·À·«À«¡"
aDay(5) = "«·«—»⁄«¡"
aDay(6) = "«·Œ„Ì”"
ArabicDay = aDay(Weekday(dDate, vbFriday) - 1)
End Function
Public Function crypt(ByVal inp As String, key As String) As String
Dim Sbox(0 To 255) As Long
Dim Sbox2(0 To 255) As Long
Dim j As Long, i As Long, t As Double
Dim K As Long, temp As Long, Outp As String
For i = 0 To 255
    Sbox(i) = i
Next i

j = 1
For i = 0 To 255
If j > Len(key) Then j = 1
Sbox2(i) = Asc(Mid(key, j, 1))
j = j + 1
Next i

j = 0
For i = 0 To 255
j = (j + Sbox(i) + Sbox2(i)) Mod 256
temp = Sbox(i)
Sbox(i) = Sbox(j)
Sbox(j) = temp
Next i

i = 0
j = 0
For X = 1 To Len(inp)
     i = (i + 1) Mod 256
     j = (j + Sbox(i)) Mod 256
     temp = Sbox(i)
    Sbox(i) = Sbox(j)
    Sbox(j) = temp
    t = (Sbox(i) + Sbox(j)) Mod 256
    K = Sbox(t)
    Outp = Outp + Chr(Asc(Mid(inp, X, 1)) Xor K)
Next X
crypt = Outp
crypt = crypt2(crypt)
crypt = StrReverse(crypt)
End Function
Public Function decrypt(ByVal inp As String, key As String) As String
Dim Sbox(0 To 255) As Long
Dim Sbox2(0 To 255) As Long
Dim j As Long, i As Long, t As Double
Dim K As Long, temp As Long, Outp As String
inp = StrReverse(inp)
inp = decrypt2(inp)

For i = 0 To 255
    Sbox(i) = i
Next i
 
 j = 1
For i = 0 To 255
    If j > Len(key) Then j = 1
    Sbox2(i) = Asc(Mid(key, j, 1))
    j = j + 1
Next i

 j = 0
 For i = 0 To 255
    j = (j + Sbox(i) + Sbox2(i)) Mod 256
    temp = Sbox(i)
    Sbox(i) = Sbox(j)
    Sbox(j) = temp
Next i

 i = 0
 j = 0
For X = 1 To Len(inp)
         i = (i + 1) Mod 256
         j = (j + Sbox(i)) Mod 256
         temp = Sbox(i)
        Sbox(i) = Sbox(j)
        Sbox(j) = temp
        t = (Sbox(i) + Sbox(j)) Mod 256
        K = Sbox(t)
        Outp = Outp + Chr(Asc(Mid(inp, X, 1)) Xor K)
Next X
decrypt = Outp
End Function
Private Function crypt2(ByVal cString) As String
'nLen = Len(cString)
'For I = 1 To nLen
'    crypt2 = crypt2 & RetZero(Asc(Mid(cString, I, 1)), 3)
'Next
crypt2 = StringToHex(cString)
End Function
Function decrypt2(ByVal cString) As String
'nLen = Len(cString)
'For I = 1 To nLen Step 3
'    decrypt2 = decrypt2 + Chr(Val(Mid(cString, I, 3)))
'Next
decrypt2 = HexToString(cString)
End Function
Function MyCode(nCode) As Double
Dim nDec
MyCode = nCode + 50
MyCode = MyCode & "2"
MyCode = MyCode / 2
MyCode = MyCode - 500
MyCode = Val(MyCode + RetDec(MyCode))
End Function
Function unMyCode(ByVal nCode)
unMyCode = Fix(nCode)
unMyCode = unMyCode + 500
unMyCode = unMyCode * 2
unMyCode = Mid(unMyCode, 1, Len(unMyCode) - 1)
unMyCode = unMyCode - 50
End Function
Public Function CRYPTDATA(ByVal inp As String, key As String) As String
Dim Sbox(0 To 255) As Long
Dim Sbox2(0 To 255) As Long
Dim j As Long, i As Long, t As Double
Dim K As Long, temp As Long, Outp As String
For i = 0 To 255
    Sbox(i) = i
Next i

j = 1
For i = 0 To 255
If j > Len(key) Then j = 1
Sbox2(i) = Asc(Mid(key, j, 1))
j = j + 1
Next i

j = 0
For i = 0 To 255
j = (j + Sbox(i) + Sbox2(i)) Mod 256
temp = Sbox(i)
Sbox(i) = Sbox(j)
Sbox(j) = temp
Next i

i = 0
j = 0
For X = 1 To Len(inp)
     i = (i + 1) Mod 256
     j = (j + Sbox(i)) Mod 256
     temp = Sbox(i)
    Sbox(i) = Sbox(j)
    Sbox(j) = temp
    t = (Sbox(i) + Sbox(j)) Mod 256
    K = Sbox(t)
    Outp = Outp + Chr(Asc(Mid(inp, X, 1)) Xor K)
Next X
CRYPTDATA = Outp
End Function
Public Function deCRYPTDATA(ByVal inp As String, key As String) As String
Dim Sbox(0 To 255) As Long
Dim Sbox2(0 To 255) As Long
Dim j As Long, i As Long, t As Double
Dim K As Long, temp As Long, Outp As String

For i = 0 To 255
    Sbox(i) = i
Next i
 
 j = 1
For i = 0 To 255
    If j > Len(key) Then j = 1
    Sbox2(i) = Asc(Mid(key, j, 1))
    j = j + 1
Next i

 j = 0
 For i = 0 To 255
    j = (j + Sbox(i) + Sbox2(i)) Mod 256
    temp = Sbox(i)
    Sbox(i) = Sbox(j)
    Sbox(j) = temp
Next i

 i = 0
 j = 0
For X = 1 To Len(inp)
         i = (i + 1) Mod 256
         j = (j + Sbox(i)) Mod 256
         temp = Sbox(i)
        Sbox(i) = Sbox(j)
        Sbox(j) = temp
        t = (Sbox(i) + Sbox(j)) Mod 256
        K = Sbox(t)
        Outp = Outp + Chr(Asc(Mid(inp, X, 1)) Xor K)
Next X
deCRYPTDATA = Outp
End Function
Private Function HexToString(ByVal strData As String) As String
    Dim strOutput As String
    Do Until Len(strData) < 2
        strOutput = strOutput + Chr$(CLng("&H" + Left$(strData, 2)))
        strData = Right$(strData, Len(strData) - 2)
    Loop
    HexToString = strOutput
End Function
Private Function StringToHex(ByVal strData As String) As String
Dim strOutput As String
Do Until Len(strData) = 0
    strOutput = strOutput + Right$(String$(2, "0") + Hex$(Asc(Left$(strData, 1))), 2)
    strData = Right$(strData, Len(strData) - 1)
Loop
StringToHex = strOutput
End Function
Function RetDec(cString)
Dim nDec, nten
For i = 1 To Len(cString)
    nDec = nDec + Val(Mid(cString, i, 1)) * (i + nMyCode3)
Next
RetDec = nDec / ("1" & String(Len(nDec), "0"))
End Function


Function addTime(pValue) As String
If Not IsDate(pValue) Then
    addTime = "NULL"
Else
    addTime = "Convert(DateTime," & MyParn(pValue) & ",108)"
End If
End Function

