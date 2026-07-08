Attribute VB_Name = "toolkit"
Sub makeMyLoad(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
Print #1, "Private Sub MyLoad()"
With pform.VSFlexGrid1
    .Rows = 0
    For I = 0 To pform.Count - 1
        If (TypeOf pform(I) Is TextBox) Or (TypeOf pform(I) Is DataCombo) Or (TypeOf pform(I) Is CheckBox) Then
            If TypeOf pform(I) Is TextBox Then
                If LCase(pform(I).Tag) = "date" Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
                    .TextMatrix(.Rows - 1, 1) = pform(I).Name & cIndex & ".text = " & "Format(CardTable" & "!" & cCardIndex & Mid(pform(I).Name, 2) & "," & myparn2("YYYY-MM-DD") & ")"
                Else
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
                    .TextMatrix(.Rows - 1, 1) = pform(I).Name & cIndex & ".text = " & " CardTable" & "!" & cCardIndex & Mid(pform(I).Name, 2) & " & " & retPar
                End If
            ElseIf TypeOf pform(I) Is CheckBox Then
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
                .TextMatrix(.Rows - 1, 1) = pform(I).Name & ".value = " & " iif(CardTable" & "!" & Mid(pform(I).Name, 2) & ",1,0)"
            Else
                .AddItem ""
                .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
                .TextMatrix(.Rows - 1, 1) = pform(I).Name & ".boundtext = " & " CardTable" & "!" & Mid(pform(I).Name, 2) & " & " & retPar
            End If
        End If
    Next
    .Select 1, 0
    .Sort = flexSortNumericAscending
    For i2 = 0 To .Rows - 1
        Print #1, .TextMatrix(i2, 1)
    Next
End With
Print #1, "End sub"
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub makeMyDefine(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
With pform.VSFlexGrid1
    .Rows = 0
    For I = 0 To pform.Count - 1
        If TypeOf pform(I) Is TextBox Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
            .TextMatrix(.Rows - 1, 1) = pform(I).Name & cIndex & ".text = " & retPar
        ElseIf TypeOf pform(I) Is CheckBox Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
            .TextMatrix(.Rows - 1, 1) = pform(I).Name & ".value = 0 "
        ElseIf TypeOf pform(I) Is DataCombo Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
            .TextMatrix(.Rows - 1, 1) = pform(I).Name & ".boundtext = " & retPar
        End If
    Next
    .Select 1, 0
    .Sort = flexSortNumericAscending
    For i2 = 0 To .Rows - 1
        Print #1, .TextMatrix(i2, 1)
    Next
End With
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub makeMyReplace(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
'Print #1, "Private Sub MyReplace()"

For I = 0 To pform.Count - 1
    If ((TypeOf pform(I) Is TextBox) Or (TypeOf pform(I) Is DataCombo) Or (TypeOf pform(I) Is CheckBox)) Then
        If LCase(pform(I).Tag) <> "ig" Then nCount = nCount + 1
    End If
Next

'Print #1, "Dim aInsert as Variant"
With pform.VSFlexGrid1
.Rows = 0
nOrder = 0
For I = 0 To pform.Count - 1
    If (TypeOf pform(I) Is TextBox) Or (TypeOf pform(I) Is DataCombo) Or (TypeOf pform(I) Is CheckBox) Then
        If TypeOf pform(I) Is TextBox And LCase(pform(I).Tag) <> "ig" Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(I).Name, 2))
            If LCase(pform(I).Tag) = "date" Then
                .TextMatrix(.Rows - 1, 2) = "addDate(" & pform(I).Name & cIndex & ".Text" & ")"
            ElseIf Trim(pform(I).Tag) = "N" Then
                .TextMatrix(.Rows - 1, 2) = "addValue(" & pform(I).Name & cIndex & ".Text" & ")"
            ElseIf Trim(pform(I).Tag) = "" Then
                .TextMatrix(.Rows - 1, 2) = "addString(" & pform(I).Name & cIndex & ".Text" & ")"
            Else
                .TextMatrix(.Rows - 1, 2) = "addString(" & pform(I).Name & cIndex & ".Text" & ")"
            End If
        ElseIf TypeOf pform(I) Is CheckBox Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(I).Name, 2))
            .TextMatrix(.Rows - 1, 2) = " iif( " & pform(I).Name & ".value = 0 , " & myparn2("FALSE") & "," & myparn2("True") & ")"
        ElseIf TypeOf pform(I) Is DataCombo Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(I).Name, 2))
            .TextMatrix(.Rows - 1, 2) = "addvalue(" & pform(I).Name & ".BoundText" & ")"
        End If
    End If
    nOrder = nOrder + 1
Next

.Select 1, 0
.Sort = flexSortNumericAscending
For i2 = 0 To .Rows - 1
    Print #1, "aInsert = addFlag(aInsert," & .TextMatrix(i2, 1) & "," & .TextMatrix(i2, 2) & ")"
Next
End With
Print #1, "End sub"
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub makeMyReplace2(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
Print #1, "Private Sub MyReplace()"

For I = 0 To pform.Count - 1
    If ((TypeOf pform(I) Is TextBox) Or (TypeOf pform(I) Is DataCombo) Or (TypeOf pform(I) Is CheckBox)) Then
        If LCase(pform(I).Tag) <> "ig" Then nCount = nCount + 1
    End If
Next

Print #1, "Dim aInsert(" & nCount - 1 & ",1)"
With pform.VSFlexGrid1
.Rows = 0
nOrder = 0
For I = 0 To pform.Count - 1
    If (TypeOf pform(I) Is TextBox) Or (TypeOf pform(I) Is DataCombo) Or (TypeOf pform(I) Is CheckBox) Then
        If TypeOf pform(I) Is TextBox And LCase(pform(I).Tag) <> "ig" Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(I).Name, 2))
            If LCase(pform(I).Tag) = "date" Then
                .TextMatrix(.Rows - 1, 2) = "addDate(" & pform(I).Name & cIndex & ".Text" & ")"
            ElseIf Trim(pform(I).Tag) = "N" Then
                .TextMatrix(.Rows - 1, 2) = "addValue(" & pform(I).Name & cIndex & ".Text" & ")"
            ElseIf Trim(pform(I).Tag) = "" Then
                .TextMatrix(.Rows - 1, 2) = "addString(" & pform(I).Name & cIndex & ".Text" & ")"
            Else
                .TextMatrix(.Rows - 1, 2) = "addString(" & pform(I).Name & cIndex & ".Text" & ")"
            End If
        ElseIf TypeOf pform(I) Is CheckBox Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(I).Name, 2))
            .TextMatrix(.Rows - 1, 2) = " iif( " & pform(I).Name & ".value = 0 , " & myparn2("FALSE") & "," & myparn2("True") & ")"
        ElseIf TypeOf pform(I) Is DataCombo Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).TabIndex
            .TextMatrix(.Rows - 1, 1) = myparn2(Mid(pform(I).Name, 2))
            .TextMatrix(.Rows - 1, 2) = "addvalue(" & pform(I).Name & ".BoundText" & ")"
        End If
    End If
    nOrder = nOrder + 1
Next

.Select 1, 0
.Sort = flexSortNumericAscending
For i2 = 0 To .Rows - 1
    Print #1, "aInsert(" & i2 & ",0) = " & .TextMatrix(i2, 1)
    Print #1, "aInsert(" & i2 & ",1) = " & .TextMatrix(i2, 2)
Next
End With
Print #1, "End sub"
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub MFocus(pform As Form, Optional sTag As String = "")
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For I = 0 To pform.Count - 1
    If (TypeOf pform(I) Is TextBox Or TypeOf pform(I) Is DataCombo) And (sTag = "" Or UCase(pform(I).Tag) = UCase(sTag)) Then
        Print #1, "Private Sub " & pform(I).Name & "_GotFocus()"
        Print #1, "mygotFocus " & pform(I).Name
        Print #1, "End Sub"
        
        Print #1, "Private Sub " & pform(I).Name & "_LostFocus()"
        Print #1, "myLostFocus " & pform(I).Name
        If UCase(pform(I).Tag) = "D" Or InStr(1, LCase(pform(I).Name), "date") Then
            Print #1, "myvalidDate " & pform(I).Name
        End If
        If TypeOf pform(I) Is DataCombo Then
            Print #1, "IF NOT " & pform(I).Name & ".MatchedWithList THEN " & pform(I).Name & ".BoundText = " & myparn2("")
        End If
        Print #1, "End Sub"
    End If
Next
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub LostFocus(pform As Form)
With pform.VSFlexGrid1
.Rows = 0
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For I = 0 To pform.Count - 1
    If (TypeOf pform(I) Is TextBox Or TypeOf pform(I) Is DataCombo) Then
        If .FindRow(pform(I).Name, , 0) = -1 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).Name
            Print #1, "Private Sub " & pform(I).Name & "_LostFocus()"
            Print #1, "myLostFocus " & pform(I).Name
            Print #1, "End Sub"
        End If
    End If
Next
Close #1
End With
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub MakeKeyDown(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For I = 0 To pform.Count - 1
    If (TypeOf pform(I) Is DataCombo) Then
        Print #1, "Private Sub " & pform(I).Name & "_KeyDown(KeyCode As Integer, Shift As Integer)"
        Print #1, "if KeyCode = 40 Then "
        Print #1, "    KeyCode = 0"
        Print #1, "    SendKeys " & myparn2("{TAB}")
        Print #1, "elseif KeyCode = 38 Then"
        Print #1, "    KeyCode = 0"
        Print #1, "    SendKeys " & myparn2("+{TAB}")
        Print #1, "End If"
        Print #1, "End Sub"
    ElseIf (TypeOf pform(I) Is TextBox) Then
        Print #1, "Private Sub " & pform(I).Name & "_KeyDown(" & IIf(pform(I).Index <> "", "Index As Integer", "") & "KeyCode As Integer, Shift As Integer)"
        Print #1, "if KeyCode = 40 Then  SendKeys " & myparn2("{TAB}")
        Print #1, "if KeyCode = 38 Then  SendKeys " & myparn2("+{TAB}")
        Print #1, "End Sub"
    End If
Next
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Sub makeMyMenu(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For I = 0 To pform.Count - 1
    Print #1, pform(I).Name
Next
Close #1
End Sub
Function retPar()
retPar = String(2, Chr(34))
End Function
Function GetText(Optional pFileName) As String
Dim TextLine
On Error GoTo myerror
Open pFileName For Input As #1    ' Open file.
Do While Not EOF(1)   ' Loop until end of file.
   Line Input #1, TextLine   ' Read line into variable.
   GetText = GetText & IIf(GetText = "", "", vbCrLf) & TextLine
Loop
Close #1   ' Close file.
Exit Function
myerror:
'MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
GetText = ""
End Function
Private Function myparn2(cString)
myparn2 = Chr(34) & cString & Chr(34)
End Function
Sub makeMyValidate(pform As Form, Optional withfocus As Boolean = False, Optional withValidate As Boolean = True, Optional withLost As Boolean = False)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
Print #1, String(50, "-")
If withValidate Then
    For I = 0 To pform.Count - 1
        If (TypeOf pform(I) Is DataCombo) Then
            Print #1, "Private Sub " & pform(I).Name & "_Validate(Cancel As Boolean)"
            Print #1, "if Not " & pform(I).Name & ".MatchedWithList Then " & pform(I).Name & ".BoundText = " & myparn2("")
            Print #1, "End Sub"
        ElseIf (TypeOf pform(I) Is TextBox) Then
            If LCase(pform(I).Tag) = "date" Then
                Print #1, "Private Sub " & pform(I).Name & "_Validate(Cancel As Boolean)"
                Print #1, "With " & pform(I).Name
                Print #1, "If (Not IsDate(.Text)) And Trim(.Text) <> " & myparn2("") & " Then "; ".text = " & myparn2("")
                'Print #1, vbTab & "Cancel = True"
               ' Print #1,
               ' Print #1, "End If"
                Print #1, ".Text = Format(.Text," & myparn2("YYYY-MM-DD") & ")"
                Print #1, "End With"
                Print #1, "End Sub"
            End If
        End If
    Next
End If

If withfocus Then
    Print #1, String(50, "-")
    For I = 0 To pform.Count - 1
        'If (TypeOf pform(I) Is DataCombo) Or TypeOf pform(I) Is TextBox Then
         If TypeOf pform(I) Is TextBox Then
            If pform(I).TabIndex > 2 Then
                Print #1, "Private Sub " & pform(I).Name & "_GotFocus()"
                'Print #1, "StatusBar1.Panels(2).Text = RetCap(Me.Name, ActiveControl.Name)"
                If TypeOf pform(I) Is TextBox Then
                    Print #1, pform(I).Name & ".SelStart = 0"
                    Print #1, pform(I).Name & ".SelLength = Len(" & pform(I).Name & ".text)"
                End If
                Print #1, "End Sub"
            End If
        End If
    Next
End If
Close #1

cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub

Private Sub wNo_GotFocus()
wNo.SelStart = 0
wNo.SelLength = Len(wNo.text)
End Sub
Private Sub xDate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.text)
End Sub
Sub mRetNumber(pform As Form)
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
With pform.VSFlexGrid1
.Rows = 0
For I = 0 To pform.Count - 1
    If .FindRow(pform(I).Name, , 0) = -1 Then
        If LCase(Left(pform(I).Name, 2)) = "m0" Or LCase(Left(pform(I).Name, 2)) = "m1" Or LCase(Left(pform(I).Name, 2)) = "m2" Then
            If pform(I).Index <> "" Then
                cIndex = "(" & pform(I).Index & ")"
            End If
        End If
        If (TypeOf pform(I) Is TextBox And pform(I).Tag <> "date") Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = pform(I).Name
            If cIndex <> "" Then
                Print #1, "Private Sub " & pform(I).Name & "_KeyPress(Index As Integer,KeyAscii As Integer)"
                Print #1, "if " & pform(I).Name & "(index).tag <> " & myparn2("dec") & " and lcase(" & pform(I).Name & "(index).tag) <>  " & myparn2("date") & " then  "
                Print #1, "if " & pform(I).Name & "(index).tag  = " & myparn2("dec"); " then " & "KeyAscii = RetNumber(KeyAscii,TRUE) else KeyAscii = RetNumber(KeyAscii)"
                Print #1, "end if"
            Else
                Print #1, "Private Sub " & pform(I).Name & "_KeyPress(KeyAscii As Integer)"
                Print #1, "KeyAscii = RetNumber(KeyAscii)"
            End If
            Print #1, "End Sub"
        End If
    End If
Next
Close #1
End With
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Function FoundOtheritem(grid1 As Variant, nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = nValue Then
            FoundOtheritem = I
            Exit Function
        End If
    End If
Next
End Function
Public Sub createAddInsert(sString, pCon As ADODB.Connection)
Dim cString As String, loctable As New ADODB.Recordset
loctable.Open sString, pCon, adOpenStatic, adLockReadOnly, adCmdText
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For I = 0 To loctable.Fields.Count - 1
    If loctable.Fields(I).Type = adDate Or loctable.Fields(I).Type = adDBTimeStamp Then
        Print #1, "aInsert = addflag(aInsert," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & ",adddate(loctable![" & loctable.Fields(I).Name & "]))"
    ElseIf loctable.Fields(I).Type = adInteger Then
        Print #1, "aInsert = addflag(aInsert," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & ",addvalue(loctable![" & loctable.Fields(I).Name & "]))"
    ElseIf loctable.Fields(I).Type = adDecimal Or loctable.Fields(I).Type = adNumeric Then
        Print #1, "aInsert = addflag(aInsert," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & ",val(loctable![" & loctable.Fields(I).Name & "] & " & Chr(34) & Chr(34) & "))"
    ElseIf loctable.Fields(I).Type = adBoolean Then
        Print #1, "aInsert = addflag(aInsert," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & ",iif(loctable![" & loctable.Fields(I).Name & "],1,0))"
    Else
        Print #1, "aInsert = addflag(aInsert," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & ",addstring(loctable![" & loctable.Fields(I).Name & "]))"
    End If
Next
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Public Sub createAddInsertTable(sString, pCon As ADODB.Connection)
Dim cString As String, loctable As New ADODB.Recordset
loctable.Open sString, pCon, adOpenStatic, adLockReadOnly, adCmdText
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For I = 0 To loctable.Fields.Count - 1
    If loctable.Fields(I).Type = adDate Or loctable.Fields(I).Type = adDBTimeStamp Then
        Print #1, "aInsert = addflag(aInsert," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & "," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & ")"
    ElseIf loctable.Fields(I).Type = adInteger Then
        Print #1, "aInsert = addflag(aInsert," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & "," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & ")"
    ElseIf loctable.Fields(I).Type = adDecimal Or loctable.Fields(I).Type = adNumeric Then
        Print #1, "aInsert = addflag(aInsert," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & "," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & ")"
    ElseIf loctable.Fields(I).Type = adBoolean Then
        Print #1, "aInsert = addflag(aInsert," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & "," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & ")"
    Else
        Print #1, "aInsert = addflag(aInsert," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & "," & Chr(34) & "[" & loctable.Fields(I).Name & "]" & Chr(34) & ")"
    End If
Next
Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
End Sub
Public Function retInsertTable(Optional pFile As String, Optional pTable As String = "locTable") As Boolean
Dim sb As New ChilkatStringBuilder
Dim strTab As New ChilkatStringTable
If pFile = "" Then
    pFile = App.Path & "\sql_data\insert.sql"
End If
success = sb.LoadFile(pFile, "utf-8")
sucsess = sb.Trim
sMarker = vbCrLf
Do Until sb.Length = 0
    cLine = Replace(Trim(sb.GetBefore(sMarker, True)), ",<", "")
    cLine = Replace(cLine, "<", "")
    cLine = Replace(cLine, ",>", "")
    cLine = Replace(cLine, ">", "")
    
    nFound = InStr(1, cLine, ",")
    pFieldtable = "[" & Trim(Mid(cLine, 1, nFound - 1)) & "]"
    pField = Trim(Mid(cLine, 1, nFound - 1))
    pType = Trim(Mid(cLine, nFound + 1))
    
    If pType = "bit" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",iif(" & pTable & "!" & pField & ",1,0))"
    ElseIf Mid(pType, 1, 7) = "decimal" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",mround(" & pTable & "!" & pField & "))"
    ElseIf LCase(Mid(pType, 1, 8)) = "nvarchar" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",addstring(" & pTable & "!" & pField & "))"
    ElseIf Mid(pType, 1, 8) = "nvarchar" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",addstring(" & pTable & "!" & pField & "))"
    ElseIf Mid(pType, 1, 7) = "varchar" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",addstring(" & pTable & "!" & pField & "))"
    ElseIf pType = "uniqueidentifier" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",addstring(" & pTable & "!" & pField & "))"
    ElseIf pType = "float" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",mround(" & pTable & "!" & pField & ",4))"
    ElseIf Left(pType, 7) = "numeric" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",mround(" & pTable & "!" & pField & ",4))"
    ElseIf Left(pType, 3) = "int" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",addvalue(" & pTable & "!" & pField & "))"
    ElseIf Left(pType, 8) = "smallInt" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",addvalue(" & pTable & "!" & pField & "))"
    ElseIf Mid(pType, 1, 4) = "date" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",adddate(" & pTable & "!" & pField & "))"
    ElseIf pType = "smalldatetime" Then
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",adddate(" & pTable & "!" & pField & "))"
    Else
        strTab.Append "aInsert = addflag(aInsert," & """" & pFieldtable & """" & ",addError_" & pType & "(" & pTable & "!" & pField & "))"
    End If
Loop
'success = strTab.SaveToFile("utf-8", 1, App.Path & "\toolkit\sql2.sql")
'success = strTab.SaveToFile("utf-8", 1, App.Path & "\toolkit\sql2.sql")
Clipboard.Clear
Clipboard.SetText strTab.GetStrings(0, 0, 0)
End Function
Sub myValidDate2(ByRef pControl As Variant)
With pControl
If IsDate(.text) Then
    .text = myFormat_p(.text)
Else
    .text = ""
End If
End With
End Sub

