Attribute VB_Name = "toolkit2"
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
                    .TextMatrix(.Rows - 1, 1) = pform(I).Name & cIndex & ".text = " & "Format(CardTable" & "!" & cCardIndex & Mid(pform(I).Name, 2) & "," & myparn2("dd-mm-yyyy") & ")"
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
Sub MFocus(pform As Form)
With pform.VSFlexGrid1
.Rows = 0
cFile = App.Path & "\temp.txt"
Open cFile For Output As #1   ' Open file for output.
For I = 0 To pform.Count - 1
    If (TypeOf pform(I) Is TextBox Or TypeOf pform(I) Is DataCombo) Then
        Print #1, "Private Sub " & pform(I).Name & "_GotFocus()"
        Print #1, "mygotFocus " & pform(I).Name
        Print #1, "End Sub"
        
        Print #1, "Private Sub " & pform(I).Name & "_LostFocus()"
        Print #1, "myLostFocus " & pform(I).Name
        Print #1, "End Sub"
    End If
Next
Close #1
End With
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
On Error GoTo myError
Open pFileName For Input As #1    ' Open file.
Do While Not EOF(1)   ' Loop until end of file.
   Line Input #1, TextLine   ' Read line into variable.
   GetText = GetText & IIf(GetText = "", "", vbCrLf) & TextLine
Loop
Close #1   ' Close file.
Exit Function
myError:
'MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
GetText = ""
End Function
Function RetAdd(sTable As String, bcon As ADODB.Connection)
cFile = App.Path & "\temp.txt"
Dim loctable As New ADODB.Recordset
Dim nLast As Integer
nLast = 20
Open cFile For Output As #1   ' Open file for output.
Print #1, "' --- ADD SUB --------------------------"

loctable.Open sTable, bcon, adOpenStatic, adLockReadOnly, adCmdTable
cString1 = "insert into " & sTable & "( "
For I = 0 To loctable.Fields.Count - 1
    cString1 = cString1 & turnFound2(cString1, loctable.Fields(0).Name, ",", "") & loctable.Fields(I).Name
Next
cString1 = Chr(34) & cString1 & ")" & Chr(34) & " & _"
cString1 = "cString  = " & cString1

Print #1, cString1

Print #1, myparn2("Values (") & " & _"
For I = 0 To loctable.Fields.Count - 1
    If loctable.Fields(I).Type = 3 Then
       cField = "addValue(x" & loctable.Fields(I).Name & ".text)"
    ElseIf loctable.Fields(I).Type = 11 Then
       cField = "iif(x" & loctable.Fields(I).Name & ".value = 0," & Chr(34) & "FALSE" & Chr(34) & "," & Chr(34) & "TRUE" & Chr(34) & " )"
    ElseIf loctable.Fields(I).Type = 7 Then
        cField = "adddate(x" & loctable.Fields(I).Name & ".text)"
    Else
       cField = "addstring(x" & loctable.Fields(I).Name & ".text)"
    End If
    
    If I = nLast Then
        Print #1, ""
        cField = "cString2 = " & cField
    End If
    
    If I = nLast - 1 And I <> loctable.Fields.Count - 1 Then
        cField = cField & " & " & myparn2(",")
    Else
        cField = cField & IIf(I = loctable.Fields.Count - 1, "& _", " & " & myparn2(",") & " & _")
    End If
    Print #1, cField
Next
Print #1, myparn2(")")
If I > nLast Then Print #1, "cString = cString & cString2"


Print #1, "' --- EDIT SUB  --------------------------"
Print #1, "cString  = " & myparn2("update " & sTable & " Set") & "& _"

For I = 0 To loctable.Fields.Count - 1
    If loctable.Fields(I).Type = 3 Then
       cField = myparn2(loctable.Fields(I).Name & " = ") & " &  addValue(x" & loctable.Fields(I).Name & ".Text )"
    ElseIf loctable.Fields(I).Type = 11 Then
       cField = myparn2(loctable.Fields(I).Name & " = ") & " & iif(x" & loctable.Fields(I).Name & ".value = 0," & myparn2("FALSE") & "," & myparn2("TRUE") & " )"
    ElseIf loctable.Fields(I).Type = 7 Then
       cField = myparn2(loctable.Fields(I).Name & " = ") & " &  addDate(x" & loctable.Fields(I).Name & ".Text )"
    Else
       cField = myparn2(loctable.Fields(I).Name & " = ") & " & addstring(x" & loctable.Fields(I).Name & ".text)"
    End If

    If I = nLast Then
        Print #1, ""
        cField = "cString2 = " & cField
    End If
    
    If I = nLast - 1 And I <> loctable.Fields.Count - 1 Then
        cField = cField & " & " & myparn2(",")
    Else
        cField = cField & IIf(I = loctable.Fields.Count, "", " & " & myparn2(",") & " & _")
    End If
    Print #1, cField
Next
Print #1, myparn2(")")
If I > nLast Then Print #1, "cString = cString & cString2"

Close #1
cText = GetText(App.Path & "\temp.txt")
Clipboard.Clear
Clipboard.SetText cText
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
                Print #1, ".Text = Format(.Text," & myparn2("dd-mm-yyyy") & ")"
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
wNo.SelLength = Len(wNo.Text)
End Sub
Private Sub xDate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
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
