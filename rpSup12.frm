VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpSup12 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   1920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1920
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   420
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   1275
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   6180
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   990
         Width           =   1680
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1680
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   3345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   315
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   765
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "rpSup12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If TypeOf ActiveControl Is DBCombo And KeyCode = 46 Then ActiveControl.BoundText = ""
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
If Trim(xCode.text) = "" Then
    MsgBox "ﬂÊœ «·„Ê—œ €Ì— „”Ã·"
    Exit Function
End If
If Not IsDate(xdate2.text) And Trim(xdate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim nBalance As Double, nRow As Integer
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

If Trim(xCode.text) <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & " FILE4_11_2.code = " & MyParn(xCode.text)
End If
          
If IsDate(xDate1.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " FILE4_11_2.date < " & DateSq(xDate1.text)
    cField1 = "(Select Sum(FILE4_11_2.Sal -  FILE4_11_2.Pay) " & _
              " from FILE4_11_2 " & _
              turnFound(cWhere) & _
              cWhere & ") as FirstBalance"
Else
    cField1 = " 0 as FirstBalance"
End If


cWhere = ""
cString = "select FILE4_11_2.*,FILE4_12.desca as file4_12desca, " & _
          cField1 & _
          " From FILE4_11_2 Left join FILE4_12 on FILE4_11_2.type = FILE4_12.code"
If Trim(xCode.text) <> "" Then
    cWhere = cWhere & turn(cWhere, " AND ") & " FILE4_11_2.code = " & MyParn(xCode.text)
    aHeader(0) = "[" & "··„Ê—œ : " & xCodeDesca.Caption & "]"
End If
          
If IsDate(xDate1.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " FILE4_11_2.date >= " & DateSq(xDate1.text)
    aHeader(1) = "[" & BetweenString(xDate1.text, xdate2.text) & "]"
End If

If IsDate(xdate2.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " FILE4_11_2.date <= " & DateSq(xdate2.text)
    aHeader(1) = "[" & BetweenString(xDate1.text, xdate2.text) & "]"
End If

cString = cString & turn(cWhere, " where ") & cWhere
cString = cString & " Order by date,Doc_id"
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not sourcetable.EOF Then
    If Val(sourcetable!FirstBalance & "") <> 0 Then
        temptable.AddNew
        nBalance = Val(sourcetable!FirstBalance)
        nRow = 1
        temptable!str2 = "—’Ìœ ”«»ﬁ"
        If Val(sourcetable!FirstBalance & "") > 0 Then
            temptable!VAL1 = Val(sourcetable!FirstBalance & "")
        Else
            temptable!val2 = Abs(Val(sourcetable!FirstBalance & ""))
        End If
        temptable!val3 = nBalance
        temptable!VAL6 = nRow
        temptable!STR21 = TurnValue(retHeader(aHeader, 0, 3))
        temptable.Update
    End If
End If
Do Until sourcetable.EOF
    temptable.AddNew
    nBalance = nBalance + Val(sourcetable!SAL & "") - Val(sourcetable!PAY & "")
    nRow = nRow + 1
    temptable!Date1 = sourcetable!Date
    temptable!str1 = sourcetable!DOC_ID
    temptable!str2 = sourcetable!Desca
    temptable!VAL1 = sourcetable!PAY
    temptable!val2 = sourcetable!SAL
    temptable!val3 = nBalance
    temptable!VAL6 = nRow
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str2 = TurnValue(sourcetable!Desca, Null, "") & "  " & sourcetable!STORE
    
'    If sourcetable!Type = "4" Then
'        cStore = GetDesca("SELECT STORE FROM FILE7_20H WHERE DOC_NO = " & MyParn(sourcetable!doc_id), con)
'        temptable!str2 = sourcetable!file4_12desca & "  " & GetDesca("SELECT DESCA FROM FILE0_40 WHERE CODE = " & MyParn(cStore), con)
'    End If
'    If sourcetable!Type = "5" Then
'        cStore = GetDesca("SELECT STORE FROM FILE7_30H WHERE DOC_NO = " & MyParn(sourcetable!doc_id), con)
'        temptable!str2 = sourcetable!file4_12desca & "  " & GetDesca("SELECT DESCA FROM FILE0_40 WHERE CODE = " & MyParn(cStore), con)
'    End If
    
    temptable.Update
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\Sup12.rpt"
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub

Private Sub Form_Load()
openCon con
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    suplookup
End If
End Sub

Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""
If xCode.text = "" Then Exit Sub
xCodeDesca.Caption = GetDesca("select desca from FILE4_10 where code = " & MyParn(xCode.text), con) & ""
End Sub
Sub myProc()
    ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End Sub
Private Sub suplookup()
    Dim Generalarray(5)
    Dim listarray(0, 4)
    Dim GrdArray(1, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "Select Code, DescA From FILE4_10"
    Generalarray(2) = "Order by file4_10.Desca"
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
    listarray(0, 1) = "(%%DESCA%%) "
    
    GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "≈”„ «·„Ê—œ"
    GrdArray(1, 1) = 3000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "«” ⁄·«„"
    Search3.Show 1
End Sub

