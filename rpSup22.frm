VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpSup22 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   3165
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
   ScaleHeight     =   3165
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   420
      Left            =   1395
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2385
      Width           =   1275
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2385
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   2265
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   6180
      Begin VB.OptionButton xin 
         Alignment       =   1  'Right Justify
         Caption         =   " — Ì» Õ”» «· «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1755
         Width           =   1860
      End
      Begin VB.OptionButton xin 
         Alignment       =   1  'Right Justify
         Caption         =   " — Ì» Õ”» «·—ﬁ„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1755
         Value           =   -1  'True
         Width           =   1860
      End
      Begin VB.TextBox xdoc1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   1680
      End
      Begin VB.TextBox xdoc2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   1680
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1350
         Width           =   1680
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   975
         Width           =   1680
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ ‘Ìﬂ"
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
         TabIndex        =   10
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈·Ï ‘Ìﬂ"
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
         TabIndex        =   9
         Top             =   630
         Width           =   735
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1395
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1035
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
Attribute VB_Name = "rpSup22"
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
Private Sub Form_Load()
openCon con

End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xdate2.Text) And Trim(xdate2.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString2 = "SELECT * from all_cash where CHECK_NO is not null "

If IsDate(xdate1.Text) Then
    cString2 = cString2 & turn(cString2) & " CHECK_date  >= " & DateSq(xdate1.Text)
End If

If IsDate(xdate2.Text) Then
    cString2 = cString2 & turn(cString2) & " CHECK_date <= " & DateSq(xdate2.Text)
End If

If xdoc1.Text <> "" Then
    cString2 = cString2 & turn(cString2) & " CHECK_no >= " & MyParn(xdoc1.Text)
End If

If xdoc2.Text <> "" Then
    cString2 = cString2 & turn(cString2) & " CHECK_no <= " & MyParn(xdoc2.Text)
End If

With sourcetable
sourcetable.Open cString2, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!DATE2 = TurnValue(sourcetable!Date, "", Null)
    temptable!str5 = TurnValue(sourcetable!doc_no, "", Null)
    temptable!str1 = TurnValue(sourcetable!DESCA, "", Null)
    temptable!str2 = TurnValue(sourcetable!Check_No, "", Null)
    temptable!str3 = TurnValue(sourcetable!desca2, "", Null)
    temptable!str4 = TurnValue(sourcetable!MOSM, "", Null)
    temptable!Date1 = sourcetable!Check_Date
    temptable!val1 = Val(!Value & "")
    temptable!str21 = TurnValue(retHeader(aHeader, 0, 3))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
temptable.Requery
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If xin(0).Value Then
    main.Report1.ReportFileName = App.Path & "\Reports\Sup22.rpt"
Else
    main.Report1.ReportFileName = App.Path & "\Reports\Sup220.rpt"
End If
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
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
If xCode.Text = "" Then Exit Sub
xCodeDesca.Caption = GetDesca("select desca from FILE4_10 where code = " & MyParn(xCode.Text)) & ""
End Sub
Sub myProc()
    ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
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

