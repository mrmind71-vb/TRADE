VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpCharge8 
   Caption         =   " ﬁ«—Ì— «·‰ﬁœÌ…"
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
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
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2340
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   5145
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2325
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2325
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xbox 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   945
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMan 
         Height          =   315
         Left            =   375
         TabIndex        =   9
         Top             =   1350
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "‘—Ìﬂ"
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·Ï  «—ÌŒ :"
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Œ“‰… :"
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1035
         Width           =   465
      End
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1395
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1890
      Width           =   1320
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1890
      Width           =   1320
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3315
      Top             =   2535
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   3810
      Top             =   2160
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   3840
      Top             =   2520
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "rpCharge8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
    doprint1
End Sub
Private Sub doprint1()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(2)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT [QFILE8_70].DATE, [QFILE8_70].DOC_NO, FILE8_71.DESCA AS DESCMAN , [QFILE8_70].DESCA, [QFILE8_70].VALUE_M, [QFILE8_70].VALUE_P FROM [QFILE8_70] INNER JOIN FILE8_71 ON [QFILE8_70].CODE = FILE8_71.CODE "
If lIsBranchStore Then cString = cString & " WHERE QFILE8_70.BRANCH = " & MyParn(cBranch)
If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & " BOX = " & MyParn(xbox.BoundText)
    aHeader(1) = "[" & xbox.Text & "]"
End If
If IsDate(xDate1.Text) Then
    cString = cString & turnFound(cString) & " [QFILE8_70].date >= " & DateSq(xDate1.Text)
    aHeader(0) = "[" & BetweenString(Format(xDate1.Text, "d-m-yyyy"), Format(xdate2.Text, "d-m-yyyy")) & "]"
End If

If IsDate(xdate2.Text) Then
    cString = cString & turnFound(cString) & " [QFILE8_70].date <= " & DateSq(xdate2.Text)
     aHeader(0) = "[" & BetweenString(Format(xDate1.Text, "d-m-yyyy"), Format(xdate2.Text, "d-m-yyyy")) & "]"
End If
    
If xMan.BoundText <> "" Then
    cString = cString & turnFound(cString) & " [QFILE8_70].CODE = " & MyParn(xMan.BoundText)
    aHeader(2) = "[" & xMan.Text & "]"
End If
   
         
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !doc_no
    temptable!Date1 = !Date
    temptable!str2 = !DESCMAN
    temptable!str3 = !DESCA
    
    temptable!VAL1 = !Value_M
    temptable!val2 = !Value_P

    temptable!val3 = !Value_M - !Value_P

    temptable!str22 = TurnValue(retHeader(aHeader, 0, 2))
    temptable!STR21 = " ›’Ì·Ï  Õ—ﬂ… ‘—ﬂ«¡ "
    temptable.Update
    sourcetable.MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\charge8.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
Data1.ConnectionString = strCon
If cBranch = "00" Then
    Data1.RecordSource = "FILE0_50"
Else
    Data1.RecordSource = "SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch)
End If
Set xbox.RowSource = Data1
xbox.BoundColumn = "CODE"
xbox.ListField = "DESCA"


Data2.ConnectionString = strCon
Data2.RecordSource = "FILE8_71"
Set xMan.RowSource = Data2
xMan.BoundColumn = "CODE"
xMan.ListField = "DESCA"

End Sub
Private Function RetCharge(pCharge)
ChargeTable.Find "Code = " & MyParn(pCharge), , adSearchForward, adBookmarkFirst
If Not ChargeTable.EOF Then RetCharge = ChargeTable!DESCA
End Function
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID() As Boolean
If (Not IsDate(xDate1.Text)) And Trim(xDate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xdate2.Text)) And Trim(xdate2.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
