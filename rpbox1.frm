VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpBox1 
   Caption         =   " ÕÊÌ·«  „‰ Œ“‰… ≈·Ï Œ“‰…"
   ClientHeight    =   1770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
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
   ScaleHeight     =   1770
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1680
      Left            =   1395
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   5055
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xBox1 
         Height          =   315
         Left            =   990
         TabIndex        =   2
         Top             =   900
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xbox2 
         Height          =   315
         Left            =   990
         TabIndex        =   3
         Top             =   1260
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰ Œ“‰… :"
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   945
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "≈·Ì Œ“‰… :"
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1305
         Width           =   825
      End
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   420
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   855
      Width           =   1320
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1305
      Width           =   1320
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   75
      Top             =   300
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
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
Attribute VB_Name = "rpBox1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
Dim aHeader(1)
If Not myValid Then Exit Sub

Dim sourcetable As New ADODB.Recordset
Dim temptable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT FILE0_50.DESCA AS Desc1 , FILE0_50_1.DESCA As Desc2 ,  FILE0_51.*  FROM (FILE0_51 LEFT JOIN FILE0_50 ON FILE0_51.NO1 = FILE0_50.CODE) LEFT JOIN FILE0_50 AS  FILE0_50_1 ON FILE0_51.NO2 = FILE0_50_1.CODE "
If IsDate(xDate1.text) Then
    cString = cString & turnFound(cString) & "Date >= " & DateSq(xDate1.text)
    aHeader(0) = "[" & BetweenString(Format(xDate1.text, "d-m-yyyy"), Format(xdate2.text, "d-m-yyyy")) & "]"
End If
If IsDate(xdate2.text) Then
    cString = cString & turnFound(cString) & "Date <= " & DateSq(xdate2.text)
    aHeader(0) = "[" & BetweenString(Format(xDate1.text, "d-m-yyyy"), Format(xdate2.text, "d-m-yyyy")) & "]"
End If


If xBox1.BoundText <> "" Then
    cString = cString & turnFound(cString) & " NO1 = " & MyParn(xBox1.BoundText)
    aHeader(1) = "[" & BetweenString(xBox1.text, xbox2.text) & "]"
End If

If xbox2.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  NO2 = " & MyParn(xbox2.BoundText)
     aHeader(1) = "[" & BetweenString(xBox1.text, xbox2.text) & "]"
End If

cString = cString & " ORDER BY DATE , NO1 , NO2 "
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
Else
    With sourcetable
    Do Until sourcetable.EOF
        temptable.AddNew
        temptable!str1 = !code
        temptable!str2 = !DESC1
        temptable!str3 = !DESC2
        temptable!STR4 = !desca
        temptable!VAL1 = !Value
        temptable!Date1 = !Date
        temptable!STR21 = retHeader(aHeader, 0, 2)
        
        temptable.Update
        sourcetable.MoveNext
    Loop
    End With
    contemp.BeginTrans
    contemp.CommitTrans
    Main.REPORT1.ReportFileName = App.Path & "\Reports\box1.rpt"
   ' MAIN.REPORT1.ReportFileName = App.Path & "\Reports\RepChrg8.rpt"
    Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    Main.REPORT1.Action = 1
End If
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
    End If
End Sub
Function myValid() As Boolean
If (Not IsDate(xDate1.text)) And Trim(xDate1.text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xdate2.text)) And Trim(xdate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
If xBox1.BoundText = "" And xbox2.BoundText = "" Then
    MsgBox "·«»œ „‰  ÕœÌœ Œ“«‰… Ê«Õœ… ⁄·Ì «·«ﬁ·"
    Exit Function
End If
myValid = True
End Function
Private Sub Form_Load()
openCon con
data1.ConnectionString = strCon
If cBranch <> "00" Then
    data1.RecordSource = "SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch)
Else
    data1.RecordSource = "FILE0_50"
End If
Set xBox1.RowSource = data1
xBox1.BoundColumn = "CODE"
xBox1.ListField = "DESCA"

Set xbox2.RowSource = data1
xbox2.BoundColumn = "CODE"
xbox2.ListField = "DESCA"
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

