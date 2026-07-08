VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpDemModelCust 
   Caption         =   " ﬁ—Ì— ÿ·»Ì«  „ÊœÌ·«   ·⁄„·«¡ "
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5295
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
   ScaleHeight     =   3525
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2670
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   5010
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1290
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1290
      End
      Begin MSDataListLib.DataCombo XMOSM 
         Height          =   315
         Left            =   75
         TabIndex        =   2
         Top             =   945
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XGROUP 
         Height          =   315
         Left            =   75
         TabIndex        =   3
         Top             =   1305
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XGRCUST 
         Height          =   315
         Left            =   90
         TabIndex        =   14
         Top             =   1665
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xbranch 
         Height          =   315
         Left            =   90
         TabIndex        =   16
         Top             =   2070
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄ :"
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2115
         Width           =   510
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… ⁄„·«¡"
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1755
         Width           =   1080
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
         Height          =   240
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "≈·Ï  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   630
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… „ÊœÌ·« "
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1395
         Width           =   1290
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„Ê”„ "
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1035
         Width           =   540
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   1080
      Top             =   2925
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
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2835
      Width           =   3615
      Begin VB.CommandButton CmdClear 
         Caption         =   " ÃœÌœ"
         Height          =   390
         Left            =   1200
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   135
         Width           =   1140
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   390
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   135
         Width           =   1140
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   3870
      Top             =   3105
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
      Left            =   5100
      Top             =   2550
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   0
      Top             =   0
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
Attribute VB_Name = "rpDemModelCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim nOption As Integer
Function MYVALID()
'If (Not IsDate(xdate1.Text)) And Trim(xdate1.Text) <> "" Then Exit Function
'If (Not IsDate(xDate2.Text)) And Trim(xDate2.Text) <> "" Then Exit Function
'If (Trim(xStore1.BoundText) = Trim(xStore2.BoundText) And xStore2.BoundText <> "") Then
'    MsgBox "·« Ì„ﬂ‰ «· ÕÊÌ· „‰ „Œ“‰ «·Ï ‰›” „Œ“‰"
'    Exit Function
'End If
MYVALID = True
End Function
Private Sub cmdClear_Click()
xDate1.Text = ""
xDate2.Text = ""
xStore1.BoundText = ""
xStore2.BoundText = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xStore1.BoundText = ""
xDate1.Text = ""
xDate2.Text = ""
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If TypeOf ActiveControl Is DBCombo Then ActiveControl.BoundText = ""
End If
End Sub
Private Sub Form_Load()
openCon con

    data1.ConnectionString = strCon
    data1.RecordSource = "Select MOSM ,DescA From MOSM "
    Set xMosm.RowSource = data1
    xMosm.ListField = "MOSM"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select CODE ,DescA From FILE1_50 WHERE ISRAW = 0 ORDER BY CODE"
    Set xGroup.RowSource = DATA2
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select CODE ,DescA From FILE3_50 ORDER BY DESCA"
    Set xGrCust.RowSource = data3
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"

    data12.ConnectionString = strCon
    data12.RecordSource = "SELECT * FROM branch "
    Set xbranch.RowSource = data12
    xbranch.ListField = "Desca"
    xbranch.BoundColumn = "Code"

End Sub
Private Sub CmdApply_Click()
Dim aHeader(1)
If Not MYVALID Then Exit Sub
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cStr1 = " SELECT  FILE3_10.CODE, FILE3_10.DESCA AS custname , FILE6_50H.DOC_NO, FILE6_50H.DATE, FILE1_10.MODEL, FILE1_10.COLOR, SUM(FILE6_50.QUANT) AS QUANT,   FILE1_10.desca " & _
        " FROM  FILE3_10 INNER JOIN  FILE6_50H ON FILE3_10.CODE = FILE6_50H.CODE INNER JOIN FILE1_10 INNER JOIN FILE6_50 ON FILE1_10.ITEM = FILE6_50.ITEM ON FILE6_50H.DOC_NO = FILE6_50.DOC_NO WHERE FILE1_10.MODEL IS NOT NULL "
If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE6_50H.MOSM = " & MyParn(xMosm.BoundText)
If xGrCust.BoundText <> "" Then cStr1 = cStr1 & " AND FILE3_10.[GROUP] = " & MyParn(xGrCust.BoundText)
If xGroup.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
If xbranch.BoundText <> "" Then cStr1 = cStr1 & " AND FILE3_10.[branch] = " & MyParn(xbranch.BoundText)
If IsDate(xDate1.Text) Then cStr1 = cStr1 & " AND FILE6_50H.[DATE] >= " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cStr1 = cStr1 & " AND FILE6_50H.[DATE] <= " & DateSq(xDate2.Text)

cStr1 = cStr1 & " GROUP BY FILE3_10.CODE, FILE3_10.DESCA, FILE6_50H.DOC_NO, FILE6_50H.DATE, FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.desca "

If IsDate(xDate1.Text) Then
    aHeader(1) = "[" & BetweenString(xDate1.Text, xDate2.Text) & "]"
End If
If IsDate(xDate2.Text) Then
    aHeader(1) = "[" & BetweenString(xDate1.Text, xDate2.Text) & "]"
End If
'aHeader(2) = XGRCUST.Text & "  " & XMOSM.Text & " " & XGROUP.Text

Set sourcetable = New ADODB.Recordset
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·⁄—÷Â«"
    GoTo lastsub
End If
Do Until .EOF
    temptable.AddNew
    temptable!str1 = !doc_no
    temptable!str2 = " ›’Ì·Ï ÿ·»Ì«  " & xMosm.Text
    temptable!str3 = !MODEL
    temptable!str4 = !DESCA
    temptable!str5 = ArbString(!CODE & "  : " & !custname)
    temptable!VAL1 = !Quant
    temptable!STR6 = !Color
    temptable!Date1 = ![Date]
    temptable!STR21 = xbranch.Text & "  " & xGrCust.Text & "  " & xGroup.Text
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\R_ORDER.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
lastsub:
    temptable.Close
    sourcetable.Close
    Set temptable = Nothing
    Set sourcetable = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
