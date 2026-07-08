VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTBalSUPL_ALL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«—’œ… Ê ≈Ã„«·Ï Õ—ﬂ… «·„Ê—œÌ‰"
   ClientHeight    =   10290
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   720
      Width           =   4920
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "VsTBalSUPL_ALL.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "VsTBalSUPL_ALL.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "VsTBalSUPL_ALL.frx":4CDD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2430
         Picture         =   "VsTBalSUPL_ALL.frx":7149
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   11070
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   8985
      Begin VB.TextBox Xcode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   622
         Width           =   1815
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5370
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGrCust 
         Height          =   315
         Left            =   465
         TabIndex        =   7
         Top             =   1035
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XMOSM 
         Height          =   315
         Left            =   4635
         TabIndex        =   19
         Top             =   990
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ê”„ : "
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
         Index           =   1
         Left            =   7335
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1035
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.Label xCustName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   615
         Width           =   4125
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·„Ê—œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   690
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… „Ê—œÌ‰ "
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
         Index           =   0
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1050
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7425
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   675
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   4
         Top             =   285
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   9960
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTBalSUPL_ALL.frx":9573
      Height          =   8010
      Left            =   240
      TabIndex        =   6
      Top             =   1575
      Width           =   19815
      _cx             =   34951
      _cy             =   14129
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12648384
      ForeColorSel    =   64
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   180
      Top             =   360
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
      Left            =   3750
      Top             =   75
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
   Begin ComctlLib.ProgressBar prog1 
      Height          =   240
      Left            =   270
      TabIndex        =   17
      Top             =   495
      Width           =   4875
      _ExtentX        =   8599
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
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
End
Attribute VB_Name = "VsTBalSUPL_ALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub CMD_PRINT_Click()
End Sub
Private Sub cmdExel_Click()
      ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub CmdExit_Click()
    Unload Me
    Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    If IsDate(xDate1.text) And IsDate(xDate2.text) Then
        myload
    Else
        MsgBox "  ÕœÌœ «·› —…  "
    End If
End Sub
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï „Êﬁ› Ê «—’œ… «·„Ê—œÌ‰ "
    If xGrCust.text <> "" Then cHead1 = cHead1 & xGrCust.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xDate2.text, "DD-MM-YYYY")
    Load PrintGrd
    PrintGrd.DOPRINT Me.grid1, 0.9, -1, cHead1, cHead2, , False, True, 9
    PrintGrd.Show 1

End Sub

Private Sub Form_Load()
    openCon con
    If nWidthScreen > 1 Then grid1.Width = 20000
    xDate1.text = "1-1-" & Year(Date)
    xDate2.text = Format(Date, "dd-mm-yyyy")
    
    data1.ConnectionString = strCon
    data1.RecordSource = "SELECT * FROM FILE4_50 "
    Set xGrCust.RowSource = data1
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    DATA3.ConnectionString = strCon
    DATA3.RecordSource = "SELECT * FROM MOSM ORDER BY DATE DESC "
    Set xMosm.RowSource = DATA3
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    
    Set grid1.DataSource = DATA2
    DATA2.ConnectionString = strCon
    
    FIXGRID
    grid1.Rows = 1
End Sub
Private Sub myload()
Dim cString As String

If IsDate(xDate1.text) Then cWhere = " date < " & DateSq(xDate1.text)
cField2 = myiif(cWhere, "[SAL] - [PAY]") & " AS F_BAL , "

cWhere = ""

If IsDate(xDate1.text) Then cWhere = " date >= " & DateSq(xDate1.text)
If IsDate(xDate2.text) Then cWhere = cWhere & turn(cWhere, " and ") & " DATE <= " & DateSq(xDate2.text)

cField3 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '4' or TYPE = '41' or TYPE = '42' )", "[SAL]") & " AS T_SAL , "
cField4 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '5')", "[PAY]") & " AS T_RET , "
cField5 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '4' or TYPE = '41' or TYPE = '42' OR TYPE = '5')", "SAL-[PAY]") & " AS NET_SAL , "
cField6 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE =  '7' or  TYPE =  '8' )", "[PAY] - [SAL]") & " AS T_CASH ,"
cField7 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = 'A' OR TYPE = 'C' )", "[PAY]") & " AS T_CHQ ,"
cField8 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '9')", "[PAY]") & " AS T_CASH0 , "
cField9 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = 'D')", "[SAL]") & " AS T_CASH00 , "
cField11 = "(SELECT SUM(VALUE) FROM FILE5_21 WHERE CLOSED = '0' AND FILE5_21.CODE1 = FILE4_10.CODE ),"

cField12 = "(SELECT SUM(FILE1_10.COSTITEM * (FILE1_11.[IN]-FILE1_11.[OUT])) FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_11.ITEM = FILE1_10.ITEM WHERE FILE1_10.CODE = FILE4_10.CODE AND (FILE4_10.ISONEST = 1 OR FILE1_10.ISONEST = 1 ) AND FILE1_11.DATE <= " & DateSq(xDate2.text) & "),"
cField14 = ",(SELECT SUM(QFILE6_20_1.cost * QFILE6_20_1.quant ) FROM FILE1_10 INNER JOIN QFILE6_20_1 ON FILE1_10.ITEM = QFILE6_20_1.ITEM WHERE FILE1_10.CODE = FILE4_10.CODE AND (FILE4_10.ISONEST = 1 OR FILE1_10.ISONEST = 1 ) AND QFILE6_20_1.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20_1.DATE <= " & DateSq(xDate2.text) & ")"
cWhere = ""
If IsDate(xDate2.text) Then cWhere = "date <= " & DateSq(xDate2.text)
cField10 = myiif(cWhere, "[sal]- [pay] ") & " AS C_BAL , "

With grid1
'                           0                           1
    cString = "  select FILE4_10.CODE AS CODE , FILE4_10.DESCA AS DESCA , " & _
                cField2 & cField3 & cField4 & cField5 & cField6 & cField7 & cField8 & cField9 & cField10 & cField11 & cField12 & "  0 " & cField14 & _
                " FROM FILE4_11_2 INNER JOIN FILE4_10 ON FILE4_11_2.CODE = FILE4_10.CODE WHERE FILE4_10.CODE IS NOT NULL "
    If xGrCust.BoundText <> "" Then cString = cString & turn(cString) & " file4_10.[group]  = " & MyParn(xGrCust.BoundText)
'   If XMOSM.BoundText <> "" Then cString = cString & turn(cString) & " FILE4_11_2.[MOSM]  = " & MyParn(XMOSM.BoundText)
    If xCode.text <> "" Then cString = cString & turn(cString) & " FILE4_10.[CODE] = " & MyParn(xCode.text)
    cString = cString & " GROUP BY FILE4_10.DESCA , FILE4_10.CODE , FILE4_10.ISONEST"
    DATA2.RecordSource = cString
    DATA2.Refresh
End With
FIXGRID
End Sub
Sub FIXGRID()
    Dim nCash As Double, nPurch As Double, nRate As Double
    With grid1
    .Cols = 15
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .FrozenCols = 2
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    
    .TextMatrix(0, 2) = "«·—’Ìœ " & xDate1.text
    .TextMatrix(0, 3) = "Ã. „‘ —Ì« "
    .TextMatrix(0, 4) = "Ã. „— Ã⁄«  "
    .TextMatrix(0, 5) = "’«›Ï ﬁÌ„… „‘ —Ì« "
    
    .TextMatrix(0, 6) = "”œ«œ ‰ﬁœÏ"
    .TextMatrix(0, 7) = "”œ«œ √.œ"
    .TextMatrix(0, 8) = " ”ÊÌ« "
    .TextMatrix(0, 9) = "„’«—Ì›"
    .TextMatrix(0, 10) = "—’Ìœ " & xDate2.text
    .TextMatrix(0, 11) = "√.œ  Õ  «· Õ’Ì·"
    
    .TextMatrix(0, 12) = "—’Ìœ »÷«⁄… «„«‰« "
    .TextMatrix(0, 13) = "«·„” Õﬁ"
    .TextMatrix(0, 14) = " ﬂ·›… „»Ì⁄«  «„«‰« "
    
    .ColWidth(0) = 900
    .ColWidth(1) = 2000
    .ColWidth(2) = 1300
    .ColWidth(3) = 1300
    .ColWidth(4) = 1100
    .ColWidth(5) = 1300
    .ColWidth(6) = 1300
    .ColWidth(7) = 1200
    .ColWidth(8) = 1100
    .ColWidth(9) = 1100
    .ColWidth(10) = 1300
    .ColWidth(11) = 1100
    .ColWidth(12) = 1100
    .ColWidth(13) = 1200
    .ColWidth(14) = 1200
    
'    .ColHidden(12) = Not bOpt5
'    .ColHidden(13) = Not bOpt5
    
    If bOpt5 Then
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 5, "#0.00", vbYellow, vbRed, True, " "
    End If
    For nRow = 1 To .Rows - 1
'         If Val(.TextMatrix(.Rows - 1, 5)) <> 0 Then
'            nRate = Round(Val(.TextMatrix(nRow, 5)) / Val(.TextMatrix(.Rows - 1, 5)) * 100, 3)
'            .TextMatrix(nRow, 12) = nRate
'         End If
'        nCash = Val(.TextMatrix(nRow, 6)) + Val(.TextMatrix(nRow, 7)) + Val(.TextMatrix(nRow, 8))
'        nPurch = Val(.TextMatrix(nRow, 2)) + Val(.TextMatrix(nRow, 5))
'        If nPurch <> 0 Then
'            nRate = Round(nCash / nPurch * 100, 2)
'            .TextMatrix(nRow, 13) = nRate
'        End If
    Next nRow
    
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 13) = Val(.TextMatrix(nRow, 10)) - Val(.TextMatrix(nRow, 12))
    Next nRow
    For nCol = 2 To .Cols - 1
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Round(Val(.TextMatrix(nRow, nCol)), 2)
        Next nRow
        .ColDataType(nCol) = flexDTDouble
        .ColFormat(nCol) = "#0.00"
        If bOpt5 Then
            .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, vbRed, True, " "
        End If
    Next nCol
    If .Rows > 1 Then
'        nRow = .Rows - 1
'        nCash = Val(.TextMatrix(nRow, 6)) + Val(.TextMatrix(nRow, 7)) + Val(.TextMatrix(nRow, 8))
'        nPurch = Val(.TextMatrix(nRow, 2)) + Val(.TextMatrix(nRow, 5))
'        If nPurch <> 0 Then
'            nRate = Round(nCash / nPurch * 100, 2)
'            .TextMatrix(nRow, 13) = nRate
'        End If
    End If
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub

Private Sub grid1_DblClick()
    If grid1.Row > 0 Then
        Load Supler_Move
        Supler_Move.xCode.text = grid1.TextMatrix(grid1.Row, 0)
        Supler_Move.XDESCA.Caption = grid1.TextMatrix(grid1.Row, 1)
        Supler_Move.xDate1.text = xDate1.text
        Supler_Move.xDate2.text = xDate2.text
        Supler_Move.Fillgrd
        Supler_Move.Show
    End If
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
xCustName.Caption = ""
If xCode.text = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 3)
xCustName.Caption = GetDesca("select desca from FILE4_10 where code = " & MyParn(xCode.text), con) & ""


End Sub
Sub myProc()
ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Unload Search3
End Sub
Private Sub CardLookup()
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

Private Sub xMosm_LostFocus()
If xMosm.BoundText <> "" Then
    xDate1.text = ""
    xDate2.text = ""
End If
End Sub
