VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsTProftSuple 
   Caption         =   "«—»«Õ «·„Ê—œÌ‰"
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   13260
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   13260
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   3885
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2520
         Picture         =   "VsTProftSuple.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   90
         Picture         =   "VsTProftSuple.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1305
         Picture         =   "VsTProftSuple.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   510
         Left            =   90
         TabIndex        =   15
         Top             =   675
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   900
         _Version        =   196610
         BackColor       =   13822956
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "VsTProftSuple.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   1260
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
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
         Left            =   7620
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Visible         =   0   'False
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
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGrCust 
         Height          =   315
         Left            =   6000
         TabIndex        =   7
         Top             =   615
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   1125
         TabIndex        =   9
         Top             =   630
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„ :"
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
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   675
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… „Ê—œÌ‰:"
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
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   1455
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
         Left            =   9555
         TabIndex        =   5
         Top             =   270
         Visible         =   0   'False
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
         Left            =   4680
         TabIndex        =   4
         Top             =   285
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   13260
      _ExtentX        =   23389
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
      Bindings        =   "VsTProftSuple.frx":9583
      Height          =   8055
      Left            =   135
      TabIndex        =   6
      Top             =   1485
      Width           =   15000
      _cx             =   26458
      _cy             =   14208
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      BackColorSel    =   16777088
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
      Left            =   2475
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   150
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
Attribute VB_Name = "VsTProftSuple"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = Me.Caption
    cHead2 = " ·„Ê”„ " & xMosm.text
    
    Load PrintGrd
    PrintGrd.DOPRINT Me.grid1, 1, 2, cHead1, cHead2, , False, True, 9
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
    Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    If xMosm.BoundText = "" Or Not IsDate(xDate1.text) Or Not IsDate(xDate2.text) Then
        MsgBox "ÌÃ»  ÕœÌœ „Ê”„ - «· «—ÌŒ "
        Exit Sub
    End If
    myload
End Sub
Private Sub Form_Load()
    openCon con
    xDate1.text = "1-1-" & Year(Date)
    xDate2.text = Format(Date, "dd-mm-yyyy")
    
    data1.ConnectionString = strCon
    data1.RecordSource = "SELECT * FROM FILE4_50 "
    Set xGrCust.RowSource = data1
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    DATA3.ConnectionString = strCon
    DATA3.RecordSource = "SELECT * FROM MOSM"
    Set xMosm.RowSource = DATA3
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "mosm"
    xMosm.BoundText = cPMosm
    
    Set grid1.DataSource = DATA2
    DATA2.ConnectionString = strCon
    Me.Caption = "«—»«Õ «·„Ê—œÌ‰"
    FIXGRID
    grid1.Rows = 1
End Sub
Private Sub myload()
Dim cString As String
    
    cW1 = " file4_11disc.mosm = " & MyParn(xMosm.BoundText) & " AND file4_11disc.CODE = FILE4_10.CODE "
    cField2 = "(SELECT SUM(T_ITEM)                  FROM    file4_11disc WHERE " & cW1 & " )  "
    cField3 = "(SELECT SUM(T_DISC + T_DISC2 )       FROM    file4_11disc WHERE " & cW1 & " ) "
    cField4 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '4' OR TYPE = '5' )", "[sal]- [PAY]") & " AS net_sal"

    cField5 = " (SELECT SUM  ( QALL_SALES.QUANT * QALL_SALES.COST  )                FROM  QALL_SALES inner join file1_10 on  QALL_SALES.item = file1_10.item WHERE file1_10.MOSM = " & MyParn(xMosm.BoundText) & " AND file1_10.CODE = FILE4_10.CODE )   "
    cField6 = " (SELECT SUM  ( QALL_SALES.TOTAL  * ((100- QALL_SALES.RATE )/100) )  FROM  QALL_SALES inner join file1_10 on  QALL_SALES.item = file1_10.item WHERE file1_10.MOSM = " & MyParn(xMosm.BoundText) & " AND file1_10.CODE = FILE4_10.CODE )   "

    cField8 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '9')", "[PAY]") & " AS T_CASH0 "
    cField10 = "(SELECT  sum(([in]-[out]) * file1_10.costITEM )   from file1_11 inner join file1_10 on file1_10.item = file1_11.item  where file1_10.code = file4_10.code and file1_10.mosm = " & MyParn(xMosm.BoundText) & " and date <= " & DateSq(xDate2.text) & " ) "

    cField11 = myiif(cWhere, "[sal]- [pay] ") & " AS C_BAL"

With grid1
'                           0                           1
    cString = " SELECT FILE4_10.CODE AS CODE , FILE4_10.DESCA AS DESCA  , " & _
                cField2 & " , " & cField3 & " , " & cField4 & " , " & cField5 & " , " & _
                cField6 & " , ' ' AS N7 , " & cField8 & " , ' ' as N9 , " & cField10 & " , " & cField11 & _
                " FROM (FILE4_11 inner JOIN file4_10 ON FILE4_11.CODE = FILE4_10.CODE) LEFT JOIN file4_50 ON FILE4_10.[group] = file4_50.CODE WHERE FILE4_10.CODE IS NOT NULL "
    If xGrCust.BoundText <> "" Then cString = cString & turn(cString) & " file4_10.[group]  = " & MyParn(xGrCust.BoundText)
    If xMosm.BoundText <> "" Then cString = cString & turn(cString) & " FILE4_11.[MOSM] = " & MyParn(xMosm.BoundText)
    cString = cString & " GROUP BY FILE4_10.DESCA , FILE4_10.CODE "
    DATA2.RecordSource = cString
    DATA2.Refresh
End With
FIXGRID
End Sub
Sub FIXGRID()
    With grid1
    .Cols = 14
    .RowHeight(0) = 1000
    .WordWrap = True
    .FrozenCols = 2
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    
    .TextMatrix(0, 2) = "’ „‘ —Ì«  „ÊœÌ·« "
    .TextMatrix(0, 3) = "Œ’„ ›Ê« Ì—"
    .TextMatrix(0, 4) = "’«›Ï «·„‘ —Ì« "
    .TextMatrix(0, 5) = " ﬂ·›… «·„»Ì⁄« "
    .TextMatrix(0, 6) = "ﬁÌ„… «·„»Ì⁄« "
    .TextMatrix(0, 7) = "—»Õ «·„»Ì⁄« "
    .TextMatrix(0, 8) = " ”ÊÌ«  „Ê—œÌ‰"
    .TextMatrix(0, 9) = "„Êﬁ› «·„Ê—œ"
    .TextMatrix(0, 10) = " ﬁÌ„ —’Ìœ „ÊœÌ·«  «·„Ê—œ"
    .TextMatrix(0, 11) = "—’Ìœ Õ”«» «·„Ê—œ"
    
    .TextMatrix(0, 12) = "‰”»… —»Õ „»Ì⁄«  «·„Ê—œ · ﬂ·›… «·„»Ì⁄« "
    .TextMatrix(0, 13) = "‰”»… „»Ì⁄«  «·„Ê—œ"
    
    
    .ColWidth(0) = 700
    .ColWidth(1) = 1800
    .ColWidth(2) = 1100
    .ColWidth(3) = 1100
    .ColWidth(4) = 1100
    .ColWidth(5) = 1100
    .ColWidth(6) = 1100
    .ColWidth(7) = 1100
    .ColWidth(8) = 1100
    .ColWidth(9) = 1100
    .ColWidth(10) = 1100
    .ColWidth(11) = 1100
    .ColWidth(12) = 1100
    .ColWidth(13) = 1100
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 7) = Val(.TextMatrix(nRow, 6)) - Val(.TextMatrix(nRow, 5))
        
        .TextMatrix(nRow, 9) = Val(.TextMatrix(nRow, 6)) + Val(.TextMatrix(nRow, 8)) - Val(.TextMatrix(nRow, 4))
        
        nRate = 0
        If Val(.TextMatrix(nRow, 5)) > 0 Then nRate = Round(Val(.TextMatrix(nRow, 7)) / Val(.TextMatrix(nRow, 5)) * 100, 2)
        .TextMatrix(nRow, 12) = nRate
        
        nRate = 0
        If Val(.TextMatrix(nRow, 4)) > 0 Then nRate = Round(Val(.TextMatrix(nRow, 5)) / Val(.TextMatrix(nRow, 4)) * 100, 2)
        .TextMatrix(nRow, 13) = nRate
        
        If Val(.TextMatrix(nRow, 9)) >= 0 Then
            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbGreen
        End If
    Next nRow
    For nCol = 2 To .Cols - 1
        .ColDataType(nCol) = flexDTDouble
        .ColFormat(nCol) = "#0.00"
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Round(Val(.TextMatrix(nRow, nCol)), 2)
        Next nRow
    Next nCol
    
    .SubtotalPosition = flexSTAbove
    For I = 2 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    .ExplorerBar = flexExSort
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
        nRow = 1
        nRate = 0
        If Val(.TextMatrix(nRow, 5)) > 0 Then nRate = Round(Val(.TextMatrix(nRow, 7)) / Val(.TextMatrix(nRow, 5)) * 100, 2)
        .TextMatrix(nRow, 12) = nRate
        
        nRate = 0
        If Val(.TextMatrix(nRow, 4)) > 0 Then nRate = Round(Val(.TextMatrix(nRow, 5)) / Val(.TextMatrix(nRow, 4)) * 100, 2)
        .TextMatrix(nRow, 13) = nRate
    End If
    End With
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
xCustName.Caption = ""
If xCode.text = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 6)
xCustName.Caption = GetDesca("select desca from FILE4_10 where code = " & MyParn(xCode.text), con) & ""

End Sub
Sub myProc()
ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Unload Search3
End Sub
