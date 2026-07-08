VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTSupp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄… «·„’«‰⁄"
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15120
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
   ScaleHeight     =   11040
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   345
      Width           =   1155
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "ÿ»«⁄…"
      Height          =   420
      Left            =   1305
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   345
      Width           =   1155
   End
   Begin VB.CommandButton cmdGo 
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
      Left            =   2520
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   345
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   3810
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
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   6000
         TabIndex        =   8
         Top             =   952
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
         Left            =   1305
         TabIndex        =   9
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGRSUPP 
         Height          =   315
         Left            =   6000
         TabIndex        =   13
         Top             =   596
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   1305
         TabIndex        =   15
         Top             =   555
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   615
         Width           =   1230
      End
      Begin VB.Label Label5 
         Caption         =   "„Ã„Ê⁄… „Ê—œÌ‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„’‰⁄"
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
         Index           =   2
         Left            =   1755
         TabIndex        =   12
         Top             =   300
         Width           =   435
      End
      Begin VB.Label Label3 
         Caption         =   "«·„’‰⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   974
         Width           =   1545
      End
      Begin VB.Label Label2 
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
         Height          =   240
         Index           =   3
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»⁄œ  «—ÌŒ :"
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   765
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10710
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   1485
      Top             =   45
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTSupp.frx":0000
      Height          =   8490
      Left            =   45
      TabIndex        =   7
      Top             =   1530
      Width           =   15000
      _cx             =   26458
      _cy             =   14975
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12648447
      ForeColorSel    =   4194368
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
      RowHeightMin    =   0
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   1188
      Top             =   45
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
      Left            =   891
      Top             =   45
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   297
      Top             =   45
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   0
      Top             =   45
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
   Begin MSAdodcLib.Adodc data6 
      Height          =   330
      Left            =   594
      Top             =   45
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
   Begin VB.Label xmodelno 
      Alignment       =   1  'Right Justify
      Caption         =   "Label7"
      Height          =   240
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   945
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "VsTSupp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3
Public cMyModel As String
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï  ⁄«„·«  «·„Ã„Ê⁄«  "
    cHead2 = " Õ Ï   «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY")
    
    Load PrintGrd
    PrintGrd.doprint Me.GRID1, 0.75, -2, cHead1, cHead2, , False, True, 6
    PrintGrd.Show 1
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    StoreTable.Open "FILE0_40", con, adOpenStatic, adLockReadOnly, adCmdTable
    xDate1.Text = Format(Date, "dd-mm-yyyy")
    
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
    Set xSection.RowSource = data1
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File4_50 order by Desca"
    Set xGRSUPP.RowSource = DATA2
    xGRSUPP.ListField = "Desca"
    xGRSUPP.BoundColumn = "Code"
    
    
    data4.ConnectionString = strCon
    data4.RecordSource = "Select mosm ,descA From mosm ORDER BY date "
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    data5.ConnectionString = strCon
    data5.RecordSource = "Select code ,desca From fact ORDER BY code "
    Set xFact.RowSource = data5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    Set GRID1.DataSource = data6
    data6.ConnectionString = strCon
    
    GRID1.Rows = 1
    GRID1.Cols = 9
    FixGrid
End Sub
Private Sub myload()
    Dim cwhere As String
    Dim cStr2 As String
    
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xGRSUPP.BoundText <> "" Then cStr2 = cStr2 & " AND FILE4_10.[GROUP]  = " & MyParn(xGRSUPP.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xSection.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & xSection.BoundText
    
    FixSubView cStr2
    cField3 = " SUM ( case when ( type = '2' OR type = '7') THEN ([IN] - [OUT])  else 0 end ) AS t_purch    , "
    cField4 = " (SELECT COUNT(DISTINCT FILE1_10_2.MODELNO ) FROM FILE1_11 AS FILE1_11_2 INNER JOIN FILE1_10 AS FILE1_10_2 ON FILE1_11_2.ITEM = FILE1_10_2.ITEM WHERE FILE1_11_2.TYPE = '6' AND FILE1_10_2.[FACT] = FACT.CODE "
    If xMosm.BoundText <> "" Then cField4 = cField4 & " AND FILE1_10_2.MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cField4 = cField4 & " AND FILE1_10_2.CODE = " & MyParn(xSupp.BoundText)
    If xFact.BoundText <> "" Then cField4 = cField4 & " AND FILE1_10_2.FACT = " & MyParn(xFact.BoundText)
    If xSection.BoundText <> "" Then cField4 = cField4 & " AND FILE1_10_2.[Section] = " & xSection.BoundText
    cField4 = cField4 & " )   AS CountSALES   , "
    cField5 = " SUM ( case when (type = '6' )                THEN ([out] )  else 0 end ) AS t_sal     , "
        
    cField7 = " SUM ( case when (type = '2' OR  type = '7') THEN (([IN] - [OUT]) * FILE1_10.COST )  else 0 end ) AS VAL_PURCH     , "
    cField8 = " SUM ( case when (type = '6' )                THEN (([OUT]) * FILE1_10.COST )  else 0 end ) AS VAL_SAL , "
    
    cField11 = " (SELECT COUNT(DISTINCT MODELNO_SHOP0.MODELNO ) FROM MODELNO_SHOP0 INNER JOIN FILE1_10 AS FILE1_10_2 ON MODELNO_SHOP0.MODELNO = FILE1_10_2.MODELNO WHERE FILE1_10_2.[FACT] = FACT.CODE "
    If xMosm.BoundText <> "" Then cField11 = cField11 & " AND FILE1_10_2.MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cField11 = cField11 & " AND FILE1_10_2.CODE = " & MyParn(xSupp.BoundText)
    If xFact.BoundText <> "" Then cField11 = cField11 & " AND FILE1_10_2.FACT = " & MyParn(xFact.BoundText)
    If xSection.BoundText <> "" Then cField11 = cField11 & " AND FILE1_10_2.[Section] = " & xSection.BoundText
    cField11 = cField11 & " )   AS CountSALES0    "
    
    '                   0                   1                            2               3                   4                              5
    
    cString = "    SELECT FACT.CODE , FACT.DESCA ,  COUNT(DISTINCT FILE1_10.MODELNO)  , "
    cString = cString & cField3 & cField4 & cField5 & _
                " sum([in] - out) as bal  , " & _
                cField7 & cField8 & _
                " sum(([IN] - [out])* FILE1_10.COST) as VAL_BAL     , ' ' AS N10  , " & cField11
    cString = cString & " FROM FILE1_10 INNER JOIN FACT ON FILE1_10.[FACT] = FACT.CODE INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM INNER JOIN MODELNO_SHOP ON FILE1_10.modelno = MODELNO_SHOP.modelno WHERE FILE1_10.MODELNO IS NOT NULL "
    cString = cString & cStr2 & " group BY FACT.CODE , FACT.DESCA ORDER BY FACT.CODE "

    data6.RecordSource = cString
    data6.Refresh
    
    FixGrid
End Sub
Sub FixGrid()
    With GRID1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .Cols = 12
    .RowHeight(0) = 1000
    .WordWrap = True
    .ColHidden(0) = True
    .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„’‰⁄ "
    
    .TextMatrix(0, 2) = "⁄œœ „ÊœÌ·«  „‘ —Ì« "
    .TextMatrix(0, 3) = "⁄œœ „‘ —Ì«  "
    
    .TextMatrix(0, 4) = "⁄œœ „»Ì⁄«  „ÊœÌ·«  "
    .TextMatrix(0, 5) = "⁄œœ „»Ì⁄«  "
    
    .TextMatrix(0, 6) = "—’Ìœ ⁄œœ "
    
    .TextMatrix(0, 7) = "ﬁÌ„… „‘ —Ì« "
    .TextMatrix(0, 8) = "ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 9) = "ﬁÌ„… «·—’Ìœ "
    
    .TextMatrix(0, 10) = "‰”»… „»Ì⁄«  "
    
    .TextMatrix(0, 11) = "⁄œœ „ÊœÌ·«  „— Ã⁄"
    
    
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 3000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 800
    .ColWidth(11) = 800
    
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 7)) <> 0 Then .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 8)) / Val(.TextMatrix(nRow, 7)) * 100, 2)
    Next nRow
    For nCol = 2 To .Cols - 3
        .ColDataType(nCol) = flexDTDouble
        .ColFormat(nCol) = "#0"
    Next nCol
    .Editable = flexEDNone
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    If .Rows > 1 Then
        .Subtotal flexSTClear
        For nCol = 2 To .Cols - 1
            .Subtotal flexSTSum, -1, nCol, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
        Next nCol
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
        nRow = 1
        If Val(.TextMatrix(nRow, 7)) <> 0 Then .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 8)) / Val(.TextMatrix(nRow, 7)) * 100, 2)
    
    End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub

Private Sub grid1_dblClick()
With GRID1
    Load VsTItem
    VsTItem.xFact.BoundText = .TextMatrix(.Row, 0)
    VsTItem.xMosm.BoundText = xMosm.BoundText
    VsTItem.xDate1.Text = xDate1.Text
'    VsTItem.xSection.BoundText = xSection.BoundText
    VsTItem.xSupp.BoundText = xSupp.BoundText
    VsTItem.Show
End With
End Sub
Private Sub grid1_EnterCell()
    xModelNo.Caption = GRID1.TextMatrix(GRID1.Row, 0)
End Sub
Sub FixSubView(cPstr)
    On Error Resume Next
    
    con.Execute " DROP VIEW Gr_Purch "
    con.Execute " DROP VIEW Gr_Sal "
    con.Execute " DROP VIEW Gr_Bal "
    Err.Clear
    

    cStr1 = " CREATE VIEW [Gr_BAL] AS SELECT     dbo.FILE1_10H.MODELNO, dbo.FILE1_10H.[GROUP] " & _
            " FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM INNER JOIN FILE1_10H ON FILE1_10.MODEL = FILE1_10H.MODEL Where FILE1_10.MODEL IS NOT NULL " & cPstr & _
            " GROUP BY FILE1_10H.[GROUP], FILE1_10H.MODELNO HAVING SUM(FILE1_11.[In] - dbo.FILE1_11.OUT) <> 0   "
    con.Execute cStr1

    cStr1 = " CREATE VIEW [Gr_PURCH] AS SELECT     dbo.FILE1_10H.MODELNO, dbo.FILE1_10H.[GROUP] " & _
            " FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM INNER JOIN FILE1_10H ON FILE1_10.MODEL = FILE1_10H.MODEL Where ( FILE1_11.TYPE = '2' OR FILE1_11.TYPE = '7' ) AND (FILE1_11.[In] - dbo.FILE1_11.OUT) <> 0 " & cPstr & _
            " GROUP BY FILE1_10H.[GROUP], FILE1_10H.MODELNO HAVING SUM(FILE1_11.[In] - dbo.FILE1_11.OUT) <> 0   "
    con.Execute cStr1

    cStr1 = " CREATE VIEW [Gr_SAL] AS SELECT     dbo.FILE1_10H.MODELNO, dbo.FILE1_10H.[GROUP] " & _
            " FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM INNER JOIN FILE1_10H ON FILE1_10.MODEL = FILE1_10H.MODEL Where ( FILE1_11.TYPE = '6' ) AND  dbo.FILE1_11.OUT  <> 0 " & cPstr & _
            " GROUP BY FILE1_10H.[GROUP], FILE1_10H.MODELNO HAVING SUM(dbo.FILE1_11.OUT) <> 0   "
    con.Execute cStr1

End Sub
Sub myProc()
If ActiveControl.Name = xSupp.Name Then
    xSupp.BoundText = oSearch.GRID1.TextMatrix(oSearch.GRID1.Row, 0)
    oSearch.Hide
End If
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.GRID1.TextMatrix(Search3.GRID1.Row, 0)
    Unload Search3
End If
End Sub


Private Sub xFACT_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FACT"
Generalarray(2) = "Order by DESCA"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1

End If
End Sub

