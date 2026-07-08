VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VsTSECmosm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÞíãÉ ÇáÑÕíÏ ãæÒÚ Úáì ÇáÇÞÓÇã æ ÇáÝÑæÚ"
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
   Begin VB.Frame Frame4 
      Height          =   1320
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   -45
      Width           =   3795
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2430
         Picture         =   "VsTSECmosm.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "ÚÑÖ"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "VsTSECmosm.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "VsTSECmosm.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   11
         Top             =   630
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   820
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
         Picture         =   "VsTSECmosm.frx":6D88
         Caption         =   "ÊÍæíá ÇßÓíá "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   240
         Left            =   45
         TabIndex        =   13
         Top             =   1080
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   423
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1305
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   -45
      Width           =   11235
      Begin VB.Frame Frame2 
         Height          =   510
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   450
         Width           =   4470
         Begin VB.OptionButton xall 
            Alignment       =   1  'Right Justify
            Caption         =   "Çáßá"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   225
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   180
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.OptionButton xall 
            Alignment       =   1  'Right Justify
            Caption         =   "ÊæßíáÇÊ ÝÞØ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1845
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   180
            Width           =   1140
         End
         Begin VB.OptionButton xall 
            Alignment       =   1  'Right Justify
            Caption         =   "ÝÑæÚ ÝÞØ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   180
            Width           =   960
         End
      End
      Begin MSDataListLib.DataCombo XMOSM 
         Height          =   315
         Left            =   5985
         TabIndex        =   5
         Top             =   585
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker xdate1 
         Height          =   375
         Left            =   7515
         TabIndex        =   12
         Top             =   180
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   441647105
         CurrentDate     =   43098.0821990741
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáãæÓã : "
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
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ÍÊì ÊÇÑíÎ :"
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
         TabIndex        =   2
         Top             =   270
         Width           =   885
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
            Key             =   ""
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
      Bindings        =   "VsTSECmosm.frx":9583
      Height          =   8895
      Left            =   90
      TabIndex        =   3
      Top             =   1305
      Width           =   15000
      _cx             =   26458
      _cy             =   15690
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
   Begin MSAdodcLib.Adodc DATA7 
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
   Begin VB.Label xmodelno 
      Alignment       =   1  'Right Justify
      Caption         =   "Label7"
      Height          =   240
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   450
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "VsTSECmosm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3
Public cMyModel As String
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "ÈíÇä ÅÌãÇáì ÑÕíÏ ÇáÇÞÓÇã ãæÒÚ Úáì ÇáãæÇÓã " & xMosm.text
    Load PrintGrd
    If grid1.Cols > 8 Then
        PrintGrd.DOPRINT Me.grid1, 1, -2, cHead1, , , False, True, 11
    Else
        PrintGrd.DOPRINT Me.grid1, 1, -2, cHead1, , , False, False, 11
    End If
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    StoreTable.Open "select * from STORE_BR order by code ", con, adOpenStatic, adLockReadOnly, adCmdText
    xDate1.Value = Date
'    DATA2.ConnectionString = strCon
'    DATA2.RecordSource = "Select Code,DescA From File4_10 order by Desca"
'    Set xSupp.RowSource = DATA2
'    xSupp.ListField = "Desca"
'    xSupp.BoundColumn = "Code"
'
'
'    data5.ConnectionString = strCon
'    data5.RecordSource = "Select code ,desca From fact ORDER BY DESCA "
'    Set xFact.RowSource = data5
'    xFact.ListField = "Desca"
'    xFact.BoundColumn = "Code"
    
    
    DATA5.ConnectionString = strCon
    DATA5.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC "
    Set xMosm.RowSource = DATA5
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    
    grid1.Rows = 1
    grid1.Cols = 1
'   FixGrid
End Sub
Private Sub myload()
    Dim cWhere As String
    Dim cStr2 As String
    Dim DataMosmTable As New ADODB.Recordset
    If xall(0).Value Then
        cStr1 = " SELECT FILE1_11.STORE , STORE_BR.desca FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_11.ITEM = FILE1_10.ITEM INNER JOIN STORE_BR ON STORE_BR.CODE = FILE1_11.STORE "
        cStr1 = cStr1 & " GROUP BY FILE1_11.STORE ,STORE_BR.desca ORDER BY STORE_BR.desca "
    End If
    If xall(1).Value Then
        cStr1 = " SELECT FR1_11.STORE , STORE_BR.desca FROM FILE1_10 INNER JOIN FR1_11 ON FR1_11.ITEM = FILE1_10.ITEM INNER JOIN STORE_BR ON STORE_BR.CODE = FR1_11.STORE "
        cStr1 = cStr1 & " GROUP BY FR1_11.STORE ,STORE_BR.desca ORDER BY STORE_BR.desca "
    End If
    If xall(2).Value Then
        cStr1 = " SELECT FILE1_11_ALL.STORE , STORE_BR.desca FROM FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_11_ALL.ITEM = FILE1_10.ITEM INNER JOIN STORE_BR ON STORE_BR.CODE = FILE1_11_ALL.STORE "
        cStr1 = cStr1 & " GROUP BY FILE1_11_ALL.STORE ,STORE_BR.desca ORDER BY STORE_BR.desca "
    End If
    DataMosmTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    If DataMosmTable.EOF And DataMosmTable.BOF Then
        Exit Sub
    End If
    DataMosmTable.MoveFirst
    Do Until DataMosmTable.EOF
        cWhere = " STORE = " & MyParn(DataMosmTable!STORE)
        If xall(1).Value Then
            cField2 = cField2 & turnFound(cField2, ",") & myiif(cWhere, "([IN] - [OUT]) * FILE1_10.COSTITEM_FR ")
        Else
            cField2 = cField2 & turnFound(cField2, ",") & myiif(cWhere, "([IN] - [OUT]) * FILE1_10.COSTITEM ")
        End If
        DataMosmTable.MoveNext
    Loop
    If xall(0).Value Then
        cString = " SELECT FILE1_10.SECTION , file1_10sc.desca  , " & cField2
        cString = cString & " FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM left join file1_10sc on file1_10sc.code = file1_10.[section] WHERE FILE1_10.ITEM IS NOT NULL "
    End If
    If xall(1).Value Then
        cString = " SELECT FILE1_10.SECTION , file1_10sc.desca  , " & cField2
        cString = cString & " FROM FILE1_10 INNER JOIN FR1_11 ON FILE1_10.ITEM = FR1_11.ITEM left join file1_10sc on file1_10sc.code = file1_10.[section] WHERE FILE1_10.ITEM IS NOT NULL "
    End If
    If xall(2).Value Then
        cString = " SELECT FILE1_10.SECTION , file1_10sc.desca  , " & cField2
        cString = cString & " FROM FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM left join file1_10sc on file1_10sc.code = file1_10.[section] WHERE FILE1_10.ITEM IS NOT NULL "
    End If
    
    If IsDate(xDate1.Value) Then cString = cString & " AND DATE <= " & DateSq(xDate1.Value)
    cString = cString & " group BY FILE1_10.SECTION , file1_10sc.desca ORDER BY FILE1_10.SECTION "
    DATA6.RecordSource = cString
    DATA6.Refresh
    FIXGRID
End Sub
Sub FIXGRID()
    Dim nTRow As Double
    Dim DataMosmTable As New ADODB.Recordset
    
    If xall(0).Value Then
        cStr1 = " SELECT FILE1_11.STORE , STORE_BR.desca FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_11.ITEM = FILE1_10.ITEM INNER JOIN STORE_BR ON STORE_BR.CODE = FILE1_11.STORE "
        cStr1 = cStr1 & " GROUP BY FILE1_11.STORE ,STORE_BR.desca ORDER BY STORE"
    End If
    If xall(1).Value Then
        cStr1 = " SELECT FR1_11.STORE , STORE_BR.desca FROM FILE1_10 INNER JOIN FR1_11 ON FR1_11.ITEM = FILE1_10.ITEM INNER JOIN STORE_BR ON STORE_BR.CODE = FR1_11.STORE "
        cStr1 = cStr1 & " GROUP BY FR1_11.STORE ,STORE_BR.desca ORDER BY STORE"
    End If
    If xall(2).Value Then
        cStr1 = " SELECT FILE1_11_ALL.STORE , STORE_BR.desca FROM FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_11_ALL.ITEM = FILE1_10.ITEM INNER JOIN STORE_BR ON STORE_BR.CODE = FILE1_11_ALL.STORE "
        cStr1 = cStr1 & " GROUP BY FILE1_11_ALL.STORE ,STORE_BR.desca ORDER BY STORE"
    End If
    DataMosmTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    
    DataMosmTable.MoveFirst
    
    With grid1
    .ColComboList(0) = StrList("SELECT CODE , DESCA FROM FILE1_10SC", con)
    If .Cols = 1 Then Exit Sub
    .FrozenCols = 1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .RowHeight(0) = 1000
    .WordWrap = True
    .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
    .TextMatrix(0, 1) = "ÇáÞÓã"
    .ColWidth(0) = 700
    .ColWidth(1) = 2000
    .ColDataType(0) = flexDTString
    nCol = 1
    Do Until DataMosmTable.EOF
       nCol = 1 + nCol
        .TextMatrix(0, nCol) = DataMosmTable!DESCA
        DataMosmTable.MoveNext
        .ColWidth(nCol) = 1600
        .ColDataType(nCol) = flexDTDouble
        .ColFormat(nCol) = "#.00"
    Loop
    .Cols = .Cols + 1
    .TextMatrix(0, .Cols - 1) = "ÇáÅÌãÇáì"
    .ColWidth(.Cols - 1) = 1400
    .ColDataType(.Cols - 1) = flexDTDouble
    .ColFormat(.Cols - 1) = "#.00"
    
    For nRow = 1 To .Rows - 1
        nTRow = 0
        For nCol = 2 To .Cols - 2
            .TextMatrix(nRow, nCol) = Format(Val(.TextMatrix(nRow, nCol)), "#0.00")
            nTRow = nTRow + Val(.TextMatrix(nRow, nCol))
        Next nCol
        .TextMatrix(nRow, .Cols - 1) = Format(nTRow, "#0.00")
    Next nRow
    .SubtotalPosition = flexSTAbove
    For nCol = 2 To .Cols - 1
        .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
    Next nCol
'    .TextMatrix(1, 1) = "ÇáÅÌãÇáì"
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_DblClick()
With grid1
    Load VsTItem
    VsTItem.xMosm.BoundText = .TextMatrix(0, .Col)
'    VsTItem.xBal.Value = True
    VsTItem.Show
End With
End Sub
Private Sub grid1_EnterCell()
    xModelNo.Caption = grid1.TextMatrix(grid1.Row, 0)
End Sub
Private Sub xSupp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then SuppLookupAll Me, oSearch
End Sub
Sub myProc()
If ActiveControl.Name = xSupp.Name Then
    xSupp.BoundText = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
End If
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
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

listarray(0, 0) = "ÇáæÕÝ"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "ÇáßæÏ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "ÇáæÕÝ"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "ÅÓÊÚáÇã "
Search3.Show 1

End If
End Sub


