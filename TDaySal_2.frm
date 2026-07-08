VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TDaySal_2 
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   FillColor       =   &H00008080&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10710
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab sstab1 
      Height          =   8250
      Left            =   135
      TabIndex        =   3
      Top             =   2430
      Width           =   15090
      _ExtentX        =   26617
      _ExtentY        =   14552
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "≈Ã„«·Ì »Ê‰«  «·»Ì⁄ "
      TabPicture(0)   =   "TDaySal_2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "print2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Grid1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " ›’Ì·Ì »Ê‰«  «·»Ì⁄"
      TabPicture(1)   =   "TDaySal_2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRID3"
      Tab(1).Control(1)=   "PRINT1"
      Tab(1).ControlCount=   2
      Begin VSFlex7Ctl.VSFlexGrid GRID3 
         Height          =   7080
         Left            =   -74820
         TabIndex        =   4
         Top             =   405
         Width           =   14775
         _cx             =   26061
         _cy             =   12488
         _ConvInfo       =   1
         Appearance      =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
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
      Begin VSFlex7Ctl.VSFlexGrid Grid1 
         Height          =   7125
         Left            =   90
         TabIndex        =   5
         Top             =   450
         Width           =   14910
         _cx             =   26300
         _cy             =   12568
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
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
      Begin Threed.SSCommand PRINT1 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   -63570
         TabIndex        =   10
         Top             =   7605
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   900
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TDaySal_2.frx":0038
         Caption         =   "ÿ»«⁄…  ›’Ì·Ï »Ê‰«  «·»Ì⁄"
         Alignment       =   4
         PictureAlignment=   10
      End
      Begin Threed.SSCommand print2 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   11475
         TabIndex        =   11
         Top             =   7695
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   900
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TDaySal_2.frx":2472
         Caption         =   "≈Ã„«·Ï »Ê‰«  «·»Ì⁄"
         Alignment       =   4
         PictureAlignment=   10
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1590
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   6270
      Begin VB.CommandButton cmdExit 
         Height          =   600
         Left            =   90
         Picture         =   "TDaySal_2.frx":48AC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo XMAN 
         Height          =   315
         Left            =   2025
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   225
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   2025
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   630
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5265
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   675
         Width           =   435
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·»«∆⁄ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   465
      End
   End
   Begin MSComCtl2.MonthView xDate 
      Height          =   2310
      Left            =   10395
      TabIndex        =   0
      Top             =   45
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12632256
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   336723975
      TitleBackColor  =   12632256
      TrailingForeColor=   12632256
      CurrentDate     =   40954
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2250
      Top             =   -90
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
      Left            =   45
      Top             =   -630
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
      Left            =   45
      Top             =   -1035
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
   Begin MSAdodcLib.Adodc DATABOX 
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
   Begin MSAdodcLib.Adodc data4 
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
Attribute VB_Name = "TDaySal_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
MYLOAD1
MYLOAD4
End Sub
Private Sub Form_Load()
    SSTab1.Tab = 1
    openCon con
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "SELECT * FROM FILE6_25"
    Set xMan.RowSource = DATA2
    xMan.ListField = "Desca"
    xMan.BoundColumn = "Code"
    
    data4.ConnectionString = strCon
    data4.RecordSource = "SELECT * FROM FILE0_40"
    Set XSTORE.RowSource = data4
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    
    If cBranch <> "00" Then
        XSTORE.BoundText = cBranchStore
        XSTORE.Enabled = False
    End If
    xDate.Visible = True
    
    Set GRID1.DataSource = data1
    data1.ConnectionString = strCon
    
    GRID1.Rows = 1
    GRID3.Rows = 1
    Fixgrd1
    Fixgrd4
    xDate.Value = Date
End Sub
Private Sub Fixgrd1()
With GRID1
.Cols = 12
.TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "«·Êﬁ "
.TextMatrix(0, 3) = "«·Œ“‰…"
.TextMatrix(0, 4) = "⁄œœ «·ﬁÿ⁄"
.TextMatrix(0, 5) = "≈Ã„«·Ï «·ﬁÌ„…"
.TextMatrix(0, 6) = "«·Œ’„"
.TextMatrix(0, 7) = "«·’«›Ï"
.TextMatrix(0, 8) = "”œ«œ ‰ﬁœÏ"
.TextMatrix(0, 9) = "”œ«œ ›Ì“«"
.TextMatrix(0, 10) = "⁄„Ì·"
.TextMatrix(0, 11) = "⁄„Ì· ﬂ«— "


.ColFormat(1) = "DD-MM-YYYY"
.ColFormat(2) = "SHORT TIME"
.ColFormat(4) = "#0"
.ColFormat(5) = "#0.00"
.ColFormat(6) = "#0.00"
.ColFormat(7) = "#0.00"
.ColFormat(8) = "#0.00"
.ColFormat(9) = "#0.00"
.ColDataType(1) = flexDTDate
.ColDataType(2) = flexDTShort
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(8) = flexDTDouble
.ColDataType(9) = flexDTDouble
.ColDataType(10) = flexDTString


.ColWidth(0) = 1200
.ColWidth(1) = 1300
.ColWidth(2) = 1500
.ColWidth(3) = 0
.ColWidth(4) = 1100
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(8) = 1100
.ColWidth(9) = 1100
.ColWidth(10) = 1000
.ColWidth(11) = 0

.FixedCols = 0
.FixedRows = 1

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 1) = DateValue(.TextMatrix(nRow, 1))
    .TextMatrix(nRow, 2) = TimeValue(.TextMatrix(nRow, 2))
Next nRow

.ExplorerBar = flexExSortShow
For nCol = 5 To .Cols - 1
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
Next nCol
End With
End Sub

Private Sub PRINT1_Click()
    Load PrintGrd
    PrintGrd.doprint GRID3, , , "≈Ã„«·Ï »Ê‰«  «·»Ì⁄ ··› —… „‰ " & xDate.Value, , , , True
    PrintGrd.Show 1
End Sub

Private Sub print2_Click()
    Load PrintGrd
    PrintGrd.doprint GRID1, 1.2, , "≈Ã„«·Ï »Ê‰«  «·»Ì⁄ ··› —… „‰ " & xDate.Value, , , , True
    PrintGrd.Show 1
End Sub

Private Sub xdate_DateClick(ByVal DateClicked As Date)
    MYLOAD1
    MYLOAD4
End Sub
Sub MYLOAD1()
    cString = "SELECT FILE6_20H.DOC_NO2,FILE6_20H.date ,FILE6_20H.TIME, ' ' AS N1 , SUM(FILE6_20.QUANT) , SUM(FILE6_20.TOTAL) ,FILE6_20H.discount , SUM(FILE6_20.TOTAL) - FILE6_20H.discount , FILE6_20H.CASH ,FILE6_20H.VISA , FILE3_10.DESCA, FILE3_20.DESCA " & _
            " FROM FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO LEFT OUTER JOIN FILE3_20 ON FILE6_20H.code2 = FILE3_20.CODE LEFT OUTER JOIN FILE3_10 ON FILE6_20H.code = FILE3_10.CODE "
    cString = cString & turn(cString) & "DATE = " & DateSq(xDate.Value)
    If xMan.BoundText <> "" Then cString = cString & turn(cString) & "FILE6_20.MAN = " & MyParn(xMan.BoundText)
    If XSTORE.BoundText <> "" Then cString = cString & turn(cString) & "FILE6_20h.store = " & MyParn(XSTORE.BoundText)
    cString = cString & " GROUP BY FILE6_20H.code2, FILE6_20H.code, FILE6_20H.DOC_NO, FILE6_20H.DOC_NO2, FILE6_20H.date, FILE6_20H.TIME, FILE6_20H.discount, FILE6_20H.CASH, FILE6_20H.VISA, FILE3_10.DESCA, FILE3_20.DESCA "
    cString = cString & " ORDER BY FILE6_20H.date, FILE6_20H.TIME, FILE6_20H.DOC_NO2 "
    data1.RecordSource = cString
    data1.Refresh
    Fixgrd1
End Sub
Private Sub MYLOAD4()
Dim cWhere As String, cString As String
If IsDate(xDate.Value) Then cWhere = turn(cWhere, " and ") & " DATE = " & DateSq(xDate.Value)

cString = "SELECT SALESDTL.* " & _
          " FROM SALESDTL"
If cWhere <> "" Then cString = cString & turn(cWhere) & cWhere
If xMan.BoundText <> "" Then cString = cString & turn(cString) & " MANDESCA = " & MyParn(xMan.text)
If XSTORE.BoundText <> "" Then cString = cString & turn(cString) & "store = " & MyParn(XSTORE.BoundText)
cString = cString & " ORDER BY DOC_NO,FLAG"

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
With GRID3
GRID3.Rows = 1
Do Until loctable.EOF
    .AddItem ""
    If loctable!Flag = 0 Then
        .TextMatrix(.Rows - 1, 0) = Mid(loctable!doc_no, 7, 4)
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Time, "HH:NN")
        .TextMatrix(.Rows - 1, 2) = loctable!MANDESCA & ""
        .TextMatrix(.Rows - 1, 3) = Format(loctable!modelfact0 & "")
        .TextMatrix(.Rows - 1, 4) = loctable!fact_Desca & ""
        .TextMatrix(.Rows - 1, 5) = loctable!Item & ""
        .TextMatrix(.Rows - 1, 6) = loctable!DESCA & ""
        .TextMatrix(.Rows - 1, 7) = loctable!Quant
        .TextMatrix(.Rows - 1, 8) = loctable!price & ""
        .TextMatrix(.Rows - 1, 9) = Format(Val(loctable!TOTAL & ""), "Fixed")
        nTotal = nTotal + Val(loctable!TOTAL & "")
    ElseIf loctable!Flag = 1 Then
        .TextMatrix(.Rows - 1, 0) = Mid(loctable!doc_no, 7, 4)
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = "«·Œ’„"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalDiscount = nTotalDiscount + Val(.TextMatrix(.Rows - 1, 8 + 1))
        If loctable!Flag = 0 Then nTotalSalesDis = nTotalSalesDis + Val(.TextMatrix(.Rows - 1, 8 + 1))
    ElseIf loctable!Flag = 3 Then
        .TextMatrix(.Rows - 1, 0) = Mid(loctable!doc_no, 7, 4)
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = "«·«Ã„«·Ì"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
    
        .AddItem ""
        .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &H8080&
    
    ElseIf loctable!Flag = 4 Then
        .TextMatrix(.Rows - 1, 0) = Mid(loctable!doc_no, 7, 4)
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = "‰ﬁœÌ…"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = Format(loctable!TOTAL, "#0.00")
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalCash = nTotalCash + Val(loctable!TOTAL & "")
    ElseIf loctable!Flag = 5 Then
        '.TextMatrix(.Rows - 1, 0) = LOCTABLE!doc_no
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = "›Ì“«"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = Format(loctable!TOTAL, "#0.00")
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalVisa = nTotalVisa + Val(loctable!TOTAL & "")
    End If
    loctable.MoveNext
Loop

If nTotal <> 0 Then
    .AddItem ""
    For i = 0 To 8
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·ÌÊ„"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = nTotal - nTotalDiscount
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If nTotalCash <> 0 Then
    .AddItem ""
    For i = 0 To 8
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·‰ﬁœÌ…"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = nTotalCash
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue

End If

If nTotalVisa <> 0 Then
    .AddItem ""
    For i = 0 To 7 + 1
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·›Ì“«"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = nTotalVisa
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If Round((nTotal - nTotalDiscount) - (nTotalCash + nTotalVisa), 0) <> 0 Then
    .AddItem ""
    For i = 0 To 8
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·¬Ã·"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = Round((nTotal - nTotalDiscount) - (nTotalCash + nTotalVisa), 2)
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

End With
Fixgrd4
End Sub
Private Sub Fixgrd4()
With GRID3
.Cols = 10
.TextMatrix(0, 0) = "—ﬁ„ „” ‰œ"
.TextMatrix(0, 1) = "«·Êﬁ "
.TextMatrix(0, 2) = "«·»«∆⁄"
.TextMatrix(0, 2 + 1) = "«·„ÊœÌ·"
.TextMatrix(0, 3 + 1) = "«·„’‰⁄"
.TextMatrix(0, 4 + 1) = "ﬂÊœ"
.TextMatrix(0, 5 + 1) = "«·’‰›"
.TextMatrix(0, 6 + 1) = "„»Ì⁄« "
.TextMatrix(0, 7 + 1) = "«·”⁄—"
.TextMatrix(0, 8 + 1) = "«·≈Ã„«·Ï"
.MergeCells = flexMergeFree
.MergeCol(0) = True
.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 1500
.ColWidth(2 + 1) = 2000
.ColWidth(3 + 1) = 1200
.ColWidth(4 + 1) = 1000
.ColWidth(5 + 1) = 2600
.ColWidth(6 + 1) = 1000
.ColWidth(7 + 1) = 1000
.ColWidth(8 + 1) = 1200
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
'.SubtotalPosition = flexSTBelow
'.Subtotal flexSTNone, -1, 0, , , &H8080&
End With
End Sub
