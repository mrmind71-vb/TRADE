VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsRepOrder 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ—Ì— ÿ·»Ì«  «·⁄„·«¡ ( ÿ·»Ì… - «·„‰›– - «·„»«⁄ )"
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   10635
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7785
         MaxLength       =   15
         TabIndex        =   11
         Top             =   630
         Width           =   1545
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4860
         MaxLength       =   15
         TabIndex        =   10
         Top             =   675
         Width           =   1545
      End
      Begin VB.TextBox XPO_NO 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1800
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   1770
      End
      Begin MSDataListLib.DataCombo XCODE 
         Height          =   315
         Left            =   4860
         TabIndex        =   12
         Top             =   225
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì· :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Ì‰  «—ÌŒ : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   9450
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   675
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "≈·Ï "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   6525
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   675
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No P.O"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8025
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1260
      Width           =   20355
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   7665
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   20130
         _cx             =   35507
         _cy             =   13520
         _ConvInfo       =   1
         Appearance      =   0
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
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
   End
   Begin VB.Frame Frame2 
      Height          =   870
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   495
      Width           =   5055
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   1485
         Picture         =   "VsRepOrder.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "⁄—÷"
         Top             =   225
         Width           =   1410
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   90
         Picture         =   "VsRepOrder.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   225
         Width           =   1410
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   600
         Left            =   3015
         TabIndex        =   6
         Top             =   180
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   1058
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
         Picture         =   "VsRepOrder.frx":495E
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
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
      Left            =   -90
      Top             =   270
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
   Begin MSAdodcLib.Adodc Adodc1 
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
   Begin ComctlLib.ProgressBar prog1 
      Height          =   285
      Left            =   90
      TabIndex        =   7
      Top             =   135
      Width           =   5010
      _ExtentX        =   8837
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
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
   Begin MSAdodcLib.Adodc DATA8 
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
   Begin MSAdodcLib.Adodc Adodc2 
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
Attribute VB_Name = "VsRepOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Variant, sM_Fact As String
Dim oSearch As New Search3
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    con.Execute " UPDATE FILE6_52 SET PRICE =  (SELECT FILE6_51.PRICE FROM FILE6_51 WHERE FILE6_51.DOC_NO = FILE6_52H.PO_NO AND FILE6_52.ITEM = FILE6_51.ITEM ) FROM FILE6_52 INNER JOIN FILE6_52H ON FILE6_52.DOC_NO = FILE6_52H.DOC_NO "
    DoDropTable
    DoTable
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    
    DATA3.ConnectionString = strCon
    DATA3.RecordSource = "Select Code,DescA From File4_10 ORDER BY DESCA"
    Set xcode.RowSource = DATA3
    xcode.ListField = "Desca"
    xcode.BoundColumn = "Code"
    
    
    Set grid1.DataSource = data6
    data6.ConnectionString = strCon
    
    grid1.Rows = 1
    grid1.Cols = 12
    Fixgrd
End Sub
Private Sub myload()
    Dim cWhere As String, cString  As String
'                           0                           1                   2                   3                                                       4                           5                                   6                                   7                                       8                               9
    cString = " SELECT  T_ORDER_REC_SALES.DOC_NO, T_ORDER_REC_SALES.DATE,FILE3_10.DESCA , COUNT(DISTINCT T_ORDER_REC_SALES.ITEM) AS CountItem, SUM(T_ORDER_REC_SALES.QUANT) , SUM(T_ORDER_REC_SALES.VALUE_ORDER) , MAX(T_ORDER_REC_SALES.DATE_REC) , COUNT(DISTINCT T_ORDER_REC_SALES.ITEM2) , SUM(T_ORDER_REC_SALES.RECEVD) , SUM(T_ORDER_REC_SALES.VALUE_REC) ,  MAX(T_ORDER_REC_SALES.DATE_SALES) , COUNT(DISTINCT T_ORDER_REC_SALES.ITEM3) AS Expr8, SUM(T_ORDER_REC_SALES.SALES) AS Expr9,  SUM(T_ORDER_REC_SALES.VALUE_SALES) AS Expr10  " & _
            " " & _
            " FROM T_ORDER_REC_SALES INNER JOIN FILE1_10 ON T_ORDER_REC_SALES.ITEM = FILE1_10.ITEM  LEFT JOIN FILE3_10 ON FILE3_10.CODE = T_ORDER_REC_SALES.CODE WHERE DOC_NO IS NOT NULL "
    If xcode.BoundText <> "" Then cString = cString & " AND  T_ORDER_REC_SALES.[code] = " & MyParn(xcode.BoundText)
    If IsDate(xDate1.Text) Then cString = cString & " AND  DATE >= " & DateSq(xDate1.Text)
    If IsDate(xDate2.Text) Then cString = cString & " AND  DATE <= " & DateSq(xDate2.Text)
    If xpo_no.Text <> "" Then cString = cString & " AND  DOC_NO = " & MyParn(xpo_no.Text)
    cString = cString & " GROUP BY FILE3_10.DESCA , T_ORDER_REC_SALES.DOC_NO, T_ORDER_REC_SALES.DATE ORDER BY T_ORDER_REC_SALES.DATE, T_ORDER_REC_SALES.DOC_NO"
    data6.RecordSource = cString
    data6.Refresh
    Fixgrd
End Sub
Sub Fixgrd()
    Dim nRate1 As Double, nRate2 As Double
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .Cols = 17
    .RowHeight(0) = 1000
    
    .TextMatrix(0, 0) = "—ﬁ„ ÿ·»Ì…"
    .TextMatrix(0, 1) = " «—ÌŒ"
    
    .TextMatrix(0, 2) = "«·⁄„Ì·"
    
    .TextMatrix(0, 3) = "⁄œœ «·√’‰«›"
    .TextMatrix(0, 4) = "«·ﬂ„Ì…"
    .TextMatrix(0, 5) = "«·ﬁÌ„…"
    
    .TextMatrix(0, 6) = " «—ÌŒ «Œ— «” ·«„"
    .TextMatrix(0, 7) = "⁄œœ «’‰«›  „ «” ·«„Â«"
    .TextMatrix(0, 8) = "ﬂ„Ì…  „ «” ·«„Â«"
    .TextMatrix(0, 9) = "ﬁÌ„…  „ «” ·«„Â«"
    
    .TextMatrix(0, 10) = " «—ÌŒ „»Ì⁄« "
    .TextMatrix(0, 11) = "⁄œœ «’‰«›  „ »Ì⁄Â«"
    .TextMatrix(0, 12) = "ﬂ„Ì…  „ »Ì⁄Â«"
    .TextMatrix(0, 13) = "ﬁÌ„… «·»Ì⁄"
    
    .TextMatrix(0, 14) = "‰”»… ﬁÌ„… «·≈” ·«„"
    .TextMatrix(0, 15) = "‰”»… ﬁÌ„… «·»Ì⁄"
    
    .TextMatrix(0, 16) = " ›’Ì·Ï «·ÿ·»Ì…"
    .ColComboList(16) = "..."
    .ColWidth(0) = 1500
    .ColWidth(1) = 1500
    .ColWidth(2) = 2000
    .ColWidth(3) = 800
    .ColWidth(4) = 800
    .ColWidth(5) = 1300
    .ColWidth(6) = 1300
    .ColWidth(7) = 800
    .ColWidth(8) = 800
    .ColWidth(9) = 1300
    
    .ColWidth(10) = 1300
    .ColWidth(11) = 800
    .ColWidth(12) = 800
    .ColWidth(13) = 1300
    
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTDate
    For nCol = 3 To .Cols - 1
        .ColDataType(nCol) = flexDTDouble
    Next nCol
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 3, "#0", vbYellow, , True, "  "
    .Subtotal flexSTSum, -1, 4, "#0", vbYellow, , True, "  "
    .Subtotal flexSTSum, -1, 5, "#0", vbYellow, , True, "  "
    .Subtotal flexSTSum, -1, 7, "#0", vbYellow, , True, "  "
    .Subtotal flexSTSum, -1, 8, "#0", vbYellow, , True, "  "
    .Subtotal flexSTSum, -1, 9, "#0", vbYellow, , True, "  "
    .Subtotal flexSTSum, -1, 11, "#0", vbYellow, , True, "  "
    .Subtotal flexSTSum, -1, 12, "#0", vbYellow, , True, "  "
    .Subtotal flexSTSum, -1, 13, "#0", vbYellow, , True, "  "
    
    For nRow = 1 To .Rows - 1
        nRate1 = 0
        nRate2 = 0
        If Val(.TextMatrix(nRow, 5)) > 0 Then nRate1 = Round((Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 5))) * 100, 2)
        If Val(.TextMatrix(nRow, 5)) > 0 Then nRate2 = Round((Val(.TextMatrix(nRow, 13)) / Val(.TextMatrix(nRow, 5))) * 100, 2)
        .TextMatrix(nRow, 14) = nRate1
        .TextMatrix(nRow, 15) = nRate2
    Next nRow
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Sub myProc()
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
If ActiveControl.Name = xcode.Name Then
    xcode.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    VsTMoveItem.xpo_no.Text = grid1.TextMatrix(grid1.Row, 0)

End Sub

Private Sub grid1_EnterCell()
With grid1
    If .Col = .Cols - 1 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
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
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdGo_Click
    End If
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Set Generalarray(0) = Me
Generalarray(1) = "Select code ,DescA From FILE3_10 "
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

Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
imgx1.Images.Clear
imgx1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub
Sub DoDropTable()
Dim cString As String
On Error Resume Next
    con.Execute " DROP TABLE T_ORDER_REC_SALES"
    Err.Clear
End Sub
Sub DoTable()
Dim cString As String
On Error Resume Next
    con.Execute " SELECT   DOC_NO, CODE, DATE, ITEM, PRICE, QUANT, RECEVD, SALES, DOC_REC, DATE_REC, DOC_SALES, DATE_SALES , VALUE_ORDER , VALUE_REC , VALUE_SALES , ITEM1 , ITEM2 , ITEM3 Into T_ORDER_REC_SALES  FROM Q_ORDER_REC_SALES"
    cString = "CREATE NONCLUSTERED INDEX [IX_T_ORDER_REC_SALES] ON [dbo].[T_ORDER_REC_SALES] (     [ITEM] ASC,     [DOC_NO] ASC  ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
    createCommand cString, con
    Err.Clear
End Sub


