VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form damageFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· «·›"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15090
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   15090
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   3150
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   1260
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "«÷«›… «’‰«› «·„Œ“‰"
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   180
         Width           =   2355
      End
   End
   Begin VB.Frame frmBar 
      Height          =   645
      Left            =   7740
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   0
      Width           =   1725
      Begin VB.CommandButton cmdPrint 
         Caption         =   "ÿ»«⁄… «·„” ‰œ"
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
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   135
         Width           =   1545
      End
   End
   Begin VB.Frame Frame6 
      Height          =   555
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   -90
      Visible         =   0   'False
      Width           =   3660
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   135
         Width           =   3540
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   9495
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   5550
      Begin VB.CommandButton CmdDelInv 
         Caption         =   "Õ–› «·„” ‰œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   180
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
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2790
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4140
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   630
      Width           =   7845
      Begin VB.TextBox xRemark 
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
         Height          =   315
         Left            =   105
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   6435
      End
      Begin VB.TextBox xDoc_No 
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
         Height          =   315
         Left            =   5265
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1290
      End
      Begin VB.TextBox xDate 
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
         Height          =   315
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2235
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   3825
         TabIndex        =   2
         Top             =   540
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   540
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   585
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "„·«ÕŸ«  "
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
         Left            =   6645
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   945
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Œ“‰ :"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   585
         Width           =   540
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   210
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   5715
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   945
      Width           =   1455
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   540
         Width           =   1320
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   1305
      Top             =   360
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   661
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
      Caption         =   "data1"
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   720
      Top             =   0
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6045
      Left            =   45
      TabIndex        =   30
      Top             =   1980
      Width           =   15000
      _cx             =   26458
      _cy             =   10663
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   4
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
   Begin VB.Frame Frame5 
      Height          =   600
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   8010
      Width           =   7890
      Begin VB.Label xTotalCost 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   180
         Width           =   1440
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «· ﬂ·›… :"
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·≈Ã„«·Ì:"
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
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
         Width           =   780
      End
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   1440
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   375
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   661
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
      Caption         =   "data1"
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
      Height          =   375
      Left            =   720
      Top             =   0
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
      _ExtentY        =   661
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
      Caption         =   "data1"
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
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   13095
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   8010
      Width           =   1950
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   150
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   570
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   150
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   150
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1455
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Move Last"
         Top             =   150
         Width           =   435
      End
   End
End
Attribute VB_Name = "DamageFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte, bEdit As Boolean
Dim Search31 As New Search3, search32 As New Search3
Dim cHeader As String, cSub As String, cTitle As String, bHideBar As Boolean, cFilter_Store As String
Dim CardTable As ADODB.Recordset
Dim tBalStore  As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca,file1_50.desca,file1_10.price,file1_10.price2 From file1_10 left join file1_50 on file1_10.[GROUP] = file1_50.code"
Generalarray(2) = "Order by file1_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter_store%' or  %%FILE1_10.DESCA%%) "

listarray(1, 0) = "«·„Ã„Ê⁄…"
listarray(1, 1) = "(%%FILE1_50.DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·’‰›"
GrdArray(0, 1) = 1500

GrdArray(1, 0) = "≈”„ «·’‰›"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "”⁄— «·Ã„·…"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "«·ﬁÿ«⁄Ì"
GrdArray(4, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„ «·«’‰«›"
Search3.Show 1
End Sub
Private Function myreplace() As Boolean
Dim aInsert(4, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(xDate.Text)

aInsert(2, 0) = "STORE"
aInsert(2, 1) = addstring(xStore.BoundText)

aInsert(3, 0) = "SECTION"
aInsert(3, 1) = addstring(xSection.BoundText)

aInsert(4, 0) = "Remark"
aInsert(4, 1) = addstring(xRemark.Text)

con.BeginTrans
If xDoc_No.Enabled Then
    xDoc_No.Text = RetZero(Val(Newflag(cHeader, "doc_no")))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, cHeader)
Else
    con.Execute CreateUpdate(aInsert, cHeader, " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myproc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(search32.grid1.TextMatrix(search32.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    grid1.TextMatrix(grid1.Row, 0) = search32.grid1.TextMatrix(search32.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = "1"
    GrdDesc grid1.Row
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 2) = ""
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 0
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.TextMatrix(grid1.Rows - 2, 2) = ""
        grid1.Select grid1.Rows - 1, 0
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_no = " & MyParn(Search31.grid1.TextMatrix(Search31.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    MyLoad
    Search31.Hide
ElseIf ActiveControl.Name = xDoc_No.Name Then
    xDoc_No.Text = Search31.grid1.TextMatrix(Search31.grid1.Row, 0)
    Search31.Hide
Else
    ActiveControl.Text = search32.grid1.TextMatrix(search32.grid1.Row, 0)
    Unload search32
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Unload Search
End Sub

Private Sub cmdBar_Click()
With grid1
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 2)) > 0 Then
            con.Execute "Insert Into ADDPRINT(Item,Quant,isPrint) " & _
                " Values(" & _
                addstring(.TextMatrix(i, 0)) & "," & _
                addvalue(.TextMatrix(i, 2)) & "," & _
                "1" & _
                ")"
        End If
    Next i
End With
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute " Delete  From " & cSub & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute " Delete  From " & cHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    CardTable.Requery
    
    CmdNewInv_Click
    Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
If MsgBox("Œ—ÊÃ !! ” ›ﬁœ ﬂ· «·»Ì«‰«  «·€Ì— „Õ›ÊŸ… ! „Ê«›ﬁ ø", vbYesNo + vbDefaultButton2) = vbYes Then Unload Me
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
cString = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA,REMARK " & _
                  " FROM " & cHeader & " INNER JOIN FILE0_40 ON " & cHeader & " .Store = FILE0_40.CODE "
If cFilter_Store <> "" Then cString = cString & turn(cString) & "STORE IN(" & cFilter_Store & ")"

Generalarray(1) = cString
Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = " @@Doc_No@@6 OR " & _
                  " ##[DATE]##"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "„Œ“‰"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "„·ÕÊŸ…"
GrdArray(4, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search31
Search31.Caption = "«” ⁄·«„"
Search31.Show 1
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    MyLoad
End If
End Sub
Private Sub CmdNewInv_Click()
mydefine
On Error Resume Next
xDoc_No.SetFocus
End Sub

Private Sub cmdPrint_Click()
    doprint
End Sub

Private Sub cmdSave_Click()
foundOther
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Requery
'CardTable.FindFirst "Doc_No = " & MyParn(xDoc_No.Text)
'If xDoc_No.Enabled Then
    'CmdNewInv_Click
'Else
    CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    MyLoad
'End If
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.BOF And CardTable.EOF Then
    mydefine
    Exit Sub
Else
    CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    MyLoad
End If
End Sub

Private Sub Command1_Click()
If Trim(xStore.BoundText) = "" Then Exit Sub
nFirstRow = grid1.Rows - 2
FillStockItems Me.grid1, xStore.BoundText, con
For i = nFirstRow To grid1.Rows - 2
    GrdDesc i
Next
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If Shift = 2 And KeyCode = 83 Then cmdSave_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
'dLastdate = lastDate("FILE1_60")
openCon con

If myPublic = 1 Then
    cSub = "FILE1_82"
    cTitle = "Â«·ﬂ"
    bHideBar = True
ElseIf myPublic = 2 Then
    cSub = "FILE1_80"
    cTitle = "Ê«—œ"
    bHideBar = flase
ElseIf myPublic = 3 Then
    cSub = "FILE1_81"
    cTitle = "’«œ—"
    bHideBar = flase
End If

cHeader = cSub & "H"
Me.Caption = "„” ‰œ " & cTitle

cFilter_Store = CreateFilter
Set CardTable = New ADODB.Recordset
cString = "SELECT * FROM " & cHeader
If cFilter_Store <> "" Then cString = cString & turn(cString) & "STORE IN(" & cFilter_Store & ")"
cString = cString & " ORDER BY DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText


cString = "select * from file0_40"
If cFilter_Store <> "" Then cString = cString & turn(cString) & "code IN(" & cFilter_Store & ")"

data1.ConnectionString = strCon
data1.RecordSource = cString
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "FILE1_10SC"
Set xSection.RowSource = DATA2
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

Set grid1.DataSource = DATA10
DATA10.ConnectionString = strCon

If Not CardTable.EOF Then
    CardTable.MoveLast
    MyLoad
Else
    mydefine
    FixGrd
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload Search3
Unload Search31
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
SetKbLayout Lang_AR
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Set GRDTABLE = Nothing
closeCon con
Err.Clear
End Sub

Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If grid1.Col = 0 Then GrdDesc grid1.Row
CalcTotals
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 0 Or grid1.Col = 2 Then
    grid1.Editable = flexEDKbdMouse
    'SetKbLayout IIf(grid1.Col = 0, Lang_EN, Lang_AR)
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_GotFocus()
With grid1
    If grid1.Row <= 1 Then
    .Select 1, 0, 1, 0
    .ShowCell 1, 0
    End If
End With
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 Then grid1.AddItem "", grid1.Row
End Sub
Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 13 And grid1.Col = 0 Then
    If grid1.Row = grid1.Rows - 1 Then
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 1
    Else
        grid1.Select grid1.Row + 1, 1
    End If
End If

If KeyAscii = 13 Then
    Select Case Col
        Case 0
            grid1.Col = 2
            grid1.Row = Row
        Case 2
            grid1.Row = Row + 1
            grid1.Col = 0
     End Select
End If

End Sub

Private Sub grid1_LostFocus()
SetKbLayout Lang_AR
End Sub

Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then grid1.AddItem ""
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 And Trim(grid1.EditText) <> "" Then
    cItem = GetDesca("select item from file1_10 where item = " & MyParn(grid1.EditText)) & ""
    If cItem = "" Then
        MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        grid1.EditText = ""
        Exit Sub
    End If
    
        
    nFound = FoundOtheritem(Row, Col, Trim(grid1.EditText))
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
        Cancel = True
    End If
End If
'If Col = 2 And myPublic = 1 Then
'    If Val(grid1.TextMatrix(Row, 3)) < Val(grid1.EditText) Then
'        MsgBox "—’Ìœ «·’‰› €Ì— ﬂ«›Ì"
'        Cancel = True
'    End If
'End If
End Sub

Private Sub xDate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CmdInform_Click
End Sub
Private Function MYVALID() As Boolean
CardTable.Find "DOC_NO = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF And xDoc_No.Enabled Then
    MsgBox "„” ‰œ »‰›” «·—ﬁ„ „‰ ﬁ»·"
    Exit Function
End If

If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
'If IsDate(dLastdate) Then
'    If DateValue(xDate.Text) <= DateValue(dLastdate) Then
'        MsgBox "«· «—ÌŒ «ﬁ· „‰ «Œ—  «—ÌŒ «€·«ﬁ"
'        Exit Function
'    End If
'End If
If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·«Ê·"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If


With grid1
For i = 1 To .Rows - 2
    If .TextMatrix(i, 0) = "" Then
        .Select i, 0, i, grid1.Cols - 1
        MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
        Exit Function
    Else
        cItem = GetDesca("select item from file1_10 where item = " & MyParn(.TextMatrix(i, 0))) & ""
        If cItem = "" Then
            MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
            Exit Function
        End If
    End If
    If Val(.TextMatrix(i, 2)) = 0 Then
        .Select i, 0, i, grid1.Cols - 1
        MsgBox "ﬂ„Ì… «·’‰› €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub MyLoad()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!store
xSection.BoundText = CardTable!Section & ""
Me.xRemark.Text = CardTable!remark & ""
myloadgrd
handleControls LoadMode
CalcTotals
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Val(Newflag(cHeader, "doc_no")))
xusername.Text = ""
xDate.Text = Format(Date, "dd-mm-yyyy")
xStore.BoundText = ""
xSection.BoundText = ""

'StatusBar1.Panels(1).Text = ""
xTotal.Caption = ""
grid1.Rows = 1
grid1.AddItem ""
handleControls DefineMode
FixGrd
End Sub
Private Sub handleControls(nMode)
cmdNewinv.Enabled = nMode = LoadMode And bEdit
CmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = nMode = LoadMode And bEdit
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
Command1.Enabled = bEdit
End Sub
Private Sub xDoc_No_LostFocus()
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.BOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad
End Sub
Private Sub Grid1_ChangeEdit()
'If Grid1.Col = 1 Then GrdDesc Grid1.Row
'CalcTotals
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim sGroup As String, cwhere As String
If KeyCode = 112 And grid1.Col = 0 Then
    cwhere = "FILE1_10.ISRAW = 1"
    If xStore.BoundText <> "" Then sGroup = GetDesca("Select [group] from file0_40 where code = " & MyParn(xStore.BoundText))
    If sGroup <> "" Then
        cwhere = cwhere & turn(cwhere, " and ") & "FILE1_50.[GROUP] = " & MyParn(sGroup)
    End If
    ItemsLookupAll Me, search32, cwhere
End If


If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And bEdit Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from " & cSub & " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case Col
    Case 0
        If KeyCode = 27 Then Exit Sub
        If KeyCode = 112 Then ItemsLookup
End Select
End Sub
Private Sub GrdDesc(Row)
Dim nBalance As Double
grid1.TextMatrix(Row, 1) = ""
grid1.TextMatrix(Row, 3) = nBalance
If grid1.TextMatrix(Row, 0) = "" Then Exit Sub


aret = aGetDesca("select desca,unit,Price from file1_10 where item = " & MyParn(grid1.TextMatrix(Row, 0)))
If UBound(aret) > 0 Then
    grid1.TextMatrix(Row, 1) = aret(1) & ""
    grid1.TextMatrix(Row, 4) = aret(2) & ""
    grid1.TextMatrix(Row, 5) = aret(3) & ""
End If

If IsDate(xDate.Text) And Trim(xStore.BoundText) <> "" And Trim(grid1.TextMatrix(Row, 0)) <> "" Then
    grid1.TextMatrix(Row, 3) = LastBalance(grid1.TextMatrix(Row, 0), xStore.BoundText, con)
Else
    grid1.TextMatrix(Row, 3) = "0"
End If

'tBalStore.Filter = "ITEM = " & MyParn(Grid1.TextMatrix(Row, 0)) & " AND STORE = " & MyParn(xstore1.BoundText)
'If Not tBalStore.EOF Then nBalance = Round(Val(tBalStore!BAL & ""), 2)

'Grid1.TextMatrix(Row, 3) = nBalance
'If Val(Grid1.TextMatrix(Row, 3)) = 0 Then MsgBox "·« ÌÊÃœ —’Ìœ ··’‰›"

End Sub
Private Function CalcTotals()
Dim nTotalQuant As Double, nTotalCost As Double
With grid1
For i = 1 To grid1.Rows - 2
'    grid1.TextMatrix(I, 5) = Val(grid1.TextMatrix(I, 2)) * Val(grid1.TextMatrix(I, 4))
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 2))
'    nTotalCost = nTotalCost + Val(grid1.TextMatrix(I, 5))
Next
'Me.StatusBar1.Panels(1).Text = IIf(nTotalQuant <> 0, "≈Ã„«·Ì «·ﬂ„Ì… : " & nTotalQuant, "")
'xtotal.Caption = nTotalQuant
End With
End Function
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub foundOther()
For i = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(i, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ ====> " & nRow
        Exit Sub
    End If
Next
End Sub
Private Sub doprint()
Dim TempTable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
TempTable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For i = 1 To grid1.Rows - 2
    TempTable.AddNew
    TempTable!STR21 = "„” ‰œ " & cTitle & Format(xDoc_No.Text)
    TempTable!str3 = TurnValue(xRemark.Text)
    TempTable!date3 = DateFix(xDate.Text)
    TempTable!str2 = TurnValue(xStore.Text)
    TempTable!str4 = TurnValue(grid1.TextMatrix(i, 0))
    TempTable!str5 = TurnValue(grid1.TextMatrix(i, 1))
    TempTable!VAL2 = TurnValue(Val(grid1.TextMatrix(i, 2)))
    TempTable!VAL1 = TurnValue(Val(grid1.TextMatrix(i, 4)))
    TempTable!VAL3 = TurnValue(Val(grid1.TextMatrix(i, 5)))
    TempTable!Val4 = TurnValue(Val(xTotalCost.Caption))
    TempTable!Val10 = i
    If Val(xTotal.Caption) <> 0 Then
        TempTable!STR6 = MyOnly(Val(xTotalCost.Caption))
    End If
    TempTable.Update
Next
If TempTable.EOF And TempTable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
TempTable.Requery
If myPublic = 1 Then
    main.Report1.ReportFileName = App.Path & "\Reports\DAMEGE.rpt"
ElseIf myPublic = 2 Then
    main.Report1.ReportFileName = App.Path & "\Reports\R_INPUT.rpt"
End If
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
TempTable.Close
Set TempTable = Nothing
End Sub
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub FixGrd()
With grid1
    .Cols = 8
    .FormatString = "ﬂÊœ|" & "«·’‰›|" & "«·ﬂ„Ì…|" & "«·—’Ìœ|" & "«·ÊÕœ…|" & "”⁄— Ã„·…|" & "«·«Ã„«·Ì|" & ""
    .ColWidth(0) = 2500
    .ColWidth(1) = 7000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1200
    .ColWidth(4) = 3000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
       
    .ColHidden(.Cols - 1) = True
    .ColHidden(3) = True
    .ColHidden(5) = True
    .ColHidden(6) = True
   
    
    
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(5) = flexAlignRightCenter
End With
End Sub
Private Sub myreplaceGrd()
Dim aInsert(4, 1)
With grid1
    For i = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 0))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = .TextMatrix(i, 2)

        aInsert(3, 0) = "cost"
        aInsert(3, 1) = Val(.TextMatrix(i, 5))

        aInsert(4, 0) = "row"
        aInsert(4, 1) = i
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, cSub)
        Else
            con.Execute CreateUpdate(aInsert, cSub, " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub myloadgrd()
cString = "Select " & cSub & ".ITEM,FILE1_10.DESCA," & cSub & ".Quant,' ' as Balance,FILE1_10.UNIT," & cSub & ".Cost, '' as Total," & cSub & ".ID" & _
          " From " & cSub & " inner join file1_10 on " & cSub & ".item = file1_10.item WHERE " & cSub & ".DOC_NO = " & MyParn(xDoc_No.Text) & _
          " Order by " & cSub & ".Row"
DATA10.RecordSource = cString
DATA10.Refresh
grid1.AddItem ""
FixGrd
End Sub

Private Sub xSection_LostFocus()
If Not xSection.MatchedWithList Then xSection.BoundText = ""
End Sub
Private Function CreateFilter() As String
Dim loctable As New ADODB.Recordset
cString = "Select code from sec_store"
If nusercode <> "" Then cString = cString & turn(cString) & " user_code = " & nusercode
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    CreateFilter = CreateFilter & turn(CreateFilter, ",") & MyParn(loctable!code)
    loctable.MoveNext
Loop
loctable.Close
Set loctable = Nothing
End Function
