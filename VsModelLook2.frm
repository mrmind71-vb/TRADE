VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsModelLook2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈” ⁄·«„ „ÊœÌ·«  & »«—ﬂÊœ"
   ClientHeight    =   6420
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   12705
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   12705
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   3555
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   45
      Width           =   9060
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   5850
         Locked          =   -1  'True
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1305
         Width           =   1590
      End
      Begin VB.TextBox xDescA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2655
         Locked          =   -1  'True
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   900
         Width           =   4785
      End
      Begin VB.TextBox xModelFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   135
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox xmodel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   135
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   540
         Visible         =   0   'False
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   5085
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XMOSM 
         Height          =   315
         Left            =   2835
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   5085
         TabIndex        =   20
         Top             =   540
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label xCodeDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1305
         Width           =   3075
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈”„ «·„ÊœÌ·  :"
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1305
         Width           =   885
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄… «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   555
         Width           =   1425
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’‰⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   195
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " ›’Ì·Ï «·„ﬁ«”"
      Height          =   1185
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   675
      Width           =   2085
      Begin VB.TextBox xbal 
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
         Left            =   90
         MaxLength       =   15
         TabIndex        =   8
         Top             =   720
         Width           =   1275
      End
      Begin Threed.SSCommand XBARCODE 
         Height          =   510
         Left            =   90
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "0"
         ButtonStyle     =   4
         PictureAlignment=   1
         BevelWidth      =   10
         ShapeSize       =   1
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1455
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   765
         Width           =   570
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "»«—ﬂÊœ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1470
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   375
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdExit 
      Height          =   510
      Left            =   45
      Picture         =   "VsModelLook2.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      Width           =   1635
   End
   Begin VB.Frame Frame4 
      Caption         =   "√—’œ… „ﬁ«”«  Ê«·Ê«‰ «·„ÊœÌ·"
      Height          =   3930
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2025
      Width           =   12480
      Begin VSFlex7Ctl.VSFlexGrid Grid1 
         Height          =   3495
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   12300
         _cx             =   21696
         _cy             =   6165
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
         FixedRows       =   3
         FixedCols       =   2
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
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   6090
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   135
      Top             =   1170
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
      Left            =   135
      Top             =   1305
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
      Left            =   180
      Top             =   1080
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
      Left            =   135
      Top             =   1395
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
      Left            =   135
      Top             =   1260
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
      Left            =   135
      Top             =   1125
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
   Begin VSFlex7Ctl.VSFlexGrid Grid2 
      Height          =   765
      Left            =   1050
      TabIndex        =   3
      Top             =   1305
      Visible         =   0   'False
      Width           =   285
      _cx             =   503
      _cy             =   1349
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      Rows            =   1
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
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
   Begin Threed.SSCommand cmdgo 
      Height          =   510
      Left            =   1800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   900
      _Version        =   196610
      ForeColor       =   0
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
      Picture         =   "VsModelLook2.frx":246C
      Caption         =   "«·—’Ìœ"
      Alignment       =   4
      ButtonStyle     =   1
      PictureAlignment=   1
      BevelWidth      =   10
      ShapeSize       =   1
   End
End
Attribute VB_Name = "VsModelLook2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cMyModel As String
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Public CardTable As New ADODB.Recordset
Private Sub cmd_look_Click()
    ModelLookupAll Me, Search3
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        ModelLookupAll Me, Search3
    End If
End Sub
Private Sub Form_Load()


con.CursorLocation = adUseClient
con.Open strCon
CardTable.Open "select file1_10h.*,file4_10.desca as supDesca from file1_10h inner JOIN file4_10 on file1_10h.code = file4_10.code   order by model", con, adOpenStatic, adLockReadOnly, acmdtext

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FACT ORDER BY DESCA "
Set xFact.RowSource = data1
xFact.ListField = "DESCA"
xFact.BoundColumn = "CODE"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM FILE1_50 ORDER BY DESCA"
Set xGroup.RowSource = DATA2
xGroup.ListField = "DESCA"
xGroup.BoundColumn = "CODE"


data4.ConnectionString = strCon
data4.RecordSource = "SELECT * FROM MOSM ORDER BY DATE DESC "
Set xMosm.RowSource = data4
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"
ModelLookupAll Me, Search3
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub GRID1_Click()
With grid1
    xBal.text = .TextMatrix(.Row, .Col)
    xbarcode.Caption = GetDesca("select item from file1_10 where model = " & MyParn(XMODEL.text) & " and scal = " & MyParn(.TextMatrix(0, .Col)) & " and c_color = " & Val(.TextMatrix(.Row, 0)), con)
End With
End Sub
Private Sub myDefineGrd()

grid2.Rows = 3
grid2.Cols = 2

grid1.Rows = 3
grid1.Cols = 2

grid1.MergeCells = flexMergeRestrictRows
grid1.TextMatrix(0, 1) = "«·„ﬁ«”"
grid1.TextMatrix(1, 1) = "”⁄— „’‰⁄"
grid1.TextMatrix(2, 1) = "”⁄— „” Â·ﬂ"
grid1.RowHidden(1) = True

End Sub
Private Sub FillItem(cFieldas As String, cField As String)
Dim GrdTable As New ADODB.Recordset
' „·∆ «·ÃœÊ·
cString = "Select c_color as [—ﬁ„ «··Ê‰] , Color  as [«··Ê‰]" & turn(cFieldas, ",") & cFieldas & _
          " From " & _
          " (Select c_color,Color,scal,item from file1_10 WHERE MODEL = " & MyParn(XMODEL.text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(item)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  " & _
          " order by pvt.c_color"

GrdTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grid1.Cols = GrdTable.Fields.Count: grid2.Cols = GrdTable.Fields.Count

For nCol = 2 To GrdTable.Fields.Count - 1
    grid1.TextMatrix(0, nCol) = GrdTable.Fields(nCol).Name
Next

Do Until GrdTable.EOF
    grid2.AddItem ""
    grid1.AddItem ""
    For nCol = 0 To GrdTable.Fields.Count - 1
        If nCol <= 1 Then
            grid1.TextMatrix(grid2.Rows - 1, nCol) = GrdTable.Fields(nCol).Value & ""
        Else
            grid2.TextMatrix(grid2.Rows - 1, nCol) = GrdTable.Fields(nCol).Value & ""
            grid1.TextMatrix(grid1.Rows - 1, nCol) = ""
        End If
    Next
    GrdTable.MoveNext
Loop
GrdTable.Close
Set GrdTable = Nothing
End Sub
Private Sub FixCost(cFieldas, cField)
' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,Price from file1_10 WHERE MODEL = " & MyParn(XMODEL.text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(price)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grid1.Cols - 1
        grid1.TextMatrix(2, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If
End Sub


Private Function myloadGrdModel() As Boolean
Dim loctable As New ADODB.Connection
Dim cFieldas As String, cField As String
cString = cString

myDefineGrd

aRet = retFields(cFieldas, cField)
If cField = "" And cFieldas = "" Then Exit Function

FillItem cFieldas, cField
FixCost cFieldas, cField

Fixgrd2
myloadGrdModel = True
End Function


Private Sub Fixgrd2()
With grid1
.ColHidden(0) = True
.RowHidden(1) = True

.TextMatrix(0, 1) = "«··Ê‰/«·„ﬁ«”"
.ColWidth(0) = 400
.ColWidth(1) = 1300
nColWidth = (.Width - 200 - .ColWidth(0) - .ColWidth(1)) / .Cols
If nColWidth < 500 Then nColWidth = 500
If nColWidth > 1200 Then nColWidth = 1200
For nCol = 2 To .Cols - 1
    .ColWidth(nCol) = nColWidth
    .ColAlignment(nCol) = flexAlignCenterCenter
Next
End With
End Sub

Private Sub DataGrdModel()
Dim loctable As New ADODB.Recordset, cString As String, cWhere As String, cOr As String, nTotal As Long, nSaveRow As Long, nSaveCol As Long
nSaveRow = grid1.Row: nSaveCol = grid1.Col

ClearGrdModel

cString = "Select FILE1_11.ITEM,SUM([IN] - [OUT]) as Balance FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM"
cString = cString & turn(cString) & " FILE1_10.MODEL = " & MyParn(XMODEL.text)
cString = cString & " GROUP BY FILE1_11.ITEM"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    For nCol = 2 To grid2.Cols - 1
        nFound = grid2.FindRow(loctable!Item, , nCol, False)
        If nFound <> -1 Then
            grid1.TextMatrix(nFound, nCol) = loctable!balance
        End If
    Next
    loctable.MoveNext
Loop
Set loctable = Nothing
If grid1.Rows > 3 Then
    grid1.AddItem ""
    grid1.Cell(flexcpBackColor, grid1.Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = &HE0E0E0
    grid1.Cols = grid1.Cols + 1
    grid1.Cell(flexcpBackColor, 0, grid1.Cols - 1, grid1.Rows - 1, grid1.Cols - 1) = &HE0E0E0
    grid1.TextMatrix(0, grid1.Cols - 1) = "≈Ã„«·Ì ·Ê‰"
    grid1.TextMatrix(grid1.Rows - 1, 1) = "≈Ã„«·Ì „ﬁ«”"
    
    For nCol = 2 To grid1.Cols - 1
        For nRow = 3 To grid1.Rows - 2
            nTotal = nTotal + Val(grid1.TextMatrix(nRow, nCol))
        Next
        grid1.TextMatrix(grid1.Rows - 1, nCol) = Myvalue(nTotal)
        nTotal = 0
    Next
    For nRow = 3 To grid1.Rows - 1
        For nCol = 2 To grid1.Cols - 2
            nTotal = nTotal + Val(grid1.TextMatrix(nRow, nCol))
        Next
        grid1.TextMatrix(nRow, grid1.Cols - 1) = Myvalue(nTotal)
        nTotal = 0
    Next
    grid1.ColAlignment(grid1.Cols - 1) = flexAlignRightCenter
End If
If nSaveRow <= grid1.Rows - 1 Then grid1.Select nSaveRow, nSaveCol
End Sub
Private Sub ClearGrdModel()
With grid1
.Rows = grid2.Rows
.Cols = grid2.Cols
For nRow = 3 To .Rows - 1
    For nCol = 2 To .Cols - 1
        .TextMatrix(nRow, nCol) = ""
    Next
Next
End With
End Sub


Private Function retFields(ByRef cFieldas As String, ByRef cField As String)
Dim aRet(1) As String
Dim FieldTable As New ADODB.Recordset
'  ⁄—Ì› «·«⁄„œ…
FieldTable.Open "Select SCAL from file1_10 where model = " & MyParn(XMODEL.text) & " GROUP BY SCAL,C_SCAL order by c_scal", con, adOpenStatic, adLockReadOnly
Do Until FieldTable.EOF
    If Not IsNull(FieldTable!scal) Then
        cFieldas = cFieldas & turn(cField, ",") & "[" & FieldTable!scal & "]" & " as " & "[" & FieldTable!scal & "]"
        cField = cField & turn(cField, ",") & "[" & FieldTable!scal & "]"
    End If
    FieldTable.MoveNext
Loop

' ⁄œ„ ÊÃÊœ «⁄„œ…
FieldTable.Close
Set FieldTable = Nothing
End Function
Sub myProc()
    XMODEL.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    CardTable.Find "Model = " & MyParn(XMODEL.text), , adSearchForward, adBookmarkFirst
    If Not CardTable.EOF Then
        xModelFact.text = DelZero(CardTable!modelfact)
        xFact.BoundText = CardTable!Fact
        xMosm.BoundText = CardTable!MOSM
        xCode.text = CardTable!CODE
        xCodeDesc.Caption = GetDesca("select desca from file4_10 where code = " & MyParn(xCode.text), con)
        xDesca.text = CardTable!DESCA
        xGroup.BoundText = CardTable!Group
        
        myloadGrdModel
        DataGrdModel
    End If
    Unload Search3
End Sub
Private Sub XBARCODE_Click()
    If Val(xbarcode.Caption) > 100 Then
        If SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Row, 1) = "" Then
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Row, 1) = Val(xbarcode.Caption)
        End If
        Unload Me
    End If
End Sub
