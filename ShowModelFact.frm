VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form ShowModelFact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„ÊœÌ·«  „‰ «·„’‰⁄ Ê  —Õ·ÌÂ« ··„Õ·"
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   16515
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
   ScaleWidth      =   16515
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_barcode_fix 
      BackColor       =   &H00FFFFC0&
      Caption         =   "≈÷«›… «·„ÊœÌ·«   ··„Õ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4230
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   675
      Width           =   1455
   End
   Begin VB.CommandButton CMD_MODEL 
      Caption         =   "⁄—÷ „ÊœÌ·«  „”Ã·…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4230
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   90
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   8880
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1350
      Width           =   16440
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   8475
         Left            =   90
         TabIndex        =   13
         Top             =   180
         Width           =   16260
         _cx             =   28681
         _cy             =   14949
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   4110
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   1462
         Picture         =   "ShowModelFact.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   1320
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   2835
         Picture         =   "ShowModelFact.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "⁄—÷"
         Top             =   225
         Width           =   1230
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   90
         Picture         =   "ShowModelFact.frx":491C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   10725
      Begin VB.CheckBox XISALL 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   " ÕœÌœ «·ﬂ·"
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
         TabIndex        =   19
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox xModelFact 
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
         Left            =   2340
         MaxLength       =   15
         TabIndex        =   6
         Top             =   540
         Width           =   2310
      End
      Begin VB.TextBox xDesca 
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
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   945
         Width           =   3390
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   2340
         TabIndex        =   2
         Top             =   180
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSEC 
         Height          =   315
         Left            =   5895
         TabIndex        =   15
         Top             =   585
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   5895
         TabIndex        =   20
         Top             =   180
         Visible         =   0   'False
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xmosm2 
         Height          =   315
         Left            =   2340
         TabIndex        =   23
         Top             =   945
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "≈÷«›… ≈·Ï „Ê”„ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   945
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ”„ : "
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
         Left            =   9495
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   585
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ÊœÌ· :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4740
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   585
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ’‰› :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   4
         Left            =   9360
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   945
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„’‰⁄ :"
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
         Left            =   9465
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   255
         Visible         =   0   'False
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ê”„ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   4725
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1230
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10710
      Width           =   16515
      _ExtentX        =   29131
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
   Begin Threed.SSCommand cmd_excel 
      Height          =   375
      Left            =   1890
      TabIndex        =   17
      Top             =   900
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   661
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
      Picture         =   "ShowModelFact.frx":6D88
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin ComctlLib.ProgressBar prog1 
      Height          =   375
      Left            =   90
      TabIndex        =   18
      Top             =   900
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "ShowModelFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Variant, sM_Fact As String
Dim oSearch As New Search3
Dim con As New ADODB.Connection
Dim confact As New ADODB.Connection
Dim confact2 As New ADODB.Connection
Dim confact3 As New ADODB.Connection
Private Sub cmd_barcode_fix_Click()
Dim nRec As Double
Dim cStrUpdate As String, nCount As Double
'On Error GoTo myerror
If xmosm2.BoundText = "" Then
    MsgBox " «Œ «— „Ê”„ «·√÷«›… "
    Exit Sub
End If
With grid1
    nRec = 0
    For nRow = 1 To .Rows - 1
        If .TextMatrix(nRow, 7) = "" And Abs(Val(.TextMatrix(nRow, 8))) = 1 Then
'            Me.Caption = ItemMosmTable!MODEL
'            Me.Caption = " Model : " & ItemMosmTable!MODEL
            myreplace_ItemFct .TextMatrix(nRow, 1), xmosm2.BoundText, "001"
        End If
    Next nRow
    MsgBox "  „ «÷«›… «·„ÊœÌ·«  «·„Õœœ…  "
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me

End Sub

Private Sub Cmd_Model_Click()
    With grid1
    prog1.Value = 0
    prog1.Max = .Rows - 1
    prog1.Min = 0
    prog1.Visible = True
    For nRow = 1 To .Rows - 1
        prog1.Value = nRow
        Set loctable = ItemFind_BARCODE2(.TextMatrix(nRow, 3), con)
        If Not (loctable.EOF And loctable.BOF) Then
            .TextMatrix(nRow, 6) = loctable!MOSM & ""
            .TextMatrix(nRow, 7) = loctable!modelfact0 & ""
        End If
    Next nRow
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
'    If xMosm.BoundText <> "" Then
        myload
'    End If
End Sub
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ »√”⁄«— «·„ÊœÌ·«  "
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 1, -2, cHead1, , , False, False, 9
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()
    openCon con
    
    openConFACT confact
'    openConFACT2 confact2
'    openConFACT3 confact3
    
    Set data4.Recordset = myRecordSet("SELECT * FROM mosm order by date desc", confact)
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "mosm"
    
    Set DATA2.Recordset = myRecordSet("SELECT * FROM mosm order by date desc", con)
    Set xmosm2.RowSource = DATA2
    xmosm2.ListField = "Desca"
    xmosm2.BoundColumn = "mosm"
    
    Set data1.Recordset = myRecordSet("SELECT * FROM SUB_SEC ORDER BY DESCA ", confact)
    Set XSEC.RowSource = data1
    XSEC.ListField = "Desca"
    XSEC.BoundColumn = "CODE"
    
    
    
    Set grid1.DataSource = data6
    
    grid1.Rows = 2
    grid1.Cols = 9
    
    
    Fixgrd
    If xFact.BoundText <> "" Then cmdGo_Click
End Sub
Private Sub myload()
    Dim cFiledShop As String
    Dim cWhere As String, cString As String
    '                       0               1                   2               3               4                   5                   6                   7               8
    
    cString = " SELECT  FILE1_10.MOSM , FILE1_10.MODEL , FILE1_10.DESCA , MIN(FILE1_10.BARCODE2 ) , MIN(FILE1_10.COST) , MIN(FILE1_10.PRICE ) " & _
            " FROM    FILE1_10 WHERE FILE1_10.MODEL IS NOT NULL "
    If xMosm.BoundText <> "" Then cString = cString & " and FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XSEC.BoundText <> "" Then cString = cString & " and  FILE1_10.[SEC] = " & MyParn(XSEC.BoundText)
    If xModelFact.text <> "" Then cString = cString & " and  FILE1_10.MODEL = " & MyParn(xModelFact.text)
'    If xFact.BoundText <> "" Then cString = cString & " and  FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XDESCA.text <> "" Then cString = cString & " and " & MyParnAnd(XDESCA.text, "file1_10.desca")
    cString = cString & " GROUP BY FILE1_10.MOSM , FILE1_10.MODEL , FILE1_10.DESCA  ORDER BY FILE1_10.MOSM , FILE1_10.MODEL "
    Set data6.Recordset = myRecordSet(cString, confact)
    Fixgrd
End Sub
Sub Fixgrd()
    Dim loctable As ADODB.Recordset
    Dim nPrice As Double
    Dim nCost As Double
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .Cols = 9
    
    .TextMatrix(0, 0) = "„Ê”„ «·„’‰⁄"
    .TextMatrix(0, 1) = "—ﬁ„ «·„’‰⁄"
    .TextMatrix(0, 2) = "’‰› «·„’‰⁄"
    .TextMatrix(0, 3) = "»«—ﬂÊœ „’‰⁄"
    
    .TextMatrix(0, 4) = "”⁄— „’‰€"
    .TextMatrix(0, 5) = "”⁄— „” Â·ﬂ"
    
    .TextMatrix(0, 6) = "„Ê”„ «·„Õ·"
    .TextMatrix(0, 7) = "„ÊœÌ· «·„Õ·"
    
    .TextMatrix(0, 8) = "≈÷«›… «·»Ì«‰« "
    
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 4000
    .ColWidth(3) = 1000
    .ColHidden(3) = True
    
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    
    .ColWidth(6) = 1000
    .ColWidth(7) = 1500
    
    .ColWidth(8) = 1000
    .ColDataType(8) = flexDTBoolean
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    If .Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_DblClick()
If ValidInt(grid1.TextMatrix(grid1.Row, 0)) And Not IsEmpty(myForm) Then
    myForm.myProc
End If
End Sub
Sub myProc()
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
End Sub
Private Sub grid1_EnterCell()
With grid1
    If .Col = 8 Then
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
Private Sub XISALL_Click()
With grid1
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 8) = IIf(XISALL.Value, 1, 0)
    Next nRow
End With
End Sub
Private Sub myreplace_ItemFct(pModel, pMosm, pFact)
Dim pItemTable As New ADODB.Recordset
pItemTable.Open "select * from file1_10 where model = " & MyParn(pModel) & " order by item ", confact, adOpenStatic, adLockReadOnly, adCmdText
With pItemTable
Dim aInsert(27, 1)
Dim sitem As String, aRet As Variant
Dim cBarCode As String
cModel = pFact & pMosm & RetZero(pModel, 10)
cModelNo = pFact & pMosm & RetZero(pModel, 10)
'If GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(cModel), con) <> "" Then Exit Sub
Do While Not pItemTable.EOF
    sitem = Val(GetDesca("SELECT MAX(ITEM) FROM FILE1_10", con)) + 1
        If Val(sitem) < 100 Then sitem = 101
    
        aInsert(0, 0) = "Model"
        aInsert(0, 1) = addstring(cModel)
    
        aInsert(1, 0) = "ModelNO"
        aInsert(1, 1) = addstring(cModelNo)
    
        aInsert(2, 0) = "Desca"
        aInsert(2, 1) = addstring(!DESCA)
    
        aInsert(3, 0) = "FACT"
        aInsert(3, 1) = addstring(pFact)
    
        aInsert(4, 0) = "MOSM"
        aInsert(4, 1) = addstring(pMosm)
    
        aInsert(5, 0) = "MODELFACT"
        aInsert(5, 1) = addstring(RetZero(pModel, 10))
    
        aInsert(6, 0) = "MODELFACT0"
        aInsert(6, 1) = addstring(DelZero(pModel))
    
        aInsert(7, 0) = "SUPP"
        aInsert(7, 1) = addstring("")
    
        aInsert(8, 0) = "[section]"
        aInsert(8, 1) = addvalue(2)
    
        aInsert(9, 0) = "Code"
        aInsert(9, 1) = addstring(pFact)
    
        aInsert(10, 0) = "[group]"
        aInsert(10, 1) = addstring("01")
    
    
    
    
        aInsert(11, 0) = "[grmodel]"
        aInsert(11, 1) = addstring(pItemTable!GRMODEL)
    
        aInsert(12, 0) = "[grdesca]"
        aInsert(12, 1) = addstring(pItemTable!GrDesca)
        
        aInsert(13, 0) = "item"
        aInsert(13, 1) = addvalue(sitem)
    
        aInsert(14, 0) = "Cost"
        If !ISDOZ Then
            aInsert(14, 1) = Val(!price / 12 & "")
        Else
            aInsert(14, 1) = Val(!price & "")
        End If
        
        aInsert(15, 0) = "Price"
        aInsert(15, 1) = Val(!PRICE2 & "")
    
        aInsert(16, 0) = "Cost2"
        If !ISDOZ Then
            aInsert(16, 1) = Val(!price / 12 & "")
        Else
            aInsert(16, 1) = Val(!price & "")
        End If
    
        aInsert(17, 0) = "Scal"
        aInsert(17, 1) = addstring(!scal)
    
        aInsert(18, 0) = "C_Scal"
        aInsert(18, 1) = addvalue(!C_SCAL)
    
        aInsert(19, 0) = "CODE_Scal"
        aInsert(19, 1) = addvalue(!C_SCAL)
    
        aInsert(20, 0) = "Color"
        aInsert(20, 1) = addstring(!color)
    
        aInsert(21, 0) = "C_Color"
        aInsert(21, 1) = 1
    
        aInsert(22, 0) = "Code_Color"
        aInsert(22, 1) = 1
    
        aInsert(23, 0) = "barcode"
        aInsert(23, 1) = addstring(!Item)
        
        aInsert(24, 0) = "barcode2"
        aInsert(24, 1) = addstring(!BARCODE2)
        
        aInsert(25, 0) = "barcode13"
        aInsert(25, 1) = addstring(!BARCODE13)
        
        aInsert(26, 0) = "barcode_gs1"
        aInsert(26, 1) = addstring(!BARCODE_GS1)
        
        aInsert(27, 0) = "Color_E"
        aInsert(27, 1) = addstring(!color)
        
        con.Execute CreateInsert(aInsert, "File1_10"), nRec
        If nRec = 1 Then Me.Caption = " Add " & pModel & " " & !scal
    .MoveNext
Loop
con.Execute " UPDATE FILE1_10 SET ITEM = ITEM WHERE MODEL =  " & MyParn(cModel), nRec

End With
End Sub


