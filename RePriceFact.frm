VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form RePriceFact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕœÌÀ «”⁄«— „‰ «·„’‰⁄"
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
   Begin VB.Frame Frame4 
      Caption         =   " ⁄œÌ·"
      Height          =   915
      Left            =   4185
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   405
      Width           =   1545
      Begin VB.CheckBox PRICE1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "”⁄— „’‰⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   1365
      End
      Begin VB.CheckBox PRICE2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "”⁄— „” Â·ﬂ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   540
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8880
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1350
      Width           =   16440
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   8475
         Left            =   90
         TabIndex        =   14
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
            Size            =   9.75
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
      TabIndex        =   10
      Top             =   0
      Width           =   4110
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   1462
         Picture         =   "RePriceFact.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   225
         Width           =   1320
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   2835
         Picture         =   "RePriceFact.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "⁄—÷"
         Top             =   225
         Width           =   1230
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   90
         Picture         =   "RePriceFact.frx":491C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
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
      Begin VB.CheckBox XISALL2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   " ÕœÌœ „Õœœ ›ﬁÿ"
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   945
         Width           =   1680
      End
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
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   945
         Width           =   1410
      End
      Begin VB.CheckBox xisprice 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "«”⁄«— „Œ ·›… ›ﬁÿ"
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
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   945
         Width           =   1860
      End
      Begin VB.CommandButton cmd_barcode_fix 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Õ›Ÿ  «·«”⁄«— «·ÃœÌœ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   450
         Width           =   2130
      End
      Begin VB.CheckBox xisbal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "«’‰«› ·Â« —’Ìœ ›ﬁÿ"
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   135
         Value           =   1  'Checked
         Width           =   2175
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
         Left            =   3105
         MaxLength       =   15
         TabIndex        =   7
         Top             =   540
         Width           =   1545
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
         TabIndex        =   6
         Top             =   945
         Width           =   3390
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   5895
         TabIndex        =   2
         Top             =   180
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
         Left            =   2340
         TabIndex        =   3
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
         TabIndex        =   18
         Top             =   585
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
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
         TabIndex        =   19
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   5
         Top             =   255
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
         TabIndex        =   4
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
      TabIndex        =   20
      Top             =   900
      Width           =   2235
      _ExtentX        =   3942
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
      Picture         =   "RePriceFact.frx":6D88
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin ComctlLib.ProgressBar prog1 
      Height          =   375
      Left            =   90
      TabIndex        =   21
      Top             =   900
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "RePriceFact"
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
With grid1
    nRec = 0
    For nRow = 2 To .Rows - 1
        If IIf(Val(.ValueMatrix(nRow, 13) & ""), 1, 0) Then
            If Val(.TextMatrix(nRow, 8) & "") <> Val(.TextMatrix(nRow, 12) & "") Or Val(.TextMatrix(nRow, 9) & "") <> Val(.TextMatrix(nRow, 11) & "") Then
            If Val(.TextMatrix(nRow, 10)) > 0 And Val(.TextMatrix(nRow, 11)) > 0 Then
                If PRICE1.Value <> 0 And PRICE2.Value <> 0 Then cStrUpdate = cStrUpdate & " UPDATE FILE1_10 SET COST = " & Val(.TextMatrix(nRow, 12)) & " , PRICE = " & Val(.TextMatrix(nRow, 11)) & " WHERE ITEM = " & Val(.TextMatrix(nRow, 0)) & " ; "
                If PRICE1.Value <> 0 And PRICE2.Value = 0 Then cStrUpdate = cStrUpdate & " UPDATE FILE1_10 SET COST = " & Val(.TextMatrix(nRow, 12)) & " WHERE ITEM = " & Val(.TextMatrix(nRow, 0)) & " ; "
                If PRICE2.Value <> 0 And PRICE1.Value = 0 Then cStrUpdate = cStrUpdate & " UPDATE FILE1_10 SET PRICE = " & Val(.TextMatrix(nRow, 11)) & " WHERE ITEM = " & Val(.TextMatrix(nRow, 0)) & " ; "
                nCount = nCount + 1
                If nCount = 100 Then
                    nCount = 0
                    If cStrUpdate <> "" Then con.Execute cStrUpdate
                    cStrUpdate = ""
                End If
                nRec = nRec + 1
                Me.Caption = .TextMatrix(nRow, 0)
                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbGreen
            End If
            End If
        End If
    Next nRow
    If cStrUpdate <> "" Then con.Execute cStrUpdate
    Inform " „  ÕœÌÀ ⁄œœ " & nRec & " »«—ﬂÊœ "
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
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
    openConFACT2 confact2
    openConFACT3 confact3
    
    data4.ConnectionString = strCon
    data4.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC "
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    data1.ConnectionString = strCon
    data1.RecordSource = "Select CODE ,descA From FILE1_10SC ORDER BY CODE "
    Set XSEC.RowSource = data1
    XSEC.ListField = "Desca"
    XSEC.BoundColumn = "CODE"
    
    data5.ConnectionString = strCon
    data5.RecordSource = "Select code ,desca From fact ORDER BY code "
    Set xFact.RowSource = data5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    xFact.BoundText = "001"
    
    xFact.BoundText = sM_Fact
    
    Set grid1.DataSource = DATA6
    
    grid1.Rows = 2
    grid1.Cols = 9
    
    
    Fixgrd
    If xFact.BoundText <> "" Then cmdGo_Click
End Sub
Private Sub myload()
    Dim cWhere As String, cString As String
    If xFact.BoundText = "001" Or xFact.BoundText = "049" Or xFact.BoundText = "095" Then
    xModelFact.text = DelZero(xModelFact.text)
    cString = " SELECT  FILE1_10.ITEM , FILE1_10.BARCODE2 , FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca AS DESCA,  FILE1_10.SCAL , FILE1_10.COLOR , SUM(FILE1_11.[IN] - FILE1_11.OUT) AS BAL , FILE1_10.COST , FILE1_10.PRICE " & _
            " FROM    FILE1_10 LEFT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM  where BARCODE2 is not null "
    If xMosm.BoundText <> "" Then cString = cString & " and FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XSEC.BoundText <> "" Then cString = cString & " and  FILE1_10.[SECTION] = " & Val(XSEC.BoundText)
    If xModelFact.text <> "" Then cString = cString & " and  FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xFact.BoundText <> "" Then cString = cString & " and  FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xDesca.text <> "" Then cString = cString & " and " & MyParnAnd(xDesca.text, "file1_10.desca")
    cString = cString & " GROUP BY FILE1_10.ITEM   , BARCODE2 , FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca, FILE1_10.COLOR , FILE1_10.SCAL , FILE1_10.PRICE , FILE1_10.COST "
'    If xisbal.Value <> 0 Then cString = cString & " having SUM(FILE1_11.[IN] - FILE1_11.OUT) > 0 "

    Set DATA6.Recordset = myRecordSet(cString, con)
        Fixgrd
    Else
        MsgBox "„’‰⁄ €Ì— ’ÕÌÕ"
    End If
End Sub
Sub Fixgrd()
    Dim nPrice As Double
    Dim nCost As Double
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 2
    .WordWrap = True
    .Cols = 14
    .RowHeight(1) = 1000
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "»«—ﬂÊœ"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "„ÊœÌ·"
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = " «·„ﬁ«”"
    .TextMatrix(0, 6) = " «··Ê‰"
    .TextMatrix(0, 7) = "«·—’Ìœ"
    
    .TextMatrix(1, 0) = "ﬂÊœ"
    .TextMatrix(1, 1) = "»«—ﬂÊœ"
    .TextMatrix(1, 2) = "„Ê”„"
    .TextMatrix(1, 3) = "„ÊœÌ·"
    .TextMatrix(1, 4) = "«·’‰›"
    .TextMatrix(1, 5) = " «·„ﬁ«”"
    .TextMatrix(1, 6) = " «··Ê‰"
    .TextMatrix(1, 7) = "«·—’Ìœ"
    
    .TextMatrix(0, 8) = " «”⁄«— ﬁ»·"
    .TextMatrix(0, 9) = " «”⁄«— ﬁ»·"
    .TextMatrix(0, 10) = " «”⁄«— »⁄œ"
    .TextMatrix(0, 11) = " «”⁄«— »⁄œ"
    .TextMatrix(0, 12) = " «”⁄«— »⁄œ"
    .TextMatrix(0, 13) = "≈Œ Ì«—"
    
    .TextMatrix(1, 8) = " ”⁄— «·„’‰⁄"
    .TextMatrix(1, 9) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(1, 10) = " ”⁄— «·„’‰⁄"
    .TextMatrix(1, 11) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(1, 12) = "”⁄— «·„’‰⁄ ··ﬁÿ⁄…"
    .TextMatrix(1, 13) = "≈Œ Ì«—"
    
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    .MergeCol(7) = True
    .MergeCol(13) = True
    
    
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1600
    .ColWidth(2) = 800
    .ColWidth(3) = 1200
    .ColWidth(4) = 3000
    
    .ColWidth(5) = 900
    .ColWidth(6) = 900
    .ColWidth(7) = 800
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    .ColWidth(10) = 800
    .ColWidth(11) = 800
    
    .ColWidth(12) = 800
    .ColWidth(13) = 800
    .ColDataType(13) = flexDTBoolean
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString
    If .Rows > 2 Then .Row = 2
    For nRow = 2 To .Rows - 1
        cBarCode = .TextMatrix(nRow, 1)
        If cBarCode <> "" Then
            If xFact.BoundText = "001" Then
                aRet = aGetDesca("SELECT PRICE2 , PRICE , ISDOZ FROM FILE1_10 WHERE BARCODE2 = " & MyParn(cBarCode), confact)
                If UBound(aRet) > 0 Then
                    nPrice = aRet(1) ' Val(GetDesca("SELECT PRICE2 FROM FILE1_10 WHERE BARCODE2 = " & MyParn(cBarCode), confact) & "")
                    nCost = aRet(2) ' Val(GetDesca("SELECT PRICE FROM FILE1_10 WHERE BARCODE2 = " & MyParn(cBarCode), confact) & "")
                    .TextMatrix(nRow, 10) = nCost
                    .TextMatrix(nRow, 11) = nPrice
'                   If GetBoolean("SELECT ISDOZ FROM FILE1_10 WHERE BARCODE2 = " & MyParn(cBarCode), confact) Then
                    If IIf(aRet(3), 1, 0) Then
                        .TextMatrix(nRow, 12) = Round(nCost / 12, 2)
                    Else
                        .TextMatrix(nRow, 12) = nCost
                    End If
                End If
            End If
            If xFact.BoundText = "095" Then
                nPrice = Val(GetDesca("SELECT PRICE2 FROM FILE1_10 WHERE BARCODE2 = " & MyParn(cBarCode), confact3) & "")
                nCost = Val(GetDesca("SELECT PRICE FROM FILE1_10 WHERE BARCODE2 = " & MyParn(cBarCode), confact3) & "")
                .TextMatrix(nRow, 10) = nCost
                .TextMatrix(nRow, 11) = nPrice
                .TextMatrix(nRow, 12) = Round(nCost / 12, 2)
            End If
            If xFact.BoundText = "049" Then
                nPrice = Val(GetDesca("SELECT PRICE2 FROM FILE1_10 WHERE BARCODE2 = " & MyParn(cBarCode), confact) & "")
                nCost = Val(GetDesca("SELECT PRICE FROM FILE1_10 WHERE BARCODE2 = " & MyParn(cBarCode), confact) & "")
                .TextMatrix(nRow, 10) = nCost
                .TextMatrix(nRow, 11) = nPrice
                .TextMatrix(nRow, 12) = nCost
            End If
        End If
        .RowHidden(nRow) = False
        If Round(Val(.TextMatrix(nRow, 8) & ""), 0) <> Round(Val(.TextMatrix(nRow, 12) & ""), 0) Or Round(Val(.TextMatrix(nRow, 9) & ""), 0) <> Round(Val(.TextMatrix(nRow, 11) & ""), 0) Then
            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbGreen
        Else
            If xisprice.Value <> 0 Then .RowHidden(nRow) = True
        End If
    Next nRow
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = 4
    If .Rows > 2 Then .Cell(flexcpAlignment, 2, 0, .Rows - 1, .Cols - 1) = 7
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
    If .Col >= 10 Then
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
    For nRow = 2 To .Rows - 1
        .TextMatrix(nRow, 13) = IIf(XISALL.Value, 1, 0)
    Next nRow
End With
End Sub
Private Sub XISALL2_Click()
With grid1
    For nRow = 2 To .Rows - 1
        If .IsSelected(nRow) Then
            .TextMatrix(nRow, 13) = IIf(XISALL2.Value, 1, 0)
        End If
    Next nRow
End With
End Sub
