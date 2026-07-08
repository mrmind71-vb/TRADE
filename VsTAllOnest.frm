VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form VsTAllOnest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈Ã„«·Ï —’Ìœ & „»Ì⁄«  «·«„«‰« "
   ClientHeight    =   10515
   ClientLeft      =   0
   ClientTop       =   375
   ClientWidth     =   10770
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
   ScaleHeight     =   10515
   ScaleWidth      =   10770
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Height          =   510
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2430
      Width           =   3840
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬂ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   180
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Caption         =   " ÊﬂÌ·«  ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Caption         =   "›—Ê⁄ ›ﬁÿ"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   180
         Value           =   -1  'True
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "’Ê— «·„ÊœÌ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   3570
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   2085
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   3678
         BackColor       =   16777215
         BorderStyle     =   4
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         AutoZoomType    =   3
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2655
      Left            =   8940
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   11235
      Begin VB.TextBox XMOSM2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   7200
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox xmosm22 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1035
         Width           =   915
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "»œÊ‰ «Ã„«·Ï „ÊœÌ·« "
         Height          =   420
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2025
         Width           =   2220
      End
      Begin VB.Frame Frame4 
         Height          =   1185
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1350
         Width           =   3795
         Begin VB.CommandButton CMD_PRINT 
            Height          =   465
            Left            =   1260
            Picture         =   "VsTAllOnest.frx":0000
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   135
            Width           =   1185
         End
         Begin VB.CommandButton cmdExit 
            Height          =   465
            Left            =   45
            Picture         =   "VsTAllOnest.frx":242A
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   135
            Width           =   1185
         End
         Begin VB.CommandButton cmdGo 
            Height          =   465
            Left            =   2475
            Picture         =   "VsTAllOnest.frx":4896
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "⁄—÷"
            Top             =   180
            Width           =   1275
         End
         Begin Threed.SSCommand cmd_excel 
            Height          =   465
            Left            =   45
            TabIndex        =   20
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
            Picture         =   "VsTAllOnest.frx":6D88
            Caption         =   " ÕÊÌ· «ﬂ”Ì· "
            Alignment       =   4
            PictureAlignment=   1
         End
         Begin ComctlLib.ProgressBar prog1 
            Height          =   465
            Left            =   2475
            TabIndex        =   21
            Top             =   630
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   820
            _Version        =   327682
            Appearance      =   1
         End
      End
      Begin VB.TextBox xdate2 
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
         Left            =   5985
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
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
         Left            =   7935
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1500
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   225
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   6000
         TabIndex        =   5
         Top             =   596
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   1350
         TabIndex        =   12
         Top             =   1035
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   1350
         TabIndex        =   14
         Top             =   596
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "‰Ê⁄ «·„Ê”„ :"
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
         Index           =   9
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1530
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "„Ê”„ œ«Œ·Ï"
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
         Index           =   8
         Left            =   8190
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Width           =   1140
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
         TabIndex        =   15
         Top             =   585
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄ :"
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1035
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
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
         TabIndex        =   8
         Top             =   270
         Width           =   765
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
         Index           =   6
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label5 
         Caption         =   "«·„Ê—œ :"
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
         TabIndex        =   6
         Top             =   585
         Width           =   1545
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10185
      Width           =   10770
      _ExtentX        =   18997
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTAllOnest.frx":9583
      Height          =   6525
      Left            =   45
      TabIndex        =   1
      Top             =   2925
      Width           =   20130
      _cx             =   35507
      _cy             =   11509
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
      BackColorSel    =   16776960
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
      Rows            =   2
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
   Begin MSAdodcLib.Adodc DATA2 
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
   Begin MSAdodcLib.Adodc data1 
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
   Begin MSAdodcLib.Adodc data5 
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
   Begin MSAdodcLib.Adodc DATA6 
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
      Caption         =   "data7"
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
   Begin MSAdodcLib.Adodc data8 
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
   Begin MSAdodcLib.Adodc data9 
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
   Begin MSAdodcLib.Adodc data10 
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
Attribute VB_Name = "VsTAllOnest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = Me.Caption & " " & xStore.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY") & " Õ Ï  «—ÌŒ " & Format(xDate2.text, "DD-MM-YYYY")
    cHead3 = ""
    If xSupp.BoundText <> "" Then cHead3 = cHead3 & " „Ê—œ " & xSupp.text & "  " & xGrCust.text
    If XSECTION.text <> "" Then cHead3 = cHead3 & " ﬁ”„ " & XSECTION.text
    If xMosm.BoundText <> "" Then cHead3 = cHead3 & " „Ê”„ " & xMosm.text
    PrintGrdNew.DOPRINT grid1, 1, , cHead1, cHead2, cHead3, , False, 8
    PrintGrdNew.Show 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
'    On Error GoTo myerror
    myload
    Exit Sub
myerror:
    MsgBox Err.Description
End Sub
Private Sub Form_Load()
    openCon con
    xmosm22.AddItem " "
    xmosm22.AddItem "S"
    xmosm22.AddItem "W"
    xmosm22.AddItem "M"
    
    data8.ConnectionString = strCon
    data8.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
    Set XSECTION.RowSource = data8
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
    DATA7.ConnectionString = strCon
    DATA7.RecordSource = "Select Code,DescA From STORE_BR ORDER BY code "
    Set xStore.RowSource = DATA7
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    If cBranch <> "00" Then
        xStore.BoundText = cBranchStore
        xStore.Enabled = False
    End If
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File4_10 order by Desca"
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    DATA4.ConnectionString = strCon
    DATA4.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC "
    Set xMosm.RowSource = DATA4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    
    grid1.Rows = 2
    grid1.Cols = 10
    FIXGRID
End Sub
Private Sub myload()
Dim I As Double
Dim cString  As String, cStr2 As String
With grid1


cWhere = ""
cWhere2 = ""

If xStore.BoundText <> "" Then cWhere = cWhere & " AND STORE = " & MyParn(xStore.BoundText)
If IsDate(xDate1.text) Then cWhere = cWhere & " AND [DATE] >= " & DateSq(xDate1.text)
If IsDate(xDate2.text) Then cWhere = cWhere & " AND [DATE] <= " & DateSq(xDate2.text)

    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M') "
    End If
If xmosm22.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(xmosm22.text)

If xStore.BoundText <> "" Then cWhere2 = cWhere2 & " AND STORE = " & MyParn(xStore.BoundText)
If IsDate(xDate2.text) Then cWhere2 = cWhere2 & " AND [DATE] <= " & DateSq(xDate2.text)


'If xall(0).Value <> 0 Then cwhere = cwhere & " AND STORE < '060'  "
'If xall(1).Value <> 0 Then cwhere = cwhere & " AND STORE > '60'  "

If xall(0).Value <> 0 Then
    cF11 = ", (SELECT SUM(QUANT)                                FROM QFILE6_20_1 WHERE QFILE6_20_1.ITEM = FILE1_10.ITEM " & cWhere & " )"
    cf12 = ", (SELECT SUM(QFILE6_20_1.QUANT * QFILE6_20_1.price_c ) FROM QFILE6_20_1 WHERE QFILE6_20_1.ITEM = FILE1_10.ITEM " & cWhere & " )"
    cf13 = ", (SELECT sum((QFILE6_20_1.QUANT * QFILE6_20_1.price_c )-QFILE6_20_1.total ) + SUM(QFILE6_20_1.TOTAL * ((QFILE6_20_1.RATE)/100)) FROM QFILE6_20_1 WHERE QFILE6_20_1.ITEM = FILE1_10.ITEM and ( QFILE6_20_1.username_disc IS NULL OR QFILE6_20_1.username_disc = '' ) " & cWhere & " )"
    cf14 = ", (SELECT Sum (QFILE6_20_1.TOTAL * (((100 - QFILE6_20_1.Rate) / 100))) FROM QFILE6_20_1 WHERE QFILE6_20_1.ITEM = FILE1_10.ITEM  " & cWhere & " )"
    cf15 = ", (SELECT Sum ([IN] - [OUT] ) FROM FILE1_11 WHERE FILE1_11.ITEM = FILE1_10.ITEM " & cWhere2 & " )"
    
    cString = "SELECT  FILE1_10.MODEL ,  file1_10.mosm , file1_10.MODELFACT0, file1_10.desca , MIN(FILE1_10.PRICE)   ,FILE1_10.COLOR , FILE1_10.SCAL , file1_10.ITEM, " & _
            " SUM(Q_ONEST.[PURCH]) , SUM(Q_ONEST.[RET]) , SUM(Q_ONEST.[purch]-Q_ONEST.[RET]) " & cF11 & cf12 & cf13 & cf14 & cf15
    cString = cString & " FROM FILE1_10 INNER JOIN Q_ONEST ON FILE1_10.ITEM = Q_ONEST.ITEM WHERE FILE1_10.ITEM IS NOT NULL "
    
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND Q_ONEST.CODE = " & MyParn(xSupp.BoundText)
    If xStore.BoundText <> "" Then cStr2 = cStr2 & " AND STORE = " & MyParn(xStore.BoundText)
    If XSECTION.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & MyParn(XSECTION.BoundText)
    If IsDate(xDate1.text) Then cStr2 = cStr2 & " AND [DATE] >= " & DateSq(xDate1.text)
    If IsDate(xDate2.text) Then cStr2 = cStr2 & " AND [DATE] <= " & DateSq(xDate2.text)
    cString = cString & cStr2 & " GROUP BY FILE1_10.MODEL , file1_10.ITEM  , file1_10.mosm , file1_10.MODELFACT0, file1_10.desca , FILE1_10.COLOR , FILE1_10.SCAL  , FILE1_10.C_SCAL ORDER BY file1_10.model, file1_10.MODELFACT0, FILE1_10.COLOR , FILE1_10.C_SCAL"
Else
    cF11 = ", (SELECT SUM(QUANT)                                FROM QFILE6_20_2 WHERE QFILE6_20_2.ITEM = FILE1_10.ITEM " & cWhere & " )"
    cf12 = ", (SELECT SUM(QFILE6_20_2.QUANT * QFILE6_20_2.price_c ) FROM QFILE6_20_2 WHERE QFILE6_20_2.ITEM = FILE1_10.ITEM " & cWhere & " )"
    cf13 = ", (SELECT sum((QFILE6_20_2.QUANT * QFILE6_20_2.price_c )-QFILE6_20_2.total ) + SUM(QFILE6_20_2.TOTAL * ((QFILE6_20_2.RATE)/100)) FROM QFILE6_20_2 WHERE QFILE6_20_2.ITEM = FILE1_10.ITEM and ( QFILE6_20_2.username_disc IS NULL OR QFILE6_20_2.username_disc = '' ) " & cWhere & " )"
    cf14 = ", (SELECT Sum (QFILE6_20_2.TOTAL * (((100 - QFILE6_20_2.Rate) / 100))) FROM QFILE6_20_2 WHERE QFILE6_20_2.ITEM = FILE1_10.ITEM  " & cWhere & " )"
    cf15 = ", (SELECT Sum ([IN] - [OUT] ) FROM FR1_11 WHERE FR1_11.ITEM = FILE1_10.ITEM " & cWhere2 & " )"

    cString = "SELECT  FILE1_10.MODEL ,  file1_10.mosm , file1_10.MODELFACT0, file1_10.desca , MIN(FILE1_10.PRICE)   ,FILE1_10.COLOR , FILE1_10.SCAL , file1_10.ITEM, " & _
            " SUM(Q_ONEST_2.[PURCH]) , SUM(Q_ONEST_2.[RET]) , SUM(Q_ONEST_2.[purch]-Q_ONEST_2.[RET]) " & cF11 & cf12 & cf13 & cf14 & cf15
    cString = cString & " FROM FILE1_10 INNER JOIN Q_ONEST_2 ON FILE1_10.ITEM = Q_ONEST_2.ITEM WHERE FILE1_10.ITEM IS NOT NULL "
    
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND Q_ONEST_2.CODE = " & MyParn(xSupp.BoundText)
    If xStore.BoundText <> "" Then cStr2 = cStr2 & " AND STORE = " & MyParn(xStore.BoundText)
    If XSECTION.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & MyParn(XSECTION.BoundText)
    If IsDate(xDate1.text) Then cStr2 = cStr2 & " AND [DATE] >= " & DateSq(xDate1.text)
    If IsDate(xDate2.text) Then cStr2 = cStr2 & " AND [DATE] <= " & DateSq(xDate2.text)
    cString = cString & cStr2 & " GROUP BY FILE1_10.MODEL , file1_10.ITEM  , file1_10.mosm , file1_10.MODELFACT0, file1_10.desca , FILE1_10.COLOR , FILE1_10.SCAL  , FILE1_10.C_SCAL ORDER BY file1_10.model, file1_10.MODELFACT0, FILE1_10.COLOR , FILE1_10.C_SCAL"
End If
    DATA6.RecordSource = cString
    DATA6.Refresh
End With
FIXGRID
End Sub
Sub FIXGRID()
With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
    .Cols = 21
    
    .TextMatrix(0, 0) = "„ÊœÌ·"
    .TextMatrix(0, 1) = "„Ê”„"
    .TextMatrix(0, 2) = "—ﬁ„"
    .TextMatrix(0, 3) = "«·„ÊœÌ· "
    .TextMatrix(0, 4) = "”⁄— „” Â·ﬂ"
    .TextMatrix(0, 5) = "«··Ê‰"
    .TextMatrix(0, 6) = "«·„ﬁ«”"
    .TextMatrix(0, 7) = "ﬂÊœ"
    
    .TextMatrix(0, 8) = "„‘ —Ì«  «„«‰« "
    .TextMatrix(0, 9) = "„— Ã⁄ «„«‰« "
    .TextMatrix(0, 10) = "’«›Ï „‘ —Ì« "
    
    .TextMatrix(0, 11) = "„»Ì⁄« "
    .TextMatrix(0, 12) = "ﬁÌ„… «·„»Ì⁄«  »”⁄— «·»Ì⁄"
    .TextMatrix(0, 13) = "Œ’„ «·„»Ì⁄« "
    .TextMatrix(0, 14) = "’«›Ï ﬁÌ„… «·„»Ì⁄« "
    
    .TextMatrix(0, 15) = "«·—’Ìœ"
    .TextMatrix(0, 16) = "„»Ì⁄«  «·«„«‰« "
    .TextMatrix(0, 17) = "„»Ì⁄«  «„«‰«  »”⁄— «·»Ì⁄"
    .TextMatrix(0, 18) = "Œ’„ „»Ì⁄«  «„«‰« "
    .TextMatrix(0, 19) = "’«›Ï ﬁÌ„… „»Ì⁄«  «·«„«‰« "
    .TextMatrix(0, 20) = "—’Ìœ «·«„«‰« "
    
    
    .ColHidden(0) = True
    .ColWidth(0) = 0
    .ColWidth(1) = 700
    .ColWidth(2) = 1500
    .ColWidth(3) = 3000
    .ColWidth(4) = 800
    .ColWidth(5) = 1500
    .ColWidth(6) = 1000
    .ColWidth(7) = 0
    
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    .ColWidth(10) = 800
    
    .ColWidth(11) = 800
    .ColWidth(12) = 1000
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .ColWidth(16) = 1000
    .ColWidth(17) = 1000
    .ColWidth(18) = 1000
    .ColWidth(19) = 1000
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For nCol = 8 To .Cols - 1
        .ColFormat(nCol) = "#0"
        .ColDataType(nCol) = flexDTDouble
    Next nCol
    .SubtotalPosition = flexSTBelow
    For I = 8 To .Cols - 1
        If Check1.Value = 0 Then
            .Subtotal flexSTSum, 0, I, "#0", &H8000000F, , True, "  "
        End If
        .Subtotal flexSTSum, -1, I, "#0", &H8000000F, , True, "  "
    Next I
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 10)) >= Val(.TextMatrix(nRow, 15)) Then
'            If Val(.TextMatrix(nRow, 15)) >= 0 Then
                .TextMatrix(nRow, 20) = Val(.TextMatrix(nRow, 15))
'            Else
'                .TextMatrix(nRow, 20) = 0
'            End If
        Else
            .TextMatrix(nRow, 20) = Val(.TextMatrix(nRow, 10))
        End If
        .TextMatrix(nRow, 16) = Val(.TextMatrix(nRow, 10)) - Val(.TextMatrix(nRow, 20))
        nRate = 0
        If Val(.TextMatrix(nRow, 11)) <> 0 Then nRate = .TextMatrix(nRow, 16) / .TextMatrix(nRow, 11)
        If nRate > 0 And nRate < 1 Then
            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = &HFF8080
        End If
        .TextMatrix(nRow, 17) = Round(Val(.TextMatrix(nRow, 12)) * nRate, 0)
        .TextMatrix(nRow, 18) = Round(Val(.TextMatrix(nRow, 13)) * nRate, 0)
        .TextMatrix(nRow, 19) = Round(Val(.TextMatrix(nRow, 14)) * nRate, 0)
    Next nRow
    For I = 8 To .Cols - 1
        If Check1.Value = 0 Then
            .Subtotal flexSTSum, 0, I, "#0", &H8000000F, , True, "  "
        End If
        .Subtotal flexSTSum, -1, I, "#0", &H8000000F, , True, "  "
    Next I
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
End Sub
Private Sub grid1_DblClick()
''    If grid1.Col <= 4 And cBranch = "00" Then
''        ITEMS.sModel = GetDesca("Select Model FROM FILE1_10H WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 1)))
''        ITEMS.Show 1
''    Else
'        ModelMove.xModel.Text = grid1.TextMatrix(grid1.Row, 2)
'        ModelMove.XSTORE.BoundText = XSTORE.BoundText
'        ModelMove.xDate = xDate1.Text
'        ModelMove.Show
''    End If
End Sub
Private Sub grid1_EnterCell()
'    xModelNo.Caption = grid1.TextMatrix(grid1.Row, 0)
'    LoadPhoto grid1.TextMatrix(grid1.Row, 0)
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

Private Sub ImgX1_Click()
If validPhoto(retFile(grid1.TextMatrix(grid1.Row, 0))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(grid1.TextMatrix(grid1.Row, 0))
    oMaxPhoto.cDesca = grid1.TextMatrix(grid1.Row, 4)
    oMaxPhoto.Show 1
End If
End Sub



