VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsTSalesMonth_YEAR 
   Caption         =   "„ﬁ«—‰… „»Ì⁄«  «·‘ÂÊ— ·”‰Ê« "
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   11280
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
   ScaleHeight     =   10365
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   510
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1440
      Width           =   3795
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
         TabIndex        =   26
         Top             =   180
         Value           =   -1  'True
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   180
         Width           =   960
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   0
      Width           =   3840
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "VsTSalesMonth_YEAR.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "VsTSalesMonth_YEAR.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2475
         Picture         =   "VsTSalesMonth_YEAR.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   21
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
         Picture         =   "VsTSalesMonth_YEAR.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   240
         Left            =   45
         TabIndex        =   22
         Top             =   1125
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   423
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   11145
      Begin VB.TextBox XYEAR2 
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
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox XYEAR1 
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
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox XMOSM2 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Left            =   90
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   135
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   585
         Width           =   915
      End
      Begin VB.Frame Frame3 
         Height          =   510
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1260
         Width           =   4695
         Begin VB.OptionButton SEC 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ﬂ·"
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
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   180
            Value           =   -1  'True
            Width           =   915
         End
         Begin VB.OptionButton SEC 
            Alignment       =   1  'Right Justify
            Caption         =   "„Ê”„"
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
            Left            =   3555
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   180
            Width           =   915
         End
         Begin VB.OptionButton SEC 
            Alignment       =   1  'Right Justify
            Caption         =   "«·‰Ê⁄"
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
            Index           =   1
            Left            =   2610
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   180
            Width           =   915
         End
         Begin VB.OptionButton SEC 
            Alignment       =   1  'Right Justify
            Caption         =   "«·›∆… «·⁄„—Ì…"
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
            Left            =   1305
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   180
            Width           =   1320
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
         TabIndex        =   13
         Top             =   225
         Width           =   1230
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
         Left            =   8340
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   6000
         TabIndex        =   4
         Top             =   1308
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
         Left            =   6000
         TabIndex        =   5
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
         Left            =   2340
         TabIndex        =   6
         Top             =   570
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   6000
         TabIndex        =   7
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
         Left            =   1305
         TabIndex        =   15
         Top             =   945
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "”‰Ê«  : "
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
         TabIndex        =   37
         Top             =   225
         Width           =   1230
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
         Left            =   990
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   180
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   630
         Width           =   1140
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
         Index           =   0
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï :"
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
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   270
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… „‰ : "
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
         TabIndex        =   12
         Top             =   270
         Width           =   870
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
         TabIndex        =   11
         Top             =   630
         Width           =   1230
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
         TabIndex        =   10
         Top             =   974
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… „ÊœÌ·«  :"
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
         Index           =   5
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1356
         Width           =   1410
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
         TabIndex        =   8
         Top             =   585
         Width           =   1545
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   11280
      _ExtentX        =   19897
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
      Bindings        =   "VsTSalesMonth_YEAR.frx":9583
      Height          =   7470
      Left            =   90
      TabIndex        =   1
      Top             =   1980
      Width           =   15090
      _cx             =   26617
      _cy             =   13176
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
      FixedRows       =   2
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
End
Attribute VB_Name = "VsTSalesMonth_YEAR"
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
    cHead1 = Me.Caption & "  " & xStore.text
    cHead2 = " „‰  «—ÌŒ " & xDate1.text & " Õ Ï  «—ÌŒ " & xDate2.text
    cHead3 = ""
    If xFact.BoundText <> "" Then cHead3 = cHead3 & " „’‰⁄ " & xFact.text
    If xSupp.BoundText <> "" Then cHead3 = cHead3 & " „Ê—œ " & xSupp.text
    If xGroup.BoundText <> "" Then cHead3 = cHead3 & " „Ã„Ê⁄…" & xGroup.text
'   If XSEC.Text <> "" Then cHead3 = cHead3 & " ﬁ”„ " & XSEC.Text
    If xStore.BoundText <> "" Then cHead3 = cHead3 & " ›—⁄" & xStore.text
    If xMosm.BoundText <> "" Then cHead3 = cHead3 & " „Ê”„ " & xMosm.text
    Load PrintGrd
    PrintGrd.DOPRINT grid1, 1, , cHead1, cHead2, cHead3, , True, 11
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
    If XYEAR1.text = "" Or XYEAR2.text = "" Then Exit Sub
    If xDate1.text = "" Then xDate1.text = "1-1"
    If xDate2.text = "" Then xDate2.text = "31-12"
    myload
End Sub

Private Sub Command1_Click()
Dim aHeader(1)
Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With grid1
For nRow = 3 To .Rows - 1
    temptable.AddNew
    temptable!str1 = RetZero(nRow - 2, 2) & " " & .TextMatrix(nRow, 0)
    temptable!VAL1 = Val(.TextMatrix(nRow, 4))
    temptable!str3 = " „‰  «—ÌŒ " & xDate1.text & " ≈·Ï " & xDate2.text
    temptable.Update
Next nRow
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\SALES_MAN.rpt"
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing

End Sub

Private Sub Form_Load()
    openCon con
    
    XYEAR1.text = 2017
    XYEAR2.text = Year(Date)
    xmosm22.AddItem " "
    xmosm22.AddItem "S"
    xmosm22.AddItem "W"
    xmosm22.AddItem "M"
    
    
    DATA7.ConnectionString = strCon
    DATA7.RecordSource = "Select Code,DescA From STORE_BR WHERE " & cUserStore
    Set xStore.RowSource = DATA7
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
'   XSTORE.BoundText = cBranchStore
    
    If cBranch <> "00" Then
        xStore.BoundText = cBranchStore
        xStore.Enabled = False
    End If
    
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File4_10 order by Desca"
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    DATA3.ConnectionString = strCon
    DATA3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = DATA3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    DATA4.ConnectionString = strCon
    DATA4.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC "
    Set xMosm.RowSource = DATA4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
'    xMosm.BoundText = cPMosm
    
    DATA5.ConnectionString = strCon
    DATA5.RecordSource = "Select code ,desca From fact ORDER BY DESCA "
    Set xFact.RowSource = DATA5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA6
    
    grid1.Rows = 2
    grid1.Cols = 3
    FIXGRID
End Sub
Private Sub myload()
Dim I As Double, nYear As Double
Dim cString  As String, cStr2 As String, DDate1 As Date, DDate2 As Date
With grid1

    If SEC(0).Value <> 0 Then
        cString = " SELECT MONTH(QFILE6_20.[DATE]) , '' , FILE1_10.MOSM  "
        For nYear = Val(XYEAR1.text) To Val(XYEAR2.text)
            DDate1 = Mid(Format(xDate1.text, "DD-MM-YYYY"), 1, 6) & nYear
            DDate2 = Mid(Format(xDate2.text, "DD-MM-YYYY"), 1, 6) & nYear
            cString = cString & " , SUM(CASE WHEN DATE >= " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2) & " THEN (QFILE6_20.QUANT) ELSE 0 END ) "
            cString = cString & " , SUM(CASE WHEN DATE >= " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2) & " THEN (QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) ELSE 0 END ) "
        Next nYear
        cString = cString & " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM WHERE FILE1_10.MODELNO IS NOT NULL "
    End If

    If SEC(1).Value <> 0 Then
        cString = " SELECT  MONTH(QFILE6_20.[DATE]) , '' ,  FILE0_81.desca "
        For nYear = Val(XYEAR1.text) To Val(XYEAR2.text)
            DDate1 = Mid(Format(xDate1.text, "DD-MM-YYYY"), 1, 6) & nYear
            DDate2 = Mid(Format(xDate2.text, "DD-MM-YYYY"), 1, 6) & nYear
            cString = cString & " , SUM(CASE WHEN DATE >= " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2) & " THEN (QFILE6_20.QUANT) ELSE 0 END )  "
            cString = cString & " , SUM(CASE WHEN DATE >= " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2) & " THEN (QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) ELSE 0 END ) "
        Next nYear
        cString = cString & " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM  INNER JOIN STORE_BR ON STORE_BR.CODE = QFILE6_20.STORE LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX   WHERE FILE1_10.MODELNO IS NOT NULL "
    End If

    If SEC(2).Value <> 0 Then
        cString = " SELECT  MONTH(QFILE6_20.[DATE]) , ' ' ,FILE0_82.desca "
        For nYear = Val(XYEAR1.text) To Val(XYEAR2.text)
            DDate1 = Mid(Format(xDate1.text, "DD-MM-YYYY"), 1, 6) & nYear
            DDate2 = Mid(Format(xDate2.text, "DD-MM-YYYY"), 1, 6) & nYear
            cString = cString & " , SUM(CASE WHEN DATE >= " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2) & " THEN (QFILE6_20.QUANT) ELSE 0 END )  "
            cString = cString & " , SUM(CASE WHEN DATE >= " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2) & " THEN (QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) ELSE 0 END ) "
        Next nYear
        cString = cString & " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM LEFT JOIN FILE0_82 ON FILE1_10.[MODELAGE] = FILE0_82.CODE WHERE FILE1_10.MODELNO IS NOT NULL "
    End If
    If SEC(3).Value <> 0 Then
        cString = " SELECT  MONTH(QFILE6_20.[DATE])  , ' ' ,  ' '   "
        For nYear = Val(XYEAR1.text) To Val(XYEAR2.text)
            DDate1 = Mid(Format(xDate1.text, "DD-MM-YYYY"), 1, 6) & nYear
            DDate2 = Mid(Format(xDate2.text, "DD-MM-YYYY"), 1, 6) & nYear
            cString = cString & " , SUM(CASE WHEN DATE >= " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2) & " THEN (QFILE6_20.QUANT) ELSE 0 END )  "
            cString = cString & " , SUM(CASE WHEN DATE >= " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2) & " THEN (QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) ELSE 0 END ) "
        Next nYear
        cString = cString & " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM INNER JOIN STORE_BR ON STORE_BR.CODE = QFILE6_20.STORE WHERE FILE1_10.MODELNO IS NOT NULL "
    End If
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xmosm22.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M') "
    End If
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xStore.BoundText <> "" Then cStr2 = cStr2 & " AND STORE = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cStr2 = cStr2 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xall(0).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 1  "
    If xall(1).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 0 '  "

    If SEC(0).Value <> 0 Then cString = cString & cStr2 & " GROUP BY MONTH(QFILE6_20.[DATE]) , FILE1_10.MOSM   ORDER BY MONTH(QFILE6_20.[DATE]) , FILE1_10.MOSM  "
    If SEC(1).Value <> 0 Then cString = cString & cStr2 & " GROUP BY MONTH(QFILE6_20.[DATE]) , FILE0_81.desca  ORDER BY MONTH(QFILE6_20.[DATE]) , FILE0_81.desca "
    If SEC(2).Value <> 0 Then cString = cString & cStr2 & " GROUP BY MONTH(QFILE6_20.[DATE]) , FILE0_82.desca  ORDER BY MONTH(QFILE6_20.[DATE]) , FILE0_82.desca "
    If SEC(3).Value <> 0 Then cString = cString & cStr2 & " GROUP BY MONTH(QFILE6_20.[DATE])  ORDER BY MONTH(QFILE6_20.[DATE])   "
    Set DATA6.Recordset = myRecordSet(cString, con)
End With
FIXGRID
If grid1.Rows > 0 Then grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID()
With grid1
    .RowHeight(0) = 500
    .RowHeight(1) = 500
    .WordWrap = True
    .FixedRows = 2
    .TextMatrix(0, 0) = "—ﬁ„"
    .TextMatrix(0, 1) = "«·‘Â—"
    
    .TextMatrix(1, 0) = "—ﬁ„"
    .TextMatrix(1, 1) = "«·‘Â—"
    
    If SEC(0).Value <> 0 Then .TextMatrix(0, 2) = "«·„Ê”„"
    If SEC(1).Value <> 0 Then .TextMatrix(0, 2) = "«·‰Ê⁄"
    If SEC(2).Value <> 0 Then .TextMatrix(0, 2) = "«·›∆… «·⁄„—Ì…"
    
    If SEC(0).Value <> 0 Then .TextMatrix(1, 2) = "«·„Ê”„"
    If SEC(1).Value <> 0 Then .TextMatrix(1, 2) = "«·‰Ê⁄"
    If SEC(2).Value <> 0 Then .TextMatrix(1, 2) = "«·›∆… «·⁄„—Ì…"
    
    If SEC(3).Value <> 0 Then .ColHidden(2) = True
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    nYear = Val(XYEAR1.text)
    .ColWidth(0) = 500
    .ColWidth(1) = 2000
    .ColWidth(2) = 2000
    For nCol = 3 To .Cols - 1 Step 2
        .TextMatrix(1, nCol) = "⁄œœ"
        .TextMatrix(1, nCol + 1) = "ﬁÌ„…"
        
        .TextMatrix(0, nCol) = nYear
        .TextMatrix(0, nCol + 1) = nYear
        .ColWidth(nCol) = 1000
        .ColWidth(nCol + 1) = 1500
        .ColFormat(nCol) = "#0"
        .ColFormat(nCol + 1) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        nYear = nYear + 1
    Next nCol
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    For nRow = 2 To .Rows - 1
        If .TextMatrix(nRow, 0) <> "" Then .TextMatrix(nRow, 1) = arbMonth(.TextMatrix(nRow, 0))
    Next nRow
    .SubtotalPosition = flexSTAbove
    For I = 3 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
End Sub
Private Sub grid1_DblClick()
    VsTSales.xDate1.text = xDate1.text
    VsTSales.xDate2.text = xDate2.text
    VsTSales.xDate2.text = xDate2.text
    VsTSales.xMosm.BoundText = xMosm.BoundText
    VsTSales.Show
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

