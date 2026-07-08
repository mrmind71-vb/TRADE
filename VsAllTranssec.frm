VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VsAllTranssec 
   Caption         =   "≈Ã„«·Ï ﬁÌ„… Ê ⁄œœ ·„ÊœÌ·«  «·«ﬁ”«„ «·Ê«—œ… ··›—⁄"
   ClientHeight    =   10230
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
   ScaleHeight     =   10230
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame5 
      Height          =   555
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   0
      Width           =   4065
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
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   1275
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   180
         Width           =   735
      End
      Begin VB.OptionButton SEC 
         Alignment       =   1  'Right Justify
         Caption         =   "«·„Ê”„"
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
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   180
         Width           =   870
      End
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
         TabIndex        =   28
         Top             =   180
         Value           =   -1  'True
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      Height          =   510
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   585
      Width           =   4065
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
         TabIndex        =   26
         Top             =   180
         Width           =   960
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
         TabIndex        =   24
         Top             =   180
         Value           =   -1  'True
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1395
      Width           =   4920
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "VsAllTranssec.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "VsAllTranssec.frx":27EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "VsAllTranssec.frx":4CDD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2430
         Picture         =   "VsAllTranssec.frx":7149
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2205
      Left            =   5535
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   9660
      Begin VB.CheckBox XOutlet 
         Alignment       =   1  'Right Justify
         Caption         =   "Outlet "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   180
         Width           =   870
      End
      Begin VB.TextBox XMOSM2 
         Alignment       =   2  'Center
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
         Left            =   1530
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   135
         Width           =   825
      End
      Begin VB.ComboBox xmosm20 
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1755
         Width           =   825
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   4605
         TabIndex        =   3
         Top             =   945
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
         Left            =   180
         TabIndex        =   4
         Top             =   540
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   4605
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
         Left            =   180
         TabIndex        =   12
         Top             =   945
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   315
         Left            =   180
         TabIndex        =   14
         Top             =   1305
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSComCtl2.DTPicker xdate1 
         Height          =   375
         Left            =   6405
         TabIndex        =   21
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   5
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
         Format          =   445448193
         CurrentDate     =   43097
      End
      Begin MSComCtl2.DTPicker xdate2 
         Height          =   375
         Left            =   4605
         TabIndex        =   22
         Top             =   180
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   661
         _Version        =   393216
         MousePointer    =   5
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
         Format          =   445448193
         CurrentDate     =   43097
      End
      Begin MSDataListLib.DataCombo xsex 
         Height          =   315
         Left            =   4590
         TabIndex        =   33
         Top             =   1710
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   4605
         TabIndex        =   38
         Top             =   1305
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xage 
         Height          =   315
         Left            =   6225
         TabIndex        =   39
         Top             =   1710
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
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
         Index           =   2
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "›∆… ⁄„—Ì… :"
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
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "›∆… ⁄„—Ì… :"
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
         Left            =   10170
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "«·‰Ê⁄ :"
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
         Index           =   7
         Left            =   5670
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1755
         Width           =   555
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
         Index           =   4
         Left            =   1215
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1800
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ "
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
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1305
         Width           =   1005
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
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   990
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… :"
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
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   540
      End
      Begin VB.Label Label2 
         Caption         =   "„Ê”„ «·„ÊœÌ·"
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
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   540
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
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   975
         Width           =   1320
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
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   7
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
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   1320
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   9900
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
      Bindings        =   "VsAllTranssec.frx":9573
      Height          =   7290
      Left            =   90
      TabIndex        =   1
      Top             =   2250
      Width           =   16080
      _cx             =   28363
      _cy             =   12859
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
   Begin ComctlLib.ProgressBar prog1 
      Height          =   240
      Left            =   90
      TabIndex        =   37
      Top             =   1080
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc data12 
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
   Begin MSAdodcLib.Adodc DATA13 
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
   Begin VB.Label xModelNo 
      Alignment       =   1  'Right Justify
      Caption         =   "Label6"
      Height          =   285
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1125
      Visible         =   0   'False
      Width           =   2310
   End
End
Attribute VB_Name = "VsAllTranssec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub CMD_PRINT_Click()
End Sub
Private Sub cmdExel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CmdExit_Click()
    Unload Me
    Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = Me.Caption & xStore.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Value, "DD-MM-YYYY") & " Õ Ï  «—ÌŒ " & Format(xDate2.Value, "DD-MM-YYYY")
    cHead3 = ""
    If xFact.BoundText <> "" Then cHead3 = cHead3 & " „’‰⁄ " & xFact.text
    If xSupp.BoundText <> "" Then cHead3 = cHead3 & " „Ê—œ " & xSupp.text
    If xGroup.BoundText <> "" Then cHead3 = cHead3 & " „Ã„Ê⁄…" & xGroup.text
    If XSECTION.text <> "" Then cHead3 = cHead3 & " ﬁ”„ " & XSECTION.text
    If xMosm.BoundText <> "" Then cHead3 = cHead3 & " „Ê”„ " & xMosm.text
    PrintGrdNew.DOPRINT grid1, 0.8, -2, cHead1, cHead2, , , , False, 9
    PrintGrdNew.Show 1
End Sub
Private Sub Form_Load()
    openCon con
    xmosm20.AddItem " "
    xmosm20.AddItem "S"
    xmosm20.AddItem "W"
    xmosm20.AddItem "M"
    
    Set DATA13.Recordset = myRecordSet("Select Code,DescA From File0_82  order by Desca", con)
    Set xage.RowSource = DATA13
    xage.ListField = "Desca"
    xage.BoundColumn = "Code"
    
    Set data12.Recordset = myRecordSet("Select Code,DescA From File0_81  order by Desca", con)
    Set xsex.RowSource = data12
    xsex.ListField = "Desca"
    xsex.BoundColumn = "Code"
    
    Set DATA7.Recordset = myRecordSet("Select Code,DescA From STORE_BR WHERE " & cUserStore, con)
    Set xStore.RowSource = DATA7
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    If cBranch <> "00" Then
        xStore.BoundText = cBranchStore
        xStore.Enabled = False
    End If
    
    Set data1.Recordset = myRecordSet("Select Code,DescA From File1_10SC ORDER BY code ", con)
    Set XSECTION.RowSource = data1
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
    Set DATA2.Recordset = myRecordSet("Select Code,DescA From File4_10 order by Desca", con)
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    Set DATA3.Recordset = myRecordSet("Select Code,DescA From File1_50 ORDER BY DESCA", con)
    Set xGroup.RowSource = DATA3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set DATA4.Recordset = myRecordSet("Select mosm ,descA From mosm ORDER BY date DESC ", con)
    Set xMosm.RowSource = DATA4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm

    Set DATA5.Recordset = myRecordSet("Select code ,desca From fact ORDER BY DESCA ", con)
    Set xFact.RowSource = DATA5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA6
    
    grid1.Rows = 2
    grid1.Cols = 16
    FIXGRID
End Sub
Private Sub myload()
Dim I As Double
Dim cString  As String, cStr2 As String, cField As String
Dim cWhere As String
If IsDate(xDate1.Value) Then cWhere = " AND DATE >= " & DateSq(xDate1.Value)
With grid1
'                           0               1                 2                3               4                5
    
    If xStore.BoundText <> "" Then
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhere, "([IN]) ")
        
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhere, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST ) ")
        
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhere, "([OUT]) ")
        
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhere, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST ) ")
        
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhere, "([IN]-[OUT]) ")
        
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' or type = 't' or type = 'f'  )" & cWhere, "(([IN]-[OUT]) * FILE1_11_ALL.COST ) ")
        
        cField = cField & ",0," & _
                myiif("(type = '6'   )" & cWhere, "([OUT]) ")
        
        cField = cField & "," & _
                myiif("(type = '6'   )" & cWhere, "([OUT] * file1_11_all.cost ) ")
    Else
        cField = cField & "," & _
                myiif("(type = '2' )" & cWhere, "([IN]) ")
        
        cField = cField & "," & _
                myiif("(type = '2' )" & cWhere, "(FILE1_11_ALL.[IN] * FILE1_11_ALL.COST ) ")
        
        cField = cField & "," & _
                myiif("(type = '7' )" & cWhere, "([OUT]) ")
        
        cField = cField & "," & _
                myiif("(type = '7' )" & cWhere, "(FILE1_11_ALL.[OUT] * FILE1_11_ALL.COST ) ")
        
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' )" & cWhere, "([IN]-[OUT]) ")
        
        cField = cField & "," & _
                myiif("(type = '2' or type = '7' )" & cWhere, "(([IN]-[OUT]) * FILE1_11_ALL.COST ) ")
        
        cField = cField & ",0," & _
                myiif("(type = '6'   )" & cWhere, "([OUT]) ")
        
        cField = cField & "," & _
                myiif("(type = '6'   )" & cWhere, "([OUT] * file1_11_all.cost ) ")
    End If
    
    
    If SEC(3).Value <> 0 Then
        cString = " SELECT  FILE1_10SC.CODE , FILE1_10SC.DESCA , ' ' AS N2 , ' ' AS N3 , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & _
                " FROM  FILE1_10 INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM  WHERE file1_10.ISNOITEM = 0 "
    End If
    
    If SEC(0).Value <> 0 Then
        cString = " SELECT  FILE1_10SC.CODE, FILE1_10SC.DESCA , FILE1_10.[MOSM], FILE1_10.[MOSM] , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & _
                " FROM  FILE1_10 INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM  WHERE file1_10.ISNOITEM = 0 "
    End If
    If SEC(1).Value <> 0 Then
        cString = " SELECT  FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA AS factdesca , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & _
                " FROM  FILE1_10 INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM  LEFT JOIN FILE0_81 ON FILE0_81.CODE = FILE1_10.MODELSEX WHERE file1_10.ISNOITEM = 0 "
    End If
    If SEC(2).Value <> 0 Then
        cString = " SELECT  FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10.MODELAGE, FILE0_82.DESCA , ' ' AS N4 " & _
                    cField & " , SUM([IN]-[OUT]) , SUM(([IN]-[OUT]) * FILE1_11_ALL.COST ) , COUNT(DISTINCT FILE1_10.MODEL) " & _
                " FROM  FILE1_10 INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM  LEFT JOIN FILE0_82 ON FILE0_82.CODE = FILE1_10.MODELAGE WHERE file1_10.ISNOITEM = 0 "
    End If
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M') "
    End If
    If XSECTION.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[SECTION] = " & (XSECTION.BoundText)
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xStore.BoundText <> "" Then cStr2 = cStr2 & " AND [store] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cStr2 = cStr2 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
'   If IsDate(xdate1.Value) Then cStr2 = cStr2 & " AND [DATE] >= " & DateSq(xdate1.Value)
    If IsDate(xDate2.Value) Then cStr2 = cStr2 & " AND [DATE] <= " & DateSq(xDate2.Value)
    
    If xmosm20.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(xmosm20.text)
    If xage.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If XOutlet.Value <> 0 Then cStr2 = cStr2 & " AND Outlet = 1 "
    
    If xall(0).Value <> 0 Then cStr2 = cStr2 & " AND ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cStr2 = cStr2 & " AND ISBRANCH = 0 "

    
    If SEC(0).Value <> 0 Then
        cString = cString & cStr2 & " GROUP BY FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10.MOSM ORDER BY FILE1_10SC.CODE , FILE1_10.MOSM "
    End If
    If SEC(1).Value <> 0 Then
        cString = cString & cStr2 & " GROUP BY FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10.MODELSEX , FILE0_81.DESCA ORDER BY FILE1_10SC.CODE "
    End If
    If SEC(2).Value <> 0 Then
        cString = cString & cStr2 & " GROUP BY FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10.MODELAGE, FILE0_82.DESCA  ORDER BY FILE1_10SC.CODE "
    End If
    If SEC(3).Value <> 0 Then
        cString = cString & cStr2 & " GROUP BY FILE1_10SC.CODE , FILE1_10SC.DESCA ORDER BY FILE1_10SC.CODE "
    End If
    Set DATA6.Recordset = myRecordSet(cString, con)
End With
FIXGRID
If grid1.Rows > 0 Then grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID()
With grid1
    .Cols = 17
    .RowHeight(0) = 1000
    .RowHidden(1) = True
    .WordWrap = True
    
    .ColHidden(0) = True
    .ColHidden(2) = True
    .ColHidden(4) = True
    
    .TextMatrix(0, 1) = "«·ﬁ”„"
    
        
    If SEC(3).Value <> 0 Then .TextMatrix(0, 3) = ""
    If SEC(2).Value <> 0 Then .TextMatrix(0, 3) = "›∆… ⁄„—Ì…"
    If SEC(1).Value <> 0 Then .TextMatrix(0, 3) = "«·‰Ê⁄"
    If SEC(0).Value <> 0 Then .TextMatrix(0, 3) = "«·„Ê”„"
    
    .TextMatrix(0, 5) = "ﬂ„Ì… Ê«—œ"
    .TextMatrix(0, 6) = "ﬁÌ„… «·Ê«—œ"
    
    .TextMatrix(0, 7) = "ﬂ„Ì… ’«œ—"
    .TextMatrix(0, 8) = "ﬁÌ„… ’«œ—"
    
    .TextMatrix(0, 9) = "’«›Ï «·ﬂ„Ì…"
    .TextMatrix(0, 10) = "’«›Ï «·ﬁÌ„…"
    
    .TextMatrix(0, 11) = "‰”»… „»Ì⁄« "
    
    .TextMatrix(0, 12) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 13) = " ﬂ·›… „»Ì⁄« "
    
    .TextMatrix(0, 14) = "—’Ìœ ⁄œœ"
    .TextMatrix(0, 15) = "ﬁÌ„… «·—’Ìœ"
    .TextMatrix(0, 16) = "⁄œœ „ÊœÌ·« "
    
    .ColHidden(6) = Not bOpt10
    .ColHidden(8) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    .ColHidden(15) = Not bOpt10
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1500
    .ColWidth(2) = 900
    .ColWidth(3) = 1500
    .ColWidth(4) = 2100
    
    .ColWidth(5) = 900
    .ColWidth(6) = 1200
    .ColWidth(7) = 900
    .ColWidth(8) = 1200
    .ColWidth(9) = 900
    .ColWidth(10) = 1200
    .ColWidth(11) = 900
    
    .ColWidth(12) = 900
    .ColWidth(13) = 1200
    
    .ColWidth(14) = 900
    .ColWidth(15) = 1200
    .ColWidth(16) = 800
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    For nCol = 5 To 9 Step 2
        .ColFormat(nCol) = "#0"
        .ColFormat(nCol + 1) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol + 1) = Format(.TextMatrix(nRow, nCol + 1), "#0.00")
        Next nRow
    Next nCol
    .ColFormat(12) = "#0"
    .ColFormat(13) = "#0.00"
    .ColDataType(12) = flexDTDouble
    .ColDataType(13) = flexDTDouble
    
    .ColFormat(14) = "#0"
    .ColFormat(15) = "#0.00"
    .ColDataType(14) = flexDTDouble
    .ColDataType(15) = flexDTDouble
    
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    Next nRow
    .SubtotalPosition = flexSTAbove
    For I = 5 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    If .Rows > 2 Then
        nRow = 2
        If Val(.TextMatrix(nRow, 9)) <> 0 Then .TextMatrix(nRow, 11) = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 9))) * 100, 2)
    End If
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
End Sub

Private Sub xSection_LostFocus()
    DATA3.ConnectionString = strCon
    If XSECTION.BoundText = "" Then
        DATA3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Else
        DATA3.RecordSource = "Select Code,DescA From File1_50 where [group] = " & Val(XSECTION.BoundText) & " ORDER BY DESCA"
    End If
    Set xGroup.RowSource = DATA3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    DATA3.Refresh
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
