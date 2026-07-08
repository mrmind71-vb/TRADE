VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsTBranshSalModel 
   Caption         =   "≈Ã„«·Ï „»Ì⁄«  „ÊœÌ·«  „Ê“⁄ ⁄·Ï «·›—Ê⁄"
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
      TabIndex        =   27
      Top             =   1350
      Width           =   3840
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   180
         Value           =   -1  'True
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   -45
      Width           =   3840
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "VsTBranshSalModel.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "VsTBranshSalModel.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2475
         Picture         =   "VsTBranshSalModel.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   24
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
         Picture         =   "VsTBranshSalModel.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   240
         Left            =   45
         TabIndex        =   26
         Top             =   1125
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   423
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1800
      Left            =   4005
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   -45
      Width           =   11100
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊœÌ·«  ·Â« „‘ —Ì«  ›ﬁÿ Œ·«· › —…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   225
         Width           =   1905
      End
      Begin VB.TextBox xDate2 
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
         TabIndex        =   25
         Top             =   225
         Width           =   1680
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
         Left            =   7980
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox xDescItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEAEA&
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
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   930
         Width           =   3015
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
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1320
         Width           =   4590
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
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   3
         Top             =   930
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   6000
         TabIndex        =   7
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
         TabIndex        =   8
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
         Left            =   2115
         TabIndex        =   9
         Top             =   180
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   6000
         TabIndex        =   10
         Top             =   596
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   2115
         TabIndex        =   11
         Top             =   540
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›‰—…"
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
         TabIndex        =   19
         Top             =   270
         Width           =   450
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
         TabIndex        =   18
         Top             =   225
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
         TabIndex        =   17
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1356
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ’‰› :"
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
         Index           =   4
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1350
         Width           =   1155
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ÊœÌ· :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   990
         Width           =   915
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
         TabIndex        =   13
         Top             =   585
         Width           =   1320
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
         Index           =   1
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   615
         Width           =   1230
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
      Bindings        =   "VsTBranshSalModel.frx":9583
      Height          =   7380
      Left            =   135
      TabIndex        =   1
      Top             =   1845
      Width           =   14955
      _cx             =   26379
      _cy             =   13017
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
End
Attribute VB_Name = "VsTBranshSalModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StoreTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me

End Sub

Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = Me.Caption
    cHead2 = " Õ Ï  «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY")
    
    Dim cHead3 As String
    cHead3 = XSECTION.text & "  " & xGroup.text & "  " & xFact.text & "  " & xSupp.text
    
    Load PrintGrd
    PrintGrd.DOPRINT grid1, 0.9, , cHead1, cHead2, cHead3, , True, 9, , Array(1, 2)
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
    If StoreTable.State = adStateOpen Then StoreTable.Close
    If xall(0).Value <> 0 Then
        If bOpt5 Then
            StoreTable.Open "select CODE , DESCA FROM STORE_BR WHERE S_BRANCH < '60' ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
        Else
            StoreTable.Open "select CODE , DESCA FROM STORE_BR WHERE S_BRANCH < '60' AND CODE IN (SELECT STORE FROM USERSHOP WHERE CODE = " & MyParn(nusercode) & " )  ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
        End If
    End If
    If xall(1).Value <> 0 Then
        If bOpt5 Then
            StoreTable.Open "select CODE , DESCA FROM STORE_BR WHERE S_BRANCH > '60' ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
        Else
            StoreTable.Open "select CODE , DESCA FROM STORE_BR WHERE S_BRANCH > '60' AND CODE IN (SELECT STORE FROM USERSHOP WHERE CODE = " & MyParn(nusercode) & " )  ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
        End If
    End If
    If xall(2).Value <> 0 Then
        If bOpt5 Then
            StoreTable.Open "select CODE , DESCA FROM STORE_BR ORDER BY S_BRANCH ", con, adOpenStatic, adLockReadOnly, adCmdText
        Else
            StoreTable.Open "select CODE , DESCA FROM STORE_BR WHERE CODE IN (SELECT STORE FROM USERSHOP WHERE CODE = " & MyParn(nusercode) & " )  ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
        End If
    End If
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    If xall(0).Value <> 0 Then
        StoreTable.Open "select CODE , DESCA FROM STORE_BR WHERE S_BRANCH < '60' ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
    End If
    If xall(1).Value <> 0 Then
        StoreTable.Open "select CODE , DESCA FROM STORE_BR WHERE S_BRANCH > '60' ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
    End If
    If xall(2).Value <> 0 Then
        StoreTable.Open "select CODE , DESCA FROM STORE_BR ORDER BY S_BRANCH ", con, adOpenStatic, adLockReadOnly, adCmdText
    End If
    
    xDate1.text = Format("1-1-" & Year(Date), "dd-mm-yyyy")
    xDate2.text = Format(Date, "dd-mm-yyyy")
    
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
    Set XSECTION.RowSource = data1
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
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
    xMosm.BoundText = cPMosm
    
    DATA5.ConnectionString = strCon
    DATA5.RecordSource = "Select code ,desca From fact ORDER BY code "
    Set xFact.RowSource = DATA5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    grid1.Rows = 2
    grid1.FixedRows = 2
    grid1.Cols = 7
'    FixGrid
End Sub
Private Sub myload()
Dim I As Double
Dim cStrall  As String

grid1.Rows = 0
grid1.Rows = 2
grid1.FixedRows = 2
grid1.Cols = 7

Dim lOk As Boolean
StoreTable.MoveFirst
Do Until StoreTable.EOF
     
    cWhere = " ( FILE1_11_ALL.TYPE = '2' OR FILE1_11_ALL.TYPE = '7'  OR FILE1_11_ALL.TYPE = 'F'  OR FILE1_11_ALL.TYPE = 'T' ) AND STORE = " & MyParn(StoreTable!CODE)
    cWhere = cWhere & " AND ( DATE >= " & DateSq(xDate1.text) & " AND DATE <= " & DateSq(xDate2.text) & " ) "
    cField2 = cField2 & " , " & myiif(cWhere, "[IN] - [OUT]") & " AS " & addstring(StoreTable!CODE)
    
    cWhere = " FILE1_11_ALL.TYPE = '6' AND STORE = " & MyParn(StoreTable!CODE)
    cWhere = cWhere & " AND ( DATE >= " & DateSq(xDate1.text) & " AND DATE <= " & DateSq(xDate2.text) & " ) "
    cField2 = cField2 & " , " & myiif(cWhere, "[OUT] ")
    
    cWhere = " STORE = " & MyParn(StoreTable!CODE)
    cWhere = cWhere & " AND ( DATE <= " & DateSq(xDate2.text) & " ) "
    cField2 = cField2 & " , " & myiif(cWhere, "[IN] - [OUT]")
    StoreTable.MoveNext
Loop

cWhere = " (FILE1_11_ALL.TYPE = '2' OR FILE1_11_ALL.TYPE = '7'  ) "
cField1 = myiif(cWhere, "[IN] - [OUT]")

cWhere = " (FILE1_11_ALL.TYPE = '6'   ) "
cField3 = myiif(cWhere, "[OUT]")



With grid1
'                           0               1                 2                3               4                        5
    cStrall = " SELECT file1_10.ModelNo, file1_10.mosm, FACT.desca , file1_10.modelfact0 , file1_10.desca  , MIN(FILE1_10.PRICE) , " & _
                cField1 & cField2 & " , " & cField3 & " , sum([in] - [OUT]) as bal " & _
            " FROM  FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE  WHERE file1_10.ISNOITEM = 0 "
    If xMosm.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.text <> "" Then cStrall = cStrall & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xSupp.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XSECTION.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.[Section] = " & XSECTION.BoundText
    If XDESCA.text <> "" Then cStrall = cStrall & " AND file1_10.DESCA LIKE ('%" & XDESCA.text & "%')   "
    
    If xall(0).Value <> 0 Then cStrall = cStrall & " AND FILE1_11_ALL.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cStrall = cStrall & " AND FILE1_11_ALL.ISBRANCH = 2 "
    If Not bOpt5 Then cStrall = cStrall & " AND STORE IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & ")"
    If Check1.Value <> 0 And IsDate(xDate1.text) And IsDate(xDate2.text) And xSupp.BoundText <> "" Then
        cStrall = cStrall & " and " & " FILE1_10.ITEM IN (SELECT ITEM FROM file7_20 inner join file7_20h on file7_20h.doc_no = file7_20.doc_no WHERE CODE = " & MyParn(xSupp.BoundText) & " AND DATE >= " & DateSq(xDate1.text) & " AND DATE <= " & DateSq(xDate2.text) & "  ) "
    End If
    cStrall = cStrall & " GROUP BY file1_10.ModelNo, FACT.desca , file1_10.mosm, file1_10.desca  , modelfact0 "
    DATA6.RecordSource = cStrall
    DATA6.Refresh
End With
FIXGRID
'If grid1.Rows > 0 Then grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID()
Dim nColS1 As Double
Dim nColS2 As Double
With grid1
'    .FixedRows = 2
    .RowHeight(0) = 700
    .RowHeight(1) = 700
    .WordWrap = True
    .MergeCells = flexMergeFixedOnly
    
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    
    .MergeRow(0) = True
    .MergeRow(1) = True
    
    .TextMatrix(0, 0) = "„ÊœÌ·"
    .TextMatrix(0, 1) = "«·„Ê”„"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "—ﬁ„ „ÊœÌ·"
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = "”⁄— „” Â·ﬂ"
    .TextMatrix(0, 6) = "„‘ —Ì« "
    
    .TextMatrix(1, 0) = "„ÊœÌ·"
    .TextMatrix(1, 1) = "«·„Ê”„"
    .TextMatrix(1, 2) = "«·„’‰⁄"
    .TextMatrix(1, 3) = "—ﬁ„ „ÊœÌ·"
    .TextMatrix(1, 4) = "«·’‰›"
    .TextMatrix(1, 5) = "”⁄— „” Â·ﬂ"
    .TextMatrix(1, 6) = "„‘ —Ì« "
    
    .ColWidth(0) = 0
    .ColWidth(1) = 700
    .ColWidth(2) = 1800
    .ColWidth(3) = 1200
    .ColWidth(4) = 2500
    .ColWidth(5) = 700
    .ColWidth(6) = 700
    
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .ColHidden(7) = True
    For nCol = 7 To .Cols - 3 Step 3
                
        .TextMatrix(1, nCol) = "Ê«—œ"
        .TextMatrix(1, nCol + 1) = "„»Ì⁄« "
        .TextMatrix(1, nCol + 2) = "—’Ìœ"
        
        .TextMatrix(0, nCol) = GetDesca("SELECT DESCA FROM STORE_BR WHERE CODE = " & MyParn(.TextMatrix(0, nCol)), con)
        .TextMatrix(0, nCol + 1) = .TextMatrix(0, nCol)
        .TextMatrix(0, nCol + 2) = .TextMatrix(0, nCol)
        .ColWidth(nCol) = 700
        .ColWidth(nCol + 1) = 700
        .ColWidth(nCol + 2) = 700
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        .ColDataType(nCol + 2) = flexDTDouble
    Next nCol
    .TextMatrix(0, .Cols - 2) = "Ã „»Ì⁄« "
    .TextMatrix(0, .Cols - 1) = "«·—’Ìœ"
    .TextMatrix(1, .Cols - 2) = "Ã „»Ì⁄« "
    .TextMatrix(1, .Cols - 1) = "«·—’Ìœ"
    .ColWidth(.Cols - 1) = 700
    .ColWidth(.Cols - 2) = 700
    .MergeCol(.Cols - 1) = True
    .MergeCol(.Cols - 2) = True
    For nRow = 2 To .Rows - 1
        For nCol = 6 To .Cols - 1
            .TextMatrix(nRow, nCol) = Format(Val(.TextMatrix(nRow, nCol)), "#0")
        Next nCol
    Next nRow
    .SubtotalPosition = flexSTAbove
    On Error Resume Next
    For I = 6 To .Cols - 1
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
        If Val(.TextMatrix(2, I)) = 0 Then .ColHidden(I) = True
    Next I
    End With
    Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    StoreTable.Close
    Set StoreTable = Nothing
    closeCon con
End Sub
Private Sub grid1_DblClick()
    If grid1.Col >= 6 Then
        Load ModelMove
        ModelMove.XMODEL.text = grid1.TextMatrix(grid1.Row, 0)
        If grid1.Col > 6 Then ModelMove.xStore.BoundText = grid1.TextMatrix(1, grid1.Col)
        ModelMove.Show
    End If
End Sub
