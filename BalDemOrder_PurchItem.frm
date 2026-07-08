VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form BalDemOrder_PurchItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ›’Ì·Ï ÿ·»Ì«  „ÊœÌ·«  „‰ „Ê—œÌ‰"
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
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   135
      Width           =   3840
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2475
         Picture         =   "BalDemOrder_PurchItem.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "BalDemOrder_PurchItem.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "BalDemOrder_PurchItem.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   26
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
         Picture         =   "BalDemOrder_PurchItem.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   45
         TabIndex        =   27
         Top             =   1125
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame4 
      Height          =   960
      Left            =   4770
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   0
      Width           =   4380
      Begin VB.OptionButton XORD 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ· ÿ·»Ì« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   0
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton XORD 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì— „‰›–…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   1
         Left            =   1687
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton XORD 
         Alignment       =   1  'Right Justify
         Caption         =   "„‰›–…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Index           =   2
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8745
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1575
      Width           =   19950
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   8475
         Left            =   270
         TabIndex        =   8
         Top             =   180
         Width           =   19500
         _cx             =   34396
         _cy             =   14949
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
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   9180
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   10905
      Begin VB.TextBox XORDER_NO 
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
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   945
         Width           =   1365
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
         Left            =   8295
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   135
         Width           =   1365
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
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo XCODE 
         Height          =   315
         Left            =   3375
         TabIndex        =   2
         Top             =   180
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   135
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
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   3375
         TabIndex        =   9
         Top             =   540
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   315
         Left            =   135
         TabIndex        =   10
         Top             =   540
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„” ‰œ ÿ·»Ì… :"
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
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   990
         Width           =   1065
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
         Left            =   9765
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   870
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
         Left            =   9810
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   450
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
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
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   540
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄ :"
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
         Left            =   6225
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Left            =   6225
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   210
         Width           =   600
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
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1050
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
      Left            =   495
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   3885
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
      Left            =   2685
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
      Left            =   1590
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
End
Attribute VB_Name = "BalDemOrder_PurchItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Variant, sM_Fact As String
Dim Search1 As New Search3, oSearch As New Search3, oSearchSup As New Search3
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me

End Sub

Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = Me.Caption & " " & XSTORE.text & " " & XSECTION.text & " " & XCODE.text
    
    Set printGrd3.grid1 = grid1
    Set printGrd3.myForm = Me
    printGrd3.doprint 0.9, 0, cHead1, , , False, True, 9, 1
    printGrd3.Show 1

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From STORE_BR where " & cUserStore
    Set XSTORE.RowSource = data1
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File1_10SC ORDER BY DESCA"
    Set XSECTION.RowSource = DATA2
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC "
    Set xMosm.RowSource = data3
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    
    data4.ConnectionString = strCon
    data4.RecordSource = "Select code ,desca From FILE4_10  ORDER BY DESCA "
    Set XCODE.RowSource = data4
    XCODE.ListField = "Desca"
    XCODE.BoundColumn = "Code"
    
    Set grid1.DataSource = data6
    
    grid1.Rows = 1
    grid1.Cols = 10
    
    Fixgrd
End Sub
Private Sub myload()
    Dim cWhere As String, cFiled1 As String, cString As String

'                                           0                       1                       2                           3                       4               5                   6               7               8           9               10              11                          12                          13                              14                                  15
    cString = " SELECT  Q_DOC_ORDER_ITEM.STORE_NAME, Q_DOC_ORDER_ITEM.DESCA, min(Q_DOC_ORDER_ITEM.DATE), Q_DOC_ORDER_ITEM.DOC_NO, Q_DOC_ORDER_ITEM.MOSM, FILE1_10SC.DESCA , FILE1_10.MODELFACT0, FILE1_10.DESCA , FILE1_10.COLOR, FILE1_10.SCAL, FILE1_10.ITEM, FILE1_10.COST, sum(Q_DOC_ORDER_ITEM.QUANT), SUM(Q_DOC_ORDER_ITEM.QUANT_2), MAX(Q_DOC_ORDER_ITEM.DATE_PURCH), Q_DOC_ORDER_ITEM.ISEND_ORDER "
    cString = cString & " FROM  FILE1_10 INNER JOIN Q_DOC_ORDER_ITEM ON FILE1_10.ITEM = Q_DOC_ORDER_ITEM.ITEM INNER JOIN FILE1_10SC ON FILE1_10.SECTION = FILE1_10SC.CODE WHERE doc_no IS NOT NULL "
    If XSECTION.BoundText <> "" Then cString = cString & " AND FILE1_10.[SECTION] = " & Val(XSECTION.BoundText)
    If XSTORE.BoundText <> "" Then cString = cString & " AND [STORE] = " & MyParn(XSTORE.BoundText)
    If XCODE.BoundText <> "" Then cString = cString & " AND Q_DOC_ORDER_ITEM.[CODE] = " & MyParn(XCODE.BoundText)
    If XORDER_NO.text <> "" Then cString = cString & " AND Q_DOC_ORDER_ITEM.[doc_no] = " & MyParn(XORDER_NO.text)
    If Not bOpt5 Then cString = cString & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If IsDate(xDate1.text) Then cString = cString & " AND DATE >=" & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cString = cString & " AND DATE <=" & DateSq(xdate2.text)
    
    cString = cString & " GROUP BY Q_DOC_ORDER_ITEM.STORE_NAME, Q_DOC_ORDER_ITEM.DESCA, Q_DOC_ORDER_ITEM.MOSM, FILE1_10SC.DESCA  , Q_DOC_ORDER_ITEM.DOC_NO, FILE1_10.MODELFACT0, FILE1_10.COLOR, FILE1_10.SCAL, FILE1_10.ITEM, FILE1_10.COST , Q_DOC_ORDER_ITEM.ISEND_ORDER , file1_10.desca "
    cString = cString & " ORDER BY Q_DOC_ORDER_ITEM.STORE_NAME, Q_DOC_ORDER_ITEM.DESCA, min(Q_DOC_ORDER_ITEM.DATE), Q_DOC_ORDER_ITEM.DOC_NO, Q_DOC_ORDER_ITEM.MOSM, FILE1_10SC.DESCA  , FILE1_10.MODELFACT0, FILE1_10.COLOR, FILE1_10.SCAL "
    Set data6.Recordset = myRecordSet(cString, con)
    Fixgrd
End Sub
Sub Fixgrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .RowHeight(0) = 1000
    .Cols = 16
    
    .TextMatrix(0, 0) = "«·›—⁄"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    .TextMatrix(0, 2) = " «—ÌŒ"
    .TextMatrix(0, 3) = " „” ‰œ"
    .TextMatrix(0, 4) = "«·„Ê”„"
    .TextMatrix(0, 5) = "«·ﬁ”„"
    .TextMatrix(0, 6) = "—ﬁ„"
    .TextMatrix(0, 7) = "„ÊœÌ·"
    .TextMatrix(0, 8) = "·Ê‰"
    .TextMatrix(0, 9) = "„ﬁ«”"
    .TextMatrix(0, 11) = "”⁄— „’‰⁄"
    .TextMatrix(0, 12) = "ﬂ„Ì…"
    .TextMatrix(0, 13) = "ﬂ„Ì… „‰›–"
    .TextMatrix(0, 14) = "«Œ—  «—ÌŒ"
    .TextMatrix(0, 15) = "„‰›–"
    
    .ColWidth(0) = 1300
    .ColWidth(1) = 1500
    .ColWidth(2) = 1300
    .ColWidth(3) = 1000
    .ColWidth(4) = 800
    .ColWidth(5) = 1000
    .ColWidth(6) = 1400
    .ColWidth(7) = 2800
    .ColWidth(8) = 1300
    .ColWidth(9) = 800
    .ColWidth(10) = 800
    .ColWidth(11) = 800
    .ColWidth(12) = 800
    .ColWidth(13) = 800
    .ColWidth(14) = 1300
    .ColWidth(15) = 600
    
    .ColDataType(14) = flexDTDate
    .ColDataType(4) = flexDTDate
    .ColDataType(15) = flexDTBoolean
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 8)) <> 0 Then
            nRate = Round((Val(.TextMatrix(nRow, 12)) / Val(.TextMatrix(nRow, 8))) * 100, 2)
            .TextMatrix(nRow, 12) = nRate
        End If
    Next nRow
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 12, "#0", vbYellow, , True, " "
    .Subtotal flexSTSum, -1, 13, "#0", vbYellow, , True, " "
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim cDoc_No As String
On Error GoTo myerror
con.BeginTrans
With grid1
If Col = 12 Then
    If MsgBox(" Õ–› «·ÿ·»Ì… »«·ﬂ«„· ", vbYesNo + vbDefaultButton2) = vbYes Then
        cStr1 = " DELETE FROM Model_Dem FROM  Model_Dem INNER JOIN FILE1_10 ON Model_Dem.ITEM = FILE1_10.ITEM  WHERE STORE = " & MyParn(.TextMatrix(.Row, 0)) & " AND FILE1_10.CODE = " & MyParn(.TextMatrix(.Row, 2))
        If Check1.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(.TextMatrix(.Row, 4))
        If Check2.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(.TextMatrix(.Row, 5))
        If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
        If XSECTION.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(XSECTION.BoundText)
        con.Execute cStr1, nRec
        Inform "  „ Õ–›  " & nRec
        .RemoveItem .Row
    End If
End If

If Col = 11 Then
    If MsgBox("  —ÕÌ· «·ÿ·»Ì… ·„” ‰œ «„— ‘—«¡ ", vbYesNo + vbDefaultButton2) = vbYes Then
        
        cDoc_No = RetZero(Newflag("FILE6_50H", "doc_no", con))
        cStr1 = "  INSERT INTO FILE6_50H (STORE, CODE, MOSM, date , doc_no  ) SELECT Model_Dem.STORE, FILE1_10.code ,MAX(FILE1_10.MOSM) , " & addDate(Date) & " , " & addstring(cDoc_No) & "  FROM Model_Dem INNER JOIN FILE1_10 ON Model_Dem.ITEM = FILE1_10.ITEM WHERE STORE = " & MyParn(.TextMatrix(.Row, 0)) & " AND FILE1_10.CODE = " & MyParn(.TextMatrix(.Row, 2))
        If Check1.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(.TextMatrix(.Row, 4))
        If Check2.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(.TextMatrix(.Row, 5))
        If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
        If XSECTION.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(XSECTION.BoundText)
        cStr1 = cStr1 & " group by file1_10.code , Model_Dem.STORE "
        con.Execute cStr1
        
        cStr1 = "  INSERT INTO FILE6_50 (doc_no , item , quant, price , discount ,model ) SELECT " & addstring(cDoc_No) & " , Model_Dem.item  , quant , file1_10.cost ,0, file1_10.model   FROM Model_Dem INNER JOIN FILE1_10 ON Model_Dem.ITEM = FILE1_10.ITEM  WHERE quant <> 0 AND STORE = " & MyParn(.TextMatrix(.Row, 0)) & " AND FILE1_10.CODE = " & MyParn(.TextMatrix(.Row, 2))
        If Check1.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(.TextMatrix(.Row, 4))
        If Check2.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(.TextMatrix(.Row, 5))
        If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
        If XSECTION.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(XSECTION.BoundText)
        con.Execute cStr1
        
        Inform "  „ ⁄„· «·„” ‰œ " & nRec

        cStr1 = " DELETE FROM Model_Dem FROM  Model_Dem INNER JOIN FILE1_10 ON Model_Dem.ITEM = FILE1_10.ITEM  WHERE STORE = " & MyParn(.TextMatrix(.Row, 0)) & " AND FILE1_10.CODE = " & MyParn(.TextMatrix(.Row, 2))
        If Check1.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(.TextMatrix(.Row, 4))
        If Check2.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(.TextMatrix(.Row, 5))
        If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
        If XSECTION.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(XSECTION.BoundText)
        con.Execute cStr1, nRec
        Inform "  „ Õ–›  " & nRec
        .RemoveItem .Row
    End If
End If
End With
con.CommitTrans
Exit Sub
con.RollbackTrans
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_DblClick()
    MYLOAD2
End Sub
Private Sub grid1_EnterCell()
With grid1
    If .Col >= 11 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Sub Fixgrd2()
    With grid2
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .RowHeight(0) = 1000
    .Cols = 11
    .ColHidden(8) = True
    .TextMatrix(0, 0) = "—ﬁ„ „ÊœÌ·"
    .TextMatrix(0, 1) = "«·„ÊœÌ·"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "«··Ê‰"
    .TextMatrix(0, 4) = "«·„ﬁ«”"
    .TextMatrix(0, 5) = "«·”⁄—"
    .TextMatrix(0, 6) = " ﬂ„Ì…"
    .TextMatrix(0, 7) = " «—ÌŒ"
    .TextMatrix(0, 8) = " «·›—⁄"
    .TextMatrix(0, 10) = " Õ–›"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1700
    .ColWidth(2) = 700
    .ColWidth(3) = 1000
    .ColWidth(4) = 700
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    .ColWidth(7) = 1300
    .ColWidth(8) = 600
    .ColWidth(9) = 600
    .ColWidth(10) = 600
    .ColComboList(10) = "..."
    .ColDataType(7) = flexDTDate
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    .ColHidden(9) = True
End With
End Sub
Sub MYLOAD2()
    With grid1
'                       0                       1               2               3               4               5               6                   7           8
    cStr1 = " SELECT FILE1_10.MODELFACT0, FILE1_10.desca, FILE1_10.mosm, FILE1_10.COLOR, FILE1_10.SCAL, FILE1_10.price , Model_Dem.QUANT, Model_Dem.DATE, Model_Dem.Branch , MODEL_DEM.ITEM  FROM    Model_Dem INNER JOIN FILE1_10 ON Model_Dem.ITEM = FILE1_10.ITEM where quant > 0 " & _
        " and STORE = " & MyParn(.TextMatrix(.Row, 0)) & _
        " AND FILE1_10.CODE = " & MyParn(.TextMatrix(.Row, 2))
        If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
        If XSECTION.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[SECTION]= " & Val(XSECTION.BoundText)
        If Check1.Value = 1 Then cStr1 = cStr1 & " AND FILE1_10.[mosm]= " & MyParn(.TextMatrix(.Row, 4))
        If Check1.Value = 2 Then cStr1 = cStr1 & " AND FILE1_10.[SECTION]= " & Val(.TextMatrix(.Row, 5))
        cStr1 = cStr1 & " ORDER BY FILE1_10.MODELFACT0, FILE1_10.desca, FILE1_10.COLOR , FILE1_10.SCAL "
    End With
    DATA7.RecordSource = cStr1
    DATA7.Refresh
    Fixgrd2
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid2
    If .Col = 6 Then
        con.Execute " UPDATE  Model_Dem SET QUANT = " & Val(.TextMatrix(.Row, 5)) & " WHERE ITEM = " & Val(.TextMatrix(.Row, 8)) & " AND STORE = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
    End If
End With
End Sub
Private Sub GRID2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Col = 10 Then
    cStr1 = " DELETE FROM Model_Dem WHERE STORE = " & MyParn(grid1.TextMatrix(grid1.Row, 0)) & " AND Model_Dem.item = " & Val(grid2.TextMatrix(grid2.Row, 9))
    con.Execute cStr1, nRec
    Inform "  „ Õ–›  " & nRec
    grid2.RemoveItem Row
End If
End Sub
Private Sub grid2_EnterCell()
With grid2
    If .Col = 6 Or .Col = 10 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub XORDER_NO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        Dim Generalarray(5)
        Dim listarray(1, 4)
        Dim GrdArray(6, 1)
        
        Set Generalarray(0) = Me
        Generalarray(1) = "SELECT     top 200 DOC_NO,[DATE] AS DATE1,DATE, FILE4_10.Desca , inv_no , ' ', STORE_BR.desca " & _
                          " FROM  file6_50h inner JOIN file4_10 ON file6_50h.CODE  = FILE4_10.CODE inner join STORE_BR on STORE_BR.code = file6_50h.store  where ISEND_ORDER = 0 "
        Generalarray(2) = "Order by DATE DESC , doc_no desc "
        Generalarray(3) = 6000
        Generalarray(5) = True
        
        listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
        listarray(0, 1) = "(inv_no Like '%cFilter%' or Doc_No Like '%cFilter%' or  FILE4_10.DESCA LIKE '%cFilter%' OR " & _
                          "##date##)"
        
        listarray(1, 0) = "«·›—⁄"
        listarray(1, 1) = "STORE_BR.desca Like '%cFilter%' "
        
        GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
        GrdArray(0, 1) = 1000
        
        GrdArray(1, 0) = "«· «—ÌŒ"
        GrdArray(1, 1) = 0
        
        GrdArray(2, 0) = "«· «—ÌŒ"
        GrdArray(2, 1) = 1500
        
        GrdArray(3, 0) = "≈”„ " & cCodeDesca
        GrdArray(3, 1) = 3000
        
        GrdArray(4, 0) = "—ﬁ„ ›« Ê—…"
        GrdArray(4, 1) = 1000
        
        GrdArray(5, 0) = "„Ê”„"
        GrdArray(5, 1) = 0
        
        GrdArray(6, 0) = "«·„Œ“‰"
        GrdArray(6, 1) = 1500
        
        searchArray = Array(Generalarray, listarray, GrdArray)
        Load Search1
        Search1.Caption = "«” ⁄·«„"
        Search1.Show 1
    End If
End Sub
Sub myProc()
If ActiveControl.Name = XORDER_NO.Name Then
    XORDER_NO.text = Search1.grid1.TextMatrix(Search1.grid1.Row, 0)
    Search1.Hide
ElseIf ActiveControl.Name = XCODE.Name Then
    ActiveControl.BoundText = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    Unload oSearchSup
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then SuppLookupAll Me, oSearchSup
End Sub

