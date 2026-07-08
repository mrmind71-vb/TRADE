VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form grditem_br 
   Caption         =   "„—«Ã⁄… «·«—’œ… ··›—Ê⁄  Ê «·”Ì—›—"
   ClientHeight    =   9405
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   14955
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
   ScaleHeight     =   9405
   ScaleWidth      =   14955
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   " ÕœÌÀ «·„” ‰œ"
      Height          =   240
      Left            =   13230
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   10620
      Width           =   1590
   End
   Begin VB.Frame Frame3 
      Caption         =   "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   315
      Width           =   2040
      Begin VB.Label xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   360
      Width           =   5820
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "grditem_br.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   510
         Left            =   1230
         Picture         =   "grditem_br.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   2340
         Picture         =   "grditem_br.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   510
         Left            =   3690
         TabIndex        =   11
         Top             =   135
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   900
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
         Picture         =   "grditem_br.frx":78EF
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   10125
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4785
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄œœ ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   405
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker xdate1 
         Height          =   375
         Left            =   1710
         TabIndex        =   13
         Top             =   180
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
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
         Format          =   162529281
         CurrentDate     =   43098.0832986111
      End
      Begin MSComCtl2.DTPicker XDATE2 
         Height          =   375
         Left            =   1710
         TabIndex        =   14
         Top             =   585
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
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
         Format          =   162529281
         CurrentDate     =   43098.0832986111
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   2520
      Top             =   405
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
      Left            =   60
      Top             =   120
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
      Left            =   300
      Top             =   120
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
      Left            =   60
      Top             =   120
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
      Height          =   6990
      Left            =   90
      TabIndex        =   9
      Top             =   1080
      Width           =   14775
      _cx             =   26061
      _cy             =   12330
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
      Rows            =   2
      Cols            =   20
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   10
      Top             =   9165
      Visible         =   0   'False
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSDataListLib.DataCombo xbranch 
      Height          =   315
      Left            =   5355
      TabIndex        =   15
      Top             =   45
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777088
      Text            =   ""
      RightToLeft     =   -1  'True
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
   Begin MSAdodcLib.Adodc DATA20 
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
   Begin MSAdodcLib.Adodc data30 
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
   Begin VSFlex7Ctl.VSFlexGrid GRID3 
      Height          =   2445
      Left            =   90
      TabIndex        =   17
      Top             =   8100
      Width           =   6360
      _cx             =   11218
      _cy             =   4313
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   8454016
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
      Rows            =   2
      Cols            =   20
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid GRID2 
      Height          =   2445
      Left            =   8640
      TabIndex        =   18
      Top             =   8100
      Width           =   6270
      _cx             =   11060
      _cy             =   4313
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   8454016
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
      Rows            =   2
      Cols            =   20
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label branchname 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   6525
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   8145
      Width           =   2040
      WordWrap        =   -1  'True
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
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   90
      Width           =   690
   End
End
Attribute VB_Name = "grditem_br"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conShop As New ADODB.Connection
Dim con As New ADODB.Connection
Dim oSearchItem As New Search31
Dim LastSalTable As New ADODB.Recordset
Dim LastImpTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim nRowTop As Double
Dim nRowEnd As Double
Dim nRowBranch As Double
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub cmdPrint_Click()
Dim cHead1 As String
Dim cHead2 As String
cHead1 = "„ «»⁄… «·›—Ê⁄"
If IsDate(xDate1.Value) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "„‰ : " & Format(xDate1.Value, "dd-mm-yyyy")
If IsDate(xdate2.Value) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "Õ Ì : " & Format(xdate2.Value, "dd-mm-yyyy")
'Set printGrd3.myform = Me
Load PrintGrd
PrintGrd.doprint grid1, 1.02, 0, cHead1, cHead2, "", False, True, 9
PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
openCon con
   
Set DATA2.Recordset = myRecordSet("SELECT CODE , DESCA FROM branch  WHERE DATA IS NOT NULL ORDER BY CODE  ", con)
Set xbranch.RowSource = DATA2
xbranch.ListField = "Desca"
xbranch.BoundColumn = "Code"
If cBranch <> "00" Then
    xbranch.BoundText = cBranchStore
    xbranch.Enabled = True
End If
xDate1.Value = DateValue(GetDesca("SELECT MIN(DATE) FROM FILE1_11", con))
xdate2.Value = DateAdd("d", -1, Date)
Set grid1.DataSource = data10
Fixgrd

Set GRID2.DataSource = data20
Set GRID3.DataSource = data30

End Sub
Private Sub myload()
Dim cString As String, cWhere As String
With grid1
If IsDate(xDate1.Value) Then
    cWhere = "DATE < " & DateSq(xDate1.Value)
    cField1 = myiif(cWhere, "[IN] - [OUT]") & " AS FIELD1"
    cField2 = myiif(cWhere, "Round(([IN]- OUT) * FILE1_10.COST,0)") & " AS FIELD2"
Else
    cField1 = "0 AS FIELD1"
    cField2 = "0 AS FIELD2"
End If
cWhere = ""
If IsDate(xDate1.Value) Then cWhere = "DATE >= " & DateSq(xDate1.Value)
If IsDate(xdate2.Value) Then cWhere = cWhere & turn(cWhere, " AND") & " DATE <= " & DateSq(xdate2.Value)

cField3 = myiif(cWhere & turn(cWhere, " AND ") & "(TYPE = '2' )", "[IN] ") & " AS PUR"
cField4 = myiif(cWhere & turn(cWhere, " AND ") & "(TYPE = '7')", "OUT") & " AS PURret"

cField5 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'T'", "[IN]") & " AS TRANS_IN"
cField6 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'T'", "ROUND([IN] * FILE1_10.COST,0)") & " AS TRANS_IN_COST"

cField7 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'F'", "[OUT]") & " AS TRANS_OUT"
cField8 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'F'", "Round([OUT] * FILE1_10.COST,0)") & " AS TRANS_OUT_COST"

cField9 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'Z'", "([IN]+OUT)") & " AS STOCK"
cField10 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'Z'", "Round(([IN]+[OUT]) * FILE1_10.COST,0)") & " AS STOCK_COST"

cField11 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = '6'", "[OUT]") & " AS SALES"
cField12 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = '6'", "Round([OUT] * FILE1_10.COST,0)") & " AS SALES_COST"
cField13 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = '6'", "Round([OUT] * FILE1_11.PRICE,0)") & " AS SALES_PRICE"
cField14 = "'' AS SALES_PRF"
cField15 = myiif("", "([IN]- OUT)") & " AS BAL_LAST"
cField16 = myiif("", "Round(([IN]- OUT) * FILE1_10.COST,0)") & " AS BAL_LAST_COST"
cField17 = "'' AS SALES_RATE"
cField18 = myiif(cWhere & turn(cWhere, " AND ") & "[TYPE] <> '6'", "Round(([IN]-OUT)  * FILE1_10.COST,0)") & " AS TOTAL_IN_COST"

cField = cField1 & "," & cField2 & "," & cField3 & "," & cField4 & "," & cField5 & "," & cField6 & "," & cField7 & _
         "," & cField8 & "," & cField9 & "," & cField10 & "," & cField11 & "," & cField12 & "," & cField13 & "," & cField14 & _
         "," & cField15 & "," & cField16 & "," & cField17 & "," & cField18

cString = "SELECT FILE1_11.STORE,FILE0_40.DESCA," & cField & _
          " FROM   FILE1_11 INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE LEFT OUTER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM  WHERE FILE0_40.ISSTOP = 0 AND FILE0_40.CODE <> '000' "
If IsDate(xdate2.Value) Then cString = cString & turn(cString) & " [DATE] <= " & DateSq(xdate2.Value)
If cBranch <> "00" Then cString = cString & turn(cString) & " STORE  = " & MyParn(cBranchStore)
If cBranch = "00" Then
    If xbranch.BoundText <> "" Then cString = cString & turn(cString) & " file0_40.branch = " & MyParn(xbranch.BoundText)
End If
cString = cString & " GROUP BY FILE1_11.STORE,FILE0_40.DESCA"
cString = cString & " ORDER BY FILE1_11.STORE"
Set data10.Recordset = myRecordSet(cString, con)

nRowTop = 1
nRowEnd = .Rows - 1
nRowBranch = nRowEnd + 2
End With

Fixgrd

If cBranch = "00" Then
    FixGridBranch
Else
    FixGridBranch
End If
With grid1
    For nRow = 2 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) = 0 Then
                .TextMatrix(nRow, nCol) = ""
            Else
                .TextMatrix(nRow, nCol) = Int(Val(.TextMatrix(nRow, nCol)))
            End If
        Next nCol
    Next nRow
    For nRow = 2 To nRowEnd
        cStore = .TextMatrix(nRow, 0)
        For NR = nRowBranch To .Rows - 1
            If .TextMatrix(NR, 0) = cStore Then
                For nCol = 2 To .Cols - 1
                     
                    If Val(.TextMatrix(nRow, nCol)) <> Val(.TextMatrix(NR, nCol)) Then
                        .Cell(flexcpBackColor, nRow, nCol) = vbRed
                        .Cell(flexcpBackColor, NR, nCol) = vbRed
                    End If
                Next nCol
            End If
        Next NR
    Next nRow
End With
Handlecontrols
End Sub
Sub Fixgrd()
Dim nTotal_Sales As Double, nTotal_in As Double
    With grid1
    .ExplorerBar = flexExNone
    .RowHeight(0) = 800
    .FrozenCols = 2
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·›—⁄"
    .TextMatrix(1, 0) = "ﬂÊœ"
    .TextMatrix(1, 1) = "«·›—⁄"

    .TextMatrix(0, 2) = "—’Ìœ «Ê·"
    .TextMatrix(0, 3) = "—’Ìœ «Ê·"
    .TextMatrix(1, 2) = "⁄œœ"
    .TextMatrix(1, 3) = "ﬁÌ„…"
    .ColHidden(3) = Check1.Value
    .TextMatrix(0, 4) = "„‘ —Ì« "
    .TextMatrix(0, 5) = "„—œÊœ „‘ —Ì« "
    .TextMatrix(1, 4) = "⁄œœ"
    .TextMatrix(1, 5) = "ﬁÌ„…"

    .TextMatrix(0, 4 + 2) = "Ê«—œ  ÕÊÌ·« "
    .TextMatrix(0, 5 + 2) = "Ê«—œ  ÕÊÌ·« "
    .TextMatrix(1, 4 + 2) = "⁄œœ"
    .TextMatrix(1, 5 + 2) = "ﬁÌ„…"
    .ColHidden(7) = Check1.Value

    .TextMatrix(0, 6 + 2) = "’«œ—  ÕÊÌ·« "
    .TextMatrix(0, 7 + 2) = "’«œ—  ÕÊÌ·« "
    .TextMatrix(1, 6 + 2) = "⁄œœ"
    .TextMatrix(1, 7 + 2) = "ﬁÌ„…"
    .ColHidden(9) = Check1.Value

    .TextMatrix(0, 8 + 2) = " ”ÊÌ«  Ã—œ"
    .TextMatrix(0, 9 + 2) = " ”ÊÌ«  Ã—œ"
    .TextMatrix(1, 8 + 2) = "⁄œœ"
    .TextMatrix(1, 9 + 2) = "ﬁÌ„…"
    .ColHidden(11) = Check1.Value

    .TextMatrix(0, 10 + 2) = "„»Ì⁄« "
    .TextMatrix(0, 11 + 2) = "„»Ì⁄« "
    .TextMatrix(0, 12 + 2) = "„»Ì⁄« "
    .TextMatrix(0, 13 + 2) = "„»Ì⁄« "
    .TextMatrix(1, 10 + 2) = "⁄œœ"
    .TextMatrix(1, 11 + 2) = "ﬁÌ„…  ﬂ·›…"
    .TextMatrix(1, 12 + 2) = "ﬁÌ„… »Ì⁄"
    .TextMatrix(1, 13 + 2) = "—»Õ"
    
    .ColHidden(13) = Check1.Value
    .ColHidden(14) = Check1.Value
    .ColHidden(15) = Check1.Value

    .TextMatrix(0, 14 + 2) = "—’Ìœ Õ«·Ì"
    .TextMatrix(0, 15 + 2) = "—’Ìœ Õ«·Ì"
    .TextMatrix(1, 14 + 2) = "⁄œœ"
    .TextMatrix(1, 15 + 2) = "ﬁÌ„…"
    .ColHidden(17) = Check1.Value

    .TextMatrix(0, 16 + 2) = "„»Ì⁄« "
    .TextMatrix(1, 16 + 2) = "‰”»…"
    .ColHidden(18) = Check1.Value
    
    '.ColFormat(15) = "#.##%"
    .ColFormat(18) = "#%"
    .ColWidth(1) = 2000
    .ColWidth(2) = 700
    .ColWidth(3) = 900
    .ColWidth(4) = 700
    .ColWidth(5) = 1000
    .ColWidth(6) = 700
    .ColWidth(7) = 1000
    .ColWidth(8) = 700
    .ColWidth(9) = 1000
    .ColWidth(10) = 700
    .ColWidth(11) = 1100
    .ColWidth(12) = 700
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .ColWidth(16) = 800
    .ColWidth(17) = 1000
    .ColWidth(18) = 0
    
    .ColHidden(0) = True
   .ColHidden(.Cols - 1) = True
    For i = 2 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    For i = 2 To .Rows - 1
        .TextMatrix(i, 15) = Val(.TextMatrix(i, 14)) - Val(.TextMatrix(i, 13))
        If Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 19)) <> 0 Then
            .TextMatrix(i, 18) = (Val(.TextMatrix(i, 13)) / (Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 19))))
        End If
    Next i
'    .SubtotalPosition = flexSTAbove
'    .Subtotal flexSTSum, -1, 2, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 3, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 5, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 6, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 7, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 8, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 9, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 10, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 11, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 12, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 13, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 14, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 15, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 16, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 17, "#0", vbRed, vbYellow, True, "  "
'    .Subtotal flexSTSum, -1, 19, "#0", vbRed, vbYellow, True, "  "
'    If cBranch = "00" Then
'        If .Rows > 2 Then
'            .TextMatrix(2, 1) = "«·≈Ã„«·Ï"
'            If Val(.TextMatrix(2, 2)) + Val(.TextMatrix(2, 19)) <> 0 Then
'                .TextMatrix(2, 18) = (Val(.TextMatrix(2, 13)) / (Val(.TextMatrix(2, 2)) + Val(.TextMatrix(2, 19))))
'            End If
'        End If
'        xTotal.Caption = IIf(grid1.Rows > 2, grid1.Rows - 2, "")
'    End If
    If Check1.Value = 1 Then
        For nCol = 2 To .Cols - 1
            .ColWidth(nCol) = 1500
        Next nCol
    End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set VsTItem = Nothing
End Sub
Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xDesca.text, 1
End If
End Sub
Private Sub xitem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    ItemsLookupAll Me, oSearchItem
End If
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xitem.text, 0
End If
End Sub
Sub myProc()
xitem.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
xDesca.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
Unload oSearchItem
End Sub
Private Sub Handlecontrols()
cmdPrint.Enabled = grid1.Rows > 1
End Sub
Sub FixGridBranch()
    Dim BranchTable As New ADODB.Recordset
    If cBranch = "00" Then
        If xbranch.BoundText <> "" Then
            BranchTable.Open "SELECT * FROM BRANCH WHERE ipsql IS NOT NULL AND CODE =  " & MyParn(xbranch.BoundText) & " ORDER BY CODE  ", con, adOpenStatic, adLockReadOnly, adCmdText
        Else
            BranchTable.Open "SELECT * FROM BRANCH WHERE ipsql IS NOT NULL AND CODE <> '00' ORDER BY CODE  ", con, adOpenStatic, adLockReadOnly, adCmdText
        End If
    Else
        BranchTable.Open "SELECT * , IPSERVER AS IPSQL , DATASERVER AS DATA FROM BRANCH WHERE ipserver IS NOT NULL  ORDER BY CODE  ", con, adOpenStatic, adLockReadOnly, adCmdText
    End If
    With BranchTable
    Do While Not .EOF
        cDataShop = !data & ""
        strConShop = LoadConStringshop(!IPsql, !data)
        If openConShop(conShop, strConShop, False) = "ok" Then
            grid1.AddItem ""
            grid1.AddItem ""
            grid1.TextMatrix(grid1.Rows - 1, 0) = !STORE
            grid1.TextMatrix(grid1.Rows - 1, 1) = !Desca
            LoadBalBranch !STORE, conShop
        End If
        .MoveNext
    Loop
    End With
    If grid1.Rows > 1 Then grid1.Cell(flexcpAlignment, 1, 0, grid1.Rows - 1, grid1.Cols - 1) = 4
End Sub
'Sub FixGridBRANCH()
'    Dim BranchTable As New ADODB.Recordset
'    BranchTable.Open "SELECT * FROM BRANCH ", con, adOpenStatic, adLockReadOnly, adCmdText
'    With BranchTable
'    Do While Not .EOF
'        cDataShop = !dataSERVER
'        strConShop = LoadConStringshop(!ipserver, !dataSERVER)
'        If openConShop(conShop, strConShop) = "ok" Then
'            grid1.AddItem ""
'            grid1.AddItem ""
'            grid1.TextMatrix(grid1.Rows - 1, 0) = !store
'            grid1.TextMatrix(grid1.Rows - 1, 1) = !desca
'            LoadBalBranch !store, conShop
'        End If
'        .MoveNext
'    Loop
'    End With
'    If grid1.Rows > 1 Then grid1.Cell(flexcpAlignment, 1, 0, grid1.Rows - 1, grid1.Cols - 1) = 4
'End Sub
Sub LoadBalBranch(pstore, pConShop)
Dim BalStoreTable As New ADODB.Recordset
cWhere = "DATE < " & DateSq(xDate1.Value)
cField1 = myiif(cWhere, "[IN] - [OUT]") & " AS FIELD1"
cWhere = ""
If IsDate(xDate1.Value) Then cWhere = "DATE >= " & DateSq(xDate1.Value)
If IsDate(xdate2.Value) Then cWhere = cWhere & turn(cWhere, " AND") & " DATE <= " & DateSq(xdate2.Value)

cField2 = myiif(cWhere & turn(cWhere, " AND ") & "(TYPE = '2')", "[IN]") & " AS PUR"
cField21 = myiif(cWhere & turn(cWhere, " AND ") & "(TYPE = '7')", "OUT") & " AS PURret"
cField3 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'T'", "[IN]") & " AS TRANS_IN"
cField4 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'F'", "[OUT]") & " AS TRANS_OUT"
cField5 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'Z'", "[out]+[in]") & " AS STOCK"
cField6 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = '6'", "[OUT]") & " AS SALES"
cField7 = myiif("", "([IN]- OUT)") & " AS BAL_LAST"

cField = cField1 & "," & cField2 & "," & cField21 & "," & cField3 & "," & cField4 & "," & cField5 & "," & cField6 & "," & cField7

cString = "SELECT " & cField & _
          " FROM   FILE1_11 INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE LEFT OUTER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM "
If IsDate(xdate2.Value) Then cString = cString & turn(cString) & " [DATE] <= " & DateSq(xdate2.Value)
cString = cString & turn(cString) & " STORE  = " & MyParn(pstore)
BalStoreTable.Open cString, pConShop, adOpenStatic, adLockReadOnly, adCmdText
If BalStoreTable.RecordCount > 0 Then
    With grid1
        .TextMatrix(.Rows - 1, 2) = BalStoreTable!FIELD1 & ""
        .TextMatrix(.Rows - 1, 4) = BalStoreTable!PUR & ""
        .TextMatrix(.Rows - 1, 5) = BalStoreTable!PURret & ""
        .TextMatrix(.Rows - 1, 6) = BalStoreTable!Trans_In & ""
        .TextMatrix(.Rows - 1, 8) = BalStoreTable!TRANS_OUT & ""
        .TextMatrix(.Rows - 1, 10) = BalStoreTable!STOCK & ""
        .TextMatrix(.Rows - 1, 12) = BalStoreTable!sales & ""
        .TextMatrix(.Rows - 1, 16) = BalStoreTable!BAL_LAST & ""
    End With
End If
End Sub

Private Function LoadConStringshop(pServer, pData)
Dim cServerName As String, cUserId As String, cPassword As String
'cUserId = decrypt(RetSetting("userId", App.Path & "\conf.txt"), "dr")
'cPassword = decrypt(RetSetting("Password", App.Path & "\conf.txt"), "dr")

cUserId = "DATA_PRO"
cPassword = "2015"

LoadConStringshop = "provider=SQLOLEDB;data source=" & pServer & ";initial " _
            & "catalog=" & pData & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Private Sub grid1_DblClick()
With grid1
    branchname.Caption = GetDesca("select desca from STORE_BR where code = " & MyParn(.TextMatrix(.Row, 0)), con)
    Select Case .Col
        Case 4
            ViewSubMove 1, .TextMatrix(.Row, 0), xDate1.Value, xdate2.Value
        Case 5
            ViewSubMove 2, .TextMatrix(.Row, 0), xDate1.Value, xdate2.Value
        Case 6
            ViewSubMove 3, .TextMatrix(.Row, 0), xDate1.Value, xdate2.Value
        Case 8
            ViewSubMove 4, .TextMatrix(.Row, 0), xDate1.Value, xdate2.Value
        Case 10
            ViewSubMove 5, .TextMatrix(.Row, 0), xDate1.Value, xdate2.Value
        Case 12
            ViewSubMove 6, .TextMatrix(.Row, 0), xDate1.Value, xdate2.Value
    End Select
End With
End Sub
Sub ViewSubMove(pType, pstore, pDate1, pDate2)
    Dim cIpServer As String, cData As String
    Select Case pType
        Case 1
            cString = "SELECT DOC_ID , DATE , SUM([IN] ) FROM FILE1_11 inner join file1_10 on file1_10.item = file1_11.item WHERE (TYPE = '2' ) AND STORE = " & MyParn(pstore) & " and date >= " & DateSq(xDate1.Value) & " and date <= " & DateSq(xdate2.Value) & _
                        " group by doc_id , date order by date , doc_id "
        Case 2
            cString = "SELECT DOC_ID , DATE , SUM([OUT] ) FROM FILE1_11 inner join file1_10 on file1_10.item = file1_11.item WHERE (TYPE = '7' ) AND STORE = " & MyParn(pstore) & " and date >= " & DateSq(xDate1.Value) & " and date <= " & DateSq(xdate2.Value) & _
                        " group by doc_id , date order by date , doc_id "
        Case 3
            cString = "SELECT DOC_ID , DATE , SUM([IN] + [OUT]) FROM FILE1_11 inner join file1_10 on file1_10.item = file1_11.item  WHERE TYPE = 'T'  AND STORE = " & MyParn(pstore) & " and date >= " & DateSq(xDate1.Value) & " and date <= " & DateSq(xdate2.Value) & _
                        " group by doc_id , date order by date , doc_id "
        Case 4
            cString = "SELECT DOC_ID , DATE , SUM([IN] + [OUT]) FROM FILE1_11 inner join file1_10 on file1_10.item = file1_11.item  WHERE TYPE = 'F'  AND STORE = " & MyParn(pstore) & " and date >= " & DateSq(xDate1.Value) & " and date <= " & DateSq(xdate2.Value) & _
                        " group by doc_id , date order by date , doc_id "
        Case 5
            cString = "SELECT DOC_ID , DATE , SUM([IN] + [OUT]) FROM FILE1_11 LEFT JOIN file1_10 on file1_10.item = file1_11.item  WHERE TYPE = 'Z'  AND STORE = " & MyParn(pstore) & " and date >= " & DateSq(xDate1.Value) & " and date <= " & DateSq(xdate2.Value) & _
                        " group by doc_id , date order by date , doc_id "
        Case 6
            cString = "SELECT DOC_ID , DATE , SUM([IN] + [OUT]) FROM FILE1_11 inner join file1_10 on file1_10.item = file1_11.item  WHERE TYPE = '6'  AND STORE = " & MyParn(pstore) & " and date >= " & DateSq(xDate1.Value) & " and date <= " & DateSq(xdate2.Value) & _
                        " group by doc_id , date order by date , doc_id "
    End Select
    GRID2.Rows = 1
    GRID3.Rows = 1
    
    Set data20.Recordset = myRecordSet(cString, con)
    FixGrid2

    cIpServer = GetDesca("SELECT IPSQL from BRANCH where store = " & MyParn(pstore), con)
    cData = GetDesca("SELECT data from BRANCH where store = " & MyParn(pstore), con)
    strConShop = LoadConStringshop(cIpServer, cData)
    If openConShop(conShop, strConShop) = "ok" Then
        GRID3.Rows = 3
        Set GRID3.DataSource = data30
        data30.ConnectionString = strConShop
        data30.RecordSource = cString
        data30.Refresh
        FixGrid3
    Else
        MsgBox "·« ÌÊÃœ « ’«· »«·›—⁄"
        Exit Sub
    End If
    
    With GRID2
        For nRow = 1 To .Rows - 1
            nFound = GRID3.FindRow(.TextMatrix(nRow, 0), 1, 0)
            If nFound > 0 Then
                If Val(.TextMatrix(nRow, 2)) = Val(GRID3.TextMatrix(nFound, 2)) Then
                    .RowHidden(nRow) = True
                    GRID3.RowHidden(nFound) = True
                Else
                    .TextMatrix(nRow, 3) = Val(GRID3.TextMatrix(nFound, 2))
                    GRID3.TextMatrix(nFound, 3) = Val(.TextMatrix(nRow, 2))
                End If
            End If
        Next nRow
    End With
End Sub
Sub FixGrid2()
    With GRID2
        .Cols = 5
        .TextMatrix(0, 0) = "„” ‰œ"
        .TextMatrix(0, 1) = " «—ÌŒ"
        .TextMatrix(0, 2) = "⁄œœ ”Ì—›—"
        .TextMatrix(0, 3) = "⁄œœ ›—⁄"
        .ColWidth(0) = 1700
        .ColWidth(1) = 1300
        .ColWidth(2) = 1200
        .ColWidth(3) = 1200
        
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, 2) = Round(Val(.TextMatrix(nRow, 2)), 2)
            .TextMatrix(nRow, 3) = Round(Val(.TextMatrix(nRow, 3)), 2)
        Next nRow
        
    End With
End Sub
Sub FixGrid3()
    With GRID3
        GRID3.Cols = 4
        .TextMatrix(0, 0) = "„” ‰œ"
        .TextMatrix(0, 1) = " «—ÌŒ"
        .TextMatrix(0, 2) = "⁄œœ ”Ì—›—"
        .TextMatrix(0, 3) = "⁄œœ ›—⁄"
        .ColWidth(0) = 2000
        .ColWidth(1) = 1400
        .ColWidth(2) = 1300
        .ColWidth(3) = 1300
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, 2) = Round(Val(.TextMatrix(nRow, 2)), 2)
            .TextMatrix(nRow, 3) = Round(Val(.TextMatrix(nRow, 3)), 2)
        Next nRow
    End With
End Sub
Private Sub Grid2_dblClick()
    Dim pstore As String, pDate As Date
    If GRID2.Col = 0 And Check2.Value <> 0 Then
        If grid1.Col = 4 Then
            cDoc = GRID2.TextMatrix(GRID2.Row, 0)
            con.Execute " INSERT INTO NEW7_20H (DOC_NO , BRANCH) values ( " & addstring(cDoc) & " ,  " & addstring(xbranch.BoundText) & ")"
            con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(xbranch.BoundText) & " FROM  FILE1_10  WHERE ITEM IN (SELECT ITEM FROM FILE7_20 WHERE DOC_NO = " & MyParn(cDoc) & " ) ", nRec
        End If
        If grid1.Col = 5 Then
            cDoc = GRID2.TextMatrix(GRID2.Row, 0)
            con.Execute " INSERT INTO NEW7_10H (DOC_NO , BRANCH) values ( " & addstring(cDoc) & " ,  " & addstring(xbranch.BoundText) & ")"
            con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(xbranch.BoundText) & " FROM  FILE1_10  WHERE ITEM IN (SELECT ITEM FROM FILE7_10 WHERE DOC_NO = " & MyParn(cDoc) & " ) ", nRec
        End If
        If grid1.Col = 6 Or grid1.Col = 8 Then
            cDoc = GRID2.TextMatrix(GRID2.Row, 0)
            con.Execute " INSERT INTO NEW1_60H (DOC_NO , BRANCH) values ( " & addstring(cDoc) & " ,  " & addstring(xbranch.BoundText) & ")"
            con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(xbranch.BoundText) & " FROM  FILE1_10  WHERE ITEM IN (SELECT ITEM FROM FILE1_60 WHERE DOC_NO = " & MyParn(cDoc) & " ) ", nRec
        End If
    End If
    
    
    If Not IsDate(grid1.TextMatrix(grid1.Row, 1)) Then Exit Sub
    pstore = grid1.TextMatrix(grid1.Row, 0)
    pDate = DateValue(GRID2.TextMatrix(GRID2.Row, 1))
    If grid1.Col = 10 Then
        If MsgBox(" ÕœÌÀ „»Ì⁄«  «·ÌÊ„ ··›—⁄", vbYesNo + vbDefaultButton2) Then
            cIpServer = GetDesca("SELECT IPSQL from BRANCH where store = " & MyParn(pstore), con)
            cData = GetDesca("SELECT data from BRANCH where store = " & MyParn(pstore), con)
            strConShop = LoadConStringshop(cIpServer, cData)
            If openConShop(conShop, strConShop) = "ok" Then
                conShop.Execute " update file6_20h set isnew = 1 where date = " & DateSq(pDate)
                con.Execute " delete from file6_20  where doc_no in ( select doc_no from file6_20h where store = " & MyParn(pstore) & " and date = " & DateSq(pDate) & " ) "
                con.Execute " delete from file6_20h where store = " & MyParn(pstore) & " and date = " & DateSq(pDate)
                Inform " ÌÃ» ”Õ» „»Ì⁄«  «·ÌÊ„ „—…«Œ—Ï  "
            End If
        End If
    End If
End Sub
Private Sub GRID2_EnterCell()
If GRID2.Col = 4 Then
    GRID2.Editable = flexEDKbdMouse
Else
    GRID2.Editable = flexEDNone
End If
End Sub

Private Sub GRID3_DblClick()
    Dim cDoc As String
    If GRID3.Col = 0 And Check2.Value <> 0 Then
        If grid1.Col = 4 Then
            cDoc = GRID3.TextMatrix(GRID3.Row, 0)
            con.Execute " INSERT INTO NEW7_20H (DOC_NO , BRANCH) values ( " & addstring(cDoc) & " ,  " & addstring(xbranch.BoundText) & ")"
            con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(xbranch.BoundText) & " FROM  FILE1_10  WHERE ITEM IN (SELECT ITEM FROM FILE7_20 WHERE DOC_NO = " & MyParn(cDoc) & " ) ", nRec
        End If
        If grid1.Col = 5 Then
            cDoc = GRID3.TextMatrix(GRID3.Row, 0)
            con.Execute " INSERT INTO NEW7_10H (DOC_NO , BRANCH) values ( " & addstring(cDoc) & " ,  " & addstring(xbranch.BoundText) & ")"
            con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(xbranch.BoundText) & " FROM  FILE1_10  WHERE ITEM IN (SELECT ITEM FROM FILE7_10 WHERE DOC_NO = " & MyParn(cDoc) & " ) ", nRec
        End If
        If grid1.Col = 6 Or grid1.Col = 8 Then
            cDoc = GRID3.TextMatrix(GRID3.Row, 0)
            con.Execute " INSERT INTO NEW1_60H (DOC_NO , BRANCH) values ( " & addstring(cDoc) & " ,  " & addstring(xbranch.BoundText) & ")"
            con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(xbranch.BoundText) & " FROM  FILE1_10  WHERE ITEM IN (SELECT ITEM FROM FILE1_60 WHERE DOC_NO = " & MyParn(cDoc) & " ) ", nRec
        End If
    
    End If
End Sub
