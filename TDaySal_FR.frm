VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form TDaySal_FR 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11280
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar PROG1 
      Height          =   300
      Left            =   135
      TabIndex        =   25
      Top             =   3780
      Width           =   4605
      _ExtentX        =   8123
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Caption         =   "„ﬁ«—‰… »Ê‰«  «·»Ì⁄ ··›—⁄"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   585
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3240
      Width           =   2625
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   90
      Width           =   2535
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   90
         Picture         =   "TDaySal_FR.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         Width           =   1140
      End
      Begin Threed.SSCommand cmdPrint 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1260
         TabIndex        =   14
         Top             =   135
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TDaySal_FR.frx":246C
         Caption         =   "ÿ»«⁄… «·ÌÊ„ "
         Alignment       =   4
         PictureAlignment=   10
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "«·› —…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   7485
      Begin VB.CommandButton cmdGo 
         Height          =   450
         Left            =   90
         Picture         =   "TDaySal_FR.frx":48A6
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   1770
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ :"
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
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
         Left            =   6525
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   315
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3300
      Left            =   4770
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   810
      Width           =   10185
      Begin VSFlex7Ctl.VSFlexGrid GridBox 
         Bindings        =   "TDaySal_FR.frx":6D98
         Height          =   2985
         Left            =   90
         TabIndex        =   2
         Top             =   135
         Width           =   10005
         _cx             =   17648
         _cy             =   5265
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         AllowSelection  =   0   'False
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
         RowHeightMax    =   300
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
         TabBehavior     =   0
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
   Begin MSComCtl2.MonthView xDate 
      Height          =   2310
      Left            =   0
      TabIndex        =   0
      Top             =   810
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12632256
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   162070535
      TitleBackColor  =   12632256
      TrailingForeColor=   12632256
      CurrentDate     =   40954
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   3915
      Top             =   -90
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
      Left            =   1710
      Top             =   -630
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   1710
      Top             =   -1035
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
   Begin TabDlg.SSTab SSTAB 
      Height          =   5745
      Left            =   45
      TabIndex        =   3
      Top             =   4095
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   10134
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "≈Ã„«·Ì »Ê‰«  «·»Ì⁄ "
      TabPicture(0)   =   "TDaySal_FR.frx":6DAB
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Grid1"
      Tab(0).Control(1)=   "PRINT1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "≈Ã„«·Ì „»Ì⁄«  «·«’‰«›"
      TabPicture(1)   =   "TDaySal_FR.frx":6DC7
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid2"
      Tab(1).Control(1)=   "SSCommand1"
      Tab(1).Control(2)=   "Picture1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   " ›’Ì·Ì »Ê‰«  «·»Ì⁄"
      TabPicture(2)   =   "TDaySal_FR.frx":6DE3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture5"
      Tab(2).Control(1)=   "PRINT4"
      Tab(2).Control(2)=   "GRID3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "„»Ì⁄«  »«·”«⁄…"
      TabPicture(3)   =   "TDaySal_FR.frx":6DFF
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "GRIDHOUR"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid GRID3 
         Height          =   5190
         Left            =   -74820
         TabIndex        =   4
         Top             =   405
         Width           =   14640
         _cx             =   25823
         _cy             =   9155
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   5145
         Left            =   -74865
         TabIndex        =   5
         Top             =   405
         Width           =   14685
         _cx             =   25903
         _cy             =   9075
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VSFlex7Ctl.VSFlexGrid Grid1 
         Height          =   5010
         Left            =   -74910
         TabIndex        =   6
         Top             =   405
         Width           =   14775
         _cx             =   26061
         _cy             =   8837
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
      Begin Threed.SSCommand PRINT1 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   -74910
         TabIndex        =   16
         Top             =   6165
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   900
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TDaySal_FR.frx":6E1B
         Alignment       =   4
         PictureAlignment=   10
      End
      Begin Threed.SSCommand PRINT4 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   -74910
         TabIndex        =   17
         Top             =   5895
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   900
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TDaySal_FR.frx":9255
         Alignment       =   4
         PictureAlignment=   10
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   2355
         Left            =   -74955
         Picture         =   "TDaySal_FR.frx":B68F
         RightToLeft     =   -1  'True
         ScaleHeight     =   2355
         ScaleWidth      =   3210
         TabIndex        =   18
         Top             =   3555
         Width           =   3210
         Begin ImgXCtrl6.ImgXCtrl imgx3 
            Height          =   2175
            Left            =   45
            TabIndex        =   19
            Top             =   45
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   3836
            BackColor       =   16777215
            BorderStyle     =   0
            AutoZoom        =   -1  'True
            SelectionLineType=   4
            Center          =   -1  'True
            ImageBorderThickness=   1
            DoubleBuffer    =   -1  'True
         End
      End
      Begin Threed.SSCommand SSCommand1 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   -74910
         TabIndex        =   20
         Top             =   5895
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   900
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TDaySal_FR.frx":ED66
         Alignment       =   4
         PictureAlignment=   10
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2355
         Left            =   -74955
         Picture         =   "TDaySal_FR.frx":111A0
         RightToLeft     =   -1  'True
         ScaleHeight     =   2355
         ScaleWidth      =   3210
         TabIndex        =   21
         Top             =   3510
         Width           =   3210
         Begin ImgXCtrl6.ImgXCtrl Imgx2 
            Height          =   2175
            Left            =   45
            TabIndex        =   22
            Top             =   45
            Width           =   3030
            _ExtentX        =   5345
            _ExtentY        =   3836
            BackColor       =   16777215
            BorderStyle     =   0
            AutoZoom        =   -1  'True
            SelectionLineType=   4
            Center          =   -1  'True
            ImageBorderThickness=   1
            DoubleBuffer    =   -1  'True
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid GRIDHOUR 
         Height          =   5145
         Left            =   3555
         TabIndex        =   23
         Top             =   405
         Width           =   11310
         _cx             =   19950
         _cy             =   9075
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
   Begin MSAdodcLib.Adodc DATABOX 
      Height          =   330
      Left            =   1665
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
   Begin MSAdodcLib.Adodc DATAHOUR 
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
Attribute VB_Name = "TDaySal_FR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conShop As New ADODB.Connection
Dim con As New ADODB.Connection
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
    doprint1
End Sub
Private Sub cmdGo_Click()
    myloadBox
    MYLOAD1 cBranchBox
    MYLOAD2 cBranchBox
    MYLOAD4 cBranchBox
    myloadHour cBranchBox
End Sub

Private Sub Form_Load()
    SSTAB.Tab = 2
    openCon con
    xDate.Value = dSalesDate
    If cUserBox = "" Then cBranchBox = ""
    
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    
    Set GRID2.DataSource = DATA2
    DATA2.ConnectionString = strCon
    
    Set GridBox.DataSource = DATABOX
    DATABOX.ConnectionString = strCon
    
    Set GRIDHOUR.DataSource = DATAHOUR
    DATAHOUR.ConnectionString = strCon
    
    xDate.Visible = bopt3
    Frame1.Visible = bopt3
    
    grid1.Rows = 1
    GRID2.Rows = 1
    GRID3.Rows = 1
    GridBox.Rows = 1
    
    Fixgrd1
    Fixgrd2
    Fixgrd4
    fixGrdBox

    xDate1.text = xDate.Value
    xdate2.text = xDate.Value
    myloadBox
    Check1.Visible = (cBranch = "00")
End Sub
Private Sub Fixgrd2()
With GRID2
.Cols = 8
.TextMatrix(0, 0) = "„’‰⁄"
.TextMatrix(0, 1) = "„Ê—œ"
.TextMatrix(0, 2) = "„Ê”„"
.TextMatrix(0, 3) = "—ﬁ„ „ÊœÌ·"
.TextMatrix(0, 4) = "«·’‰›"
.TextMatrix(0, 5) = "⁄œœ „»Ì⁄« "
.TextMatrix(0, 6) = "ﬁÌ„… „»Ì⁄« "
.FixedCols = 0
.FixedRows = 1
.ColHidden(.Cols - 1) = True

For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 3) = DelZero(.TextMatrix(nRow, 3))
Next nRow

.ColWidth(0) = 2000
.ColWidth(1) = 2000
.ColWidth(2) = 700
.ColWidth(3) = 2000
.ColWidth(4) = 5000
.ColWidth(5) = 800
.ColWidth(6) = 1500

.MergeCells = flexMergeFree
.MergeCol(0) = True
.MergeCol(1) = True
.MergeCol(2) = True

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTBelow
For nCol = 5 To .Cols - 1
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
    .Subtotal flexSTSum, 0, nCol, "#0.00", vbYellow, , True, " "
Next nCol
End With
End Sub
Private Sub Fixgrd1()

With grid1
.Cols = 16
.TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "«·Êﬁ "
.TextMatrix(0, 3) = "«·Œ“‰…"
.TextMatrix(0, 4) = "⁄œœ «·ﬁÿ⁄"
.TextMatrix(0, 5) = "≈Ã„«·Ï «·ﬁÌ„…"
.TextMatrix(0, 6) = "«·Œ’„"
.TextMatrix(0, 7) = "«·’«›Ï"
.TextMatrix(0, 8) = "”œ«œ ‰ﬁœÏ"
.TextMatrix(0, 9) = "”œ«œ ›Ì“«"
.TextMatrix(0, 10) = "«·›—ﬁ"

.TextMatrix(0, 11) = "’·«ÕÌ… Œ’„"
.TextMatrix(0, 12) = "’·«ÕÌ… „— Ã⁄"
.TextMatrix(0, 13) = "„” ‰œ »Ì⁄"

.TextMatrix(0, 14) = "≈Ã„«·Ï ›—⁄"


.ColFormat(1) = "DD-MM-YYYY"
.ColFormat(2) = "SHORT TIME"
.ColFormat(4) = "#0"
.ColFormat(5) = "#0.00"
.ColFormat(6) = "#0.00"
.ColFormat(7) = "#0.00"
.ColFormat(8) = "#0.00"
.ColFormat(9) = "#0.00"
.ColFormat(10) = "#0.00"
.ColDataType(1) = flexDTDate
.ColDataType(2) = flexDTShort
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(8) = flexDTDouble
.ColDataType(9) = flexDTDouble
.ColDataType(10) = flexDTDouble
.ColDataType(14) = flexDTDouble
.ColHidden(14) = True
.ColHidden(15) = True
.ColWidth(0) = 1200
.ColWidth(1) = 1300
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.ColWidth(4) = 1100
.ColWidth(5) = 1100
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(8) = 1100
.ColWidth(9) = 1100

.ColWidth(10) = 1500
.ColWidth(11) = 1500
.ColWidth(12) = 2200

.ColWidth(14) = 1200

.FixedCols = 0
.FixedRows = 1


For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 10) = Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)) - Val(.TextMatrix(nRow, 9))
    .TextMatrix(nRow, 1) = DateValue(.TextMatrix(nRow, 1))
    .TextMatrix(nRow, 2) = TimeValue(.TextMatrix(nRow, 2))
Next nRow

.ExplorerBar = flexExSortShow
.Subtotal flexSTSum, -1, 4, "#0", vbYellow, , True, " "
For nCol = 5 To 9
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
Next nCol
End With
End Sub
Private Sub fixGrdBox()
With GridBox
.FixedCols = 0
.FixedRows = 1
.Cols = 11
If cBranch = "00" Then
    .TextMatrix(0, 0) = "«·›—⁄"
    .TextMatrix(0, 1) = "«· «—ÌŒ"
    .TextMatrix(0, 2) = "«·›—⁄"
Else
    .TextMatrix(0, 0) = "«·Œ“‰…"
    .TextMatrix(0, 1) = "«· «—ÌŒ"
    .TextMatrix(0, 2) = "«·›—⁄"
End If
.TextMatrix(0, 3) = "⁄œœ »Ê‰« "
.TextMatrix(0, 4) = "⁄œœ ﬁÿ⁄ „»Ì⁄« "
.TextMatrix(0, 5) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 6) = "«·Œ’„"
.TextMatrix(0, 7) = "«·’«›Ï"
.TextMatrix(0, 8) = "‰ﬁœÏ"
.TextMatrix(0, 9) = "›Ì“«"
.TextMatrix(0, 10) = "«·›—ﬁ"
.RowHeight(0) = 800
.WordWrap = True
.ColWidth(1) = 1200
.ColWidth(2) = 1400
.ColWidth(3) = 600
.ColWidth(4) = 600
.ColWidth(5) = 1100
.ColWidth(6) = 900
.ColWidth(7) = 1100
.ColWidth(8) = 1100
.ColWidth(9) = 1100
.ColWidth(10) = 700
.ColHidden(0) = True

For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 10) = Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)) - Val(.TextMatrix(nRow, 9))
    For nCol = 5 To .Cols - 1
        .TextMatrix(nRow, nCol) = Format(Val(.TextMatrix(nRow, nCol)), "#0.00")
    Next nCol
Next nRow
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTAbove
For nCol = 3 To .Cols - 1
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
Next nCol
.Subtotal flexSTSum, -1, 3, "#0", vbYellow, , True, " "
.Subtotal flexSTSum, -1, 4, "#0", vbYellow, , True, " "
End With

End Sub

Private Sub grid2_EnterCell()
MyLoadPhoto Mid(GRID2.TextMatrix(GRID2.Row, GRID2.Cols - 1), 1, 10), Imgx2

End Sub

Private Sub GRID3_EnterCell()
MyLoadPhoto Mid(GRID3.TextMatrix(GRID3.Row, 3), 1, 10), imgx3
End Sub

Private Sub gridbox_DblClick()
If grid1.Rows > 1 Then
    If cBranch = "00" Then
        MYLOAD1 "", GridBox.TextMatrix(GridBox.Row, 0)
        MYLOAD2 "", GridBox.TextMatrix(GridBox.Row, 0)
        MYLOAD4 "", GridBox.TextMatrix(GridBox.Row, 0)
        myloadHour GridBox.TextMatrix(GridBox.Row, 0)
    Else
        MYLOAD1 GridBox.TextMatrix(GridBox.Row, 0)
        MYLOAD2 GridBox.TextMatrix(GridBox.Row, 0)
        MYLOAD4 GridBox.TextMatrix(GridBox.Row, 0)
        myloadHour GridBox.TextMatrix(GridBox.Row, 0)
    End If
End If
End Sub

Private Sub Imgx2_DblClick()
If validPhoto(retFile(Mid(GRID2.TextMatrix(GRID2.Row, GRID2.Cols - 1), 1, 10))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(Mid(GRID2.TextMatrix(GRID2.Row, GRID2.Cols - 1), 1, 10))
    oMaxPhoto.cDesca = GRID2.TextMatrix(GRID2.Row, 4)
    oMaxPhoto.Show 1
End If

End Sub

Private Sub imgx3_DblClick()
If validPhoto(retFile(Mid(GRID3.TextMatrix(GRID3.Row, 3), 1, 10))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(Mid(GRID3.TextMatrix(GRID3.Row, 3), 1, 10))
    oMaxPhoto.cDesca = GRID3.TextMatrix(GRID3.Row, 6)
    oMaxPhoto.Show 1
End If
End Sub

Private Sub PRINT1_Click()
Load PrintGrd
PrintGrd.doprint grid1, , , "≈Ã„«·Ï »Ê‰«  «·»Ì⁄ ··› —… „‰ " & xDate1.text & " Õ Ï " & xdate2.text, , , , True
PrintGrd.Show 1
End Sub

Private Sub print2_Click()
Load PrintGrd
PrintGrd.doprint GRID2, , , "≈Ã„«·Ï „»Ì⁄«  «·„ÊœÌ·«  ··› —… „‰ " & xDate1.text & " Õ Ï " & xdate2.text, , , , True
PrintGrd.Show 1
End Sub
Private Sub PRINT4_Click()
    PrintGrd.doprint GRID3, , , " ›’Ì·Ì »Ê‰«  «·»Ì⁄ " & xDate1.text & " Õ Ï " & xdate2.text, , , , True
    PrintGrd.Show 1
End Sub
Private Sub xdate_DateClick(ByVal DateClicked As Date)
    xDate1.text = Format(xDate.Value, "DD-MM-YYYY")
    xdate2.text = Format(xDate.Value, "DD-MM-YYYY")
    myloadBox
End Sub
Sub myloadBox()
cStr1 = "SELECT t_Salesdoc_FR.STORE , date , STORENAME , COUNT(DOC_NO) , SUM(T_QUANT) , SUM(T_TOTAL) , SUM(discount) , SUM(T_TOTAL-discount), SUM(CASH) , SUM(VISA) FROM  t_Salesdoc_FR"
cStr1 = cStr1 & turn(cStr1) & "DATE >= " & DateSq(xDate1.text)
cStr1 = cStr1 & turn(cStr1) & "DATE <= " & DateSq(xdate2.text)
If Not bOpt5 Then cStr1 = cStr1 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "

cStr1 = cStr1 & " GROUP BY date,STORENAME , STORE ORDER BY STORE "
DATABOX.RecordSource = cStr1
DATABOX.Refresh

fixGrdBox

MYLOAD2 cBranchBox
MYLOAD1 cBranchBox
MYLOAD4 cBranchBox
myloadHour cBranchBox

End Sub
Sub MYLOAD2(Optional cBox As String, Optional cStore As String)
    
    cString = "SELECT  FACTDESCA, SUPPDESCA,FILE1_10.MOSM , FILE1_10.MODELFACT, FILE1_10.desca, SUM(QUANT) AS T_QTY, SUM(TOTAL) AS T_TOTAL,FILE1_10.MODELNO FROM SALES_MODEL_FR INNER JOIN FILE1_10 ON SALES_MODEL_FR.ITEM = FILE1_10.ITEM"
    If IsDate(xDate1.text) Then cString = cString & turn(cString) & "DATE >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cString = cString & turn(cString) & "DATE <= " & DateSq(xdate2.text)
    If cBox <> "" Then cString = cString & turn(cString) & "SALES_MODEL_FR.BOX = " & MyParn(cBox)
    If cStore <> "" Then cString = cString & turn(cString) & "SALES_MODEL_FR.STORE = " & MyParn(cStore)
    If Not bOpt5 Then cString = cString & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    cString = cString & " GROUP BY FILE1_10.MOSM, SUPPDESCA, FACTDESCA, FILE1_10.MODELFACT, FILE1_10.desca,MODELNO order by SUPPDESCA, FACTDESCA , FILE1_10.MOSM,FILE1_10.MODELFACT"
    DATA2.RecordSource = cString
    DATA2.Refresh
    Fixgrd2

End Sub
Sub MYLOAD1(Optional cBox As String, Optional cStore As String)
'    aInsert(20, 0) = "username_disc"
'    aInsert(21, 0) = "username_ret"
'    aInsert(22, 0) = "SALES_RET"

'                           0                   1               2               3                   4                   5                   6                   7                                           8           9               10              11                      12              13
    cString = "SELECT FR6_20H.DOC_NO2 ,FR6_20H.date ,FR6_20H.TIME, FILE0_50FR.DESCA , SUM(FR6_20.QUANT) , SUM(FR6_20.TOTAL) ,FR6_20H.discount , SUM(FR6_20.TOTAL) - FR6_20H.discount , FR6_20H.CASH ,FR6_20H.VISA , ' ' AS N1 , FR6_20H.username_disc , username_ret , SALES_RET , ' ' as n14 , FR6_20H.DOC_NO" & _
    " FROM (FR6_20H INNER JOIN FR6_20 ON FR6_20H.DOC_NO = FR6_20.DOC_NO) LEFT JOIN FILE0_50FR ON FR6_20H.BOX = FILE0_50FR.CODE "
    If IsDate(xDate1.text) Then cString = cString & turn(cString) & "DATE >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cString = cString & turn(cString) & "DATE <= " & DateSq(xdate2.text)
    If cBox <> "" Then cString = cString & turn(cString) & "FR6_20H.BOX = " & MyParn(cBox)
    If cStore <> "" Then cString = cString & turn(cString) & "FR6_20H.STORE = " & MyParn(cStore)
    If Not bOpt5 Then cString = cString & " AND FR6_20H.[store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    cString = cString & " GROUP BY " & "FR6_20H.DOC_NO,FR6_20H.DOC_NO2,FR6_20H.date,FR6_20H.TIME,FILE0_50FR.DESCA,FR6_20H.DISCOUNT,FR6_20H.CASH,FR6_20H.VISA , FR6_20H.username_disc , username_ret , SALES_RET , FR6_20H.STORE "
    cString = cString & " ORDER BY DATE , TIME,FR6_20H.DOC_NO2"
    data1.RecordSource = cString
    data1.Refresh
    
    Fixgrd1
    With grid1
        If cStore <> "" And Check1.Value <> 0 Then
            IpShop = GetDesca("select ipsql from branch where store = " & MyParn(cStore), con)
            cDataShop = GetDesca("select DATA from branch where store = " & MyParn(cStore), con)
            strConShop = LoadConStringshop
            If openConShop(conShop, strConShop) = "ok" Then
                prog1.Visible = True
                prog1.Min = 0
                prog1.Max = .Rows - 1
                prog1.Value = 0
                .ColHidden(14) = False
                For nRow = 2 To .Rows - 1
                    prog1.Value = nRow
                    .TextMatrix(nRow, 14) = Val(GetDesca("SELECT SUM(T_TOTAL-DISCOUNT) FROM t_Salesdoc_FR WHERE DOC_NO = " & MyParn(.TextMatrix(nRow, 15)), conShop) & "")
                    If Val(.TextMatrix(nRow, 14)) <> Val(.TextMatrix(nRow, 7)) Then
                        .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = &HFF&
                    End If
                Next nRow
            Else
                .ColHidden(14) = True
                Inform "·« ÌÊÃœ « ’«· »«·›—⁄"
            End If
        End If
    End With
End Sub
Private Sub doprint1()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(1)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT  SUPPDESCA, FACTDESCA, SUM(QUANT) AS TQUANT, SUM(TOTAL) AS TTOTAL, BOX FROM         SALES_MODEL_FR "
If IsDate(xDate1.text) Then
    cString = cString & turn(cString) & " date >= " & DateSq(xDate1.text)
    aHeader(0) = "[" & BetweenString(Format(xDate1.text, "d-m-yyyy"), Format(xdate2.text, "d-m-yyyy")) & "]"
End If

If IsDate(xdate2.text) Then
    cString = cString & turn(cString) & " date <= " & DateSq(xdate2.text)
     aHeader(0) = "[" & BetweenString(Format(xDate1.text, "d-m-yyyy"), Format(xdate2.text, "d-m-yyyy")) & "]"
End If

If cBox <> "" Then
    cString = cString & turnFound(cString) & " AND BOX = " & MyParn(cBox)
    aHeader(1) = "[" & SalesFrm.xBox.text & "]"
End If
    
cString = cString & " GROUP BY SUPPDESCA, FACTDESCA, BOX "
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.Value
    temptable!str1 = !BOX
    temptable!str2 = GetDesca("SELECT DESCA FROM FILE0_50FR WHERE CODE = " & MyParn(!BOX), con)
    temptable!str5 = !SUPPDESCA
    If !SUPPDESCA <> !FACTDESCA Then temptable!str6 = !FACTDESCA
    temptable!VAL1 = !TQUANT
    temptable!val3 = !TTOTAL
    temptable.Update
    sourcetable.MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\TDAY.RPT"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
End Sub
Private Sub MYLOAD4(Optional cBox As String, Optional cStore As String)
Dim cWhere As String, cString As String
If IsDate(xDate1.text) Then cWhere = " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xdate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xdate2.text)
If cBox <> "" Then cWhere = cWhere & " AND  BOX = " & MyParn(cBox)
If cStore <> "" Then cWhere = cWhere & " AND  STORE = " & MyParn(cStore)
If Not bOpt5 Then cWhere = cWhere & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "

cString = "SELECT SALESDTL_FR.* " & _
          " FROM SALESDTL_FR WHERE DOC_NO IS NOT NULL "
If cWhere <> "" Then cString = cString & cWhere
cString = cString & " ORDER BY DOC_NO,FLAG"
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
With GRID3
GRID3.Rows = 1
Do Until loctable.EOF
    .AddItem ""
    If loctable!Flag = 0 Then
        .TextMatrix(.Rows - 1, 0) = Mid(loctable!doc_no, 7, 5)
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Time, "HH:NN")
        .TextMatrix(.Rows - 1, 2) = loctable!MANDESCA & ""
        .TextMatrix(.Rows - 1, 3) = Format(loctable!modelfact0 & "")
        .TextMatrix(.Rows - 1, 4) = loctable!fact_Desca & ""
        .TextMatrix(.Rows - 1, 5) = loctable!Item & ""
        .TextMatrix(.Rows - 1, 6) = loctable!DESCA & ""
        If loctable!Quant > 0 Then
            .TextMatrix(.Rows - 1, 7) = loctable!Quant
        Else
            .TextMatrix(.Rows - 1, 8) = loctable!Quant * -1
        End If
        .TextMatrix(.Rows - 1, 9) = loctable!price & ""
        .TextMatrix(.Rows - 1, 10) = Format(Val(loctable!TOTAL & ""), "Fixed")
        .TextMatrix(.Rows - 1, 11) = loctable!PRICE_C & ""
        
        
        
        If loctable!price <> loctable!PRICE_C Then
            .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
            .Cell(flexcpBackColor, .Rows - 1, 11, .Rows - 1, 11) = vbYellow
        End If
        nTotal = nTotal + Val(loctable!TOTAL & "")
    ElseIf loctable!Flag = 1 Then
        .TextMatrix(.Rows - 1, 0) = DelZero(Mid(loctable!doc_no, 7, 5))
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = loctable!DESCA
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HD3BD78
        
'        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
'        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalDiscount = nTotalDiscount + Val(.TextMatrix(.Rows - 1, 8 + 1))
        If loctable!Flag = 0 Then nTotalSalesDis = nTotalSalesDis + Val(.TextMatrix(.Rows - 1, 8 + 1))
    ElseIf loctable!Flag = 3 Then
        .TextMatrix(.Rows - 1, 0) = DelZero(Mid(loctable!doc_no, 7, 5))
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = "«·«Ã„«·Ì"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
    ElseIf loctable!Flag = 4 Then
        .TextMatrix(.Rows - 1, 0) = DelZero(Mid(loctable!doc_no, 7, 5))
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = "‰ﬁœÌ…"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = Format(loctable!TOTAL, "#0.00")
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalCash = nTotalCash + Val(loctable!TOTAL & "")
    ElseIf loctable!Flag = 5 Then
        '.TextMatrix(.Rows - 1, 0) = LOCTABLE!doc_no
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = "›Ì“«"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = Format(loctable!TOTAL, "#0.00")
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalVisa = nTotalVisa + Val(loctable!TOTAL & "")
    End If
    .TextMatrix(.Rows - 1, 12) = loctable!username_disc & ""
    .TextMatrix(.Rows - 1, 13) = loctable!username_ret & ""
    .TextMatrix(.Rows - 1, 14) = loctable!SALES_RET & ""
    loctable.MoveNext
Loop

If nTotal <> 0 Then
    .AddItem ""
    For i = 0 To 8
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·ÌÊ„"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = nTotal - nTotalDiscount
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If nTotalCash <> 0 Then
    .AddItem ""
    For i = 0 To 8
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·‰ﬁœÌ…"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = nTotalCash
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If nTotalVisa <> 0 Then
    .AddItem ""
    For i = 0 To 7 + 1
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·›Ì“«"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = nTotalVisa
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If Round((nTotal - nTotalDiscount) - (nTotalCash + nTotalVisa), 0) <> 0 Then
    .AddItem ""
    For i = 0 To 8
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·¬Ã·"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = Round((nTotal - nTotalDiscount) - (nTotalCash + nTotalVisa), 2)
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If
End With
Fixgrd4
End Sub
Private Sub Fixgrd4()
With GRID3
.Cols = 15
.TextMatrix(0, 0) = "—ﬁ„ „” ‰œ"
.TextMatrix(0, 1) = "«·Êﬁ "
.TextMatrix(0, 2) = "«·»«∆⁄"
.TextMatrix(0, 3) = "«·„ÊœÌ·"
.TextMatrix(0, 4) = "«·„’‰⁄"
.TextMatrix(0, 5) = "»«—ﬂÊœ"
.TextMatrix(0, 6) = "«·’‰›"
.TextMatrix(0, 7) = "„»Ì⁄« "
.TextMatrix(0, 8) = "„— Ã⁄« "
.TextMatrix(0, 9) = "«·”⁄—"
.TextMatrix(0, 10) = "«·≈Ã„«·Ï"
.TextMatrix(0, 11) = "”⁄— „” Â·ﬂ"

.TextMatrix(0, 12) = "’·«ÕÌ… Œ’„"
.TextMatrix(0, 13) = "’·«ÕÌ… „— Ã⁄"
.TextMatrix(0, 14) = "„” ‰œ »Ì⁄"


.MergeCells = flexMergeFree
.MergeCol(0) = True

.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 1200
.ColWidth(2) = 800
.ColWidth(3) = 1800
.ColWidth(4) = 1000
.ColWidth(5) = 1200
.ColWidth(6) = 3000
.ColWidth(7) = 800
.ColWidth(8) = 1100
.ColWidth(9) = 1100
.ColWidth(10) = 1100
.ColWidth(11) = 1100

.ColWidth(12) = 1300
.ColWidth(13) = 1300
.ColWidth(14) = 1800

.RowHeight(0) = 700
.WordWrap = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
End With
End Sub
Sub myloadHour(Optional cBox As String, Optional cStore As String)
cStr1 = "SELECT { fn HOUR(TIME) } , COUNT(DOC_NO) , SUM(T_QUANT) , SUM(T_TOTAL) , SUM(discount) , SUM(T_TOTAL-discount), SUM(CASH) , SUM(VISA) FROM  t_Salesdoc_FR"
cStr1 = cStr1 & turn(cStr1) & "DATE >= " & DateSq(xDate1.text)
cStr1 = cStr1 & turn(cStr1) & "DATE <= " & DateSq(xdate2.text)
If Not bOpt5 Then cStr1 = cStr1 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
'If cBox <> "" Then cStr1 = cStr1 & turn(cStr1) & "box = " & MyParn(cBox)
If cStore <> "" Then cStr1 = cStr1 & turn(cStr1) & "STORE = " & MyParn(cStore)
cStr1 = cStr1 & " GROUP BY { fn HOUR(TIME) } ORDER BY { fn HOUR(TIME) }  "
DATAHOUR.RecordSource = cStr1
DATAHOUR.Refresh
fixGrdHour
End Sub
Private Sub fixGrdHour()
With GRIDHOUR
.FixedCols = 0
.FixedRows = 1
.Cols = 9
.TextMatrix(0, 0) = "«·”«⁄…"
.TextMatrix(0, 1) = "⁄œœ »Ê‰« "
.TextMatrix(0, 2) = "⁄œœ ﬁÿ⁄ „»Ì⁄« "
.TextMatrix(0, 3) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 4) = "«·Œ’„"
.TextMatrix(0, 5) = "«·’«›Ï"
.TextMatrix(0, 6) = "‰ﬁœÏ"
.TextMatrix(0, 7) = "›Ì“«"
.TextMatrix(0, 8) = "«·›—ﬁ"
.RowHeight(0) = 800
.WordWrap = True
.ColWidth(1) = 1000
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1000

For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 8) = Val(.TextMatrix(nRow, 5)) - Val(.TextMatrix(nRow, 6)) - Val(.TextMatrix(nRow, 7))
Next nRow
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTAbove
For nCol = 1 To .Cols - 1
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
Next nCol
End With
End Sub
