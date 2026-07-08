VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TDaySal2 
   ClientHeight    =   8520
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
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
   ScaleHeight     =   8520
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComCtl2.MonthView xDate 
      Height          =   2370
      Left            =   6525
      TabIndex        =   30
      Top             =   90
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   6470346
      Appearance      =   1
      StartOfWeek     =   71041031
      TitleBackColor  =   3778241
      TrailingForeColor=   3778241
      CurrentDate     =   40954
   End
   Begin VB.CommandButton cmdCloseday 
      Caption         =   "«€·«ﬁ «·ÌÊ„"
      Height          =   330
      Left            =   11430
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4320
      Width           =   3750
   End
   Begin VSPrinter7LibCtl.VSPrinter vp 
      Height          =   1065
      Left            =   14085
      TabIndex        =   4
      Top             =   5355
      Visible         =   0   'False
      Width           =   960
      _cx             =   1693
      _cy             =   1879
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Traditional Arabic"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   0   'False
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   100
      MarginTop       =   100
      MarginRight     =   100
      MarginBottom    =   100
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   200
      ZoomMode        =   0
      ZoomMax         =   400
      ZoomMin         =   200
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   1
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483626
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   135
      TabIndex        =   0
      Top             =   2565
      Width           =   11145
      _ExtentX        =   19659
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "≈Ã„«·Ì »Ê‰«  «·»Ì⁄ "
      TabPicture(0)   =   "TDaySal2.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Grid2"
      Tab(0).Control(1)=   "Command1"
      Tab(0).Control(2)=   "Command4(0)"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "≈Ã„«·Ì „»Ì⁄«  «·«’‰«›"
      TabPicture(1)   =   "TDaySal2.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Grid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Command2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command4(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Œ—ÊÃ"
         Height          =   375
         Index           =   1
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   7695
         Width           =   1320
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Œ—ÊÃ"
         Height          =   375
         Index           =   0
         Left            =   -74910
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   7695
         Width           =   1320
      End
      Begin VB.CommandButton Command2 
         Caption         =   "ÿ»«⁄…"
         Height          =   375
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   7695
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ÿ»«⁄…"
         Height          =   375
         Left            =   -73560
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   7695
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VSFlex7LCtl.VSFlexGrid Grid2 
         Height          =   7260
         Left            =   -74910
         TabIndex        =   1
         Top             =   405
         Width           =   10995
         _cx             =   19394
         _cy             =   12806
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         Rows            =   50
         Cols            =   10
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
         AutoResize      =   -1  'True
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
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7LCtl.VSFlexGrid Grid1 
         Height          =   7260
         Left            =   855
         TabIndex        =   2
         Top             =   405
         Width           =   10140
         _cx             =   17886
         _cy             =   12806
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         Rows            =   50
         Cols            =   10
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
         AutoResize      =   -1  'True
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
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin Threed.SSPanel SSPanel1 
      Height          =   4200
      Left            =   11430
      TabIndex        =   3
      Top             =   90
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   7408
      _Version        =   196610
      BackColor       =   13822956
      Windowless      =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "’«›Ì „»Ì⁄« "
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   3780
         Width           =   1950
      End
      Begin VB.Label xSumofTotalNet 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   3780
         Width           =   1410
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≈Ã„«·Ì Œ’„ "
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   3375
         Width           =   1950
      End
      Begin VB.Label xSumofTotalDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   3375
         Width           =   1410
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≈Ã„«·Ì „»Ì⁄« "
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   2970
         Width           =   1950
      End
      Begin VB.Label xsumofTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   2970
         Width           =   1410
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≈Ã„«·Ì ’«›Ì ¬Ã·"
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2565
         Width           =   1950
      End
      Begin VB.Label xsumofDelayNet 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2565
         Width           =   1410
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≈Ã„«·Ì Œ’„ ¬Ã·"
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   2160
         Width           =   1950
      End
      Begin VB.Label xsumofDelayDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2160
         Width           =   1410
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≈Ã„«·Ì Œ’„ ‰ﬁœÌ"
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   1905
      End
      Begin VB.Label xSumofCashDiscount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   945
         Width           =   1410
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≈Ã„«·Ì „»Ì⁄«  ‰ﬁœÌ "
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   540
         Width           =   1905
      End
      Begin VB.Label xsumofCash 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   540
         Width           =   1410
      End
      Begin VB.Label xsumofDelay 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1755
         Width           =   1410
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≈Ã„«·Ì „»Ì⁄«  √Ã·"
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1755
         Width           =   1950
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≈Ã„«·Ì ⁄œœ «·»Ê‰«  "
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   135
         Width           =   1905
      End
      Begin VB.Label xSalesCount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   135
         Width           =   1410
      End
      Begin VB.Label xsumofCashNet 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1350
         Width           =   1410
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "≈Ã„«·Ì ’«›Ì ‰ﬁœÌ"
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1350
         Width           =   1950
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   225
      Top             =   135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   270
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   270
      Top             =   990
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
      Left            =   270
      Top             =   585
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
Attribute VB_Name = "TDaySal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bShowPrf As Boolean
Dim con As New ADODB.Connection
Dim TSAlTable As Recordset
Dim CountTable As Recordset
Dim SalTable As Recordset
Dim Sal2Table As Recordset
Dim temptable As Recordset
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmd_close_Click()
'    cString = " UPDATE FILE6_20 SET FILE6_20.CLOSED = TRUE  Where Date = " & DATESQ(xDate.Value)
'    mydb.Execute cString
End Sub
Private Sub myload()
Dim bNew As Boolean, DDate1 As String, DDate2 As String
If IsDate(xDate.Value) Then cwhere = cwhere & turn(cwhere, " AND ") & "FILE6_20H.DATE = " & DateSq(xDate.Value)
cField1 = "Select SUM(FILE6_20H.DISCOUNT) FROM FILE6_20H" & turn(cwhere) & cwhere
cField1 = "(" & cField1 & ")" & " as SumofDiscount"
cField2 = myiif("FILE6_20.Quant ", "QUANT") & " AS SalesQuant"
cField3 = myiif("FILE6_20.Quant ", "QUANT * FILE6_20.PRICE") & " AS SalesValue"

cString = "Select " & cField1 & "," & cField2 & "," & cField3 & _
          ", FILE6_20.ITEM,FILE1_10.DESCA,SUM(FILE6_20.QUANT) AS salesQuantNet,SUM(FILE6_20.QUANT * FILE6_20.PRICE) AS SalesValueNet FROM (FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO) INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM"
cString = cString & turn(cwhere) & cwhere
cString = cString & " group by FILE6_20.ITEM,FILE1_10.DESCA"


Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText

Dim nTotalValue As Double
With grid1
.Rows = 1
Do Until sourcetable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = sourcetable!Item
    .TextMatrix(.Rows - 1, 1) = sourcetable!DESCA & ""
    
    .TextMatrix(.Rows - 1, 2) = Format(sourcetable!SalesQuant, "Fixed")
    .TextMatrix(.Rows - 1, 3) = Format(sourcetable!SalesValue, "fixed")
    nTotalValue = Val(sourcetable!SalesValueNet & "") + nTotalValue
    nDiscount = sourcetable!sumOfDiscount
    sourcetable.MoveNext
Loop

If nDiscount <> 0 Then
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = ""
    .TextMatrix(.Rows - 1, 1) = "«·Œ’„"
    .TextMatrix(.Rows - 1, 2) = ""
    .TextMatrix(.Rows - 1, 7) = -1 * nDiscount
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFFF
    nTotalValue = nTotalValue - nDiscount
End If

If Not (sourcetable.EOF And sourcetable.BOF) Then
    .AddItem ""
    .TextMatrix(.Rows - 1, 1) = "«·«Ã„«·Ì"
    .TextMatrix(.Rows - 1, 6) = nTotalQuant
    .TextMatrix(.Rows - 1, 7) = nTotalValue
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
    .Cell(flexcpForeColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbBlue
End If
End With
myload2
MyLoadTotal
End Sub
Private Sub myload2()
Dim bNew As Boolean, DDate1 As String, DDate2 As String
If IsDate(xDate.Value) Then cwhere = cwhere & turn(cwhere, " AND ") & "FILE6_20H.DATE = " & DateSq(Format(xDate.Value, "dd-mm-yyyy"))

cField1 = myiif("FILE6_20.Quant > 0", "QUANT * FILE6_20.PRICE") & " AS SalesValue"
cField2 = myiif("FILE6_20.Quant < 0", "-1 * QUANT * FILE6_20.PRICE") & " AS SalesValueRet"
cField3 = "SUM(FILE6_20.QUANT * FILE6_20.PRICE) - FILE6_20H.DISCOUNT AS salesValueNet"

cString = "Select " & cField1 & "," & cField2 & "," & cField3 & _
          ", FILE6_20H.DOC_NO,FILE6_20H.DISCOUNT,CONVERT(VARCHAR(5),FILE6_20H.[TIME],108) as mytime,FILE6_20H.DISCOUNT,FILE6_25.DESCA AS MANDESCA" & _
          " FROM (FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO) LEFT JOIN FILE6_25 ON FILE6_20H.MAN = FILE6_25.CODE"

If cwhere <> "" Then
    cString = cString & turn(cString) & cwhere
End If
cString = cString & " group by  FILE6_20H.DOC_NO,FILE6_20H.DISCOUNT,CONVERT(VARCHAR(5),FILE6_20H.[TIME],108),FILE6_25.DESCA"
cString = cString & " Order by FILE6_20H.DOC_NO"

Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText

Dim nTotalValue As Double
With grid2
.Rows = 1
Do Until sourcetable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = sourcetable!doc_no
    .TextMatrix(.Rows - 1, 1) = sourcetable!mytime & ""
    .TextMatrix(.Rows - 1, 2) = Format(sourcetable!SalesValue, "fixed")
    .TextMatrix(.Rows - 1, 3) = Format(sourcetable!SalesValueRet, "fixed")
    .TextMatrix(.Rows - 1, 4) = Format(sourcetable!discount, "fixed")
    .TextMatrix(.Rows - 1, 5) = Format(sourcetable!SalesValueNet, "fixed")
    .TextMatrix(.Rows - 1, 6) = sourcetable!MANDESCA & ""
    nTotalSales = Val(sourcetable!SalesValue & "") + nTotalSales
    nTotalSalesRet = Val(sourcetable!SalesValueRet & "") + nTotalSalesRet
    nTotalValue = Val(sourcetable!SalesValueNet & "") + nTotalValue
    nTotalDiscount = Val(sourcetable!discount & "") + nTotalDiscount
    sourcetable.MoveNext
Loop

If Not (sourcetable.EOF And sourcetable.BOF) Then
    .AddItem ""
    .TextMatrix(.Rows - 1, 1) = "«·«Ã„«·Ì"
    .TextMatrix(.Rows - 1, 2) = nTotalSales
    .TextMatrix(.Rows - 1, 3) = nTotalSalesRet
    .TextMatrix(.Rows - 1, 4) = nTotalDiscount
    .TextMatrix(.Rows - 1, 5) = nTotalValue
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
    .Cell(flexcpForeColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = vbBlue
End If
End With
End Sub


Private Sub cmdCloseDay_Click()
'doprint3
'If IsDate(xDate.Value) Then CloseDay.sDate = Format(xDate.Value, "dd-mm-yyyy")
'CloseDay.Show 1
End Sub
Private Sub Command1_Click()
doprint2
End Sub
Private Sub Command2_Click()
If Me.SSTab1.TabIndex = 0 Then
    doprint1
Else
    doprint2
End If
End Sub

Private Sub Command3_Click()
Dim loctable As New ADODB.Recordset
On Error GoTo myerror
Me.MousePointer = 11
loctable.Open "select file6_20.doc_no,file6_20.item,file6_20h.date,file6_20.row from file6_20 inner join file6_20h on file6_20.doc_no = file6_20h.doc_no", con, adOpenStatic, adLockReadOnly
con.BeginTrans
Do
    con.Execute "update file6_20 set file6_20.cost = " & Round(itemCost(loctable!Item, Format(loctable!Date, "dd-mm-yyyy")), 2) & _
                 " where file6_20.doc_no = " & MyParn(loctable!doc_no) & " and " & _
                 " file6_20.row = " & Val(loctable!Row & "")
    loctable.MoveNext
Loop Until loctable.EOF
con.CommitTrans
Me.MousePointer = 0
MsgBox "done ..."
Exit Sub
Me.MousePointer = 0
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub

Private Sub Command4_Click(Index As Integer)
Unload Me
End Sub
Private Sub Form_Load()
openCon con
 xDate.Visible = True
data1.ConnectionString = strCon
Fixgrd
'VsSales.RowHeight(0) = 400
'VsSales.Sort = flexSortNone
'VsSales.ExplorerBar = flexExSortShow
xDate.Value = DateValue(Format(SALESFRM.xDate.Text, "dd-mm-yyyy"))
MakeDataString
Set xMan.RowSource = data1
xMan.ListField = "Desca"
xMan.BoundColumn = "CODE"
Dim cdefman As String
cdefman = GetDesca("Select code from file6_25 where usercode = " & Val(nusercode))
xMan.BoundText = cdefman
myload
End Sub
Private Sub CMD_PRINT_Click()
    Load REP1_27
    REP1_27.date1.Text = Format(xDate.Value, "dd-mm-yyyy")
    REP1_27.date2.Text = Format(xDate.Value, "dd-mm-yyyy")
    REP1_27.Show 1
End Sub
Private Sub SSCommand3_Click()
    Load Rep_107
    Rep_107.xDate1.Text = xDate.Value
    Rep_107.xdate2.Text = xDate.Value
    Rep_107.xStore.BoundText = xStore.BoundText
    Rep_107.Show 1
End Sub
Private Sub VsSales_DBLClick()
    ViewSale.Show 1
    myload
End Sub

Private Sub SSCommand7_Click()

End Sub

Private Sub X5_Click()
End Sub

Private Sub Grid2_dblClick()
'Dim osalesfrm As New SalesFrm
'osalesfrm.myPublic = 0
'osalesfrm.sDoc_No = Grid2.TextMatrix(Grid2.Row, 0)
'osalesfrm.Show 1
salesdtlfrm.cDoc_No = grid2.TextMatrix(grid2.Row, 0)
salesdtlfrm.Show 1
End Sub

Private Sub Grid2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Grid2_dblClick
End Sub

Private Sub xDate_BeforeUpdate(Cancel As Integer)
   ' MyLoad
   ' MyLoadTotal
End Sub
Private Sub xDate_Click()
    myload
    MyLoadTotal

End Sub

Private Sub xStore_Click(Area As Integer)
    myload
End Sub
Private Sub MyLoadTotal()
If IsDate(xDate.Value) Then
    cwhere = cwhere & turn(cwhere, " and ") & " FILE6_20H.DATE = " & DateSq(Format(xDate.Value, "DD-MM-YYYY"))
End If

cField1 = "Sum(1) as SalesCount"

cField2 = "(Select Sum(file6_20.total) from file6_20 inner join file6_20h on file6_20.doc_no = file6_20h.doc_no where file6_20h.code = '000'"
If cwhere <> "" Then cField2 = cField2 & turn(cwhere, " and ") & cwhere
cField2 = cField2 & ") as sumofCash"

cField3 = "(Select Sum(file6_20.total) from file6_20 inner join file6_20h on file6_20.doc_no = file6_20h.doc_no where file6_20h.code <> '000'"
If cwhere <> "" Then cField3 = cField3 & turn(cwhere, " and ") & cwhere
cField3 = cField3 & ") as sumofDelay"

cField4 = "(Select Sum(file6_20.total) from file6_20 inner join file6_20h on file6_20.doc_no = file6_20h.doc_no"
If cwhere <> "" Then cField4 = cField4 & turn(cwhere, " and ") & cwhere
cField4 = cField4 & ") as sumofTotal"


cField5 = myiif("FILE6_20H.CODE = '000'", "DISCOUNT") & " sumofCashDiscount"
cField6 = myiif("FILE6_20H.CODE <> '000'", "DISCOUNT") & " sumofDelayDiscount"
cField7 = "SUM(FILE6_20H.DISCOUNT) AS sumofTotalDiscount"

cString = "SELECT " & _
            cField1 & _
           "," & cField2 & _
           "," & cField3 & _
           "," & cField4 & _
           "," & cField5 & _
           "," & cField6 & _
           "," & cField7 & _
          " FROM FILE6_20H"

cString = cString & turn(cwhere) & cwhere
                                                     
Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenKeyset, adCmdText
                            
With sourcetable
If Not (sourcetable.EOF And sourcetable.BOF) Then
    xSalesCount.Caption = Format(sourcetable!SalesCount, "Fixed")
    
    xsumofCash.Caption = Format(sourcetable!SumOfCash, "Fixed")
    xSumofCashDiscount.Caption = Format(sourcetable!sumofCashDiscount, "Fixed")
    xsumofCashNet.Caption = Val(sourcetable!SumOfCash & "") - Val(sourcetable!sumofCashDiscount & "")
    
    xsumofDelay.Caption = Format(sourcetable!SumOfDelay, "Fixed")
    xsumofDelayDiscount.Caption = Format(sourcetable!sumofDelayDiscount, "Fixed")
    xsumofDelayNet.Caption = Format(Val(sourcetable!SumOfDelay & "") - Val(sourcetable!sumofDelayDiscount & ""), "Fixed")
    
    
    
    xsumofTotal.Caption = Format(sourcetable!sumofTotal, "Fixed")
    xSumofTotalDiscount.Caption = Format(sourcetable!sumofTotalDiscount, "Fixed")
    xSumofTotalNet.Caption = Format(Val(sourcetable!sumofTotal & "") - Val(sourcetable!sumofTotalDiscount & ""), "Fixed")
End If
End With
End Sub
Private Sub doprint1()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For i = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!str21 = "≈Ã„«·Ì „»Ì⁄«  «’‰«› ÌÊ„ " & NameOfDay(xDate.Value) & " «·„Ê«›ﬁ : " & Format(xDate.Value, "dd-mm-yy")
    If xMan.MatchedWithList Then temptable!str22 = "«·ﬂ«‘Ì— : " & xMan.Text
    temptable!str1 = TurnValue(grid1.TextMatrix(i, 0))
    temptable!str2 = TurnValue(grid1.TextMatrix(i, 1))
    temptable!val1 = Val(grid1.TextMatrix(i, 2))
    temptable!val2 = Val(grid1.TextMatrix(i, 3))
    temptable!val3 = Val(grid1.TextMatrix(i, 4))
    temptable!VAL4 = Val(grid1.TextMatrix(i, 5))
    temptable!val5 = Val(grid1.TextMatrix(i, 6))
    temptable!Val6 = Val(grid1.TextMatrix(i, 7))
    temptable.Update
Next

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\salesday1.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub doprint2()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For i = 1 To grid2.Rows - 2
    temptable.AddNew
    temptable!str21 = "≈Ã„«·Ì »Ê‰«  »Ì⁄ ÌÊ„ " & NameOfDay(xDate.Value) & " «·„Ê«›ﬁ : " & Format(xDate.Value, "dd-mm-yy")
    If xMan.MatchedWithList Then temptable!str22 = "«·ﬂ«‘Ì— : " & xMan.Text
    temptable!str1 = TurnValue(grid2.TextMatrix(i, 0))
    temptable!str2 = TurnValue(grid2.TextMatrix(i, 1))
    temptable!val1 = Val(grid2.TextMatrix(i, 2))
    temptable!val2 = Val(grid2.TextMatrix(i, 3))
    temptable!val3 = Val(grid2.TextMatrix(i, 4))
    temptable!VAL4 = Val(grid2.TextMatrix(i, 5))
    temptable.Update
Next

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\salesday2.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub doprint3()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

If IsDate(xDate.Value) Then cwhere = turn(cwhere, " AND ") & " FILE6_20H.DATE = " & DateSq(Format(xDate.Value, "dd-mm-yyyy"))
If xMan.BoundText <> "" Then cwhere = cwhere & turn(cwhere, " AND ") & " FILE6_20H.MAN = " & MyParn(xMan.BoundText)
cString = "SELECT FILE6_20H.DOC_NO,FILE6_20H.MAN,FILE6_25.DESCA,FILE6_20H.CASH," & _
          " Sum(Val(FILE6_20.QUANT & '')* Val(FILE6_20.PRICE & '')) - VAL(FILE6_20H.DISCOUNT & '') + VAL(FILE6_20H.TAX & '') as SumOfTotal,FILE6_20H.VISA  " & _
          " FROM (FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO) LEFT JOIN FILE6_25 ON FILE6_20H.MAN = FILE6_25.CODE "
cString = cString & TurnWhere(cwhere) & cwhere
cString = cString & " GROUP BY FILE6_20H.DOC_NO,FILE6_20H.DISCOUNT,FILE6_20H.CASH,FILE6_20H.TAX,FILE6_20H.MAN,FILE6_25.DESCA,FILE6_20H.VISA ORDER BY FILE6_20H.DOC_NO"
                       
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
    If xMan.BoundText = "" Then
       temptable!str21 = "≈€·«ﬁ ÌÊ„Ì… " & NameOfDay(xDate.Value) & " «·„Ê«›ﬁ : " & Format(xDate.Value, "dd-mm-yy")
    Else
       temptable!str21 = "≈€·«ﬁ ÌÊ„Ì… «·ﬂ«‘Ì— " & xMan.Text & " ·ÌÊ„ " & NameOfDay(xDate.Value) & " «·„Ê«›ﬁ : " & Format(xDate.Value, "dd-mm-yy")
    End If
    temptable!str1 = sourcetable!MAN
    temptable!str2 = sourcetable!DESCA
    temptable!str3 = sourcetable!doc_no
    temptable!val1 = sourcetable!sumofTotal
    temptable!val2 = Val(sourcetable!CASH & "")
    temptable!val3 = Val(sourcetable!Visa & "")
    temptable!VAL4 = sourcetable!sumofTotal - (Val(sourcetable!CASH & "") + Val(sourcetable!Visa & ""))
    temptable!val11 = 0
    temptable!val12 = IIf(xMan.BoundText = "", 0, 1)
    temptable!str11 = "„»Ì⁄« "
    temptable.Update
    sourcetable.MoveNext
Loop

cwhere = ""

If IsDate(xDate.Value) Then cwhere = TurnAnd(cwhere) & " FILE6_10H.DATE = " & DateSq(Format(xDate.Value, "dd-mm-yyyy"))
If xMan.BoundText <> "" Then cwhere = cwhere & TurnAnd(cwhere) & " FILE6_10H.MAN = " & MyParn(xMan.BoundText)

cString = "SELECT FILE6_10H.DOC_NO,FILE6_10H.MAN,FILE6_25.DESCA,FILE6_10H.CASH," & _
          " Sum(Val(FILE6_10.QUANT & '')* Val(FILE6_10.PRICE & '')) - VAL(FILE6_10H.DISCOUNT & '') + VAL(FILE6_10H.TAX & '') as SumOfTotal " & _
          " FROM (FILE6_10H INNER JOIN FILE6_10 ON FILE6_10H.DOC_NO = FILE6_10.DOC_NO) LEFT JOIN FILE6_25 ON FILE6_10H.MAN = FILE6_25.CODE "
cString = cString & TurnWhere(cwhere) & cwhere
cString = cString & " GROUP BY FILE6_10H.DOC_NO,FILE6_10H.DISCOUNT,FILE6_10H.TAX,FILE6_10H.CASH,FILE6_10H.MAN,FILE6_25.DESCA ORDER BY FILE6_10H.DOC_NO"
sourcetable.Close
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
    If xMan.BoundText = "" Then
        temptable!str21 = "≈€·«ﬁ ÌÊ„Ì… " & NameOfDay(xDate.Value) & " «·„Ê«›ﬁ : " & Format(xDate.Value, "dd-mm-yy")
    Else
       temptable!str21 = "≈€·«ﬁ ÌÊ„Ì… «·»«∆⁄ " & xMan.Text & " ·ÌÊ„ " & NameOfDay(xDate.Value) & " «·„Ê«›ﬁ : " & Format(xDate.Value, "dd-mm-yy")
    End If
    temptable!str1 = sourcetable!MAN
    temptable!str2 = sourcetable!DESCA
    temptable!str3 = sourcetable!doc_no
    temptable!val1 = -1 * sourcetable!sumofTotal
    temptable!val2 = -1 * sourcetable!CASH
    temptable!val11 = 1
    temptable!val3 = 0
    temptable!VAL4 = -1 * (sourcetable!sumofTotal - sourcetable!CASH)
    temptable!val12 = IIf(xMan.BoundText = "", 0, 1)
    temptable!str11 = "„—œÊœ „»Ì⁄« "
    temptable.Update
    sourcetable.MoveNext
Loop

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\daySal3.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub

Private Sub xMan_Change()
If xMan.MatchedWithList Or Trim(xMan.Text) = "" Then
    myload
    MyLoadTotal
'    Me.Caption = Val(Me.Caption) + 1
End If
End Sub

Private Sub xMan_Click(Area As Integer)
'If Area = 2 Then
'    MyLoad
'    MyLoadTotal
'End If
End Sub
Private Sub MakeDataString()
data1.RecordSource = "SELECT Distinct code,FILE6_25.DESCA FROM FILE6_25 "
End Sub
Private Sub xMan_LostFocus()
If Not xMan.MatchedWithList Then xMan.BoundText = ""
End Sub
Private Sub Fixgrd()
With grid1
.Cols = 8
.TextMatrix(0, 0) = "ﬂÊœ «·’‰›"
.TextMatrix(0, 1) = "«·«”„"
.TextMatrix(0, 2) = "ﬂ„Ì… »Ì⁄"
.TextMatrix(0, 3) = "ﬁÌ„… »Ì⁄"
.TextMatrix(0, 4) = "ﬂ„Ì… „— Ã⁄"
.TextMatrix(0, 5) = "ﬁÌ„… „— Ã⁄"
.TextMatrix(0, 6) = "’«›Ì ﬂ »Ì⁄"
.TextMatrix(0, 7) = "’«›Ì ﬁ »Ì⁄"

.ColWidth(0) = 1600
.ColWidth(1) = 3000
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1000
.ColWidth(7) = 1000

.MergeCells = flexMergeFree
.MergeCol(0) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
End With

With grid2
.Cols = 7
.TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
.TextMatrix(0, 1) = "«·Êﬁ "
.TextMatrix(0, 2) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 3) = "ﬁÌ„… „—œÊœ „»»⁄« "
.TextMatrix(0, 4) = "«·Œ’„"
.TextMatrix(0, 5) = "’«›Ì «·„»Ì⁄« "
.TextMatrix(0, 6) = "«·»«∆⁄"

.ColWidth(0) = 1400
.ColWidth(1) = 1400
.ColWidth(2) = 1400
.ColWidth(3) = 1400
.ColWidth(4) = 1400
.ColWidth(5) = 1400
.ColWidth(6) = 2000

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
.ExplorerBar = flexExSortShow
End With
End Sub

Private Sub xSalesValue_Click()

End Sub

Private Sub xSalesValueRet_Click()

End Sub


