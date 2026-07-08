VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form TSal_DashBoard 
   BackColor       =   &H00E0E0E0&
   Caption         =   "„ Ê”ÿ „»Ì⁄«  «·›—Ê⁄ Ê «· ÊﬂÌ·« "
   ClientHeight    =   10230
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   13380
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
   ScaleWidth      =   13380
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
      Height          =   1185
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   4785
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "TSal_DashBoard.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   180
         Width           =   1050
      End
      Begin VB.CommandButton cmdGo 
         Height          =   495
         Left            =   1170
         Picture         =   "TSal_DashBoard.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   825
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   810
         Width           =   1275
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   765
         Width           =   1230
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   510
         Left            =   2205
         TabIndex        =   17
         Top             =   180
         Width           =   2370
         _ExtentX        =   4180
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
         Picture         =   "TSal_DashBoard.frx":495E
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
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
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   810
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   765
         Width           =   780
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "„‰«ÿﬁ"
      Height          =   2400
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   3375
      Width           =   4965
      Begin VSFlex7Ctl.VSFlexGrid GRID4 
         Height          =   2175
         Left            =   45
         TabIndex        =   19
         Top             =   180
         Width           =   4830
         _cx             =   8520
         _cy             =   3836
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
         BackColorSel    =   16761024
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
         SelectionMode   =   1
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
   Begin VB.Frame Frame5 
      Caption         =   "„Ê—œÌ‰"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5145
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   5670
      Width           =   5910
      Begin VSFlex7Ctl.VSFlexGrid GRID3 
         Height          =   4740
         Left            =   45
         TabIndex        =   16
         Top             =   270
         Width           =   5775
         _cx             =   10186
         _cy             =   8361
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
         BackColorSel    =   16761024
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
         SelectionMode   =   1
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
   Begin VB.Frame Frame4 
      Caption         =   "«ﬁ”«„ - „Ê«”„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5145
      Left            =   5985
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   5670
      Width           =   6810
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   4740
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   6675
         _cx             =   11774
         _cy             =   8361
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
   Begin VB.Frame Frame3 
      Caption         =   "«ﬁ”«„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5145
      Left            =   12825
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   5670
      Width           =   7395
      Begin VSFlex7Ctl.VSFlexGrid GRID1 
         Height          =   4740
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   7260
         _cx             =   12806
         _cy             =   8361
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
         BackColorSel    =   16761024
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
         SelectionMode   =   1
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
   Begin MSComctlLib.ProgressBar PROG1 
      Height          =   210
      Left            =   135
      TabIndex        =   9
      Top             =   3420
      Visible         =   0   'False
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Height          =   5505
      Left            =   4950
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   15180
      Begin VSFlex7Ctl.VSFlexGrid GridBox 
         Bindings        =   "TSal_DashBoard.frx":7159
         Height          =   5235
         Left            =   90
         TabIndex        =   2
         Top             =   135
         Width           =   15000
         _cx             =   26458
         _cy             =   9234
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
      Left            =   135
      TabIndex        =   0
      Top             =   1080
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12632256
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   110624775
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
   Begin MSAdodcLib.Adodc DATA4 
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
Attribute VB_Name = "TSal_DashBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conShop As New ADODB.Connection
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
    ToFileExel2 GridBox, , , , , 1.1, , , , , , Me
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myloadBox ""
    MYLOAD1
    MYLOAD2
    MYLOAD3
    MYLOAD4
End Sub
Private Sub Form_Load()
    openCon con
    xDate.Value = dSalesDate
    
    Set grid1.DataSource = data1
    Set GRID2.DataSource = DATA2
    Set GRID3.DataSource = data3
    Set GRID4.DataSource = data4
    Set GridBox.DataSource = DATABOX
    
    grid1.Rows = 1
    GRID2.Rows = 1
    GRID3.Rows = 1
    GRID4.Rows = 1
    GridBox.Rows = 1
    
    fixGrdBox
    Fixgrd1
    Fixgrd2
    Fixgrd3
    Fixgrd4

    xDate1.text = xDate.Value
    xdate2.text = xDate.Value
    myloadBox ""
End Sub
Private Sub Fixgrd2()
With GRID2
.Cols = 6
.WordWrap = True
.RowHeight(0) = 600
.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·ﬁ”„"
.TextMatrix(0, 2) = "„Ê”„"
.TextMatrix(0, 3) = "⁄œœ «·ﬁÿ⁄"
.TextMatrix(0, 4) = "ﬁÌ„… «·„»Ì⁄« "
.TextMatrix(0, 5) = "„ Ê”ÿ ﬁÌ„… «·ﬁÿ⁄…"

.ColFormat(3) = "#0"
.ColFormat(4) = "#0.00"
.ColFormat(5) = "#0.00"
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble

.ColWidth(0) = 500
.ColWidth(1) = 2000
.ColWidth(2) = 700
.ColWidth(3) = 800
.ColWidth(4) = 1200
.ColWidth(5) = 900
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
For nRow = 1 To .Rows - 1
    If Val(.TextMatrix(nRow, 3)) <> 0 Then .TextMatrix(nRow, 5) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 3)), 2)
Next nRow
.SubtotalPosition = flexSTBelow
.ExplorerBar = flexExSortShow
.Subtotal flexSTSum, -1, 3, "#0", vbYellow, , True, " "
.Subtotal flexSTSum, -1, 4, "#0", vbYellow, , True, " "
If .Rows > 1 Then
    nRow = .Rows - 1
    If Val(.TextMatrix(nRow, 3)) <> 0 Then .TextMatrix(nRow, 5) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 3)), 2)
End If
End With
End Sub
Private Sub Fixgrd1()
With grid1
.Cols = 7
.WordWrap = True
.RowHeight(0) = 600
.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·ﬁ”„"
.TextMatrix(0, 2) = "⁄œœ «·ﬁÿ⁄"
.TextMatrix(0, 3) = "ﬁÌ„… «·„»Ì⁄« "
.TextMatrix(0, 4) = "„ Ê”ÿ ﬁÌ„… «·ﬁÿ⁄…"

.TextMatrix(0, 5) = "⁄œœ «Ê ·Ì "
.TextMatrix(0, 6) = "ﬁÌ„… «Ê ·Ì "
.ColHidden(0) = True
.ColFormat(2) = "#0"
.ColFormat(3) = "#0.00"
.ColFormat(4) = "#0.00"
.ColDataType(2) = flexDTDouble
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble

.ColWidth(0) = 500
.ColWidth(1) = 2000
.ColWidth(2) = 800
.ColWidth(3) = 1200
.ColWidth(4) = 900

.ColWidth(5) = 800
.ColWidth(6) = 1100

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
For nRow = 1 To .Rows - 1
    If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 4) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
Next nRow
.SubtotalPosition = flexSTBelow
.ExplorerBar = flexExSortShow
.Subtotal flexSTSum, -1, 2, "#0", vbYellow, , True, " "
.Subtotal flexSTSum, -1, 3, "#0", vbYellow, , True, " "
.Subtotal flexSTSum, -1, 5, "#0", vbYellow, , True, " "
.Subtotal flexSTSum, -1, 6, "#0", vbYellow, , True, " "

If .Rows > 1 Then
    nRow = .Rows - 1
    If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 4) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
End If
End With
End Sub
Private Sub fixGrdBox()
With GridBox
.FixedCols = 0
.FixedRows = 1
.Cols = 14
.RowHeight(0) = 1000
.WordWrap = True

.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·›—⁄"
.TextMatrix(0, 2) = "⁄œœ »Ê‰« "
.TextMatrix(0, 3) = "⁄œœ ﬁÿ⁄ „»Ì⁄« "
.TextMatrix(0, 4) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 5) = "‰ﬁœÏ"
.TextMatrix(0, 6) = "›Ì“«"
.TextMatrix(0, 7) = "point"
.TextMatrix(0, 8) = "ONLINE"

.TextMatrix(0, 9) = "„ Ê”ÿ ⁄œœ «·ﬁÿ⁄ ··»Ê‰"
.TextMatrix(0, 10) = "„ Ê”ÿ ﬁÌ„… «·»Ê‰"
.TextMatrix(0, 11) = "„ Ê”ÿ ”⁄— «·ﬁÿ⁄…"


.TextMatrix(0, 12) = "⁄œœ «Ê ·Ì "
.TextMatrix(0, 13) = "ﬁÌ„… «Ê  ·Ì "


'.RowHeight(0) = 800
'.WordWrap = True
.ColWidth(0) = 600
.ColWidth(1) = 1500
.ColWidth(2) = 800
.ColWidth(3) = 800
.ColWidth(4) = 1400
.ColWidth(5) = 1400
.ColWidth(6) = 1400
.ColWidth(7) = 1100
.ColWidth(8) = 950
.ColWidth(9) = 900
.ColWidth(9) = 900
.ColWidth(10) = 900
.ColWidth(11) = 900
.ColWidth(12) = 900
.ColWidth(13) = 1000

For nRow = 1 To .Rows - 1
    If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
    If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 2)), 2)
    If Val(.TextMatrix(nRow, 3)) <> 0 Then .TextMatrix(nRow, 11) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 3)), 2)
    For nCol = 4 To .Cols - 1
        .TextMatrix(nRow, nCol) = Format(Val(.TextMatrix(nRow, nCol)), "#0.00")
    Next nCol
Next nRow
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
If .Rows = 1 Then Exit Sub
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTBelow
For nCol = 2 To 8
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
Next nCol
.Subtotal flexSTSum, -1, 12, "#0", vbYellow, , True, " "
.Subtotal flexSTSum, -1, 13, "#0.00", vbYellow, , True, " "

If Not bOpt5 Then Exit Sub
nRow = .Rows - 1

If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 2)), 2)
If Val(.TextMatrix(nRow, 3)) <> 0 Then .TextMatrix(nRow, 11) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 3)), 2)

.AddItem ""
.AddItem ""
nRow = .Rows - 1
.TextMatrix(nRow, 1) = "≈Ã„«·Ï ›—Ê⁄"
cSubStr1 = ""
If Not bOpt5 Then cSubStr1 = " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "


cF1 = ",(SELECT SUM(CASH) FROM QFILE6_20H WHERE QFILE6_20H.STORE  <> '000' AND LEN(QFILE6_20H.STORE ) = 3 AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & cSubStr1 & " )"
cF2 = ",(SELECT SUM(VISA) FROM QFILE6_20H WHERE QFILE6_20H.STORE  <> '000' AND LEN(QFILE6_20H.STORE ) = 3 AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & cSubStr1 & " )"
cF3 = ",(SELECT SUM(POINT) FROM QFILE6_20H WHERE QFILE6_20H.STORE  <> '000' AND LEN(QFILE6_20H.STORE ) = 3 AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & cSubStr1 & " )"
cF4 = ",(SELECT SUM(LATE ) FROM QFILE6_20H WHERE QFILE6_20H.STORE  <> '000' AND LEN(QFILE6_20H.STORE ) = 3 AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & cSubStr1 & " )"
cStr1 = " SELECT ' ' , ' ' , " & _
            "  COUNT(DISTINCT QFILE6_20_1.DOC_NO), SUM(QFILE6_20_1.QUANT) AS t_q,SUM(QFILE6_20_1.TOTAL * ((100-QFILE6_20_1.RATE )/100)) AS t_sales  " & cF1 & cF2 & cF3 & cF4 & _
            " FROM QFILE6_20_1 WHERE QFILE6_20_1.STORE  <> '000' "


If Not bOpt5 Then cStr1 = cStr1 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
cStr1 = cStr1 & turn(cStr1) & "DATE >= " & DateSq(xDate1.text)
cStr1 = cStr1 & turn(cStr1) & "DATE <= " & DateSq(xdate2.text)
aRet = aGetDesca(cStr1, con)

.TextMatrix(nRow, 2) = Round(Val(aRet(3) & ""), 2)
.TextMatrix(nRow, 3) = Round(Val(aRet(4) & ""), 2)
.TextMatrix(nRow, 4) = Round(Val(aRet(5) & ""), 2)
.TextMatrix(nRow, 5) = Round(Val(aRet(6) & ""), 2)
.TextMatrix(nRow, 6) = Round(Val(aRet(7) & ""), 2)
.TextMatrix(nRow, 7) = Round(Val(aRet(8) & ""), 2)
.TextMatrix(nRow, 8) = Round(Val(aRet(9) & ""), 2)

If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 2)), 2)
If Val(.TextMatrix(nRow, 3)) <> 0 Then .TextMatrix(nRow, 11) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 3)), 2)

.AddItem ""
nRow = .Rows - 1
.TextMatrix(nRow, 1) = "≈Ã„«·Ï  ÊﬂÌ·«  "
cF1 = ",(SELECT SUM(CASH) FROM QFILE6_20H WHERE LEN(QFILE6_20H.STORE ) = 2 AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & " )"
cF2 = ",(SELECT SUM(VISA) FROM QFILE6_20H WHERE LEN(QFILE6_20H.STORE ) = 2 AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & " )"
cF3 = ",(SELECT SUM(POINT) FROM QFILE6_20H WHERE LEN(QFILE6_20H.STORE ) = 2 AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & " )"
cF4 = ",(SELECT SUM(LATE ) FROM QFILE6_20H WHERE LEN(QFILE6_20H.STORE ) = 2 AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & " )"
cStr1 = " SELECT ' ', ' ', " & _
            "  COUNT(DISTINCT QFILE6_20_2.DOC_NO), SUM(QFILE6_20_2.QUANT) AS t_q,SUM(QFILE6_20_2.TOTAL * ((100-QFILE6_20_2.RATE )/100)) AS t_sales  " & cF1 & cF2 & cF3 & cF4 & _
            " FROM QFILE6_20_2 WHERE QFILE6_20_2.doc_no IS NOT NULL "
cStr1 = cStr1 & turn(cStr1) & "DATE >= " & DateSq(xDate1.text)
cStr1 = cStr1 & turn(cStr1) & "DATE <= " & DateSq(xdate2.text)
aRet = aGetDesca(cStr1, con)

.TextMatrix(nRow, 2) = Round(Val(aRet(3) & ""), 2)
.TextMatrix(nRow, 3) = Round(Val(aRet(4) & ""), 2)
.TextMatrix(nRow, 4) = Round(Val(aRet(5) & ""), 2)
.TextMatrix(nRow, 5) = Round(Val(aRet(6) & ""), 2)
.TextMatrix(nRow, 6) = Round(Val(aRet(7) & ""), 2)
.TextMatrix(nRow, 7) = Round(Val(aRet(8) & ""), 2)
.TextMatrix(nRow, 8) = Round(Val(aRet(9) & ""), 2)

If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 8) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 2)), 2)
If Val(.TextMatrix(nRow, 3)) <> 0 Then .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 3)), 2)
If .Rows > 1 Then
    nRow = .Rows - 1
    If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
    If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 2)), 2)
    If Val(.TextMatrix(nRow, 3)) <> 0 Then .TextMatrix(nRow, 11) = Round(Val(.TextMatrix(nRow, 4)) / Val(.TextMatrix(nRow, 3)), 2)
End If
End With
End Sub

Private Sub grid1_DblClick()
With grid1
    myloadBox .TextMatrix(.Row, 0)
    MYLOAD2 , , .TextMatrix(.Row, 0)
    MYLOAD3 , .TextMatrix(.Row, 0)
End With
End Sub
Private Sub GRID3_DblClick()
With GRID3
    MYLOAD1 , .TextMatrix(.Row, 0)
    MYLOAD2 , .TextMatrix(.Row, 0)
End With
End Sub

Private Sub GRID4_DblClick()
If grid1.Rows > 1 Then
    myloadBox "", GRID4.TextMatrix(GRID4.Row, 0)


    MYLOAD1 , , GRID4.TextMatrix(GRID4.Row, 0)
    MYLOAD2 , , , GRID4.TextMatrix(GRID4.Row, 0)
    MYLOAD3 , , GRID4.TextMatrix(GRID4.Row, 0)

End If
End Sub

Private Sub gridbox_DblClick()
If grid1.Rows > 1 Then
    MYLOAD1 GridBox.TextMatrix(GridBox.Row, 0)
    MYLOAD2 GridBox.TextMatrix(GridBox.Row, 0)
    MYLOAD3 GridBox.TextMatrix(GridBox.Row, 0)
End If
End Sub
Private Sub xdate_DateClick(ByVal DateClicked As Date)
    xDate1.text = Format(xDate.Value, "DD-MM-YYYY")
    xdate2.text = Format(xDate.Value, "DD-MM-YYYY")
    myloadBox ""
End Sub
Sub myloadBox(pSec, Optional CGROUP As String)
Dim cStr1 As String
cF1 = ",(SELECT SUM(CASH) FROM QFILE6_20H WHERE QFILE6_20H.STORE = QFILE6_20.STORE AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & " )"
cF2 = ",(SELECT SUM(VISA) FROM QFILE6_20H WHERE QFILE6_20H.STORE = QFILE6_20.STORE AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & " )"
cF3 = ",(SELECT SUM(POINT) FROM QFILE6_20H WHERE QFILE6_20H.STORE = QFILE6_20.STORE AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & " )"
cF4 = ",(SELECT SUM(LATE) FROM QFILE6_20H WHERE QFILE6_20H.STORE = QFILE6_20.STORE AND QFILE6_20H.DATE >= " & DateSq(xDate1.text) & " AND QFILE6_20H.DATE <= " & DateSq(xdate2.text) & " )"

cF5 = myiif("QFILE6_20.Outlet = 1", "(QFILE6_20.QUANT)")
cF6 = myiif("QFILE6_20.Outlet = 1", "(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100))")

cStr1 = " SELECT QFILE6_20.store , STORE_BR.DESCA , " & _
            "  COUNT(DISTINCT QFILE6_20.DOC_NO), SUM(QFILE6_20.QUANT) AS t_q,SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS t_sales  " & cF1 & cF2 & cF3 & cF4 & " ,0,0,0, " & cF5 & " , " & cF6 & _
            " FROM QFILE6_20 INNER JOIN STORE_BR ON qFILE6_20.store = STORE_BR.CODE WHERE QFILE6_20.STORE <> '000' "
If pSec <> "" Then
    cStr1 = cStr1 & " AND [SECTION] = " & Val(pSec)
End If
If CGROUP <> "" Then cStr1 = cStr1 & " AND QFILE6_20.store  IN ( SELECT CODE FROM STORE_BR2 WHERE BRANCH_GR = " & MyParn(CGROUP) & " ) "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xDate1.text)
cStr1 = cStr1 & " AND DATE <= " & DateSq(xdate2.text)
If Not bOpt5 Then cStr1 = cStr1 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
cStr1 = cStr1 & " GROUP BY QFILE6_20.store , STORE_BR.DESCA ORDER BY QFILE6_20.store  "
Set DATABOX.Recordset = myRecordSet(cStr1, con)

fixGrdBox

MYLOAD1
MYLOAD2
MYLOAD3
MYLOAD4

End Sub
Sub MYLOAD2(Optional cStore As String, Optional cCode As String, Optional cSec As String, Optional CGROUP As String)
    Dim cString As String
    cString = " SELECT QFILE6_20.[SECTION] , FILE1_10SC.DESCA , QFILE6_20.[MOSM]  , " & _
                "  SUM(QFILE6_20.QUANT) AS t_q,SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS t_sales " & _
                " FROM QFILE6_20 INNER JOIN FILE1_10SC ON qFILE6_20.[SECTION] = FILE1_10SC.CODE WHERE QFILE6_20.STORE <> '000' "
    cString = cString & turn(cString) & "DATE >= " & DateSq(xDate1.text)
    cString = cString & turn(cString) & "DATE <= " & DateSq(xdate2.text)
    If Not bOpt5 Then cString = cString & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If cStore <> "" Then cString = cString & turn(cString) & "STORE = " & MyParn(cStore)
    If cCode <> "" Then cString = cString & turn(cString) & " qFILE6_20.CODE = " & MyParn(cCode)
    If cSec <> "" Then cString = cString & turn(cString) & " qFILE6_20.[SECTION]= " & MyParn(cSec)
    If CGROUP <> "" Then cStr1 = cStr1 & " AND QFILE6_20.store  IN ( SELECT CODE FROM STORE_BR2 WHERE BRANCH_GR = " & MyParn(CGROUP) & " ) "
    cString = cString & " GROUP BY QFILE6_20.[SECTION] , FILE1_10SC.DESCA , QFILE6_20.[MOSM]  ORDER BY QFILE6_20.[SECTION] , QFILE6_20.[MOSM]  "
    Set DATA2.Recordset = myRecordSet(cString, con)
    Fixgrd2
End Sub
Sub MYLOAD1(Optional cStore As String, Optional cCode As String, Optional CGROUP As String)
Dim cString As String
cF4 = myiif("QFILE6_20.Outlet = 1", "(QFILE6_20.QUANT)")
cF5 = myiif("QFILE6_20.Outlet = 1", "(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100))")

cString = " SELECT QFILE6_20.[SECTION] , FILE1_10SC.DESCA , " & _
            "  SUM(QFILE6_20.QUANT) AS t_q,SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS t_sales , 0 , " & cF4 & " , " & cF5 & _
            " FROM QFILE6_20 INNER JOIN FILE1_10SC ON qFILE6_20.[SECTION] = FILE1_10SC.CODE WHERE QFILE6_20.STORE <> '000' "


cString = cString & turn(cString) & "DATE >= " & DateSq(xDate1.text)
cString = cString & turn(cString) & "DATE <= " & DateSq(xdate2.text)
If Not bOpt5 Then cString = cString & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
If cStore <> "" Then cString = cString & turn(cString) & "STORE = " & MyParn(cStore)
If cCode <> "" Then cString = cString & turn(cString) & " qFILE6_20.CODE = " & MyParn(cCode)
If CGROUP <> "" Then cStr1 = cStr1 & " AND QFILE6_20.store  IN ( SELECT CODE FROM STORE_BR2 WHERE BRANCH_GR = " & MyParn(CGROUP) & " ) "

cString = cString & " GROUP BY QFILE6_20.[SECTION] , FILE1_10SC.DESCA ORDER BY QFILE6_20.[SECTION] "
Set data1.Recordset = myRecordSet(cString, con)
Fixgrd1
End Sub
Sub MYLOAD3(Optional cStore As String, Optional cSec As String, Optional CGROUP As String)
Dim cString As String
cString = " SELECT QFILE6_20.[CODE] , FILE4_10.DESCA , " & _
            "  SUM(QFILE6_20.QUANT) AS t_q,SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS t_sales " & _
            " FROM QFILE6_20 INNER JOIN FILE4_10 ON qFILE6_20.[CODE] = FILE4_10.CODE WHERE QFILE6_20.STORE <> '000' "
If Not bOpt5 Then cString = cString & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
cString = cString & turn(cString) & "DATE >= " & DateSq(xDate1.text)
cString = cString & turn(cString) & "DATE <= " & DateSq(xdate2.text)
If cStore <> "" Then cString = cString & turn(cString) & "STORE = " & MyParn(cStore)
If cSec <> "" Then cString = cString & turn(cString) & " qFILE6_20.SECTION  = " & MyParn(cSec)
If CGROUP <> "" Then cStr1 = cStr1 & " AND QFILE6_20.store  IN ( SELECT CODE FROM STORE_BR2 WHERE BRANCH_GR = " & MyParn(CGROUP) & " ) "
cString = cString & " GROUP BY QFILE6_20.[CODE] , FILE4_10.DESCA ORDER BY QFILE6_20.[CODE] "
Set data3.Recordset = myRecordSet(cString, con)
Fixgrd3
End Sub
Sub MYLOAD4()
Dim cString As String
cString = " SELECT BRANCH_GR.[CODE] , BRANCH_GR.DESCA , " & _
            "  SUM(QFILE6_20.QUANT) AS t_q,SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS t_sales " & _
            "  FROM QFILE6_20 INNER JOIN STORE_BR2 ON qFILE6_20.[store] = STORE_BR2.CODE LEFT JOIN BRANCH_GR ON BRANCH_GR.CODE = STORE_BR2.BRANCH_GR   WHERE QFILE6_20.STORE <> '000' "
If Not bOpt5 Then cString = cString & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
cString = cString & turn(cString) & "DATE >= " & DateSq(xDate1.text)
cString = cString & turn(cString) & "DATE <= " & DateSq(xdate2.text)
cString = cString & " GROUP BY BRANCH_GR.[CODE] , BRANCH_GR.DESCA ORDER BY BRANCH_GR.[CODE]  "
Set data4.Recordset = myRecordSet(cString, con)
Fixgrd4
End Sub



Private Sub Fixgrd3()
With GRID3
.Cols = 5
.WordWrap = True
.RowHeight(0) = 600
.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·„Ê—œ"
.TextMatrix(0, 2) = "⁄œœ «·ﬁÿ⁄"
.TextMatrix(0, 3) = "ﬁÌ„… «·„»Ì⁄« "
.TextMatrix(0, 4) = "„ Ê”ÿ ﬁÌ„… «·ﬁÿ⁄…"

.ColFormat(2) = "#0"
.ColFormat(3) = "#0.00"
.ColFormat(4) = "#0.00"
.ColDataType(2) = flexDTDouble
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble

.ColWidth(0) = 500
.ColWidth(1) = 2000
.ColWidth(2) = 800
.ColWidth(3) = 1200
.ColWidth(4) = 900
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
For nRow = 1 To .Rows - 1
    If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 4) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
Next nRow
.SubtotalPosition = flexSTBelow
.ExplorerBar = flexExSortShow
.Subtotal flexSTSum, -1, 2, "#0", vbYellow, , True, " "
.Subtotal flexSTSum, -1, 3, "#0", vbYellow, , True, " "
If .Rows > 1 Then
    nRow = .Rows - 1
    If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 4) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
End If
End With
End Sub

Private Sub Fixgrd4()
With GRID4
.Cols = 5
.WordWrap = True
.RowHeight(0) = 600
.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·„‰ÿﬁ…"
.TextMatrix(0, 2) = "⁄œœ «·ﬁÿ⁄"
.TextMatrix(0, 3) = "ﬁÌ„… «·„»Ì⁄« "
.TextMatrix(0, 4) = "„ Ê”ÿ ﬁÌ„… «·ﬁÿ⁄…"
.ColHidden(0) = True
.ColFormat(2) = "#0"
.ColFormat(3) = "#0.00"
.ColFormat(4) = "#0.00"
.ColDataType(2) = flexDTDouble
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble

.ColWidth(0) = 500
.ColWidth(1) = 1300
.ColWidth(2) = 800
.ColWidth(3) = 1500
.ColWidth(4) = 800
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
For nRow = 1 To .Rows - 1
    If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 4) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
Next nRow
.SubtotalPosition = flexSTBelow
.ExplorerBar = flexExSortShow
.Subtotal flexSTSum, -1, 2, "#0", vbYellow, , True, " "
.Subtotal flexSTSum, -1, 3, "#0", vbYellow, , True, " "
If .Rows > 1 Then
    nRow = .Rows - 1
    If Val(.TextMatrix(nRow, 2)) <> 0 Then .TextMatrix(nRow, 4) = Round(Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2)), 2)
End If
End With
End Sub

