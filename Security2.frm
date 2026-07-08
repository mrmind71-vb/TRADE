VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form Security 
   Caption         =   "«·’·«ÕÌ« "
   ClientHeight    =   9795
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   10050
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   9795
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   3330
      Width           =   5595
      Begin VB.CheckBox Xisbranch 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„” Œœ„ ·ﬂ· «·›—Ê⁄"
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
         Height          =   330
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   270
         Width           =   1905
      End
      Begin VB.CommandButton cmdDate 
         Caption         =   " €ÌÌ— «· «—ÌŒ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   180
         Width           =   2175
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5010
      Left            =   90
      TabIndex        =   13
      ToolTipText     =   "450"
      Top             =   4095
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   8837
      _Version        =   393216
      Tabs            =   10
      TabsPerRow      =   10
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
      TabCaption(0)   =   "«Œ—Ì"
      TabPicture(0)   =   "Security2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grid1(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Œœ„« "
      TabPicture(1)   =   "Security2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid1(8)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "«· ﬁ«—Ì—"
      TabPicture(2)   =   "Security2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Grid1(7)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "«·»‰Êﬂ"
      TabPicture(3)   =   "Security2.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Grid1(6)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "‰ﬁœÌ…"
      TabPicture(4)   =   "Security2.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Grid1(5)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Online"
      TabPicture(5)   =   "Security2.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Grid1(4)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "›Ê« Ì—"
      TabPicture(6)   =   "Security2.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Grid1(3)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "„Ê—œÌ‰"
      TabPicture(7)   =   "Security2.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Grid1(2)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "⁄„·«¡"
      TabPicture(8)   =   "Security2.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Grid1(1)"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "«’‰«›"
      TabPicture(9)   =   "Security2.frx":00FC
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Grid1(0)"
      Tab(9).ControlCount=   1
      Begin VSFlex7LCtl.VSFlexGrid Grid1 
         Height          =   4605
         Index           =   0
         Left            =   -74955
         TabIndex        =   30
         Top             =   360
         Width           =   9780
         _cx             =   17251
         _cy             =   8123
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   2
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
         Editable        =   2
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
         Height          =   4605
         Index           =   1
         Left            =   -74955
         TabIndex        =   31
         Top             =   360
         Width           =   9780
         _cx             =   17251
         _cy             =   8123
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   2
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
         Editable        =   2
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
         Height          =   4605
         Index           =   2
         Left            =   -74955
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   360
         Width           =   9780
         _cx             =   17251
         _cy             =   8123
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   2
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
         Editable        =   2
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
         Height          =   4605
         Index           =   3
         Left            =   -74955
         TabIndex        =   33
         Top             =   360
         Width           =   9780
         _cx             =   17251
         _cy             =   8123
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   2
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
         Editable        =   2
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
         Height          =   4605
         Index           =   4
         Left            =   -74955
         TabIndex        =   34
         Top             =   360
         Width           =   9780
         _cx             =   17251
         _cy             =   8123
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   2
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
         Editable        =   2
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
         Height          =   4605
         Index           =   5
         Left            =   -74955
         TabIndex        =   35
         Top             =   360
         Width           =   9780
         _cx             =   17251
         _cy             =   8123
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   2
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
         Editable        =   2
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
         Height          =   4605
         Index           =   9
         Left            =   45
         TabIndex        =   36
         Top             =   360
         Width           =   9780
         _cx             =   17251
         _cy             =   8123
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   2
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
         Editable        =   2
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
         Height          =   4605
         Index           =   6
         Left            =   -74955
         TabIndex        =   37
         Top             =   360
         Width           =   9780
         _cx             =   17251
         _cy             =   8123
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   2
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
         Editable        =   2
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
         Height          =   4605
         Index           =   7
         Left            =   -74955
         TabIndex        =   38
         Top             =   360
         Width           =   9780
         _cx             =   17251
         _cy             =   8123
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   2
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
         Editable        =   2
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
         Height          =   4605
         Index           =   8
         Left            =   -74955
         TabIndex        =   39
         Top             =   360
         Width           =   9780
         _cx             =   17251
         _cy             =   8123
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   2
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
         Editable        =   2
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
   Begin VB.PictureBox SSPanel2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   10050
      TabIndex        =   8
      Top             =   0
      Width           =   10050
      Begin VB.CommandButton CmdExit 
         Height          =   465
         Left            =   810
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security2.frx":0118
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdSave 
         Height          =   465
         Left            =   4230
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security2.frx":0262
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Õ›Ÿ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         Height          =   465
         Left            =   3090
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security2.frx":06A4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   465
         Left            =   1950
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security2.frx":2F3E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   465
         Left            =   5370
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security2.frx":54B7
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         Height          =   465
         Left            =   6510
         Picture         =   "Security2.frx":7A63
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   45
         Width           =   1140
      End
      Begin Threed.SSCommand cmdFix 
         Height          =   465
         Left            =   7920
         TabIndex        =   21
         Top             =   45
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   820
         _Version        =   196610
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
         Picture         =   "Security2.frx":A236
         Caption         =   " «⁄«œ… ÷»ÿ «·ﬁ«∆„…"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   285
      Left            =   45
      TabIndex        =   14
      Top             =   5895
      Visible         =   0   'False
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   585
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin VB.Frame Frame8 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   9135
      Width           =   3030
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   23
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "Security2.frx":C5F0
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security2.frx":E7C0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   24
         Top             =   135
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "Security2.frx":10908
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security2.frx":12AD0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1530
         TabIndex        =   25
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "Security2.frx":14C1F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security2.frx":16DFF
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2295
         TabIndex        =   26
         Top             =   135
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "Security2.frx":18F5A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security2.frx":1B116
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2490
      Left            =   4320
      TabIndex        =   9
      Top             =   855
      Width           =   5595
      Begin VB.TextBox xPassword 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2025
         MaxLength       =   40
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   2175
      End
      Begin VB.TextBox xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   270
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   945
         Width           =   3930
      End
      Begin VB.CheckBox xShow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "≈ŸÂ«— "
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
         Height          =   330
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   585
         Width           =   825
      End
      Begin VB.CheckBox XSUPPER 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Supervisor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   1590
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3420
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Tag             =   "2"
         Top             =   225
         Width           =   780
      End
      Begin MSDataListLib.DataCombo xbox 
         Height          =   315
         Left            =   270
         TabIndex        =   6
         Top             =   1665
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   270
         TabIndex        =   7
         Top             =   2025
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XBRANCH 
         Height          =   315
         Left            =   270
         TabIndex        =   5
         Top             =   1305
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰"
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
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2025
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1710
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·«”„"
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
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«·—ﬁ„"
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
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   300
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ﬂ·„… «·”—"
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
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   660
         Width           =   720
      End
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin VSFlex7Ctl.VSFlexGrid VSSHOP 
      Height          =   3075
      Left            =   90
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   945
      Width           =   4155
      _cx             =   7329
      _cy             =   5424
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
      Cols            =   5
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
   Begin VB.Label xRecords 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
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
      Height          =   465
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   9180
      Width           =   6630
   End
End
Attribute VB_Name = "Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim CardTable As New ADODB.Recordset, oSearch As New Search3
Const LoadMode = 1, DefineMode = 2
Dim cFilter As String
Dim con As New ADODB.Connection
Dim nUser As Integer
Dim conmdb As New ADODB.Connection, secmdb As New ADODB.Connection
Enum enUser
     Admin = 3
     Super = 2
     User = 1
End Enum
Sub Handlecontrols(nMode)
xCode.Tag = nMode
Dim nRecord As Long, nRecords As Long
retRecords xCode.text, nRecords, nRecord

If nMode = LoadMode Then
    xRecords.Caption = "”Ã· " & nRecord & " „‰ " & nRecords
Else
    xRecords.Caption = "«÷«›… ”Ã· " & (nRecords + 1)
End If

cmdAdd.Enabled = (nMode = LoadMode)
cmddel.Enabled = (nMode = LoadMode)

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1 And sCode = ""
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords And sCode = ""
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2 And sCode = ""
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2 And sCode = ""

xCode.Enabled = Not (nMode = LoadMode)
End Sub
Sub myDefine()
xCode.text = Val(Newflag("USERS", "CODE", con))
xStore.BoundText = ""
XBRANCH.BoundText = ""
XISBRANCH.Value = 0
If cBranch <> "00" Then xStore.BoundText = cBranchStore
xPassword.text = ""
xDescA.text = ""
XSUPPER.Value = 0

xBox.BoundText = ""
Handlecontrols DefineMode

myDefineGrd
MyLoadShop
End Sub
Sub myload()
xCode.text = CardTable!code
xPassword.text = CardTable!PassWord & ""
xDescA.text = CardTable!desca & ""
XBRANCH.BoundText = CardTable!branch & ""
xBox.BoundText = CardTable!BOX & ""

xStore.BoundText = CardTable!STORE & ""
XSUPPER.Value = IIf(CardTable!SUPPER, 1, 0)
XISBRANCH.Value = IIf(CardTable!isbranch, 1, 0)

myDefineGrd
grid1(grid1.Count - 1).TextMatrix(1, 2) = IIf(CardTable!Option1, 1, 0)
grid1(grid1.Count - 1).TextMatrix(2, 2) = IIf(CardTable!Option2, 1, 0)
grid1(grid1.Count - 1).TextMatrix(3, 2) = IIf(CardTable!Option3, 1, 0)
grid1(grid1.Count - 1).TextMatrix(4, 2) = IIf(CardTable!Option4, 1, 0)
grid1(grid1.Count - 1).TextMatrix(5, 2) = IIf(CardTable!Option5, 1, 0)
grid1(grid1.Count - 1).TextMatrix(6, 2) = IIf(CardTable!Option6, 1, 0)
grid1(grid1.Count - 1).TextMatrix(7, 2) = IIf(CardTable!Option7, 1, 0)
grid1(grid1.Count - 1).TextMatrix(8, 2) = IIf(CardTable!Option8, 1, 0)
grid1(grid1.Count - 1).TextMatrix(9, 2) = IIf(CardTable!Option9, 1, 0)
grid1(grid1.Count - 1).TextMatrix(10, 2) = IIf(CardTable!Option10, 1, 0)

If nUser = enUser.Admin Then
    grid1(grid1.Count - 1).TextMatrix(11, 2) = IIf(CardTable!Option11, 1, 0)
End If

myLoadGrd
MyLoadShop

Handlecontrols LoadMode
End Sub
Private Sub incase()
Set xBox.RowSource = data1
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"
If Not xBox.MatchedWithList Then
    xBox.BoundText = sBound
End If


sBound = xStore.BoundText
data3.ConnectionString = strCon

If XBRANCH.BoundText = "" Then
    If cBranch = "00" Or lIsBranchStore Then
        Set data3.Recordset = myRecordSet("SELECT * FROM FILE0_40 WHERE BRANCH IS NULL ", con)
    Else
        Set data3.Recordset = myRecordSet("SELECT * FROM FILE0_40 ", con)
    End If
Else
    Set data3.Recordset = myRecordSet("SELECT * FROM FILE0_40 WHERE BRANCH = " & MyParn(XBRANCH.BoundText), con)
End If

Set xStore.RowSource = data3
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

If Not xStore.MatchedWithList Then
    xStore.BoundText = sBound
End If
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant

aInsert = AddFlag(Empty, "[password]", addstring(UCase(xPassword.text)))
aInsert = AddFlag(aInsert, "[desca]", addstring(xDescA.text))
aInsert = AddFlag(aInsert, "[Box]", addstring(xBox.BoundText))
aInsert = AddFlag(aInsert, "[STORE]", addstring(xStore.BoundText))

aInsert = AddFlag(aInsert, "[option1]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(1, 2)) = 0, "0", "1"))
aInsert = AddFlag(aInsert, "[option2]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(2, 2)) = 0, "0", "1"))
aInsert = AddFlag(aInsert, "[option3]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(3, 2)) = 0, "0", "1"))
aInsert = AddFlag(aInsert, "[option4]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(4, 2)) = 0, "0", "1"))
aInsert = AddFlag(aInsert, "[option5]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(5, 2)) = 0, "0", "1"))
aInsert = AddFlag(aInsert, "[option6]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(6, 2)) = 0, "0", "1"))

aInsert = AddFlag(aInsert, "[option7]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(7, 2)) = 0, "0", "1"))

aInsert = AddFlag(aInsert, "[option8]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(8, 2)) = 0, "0", "1"))
aInsert = AddFlag(aInsert, "[option9]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(9, 2)) = 0, "0", "1"))
aInsert = AddFlag(aInsert, "[option10]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(10, 2)) = 0, "0", "1"))

If cBranch = "00" Then
    aInsert = AddFlag(aInsert, "[option11]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(11, 2)) = 0, "0", "1"))
End If

aInsert = AddFlag(aInsert, "[SUPPER]", IIf(XSUPPER.Value, 1, 0))
aInsert = AddFlag(aInsert, "[BRANCH]", addstring(XBRANCH.BoundText))
aInsert = AddFlag(aInsert, "[isbranch]", IIf(XISBRANCH.Value, 1, 0))

con.BeginTrans
On Error GoTo myError
If xCode.Enabled Then
    Dim nCode As Long
    aInsert = AddFlag(aInsert, "code", xCode.text)
    con.Execute addInsert(aInsert, "users")
Else
    con.Execute addUpdate(aInsert, "users", " code = " & xCode.text)
End If
myreplaceGrd
MyReplaceShop
con.CommitTrans
myreplace = True
Exit Function
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Function myValid() As Boolean
If Not IsNumeric(xCode.text) Then
    MsgBox "—ﬁ„ «·„” Œœ„ €Ì— ”·Ì„"
    Exit Function
End If

If xDescA.text = "" Then
    MsgBox "«”„ «·„” Œœ„ €Ì— „”Ã·"
    Exit Function
End If

If xPassword.text = "" Then
    MsgBox "ﬂ·„… «·”— €Ì— „”Ã·…"
    Exit Function
End If
myValid = True
End Function
Private Sub CmdAdd_Click()
    myDefine
    xCode.SetFocus
End Sub
Private Sub cmdDate_Click()
Dim dDate As String
dDate = Format(InputBox("", " €Ì—  «—ÌŒ «·„»Ì⁄« ", dSalDate), "DD-MM-YYYY")
If IsDate(dDate) Then
    CD = "”Ê› Ì „  €Ì— «· «—ÌŒ «·Õ«·Ï " & Format(dSalesDate, "DD-MM-YYYY") & " ≈·Ï " & Format(dDate, "DD-MM-YYYY")
    If MsgBox(CD, vbOKCancel + vbDefaultButton2) = vbOK Then
        dSalesDate = dDate
        If lIsBranchStore Then
            con.Execute " UPDATE DSALES SET DSALES = " & addDate(dSalesDate) & " WHERE BRANCH = " & MyParn(cBranch), nRec
            If nRec = 0 Then con.Execute " INSERT DSALES (DSALES , BRANCH ) VALUES( " & addDate(dSalesDate) & " , " & addstring(cBranch) & ")"
        Else
            con.Execute " UPDATE DSALES SET DSALES = " & addDate(dSalesDate)
        End If
        Firsttitle = "JUONOUR " & myFormat_p(dSalesDate)
        Main.Caption = Firsttitle
        cmdDate.Caption = myFormat_p(dSalesDate)
    Else
        Exit Sub
    End If
End If
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdDel_Click()
On Error GoTo myError
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.BeginTrans
    
    con.Execute "Delete  From menusetting  Where code = " & Val(xCode.text)
    con.Execute "Delete  From users  Where code = " & Val(xCode.text)
    con.Execute "Delete  From USERSHOP  Where code = " & Val(xCode.text)
    
    con.CommitTrans
    If Not openCardTable(tbMode.tbPrevious, xCode.text) Then
        If Not openCardTable(tbMode.tbFirst) Then
            myDefine
        End If
    End If
End If
Exit Sub
myError:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub CmdFirst_Click()
If Not openCardTable(tbMode.tbFirst) Then
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
If Not openCardTable Then
    myDefine
End If
End Sub
Private Sub CmdNext_Click()
If Not openCardTable(tbMode.tbNext, xCode.text) Then
    If Not openCardTable(tbMode.tblast) Then
        myDefine
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openCardTable(tbMode.tbPrevious, xCode.text) Then
    If Not openCardTable(tbMode.tbFirst) Then
        myDefine
    End If
End If
End Sub

Private Sub cmdFix_Click()
FixControl
If FillMenuFile Then MsgBox " „ ÷»ÿ «·ﬁ«∆„… »‰Ã«Õ"
FillMenu
Filloption
myload
End Sub

Private Sub CmdInform_Click()
    CardLookup
End Sub
Private Sub cmdSave_Click()
If Not myValid Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
If Not openCardTable(tbMode.tbFind, xCode.text) Then
    If Not openCardTable Then myDefine
End If
End Sub
Private Sub CmdUndo_Click()
If xCode.Tag = DefineMode Then
    If Not openCardTable Then
        CmdAdd_Click
    End If
Else
    If Not openCardTable(tbMode.tbFind, xCode.text) Then
        If Not openCardTable Then
            myDefine
        End If
    End If
End If
End Sub
Private Sub Command1_Click()
CopyData.Show 1
Exit Sub
FixControl
If FillMenuFile Then MsgBox " „ ÷»ÿ «·ﬁ«∆„… »‰Ã«Õ"
FillMenu
Filloption
myload
End Sub
Private Sub Form_Load()
Dim ShopTable As New ADODB.Recordset
On Error GoTo myError
openCon con
If bSupermode And cBranch = "00" Then
    nUser = enUser.Admin
Else
    nUser = enUser.User
End If

FixList1
FixList2

Set xBox.RowSource = data1
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"

Set xStore.RowSource = data3
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

'con.Execute " UPDATE users SET option7 = 0 WHERE option7 IS NULL "
'con.Execute " UPDATE users SET option8 = 0 WHERE option8 IS NULL "

cmdDate.Caption = myFormat_p(dSalesDate)
'If GetDesca("select id from [option] where id = 9  ", con) = "" Then con.Execute "   insert INTO [option] (desca) VALUES        ('’·«ÕÌ… › Õ „»»Ì⁄«  ”«»ﬁ…')"
'If GetDesca("select id from [option] where id = 10 ", con) = "" Then con.Execute "   insert INTO [option] (desca) VALUES        ('’·«ÕÌ…  ﬂ·›… „»»Ì⁄«  ')"

Set data4.Recordset = cmd("SELECT * FROM BRANCH WHERE ISBRANCH = 1 ORDER BY DESCA", con).Execute

Set XBRANCH.RowSource = data4
XBRANCH.ListField = "Desca"
XBRANCH.BoundColumn = "Code"

Set VsShop.DataSource = DATA5

For ntab = 0 To SSTab1.Tabs - 1
    grid1(ntab).Cols = 5
    grid1(ntab).ColWidth(1) = 5000
    grid1(ntab).ColWidth(2) = 1000
    grid1(ntab).ColWidth(3) = 1000
    grid1(ntab).ColDataType(2) = flexDTBoolean
    grid1(ntab).ColDataType(3) = flexDTBoolean
    grid1(ntab).TextMatrix(0, 1) = "«·»‰œ"
    grid1(ntab).TextMatrix(0, 2) = "≈ŸÂ«—"
    grid1(ntab).TextMatrix(0, 3) = " ⁄œÌ·"
    grid1(ntab).ColHidden(0) = True
    grid1(ntab).ColHidden(grid1(ntab).Cols - 1) = True
    grid1(ntab).ColHidden(3) = True
    For i = 0 To 3
        grid1(ntab).ColAlignment(i) = flexAlignRightCenter
    Next
    grid1(ntab).Cell(flexcpChecked, 0, 2, 0, grid1(ntab).Cols - 1) = 2
    
Next

FillMenu
Filloption

'If cBranch = "00" Then
'    ShopTable.Open "SELECT * FROM STORE_BR ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
'Else
'    ShopTable.Open "SELECT * FROM file0_40 where branch = " & MyParn(cBranch), con, adOpenStatic, adLockReadOnly, adCmdText
'End If


If Not openCardTable Then myDefine


'If nUser = enUser.user Then
'    CardTable.Open "SELECT * FROM users WHERE BRANCH = " & MyParn(cBranch), con, adOpenStatic, adLockReadOnly, adCmdText
'Else
'    CardTable.Open "users", con, adOpenStatic, adLockReadOnly, adCmdTable
'End If

Exit Sub
myError:
End Sub
Private Function openCardTable(Optional pMode As Integer = tbMode.tblast, Optional pCode As String = "", Optional bDefine As Boolean = True) As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pCode) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

On Error GoTo myError
cFilter = ""

If lServerOnLine Then
    If nUser = enUser.User Then
        If cBranch = "00" Then
            cFilter = "(BRANCH = " & MyParn(cBranch) & " OR BRANCH IS NULL" & ")"
        Else
            cFilter = "BRANCH = " & MyParn(cBranch)
        End If
    End If
End If

If sCode <> "" Then cFilter = "USERS.CODE = " & sCode

cString = "SELECT TOP 1 * " & _
          " FROM USERS"

If pMode = tbMode.tbFirst Then
    cOrder = "Order by USERS.CODE"
ElseIf pMode = tbMode.tblast Then
    cOrder = "Order by USERS.CODE DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "CODE = " & pCode
ElseIf pMode = tbMode.tbPrevious Then
    cWhere = "CODE < " & pCode
    cOrder = "Order by CODE desc"
ElseIf pMode = tbMode.tbNext Then
    cWhere = "CODE > " & pCode
    cOrder = "Order by CODE"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

If cFilter <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilter
End If

cString = cString & " " & cOrder

Set CardTable = mycmd(cString, con, adText, , 600)

If (Not CardTable.EOF) Then
    myload
    openCardTable = True
End If
Me.MousePointer = 0
Exit Function
myError:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function

Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub

Private Sub Grid1_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
If Row = 0 Then
    With grid1(Index)
    For i = 1 To .Rows - 1
        .TextMatrix(i, Col) = IIf(grid1(Index).Cell(flexcpChecked, 0, Col) = 1, -1, 0)
    Next
    End With
End If
End Sub

Private Sub VSSHOP_EnterCell()
If VsShop.Col = VsShop.Cols - 1 Then
    VsShop.Editable = flexEDKbdMouse
Else
    VsShop.Editable = flexEDNone
End If
End Sub

Private Sub xbox_GotFocus()
'FixList1 XBRANCH.BoundText
End Sub

Private Sub xbox_LostFocus()
If Not xBox.MatchedWithList Then xBox.BoundText = ""
'FixList1
End Sub
Private Sub XBRANCH_Change()
FixList1 XBRANCH.BoundText
FixList2 XBRANCH.BoundText
End Sub

Private Sub xShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
xPassword.PasswordChar = IIf(xShow.Value, "", "*")
End Sub
Private Sub XSTORE_GotFocus()
'FixList2 XBRANCH.BoundText
End Sub
Private Sub XSTORE_LostFocus()
'FixList2
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xcode_LostFocus()
myLostFocus xCode
If Not ValidNum(xCode.text) Then
    xCode.text = ""
     If xCode.Tag = LoadMode Then
        myDefine
    End If
ElseIf Not openCardTable(tbMode.tbFind, xCode.text) Then
   If xCode.Tag = LoadMode Then
        myDefine
   End If
End If
End Sub
Private Sub myDefineGrd()
For ntab = 0 To SSTab1.Tabs - 1
     With grid1(ntab)
        For i = 1 To .Rows - 1
            .TextMatrix(i, 2) = 0
            .TextMatrix(i, 3) = 0
            .TextMatrix(i, 4) = ""
        Next
     End With
Next
End Sub
Private Sub FillMenu()
Dim GrdTable As New ADODB.Recordset
Set GrdTable = cmd("Select * from menu  Order by [Order]", con).Execute

If GrdTable.EOF And GrdTable.BOF Then Exit Sub

For ntab = 0 To 8
    With grid1(ntab)
        .Rows = 1
        grid1(ntab).Cell(flexcpChecked, 0, 2, 0, grid1(ntab).Cols - 1) = 2
        GrdTable.Filter = " menuNo = " & ntab + 1
        Do Until GrdTable.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = GrdTable!Control & ""
            .TextMatrix(.Rows - 1, 1) = GrdTable!desca & ""
            .TextMatrix(.Rows - 1, 2) = 0
            .TextMatrix(.Rows - 1, 3) = 0
            GrdTable.MoveNext
        Loop
    End With
Next
GrdTable.Close
Set GrdTable = Nothing
End Sub
Private Sub Filloption()
Dim GrdTable As New ADODB.Recordset
Set GrdTable = cmd("select * from [Option] order by id", con, adText).Execute
If Not GrdTable.EOF Then
    With grid1(SSTab1.Tabs - 1)
        .Rows = 1
        Do Until GrdTable.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = ""
            .TextMatrix(.Rows - 1, 1) = GrdTable!desca
            .TextMatrix(.Rows - 1, 2) = 0
            .TextMatrix(.Rows - 1, 3) = 0
            GrdTable.MoveNext
        Loop
    End With
End If
'Inform Grid1(8).Rows
GrdTable.Close
Set GrdTable = Nothing
End Sub
Private Sub myLoadGrd()
Dim GrdTable As New ADODB.Recordset
Set GrdTable = cmd("select * From menusetting where code = " & Val(xCode.text), con).Execute
If GrdTable.EOF And GrdTable.BOF Then Exit Sub
For ntab = 0 To SSTab1.Tabs - 2
    With grid1(ntab)
    For i = 1 To .Rows - 1
        GrdTable.Find "control = " & MyParn(.TextMatrix(i, 0)), , adSearchForward, adBookmarkFirst
        If Not GrdTable.EOF Then
            .TextMatrix(i, 2) = IIf(GrdTable!Visible, -1, 0)
            .TextMatrix(i, 3) = IIf(GrdTable!Editable, -1, 0)
            .TextMatrix(i, .Cols - 1) = GrdTable!ID
        End If
    Next
    End With
Next
End Sub
Private Sub myreplaceGrd()
Dim cString As New ChilkatStringBuilder

con.Execute "delete from menuSetting where code = " & xCode.text

cString.Append "INSERT INTO MENUSETTING(" & _
                "CODE," & _
                "CONTROL," & _
                "VISIBLE," & _
                "EDITABLE" & _
                ")"
cString.Append " VALUES "
With grid1
For ntab = 0 To grid1.Count - 2
    With grid1(ntab)
        For i = 1 To .Rows - 1
            cString.Append "("
            cString.Append addvalue(xCode.text) & ","
            cString.Append addstring(.TextMatrix(i, 0)) & ","
            cString.Append IIf(.ValueMatrix(i, 2) = 0, "0", "1") & ","
            cString.Append IIf(.ValueMatrix(i, 3) = 0, "0", "1")
            cString.Append ")"
            cString.Append ","
        Next
    End With
Next
cString.Shorten 1
con.Execute cString.GetAsString
End With
End Sub
Private Function FillMenuFile() As Boolean
Dim MenuNo As Integer, Order As Integer
On Error GoTo myError
con.BeginTrans
con.Execute "Delete  from menu"
For i = 0 To Main.Count - 1
     If TypeOf Main(i) Is Menu Then
        If Mid(Main(i).Name, 1, 2) = "mn" Then
           MainMenu = Main(i).Name
           Order = 1
           MenuNo = MenuNo + 1
        Else
            con.Execute "insert into menu(control,Desca,menuNo,[order],mainmenu)" & _
                           " values(" & _
                           addstring(Main(i).Name) & "," & _
                           addstring(Main(i).Caption) & "," & _
                           MenuNo & "," & _
                           Order & "," & _
                           addstring(MainMenu) & _
                            ")"
            Order = Order + 1
        End If
    End If
Next
con.CommitTrans
FillMenuFile = True
Exit Function
myError:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Function
Private Function FixControl() As Boolean
Dim loctable As New ADODB.Recordset
Dim cString As String
cString = "select menusetting.Id " & _
           " from menusetting left join menu on menu.control = menusetting.control " & _
            "where (menu.control is null)"
Set loctable = cmd(cString, con).Execute
On Error GoTo myError
con.BeginTrans
Do Until loctable.EOF
    con.Execute "delete  from menusetting where ID = " & loctable!ID
    loctable.MoveNext
Loop
con.CommitTrans
loctable.Close
Set loctable = Nothing
Exit Function
myError:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Function
Sub MyLoadShop()
Dim cString As String
cString = "SELECT CODE,DESCA,CASE WHEN [OK] IS NULL 0 ELSE USERSHOP = [OK] "
If cBranch = "00" Then
    cString = "SELECT STORE_BR.CODE," & _
              "STORE_BR.DESCA," & _
              "CASE WHEN [OK] IS NULL THEN 0 ELSE USERSHOP.[OK]  END" & _
              " FROM STORE_BR LEFT JOIN USERSHOP ON STORE_BR.CODE = USERSHOP.STORE" & _
              " AND USERSHOP.CODE = " & addvalue(xCode.text)
Else
    cString = "SELECT FILE0_40.CODE," & _
              "FILE0_40.DESCA," & _
              "CASE WHEN [OK] IS NULL THEN 0 ELSE USERSHOP.[OK]  END" & _
              " FROM FILE0_40 LEFT JOIN USERSHOP ON FILE0_40.CODE = USERSHOP.STORE" & _
              " AND USERSHOP.CODE = " & addvalue(xCode.text)
End If
Set DATA5.Recordset = cmd(cString, con).Execute
fixGrdShop
End Sub
Sub MyReplaceShop()
    With VsShop
        con.Execute " DELETE FROM USERSHOP WHERE CODE = " & xCode.text
        For nRow = 1 To .Rows - 1
            If TurnValue(.TextMatrix(nRow, 2), "", False) Then
                cStr1 = " INSERT INTO USERSHOP (STORE, CODE, [OK]) " & _
                        " VALUES     ( " & addstring(.TextMatrix(nRow, 0)) & " , " & xCode.text & "  , 1   ) "
                con.Execute cStr1
            End If
        Next nRow
    End With
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From users"
'If lIsBranchStore Then Generalarray(1) = Generalarray(1) & " WHERE BRANCH = " & MyParn(cBranch)
Generalarray(2) = "Order by desca "
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "»ÕÀ"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·≈”„"
GrdArray(1, 1) = 4000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„ "
oSearch.Show 1
End Sub
Sub myProc()
openCardTable tbMode.tbFind, oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
Unload oSearch
End Sub
Private Sub FixList1(Optional pBranch As String)
Dim cString As String
Dim sBound As String
sBound = xBox.BoundText
cString = "SELECT * FROM FILE0_50"
If pBranch <> "" Then
    cString = cString & " WHERE BRANCH = " & MyParn(pBranch)
End If

If data1.Recordset Is Nothing Then
    Set data1.Recordset = cmd(cString, con).Execute
ElseIf data1.Recordset.Source <> cString Then
    Set data1.Recordset = cmd(cString, con).Execute
Else
    Exit Sub
End If
If sBound <> "" Then
    xBox.BoundText = sBound
    If Not xBox.MatchedWithList Then xBox.BoundText = ""
End If

End Sub
Private Sub FixList2(Optional pBranch As String)
Dim cString As String
Dim sBound As String

sBound = xStore.BoundText
cString = "SELECT * FROM FILE0_40"
If pBranch <> "" Then
    cString = cString & " WHERE BRANCH = " & MyParn(pBranch)
End If

If data3.Recordset Is Nothing Then
    Set data3.Recordset = cmd(cString, con).Execute
ElseIf data3.Recordset.Source <> cString Then
    Set data3.Recordset = cmd(cString, con).Execute
Else
    Exit Sub
End If
If sBound <> "" Then
    xStore.BoundText = sBound
    If Not xStore.MatchedWithList Then xBox.BoundText = ""
End If
End Sub

Private Function retRecords(pCode, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If ValidNum(pCode) Then
    cString = "SELECT Count(*) AS records,COUNT(CASE WHEN CODE <= " & pCode & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM USERS " & Tr(cFilter, " WHERE ") & cFilter

Set loctable = mycmd(cString, con)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function
Private Sub fixGrdShop()
With VsShop
.TextMatrix(0, 0) = "—ﬁ„"
.TextMatrix(0, 1) = "«·›—⁄"
.TextMatrix(0, 2) = "⁄—÷"
.ColWidth(0) = 0
.ColWidth(1) = 2400
.ColWidth(2) = 600
.ColDataType(2) = flexDTBoolean
End With
End Sub
