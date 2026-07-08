VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin VB.CheckBox Xisbranch 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "„” Œœ„ ·ﬂ· «·›—Ê⁄"
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
      Left            =   8010
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   3240
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
      Height          =   780
      Left            =   4365
      Picture         =   "Security.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   3240
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5010
      Left            =   90
      TabIndex        =   6
      ToolTipText     =   "450"
      Top             =   4095
      Width           =   9870
      _ExtentX        =   17410
      _ExtentY        =   8837
      _Version        =   393216
      Tabs            =   9
      TabsPerRow      =   9
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
      TabPicture(0)   =   "Security.frx":321D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Grid1(8)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Œœ„« "
      TabPicture(1)   =   "Security.frx":3239
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid1(7)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "«· ﬁ«—Ì—"
      TabPicture(2)   =   "Security.frx":3255
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Grid1(6)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "«·»‰Êﬂ"
      TabPicture(3)   =   "Security.frx":3271
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Grid1(5)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "‰ﬁœÌ…"
      TabPicture(4)   =   "Security.frx":328D
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Grid1(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "›Ê« Ì—"
      TabPicture(5)   =   "Security.frx":32A9
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Grid1(3)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "„Ê—œÌ‰"
      TabPicture(6)   =   "Security.frx":32C5
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Grid1(2)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "⁄„·«¡"
      TabPicture(7)   =   "Security.frx":32E1
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Grid1(1)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "√’‰«›"
      TabPicture(8)   =   "Security.frx":32FD
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Grid1(0)"
      Tab(8).ControlCount=   1
      Begin VSFlex7LCtl.VSFlexGrid Grid1 
         Height          =   4455
         Index           =   2
         Left            =   -74865
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   450
         Width           =   9600
         _cx             =   16933
         _cy             =   7858
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
         Height          =   4455
         Index           =   3
         Left            =   -74895
         TabIndex        =   8
         Top             =   450
         Width           =   9645
         _cx             =   17013
         _cy             =   7858
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
         Height          =   4500
         Index           =   4
         Left            =   -74910
         TabIndex        =   9
         Top             =   405
         Width           =   9690
         _cx             =   17092
         _cy             =   7937
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
         Height          =   4455
         Index           =   5
         Left            =   -74895
         TabIndex        =   10
         Top             =   450
         Width           =   9645
         _cx             =   17013
         _cy             =   7858
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
         Height          =   4455
         Index           =   6
         Left            =   -74895
         TabIndex        =   11
         Top             =   450
         Width           =   9645
         _cx             =   17013
         _cy             =   7858
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
         Height          =   4455
         Index           =   7
         Left            =   -74895
         TabIndex        =   12
         Top             =   450
         Width           =   9645
         _cx             =   17013
         _cy             =   7858
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
         Height          =   4455
         Index           =   8
         Left            =   105
         TabIndex        =   13
         Top             =   450
         Width           =   9645
         _cx             =   17013
         _cy             =   7858
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
         Height          =   4455
         Index           =   1
         Left            =   -74850
         TabIndex        =   14
         Top             =   450
         Width           =   9600
         _cx             =   16933
         _cy             =   7858
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
         Height          =   4410
         Index           =   0
         Left            =   -74850
         TabIndex        =   15
         Top             =   450
         Width           =   9600
         _cx             =   16933
         _cy             =   7779
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
      TabIndex        =   1
      Top             =   0
      Width           =   10050
      Begin VB.CommandButton CmdExit 
         Height          =   465
         Left            =   810
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security.frx":3319
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Picture         =   "Security.frx":3463
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Õ›Ÿ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         Height          =   465
         Left            =   3090
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Security.frx":38A5
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Picture         =   "Security.frx":613F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "Security.frx":86B8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         Height          =   465
         Left            =   6510
         Picture         =   "Security.frx":AC64
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   45
         Width           =   1140
      End
      Begin Threed.SSCommand cmdFix 
         Height          =   465
         Left            =   7920
         TabIndex        =   23
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
         Picture         =   "Security.frx":D437
         Caption         =   " «⁄«œ… ÷»ÿ «·ﬁ«∆„…"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   285
      Left            =   45
      TabIndex        =   16
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
      TabIndex        =   26
      Top             =   9135
      Width           =   3030
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   27
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
         Picture         =   "Security.frx":F7F1
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security.frx":119C1
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   825
         TabIndex        =   28
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
         Picture         =   "Security.frx":13B09
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security.frx":15CD1
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1530
         TabIndex        =   29
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
         Picture         =   "Security.frx":17E20
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security.frx":1A000
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2295
         TabIndex        =   30
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
         Picture         =   "Security.frx":1C15B
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Security.frx":1E317
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
      Height          =   2625
      Left            =   4365
      TabIndex        =   2
      Top             =   540
      Width           =   5595
      Begin VB.TextBox xPassword 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   2070
         MaxLength       =   40
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   315
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   975
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
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   540
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3465
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   780
      End
      Begin MSDataListLib.DataCombo xbox 
         Height          =   315
         Left            =   315
         TabIndex        =   31
         Top             =   1350
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
         Left            =   315
         TabIndex        =   33
         Top             =   1710
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
         Left            =   315
         TabIndex        =   39
         Top             =   2070
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
         Caption         =   "«·›—⁄ : "
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   2025
         Width           =   570
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰ :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1665
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰… :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1305
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·«”„ :"
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
         TabIndex        =   5
         Top             =   1065
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "«·—ﬁ„ :"
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
         TabIndex        =   4
         Top             =   300
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "ﬂ·„… «·”— :"
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
         TabIndex        =   3
         Top             =   660
         Width           =   840
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
   Begin VSFlex7LCtl.VSFlexGrid VsShop 
      Height          =   3420
      Left            =   90
      TabIndex        =   38
      Top             =   630
      Width           =   4200
      _cx             =   7408
      _cy             =   6032
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
End
Attribute VB_Name = "Security"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim formMode As Byte
Dim CardTable As New ADODB.Recordset, oSearch As New Search3
Const LoadMode = 1, DefineMode = 2
Dim con As New ADODB.Connection
Dim conmdb As New ADODB.Connection, secmdb As New ADODB.Connection
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xCode.Enabled = Not (nMode = LoadMode)
End Sub
Sub myDefine()
If cBranch = "00" Then
    xCode.text = Val(Newflag("USERS", "CODE", con))
Else
    xCode.text = IncRec(GetDesca("SELECT MAX(CODE) FROM USERS WHERE ISBRANCH = 0 ", con))
End If
xStore.BoundText = ""
XBRANCH.BoundText = ""
XISBRANCH.Value = 0
If cBranch <> "00" Then xStore.BoundText = cBranchStore
xPassword.text = ""
XDESCA.text = ""
XSUPPER.Value = 0

xBox.BoundText = ""
Handlecontrols DefineMode
myDefineGrd
End Sub
Sub myload()
xCode.text = CardTable!CODE
xPassword.text = CardTable!PassWord & ""
XDESCA.text = CardTable!DESCA & ""
xBox.BoundText = CardTable!BOX & ""
xStore.BoundText = CardTable!STORE & ""
XSUPPER.Value = IIf(CardTable!SUPPER, 1, 0)
XBRANCH.BoundText = CardTable!branch & ""
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
If cBranch = "00" Then grid1(grid1.Count - 1).TextMatrix(11, 2) = IIf(CardTable!Option11, 1, 0)

sBound = xBox.BoundText
If XBRANCH.BoundText = "" Then
    If cBranch = "00" Or lIsBranchStore Then
        Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_50 WHERE BRANCH IS NULL ", con)
    Else
        Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_50 ", con)
    End If
Else
    Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(XBRANCH.BoundText), con)
End If
Set xBox.RowSource = data1
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"
If Not xBox.MatchedWithList Then
    xBox.BoundText = sBound
End If


sBound = xStore.BoundText
DATA3.ConnectionString = strCon
If XBRANCH.BoundText = "" Then
    If cBranch = "00" Or lIsBranchStore Then
        Set DATA3.Recordset = myRecordSet("SELECT * FROM FILE0_40 WHERE BRANCH IS NULL ", con)
    Else
        Set DATA3.Recordset = myRecordSet("SELECT * FROM FILE0_40 ", con)
    End If
Else
    Set DATA3.Recordset = myRecordSet("SELECT * FROM FILE0_40 WHERE BRANCH = " & MyParn(XBRANCH.BoundText), con)
End If

Set xStore.RowSource = DATA3
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

If Not xStore.MatchedWithList Then
    xStore.BoundText = sBound
End If

myloadgrd
MyLOadShop


Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant

aInsert = AddFlag(Empty, "[password]", addstring(UCase(xPassword.text)))
aInsert = AddFlag(aInsert, "[desca]", addstring(XDESCA.text))
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
If cBranch = "00" Then aInsert = AddFlag(aInsert, "[option11]", IIf(Val(grid1(SSTab1.Tabs - 1).TextMatrix(11, 2)) = 0, "0", "1"))

aInsert = AddFlag(aInsert, "[SUPPER]", IIf(XSUPPER.Value, 1, 0))
aInsert = AddFlag(aInsert, "[BRANCH]", addstring(XBRANCH.BoundText))
aInsert = AddFlag(aInsert, "[isbranch]", IIf(XISBRANCH.Value, 1, 0))
con.BeginTrans

On Error GoTo myerror
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
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Function MYVALID() As Boolean
If Not IsNumeric(xCode.text) Then
    MsgBox "—ﬁ„ «·„” Œœ„ €Ì— ”·Ì„"
    Exit Function
End If
If XDESCA.text = "" Then
    MsgBox "«”„ «·„” Œœ„ €Ì— „”Ã·"
    Exit Function
End If

If xPassword.text = "" Then
    MsgBox "ﬂ·„… «·”— €Ì— „”Ã·…"
    Exit Function
End If
MYVALID = True
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
        Firsttitle = "JUONOUR " & Format(dSalesDate, "DD-MM-YYYY")
        Main.Caption = Firsttitle
        cmdDate.Caption = ArbString(Format(dSalesDate, "DD-MM-YYYY"))
    Else
        Exit Sub
    End If
End If
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.BeginTrans
    
    con.Execute "Delete  From menusetting  Where code = " & Val(xCode.text)
    con.Execute "Delete  From users  Where code = " & Val(xCode.text)
    con.Execute "Delete  From USERSHOP  Where code = " & Val(xCode.text)
    
    con.CommitTrans
    CardTable.Requery
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & Val(xCode.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        myDefine
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub CmdFirst_Click()
'On Error GoTo myError
CardTable.MoveFirst
myload
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

Private Sub CmdLast_Click()
'On Error GoTo myError
CardTable.MoveLast
myload
End Sub
Private Sub CmdNext_Click()
'On Error GoTo myError
CardTable.MoveNext
If CardTable.EOF Then CardTable.MovePrevious Else myload
Exit Sub
End Sub
Private Sub CmdPrevious_Click()
'On Error GoTo myError
CardTable.MovePrevious
If CardTable.BOF Then CardTable.MoveNext Else myload
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
CardTable.Requery
CardTable.Find "code = " & Val(xCode.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    myDefine
Else
    If xCode.Enabled Then
        CardTable.MoveLast
    Else
        CardTable.Find "code = " & Val(xCode.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    myload
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

Private Sub DataCombo1_Click(Area As Integer)

End Sub
Private Sub Form_Load()
Dim ShopTable As New ADODB.Recordset
On Error GoTo myerror
openCon con

'con.Execute " UPDATE users SET option7 = 0 WHERE option7 IS NULL "
'con.Execute " UPDATE users SET option8 = 0 WHERE option8 IS NULL "

If lIsBranchStore Then
    CardTable.Open "SELECT * FROM users WHERE BRANCH = " & MyParn(cBranch), con, adOpenStatic, adLockReadOnly, adCmdText
Else
    CardTable.Open "users", con, adOpenStatic, adLockReadOnly, adCmdTable
End If
cmdDate.Caption = ArbString(Format(dSalesDate, "DD-MM-YYYY"))
If GetDesca("select id from [option] where id = 9  ", con) = "" Then con.Execute "   insert INTO [option] (desca) VALUES        ('’·«ÕÌ… › Õ „»»Ì⁄«  ”«»ﬁ…')"
If GetDesca("select id from [option] where id = 10 ", con) = "" Then con.Execute "   insert INTO [option] (desca) VALUES        ('’·«ÕÌ…  ﬂ·›… „»»Ì⁄«  ')"


If cBranch = "00" Or Not lIsBranchStore Then
    Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_50", con)
Else
    Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch), con)
End If
Set xBox.RowSource = data1
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"

If cBranch = "00" Or Not lIsBranchStore Then
    Set DATA3.Recordset = myRecordSet("SELECT * FROM FILE0_40 ", con)
    DATA3.RecordSource = "SELECT * FROM FILE0_40 "
Else
    Set DATA3.Recordset = myRecordSet("SELECT * FROM FILE0_40 WHERE BRANCH = " & MyParn(cBranch), con)
End If
Set xStore.RowSource = DATA3
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

If lIsBranchStore Or cBranch = "00" Then
    Set DATA4.Recordset = myRecordSet("SELECT * FROM BRANCH WHERE ISBRANCH = 1 ORDER BY CODE ", con)
    Set XBRANCH.RowSource = DATA4
    XBRANCH.ListField = "Desca"
    XBRANCH.BoundColumn = "Code"
End If

If cBranch <> "00" Then
    xStore.BoundText = cBranchStore
'    XSTORE.Enabled = False
End If

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
    For I = 0 To 3
        grid1(ntab).ColAlignment(I) = flexAlignRightCenter
    Next
    grid1(ntab).Cell(flexcpChecked, 0, 2, 0, grid1(ntab).Cols - 1) = 2
    
Next
FillMenu
Filloption
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
End If

If cBranch = "00" Then
    ShopTable.Open "SELECT * FROM STORE_BR ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
Else
    ShopTable.Open "SELECT * FROM file0_40 where branch = " & MyParn(cBranch), con, adOpenStatic, adLockReadOnly, adCmdText
End If
With VsShop
    .Cols = 3
    .Rows = 1
    .FixedCols = 2
    .TextMatrix(0, 0) = "—ﬁ„"
    .TextMatrix(0, 1) = "«·›—⁄"
    .TextMatrix(0, 2) = "⁄—÷"
    .ColWidth(0) = 0
    .ColWidth(1) = 2400
    .ColWidth(2) = 600
    .ColDataType(2) = flexDTBoolean
    Do While Not ShopTable.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = ShopTable!CODE
        .TextMatrix(.Rows - 1, 1) = ShopTable!DESCA
        ShopTable.MoveNext
    Loop
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
End With
Exit Sub
myerror:
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub

Private Sub grid1_AfterEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long)
If Row = 0 Then
    With grid1(Index)
    For I = 1 To .Rows - 1
        .TextMatrix(I, Col) = IIf(grid1(Index).Cell(flexcpChecked, 0, Col) = 1, -1, 0)
    Next
    End With
End If
End Sub

Private Sub xBox_LostFocus()
If Not xBox.MatchedWithList Then xBox.BoundText = ""
End Sub
Private Sub xbranch_LostFocus()
If lIsBranchStore Or cBranch = "00" Then
    sBound = xBox.BoundText
    If XBRANCH.BoundText = "" Then
        Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_50 WHERE BRANCH IS NULL ", con)
    Else
        Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(XBRANCH.BoundText), con)
    End If
    Set xBox.RowSource = data1
    xBox.ListField = "Desca"
    xBox.BoundColumn = "Code"
    
    If Not xBox.MatchedWithList Then
        xBox.BoundText = sBound
    End If
    
    
    sBound = xStore.BoundText
    If XBRANCH.BoundText = "" Then
        Set DATA3.Recordset = myRecordSet("SELECT * FROM FILE0_40 WHERE BRANCH IS NULL ", con)
    Else
        Set DATA3.Recordset = myRecordSet("SELECT * FROM FILE0_40 WHERE BRANCH = " & MyParn(XBRANCH.BoundText), con)
    End If
    Set xStore.RowSource = DATA3
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    If Not xStore.MatchedWithList Then
        xStore.BoundText = sBound
    End If
End If
End Sub

Private Sub xShow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
xPassword.PasswordChar = IIf(xShow.Value, "", "*")
End Sub

Private Sub XSTORE_LostFocus()
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xCode_LostFocus()
If xCode.text = "" Then Exit Sub
CardTable.Find "code = " & xCode.text, , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub myDefineGrd()
For ntab = 0 To SSTab1.Tabs - 1
     With grid1(ntab)
        For I = 1 To .Rows - 1
            .TextMatrix(I, 2) = 0
            .TextMatrix(I, 3) = 0
            .TextMatrix(I, 4) = ""
        Next
     End With
Next
With VsShop
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 2) = False
    Next nRow
End With
End Sub
Private Sub FillMenu()
Dim GrdTable As New ADODB.Recordset
GrdTable.Open "Select * from menu  Order by [Order]", con, adOpenStatic, adLockReadOnly

If GrdTable.EOF And GrdTable.BOF Then Exit Sub

For ntab = 0 To 7
    With grid1(ntab)
        .Rows = 1
        grid1(ntab).Cell(flexcpChecked, 0, 2, 0, grid1(ntab).Cols - 1) = 2
        GrdTable.Filter = " menuNo = " & ntab + 1
        Do Until GrdTable.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = GrdTable!Control & ""
            .TextMatrix(.Rows - 1, 1) = GrdTable!DESCA & ""
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
GrdTable.Open "[Option]", con, adOpenStatic, adLockReadOnly, adCmdTable
If Not GrdTable.EOF Then
    With grid1(SSTab1.Tabs - 1)
        .Rows = 1
        Do Until GrdTable.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = ""
            .TextMatrix(.Rows - 1, 1) = GrdTable!DESCA
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
Private Sub myloadgrd()
Dim GrdTable As New ADODB.Recordset
GrdTable.Open "select * From menusetting where code = " & Val(xCode.text), con, adOpenStatic, adLockReadOnly, adCmdText
If GrdTable.EOF And GrdTable.BOF Then Exit Sub
For ntab = 0 To SSTab1.Tabs - 2
    With grid1(ntab)
    For I = 1 To .Rows - 1
        GrdTable.Find "control = " & MyParn(.TextMatrix(I, 0)), , adSearchForward, adBookmarkFirst
        If Not GrdTable.EOF Then
            .TextMatrix(I, 2) = IIf(GrdTable!Visible, -1, 0)
            .TextMatrix(I, 3) = IIf(GrdTable!Editable, -1, 0)
            .TextMatrix(I, .Cols - 1) = GrdTable!ID
        End If
    Next
    End With
Next
End Sub
Private Sub myreplaceGrd()
Dim aInsert(3, 1), cStrSec As String

For ntab = 0 To grid1.Count - 2
    With grid1(ntab)
        For I = 1 To .Rows - 1
            aInsert(0, 0) = "CODE"
            aInsert(0, 1) = Val(xCode.text)
            
            aInsert(1, 0) = "CONTROL"
            aInsert(1, 1) = addstring(.TextMatrix(I, 0))
            
            aInsert(2, 0) = "VISIBLE"
            aInsert(2, 1) = IIf(Val(.TextMatrix(I, 2)) = 0, "0", "1")
            
            aInsert(3, 0) = "EDITABLE"
            aInsert(3, 1) = IIf(Val(.TextMatrix(I, 3)) = 0, "0", "1")
                       
            If .TextMatrix(I, .Cols - 1) = "" Then
                cStrSec = cStrSec & CreateInsert(aInsert, "menuSetting") & " ; "
'                con.Execute CreateInsert(aInsert, "menuSetting")
            Else
                cStrSec = cStrSec & CreateUpdate(aInsert, "menuSetting", " WHERE ID = " & .TextMatrix(I, .Cols - 1), Array(-1)) & " ; "
'                con.Execute CreateUpdate(aInsert, "menuSetting", " WHERE ID = " & .TextMatrix(i, .Cols - 1), Array(-1))
            End If
        Next
        If cStrSec <> "" Then con.Execute cStrSec
    End With
Next
End Sub
Private Function FillMenuFile() As Boolean
Dim MenuNo As Integer, Order As Integer
On Error GoTo myerror
con.BeginTrans
con.Execute "Delete  from menu"
For I = 0 To Main.Count - 1
     If TypeOf Main(I) Is Menu Then
        If Mid(Main(I).Name, 1, 2) = "mn" Then
           MainMenu = Main(I).Name
           Order = 1
           MenuNo = MenuNo + 1
        Else
            con.Execute "insert into menu(control,Desca,menuNo,[order],mainmenu)" & _
                           " values(" & _
                           addstring(Main(I).Name) & "," & _
                           addstring(Main(I).Caption) & "," & _
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
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Function
Private Function FixControl() As Boolean
Dim loctable As New ADODB.Recordset
loctable.Open "select menusetting.Id from menusetting left join menu on menu.control = menusetting.control " & _
              "where (menu.control is null)", con, adOpenStatic, adLockReadOnly, adCmdText
On Error GoTo myerror
con.BeginTrans
Do Until loctable.EOF
    con.Execute "delete  from menusetting where ID = " & loctable!ID
    loctable.MoveNext
Loop
con.CommitTrans
loctable.Close
Set loctable = Nothing
Exit Function
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Function
Sub MyLOadShop()
    With VsShop
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, 2) = GetDesca("SELECT [OK] FROM USERSHOP WHERE CODE = " & xCode.text & " AND [STORE] =  " & MyParn(.TextMatrix(nRow, 0)), con)
        Next nRow
    End With
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

Generalarray(1) = "Select code ,DescA From users "
If lIsBranchStore Then Generalarray(1) = Generalarray(1) & " WHERE BRANCH = " & MyParn(cBranch)
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
   CardTable.Find "CODE = " & MyParn(oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
   myload
   Unload oSearch
End Sub

