VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form VsTSales_DISC_DOC 
   Caption         =   "≈Ã„«·Ï „»Ì⁄«  „ÊœÌ·«  ·› —… - ⁄—Ê÷ ›ﬁÿ"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   6225
      Left            =   135
      TabIndex        =   53
      Top             =   3285
      Width           =   15000
      _ExtentX        =   26458
      _ExtentY        =   10980
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "≈Ã„«·Ï „” ‰œ« "
      TabPicture(0)   =   "VsTSales_DISC_DOC.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GRID5"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "≈Ã„«·Ï ›—Ê⁄"
      TabPicture(1)   =   "VsTSales_DISC_DOC.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRID4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "≈Ã„«·Ï „Ê—œÌ‰"
      TabPicture(2)   =   "VsTSales_DISC_DOC.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GRID3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "≈Ã„«·Ï „ÊœÌ·«  - ⁄—Ê÷"
      TabPicture(3)   =   "VsTSales_DISC_DOC.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GRID2"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "≈Ã„«·Ï „ÊœÌ·« "
      TabPicture(4)   =   "VsTSales_DISC_DOC.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "grid1"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   5715
         Left            =   90
         TabIndex        =   55
         Top             =   405
         Width           =   14820
         _cx             =   26141
         _cy             =   10081
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
      Begin VSFlex7Ctl.VSFlexGrid GRID2 
         Height          =   5715
         Left            =   -74910
         TabIndex        =   56
         Top             =   405
         Width           =   14820
         _cx             =   26141
         _cy             =   10081
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
      Begin VSFlex7Ctl.VSFlexGrid GRID3 
         Height          =   5715
         Left            =   -74910
         TabIndex        =   57
         Top             =   405
         Width           =   14820
         _cx             =   26141
         _cy             =   10081
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
      Begin VSFlex7Ctl.VSFlexGrid GRID4 
         Height          =   5715
         Left            =   -74910
         TabIndex        =   58
         Top             =   405
         Width           =   14820
         _cx             =   26141
         _cy             =   10081
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
      Begin VSFlex7Ctl.VSFlexGrid GRID5 
         Height          =   5715
         Left            =   -74910
         TabIndex        =   59
         Top             =   405
         Width           =   14820
         _cx             =   26141
         _cy             =   10081
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
   Begin VB.Frame Frame2 
      Height          =   510
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   2430
      Width           =   3840
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
         Top             =   180
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "’Ê— «·„ÊœÌ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   3570
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   2085
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   3678
         BackColor       =   16777215
         BorderStyle     =   4
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         AutoZoomType    =   3
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "amr123"
         LicenseRegCode  =   "íß“µª∫≥ºΩ∫æ“±æß´¥Ω∏•OANH-FAZOHIZB-EFTP6gI"
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3285
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
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
         Left            =   1215
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1350
         Width           =   870
      End
      Begin VB.Frame Frame7 
         Height          =   555
         Left            =   7335
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   2655
         Width           =   3795
         Begin VB.OptionButton xonest 
            Alignment       =   1  'Right Justify
            Caption         =   "€Ì— «„«‰« "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   2295
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   180
            Width           =   1185
         End
         Begin VB.OptionButton xonest 
            Alignment       =   1  'Right Justify
            Caption         =   "«„«‰« "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   1
            Left            =   1215
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   180
            Width           =   825
         End
         Begin VB.OptionButton xonest 
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
            Height          =   285
            Index           =   2
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   47
            Top             =   180
            Value           =   -1  'True
            Width           =   780
         End
      End
      Begin VB.TextBox XMOSM2 
         Alignment       =   2  'Center
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
         Left            =   135
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   945
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   540
         Width           =   915
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»Ì⁄«  ⁄—÷ 2+1"
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox XCOUNT 
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
         TabIndex        =   32
         Top             =   2430
         Width           =   1095
      End
      Begin VB.TextBox xdoc_no 
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
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   180
         Width           =   1725
      End
      Begin VB.Frame Frame4 
         Height          =   690
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   2565
         Width           =   5595
         Begin VB.CommandButton CMD_PRINT 
            Height          =   465
            Left            =   1260
            Picture         =   "VsTSales_DISC_DOC.frx":008C
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   135
            Width           =   1185
         End
         Begin VB.CommandButton cmdExit 
            Height          =   465
            Left            =   45
            Picture         =   "VsTSales_DISC_DOC.frx":24B6
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   135
            Width           =   1185
         End
         Begin VB.CommandButton cmdGo 
            Height          =   465
            Left            =   2475
            Picture         =   "VsTSales_DISC_DOC.frx":4922
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "⁄—÷"
            Top             =   135
            Width           =   1275
         End
         Begin Threed.SSCommand cmd_excel 
            Height          =   465
            Left            =   3780
            TabIndex        =   28
            Top             =   135
            Width           =   1740
            _ExtentX        =   3069
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
            Picture         =   "VsTSales_DISC_DOC.frx":6E14
            Caption         =   " ÕÊÌ· «ﬂ”Ì· "
            Alignment       =   4
            PictureAlignment=   1
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
         TabIndex        =   12
         Top             =   225
         Width           =   1815
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
         Left            =   7935
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1500
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Top             =   1310
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
         TabIndex        =   4
         Top             =   953
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
         Left            =   2250
         TabIndex        =   5
         Top             =   573
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   6000
         TabIndex        =   6
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
         Left            =   2250
         TabIndex        =   15
         Top             =   1329
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   2250
         TabIndex        =   17
         Top             =   951
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xcust 
         Height          =   315
         Left            =   6000
         TabIndex        =   20
         Top             =   1667
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGrCust 
         Height          =   315
         Left            =   5985
         TabIndex        =   22
         Top             =   2025
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
         Left            =   2835
         TabIndex        =   38
         Top             =   1710
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xsex 
         Height          =   315
         Left            =   315
         TabIndex        =   41
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
      Begin MSDataListLib.DataCombo XDOC_DISC 
         Height          =   315
         Left            =   405
         TabIndex        =   52
         Top             =   2115
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Œ’„ ⁄—Ê÷ :"
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
         Index           =   10
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   2160
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
         Left            =   1035
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1035
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
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   585
         Width           =   1140
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
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1755
         Width           =   555
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
         Index           =   4
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«⁄·Ï ﬁÌ„… „»Ì⁄«  "
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
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2475
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "›« Ê—… —ﬁ„ :"
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… „Ê—œÌ‰:"
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "«·⁄„Ì· :"
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
         TabIndex        =   21
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Label xModelNo 
         Alignment       =   1  'Right Justify
         Caption         =   "Label6"
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Visible         =   0   'False
         Width           =   645
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
         Index           =   0
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1020
         Width           =   1230
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1395
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
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
         TabIndex        =   11
         Top             =   270
         Width           =   765
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
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
   Begin MSAdodcLib.Adodc DATA21 
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
   Begin MSAdodcLib.Adodc data8 
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
   Begin MSAdodcLib.Adodc data9 
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
   Begin MSAdodcLib.Adodc data10 
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
   Begin MSAdodcLib.Adodc Adodc2 
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
   Begin MSAdodcLib.Adodc DATA11 
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
   Begin MSAdodcLib.Adodc DATA12 
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
   Begin ComctlLib.ProgressBar prog1 
      Height          =   285
      Left            =   45
      TabIndex        =   54
      Top             =   2970
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc DATA22 
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
   Begin MSAdodcLib.Adodc DATA23 
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
   Begin MSAdodcLib.Adodc DATA24 
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
   Begin MSAdodcLib.Adodc DATA25 
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
Attribute VB_Name = "VsTSales_DISC_DOC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3, Search31 As New Search3
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
    If SSTab1.Tab = 4 Then ToFileExel2 grid1, , , , , 1.1, , , , , , Me
    If SSTab1.Tab = 3 Then ToFileExel2 GRID2, , , , , 1.1, , , , , , Me
    If SSTab1.Tab = 2 Then ToFileExel2 GRID3, , , , , 1.1, , , , , , Me
    If SSTab1.Tab = 1 Then ToFileExel2 GRID4, , , , , 1.1, , , , , , Me
    If SSTab1.Tab = 0 Then ToFileExel2 grid5, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = "≈Ã„«·Ï „»Ì⁄«  „ÊœÌ·«  ·› —… " & XSTORE.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY") & " Õ Ï  «—ÌŒ " & Format(xdate2.text, "DD-MM-YYYY")
    cHead3 = ""
    If xFact.BoundText <> "" Then cHead3 = cHead3 & " „’‰⁄ " & xFact.text
    If xSupp.BoundText <> "" Then cHead3 = cHead3 & " „Ê—œ " & xSupp.text & "  " & xGrCust.text
    If xGroup.BoundText <> "" Then cHead3 = cHead3 & " „Ã„Ê⁄…" & xGroup.text & "  " & xage.text & "  " & xsex.text
    If xSection.text <> "" Then cHead3 = cHead3 & " ﬁ”„ " & xSection.text
    
    If xMosm.BoundText <> "" Then cHead3 = cHead3 & " „Ê”„ " & xMosm.text
    
    PrintGrdNew.doprint grid1, 1, , cHead1, cHead2, cHead3, , False, 8
    PrintGrdNew.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
'    On Error GoTo myerror
    MYLOAD1
    MYLOAD2
    MYLOAD3
    MYLOAD4
    myload5
    Exit Sub
myerror:
    MsgBox Err.Description
End Sub
Private Sub Form_Load()
    openCon con
    
    xmosm22.AddItem " "
    xmosm22.AddItem "S"
    xmosm22.AddItem "W"
    xmosm22.AddItem "M"
    
    Set DATA11.Recordset = myRecordSet("Select Code,DescA From File0_82  order by Desca", con)
    Set xage.RowSource = DATA11
    xage.ListField = "Desca"
    xage.BoundColumn = "Code"
    
    Set DATA12.Recordset = myRecordSet("Select Code,DescA From File0_81  order by Desca", con)
    Set xsex.RowSource = DATA12
    xsex.ListField = "Desca"
    xsex.BoundColumn = "Code"
    
    Set data9.Recordset = myRecordSet("Select Code,DescA From File3_10  order by Desca", con)
    Set xcust.RowSource = data9
    xcust.ListField = "Desca"
    xcust.BoundColumn = "Code"
    
    Set data1.Recordset = myRecordSet("SELECT * FROM FILE4_50 ", con)
    Set xGrCust.RowSource = data1
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    Set data8.Recordset = myRecordSet("Select Code,DescA From File1_10SC order by Desca", con)
    Set xSection.RowSource = data8
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    Set DATA7.Recordset = myRecordSet("Select Code,DescA From STORE_BR WHERE " & cUserStore, con)
    Set XSTORE.RowSource = DATA7
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    If cBranch <> "00" Then
        XSTORE.BoundText = cBranchStore
        XSTORE.Enabled = False
    End If
    
    Set DATA2.Recordset = myRecordSet("Select Code,DescA From File4_10 order by Desca", con)
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    Set data3.Recordset = myRecordSet("Select Code,DescA From File1_50 ORDER BY DESCA", con)
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set data4.Recordset = myRecordSet("Select mosm ,descA From mosm ORDER BY date DESC ", con)
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    
    
    Set data5.Recordset = myRecordSet("Select code ,desca From fact ORDER BY DESCA ", con)
    Set xFact.RowSource = data5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    Set DATA13.Recordset = myRecordSet("Select DOC_NO ,DescA From File0_90H order by DOC_NO DESC ", con)
    Set xDOC_DISC.RowSource = DATA13
    xDOC_DISC.ListField = "Desca"
    xDOC_DISC.BoundColumn = "DOC_NO"
    
    Set grid1.DataSource = DATA21
    Set GRID2.DataSource = DATA22
    Set GRID3.DataSource = DATA23
    Set GRID4.DataSource = DATA24
    Set grid5.DataSource = data25

    FixGrid1
    FixGrid2

End Sub
Private Sub MYLOAD1()
Dim i As Double
Dim cString  As String, cStr2 As String
With grid1
'                           0               1                 2                3               4                5
    
    If Val(XCOUNT.text) > 0 Then
        cString = " SELECT TOP " & Val(XCOUNT.text) & "  "
    Else
        cString = " SELECT "
    End If
    cString = cString & " ' ' AS N1 , FACT.desca , file1_10.ModelNo , file1_10.mosm , file1_10.MODELFACT0, file1_10.desca ,   " & _
            " SUM(QFILE6_20.QUANT) AS t_q, SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS t_sales, SUM(QFILE6_20.QUANT * QFILE6_20.cost) AS t_cost, SUM((QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100))  - (QFILE6_20.QUANT * QFILE6_20.cost)) AS t_proft , ' ' , SUM((QFILE6_20.QUANT * QFILE6_20.price_c )-QFILE6_20.total ) , SUM(QFILE6_20.TOTAL * ((QFILE6_20.RATE )/100)) , SUM(QFILE6_20.QUANT * QFILE6_20.PRICE_C)" & _
            " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE inner join file4_10 on file4_10.code = file1_10.code  LEFT JOIN FILE0_90H ON FILE0_90H.DOC_NO = QFILE6_20.DOC_DISC WHERE (QFILE6_20.rate <> 0 or QFILE6_20.price <> QFILE6_20.price_c)   and QFILE6_20.DOC_DISC is not null AND file0_90h.DOC_NO <> '' "
    If xcust.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.code = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M') "
    End If
    If xmosm22.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    
    If xDOC_DISC.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.DOC_DISC = " & MyParn(xDOC_DISC.BoundText)
    
    If XOutlet.Value <> 0 Then cStr2 = cStr2 & " AND file1_10.Outlet = 1 "
    
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cStr2 = cStr2 & " AND FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cStr2 = cStr2 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    
    If xdoc_no.text <> "" Then cStr2 = cStr2 & " AND QFILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cStr2 = cStr2 & " AND STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If Check2.Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.[ISDISC2_1] = 1 "
    
    If xall(0).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 0 "
    
    If xonest(0).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 1 "
    
    cString = cString & cStr2 & " GROUP BY file1_10.ModelNo, FACT.desca , file1_10.mosm, file1_10.MODELFACT0, file1_10.desca  ORDER BY t_sales DESC "
    Set DATA21.Recordset = myRecordSet(cString, con)
End With
FixGrid1
End Sub
Sub FixGrid1()
With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
    .Cols = 15
    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„ÊœÌ·"
    .TextMatrix(0, 3) = "„Ê”„"
    .TextMatrix(0, 4) = "„ÊœÌ· "
    .TextMatrix(0, 5) = "«·’‰›"
    
    .TextMatrix(0, 6) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 7) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 8) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 9) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 10) = "‰”»… «·—»Õ"
    .TextMatrix(0, 11) = "Œ’„ «’‰«›"
    .TextMatrix(0, 12) = "Œ’„ »Ê‰"
    
    .TextMatrix(0, 13) = "ﬁÌ„… »”⁄— „” Â·ﬂ"
    .TextMatrix(0, 14) = "‰”»… „” Â·ﬂ"
    
    .ColWidth(0) = 2500
    
    .ColWidth(0) = 500
    .ColWidth(2) = 0
    .ColWidth(3) = 800
    .ColWidth(4) = 1500
    .ColWidth(5) = 2800
    
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 800
    .ColWidth(11) = 800
    .ColWidth(12) = 800
    .ColWidth(13) = 1000
    .ColWidth(14) = 800
    
    .ColHidden(8) = Not bOpt10
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
   
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 0) = nRow
        nRate = 0
        .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)), 2)
        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .TextMatrix(nRow, 10) = nRate
        If Val(.TextMatrix(nRow, 13)) <> 0 Then
            NP = Val(.TextMatrix(nRow, 13)) - Val(.TextMatrix(nRow, 8))
            nRate = Round((NP / Val(.TextMatrix(nRow, 13))) * 100, 2)
            .TextMatrix(nRow, 14) = nRate
        End If
    Next nRow
    For nCol = 6 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
        Next nRow
    Next nCol
    .SubtotalPosition = flexSTBelow
    For i = 6 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    If .Rows > 1 Then
        nRow = .Rows - 1
        nRate = 0
        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .TextMatrix(nRow, 10) = nRate
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
End Sub
Private Sub grid1_DblClick()
    If Not xall(1).Value Then
        ModelMove.XMODEL.text = grid1.TextMatrix(grid1.Row, 2)
        ModelMove.XSTORE.BoundText = XSTORE.BoundText
        ModelMove.xDate = xDate1.text
        ModelMove.Show
    Else
        ModelMove_FR.XMODEL.text = grid1.TextMatrix(grid1.Row, 2)
        ModelMove_FR.XSTORE.BoundText = XSTORE.BoundText
        ModelMove.xDate = xDate1.text
        ModelMove_FR.Show
    End If
End Sub
Private Sub grid1_EnterCell()
    xModelNo.Caption = grid1.TextMatrix(grid1.Row, 0)
    LoadPhoto grid1.TextMatrix(grid1.Row, 0)
End Sub

Private Sub XDOC_DISC_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        DocDisc_Look
    End If
End Sub
Private Sub xSection_LostFocus()
    data3.ConnectionString = strCon
    If xSection.BoundText = "" Then
        data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Else
        data3.RecordSource = "Select Code,DescA From File1_50 where [group] = " & Val(xSection.BoundText) & " ORDER BY DESCA"
    End If
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    data3.Refresh
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
If ActiveControl.Name = xDOC_DISC.Name Then
    xDOC_DISC.BoundText = Search31.grid1.TextMatrix(Search31.grid1.Row, 0)
    Unload Search31
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
Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
ImgX1.Images.Clear
ImgX1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub

Private Sub ImgX1_Click()
If validPhoto(retFile(grid1.TextMatrix(grid1.Row, 0))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(grid1.TextMatrix(grid1.Row, 0))
    oMaxPhoto.cDesca = grid1.TextMatrix(grid1.Row, 4)
    oMaxPhoto.Show 1
End If
End Sub
Sub DocDisc_Look()
    Dim Generalarray(5)
    Dim listarray(0, 4)
    Dim GrdArray(6, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),FILE0_90H.DESCA , isstop , date1 , date2 " & _
                      " FROM FILE0_90H "
    If cFilter <> "" Then
        Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
    End If
    
    Generalarray(2) = "Order by Date , DOC_NO "
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    
    listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·»Ì«‰"
    listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                      " ##[DATE]## or %%FILE0_90h.DESCA%%)"
    
    GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«· «—ÌŒ"
    GrdArray(1, 1) = 0
    
    GrdArray(2, 0) = "«· «—ÌŒ"
    GrdArray(2, 1) = 1500
    
    GrdArray(3, 0) = "«·»Ì«‰"
    GrdArray(3, 1) = 4000
    
    GrdArray(4, 0) = "„ Êﬁ›"
    GrdArray(4, 1) = 800
    
    GrdArray(5, 0) = "„‰  «—ÌŒ"
    GrdArray(5, 1) = 1500
    
    GrdArray(6, 0) = "≈·Ï"
    GrdArray(6, 1) = 1500
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search31
    Search31.Caption = "«” ⁄·«„"
    Search31.Show 1
End Sub
Private Sub MYLOAD2()
Dim i As Double
Dim cString  As String, cStr2 As String
'                           0               1                 2                3               4                5
    
    cString = " SELECT file0_90h.desca , FACT.desca , file1_10.ModelNo , file1_10.mosm , file1_10.MODELFACT0, file1_10.desca ,   " & _
            " SUM(QFILE6_20.QUANT) AS t_q, SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS t_sales, SUM(QFILE6_20.QUANT * QFILE6_20.cost) AS t_cost, SUM((QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100))  - (QFILE6_20.QUANT * QFILE6_20.cost)) AS t_proft , ' ' , SUM((QFILE6_20.QUANT * QFILE6_20.price_c )-QFILE6_20.total ) , SUM(QFILE6_20.TOTAL * ((QFILE6_20.RATE )/100)) , SUM(QFILE6_20.QUANT * QFILE6_20.PRICE_C)" & _
            " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE inner join file4_10 on file4_10.code = file1_10.code  LEFT JOIN FILE0_90H ON FILE0_90H.DOC_NO = QFILE6_20.DOC_DISC WHERE (QFILE6_20.rate <> 0 or QFILE6_20.price <> QFILE6_20.price_c)   and QFILE6_20.DOC_DISC is not null AND file0_90h.DOC_NO <> '' "
    If xcust.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.code = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M') "
    End If
    If xmosm22.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    
    If xDOC_DISC.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.DOC_DISC = " & MyParn(xDOC_DISC.BoundText)
    
    If XOutlet.Value <> 0 Then cStr2 = cStr2 & " AND file1_10.Outlet = 1 "
    
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cStr2 = cStr2 & " AND FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cStr2 = cStr2 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    
    If xdoc_no.text <> "" Then cStr2 = cStr2 & " AND QFILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cStr2 = cStr2 & " AND STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If Check2.Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.[ISDISC2_1] = 1 "
    
    If xall(0).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 0 "
    
    If xonest(0).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 1 "
    
    cString = cString & cStr2 & " GROUP BY file0_90h.desca ,file1_10.ModelNo, FACT.desca , file1_10.mosm, file1_10.MODELFACT0, file1_10.desca  ORDER BY file0_90h.desca ,t_sales DESC "
    Set DATA22.Recordset = myRecordSet(cString, con)

    FixGrid2
End Sub


Sub FixGrid2()
With GRID2
    .RowHeight(0) = 1000
    .WordWrap = True
    .Cols = 15
    .TextMatrix(0, 0) = "⁄—÷ «·Œ’„"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„ÊœÌ·"
    .TextMatrix(0, 3) = "„Ê”„"
    .TextMatrix(0, 4) = "„ÊœÌ· "
    .TextMatrix(0, 5) = "«·’‰›"
    
    .TextMatrix(0, 6) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 7) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 8) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 9) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 10) = "‰”»… «·—»Õ"
    .TextMatrix(0, 11) = "Œ’„ «’‰«›"
    .TextMatrix(0, 12) = "Œ’„ »Ê‰"
    
    .TextMatrix(0, 13) = "ﬁÌ„… »”⁄— „” Â·ﬂ"
    .TextMatrix(0, 14) = "‰”»… „” Â·ﬂ"
    
    .ColWidth(0) = 2500
    
    .ColWidth(0) = 2000
    .ColWidth(2) = 0
    .ColWidth(3) = 800
    .ColWidth(4) = 1500
    .ColWidth(5) = 2800
    
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 800
    .ColWidth(11) = 800
    .ColWidth(12) = 800
    .ColWidth(13) = 1000
    .ColWidth(14) = 800
    
    .ColHidden(8) = Not bOpt10
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
   
    For nRow = 1 To .Rows - 1
        nRate = 0
        .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)), 2)
        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .TextMatrix(nRow, 10) = nRate
        If Val(.TextMatrix(nRow, 13)) <> 0 Then
            NP = Val(.TextMatrix(nRow, 13)) - Val(.TextMatrix(nRow, 8))
            nRate = Round((NP / Val(.TextMatrix(nRow, 13))) * 100, 2)
            .TextMatrix(nRow, 14) = nRate
        End If
    Next nRow
    For nCol = 6 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
        Next nRow
    Next nCol
    .SubtotalPosition = flexSTBelow
    For i = 6 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
        .Subtotal flexSTSum, 0, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    If .Rows > 1 Then
        nRow = .Rows - 1
        nRate = 0
        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .TextMatrix(nRow, 10) = nRate
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
    End With
End Sub
Private Sub MYLOAD3()
Dim i As Double
Dim cString  As String, cStr2 As String
'                           0               1                 2                3               4                5
    cString = " SELECT FILE4_10.CODE , FILE4_10.DESCA ,' ' , ' ', ' ' , ' ' " & _
            " , SUM(QFILE6_20.QUANT) AS t_q, SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS t_sales, SUM(QFILE6_20.QUANT * QFILE6_20.cost) AS t_cost, SUM((QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100))  - (QFILE6_20.QUANT * QFILE6_20.cost)) AS t_proft , ' ' , SUM((QFILE6_20.QUANT * QFILE6_20.price_c )-QFILE6_20.total ) , SUM(QFILE6_20.TOTAL * ((QFILE6_20.RATE )/100)) , SUM(QFILE6_20.QUANT * QFILE6_20.PRICE_C)" & _
            " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE inner join file4_10 on file4_10.code = file1_10.code  LEFT JOIN FILE0_90H ON FILE0_90H.DOC_NO = QFILE6_20.DOC_DISC WHERE (QFILE6_20.rate <> 0 or QFILE6_20.price <> QFILE6_20.price_c)   and QFILE6_20.DOC_DISC is not null AND file0_90h.DOC_NO <> '' "
    If xcust.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.code = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M') "
    End If
    If xmosm22.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    
    If xDOC_DISC.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.DOC_DISC = " & MyParn(xDOC_DISC.BoundText)
    
    If XOutlet.Value <> 0 Then cStr2 = cStr2 & " AND file1_10.Outlet = 1 "
    
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cStr2 = cStr2 & " AND FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cStr2 = cStr2 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    
    If xdoc_no.text <> "" Then cStr2 = cStr2 & " AND QFILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cStr2 = cStr2 & " AND STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If Check2.Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.[ISDISC2_1] = 1 "
    
    If xall(0).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 0 "
    
    If xonest(0).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 1 "
    
    cString = cString & cStr2 & " GROUP BY FILE4_10.CODE , FILE4_10.DESCA  ORDER BY FILE4_10.CODE , FILE4_10.DESCA "
    Set DATA23.Recordset = myRecordSet(cString, con)
    FixGrid3
End Sub
Sub FixGrid3()
With GRID3
    .RowHeight(0) = 1000
    .WordWrap = True
    .Cols = 15
    .TextMatrix(0, 0) = "ﬂÊœ "
    .TextMatrix(0, 1) = "«·„Ê—œ"
    
    .TextMatrix(0, 6) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 7) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 8) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 9) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 10) = "‰”»… «·—»Õ"
    .TextMatrix(0, 11) = "Œ’„ «’‰«›"
    .TextMatrix(0, 12) = "Œ’„ »Ê‰"
    
    .TextMatrix(0, 13) = "ﬁÌ„… »”⁄— „” Â·ﬂ"
    .TextMatrix(0, 14) = "‰”»… „” Â·ﬂ"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 2000
    
    .ColHidden(2) = True
    .ColHidden(3) = True
    .ColHidden(4) = True
    .ColHidden(5) = True
    
    
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    .ColWidth(12) = 1200
    .ColWidth(13) = 1200
    .ColWidth(14) = 1200
    
    .ColHidden(8) = Not bOpt10
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
   
    For nRow = 1 To .Rows - 1
        nRate = 0
        .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)), 2)
        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .TextMatrix(nRow, 10) = nRate
        If Val(.TextMatrix(nRow, 13)) <> 0 Then
            NP = Val(.TextMatrix(nRow, 13)) - Val(.TextMatrix(nRow, 8))
            nRate = Round((NP / Val(.TextMatrix(nRow, 13))) * 100, 2)
            .TextMatrix(nRow, 14) = nRate
        End If
    Next nRow
    For nCol = 6 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
        Next nRow
    Next nCol
    .SubtotalPosition = flexSTBelow
    For i = 6 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    If .Rows > 1 Then
        nRow = .Rows - 1
        nRate = 0
        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .TextMatrix(nRow, 10) = nRate
    
        If Val(.TextMatrix(nRow, 13)) <> 0 Then
            NP = Val(.TextMatrix(nRow, 13)) - Val(.TextMatrix(nRow, 8))
            nRate = Round((NP / Val(.TextMatrix(nRow, 13))) * 100, 2)
            .TextMatrix(nRow, 14) = nRate
        End If
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    
    End If
    End With
End Sub
Private Sub MYLOAD4()
Dim i As Double
Dim cString  As String, cStr2 As String
'                           0               1                 2                3               4                5
    
    cString = " SELECT STORE_BR.CODE , STORE_BR.DESCA ,' ' , ' ', ' ' , ' ' " & _
            " , SUM(QFILE6_20.QUANT) AS t_q, SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS t_sales, SUM(QFILE6_20.QUANT * QFILE6_20.cost) AS t_cost, SUM((QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100))  - (QFILE6_20.QUANT * QFILE6_20.cost)) AS t_proft , ' ' , SUM((QFILE6_20.QUANT * QFILE6_20.price_c )-QFILE6_20.total ) , SUM(QFILE6_20.TOTAL * ((QFILE6_20.RATE )/100)) , SUM(QFILE6_20.QUANT * QFILE6_20.PRICE_C)" & _
            " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE inner join file4_10 on file4_10.code = file1_10.code INNER JOIN STORE_BR ON STORE_BR.CODE = QFILE6_20.STORE  LEFT JOIN FILE0_90H ON FILE0_90H.DOC_NO = QFILE6_20.DOC_DISC WHERE (QFILE6_20.rate <> 0 or QFILE6_20.price <> QFILE6_20.price_c)   and QFILE6_20.DOC_DISC is not null AND file0_90h.DOC_NO <> '' "
    If xcust.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.code = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M') "
    End If
    If xmosm22.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    
    If xDOC_DISC.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.DOC_DISC = " & MyParn(xDOC_DISC.BoundText)
    
    If XOutlet.Value <> 0 Then cStr2 = cStr2 & " AND file1_10.Outlet = 1 "
    
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cStr2 = cStr2 & " AND FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cStr2 = cStr2 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    
    If xdoc_no.text <> "" Then cStr2 = cStr2 & " AND QFILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cStr2 = cStr2 & " AND STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If Check2.Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.[ISDISC2_1] = 1 "
    
    If xall(0).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 0 "
    
    If xonest(0).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 1 "
    
    cString = cString & cStr2 & " GROUP BY STORE_BR.CODE , STORE_BR.DESCA   ORDER BY STORE_BR.CODE , STORE_BR.DESCA  "
    Set DATA24.Recordset = myRecordSet(cString, con)
FIXGRID4
End Sub
Sub FIXGRID4()
With GRID4
    .RowHeight(0) = 1000
    .WordWrap = True
    .Cols = 15
    .TextMatrix(0, 0) = "ﬂÊœ "
    .TextMatrix(0, 1) = "«·›—⁄"
    
    .TextMatrix(0, 6) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 7) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 8) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 9) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 10) = "‰”»… «·—»Õ"
    .TextMatrix(0, 11) = "Œ’„ «’‰«›"
    .TextMatrix(0, 12) = "Œ’„ »Ê‰"
    
    .TextMatrix(0, 13) = "ﬁÌ„… »”⁄— „” Â·ﬂ"
    .TextMatrix(0, 14) = "‰”»… „” Â·ﬂ"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 2000
    
    .ColHidden(2) = True
    .ColHidden(3) = True
    .ColHidden(4) = True
    .ColHidden(5) = True
    
    
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    .ColWidth(12) = 1200
    .ColWidth(13) = 1200
    .ColWidth(14) = 1200
    
    .ColHidden(8) = Not bOpt10
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
   
    For nRow = 1 To .Rows - 1
        nRate = 0
        .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)), 2)
        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .TextMatrix(nRow, 10) = nRate
        If Val(.TextMatrix(nRow, 13)) <> 0 Then
            NP = Val(.TextMatrix(nRow, 13)) - Val(.TextMatrix(nRow, 8))
            nRate = Round((NP / Val(.TextMatrix(nRow, 13))) * 100, 2)
            .TextMatrix(nRow, 14) = nRate
        End If
    Next nRow
    For nCol = 6 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
        Next nRow
    Next nCol
    .SubtotalPosition = flexSTBelow
    For i = 6 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    If .Rows > 1 Then
        nRow = .Rows - 1
        nRate = 0
        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .TextMatrix(nRow, 10) = nRate
    
        If Val(.TextMatrix(nRow, 13)) <> 0 Then
            NP = Val(.TextMatrix(nRow, 13)) - Val(.TextMatrix(nRow, 8))
            nRate = Round((NP / Val(.TextMatrix(nRow, 13))) * 100, 2)
            .TextMatrix(nRow, 14) = nRate
        End If
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    
    End If
    End With
End Sub



Private Sub myload5()
Dim i As Double
Dim cString  As String, cStr2 As String
'                           0               1                 2                3               4                5
    
    cString = " SELECT FILE0_90H.DOC_NO , FILE0_90H.DESCA ,' ' , ' ', ' ' , ' ' " & _
            " , SUM(QFILE6_20.QUANT) AS t_q, SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS t_sales, SUM(QFILE6_20.QUANT * QFILE6_20.cost) AS t_cost, SUM((QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100))  - (QFILE6_20.QUANT * QFILE6_20.cost)) AS t_proft , ' ' , SUM((QFILE6_20.QUANT * QFILE6_20.price_c )-QFILE6_20.total ) , SUM(QFILE6_20.TOTAL * ((QFILE6_20.RATE )/100)) , SUM(QFILE6_20.QUANT * QFILE6_20.PRICE_C)" & _
            " FROM FILE1_10 INNER JOIN QFILE6_20 ON FILE1_10.ITEM = QFILE6_20.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE inner join file4_10 on file4_10.code = file1_10.code  LEFT JOIN FILE0_90H ON FILE0_90H.DOC_NO = QFILE6_20.DOC_DISC WHERE (QFILE6_20.rate <> 0 or QFILE6_20.price <> QFILE6_20.price_c)   and QFILE6_20.DOC_DISC is not null AND file0_90h.DOC_NO <> '' "
    If xcust.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.code = " & MyParn(xcust.BoundText)
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If XMOSM2.text <> "W" Then
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND SUBSTRING(FILE1_10.MOSM ,1,1) = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND ( SUBSTRING(FILE1_10.MOSM ,1,1) = 'W' OR SUBSTRING(FILE1_10.MOSM ,1,1) = 'M') "
    End If
    If xmosm22.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(xmosm22.text)
    If xsex.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xage.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    
    If xDOC_DISC.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.DOC_DISC = " & MyParn(xDOC_DISC.BoundText)
    
    If XOutlet.Value <> 0 Then cStr2 = cStr2 & " AND file1_10.Outlet = 1 "
    
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGrCust.BoundText <> "" Then cStr2 = cStr2 & " AND FILE4_10.[group] = " & MyParn(xGrCust.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If Not bOpt5 Then cStr2 = cStr2 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    
    If xdoc_no.text <> "" Then cStr2 = cStr2 & " AND QFILE6_20.doc_no = " & MyParn(xdoc_no.text)
    If XSTORE.BoundText <> "" Then cStr2 = cStr2 & " AND STORE = " & MyParn(XSTORE.BoundText)
    If xSection.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & MyParn(xSection.BoundText)
    If IsDate(xDate1.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] <= " & DateSq(xdate2.text)
    If Check2.Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.[ISDISC2_1] = 1 "
    
    If xall(0).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 0 "
    
    If xonest(0).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 1 "
    
    cString = cString & cStr2 & " GROUP BY FILE0_90H.DOC_NO , FILE0_90H.DESCA  ORDER BY FILE0_90H.DOC_NO , FILE0_90H.DESCA  "
    Set data25.Recordset = myRecordSet(cString, con)
    FIXGRID5
End Sub
Sub FIXGRID5()
With grid5
    .RowHeight(0) = 1000
    .WordWrap = True
    .Cols = 15
    .TextMatrix(0, 0) = "„” ‰œ"
    .TextMatrix(0, 1) = "⁄—Ê÷ «·Œ’„"
    
    .TextMatrix(0, 6) = "⁄œœ „»Ì⁄« "
    .TextMatrix(0, 7) = "ﬁÌ„… „»Ì⁄« "
    
    .TextMatrix(0, 8) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(0, 9) = "—»Õ „»Ì⁄« "
    .TextMatrix(0, 10) = "‰”»… «·—»Õ"
    .TextMatrix(0, 11) = "Œ’„ «’‰«›"
    .TextMatrix(0, 12) = "Œ’„ »Ê‰"
    
    .TextMatrix(0, 13) = "ﬁÌ„… »”⁄— „” Â·ﬂ"
    .TextMatrix(0, 14) = "‰”»… „” Â·ﬂ"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 3000
    
    .ColHidden(2) = True
    .ColHidden(3) = True
    .ColHidden(4) = True
    .ColHidden(5) = True
    
    
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    .ColWidth(12) = 1200
    .ColWidth(13) = 1200
    .ColWidth(14) = 1200
    
    .ColHidden(8) = Not bOpt10
    .ColHidden(9) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
   
    For nRow = 1 To .Rows - 1
        nRate = 0
        .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)), 2)
        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .TextMatrix(nRow, 10) = nRate
        If Val(.TextMatrix(nRow, 13)) <> 0 Then
            NP = Val(.TextMatrix(nRow, 13)) - Val(.TextMatrix(nRow, 8))
            nRate = Round((NP / Val(.TextMatrix(nRow, 13))) * 100, 2)
            .TextMatrix(nRow, 14) = nRate
        End If
    Next nRow
    For nCol = 6 To .Cols - 1
        .ColFormat(nCol) = "#0.00"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(.TextMatrix(nRow, nCol), "#0.00")
        Next nRow
    Next nCol
    .SubtotalPosition = flexSTBelow
    For i = 6 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    If .Rows > 1 Then
        nRow = .Rows - 1
        nRate = 0
        If Val(.TextMatrix(nRow, 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)) * 100, 2)
        .TextMatrix(nRow, 10) = nRate
    
        If Val(.TextMatrix(nRow, 13)) <> 0 Then
            NP = Val(.TextMatrix(nRow, 13)) - Val(.TextMatrix(nRow, 8))
            nRate = Round((NP / Val(.TextMatrix(nRow, 13))) * 100, 2)
            .TextMatrix(nRow, 14) = nRate
        End If
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    
    End If
    End With
End Sub




