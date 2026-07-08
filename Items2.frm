VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form items 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„ÊœÌ·« "
   ClientHeight    =   12600
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   15000
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
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
   ScaleHeight     =   12600
   ScaleWidth      =   15000
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdModelFix 
      Caption         =   "÷»ÿ »Ì«‰«  «·„ÊœÌ·« "
      Height          =   510
      Left            =   2655
      TabIndex        =   98
      Top             =   10080
      Width           =   2040
   End
   Begin VB.CommandButton CMD_COPYPICT 
      Caption         =   "‰ﬁ· «·’Ê— ··”Ì—›— «·ÃœÌœ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11070
      TabIndex        =   97
      Top             =   11475
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.CommandButton FIXBAR13 
      Caption         =   "FIX BARCODE13"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1980
      TabIndex        =   95
      Top             =   5445
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton FIXBARCODE2 
      Caption         =   "FIX BARCODE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   94
      Top             =   5445
      Width           =   1905
   End
   Begin VB.Frame Frame12 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   14940
      TabIndex        =   89
      Top             =   45
      Width           =   1500
      Begin VB.CommandButton CMD_LOOK2 
         Caption         =   "≈” ⁄·«„ «’‰«› «’Ê·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   90
         TabIndex        =   91
         Top             =   720
         Width           =   1320
      End
      Begin VB.CommandButton CMD_LOOK1 
         Caption         =   "≈” ⁄·«„ «’‰«› Œœ„…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   90
         TabIndex        =   90
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmd_add_all 
      Caption         =   "«÷«›… „ÊœÌ·«  «·„Ê”„"
      Height          =   510
      Left            =   135
      TabIndex        =   84
      Top             =   10080
      Width           =   2490
   End
   Begin VB.Frame Frame9 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   -1485
      TabIndex        =   26
      Top             =   5940
      Visible         =   0   'False
      Width           =   4245
      Begin VB.CommandButton xFIX 
         Caption         =   "÷»ÿ  ”⁄Ì— √Êﬂ«“ÌÊ‰"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   135
         Width           =   1995
      End
      Begin VB.CommandButton Cmd_Okaz 
         Caption         =   "÷»ÿ ”⁄— «·√Êﬂ«“ÌÊ‰"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2070
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   135
         Width           =   1860
      End
      Begin VB.CommandButton xCmd_Fix 
         Caption         =   "Õ›Ÿ «·”⁄— «·„ÊÕœ"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   180
         Width           =   2040
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "‰”»… «·—»Õ"
      Height          =   780
      Left            =   6480
      TabIndex        =   65
      Top             =   5220
      Width           =   1905
      Begin VB.Label xProfit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         TabIndex        =   66
         Top             =   315
         Width           =   1590
      End
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   4365
      TabIndex        =   60
      Top             =   45
      Width           =   1410
      Begin VB.CommandButton cmdCopy 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items2.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "«·„ﬁ«”« "
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   45
      TabIndex        =   54
      Top             =   6075
      Width           =   14865
      Begin VSFlex7Ctl.VSFlexGrid vsScal 
         Height          =   1920
         Left            =   135
         TabIndex        =   14
         Top             =   315
         Width           =   14625
         _cx             =   25797
         _cy             =   3387
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   9
         Cols            =   10
         FixedRows       =   0
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
         Editable        =   2
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
   Begin VB.Frame FramProg 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      TabIndex        =   57
      Top             =   5265
      Width           =   6405
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   420
         Left            =   90
         TabIndex        =   58
         Top             =   225
         Width           =   6225
         _ExtentX        =   10980
         _ExtentY        =   741
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "«·√·Ê«‰"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   45
      TabIndex        =   55
      Top             =   8415
      Width           =   14865
      Begin VSFlex7Ctl.VSFlexGrid vsColor 
         Height          =   1290
         Left            =   135
         TabIndex        =   59
         Top             =   225
         Width           =   14640
         _cx             =   25823
         _cy             =   2275
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
         Editable        =   2
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
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   8415
      TabIndex        =   52
      Top             =   5220
      Width           =   3480
      Begin VB.CommandButton cmdFixModel 
         Caption         =   "÷»ÿ «·„ÊœÌ·« "
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1755
         TabIndex        =   56
         Top             =   180
         Width           =   1635
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   510
         Left            =   90
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   180
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         PictureFrames   =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Items2.frx":3238
         Caption         =   "—’Ìœ «·„ÊœÌ·"
         Alignment       =   4
         ButtonStyle     =   1
         PictureAlignment=   1
         BevelWidth      =   10
         ShapeSize       =   1
      End
   End
   Begin VB.Frame FramInform 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   10755
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   0
      Width           =   4155
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   90
         Picture         =   "Items2.frx":575A
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   270
         Width           =   1140
      End
      Begin VB.TextBox xItem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   465
         Left            =   1305
         MaxLength       =   40
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»«—ﬂÊœ :"
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
         Left            =   3330
         TabIndex        =   78
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5805
      TabIndex        =   43
      Top             =   45
      Width           =   4920
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   3645
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items2.frx":7F2D
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items2.frx":A290
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items2.frx":C809
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1275
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items2.frx":EC75
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   4920
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items2.frx":1150F
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   225
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   11925
      TabIndex        =   37
      Top             =   5220
      Width           =   2985
      Begin Threed.SSCommand calc 
         Height          =   540
         Index           =   1
         Left            =   795
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   180
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   953
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1/10"
      End
      Begin Threed.SSCommand calc 
         Height          =   540
         Index           =   0
         Left            =   90
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   180
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   953
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1/12"
      End
      Begin Threed.SSCommand calc 
         Height          =   540
         Index           =   2
         Left            =   1500
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   180
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   953
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1◊ 12"
      End
      Begin Threed.SSCommand calc 
         Height          =   540
         Index           =   3
         Left            =   2205
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   180
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   953
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Impact"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1◊10"
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
      Height          =   960
      Left            =   4365
      TabIndex        =   16
      Top             =   900
      Width           =   10545
      Begin VB.ComboBox XMOSM2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   135
         TabIndex        =   75
         Top             =   135
         Width           =   1230
      End
      Begin VB.TextBox xModelNo 
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
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   135
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   540
         Width           =   4155
      End
      Begin VB.TextBox xModel 
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
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   5625
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   540
         Width           =   4065
      End
      Begin VB.TextBox xModelFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   2745
         MaxLength       =   10
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   7605
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
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
      Begin MSDataListLib.DataCombo XMOSM 
         Height          =   315
         Left            =   5625
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
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
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ÊœÌ· «·„’‰⁄ :"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ÊœÌ· :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   585
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„ÊœÌ· :"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " ’‰Ì› «·œ«Œ·Ï"
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
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„ :"
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
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’‰⁄ :"
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
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   180
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   4365
      TabIndex        =   21
      Top             =   1890
      Width           =   10545
      Begin VB.CheckBox XISONLINE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„»Ì⁄«  «Ê‰ ·«Ì‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   1395
         Width           =   1725
      End
      Begin VB.TextBox xdesce 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   135
         MaxLength       =   40
         TabIndex        =   88
         Top             =   540
         Width           =   4245
      End
      Begin VB.CheckBox XISNOITEM2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "’‰› «’Ê·"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6750
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   2295
         Width           =   1230
      End
      Begin VB.CheckBox XOutlet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Outlet "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   2295
         Width           =   1050
      End
      Begin VB.CheckBox XISONEST 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«„«‰«  "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   2295
         Width           =   870
      End
      Begin VB.CommandButton FIXGRMODEL 
         Caption         =   "FIX GRMODEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         TabIndex        =   83
         Top             =   2700
         Width           =   1185
      End
      Begin VB.TextBox xGrModel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   4725
         MaxLength       =   11
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   2745
         Width           =   1545
      End
      Begin VB.TextBox xGrDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1350
         MaxLength       =   11
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   2745
         Width           =   2130
      End
      Begin VB.TextBox xSupp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   6120
         MaxLength       =   10
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   1800
         Width           =   510
      End
      Begin VB.CheckBox XISNOITEM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "’‰› Œœ„… "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   2295
         Width           =   1230
      End
      Begin VB.CheckBox XISNODEM 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "’‰› »œÊ‰ ÿ·»Ì…"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   225
         Width           =   1635
      End
      Begin VB.CheckBox ISNEWPRICE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " „  ⁄œÌ· «·”⁄—"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   2295
         Width           =   1410
      End
      Begin VB.CheckBox xIsOkazItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·»Ì⁄ »”⁄— «·«Êﬂ«“ÊÌ‰"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   2295
         Width           =   1950
      End
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5580
         TabIndex        =   64
         Top             =   180
         Width           =   330
      End
      Begin VB.TextBox xOkaz_2 
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
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   135
         MaxLength       =   10
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CheckBox xisstop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         Caption         =   "’‰› „ Êﬁ›"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         TabIndex        =   42
         Top             =   2205
         Width           =   1815
      End
      Begin VB.TextBox xOkaz 
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
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   135
         MaxLength       =   10
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   900
         Width           =   1815
      End
      Begin VB.TextBox xDescA 
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
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   4455
         MaxLength       =   40
         TabIndex        =   7
         Top             =   540
         Width           =   4290
      End
      Begin VB.TextBox xCode 
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
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   7965
         Locked          =   -1  'True
         MaxLength       =   40
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   900
         Width           =   780
      End
      Begin VB.TextBox xRate 
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
         ForeColor       =   &H80000012&
         Height          =   330
         Left            =   7155
         MaxLength       =   10
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1590
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   390
         Left            =   5940
         TabIndex        =   6
         Top             =   180
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XCHARGE 
         Height          =   315
         Left            =   7515
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   2745
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xModelAge 
         Height          =   315
         Left            =   6120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1350
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xModelSex 
         Height          =   315
         Left            =   135
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1350
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   315
         Left            =   135
         TabIndex        =   92
         Top             =   180
         Width           =   2670
         _ExtentX        =   4710
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
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„’—Ê›"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9810
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   2745
         Width           =   630
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄… „ÊœÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6300
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   2790
         Width           =   1155
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»Ì«‰ «·„Ã„Ê⁄… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   2790
         Width           =   1200
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·›∆… «·⁄„—Ì… : "
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
         Left            =   8775
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   1350
         Width           =   1110
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·‰Ê⁄ : "
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   1395
         Width           =   585
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… √Êﬂ«“ÌÊ‰  «·⁄—÷ :"
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   1890
         Width           =   1845
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   225
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄… «·„ÊœÌ· :"
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
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… ”⁄— √Êﬂ«“ÌÊ‰ :"
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
         Left            =   2025
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   945
         Width           =   1605
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·„Ê—œ :"
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
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   915
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈”„ «·„ÊœÌ·  :"
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
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   555
         Width           =   1095
      End
      Begin VB.Label xCodeDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   4455
         TabIndex        =   15
         Top             =   900
         Width           =   3480
      End
      Begin VB.Label EQWEW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·„” Â·ﬂ :"
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
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1800
         Width           =   1245
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   7245
      Top             =   5490
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
      Left            =   7650
      Top             =   5175
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
      Left            =   10440
      Top             =   9855
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
      Left            =   7110
      Top             =   4950
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   945
      _ExtentX        =   847
      _ExtentY        =   847
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
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5280
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   -45
      Width           =   4290
      Begin VB.CommandButton cmdDelPhoto 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–›"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1485
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3690
         Width           =   825
      End
      Begin VB.CommandButton cmdRotate 
         Caption         =   "„‰ „·›"
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2340
         TabIndex        =   33
         Top             =   3690
         Width           =   915
      End
      Begin VB.CommandButton cmdScan 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3285
         Picture         =   "Items2.frx":13ABB
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   3690
         Width           =   825
      End
      Begin VB.Frame Frame13 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   4140
         Width           =   4110
         Begin VB.CheckBox chkEdit 
            Appearance      =   0  'Flat
            Caption         =   " ⁄œÌ·"
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   135
            TabIndex        =   104
            Top             =   225
            Width           =   690
         End
         Begin Threed.SSCommand cmdFirst 
            Height          =   420
            Left            =   3240
            TabIndex        =   100
            TabStop         =   0   'False
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
            Picture         =   "Items2.frx":161F9
            Caption         =   "√Ê·"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "Items2.frx":183A0
         End
         Begin Threed.SSCommand cmdPrevious 
            Height          =   420
            Left            =   2430
            TabIndex        =   101
            TabStop         =   0   'False
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
            Picture         =   "Items2.frx":1A3E7
            Caption         =   "”«»ﬁ"
            ButtonStyle     =   3
            PictureAlignment=   10
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "Items2.frx":1C4D2
         End
         Begin Threed.SSCommand cmdNext 
            Height          =   420
            Left            =   1665
            TabIndex        =   102
            TabStop         =   0   'False
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
            Picture         =   "Items2.frx":1E4CC
            Caption         =   "·«Õﬁ"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "Items2.frx":205DD
         End
         Begin Threed.SSCommand cmdLast 
            Height          =   420
            Left            =   900
            TabIndex        =   103
            TabStop         =   0   'False
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
            Picture         =   "Items2.frx":225D7
            Caption         =   "√ŒÌ—"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "Items2.frx":247FB
         End
      End
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   3390
         Left            =   90
         TabIndex        =   35
         Top             =   225
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   5980
         BackColor       =   16777215
         BorderStyle     =   4
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         AutoZoomType    =   2
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
      Begin VB.CommandButton BrtPlus 
         Caption         =   "B+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1305
         TabIndex        =   62
         Top             =   2205
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton BrtMinus 
         Caption         =   "B-"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1305
         TabIndex        =   63
         Top             =   1755
         Visible         =   0   'False
         Width           =   645
      End
      Begin Threed.SSCommand cmdColor 
         Height          =   375
         Left            =   1350
         TabIndex        =   107
         Top             =   4815
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   661
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "»œÊ‰ ·Ê‰"
         TagVariant      =   "»œÊ‰ ·Ê‰"
         ButtonStyle     =   3
      End
      Begin VB.Label lblRecords 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   4815
         Width           =   1230
      End
      Begin VB.Label xId 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   3645
         Visible         =   0   'False
         Width           =   690
      End
   End
End
Attribute VB_Name = "items"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents twain As ImgXTwain
Attribute twain.VB_VarHelpID = -1
Dim con As New ADODB.Connection
Dim oSearchColor As New Search_abd
Dim con_image As New ADODB.Connection
Dim bIg As Boolean
Dim confact As New ADODB.Connection
Dim confact2 As New ADODB.Connection
Dim confact3 As New ADODB.Connection
Dim ISNEWMODEL As Boolean
Dim bAdded As Boolean
Public CardTable As ADODB.Recordset
Public bEdit As Boolean
Public nFlag As Integer
Public nSection As Integer
Dim fs As New FileSystemObject
Dim rsPhotos As ADODB.Recordset
Public SFact As String, SMosm As String, sSupp As String, SModelFact As String, sCode As String, sModel As String, sSection  As String, sModelItem As Double
Public bAddModel As Boolean
Dim oSearchItem As New Search3, sControl As String
Dim sOkaz As String, sRate As String, sBig As String, sRedem As String
Dim ColorTable As New ADODB.Recordset
Dim cFilter As String
Dim nColWidth As Long
Dim nRecordNumber As Integer
Const LoadMode = 1, DefineMode = 2
Sub myDefine()
ISNEWMODEL = False
XMODEL.text = ""
myDefinePhotos
XSECTION.BoundText = ""
ISNEWPRICE.Value = 0
xFact.BoundText = ""
xModelFact.text = ""
xMosm.BoundText = ""
xGroup.BoundText = ""
xCharge.BoundText = ""
xmodelsex.BoundText = ""
xmodelage.BoundText = ""

xGrDesca.text = ""
XGRMODEL.text = ""

xSupp.text = ""
xCode.text = ""
xCodeDesc.Caption = ""
xisnodem.Value = 0
XISNOITEM.Value = 0
XISNOITEM2.Value = 0
xIsOkazItem.Value = 0
XOutlet.Value = 0
XISONLINE.Value = 0

'xbig.Text = sBig
XOKAZ.text = sOkaz
xOkaz_2.text = ""

xDescA.text = ""
xdesce.text = ""

vsScal.Rows = 1
vsColor.Rows = 1
xitem.text = ""
xProfit.Caption = ""
ImgX1.Images.Clear
xisstop.Value = 0
xIsOkazItem.Value = 0
myDefineGrd
Handlecontrols DefineMode
End Sub
Sub myProc()
If ActiveControl.Name = cmdInform.Name Or ActiveControl.Name = CMD_LOOK1.Name Or ActiveControl.Name = CMD_LOOK2.Name Then
    XMODEL.text = oSearchItem.Grid1.TextMatrix(oSearchItem.Grid1.Row, 0)
    Unload oSearchItem
    openCardTable
    myUndo
ElseIf ActiveControl.Name = xCharge.Name Then
    xCharge.BoundText = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)
    Unload Search3
ElseIf ActiveControl.Name = cmdColor.Name Then
    If oSearchColor.Grid1.TextMatrix(oSearchColor.Grid1.Row, 0) = "" Then
        cmdColor.Tag = ""
        cmdColor.Caption = cmdColor.TagVariant
    Else
        cmdColor.Tag = oSearchColor.Grid1.TextMatrix(oSearchColor.Grid1.Row, 0)
        cmdColor.Caption = oSearchColor.Grid1.TextMatrix(oSearchColor.Grid1.Row, 1)
    End If
    
    con_image.Execute "UPDATE MODEL_IMAGES " & _
                      " SET C_COLOR = " & addvalue(cmdColor.Tag) & _
                      " WHERE ID = " & xId.Caption
    
    If Not openTablePhotos(tbMode.tbFind, xId.Caption) Then
        If Not openTablePhotos Then myDefine
    End If
    oSearchColor.Hide
End If
End Sub
Public Sub myload()
On Error GoTo myError
XMODEL.text = CardTable!MODEL
xModelFact.text = CardTable!modelfact & ""
xModelNo.text = CardTable!MODELNO & ""
xSupp.text = CardTable!SUPP & ""
xMosm.BoundText = CardTable!MOSM & ""
XMOSM2.text = CardTable!MOSM2 & ""
XSECTION.BoundText = CardTable!Section & ""
xCharge.BoundText = CardTable!CHARGE & ""
data2.ConnectionString = strCon

xGroup.BoundText = CardTable!Group & ""

ISNEWPRICE.Value = 0
xFact.BoundText = CardTable!Fact & ""
XOKAZ.text = Myvalue(CardTable!okaz & "")
xOkaz_2.text = Myvalue(CardTable!okaz_2 & "")
xRate.text = Myvalue(CardTable!Rate, "#0.00")

xmodelsex.BoundText = CardTable!MODELSEX & ""
xmodelage.BoundText = CardTable!MODELAGE & ""


xCode.text = CardTable!code & ""
xCodeDesc.Caption = CardTable!SUPPNAME & ""
xDescA.text = CardTable!desca & ""
xdesce.text = CardTable!DESCe & ""
xisstop.Value = IIf(CardTable!ISSTOP & "", 1, 0)

If cBranch = "00" Then
    xGrDesca.text = CardTable!GrDesca & ""
    XGRMODEL.text = CardTable!GRMODEL & ""
    XISONEST.Value = IIf(CardTable!ISONEST & "", 1, 0)
    XOutlet.Value = IIf(CardTable!Outlet & "", 1, 0)
    XISNOITEM2.Value = IIf(CardTable!ISNOITEM2, 1, 0)
    XISONLINE.Value = IIf(CardTable!IsOnLine, 1, 0)
End If

xisnodem.Value = IIf(CardTable!ISNODEM, 1, 0)
XISNOITEM.Value = IIf(CardTable!ISNOITEM, 1, 0)
xIsOkazItem.Value = IIf(CardTable!ISOKAZITEM, 1, 0)

'myLoadPicture xModelNo.text

myLoadGrd

myloadgrd2
'xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub BrtMinus_Click()
On Error GoTo myError
If validPhoto(retFile(XMODEL.text)) Then
    ImgX1.Filters.Brightness (-5)
    ImgX1.Export.ToFile retFile(XMODEL.text), ixfsJPG
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub BrtPlus_Click()
On Error GoTo myError
If validPhoto(retFile(XMODEL.text)) Then
    ImgX1.Filters.Brightness (5)
    ImgX1.Export.ToFile retFile(XMODEL.text), ixfsJPG
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Cmd_Undo_Click()
myUndo
End Sub

Private Sub calc_Click(Index As Integer)
With vsScal
Dim nValue As Double
Select Case Index
Case 0
    nValue = (1 / 12)
Case 1
    nValue = (1 / 10)
Case 2
    nValue = 10
Case 3
    nValue = 12
End Select
For i = 1 To vsScal.Rows - 2
  .TextMatrix(i, .Col) = Myvalue(Round(nValue * Val(.TextMatrix(i, .Col)), 2))
Next
'If Val(.TextMatrix(.Row, .Col)) > 0 Then .TextMatrix(.Row, .Col) = Myvalue(Val(.TextMatrix(.Row, .Col)) / 12, "#0.00")
End With
End Sub

Private Sub cmd_add_all_Click()
    Dim sFactMosm  As String, i As Double
    Dim sShopMosm  As String
    Dim ItemMosmTable As New ADODB.Recordset
    Dim sSecModel  As String
    
    sFactMosm = InputBox("", " „Ê”„ «·„’‰⁄ ", "")
    sShopMosm = InputBox("", " „Ê”„ «·„Õ·«  ", "")
    sSecModel = InputBox("", " ﬁ”„ «·„ÊœÌ·  ", "")
    ItemMosmTable.Open "SELECT MODEL FROM FILE1_10 WHERE MOSM = " & MyParn(sFactMosm) & " AND SEC = " & addstring(sSecModel) & " GROUP BY MODEL ORDER BY MODEL ", confact, adOpenStatic, adLockReadOnly, adCmdText
    prog1.Min = 0
    prog1.Max = ItemMosmTable.RecordCount
    prog1.Value = 0
    With ItemMosmTable
        Do While Not .EOF
            i = i + 1
            prog1.Value = i
            Me.Caption = ItemMosmTable!MODEL
            Me.Caption = " Model : " & ItemMosmTable!MODEL
            myreplace_ItemFct ItemMosmTable!MODEL, sShopMosm, "001"
            .MoveNext
        Loop
        con.Execute " UPDATE FILE1_10 SET ITEM = ITEM WHERE MOSM = " & MyParn(sShopMosm)
    End With
End Sub
Private Sub CMD_COPYPICT_Click()
Dim MyPictTable As New ADODB.Recordset
    MyPictTable.Open "select top 500 * from  PICT", conPict, adOpenStatic, adLockReadOnly, adCmdText
    nRecordCount = MyPictTable.RecordCount
    MsgBox " Record : " & nRecordCount
    prog1.Value = 0
    prog1.Min = 0
    prog1.Max = nRecordCount
    With MyPictTable
        Do While Not .EOF
            i = i + 1
            prog1.Value = i
            Me.Caption = " " & !MODELNO
            xModelNo.text = !MODELNO
            If xModelNo.text <> "" Then
                myLoadPicture (MyPictTable!MODELNO)
                AddFromImageFile2 Common1.FileName
                Inform " „ Õ›Ÿ «·’Ê—… " & i
            End If
            MyPictTable.MoveNext
        Loop
    End With
End Sub

Private Sub CMD_LOOK1_Click()
    ModelLookupAll_Ser1 Me, oSearchItem
End Sub
Private Sub CMD_LOOK2_Click()
    ModelLookupAll_Ser2 Me, oSearchItem
End Sub

Private Sub cmdClient_Click()

End Sub
Private Sub cmdColor_Click()
ColorLookup Me, oSearchColor, "FILE1_10.MODEL = " & MyParn(XMODEL.text), , IIf(cmdColor.Tag = "", "", "»œÊ‰ ·Ê‰")
End Sub

Private Sub cmdcopy_Click()
    Dim cModel As String
    cModel = DelZero(xModelFact.text)
    If xFact.BoundText = "001" And cModel <> "" And Trim(xDescA.text) = "" Then
        FillItemFact (cModel)
    End If
    If xCode.text = "049" And cModel <> "" And Trim(xDescA.text) = "" Then
        FillItemFact2 (cModel)
    End If
    If xCode.text = "049" And cModel <> "" And Trim(xDescA.text) = "" Then
        FillItemFact2 (cModel)
    End If
    If xCode.text = "095" And cModel <> "" And Trim(xDescA.text) = "" Then
        FillItemFact3 (cModel)
    End If
End Sub

Private Sub CmdDel_Click()
On Error GoTo myError
If cBranch <> "00" Then Exit Sub
If GetDesca("SELECT FILE1_10.ITEM FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_11.ITEM = FILE1_10.ITEM WHERE FILE1_10.MODEL = " & MyParn(XMODEL.text), con) <> "" Then
    MsgBox "ÌÊÃœ Õ—ﬂ… ⁄·Ï «·„ÊœÌ·"
    Exit Sub
End If
If MsgBox("”Ì „ «·€«¡ «·„ÊœÌ· »«·ﬂ«„· ", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
'    con.BeginTrans
    cString = " DELETE  FROM FILE1_10 WHERE MODEL = " & MyParn(XMODEL.text)
    con.Execute cString

    cString = " DELETE  FROM FILE1_10h WHERE MODEL = " & MyParn(XMODEL.text)
    con.Execute cString

'    con.CommitTrans
End If
AddLod_Data cusername, 2, " Õ–› «·„ÊœÌ·«  ", con, XMODEL.text, , , xDescA.text

myDefine
openCardTable
MsgBox "  „ Õ–› «·„ÊœÌ· "
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
'con.RollbackTrans
End Sub
Private Sub cmdDelPhoto_Click()
If MsgBox("Õ–› «·’Ê—… ", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
On Error GoTo myError
con_image.Execute "Delete from model_images where id = " & xId.Caption

If Not openTablePhotos(tbMode.tbPrevious, xId.Caption) Then
    If Not openTablePhotos(tbMode.tbFirst) Then
        myDefinePhotos
    End If
End If
myInform "done..."
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub



Private Sub cmdGroup_Click()
Dim sBound As String
sBound = xGroup.BoundText
GrItem.Show 1
data2.Recordset.Requery
xGroup.BoundText = sBound
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""

'ReDim aLocal(6)
'cCode = xGroup.BoundText
'aLocal(0) = "FILE1_50"
'aLocal(1) = "Code"
'aLocal(2) = "Desca"
'aLocal(3) = "ﬂÊœ „Ã„Ê⁄…"
'aLocal(4) = "»Ì«‰ „Ã„Ê⁄« "
'aLocal(5) = " ”ÃÌ· „Ã„Ê⁄« "
'aLocal(6) = 2
'FlagFrm2.bEdit = True
'FlagFrm2.myPublic = aLocal
'FlagFrm2.Show 1
'DATA2.Refresh


'xGroup.BoundText = cCode
'If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub

Private Sub CmdInform_Click()
    ModelLookupAll Me, oSearchItem
End Sub
Private Sub OLDcmdRotate_Click()
'''On Error GoTo myerror
'''If validPhoto(retFile(xmodelno.Text)) Then
'''    ImgX1.Effects.Rotate 90
'''    ImgX1.Export.ToFile retFile(xmodelno.Text), ixfsJPG
'''End If
'''Exit Sub
'''myerror:
'''MsgBox Err.Description
'''Err.Clear
End Sub
Private Sub cmdModelFix_Click()
Dim cmd As New ADODB.command
Set cmd = mycmdEx("dbo.sp_model_fix", con, AddFlag(Empty, "MODEL", Null))
Inform " „  ⁄œÌ· " & cmd.Parameters("@COUNT_INSERT").Value & " „ÊœÌ·"
End Sub

Private Sub cmdRotate_Click()
    Dim cFile As String, cNewFile As String
    Common1.FileName = ""
    'Common1.InitDir = doc_dir()
    Common1.Filter = "Pictures (*.Jpg)|*.Jpg"
    Common1.ShowOpen
    If Common1.FileTitle <> "" Then
        cFile = Common1.FileName
        If cFile <> "" Then
            myReplaceImage IIf(chkEdit.Value = 1, xId.Caption, ""), cFile
        End If
    End If
    
End Sub
Private Sub cmdSave_Click()
If cBranch <> "00" Then Exit Sub
If Not myValid Then Exit Sub
If myreplace Then
    If xFact.BoundText = "001" Then FixBarCode2_Fact XMODEL.text
    If xFact.BoundText = "095" Then FixBarCode2_Fact3 XMODEL.text
    If Not ISNEWMODEL And ISNEWPRICE.Value = 1 Then
        UpdatePrice
        ISNEWPRICE.Value = 0
    End If

    Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
    AddLod_Data cusername, 1, " Õ÷Ÿ «·„ÊœÌ·«  ", con, XMODEL.text, , , xDescA.text

End If
End Sub

Private Sub cmdScan_Click()
'If xFact.text <> "" And XMOSM.BoundText <> "" And xFact.BoundText <> "" Then
'    Dim sModelNo As String
'    sModelNo = xFact.BoundText & XMOSM.BoundText & RetZero(xModelFact.text, 10)
'    If Not XMOSM.MatchedWithList Then MsgBox "«·„Ê”„ €Ì— ’ÕÌÕ"
'    If Len(XMOSM.BoundText) < 3 Then MsgBox "«·„Ê”„ €Ì— ’ÕÌÕ"
'    MyCreateFolder App.Path & "\PICT" & turn(XMOSM.BoundText, "\") & XMOSM.BoundText
'    If ScanImage Then cmdSave_Click
'End If
nOption = 0
ScanImage
End Sub
Private Sub CmdUndo_Click()
If XMODEL.Tag = DefineMode And bAddModel Then
    myAddModel
Else
    openCardTable
    myUndo
End If
End Sub
Private Sub cmdFixModel_click()
Dim cMod As String, cMod0 As String, cModel As String, cModelNo  As String, cString As String, cCaption As String
Dim loctable As New ADODB.Recordset
cString = "Select * from file1_10 "
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
con.BeginTrans
'On Error GoTo myerror
Me.FramProg.Visible = True
Dim aInsert(3, 1)
With loctable
    Dim nRecord As Long, nRecordCount As Long
    nRecordCount = loctable.RecordCount
    cCaption = Me.Caption
    Do Until .EOF
        nRecord = nRecord + 1
        Me.Caption = cCaption & " ”Ã· —ﬁ„ " & nRecord & " „‰ " & nRecordCount
        prog1.Value = Round(nRecord / nRecordCount, 2) * 100

        cMod = UCase(Trim(!modelfact & ""))
        cMod = RetZero(Replace(cMod, "'", "-"), 10)
'        cMod = Mid(cMod, 7, 10)
        cMod0 = DelZero(cMod)
        cModelNo = UCase(!Fact & !MOSM & cMod)
        cModel = UCase(!Fact & !MOSM & cMod & Trim(!SUPP & ""))
        
        aInsert(0, 0) = "ModelFact"
        aInsert(0, 1) = addstring(cMod)
        
        aInsert(1, 0) = "ModelFact0"
        aInsert(1, 1) = addstring(cMod0)
        
        aInsert(2, 0) = "MODEL"
        aInsert(2, 1) = addstring(cModel)
        
        aInsert(3, 0) = "MODELNO"
        aInsert(3, 1) = addstring(cModelNo)
        
        con.Execute CreateUpdate(aInsert, "FILE1_10", " where item = " & !Item, Array(-1))
        
'        con.Execute " DELETE FROM ITEM_BRANCH WHERE ITEM = " & sitem
'        For nRow = 1 To 4
'            sStore = RetZero(nRow, 2)
'            con.Execute "  INSERT INTO ITEM_BRANCH (ITEM, BRANCH) VALUES        ( " & addvalue(sitem) & " , " & addstring(sStore) & " )  "
'        Next nRow
        
        .MoveNext
    Loop
End With
con.CommitTrans
Set loctable = Nothing
Inform " „ ÷»ÿ »Ì«‰ «·„ÊœÌ·«  »‰Ã«Õ"
Me.FramProg.Visible = False
Me.Caption = cCaption
Exit Sub
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
Me.Caption = cCaption
End Sub
Private Sub cmdFixModel2_click()
Dim cString As String, sModel As String, nCode_Scal As Long
Dim loctable As New ADODB.Recordset
loctable.Open "SELECT FILE1_10.* FROM FILE1_10 ORDER BY FILE1_10.MODEL,FILE1_10.C_SCAL", con, adOpenStatic, adLockReadOnly, adCmdText
On Error GoTo myError
Me.FramProg.Visible = True
sModel = "Dummy123"
con.BeginTrans
On Error GoTo myError
With loctable
    Dim nRecord As Long, nRecordCount As Long
    nRecordCount = loctable.RecordCount
    Do Until .EOF
        If sModel <> loctable!MODEL Then
            nCode_Scal = 1
            sModel = loctable!MODEL
        End If
        nRecord = nRecord + 1
        prog1.Value = Round(nRecord / nRecordCount, 2) * 100
        cString = " UPDATE FILE1_10 SET CODE_SCAL = " & nCode_Scal & " FROM FILE1_10 where item = " & !Item
        con.Execute cString
        nCode_Scal = nCode_Scal + 1
        .MoveNext
    Loop
End With
con.CommitTrans
lastsub:
Me.FramProg.Visible = False
Set loctable = Nothing
Inform " „ ÷»ÿ »Ì«‰ «·„ÊœÌ·«  »‰Ã«Õ"
Exit Sub
myError:
Err.Clear
MsgBox Err.Description
GoTo lastsub
End Sub

Private Sub Command1_Click()
Dim cString As String, loctable As New ADODB.Recordset
cString = "SELECT * FROM FILE1_10"
cString = cString & turn(cString) & "FILE1_10.MODEL = " & MyParn(XMODEL.text)
'cString = cString & turn(cString) & "FILE1_10.CODE_SCAL IS NULL"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
'con.BeginTrans
On Error Resume Next
Do Until loctable.EOF
    cString = "DELETE FROM FILE1_10"
    cString = cString & turn(cString) & " FILE1_10.item = " & loctable!Item
    con.Execute cString, nRecord
    loctable.MoveNext
Loop
'con.CommitTrans
loctable.Close
Set loctable = Nothing
Exit Sub
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub Command2_Click()
Dim fs, f, f1, fc, s, photoDir As String, cFile As String, SMosm As String
contemp.Execute "Delete from temp"
photoDir = App.Path & "\pict2"
Set fs = CreateObject("Scripting.FileSystemObject")
Set f = fs.GetFolder(photoDir)
Set fc = f.Files
nCount = fc.Count
For Each f1 In fc
i = i + 1
If InStr(1, LCase(photoDir & "\" & f1.Name), "jpg") <> 0 Then
    cFile = Replace(LCase(f1.Name), ".jpg", "")
    If Len(cFile) = 10 Then
        SMosm = Mid(cFile, 4, 3)
        cFile = retFile(cFile)
        MyCreateFolder App.Path & "\PICT" & turn(SMosm, "\") & SMosm
        On Error Resume Next
        fs.CopyFile photoDir & "\" & f1.Name, cFile
        If Err.Number <> 0 Then
            Err.Clear
            contemp.Execute "INSERT INTO temp(str1) " & _
                         "VALUES(" & _
                         addstring(f1.Name) & _
            ")"
        End If
    End If
End If
Me.Caption = i
Next
End Sub
Private Sub ss()
End Sub
Private Sub Command3_Click()
'con.BeginTrans
con.Execute "UPDATE FILE7_20 SET FILE7_20.ITEM = FILE7_20.ITEM FROM FILE7_20"
con.Execute "UPDATE FILE7_10 SET FILE7_10.ITEM = FILE7_10.ITEM FROM FILE7_10"
con.Execute "UPDATE FILE6_20 SET FILE6_20.ITEM = FILE6_20.ITEM "
con.Execute "UPDATE FILE0_10H SET FILE0_10H.CLOSED = FILE0_10H.CLOSED"
'con.CommitTrans

'ADDFILE "FILE7_20"
End Sub

Private Sub FIXBARCODE_Click()
    FixBarcodeFRM.pModelNo = XMODEL
    FixBarcodeFRM.Show 1
End Sub

Private Sub OLDFIXBARCODE2_Click()
Dim cBarCodE2 As String
Dim ItemFactTable As New ADODB.Recordset

ItemFactTable.Open "SELECT * FROM FILE1_10 WHERE CODE = '001' ORDER BY ITEM ", con, adOpenStatic, adLockReadOnly, adCmdText
With ItemFactTable
    Do While Not .EOF
        cBarCodE2 = GetDesca("SELECT ITEM_BARCODE FROM FILE1_10 WHERE ITEM = " & MyParn(!BARCODE), confact)
        con.Execute " UPDATE FILE1_10 SET BARCODE2 = " & addstring(cBarCodE2) & " WHERE ITEM = " & !Item
        Me.Caption = !Item & "  " & cBarCodE2
        .MoveNext
    Loop
End With


ItemFactTable.Close
ItemFactTable.Open "SELECT * FROM FILE1_10 WHERE CODE = '049' ORDER BY ITEM ", con, adOpenStatic, adLockReadOnly, adCmdText
With ItemFactTable
    Do While Not .EOF
        cBarCodE2 = GetDesca("SELECT ITEM_BARCODE FROM FILE1_10 WHERE ITEM = " & MyParn(!BARCODE), confact2)
        con.Execute " UPDATE FILE1_10 SET BARCODE2 = " & addstring(cBarCodE2) & " WHERE ITEM = " & !Item
        Me.Caption = !Item & "  " & cBarCodE2
        .MoveNext
    Loop
End With
End Sub

Private Sub oldFIXBAR13_Click()
'    Dim i As Double
'    con.Execute " update file1_10 set barcode13 = null where barcode13 ='' "
'    Dim MyItem13 As New ADODB.Recordset
'    MyItem13.Open " SELECT MODEL FROM FILE1_10 WHERE CODE = '001' AND barcode13 IS NULL ORDER BY MODEL ", con, adOpenStatic, adCmdText
'    If MyItem13.RecordCount = 0 Then Exit Sub
'    prog1.Min = 0
'    prog1.Max = MyItem13.RecordCount
'    prog1.Value = 0
'    i = 0
'    MyItem13.MoveFirst
'    Do While Not MyItem13.EOF
'        i = i + 1
'        prog1.Value = i
'        FixBarCode2_Fact (MyItem13!MODEL)
'        MyItem13.MoveNext
'    Loop
End Sub
Private Sub FIXBAR13_Click()
    Dim i As Double
    con.Execute " update file1_10 set barcode13 = null where barcode13 ='' "
    con.Execute " UPDATE FILE1_10 SET BARCODE2 = NULL WHERE BARCODE2 ='' "
    Dim MyItem13 As New ADODB.Recordset
    MyItem13.Open " SELECT * FROM FILE1_10 WHERE CODE = '001' AND MOSM = " & MyParn(cPMosm) & " ORDER BY MODEL ", con, adOpenStatic, adCmdText
    
    If MyItem13.RecordCount = 0 Then Exit Sub
    prog1.Min = 0
    prog1.Max = MyItem13.RecordCount
    prog1.Value = 0
    i = 0
    MyItem13.MoveFirst
    Do While Not MyItem13.EOF
        i = i + 1
        prog1.Value = i
        If Not IsNull(MyItem13!MODEL) Then
            If MyItem13!code = "001" Then FixBarCode2_Fact (MyItem13!MODEL)
        End If
        MyItem13.MoveNext
    Loop
End Sub
Private Sub FIXBAR13_S_Click()
    Dim i As Double
    con.Execute " update file1_10 set barcode13 = null where barcode13 ='' "
    con.Execute " UPDATE FILE1_10 SET BARCODE2 = NULL WHERE BARCODE2 ='' "
    Dim MyItem13 As New ADODB.Recordset
    MyItem13.Open " SELECT * FROM FILE1_10 WHERE CODE = '095' AND MOSM = " & MyParn(cPMosm) & " ORDER BY MODEL ", con, adOpenStatic, adCmdText
    
    If MyItem13.RecordCount = 0 Then Exit Sub
    prog1.Min = 0
    prog1.Max = MyItem13.RecordCount
    prog1.Value = 0
    i = 0
    MyItem13.MoveFirst
    Do While Not MyItem13.EOF
        i = i + 1
        prog1.Value = i
        If Not IsNull(MyItem13!MODEL) Then
            FixBarCode2_Fact3 (MyItem13!MODEL)
        End If
        MyItem13.MoveNext
    Loop
End Sub

Private Sub old_FIXBARCODE2_Click()
    If MsgBox(" ÕœÌÀ »«—ﬂÊœ «·„’‰⁄ ·ﬂ· „ÊœÌ·«  «·„’‰⁄", vbYesNo + vbDefaultButton2) = vbYes Then
        
        
        con.Execute " update file1_10 set barcode13 = null where barcode13 ='' "
        Dim TableFact As New ADODB.Recordset
        Dim cScal As String, cModelFact As String
        TableFact.Open " SELECT * FROM FILE1_10 WHERE CODE = '001' AND ( BARCODE2 IS NULL OR BARCODE13 IS NULL ) ", con, adOpenStatic, adCmdText
        If TableFact.RecordCount = 0 Then Exit Sub
        TableFact.MoveLast
        prog1.Min = 0
        prog1.Max = TableFact.RecordCount
        prog1.Value = 0
        TableFact.MoveFirst
        With TableFact
        Do While Not .EOF
            If TableFact!MODELNO = "001W2003-2001-19" Then
                A = A
            End If
            
            i = i + 1
            prog1.Value = i
            Me.Caption = !Item
            cModelFact = UCase(DelZero(!modelfact))
'            If Mid(cModelFact, 1, 1) <> "I" And Mid(cModelFact, 1, 1) <> "G" And Mid(cModelFact, 1, 1) <> "B" And (Not IsNumeric(cModelFact) Or Len(cModelFact) < 8) Then
'                cModelFact = "00-" & cModelFact
'                cModelFact = Replace(cModelFact, "/", "\")
'            End If
            cScal = !SCAL
            If cScal = "0\3" Then cScal = "03"
            If cScal = "3\6" Then cScal = "06"
            If cScal = "6\12" Then cScal = "12"
            If cScal = "12\18" Then cScal = "18"
            If cScal = "18\24" Then cScal = "24"
            
            If cScal = "3/0" Then cScal = "03"
            If cScal = "6/3" Then cScal = "06"
            If cScal = "12/6" Then cScal = "12"
            If cScal = "18/12" Then cScal = "18"
            If cScal = "24/18" Then cScal = "24"
            cModelFact = UCase(DelZero(!modelfact))
                    
            If (!color) <> "" And Len(cScal) = 3 Then cScal = RetZero(DelZero(cScal), 2)
            cBarCode = GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(cModelFact) & " AND CODE_COLOR = " & MyParn(!color) & " AND SCAL = " & MyParn(cScal), confact)
            If cBarCode = "" Then
                If cScal = "---" Then cScal = "-"
                cBarCode = GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(cModelFact) & " AND CODE_COLOR = " & MyParn(!color) & " AND SCAL = " & MyParn(cScal), confact)
            End If
            If cBarCode <> "" Then
                aRet = aGetDesca("SELECT ITEM_BARCODE  ,BARCODE13 FROM FILE1_10 WHERE ITEM = " & MyParn(cBarCode), confact)
                If UBound(aRet) > 0 Then
                    cBarCodE2 = aRet(1)
                    If cBarCodE2 <> "" Then
                        con.Execute " UPDATE FILE1_10 SET BARCODE = " & addstring(cBarCode) & " , BARCODE2 = " & addstring(cBarCodE2) & "  , BARCODE13 = " & addstring(aRet(2)) & " , ISDATE = GETDATE() WHERE ITEM = " & Val(!Item)
                        confact.Execute " UPDATE FILE1_10 SET BARCODE_SHOP = " & Val(!Item) & " WHERE ITEM = " & addstring(cBarCode)
                    End If
                End If

''''                aRet = aGetDesca("SELECT ITEM_BARCODE  , BARCODE13 FROM FILE1_10 WHERE ITEM = " & MyParn(cBarCode), confact)
''''                If UBound(aRet) > 0 Then
''''                    cBarCodE2 = aRet(1)
''''                    If cBarCodE2 <> "" Then
''''                        If Not IsNull(aRet(2)) Then
''''                            Me.Caption = aRet(2) & ""
''''                        End If
''''
''''                        con.Execute " UPDATE FILE1_10 SET BARCODE = " & addstring(cBarCode) & " , BARCODE2 = " & addstring(cBarCodE2) & " , BARCODE13 = " & addstring(aRet(2)) & "  , ISDATE = GETDATE() WHERE ITEM = " & Val(!Item)
''''                        confact.Execute " UPDATE FILE1_10 SET BARCODE_SHOP = " & Val(!Item) & " WHERE ITEM = " & addstring(cBarCode)
''''
'''''                        con.Execute " DELETE FROM NEW_ITEM WHERE ITEM = " & sitem
'''''                        For nRow = 2 To nCountBranch
'''''                            sStore = RetZero(nRow, 2)
'''''                            con.Execute "  INSERT INTO NEW_ITEM (ITEM, BRANCH) VALUES        ( " & addvalue(sitem) & " , " & addstring(sStore) & " )  "
'''''                        Next nRow
''''
''''                    End If
''''                End If
            Else
                Me.Caption = "Erroe " & !Item & " " & !desca
            End If
            .MoveNext
        Loop
        End With
    End If


End Sub
Private Sub FIXBARCODE2_Click()
'    If MsgBox(" ÕœÌÀ »«—ﬂÊœ «·„’‰⁄ ·ﬂ· „ÊœÌ·«  «·„’‰⁄", vbYesNo + vbDefaultButton2) = vbYes Then
'        cStr1 = " Update file1_10 " & _
'                " set barcode =  (select top 1 fact.dbo.file1_10.item from fact.dbo.file1_10 where fact.dbo.file1_10.barcode2 = JUNIOR_SHOP.dbo.file1_10.barcode2 ) " & _
'                " ,barcode13 =  (select top 1 fact.dbo.file1_10.barcode13 from fact.dbo.file1_10 where fact.dbo.file1_10.barcode2 = JUNIOR_SHOP.dbo.file1_10.barcode2 ) " & _
'                " ,barcode_gs1 =  (select top 1 fact.dbo.file1_10.barcode_gs1 from fact.dbo.file1_10 where fact.dbo.file1_10.barcode2 = JUNIOR_SHOP.dbo.file1_10.barcode2 ) " & _
'                " where barcode2 is not null and code = '001'  and substring(barcode2,1,1) = 'j' "
'        con.Execute cStr1, nRec
'        MsgBox "   „  ⁄œÌ· ⁄œœ  " & nRec
'
'        cStr1 = " Update file1_10 " & _
'                " set barcode =  (select top 1 FACT_SOCK.dbo.file1_10.item from FACT_SOCK.dbo.file1_10 where FACT_SOCK.dbo.file1_10.barcode2 = JUNIOR_SHOP.dbo.file1_10.barcode2 ) " & _
'                " ,barcode13 =  (select top 1 FACT_SOCK.dbo.file1_10.barcode13 from FACT_SOCK.dbo.file1_10 where FACT_SOCK.dbo.file1_10.barcode2 = JUNIOR_SHOP.dbo.file1_10.barcode2 ) " & _
'                " ,barcode_gs1 =  (select top 1 FACT_SOCK.dbo.file1_10.barcode_gs1 from FACT_SOCK.dbo.file1_10 where FACT_SOCK.dbo.file1_10.barcode2 = JUNIOR_SHOP.dbo.file1_10.barcode2 ) " & _
'                " where barcode2 is not null and code = '095' and substring(barcode2,1,1) = 'S' "
'        con.Execute cStr1, nRec
'        MsgBox "   „  ⁄œÌ· ⁄œœ  " & nRec
'    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 19 And CanSave Then
        cmdSave_Click
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) And ActiveControl.Name <> xGroup.Name Then
        KeyCode = 0
        SendKeys "{tAB}"
    End If
End Sub
Private Sub Form_Load()
'SetKbLayout Lang_EN
Dim ImgXTmp As ImgX
Set ImgXTmp = New ImgX
'ImgXTmp.Import.FromMemoryFile LoadResData("BACKGROUND", "CUSTOM")
Set ImgX1.BackgroundImage = ImgXTmp.Image
Set ImgXTmp = Nothing
bEdit = True
ImgX1.AutoZoom = True
ImgX1.AutoZoomType = 1
nColWidth = vsScal.Width / 16
openCon con
openConDoc con_image

XMOSM2.AddItem "S"
XMOSM2.AddItem "W"
XMOSM2.AddItem "M"
If cBranch = "00" Then
    'openConFACT confact
    'openConFACT2 confact2
    'openConFACT3 confact3
Else
    cmdSave.Visible = False
    cmddel.Visible = False
    cmdcopy.Visible = False
    Frame7.Visible = False
End If

Set data1.Recordset = myRecordSet("SELECT * FROM FACT ORDER BY DESCA", con)
Set xFact.RowSource = data1
xFact.ListField = "DESCA"
xFact.BoundColumn = "CODE"

Set data2.Recordset = myRecordSet("SELECT * FROM FILE1_50 ORDER BY DESCA", con)
Set xGroup.RowSource = data2
xGroup.ListField = "DESCA"
xGroup.BoundColumn = "CODE"

Set data3.Recordset = myRecordSet("SELECT * FROM FILE1_10SC ORDER BY DESCA", con)
Set XSECTION.RowSource = data3
XSECTION.ListField = "DESCA"
XSECTION.BoundColumn = "CODE"

Set data4.Recordset = myRecordSet("SELECT * FROM MOSM ORDER BY DATE DESC ", con)
Set xMosm.RowSource = data4
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"

Set DATA11.Recordset = myRecordSet("SELECT * FROM FILE0_82 ORDER BY DESCA", con)
Set xmodelage.RowSource = DATA11
xmodelage.ListField = "DESCA"
xmodelage.BoundColumn = "CODE"

Set data12.Recordset = myRecordSet("SELECT * FROM FILE0_81 ORDER BY DESCA", con)
Set xmodelsex.RowSource = data12
xmodelsex.ListField = "DESCA"
xmodelsex.BoundColumn = "CODE"

Set DATA13.Recordset = myRecordSet("SELECT * FROM FILE8_51 ORDER BY DESCA", con)
Set xCharge.RowSource = DATA13
xCharge.ListField = "DESCA"
xCharge.BoundColumn = "CODE"

myDefineGrd
myDefine
If sModelItem <> 0 Then
    sModel = GetDesca("SELECT MODEL FROM FILE1_10 WHERE ITEM = " & sModelItem, con)
'    sModel = RetField("FILE1_10", "MODEL", "ITEM = " & sModelItem)
End If
If sModel <> "" Then
    myAddModel
    'vsScal.SetFocus
Else
'    openCardTable
'    myUndo
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set items = Nothing
closeCon con
End Sub

Private Sub ImgX1_Click()
Set show_model_photofrm.ImgX1.Picture = ImgX1.Picture
show_model_photofrm.Show 1
End Sub

Private Sub vsColor_GotFocus()
If vsColor.Row < 1 Then vsColor.Row = 1
vsColor.Col = vsColor.Cols - 1
End Sub

Private Sub vsColor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CellPos2 KeyAscii, vsColor.Row, vsColor.Col
End Sub

Private Sub vsColor_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 13 Then CellPos2 KeyAscii, Row, Col
End Sub

Private Sub vsColor_KeyupEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'If KeyCode = 13 Then CellPos2 KeyCode, vsColor.Row, vsColor.Col
End Sub
Private Sub vsColor_KeyUp(KeyCode As Integer, Shift As Integer)
With vsColor
If cmdSave.Enabled = False Then Exit Sub
'If KeyCode = 45 And .Col <> .Cols - 1 And validCol2(.Col) Then
'    Dim nMax As Long
'    InsertCol vsColor, .Col
'    FixCol2 .Col
'End If

If KeyCode = 46 And .Col <> .Cols - 1 And .Col <> 0 And .Row = 1 Then
    If MsgBox("Õ–› «··Ê‰ ø", vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
    If Trim(.TextMatrix(.Rows - 1, .Col)) = "" Then
        DelCol vsColor, .Col
        FixCol2 .Col
    Else
        On Error GoTo myError
        cString = "DELETE  FROM FILE1_10 WHERE MODEL = " & MyParn(XMODEL.text) & " AND c_color = " & MyParn(.TextMatrix(.Rows - 1, .Col))
        con.Execute cString, nAffect
        con.Execute " update file1_10 set item = item where model = " & MyParn(XMODEL.text)
        nCol = .Col
        myloadgrd2
        .Col = nCol
    End If
ElseIf KeyCode = 13 Then
'    CellPos2 KeyCode, Row, Col
End If
End With
Exit Sub
myError:
'con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub vsScal_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With vsScal
If Row = 0 Then
    If Col = .Cols - 1 Then
        If validCol(Col) Then
            .Cols = .Cols + 1
            CalcCol Col, 1
            FixCol Col
            FixBackColor
        End If
    ElseIf Row = 0 And .Cols = 3 And Col = 1 And Trim(.TextMatrix(0, .Cols - 1)) = "" And Trim(.TextMatrix(0, Col)) <> "---" Then
        .Col = .Cols - 1
    End If
Else
'    CalcRate Col, Row
End If
If Row = 2 Then ISNEWPRICE.Value = 1
CalcTotals
If Val(.TextMatrix(4, Col)) = 0 Then .TextMatrix(4, Col) = Val(.TextMatrix(1, Col))
End With
End Sub
Private Sub VsColor_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With vsColor
If Col = .Cols - 1 Then
    If validCol2(Col) Then
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 2) = Val(.TextMatrix(0, .Cols - 3)) + 1
        FixCol2 Col
        .Col = Col + 1
        FixBackColor2
    End If
End If
End With
End Sub
Private Sub VsColor_Validate(Cancel As Boolean)
If (Not validCol2(vsColor.Col)) And vsColor.Col <> vsColor.Cols - 1 And vsColor.Col <> 0 Then DelCol vsColor, vsColor.Col
End Sub
Private Sub vsScal_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With vsScal
If OldCol <> NewCol And OldCol <> .Cols - 1 And OldCol <> 0 And vsScal.TextMatrix(vsScal.Rows - 1, OldCol) = "" Then
    If Not validCol(OldCol) Then
        DelCol vsScal, OldCol
    End If
End If
'If OldCol <> NewCol And NewCol < .Cols - 1 Then CalcCol NewCol, NewRow
End With
End Sub
Private Sub VsColor_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With vsColor
If OldCol <> NewCol And OldCol <> .Cols - 1 And OldCol <> 0 Then
    If Not validCol2(OldCol) Then
        DelCol vsColor, OldCol
        FixCol2 OldCol
    End If
End If
End With
End Sub
Private Sub vsScal_EnterCell()
'If (vsScal.Row = 0 And vsScal.TextMatrix(vsScal.Rows - 1, vsScal.Col) <> "") Or cmdSave.Enabled = False Then
If cmdSave.Enabled = False Then
    vsScal.Editable = flexEDNone
Else
    vsScal.Editable = flexEDKbdMouse
End If
If vsScal.Row = 1 And vsScal.Editable = flexEDKbdMouse Then
    CalcCol vsScal.Col, vsScal.Row
End If

End Sub

Private Sub vsScal_GotFocus()
If vsScal.Row < 1 Then vsScal.Row = 1
vsScal.Col = vsScal.Cols - 2
End Sub
Private Sub vsScal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And vsScal.Col = vsScal.Cols - 1 And vsScal.TextMatrix(vsScal.Row, vsScal.Col) = "" And vsScal.Row = 5 Then
'    On Error Resume Next
    KeyAscii = 0
   ' cmdScan_Click
    Exit Sub
End If
If KeyAscii = 13 Then
    CellPos KeyAscii, vsScal.Row, vsScal.Col
End If
End Sub
Private Sub vsScal_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 13 Then
    CellPos KeyAscii, Row, Col
End If
End Sub
Private Sub vsScal_KeyUp(KeyCode As Integer, Shift As Integer)
With vsScal
    If cmdSave.Enabled = False Then Exit Sub
    If KeyCode = 45 And .Col <> .Cols - 1 And validCol(.Col) Then
        InsertCol vsScal, .Col
        CalcTotals
        FixCol .Col
   ElseIf KeyCode = 46 And .Col <> .Cols - 1 And .Col <> 0 And .Row = 0 Then
        If Not IsNumeric(.TextMatrix(.Rows - 1, .Col)) Then
            DelCol vsScal, .Col
            CalcTotals
            FixCol .Col
        Else
            If MsgBox("Õ–› «·„ﬁ«” ø", vbYesNo + vbDefaultButton2) <> vbYes Then
                bAdd = False
                Exit Sub
            End If
            On Error GoTo myError
            cString = "DELETE  FROM FILE1_10 WHERE MODEL = " & MyParn(XMODEL.text) & " AND CODE_SCAL = " & .TextMatrix(.Rows - 1, .Col)
            con.Execute cString
            con.Execute " update file1_10 set item = item where model = " & MyParn(XMODEL.text)
            nCol = .Col
            myLoadGrd
            .Col = nCol
        End If
    End If
End With
Exit Sub
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long, Optional bCalc As Boolean = True)
With vsScal
If cmdSave.Enabled = False Then Exit Sub
If (Not validCol(Col)) Then Exit Sub
'--- Õ«·… «Ê· ’› Ê«·„ﬁ«”
If Row = 0 And .Cols = 3 And Col = 1 And Trim(.TextMatrix(0, .Cols - 1)) = "" And Trim(.TextMatrix(0, Col)) = "---" Then
    Exit Sub
End If
KeyCode = 0
If Col < vsScal.Cols - 1 Then vsScal.Col = Col + 1
If vsScal.Col = vsScal.Cols - 1 Then
    If Row = 0 Then CalcCol vsScal.Col, 0
    If validCol(.Cols - 1) Then
        .Cols = .Cols + 1
        CalcCol vsScal.Col, 1
        FixCol .Cols - 1
    Else
        If Row = 1 Or Row = 2 Then
            .Row = Row + 1
            .Col = 1
            If CanSave Then cmdSave_Click
        ElseIf Row = 0 And Col > 1 Then
            .Row = .Row + IIf(Row < 3, 1, 2)
            .Col = 1
        End If
    End If
Else
'    If emptyRow(vsScal, Row, .Col) Then
'        Dim i As Long
'        For i = vsScal.Col To .Cols - 2
'            CalcCol i, Row
'        Next
'    Else
    CalcCol vsScal.Col, Row
'    End If
End If
End With
End Sub
Private Sub CellPos2(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
If Not validCol2(Col) Then Exit Sub
If Col < vsColor.Cols - 1 Then vsColor.Col = Col + 1
KeyCode = 0
'VsColor_AfterEdit Row, Col
End Sub
Private Sub DelCol(pGrid As Variant, pos As Long)
With pGrid
.Cols = .Cols + 1
.ColPosition(pos) = .Cols - 1
.Cols = .Cols - 2
End With
End Sub
Private Sub FixBackColor()
With Me.vsScal
.Cell(flexcpBackColor, 0, 1, 0, .Cols - 1) = &H8000000F
For i = 1 To .Cols - 1
    .Cell(flexcpBackColor, 0, i, 0, i) = IIf(Trim(.TextMatrix(.Rows - 1, i)) = "", &HFFFFFF, &HE0E0E0)
Next
End With
End Sub
Private Sub VsColor_EnterCell()
With vsColor
    If (.Row = 1 Or .Row = 2) And cmdSave.Enabled Then .Editable = flexEDKbdMouse Else .Editable = flexEDNone
End With
End Sub
Private Sub VsColor_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'With vsColor
'If Col = .Cols - 1 Then
'    .Cols = .Cols + 1
'    FixBackColor2
'End If
'End With
End Sub
Private Sub VsColor_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With vsColor
If Trim(.EditText) = "" Then
    MsgBox "·Ê‰ €Ì— „”Ã·"
    Cancel = True
Else
    .EditText = Trim(.EditText)
    For i = 1 To .Cols - 2
        If i <> Col Then
            If Trim(LCase(.EditText)) = LCase(.TextMatrix(1, i)) Then
                MsgBox " „ «œŒ«· ‰›” «··Ê‰ „‰ ﬁ»·"
                Cancel = True
            End If
        End If
    Next
End If
End With
End Sub
Private Function validCol(nCol)
With vsScal
If Trim(vsScal.TextMatrix(0, nCol)) = "" Then Exit Function
For i = 1 To .Cols - 1
    If i <> nCol Then
        If Trim(.TextMatrix(0, nCol)) = .TextMatrix(0, i) Then
            'MsgBox " „ «œŒ«· ‰›” «·„ﬁ«” „‰ ﬁ»·"
            Cancel = True
            Exit Function
        End If
    End If
Next
End With
validCol = True
End Function
Private Function validCol2(nCol) As Boolean
If Trim(vsColor.TextMatrix(1, nCol)) = "" Then Exit Function
validCol2 = True
End Function
Private Sub vsScal_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'With vsScal
'If Col = .Cols - 1 Then
'    .Cols = .Cols + 1
'    FixBackColor
'End If
'End With
End Sub
Private Sub vsScal_Validate(Cancel As Boolean)
If (Not validCol(vsScal.Col)) And vsScal.Col <> vsScal.Cols - 1 And vsScal.Col <> 0 And vsScal.TextMatrix(vsScal.Rows - 1, vsScal.Col) = "" Then DelCol vsScal, vsScal.Col
End Sub
Private Sub VsScal_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With vsScal
If Row = 0 Then
    If Trim(.EditText) = "" Then
        'MsgBox "„ﬁ«” €Ì— „”Ã·"
        'Cancel = True
    Else
        .EditText = UCase(.EditText)
        For i = 1 To .Cols - 2
            If i <> Col Then
                If Trim(.EditText) = .TextMatrix(0, i) Then
                    MsgBox " „ «œŒ«· ‰›” «·„ﬁ«” „‰ ﬁ»·"
                    Cancel = True
                End If
            End If
        Next
    End If

End If
If Row = 1 Then
    If Val(.TextMatrix(1, Col)) = Val(.TextMatrix(4, Col)) Then
        .TextMatrix(4, Col) = Val(.EditText)
    End If
End If
End With

End Sub
Private Sub CmdFirst_Click()
If Not openTablePhotos(tbMode.tbFirst) Then
    myDefinePhotos
End If
End Sub
Private Sub CmdLast_Click()
If Not openTablePhotos Then
    myDefinePhotos
End If
End Sub
Private Sub CmdNext_Click()
If Not openTablePhotos(tbMode.tbNext, xId.Caption) Then
    If Not openTablePhotos(tbMode.tblast) Then
        myDefinePhotos
    End If
End If
End Sub
Private Sub CmdPrevious_Click()
If Not openTablePhotos(tbMode.tbPrevious, xId.Caption) Then
       If Not openTablePhotos(tbMode.tbFirst) Then
        myDefinePhotos
    End If
End If
End Sub
Private Sub myLoadGrd()
Dim loctable As New ADODB.Recordset, i As Integer
loctable.CursorLocation = adUseClient

'                      0    1                   2                           3                       4                   5                           6                               7                               8                   9           10
loctable.Open "select scal,max(Cost)  as cost ,max(Price ) as price ,max(PRICE2) as price2 ,max(Cost2) as cost2 , MAX(PRICE_2) as price_2 , MAX(COSTITEM) as COSTITEM , MAX(COSTITEM_FR) as COSTITEM_FR , MAX(PRICE_P) AS PRICE_P, code_scal ,min(Price ) as maxprice  from file1_10 where model = " & MyParn(XMODEL.text) & " group by file1_10.c_scal,scal,file1_10.code_Scal order by file1_10.c_scal ", con, adOpenForwardOnly, adLockReadOnly, adCmdText
With vsScal
.Cols = 1
Do Until loctable.EOF
    .Cols = .Cols + 1
    For i = 0 To 9
        .TextMatrix(i, .Cols - 1) = loctable.Fields(i) & ""
        If i = 2 Then .TextMatrix(i, .Cols - 1) = loctable!maxprice
    Next
    loctable.MoveNext
Loop
If cmdSave.Enabled Then .Cols = .Cols + 1
fixGrd
CalcTotals
loctable.Close
Set loctable = Nothing
End With
End Sub
Private Sub myloadgrd2()
Dim loctable As New ADODB.Recordset
loctable.CursorLocation = adUseServer
loctable.Open "select file1_10.C_Color,Color,Color_E,FILE1_10.CODE_COLOR from file1_10 where model = " & MyParn(XMODEL.text) & " group by file1_10.C_color,COLOR,COLOR_E,CODE_COLOR order by file1_10.C_color", con, adOpenStatic, adLockReadOnly, adCmdText
With vsColor
.Cols = 1
Do Until loctable.EOF
    .Cols = .Cols + 1
    .TextMatrix(0, .Cols - 1) = loctable!c_Color & ""
    .TextMatrix(1, .Cols - 1) = loctable!color & ""
    .TextMatrix(2, .Cols - 1) = loctable!Color_E & ""
    .TextMatrix(.Rows - 1, .Cols - 1) = loctable!CODE_COLOR & ""
    loctable.MoveNext
Loop
If cmdSave.Enabled Then .Cols = .Cols + 1
Fixgrd2
End With
loctable.Close
Set loctable = Nothing
End Sub
Private Sub fixGrd()
With vsScal
.ColWidth(0) = 1300
For i = 1 To .Cols - 1
    .ColWidth(i) = nColWidth
    .ColAlignment(i) = flexAlignRightCenter
Next
'.RowHidden(4) = True

.RowHidden(.Rows - 1) = True
End With
FixBackColor
End Sub
Private Sub Fixgrd2()
With vsColor
.ColWidth(0) = 2000
For i = 1 To .Cols - 1
    .ColWidth(i) = 1500
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
FixBackColor2
End Sub
Private Sub FixBackColor2()
With Me.vsColor
'.Cell(flexcpBackColor, 0, 1, 0, .Cols - 1) = &H8000000F
'For i = 1 To .Cols - 1
'    .Cell(flexcpBackColor, 0, i, 0, i) = IIf(Trim(.TextMatrix(0, i)) = "", &HFFFFFF, &H8000000F)
'Next
End With
End Sub
Private Sub InsertCol(pGrid As Variant, pos%)
With pGrid
.Cols = .Cols + 1
.ColPosition(.Cols - 1) = pos
End With
End Sub
Private Function myreplace() As Boolean
    myreplaceGrd
    myreplace = True
    Exit Function
myError:
    MsgBox Err.Description
    Err.Clear
End Function
Private Function myValid() As Boolean
If xCharge.BoundText <> "" And XISNOITEM.Value = 0 And XISNOITEM2.Value = 0 Then
    MsgBox "„’—Ê› ·’‰› €Ì— Œœ„… «Ê «’· "
    Exit Function
End If

'If XISNOITEM.Value <> 0 And XISNOITEM2.Value = 0 Then
'    MsgBox "’‰› Œœ„… Ê «’Ê· ÌÃ» «· ⁄œÌ·"
'    Exit Function
'End If

If xFact.BoundText = "" Then
    MsgBox "«·„’‰⁄"
    Exit Function
End If
If XISNOITEM2.Value <> 0 And XISNOITEM.Value <> 0 Then
    MsgBox "€Ì—  „”„ÊÕ ’‰› Œœ„… Ê «’· "
    Exit Function
End If
If DelZero(xModelFact.text) = "" Then
    MsgBox "—ﬁ„ «·„ÊœÌ·"
    Exit Function
End If
If xMosm.BoundText = "" Then
    MsgBox "«·„Ê”„ €Ì— „”Ã·"
    Exit Function
End If

If XSECTION.BoundText = "" Then
    MsgBox "«·ﬁ”„ €Ì— „”Ã·"
    Exit Function
End If
If GetDesca("select [group] from file1_50 where code = " & MyParn(xGroup.BoundText), con) <> XSECTION.BoundText Then
    MsgBox "«·„Ã„Ê⁄… „Õ ·›… ⁄‰ «·ﬁ”„ "
    Exit Function
End If
If xDescA.text = "" Then
    MsgBox "≈”„ «·’‰› €Ì— „”Ã·"
    Exit Function
End If

If xGroup.BoundText = "" Then
    MsgBox "«·„Ã„Ê⁄… €Ì— „”Ã·…"
    Exit Function
End If

With vsScal
    For nCol = 1 To .Cols - 2
        If Val(.TextMatrix(1, nCol)) = 0 Then
           MsgBox "”⁄— „’‰⁄ €Ì— „”Ã·"
        End If
        
        If Val(.TextMatrix(2, nCol)) = 0 Then
           MsgBox "”⁄— „” Â·ﬂ €Ì— „”Ã·"
        End If
        
        If .TextMatrix(0, nCol) = "" Then
            MsgBox "„ﬁ«” €Ì— „”Ã·"
            Exit Function
        End If
        For K = nCol + 1 To .Cols - 2
            If Trim(.TextMatrix(0, nCol)) = Trim(.TextMatrix(0, K)) Then
                MsgBox "«·„ﬁ«” „ﬂ——"
                Exit Function
            End If
        Next
    Next
End With
      
With vsColor
    For nCol = 1 To .Cols - 2
        If .TextMatrix(0, nCol) = "" Then
            MsgBox "·Ê‰ €Ì— „”Ã·"
            Exit Function
        End If
        For K = nCol + 1 To .Cols - 2
            If Trim(.TextMatrix(0, nCol)) = Trim(.TextMatrix(0, K)) Then
                MsgBox "«··Ê‰ „ﬂ——"
                Exit Function
            End If
        Next
    Next
End With
myValid = True
End Function
Private Sub Handlecontrols(nMode)
cmddel.Enabled = (nMode = LoadMode) And bEdit
cmdSave.Enabled = bEdit
cmdUndo.Enabled = bEdit
'CmdInform.Enabled = (nMode = LoadMode)
XMODEL.Tag = nMode
xModelFact.Enabled = False
'xSection.Locked = sSection <> ""
'xSection.TabStop = sSection = ""
XMODEL.Tag = nMode
XMODEL.Enabled = Not (nMode = LoadMode)
End Sub
Private Sub myDefineGrd()
With vsScal
    .Cols = 1
    .Rows = 10
    .TextMatrix(0, 0) = "«·„ﬁ«”"
    .TextMatrix(1, 0) = "”⁄— «·„’‰⁄"
    .TextMatrix(2, 0) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(3, 0) = "”⁄— √Êﬂ«“ÌÊ‰"
    .TextMatrix(4, 0) = "”⁄— Ã„·…"
    .TextMatrix(5, 0) = "”⁄— «Êﬂ«“ÊÌ‰ 2"
    
    .TextMatrix(6, 0) = " ﬂ·›… ›—Ê⁄"
    .TextMatrix(7, 0) = " ﬂ·›…  ÊﬂÌ·« "
    .TextMatrix(8, 0) = "”⁄— ÿ»«⁄…2"
    
    .TextMatrix(9, 0) = ""
    
'    .RowHidden(4) = (cBranch > "60")
'    .RowHidden(7) = (cBranch > "60")
'    .RowHidden(8) = (cBranch <> "00")
    .RowHidden(1) = Not bOpt10
    'If cBranch > "60" Then .TextMatrix(6, 0) = " ﬂ·›… "
    .RowHidden(6) = True
    .RowHidden(7) = True
    .RowHidden(8) = True
    
    .FixedCols = 1
    .Cols = .Cols + 1
    .TextMatrix(0, 1) = "---"
    .Cols = .Cols + 1
    fixGrd
End With
With vsColor
    .Cols = 1
    .Rows = 4
    .TextMatrix(0, 0) = "«·ﬂÊœ"
    .TextMatrix(1, 0) = "«··Ê‰"
    .TextMatrix(2, 0) = "COLOR"
    .Cols = .Cols + 1
    .TextMatrix(1, 1) = "---"
    .TextMatrix(0, 1) = 1
    .Cols = .Cols + 1
    'vsColor.RowHidden(vsColor.Rows - 1) = True
    Fixgrd2
End With
End Sub

Private Sub XCHARGE_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then ChargeLook
End Sub
Private Sub xGroup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    If Trim(xGroup.text) <> "" And Not xGroup.MatchedWithList Then
        If MsgBox("·« ÌÊÃœ „Ã„Ê⁄… »Â–« «·«”„ !! «÷«›…", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
        
        Dim sBound As String
        KeyCode = 0
        
        Dim aInsert As Variant, sCode As String
        sCode = RetZero(Newflag("FILE1_50", "CODE", con), 2)
        aInsert = AddFlag(Empty, "CODE", addstring(sCode))
        aInsert = AddFlag(aInsert, "DESCA", addstring(xGroup.text))
        con.BeginTrans
        
        On Error GoTo myError
        con.Execute addInsert(aInsert, "FILE1_50")
        con.CommitTrans
    
        data2.Recordset.Requery
        xGroup.BoundText = sCode
        If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
    Else
        xDescA.SetFocus
    End If
    If Trim(xDescA.text) = "" Then xDescA.text = xGroup.text
End If
Exit Sub
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub

Private Sub xGroup_Validate(Cancel As Boolean)
On Error GoTo myError
Dim cGr As String
'If (Not xGroup.MatchedWithList) And xGroup.BoundText <> "" Then
'    If MsgBox("„Ã„Ê⁄… ÃœÌœ… Â·  —Ìœ  ”ÃÌ·Â« ", vbYesNo) = vbYes Then
'        Dim sGroup As String
'        sGroup = RetZero(Newflag("FILE1_50", "CODE", con), 3)
'        con.BeginTrans
'        cString = " INSERT INTO FILE1_50 (CODE,DESCA)" & _
'                  " VALUES(" & addstring(sGroup) & "," & addstring(xGroup.Text) & ")"
'        con.Execute cString
'        con.CommitTrans
'        data2.Refresh
'        xGroup.BoundText = sGroup
'    End If
'End If
If Trim(xDescA.text) = "" Then xDescA.text = xGroup.text
Exit Sub
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(xitem) Then
        If cFilter <> "" Then
            cFilter = ""
            openCardTable
        End If
        With CardTable
            Dim cString As String
            cString = "SELECT MODEL FROM FILE1_10"
            cString = cString & turn(cString) & " ITEM = " & Val(xitem.text)
            XMODEL.text = GetDesca(cString, con)
            openCardTable
            If XMODEL.text <> "" Then
                CardTable.Find "Model = " & MyParn(XMODEL.text), , adSearchForward, adBookmarkFirst
                If Not CardTable.EOF Then myload Else myDefine
            Else
                myDefine
            End If
        End With
    End If
End Sub
Private Sub LoadPhoto()
On Error GoTo myError
ImgX1.Images.Clear
ImgX1.Import.FromFile retFile(xModelNo.text)
Exit Sub
myError:
Err.Clear
End Sub

Private Sub Twain_TwainError(ByVal erNum As Long, ByVal erSource As String, ByVal Description As String)
MsgBox "Error Number:  " & erNum & vbCrLf & Description, vbInformation, erSource
End Sub

Private Sub FixCol(ByVal pCol As Long, Optional nMax As Long)
With vsScal
For i = pCol To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
    .ColWidth(i) = nColWidth
Next
End With
End Sub
Private Sub FixCol2(ByVal pCol As Long, Optional nMax As Long)
With vsColor
For i = pCol To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
    .ColWidth(i) = nColWidth
Next
End With
End Sub
Private Sub xitem_GotFocus()
xitem.SelStart = 0
xitem.SelLength = Len(xitem.text)
xitem.BackColor = &H80FFFF
End Sub
Private Sub xModel_Change()
If Not openTablePhotos Then
    myDefinePhotos
End If
End Sub

Private Sub xModelFact_Validate(Cancel As Boolean)
If xModelFact.text = "" Then
    Cancel = True
    Exit Sub
End If
xModelFact.text = RetZero(xModelFact.text, 10)
XMODEL.text = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.text, 10) & xSupp.text
xModelNo.text = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.text, 10)
openCardTable

CardTable.Find "MODEL = " & MyParn(XMODEL.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    xModelFact.Enabled = False
    myload
End If
End Sub
Private Sub xModelNo_GotFocus()
xModelNo.SelStart = 0
xModelNo.SelLength = Len(xModelNo.text)
xModelNo.BackColor = &H80FFFF
End Sub
Private Sub xModel_GotFocus()
XMODEL.SelStart = 0
XMODEL.SelLength = Len(XMODEL.text)
XMODEL.BackColor = &H80FFFF
End Sub
Private Sub xModelFACT_GotFocus()
xModelFact.SelStart = 0
xModelFact.SelLength = Len(xModelFact.text)
xModelFact.BackColor = &H80FFFF
End Sub

Private Sub xOkaz_2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With vsScal
    For nCol = 1 To .Cols - 2
        .TextMatrix(5, nCol) = Round(Val(Val(.TextMatrix(2, nCol))) * ((100 - Val(xOkaz_2.text)) / 100), 0)
    Next nCol
End With

End If
End Sub
Private Sub xOkaz_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With vsScal
    For nCol = 1 To .Cols - 2
        .TextMatrix(3, nCol) = Round(Val(Val(.TextMatrix(2, nCol))) * ((100 - Val(XOKAZ.text)) / 100), 0)
    Next nCol
End With

End If
End Sub

Private Sub xRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MsgBox(" €ÌÌ— «·‰”»… ·ﬂ· «·«’‰«›", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
    Dim nRow As Long, nCol As Long
    For nCol = 1 To vsScal.Cols - 1
        For nRow = 2 To 3
            vsScal.TextMatrix(nRow, nCol) = ""
        Next
    Next
    For nCol = 1 To vsScal.Cols - 1
        CalcCol nCol, 1
    Next
End If
End Sub

Private Sub xSupp_GotFocus()
xSupp.SelStart = 0
xSupp.SelLength = Len(xSupp.text)
xSupp.BackColor = &H80FFFF
End Sub
Private Sub xFact_GotFocus()
xFact.BackColor = &H80FFFF
End Sub
Private Sub xMosm_GotFocus()
xMosm.BackColor = &H80FFFF
End Sub
Private Sub xOkaz_GotFocus()
    XOKAZ.SelStart = 0
    XOKAZ.SelLength = Len(XOKAZ.text)
    XOKAZ.BackColor = &H80FFFF
End Sub
Private Sub xRedem_GotFocus()
xRedem.SelStart = 0
xRedem.SelLength = Len(xRedem.text)
xRedem.BackColor = &H80FFFF
End Sub
Private Sub xDescA_GotFocus()
If Trim(xDescA.text) = Trim(xGroup.text) Then
    xDescA.text = xGroup.text & " "
    xDescA.SelStart = Len(xDescA.text)
    xDescA.SelLength = 1
Else
    xDescA.SelLength = Len(xDescA.text)
End If
xDescA.BackColor = &H80FFFF
End Sub
Private Sub xcode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.text)
xCode.BackColor = &H80FFFF
End Sub
Private Sub xRate_GotFocus()
xRate.SelStart = 0
xRate.SelLength = Len(xRate.text)
xRate.BackColor = &H80FFFF
End Sub
Private Sub xGroup_GotFocus()
'    DATA2.ConnectionString = strCon
'    If XSECTION.BoundText <> "" Then
'        DATA2.RecordSource = "SELECT * FROM FILE1_50 WHERE [GROUP] = " & Val(XSECTION.BoundText) & " ORDER BY DESCA"
'    Else
'        DATA2.RecordSource = "SELECT * FROM FILE1_50 ORDER BY DESCA"
'    End If
'    Set xGroup.RowSource = DATA2
'    xGroup.ListField = "DESCA"
'    xGroup.BoundColumn = "CODE"
'    DATA2.Refresh

xGroup.BackColor = &H80FFFF
End Sub
Private Sub xSection_GotFocus()
XSECTION.BackColor = &H80FFFF
End Sub
Private Sub xitem_LostFocus()
xitem.BackColor = &H80000005
End Sub
Private Sub xModelNO_LostFocus()
xModelNo.BackColor = &H80000005
End Sub
Private Sub XModel_LostFocus()
XMODEL.BackColor = &H80000005
End Sub
Private Sub XModelFACT_LostFocus()
xModelFact.BackColor = &H80000005
End Sub
Private Sub xSupp_LostFocus()
xSupp.BackColor = &H80000005
End Sub
Private Sub xFACT_LostFocus()
xFact.BackColor = &H80000005
End Sub
Private Sub xMosm_LostFocus()
xMosm.BackColor = &H80000005
End Sub
Private Sub xOkaz_LostFocus()
    XOKAZ.BackColor = &H80000005
End Sub
Private Sub xRedem_LostFocus()
xRedem.BackColor = &H80000005
End Sub
Private Sub xDesca_LostFocus()
xDescA.BackColor = &H80000005
End Sub
Private Sub xcode_LostFocus()
xCode.BackColor = &H80000005
End Sub
Private Sub xRate_LostFocus()
xRate.BackColor = &H80000005
End Sub

Private Sub xgroup_LostFocus()
    xGroup.BackColor = &H80000005
    If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub xSection_LostFocus()
XSECTION.BackColor = &H80000005
End Sub
Private Sub myUndo()
'On Error GoTo myerror
If (CardTable.BOF And CardTable.EOF) Then
    myDefine
Else
    If Trim(XMODEL.text) <> "" Then
        CardTable.Find "MODEL = " & MyParn(XMODEL.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "select TOP 1 FILE1_10H.* , FILE4_10.DESCA AS SUPPNAME from file1_10H LEFT JOIN FILE4_10 ON FILE4_10.CODE = FILE1_10H.CODE "
If XMODEL.text <> "" Then cString = cString & turn(cString) & "FILE1_10H.MODEL = " & MyParn(XMODEL.text)
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " order by model desc"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Function LastCol(pGrid As Variant) As Integer
With pGrid
For i = 1 To .Cols - 1
    If Val(.TextMatrix(.Rows - 1, i)) > LastCol Then LastCol = Val(.TextMatrix(.Rows - 1, i))
Next
LastCol = LastCol + 1
End With
End Function
Private Sub FixCode(pGrid As Variant)
With pGrid
For i = 1 To .Cols - 1
    If .TextMatrix(.Rows - 1, i) = "" Then .TextMatrix(.Rows - 1, i) = LastCol(pGrid)
Next
End With
End Sub
Private Sub CalcCol(Col As Long, Row As Long)
Dim sValue As String
With vsScal
If Col > 1 Then
    If Row = 0 Then
       If IsNumeric(.TextMatrix(0, Col - 1)) And IsNumeric(.TextMatrix(0, Col - 2)) And Val(.TextMatrix(0, Col - 1)) > Val(.TextMatrix(0, Col - 2)) And .TextMatrix(0, Col) = "" Then
            sValue = Val(.TextMatrix(0, Col - 1)) + (Val(.TextMatrix(0, Col - 1)) - Val(.TextMatrix(0, Col - 2)))
        ElseIf (Not IsNumeric(.TextMatrix(0, Col - 1))) And .TextMatrix(0, Col) = "" Then
            Select Case .TextMatrix(0, Col - 1)
                Case "XS"
                    sValue = "S"
                Case "S"
                    sValue = "M"
                Case "M"
                    sValue = "L"
                Case "L"
                    sValue = "XL"
                Case "XL"
                    sValue = "XXL"
                Case "XXL"
                    sValue = "3XL"
                Case "3XL"
                    sValue = "4XL"
                Case "4XL"
                    sValue = "5XL"
            End Select
        End If
        If sValue <> "" Then .TextMatrix(0, Col) = sValue
    End If
    If Col < .Cols - 1 Then
        If (Row = 1 Or Row = 2 Or Row = 3 Or Row = 4) And Col > 1 Then
            If IsNumeric(.TextMatrix(Row, Col - 1)) And IsNumeric(.TextMatrix(Row, Col - 2)) And Val(.TextMatrix(Row, Col - 1)) >= Val(.TextMatrix(Row, Col - 2)) And .TextMatrix(Row, Col) = "" Then
                .TextMatrix(Row, Col) = Val(.TextMatrix(Row, Col - 1)) + (Val(.TextMatrix(Row, Col - 1)) - Val(.TextMatrix(Row, Col - 2)))
                .TextMatrix(4, Col) = Val(.TextMatrix(4, Col - 1)) + (Val(.TextMatrix(4, Col - 1)) - Val(.TextMatrix(4, Col - 2)))
                 
            End If
         End If
    End If
End If
CalcRate Col, Row
End With
End Sub
Private Sub myAddModel()
XMODEL.text = sModel
openCardTable
If Not (CardTable.EOF And CardTable.BOF) Then
    myUndo
    Exit Sub
End If

Dim sModelNo As String
XMODEL.text = ""
sModelNo = SFact & SMosm & RetZero(SModelFact, 10)
cFilter = "ModelNo = " & MyParn(sModelNo)
openCardTable
myUndo
If (XMODEL.text <> sModel) Then
    XMODEL.Tag = DefineMode
    xFact.BoundText = SFact
    xMosm.BoundText = SMosm
    xModelFact.text = SModelFact
    xModelNo.text = SFact & SMosm & RetZero(SModelFact, 10)
    XMODEL.text = sModel
    xSupp.text = sSupp
    XSECTION.BoundText = sSection
    xCode.text = sCode
    If xCode.text <> "" Then
        Dim aRet As Variant
        aRet = aGetDesca("Select desca,Rate from file4_10 where code = " & MyParn(xCode.text), con)
        If UBound(aRet) > 0 Then
            xCodeDesc.Caption = aRet(1)
            xRate.text = aRet(2) & ""
        End If
    End If
    For i = 1 To vsScal.Cols - 1
        vsScal.TextMatrix(vsScal.Rows - 1, i) = ""
    Next
    FixBackColor
End If
cFilter = ""
'cFilter = "FILE1_10H.MODEL = " & MyParn(sModel)
End Sub
Private Function retColValue(Row As Long, Col As Long)
With vsScal
If IsNumeric(.TextMatrix(0, Col)) And IsNumeric(.TextMatrix(0, Col - 1)) And Val(.TextMatrix(0, Col)) > Val(.TextMatrix(0, Col - 1)) And Col < .Cols - 1 And .TextMatrix(0, Col + 1) = "" Then
    sValue = Val(.TextMatrix(0, Col)) + (Val(.TextMatrix(0, Col)) - Val(.TextMatrix(0, Col - 1)))
ElseIf (Not IsNumeric(.TextMatrix(0, Col))) And .TextMatrix(0, Col + 1) = "" Then
    Select Case .TextMatrix(0, Col)
        Case "XS"
            sValue = "S"
        Case "S"
            sValue = "M"
        Case "M"
            sValue = "L"
        Case "L"
            sValue = "XL"
        Case "XL"
            sValue = "XXL"
        Case "XXL"
            sValue = "3XL"
        Case "3XL"
            sValue = "4XL"
        Case "4XL"
            sValue = "5XL"
    End Select
End If
For i = 1 To .Cols - 1
    If sValue = Trim(UCase(.TextMatrix(0, i))) Then Exit Function
Next
retColValue = sValue
End With
End Function
Private Sub FixRate()
Dim loctable As New ADODB.Recordset
loctable.Open "select * from ITEM_DEF WHERE CODE = 1", con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    sRate = Myvalue(loctable!Rate)
    sRedem = Myvalue(loctable!Redem)
    sOkaz = Myvalue(loctable!okaz)
    sBig = Myvalue(loctable!bIg)
End If
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CalcRate(Col, Row)
With vsScal
If Col < .Cols - 1 And Col > 0 Then
    'If Row <> 2 And Trim(.TextMatrix(2, Col)) = "" Then
     If Row <> 2 Then
        If Val(xRate.text) <> 0 And Val(.TextMatrix(1, Col)) > 0 And (Val(.TextMatrix(3, Col)) = 0 Or Row = 1) Then
'            Dim nPrice As Single, nRem As Single
'            nPrice = Val(.TextMatrix(1, Col)) * (100 + Val(xRate.Text)) / 100
'            nRem = nPrice - Int(nPrice)
'            Select Case nRem
'                Case Is = 0
'                    nPrice = Int(nPrice)
'                Case Is > 0
'                    nPrice = Int(nPrice) + 1
'                Case Is < 0.5
'                    nPrice = Int(nPrice) + 0.5
'            End Select
'            .TextMatrix(2, Col) = Myvalue(nPrice, "#0.00")
        End If
    End If
    
    'If Row <> 3 And Trim(.TextMatrix(3, Col)) = "" Then
     If Row <> 3 Then
'        If Val(xOkaz.Text) <> 0 Then
'            If (Val(.TextMatrix(3, Col)) = 0 Or Row = 1) And Val(xOkaz.Text) > 0 Then
''                .TextMatrix(3, Col) = Myvalue(Val(.TextMatrix(1, Col)) * (100 + Val(xOkaz.Text)) / 100, "#0.00")
'            End If
'        Else
'            .TextMatrix(3, Col) = 0
'        End If
    End If
    'If Row <> 4 And Trim(.TextMatrix(4, Col)) = "" Then
'     If Row <> 4 Then
'        If Val(xbig.Text) <> 0 And Val(.TextMatrix(1, Col)) > 0 And (Val(.TextMatrix(4, Col)) = 0 Or Row = 1) Then
'            nPrice = Val(.TextMatrix(1, Col)) * (100 + Val(xbig.Text)) / 100
'            nRem = nPrice - Int(nPrice)
'            Select Case nRem
'                Case Is = 0
'                    nPrice = Int(nPrice)
'                Case Is > 0
'                    nPrice = Int(nPrice) + 1
'                Case Is < 0.5
'                    nPrice = Int(nPrice) + 0.5
'            End Select
'            .TextMatrix(4, Col) = Myvalue(nPrice, "#0.00")
'        End If
'    End If
End If
End With
End Sub
Private Sub MakeCopy()
cFilter = ""
xModelNo.text = ""
With vsScal
For nCol = 1 To vsScal.Cols - 1
    .TextMatrix(.Rows - 1, nCol) = ""
Next
End With

With vsColor
For nCol = 1 To vsScal.Cols - 1
    .TextMatrix(.Rows - 1, nCol) = ""
Next
xModelNo.text = ""
xModelFact.text = ""
xModelFact.Enabled = True
FixBackColor
FixBackColor2
End With
End Sub
Private Sub myreplaceGrd()
FixCode vsScal
FixCode vsColor
Dim aInsert(42, 1), nScal As Double
Dim sitem As String, aRet As Variant
Dim cBarCode As String
If xFact.BoundText = "001" Then
    cModel = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.text, 10)
    cModelNo = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.text, 10)
Else
    cModel = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.text, 10) & retCode(xCode.text, con)
    cModelNo = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.text, 10)
End If
aInsert(0, 0) = "Model"
aInsert(0, 1) = addstring(cModel)

aInsert(1, 0) = "ModelNO"
aInsert(1, 1) = addstring(cModelNo)

aInsert(2, 0) = "Desca"
aInsert(2, 1) = addstring(xDescA.text)

aInsert(3, 0) = "FACT"
aInsert(3, 1) = addstring(xFact.BoundText)

aInsert(4, 0) = "MOSM"
aInsert(4, 1) = addstring(xMosm.BoundText)

aInsert(5, 0) = "MODELFACT"
aInsert(5, 1) = addstring(RetZero(xModelFact.text, 10))

aInsert(6, 0) = "MODELFACT0"
aInsert(6, 1) = addstring(DelZero(xModelFact.text))

aInsert(7, 0) = "SUPP"
aInsert(7, 1) = addstring(xSupp.text)

aInsert(8, 0) = "[section]"
aInsert(8, 1) = addstring(XSECTION.BoundText)

aInsert(9, 0) = "OKAZ"
aInsert(9, 1) = Val(XOKAZ.text)

aInsert(10, 0) = "RATE"
aInsert(10, 1) = Val(xRate.text)

aInsert(11, 0) = "Isstop"
aInsert(11, 1) = Val(xisstop.Value)

aInsert(12, 0) = "Code"
aInsert(12, 1) = addstring(xCode.text)

aInsert(13, 0) = "[group]"
aInsert(13, 1) = addstring(xGroup.BoundText)

aInsert(14, 0) = "CHARGE"
aInsert(14, 1) = addstring(xCharge.BoundText)

aInsert(15, 0) = "MOSM2"
aInsert(15, 1) = addstring(XMOSM2.text)

aInsert(16, 0) = "OKAZ_2"
aInsert(16, 1) = Val(xOkaz_2.text)

aInsert(17, 0) = "ISNODEM"
aInsert(17, 1) = Val(xisnodem.Value)

aInsert(18, 0) = "ISNOitem"
aInsert(18, 1) = Val(XISNOITEM.Value)

aInsert(19, 0) = "[ModelSex]"
aInsert(19, 1) = Val(xmodelsex.BoundText)

aInsert(20, 0) = "[ModelAge]"
aInsert(20, 1) = Val(xmodelage.BoundText)

aInsert(21, 0) = "[grmodel]"
aInsert(21, 1) = addstring(XGRMODEL.text)

aInsert(22, 0) = "[grdesca]"
aInsert(22, 1) = addstring(xGrDesca.text)

aInsert(23, 0) = "ISONEST"
aInsert(23, 1) = Val(XISONEST.Value)

aInsert(24, 0) = "Outlet"
aInsert(24, 1) = Val(XOutlet.Value)

aInsert(25, 0) = "ISNOITEM2"
aInsert(25, 1) = Val(XISNOITEM2.Value)

aInsert(26, 0) = "DESCE"
aInsert(26, 1) = addstring(xdesce.text)

aInsert(42, 0) = "ISONLINE"
aInsert(42, 1) = IIf(XISONLINE.Value, 1, 0)

ISNEWMODEL = False
prog1.Min = 0
prog1.Max = vsScal.Cols - 2
prog1.Value = 0
For nScal = 1 To vsScal.Cols - 2
    prog1.Value = nScal
    If validCol(nScal) Then
        For nColor = 1 To vsColor.Cols - 2
            lNew = False
            If Val(vsColor.TextMatrix(0, nColor)) <> 0 And vsScal.TextMatrix(0, nScal) <> "" Then
                sitem = ItemFind_Model_S_C(XMODEL.text, Val(vsColor.TextMatrix(vsColor.Rows - 1, nColor)), Val(vsScal.TextMatrix(vsScal.Rows - 1, nScal)), con)
                If sitem = 0 Then
                    sitem = Val(GetDesca("SELECT MAX(ITEM) FROM FILE1_10", con)) + 1
                    If Val(sitem) < 100 Then sitem = 101
                    lNew = True
                End If
                
                aInsert(27, 0) = "item"
                aInsert(27, 1) = addvalue(sitem)
                
                aInsert(28, 0) = "Cost"
                aInsert(28, 1) = Val(vsScal.TextMatrix(1, nScal))
                
                aInsert(29, 0) = "Price"
                aInsert(29, 1) = Val(vsScal.TextMatrix(2, nScal))
                
                aInsert(30, 0) = "Price2"
                aInsert(30, 1) = Val(vsScal.TextMatrix(3, nScal))
                
                aInsert(31, 0) = "Cost2"
                aInsert(31, 1) = Val(vsScal.TextMatrix(4, nScal))
            
                aInsert(32, 0) = "Scal"
                aInsert(32, 1) = addstring(vsScal.TextMatrix(0, nScal))
            
                aInsert(33, 0) = "C_Scal"
                aInsert(33, 1) = nScal
                
                aInsert(34, 0) = "Color"
                aInsert(34, 1) = addstring(UCase(vsColor.TextMatrix(1, nColor)))
        
                aInsert(35, 0) = "COLOR_E"
                aInsert(35, 1) = addstring(vsColor.TextMatrix(2, nColor))
                
                aInsert(36, 0) = "C_Color"
                aInsert(36, 1) = nColor
        
                aInsert(37, 0) = "Code_Color"
                aInsert(37, 1) = Val(vsColor.TextMatrix(vsColor.Rows - 1, nColor))
        
                aInsert(38, 0) = "IsOkazItem"
                aInsert(38, 1) = Val(xIsOkazItem.Value)
        
                aInsert(39, 0) = "Code_Scal"
                aInsert(39, 1) = Val(vsScal.TextMatrix(vsScal.Rows - 1, nScal))
        
                aInsert(40, 0) = "PRICE_2"
                aInsert(40, 1) = Val(vsScal.TextMatrix(5, nScal))

                aInsert(41, 0) = "PRICE_P"
                aInsert(41, 1) = Val(vsScal.TextMatrix(8, nScal))

                If lNew Then
                    con.Execute CreateInsert(aInsert, "File1_10")
                    ISNEWMODEL = True
                Else
                    con.Execute CreateUpdate(aInsert, "File1_10", " where item = " & sitem)
                End If
            End If
        Next
    End If
Next nScal
con.Execute " update file1_10 set item = item where model = " & addstring(XMODEL.text)
con.Execute " DELETE FROM NEW_ITEM FROM  NEW_ITEM INNER JOIN FILE1_10 ON NEW_ITEM.ITEM = FILE1_10.ITEM WHERE    FILE1_10.MODEL = " & MyParn(XMODEL.text)
con.Execute " INSERT INTO NEW_ITEM (ITEM, BRANCH) SELECT FILE1_10.ITEM, branch.code FROM FILE1_10 CROSS JOIN branch WHERE BRANCH.IPSQL IS NOT NULL  AND FILE1_10.MODEL = " & MyParn(XMODEL.text)

con.Execute " DELETE FROM NEW_ITEM_FR FROM  NEW_ITEM_FR INNER JOIN FILE1_10 ON NEW_ITEM_FR.ITEM = FILE1_10.ITEM WHERE    FILE1_10.MODEL = " & MyParn(XMODEL.text)
con.Execute " INSERT INTO NEW_ITEM_FR (ITEM, BRANCH) SELECT FILE1_10.ITEM, BRANCH_FR.code FROM FILE1_10 CROSS JOIN BRANCH_FR WHERE BRANCH_FR.ISSTOP = 0 AND FILE1_10.MODEL = " & MyParn(XMODEL.text)

End Sub
Private Sub ADDFILE(pFile As String)
Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT TOP 1000 * FROM " & pFile
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
cCaption = "from " & loctable.RecordCount
con.BeginTrans
Do Until loctable.EOF
    i = i + 1
    Me.Caption = i & " " & cCaption
    cString = "UPDATE " & pFile & " SET TRIG = " & i & " FROM " & pFile
    cString = cString & turn(cString) & "ID = " & loctable!ID
    con.Execute cString
    loctable.MoveNext
Loop
con.CommitTrans
End Sub
Private Function CanSave() As Boolean
With vsScal
For i = 1 To .Cols - 2
    If (Not IsNumeric(.TextMatrix(1, i))) Or (Not IsNumeric(.TextMatrix(2, i))) Then
        Exit Function
    End If
Next
End With
CanSave = True
End Function
Private Sub CalcTotals()
Dim nRate As Double
With vsScal
xProfit.Caption = ""
If .Cols > 2 And .Rows > 2 Then
    If Val(.TextMatrix(1, 1)) <> 0 And Val(.TextMatrix(2, 1)) <> 0 Then
        nRate = Round((Val(.TextMatrix(2, 1)) - Val(.TextMatrix(1, 1))) / Val(.TextMatrix(1, 1)) * 100, 2)
        If nRate <> 0 Then
            xProfit.Caption = nRate & "%"
        End If
    End If
End If
End With
End Sub
Sub FillItemFact(cModel)
Dim ColorTable As New ADODB.Recordset
Dim ScalTable As New ADODB.Recordset
cStrColor = " SELECT FILE1_10.MODEL,  FILE1_10.COLOR, FILE1_10.C_COLOR , code_Color FROM FILE1_10 WHERE MODELFACT0 = " & MyParn(cModel) & " GROUP BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_COLOR , FILE1_10.DESCA , code_Color ORDER BY C_COLOR "
cStrScal = " SELECT FILE1_10.DESCA ,FILE1_10.SCAL, FILE1_10.C_SCAL, FILE1_10.PRICE , FILE1_10.PRICE2 , file1_10.GrModel , file1_10.GrDesca FROM FILE1_10 WHERE MODELFACT0 = " & MyParn(cModel) & " GROUP BY FILE1_10.DESCA ,FILE1_10.SCAL, FILE1_10.C_SCAL, FILE1_10.PRICE, FILE1_10.PRICE2 , file1_10.GrModel , file1_10.GrDesca ORDER BY C_SCAL "
ColorTable.Open cStrColor, confact, adOpenKeyset, adLockReadOnly, adCmdText
ScalTable.Open cStrScal, confact, adOpenKeyset, adLockReadOnly, adCmdText
If ScalTable.RecordCount > 0 Then
With vsScal
    .Cols = 1
    ScalTable.MoveFirst
    xDescA.text = ScalTable!desca & ""
    XGRMODEL.text = ScalTable!GRMODEL & ""
    xGrDesca.text = ScalTable!GrDesca & ""
    Do While True
        .Cols = .Cols + 1
        .TextMatrix(6, .Cols - 1) = .Cols - 1
        .TextMatrix(0, .Cols - 1) = ScalTable!SCAL
        .TextMatrix(1, .Cols - 1) = Format(ScalTable!price, "##0.00")
        .TextMatrix(2, .Cols - 1) = Format(ScalTable!PRICE2, "##0.00")
        .TextMatrix(4, .Cols - 1) = Format(ScalTable!price, "##0.00")
        ScalTable.MoveNext
        If ScalTable.EOF Then Exit Do
    Loop
    .Cols = .Cols + 1
    fixGrd
    End With

    With vsColor
    .Cols = 1
    ColorTable.MoveFirst
    Do While True
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = ColorTable!c_Color
        .TextMatrix(1, .Cols - 1) = ColorTable!CODE_COLOR & ""
        ColorTable.MoveNext
        If ColorTable.EOF Then Exit Do
    Loop
    .Cols = .Cols + 1
    Fixgrd2
    End With
End If
End Sub
Sub FillItemFact2(cModel)
Dim ColorTable As New ADODB.Recordset
Dim ScalTable As New ADODB.Recordset
cStrColor = " SELECT FILE1_10.MODEL,  FILE1_10.COLOR, FILE1_10.C_COLOR FROM FILE1_10 WHERE MODEL = " & MyParn(cModel) & " GROUP BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_COLOR , FILE1_10.DESCA ORDER BY C_COLOR "
cStrScal = " SELECT FILE1_10.DESCA ,FILE1_10.SCAL, FILE1_10.C_SCAL, FILE1_10.PRICE , FILE1_10.PRICE2 FROM FILE1_10 WHERE MODEL = " & MyParn(cModel) & " GROUP BY FILE1_10.DESCA ,FILE1_10.SCAL, FILE1_10.C_SCAL, FILE1_10.PRICE, FILE1_10.PRICE2 ORDER BY C_SCAL "
ColorTable.Open cStrColor, confact2, adOpenKeyset, adLockReadOnly, adCmdText
ScalTable.Open cStrScal, confact2, adOpenKeyset, adLockReadOnly, adCmdText
If ScalTable.RecordCount > 0 Then

'    .TextMatrix(0, 0) = "«·„ﬁ«”"
'    .TextMatrix(1, 0) = "”⁄— «·„’‰⁄"
'    .TextMatrix(2, 0) = "”⁄— «·„” Â·ﬂ"
'    .TextMatrix(3, 0) = "”⁄— √Êﬂ«“ÌÊ‰"
'    .TextMatrix(4, 0) = "”⁄— Ã„·…"
'    .TextMatrix(5, 0) = "ﬁ”„"
'    .TextMatrix(6, 0) = ""

With vsScal
    .Cols = 1
    ScalTable.MoveFirst
    xDescA.text = ScalTable!desca
    Do While True
        .Cols = .Cols + 1
        .TextMatrix(6, .Cols - 1) = .Cols - 1
        .TextMatrix(0, .Cols - 1) = ScalTable!SCAL
        .TextMatrix(1, .Cols - 1) = Format(ScalTable!price, "##0.00")
        .TextMatrix(2, .Cols - 1) = Format(ScalTable!PRICE2, "##0.00")
        .TextMatrix(4, .Cols - 1) = Format(ScalTable!price, "##0.00")
        ScalTable.MoveNext
        If ScalTable.EOF Then Exit Do
    Loop
    .Cols = .Cols + 1
    fixGrd
    End With
End If
End Sub
Sub UpdatePrice()
On Error Resume Next
    con.Execute " DELETE FROM NEWPRICEITEM          WHERE   MODEL =     " & MyParn(XMODEL.text)
    con.Execute " INSERT INTO NEWPRICEITEM         (MODEL ) VALUES ( " & addstring(XMODEL.text) & " ) "

Err.Clear
End Sub
Private Sub AddFromImageFile(sPath As String)
If Not fs.FileExists(sPath) Then
    Exit Sub
End If
ImgX1.Images.Clear
ImgX1.Import.FromFile sPath
ImgX1.Refresh
AddToDataSql
'Handlecontrols
End Sub

Private Sub AddFromImageFile2(sPath As String)
AddToDataSql2
'Handlecontrols
End Sub
Private Function AddToDataSql() As Boolean
Dim bytes() As Byte, nSize As Long, loctable As New ADODB.Recordset
ImgX1.Export.ToMemoryFile bytes, ixmfJPG
con.Execute " DELETE FROM PICT WHERE MODELNO = " & MyParn(xModelNo.text)
Dim cmdUpdate As New ADODB.command
Set cmdUpdate.ActiveConnection = con
cmdUpdate.CommandType = adCmdStoredProc
cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("pict", adLongVarBinary, adParamInput, UBound(bytes) + 1, bytes)
cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("size", adInteger, adParamInput, 8, UBound(bytes) + 1)
cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("modelno", adVarWChar, adParamInput, 16, xModelNo.text)
cmdUpdate.CommandText = "ADD_IMAGE"
cmdUpdate.Execute
End Function
Private Function AddToDataSql2() As Boolean
Dim bytes() As Byte, nSize As Long, loctable As New ADODB.Recordset
ImgX1.Export.ToMemoryFile bytes, ixmfJPG
con.Execute " DELETE FROM PICT WHERE MODELNO = " & MyParn(xModelNo.text)
Dim cmdUpdate As New ADODB.command
Set cmdUpdate.ActiveConnection = con
cmdUpdate.CommandType = adCmdStoredProc
cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("pict", adLongVarBinary, adParamInput, UBound(bytes) + 1, bytes)
cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("size", adInteger, adParamInput, 8, UBound(bytes) + 1)
cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("modelno", adVarWChar, adParamInput, 16, xModelNo.text)
cmdUpdate.CommandText = "ADD_IMAGE"
cmdUpdate.Execute

conPict.Execute " DELETE FROM PICT WHERE MODELNO = " & MyParn(xModelNo.text)
End Function
Private Sub myLoadPicture(pModel As String)
    On Error Resume Next
    If Not lServerPict Then Exit Sub
    ImgX1.Images.Clear
    If pModel = "" Then Exit Sub
    If conPict Is Nothing Then Exit Sub
    Dim loctable As Recordset
    Dim bytes() As Byte, nSize As Long
    ImgX1.Images.Clear
    Set loctable = myRecordSet("select PICT,size from PICT where MODELNO = " & MyParn(pModel), conPict)
    If loctable.EOF Then Exit Sub
    nSize = CLng(loctable("size").Value)
    bytes = loctable("PICT").GetChunk(nSize)
    ImgX1.Import.FromMemoryFile bytes, ixmfJPG
    Err.Clear
End Sub
Sub FixBarCode2_Fact(pModel)
        Dim TableFact As New ADODB.Recordset
        Dim cScal As String, cModelFact As String
        TableFact.Open " SELECT * FROM FILE1_10 WHERE MODEL = " & MyParn(pModel) & " ORDER BY ITEM ", con, adOpenStatic, adCmdText
        If TableFact.RecordCount = 0 Then Exit Sub
        With TableFact
        Do While Not .EOF
            Me.Caption = !Item
            cModelFact = UCase(DelZero(!modelfact))
            cScal = !SCAL
            Dim loctable As ADODB.Recordset
            Set loctable = ItemFind_Fact(cModelFact, cScal, confact)
            If Not (loctable.EOF And loctable.BOF) Then
                con.Execute " UPDATE FILE1_10 SET BARCODE =  " & addstring(loctable!Item) & " , BARCODE2 = " & addstring(loctable!Item_BARCODE) & " , BARCODE13 = " & addstring(loctable!BARCODE13) & " , BARCODE_GS1 = " & addstring(loctable!BARCODE_GS1) & " WHERE ITEM = " & !Item
            End If
            .MoveNext
        Loop
        End With
End Sub
Sub FixBarCode2_Fact3(pModel)
        Dim TableFact As New ADODB.Recordset
        Dim cScal As String, cModelFact As String
'        TableFact.Open " SELECT * FROM FILE1_10 WHERE MODEL = " & MyParn(pModel) & " AND BARCODE2 IS NULL AND BARCODE13 IS NULL  AND BARCODE_GS1 IS NULL ORDER BY ITEM  ORDER BY ITEM ", con, adOpenStatic, adCmdText
        TableFact.Open " SELECT * FROM FILE1_10 WHERE MODEL = " & MyParn(pModel) & " ORDER BY ITEM ", con, adOpenStatic, adCmdText
        If TableFact.RecordCount = 0 Then Exit Sub
        With TableFact
        Do While Not .EOF
            Me.Caption = !Item
            cModelFact = UCase(DelZero(!modelfact))
            cScal = !SCAL
            Dim loctable As ADODB.Recordset
            Set loctable = ItemFind_Fact(cModelFact, cScal, confact3)
            If Not (loctable.EOF And loctable.BOF) Then
                con.Execute " UPDATE FILE1_10 SET BARCODE =  " & addstring(loctable!Item) & " , BARCODE2 = " & addstring(loctable!Item_BARCODE) & " , BARCODE13 = " & addstring(loctable!BARCODE13) & " , BARCODE_GS1 = " & addstring(loctable!BARCODE_GS1) & " WHERE ITEM = " & !Item
            End If
            .MoveNext
        Loop
        End With
End Sub
Private Sub FIXGRMODEL_Click()
    Dim i As Double
    Dim GrModelFact As New ADODB.Recordset, cModel As String, cMosmFact As String
    cMosmFact = InputBox("„Ê”„ «·„’‰⁄", , cPMosm)
    GrModelFact.Open " SELECT MODEL , GRMODEL , GRDESCA , MIN(BARCODE2) AS BARCODE2  FROM FILE1_10 WHERE GRMODEL IS NOT NULL  and mosm = " & MyParn(cMosmFact) & "  GROUP BY MODEL , GRMODEL , GRDESCA ", confact, adOpenStatic, adCmdText
    If GrModelFact.RecordCount = 0 Then Exit Sub
    prog1.Min = 0
    prog1.Max = GrModelFact.RecordCount
    prog1.Value = 0
    i = 0
    GrModelFact.MoveFirst
    Do While Not GrModelFact.EOF
        i = i + 1
        prog1.Value = i
        cModel = GetDesca("SELECT MODEL FROM FILE1_10 WHERE BARCODE2 = " & MyParn(GrModelFact!BARCODE2), con)
        con.Execute " UPDATE FILE1_10 SET GRMODEL = " & addstring(GrModelFact!GRMODEL) & " , GRDESCA = " & addstring(GrModelFact!GrDesca) & " WHERE MODEL = " & MyParn(cModel)
        GrModelFact.MoveNext
    Loop
    Inform " „  ⁄œÌ· " & i & " „ÊœÌ· "
End Sub
Sub AddModelToShop(pModel, pMosm, pFact)
    Dim ItemFactTable As New ADODB.Recordset
    ItemFactTable.Open "SELECT * FROM FILE1_10 WHERE MODEL = " & MyParn(pModel) & " ORDER BY SCAL ", condef, adOpenStatic, adLockReadOnly, adCmdText
    With ItemFactTable
        .MoveFirst
        Do While Not .EOF
            .MoveNext
        Loop
    End With
End Sub

Private Sub ChargeLook()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select FILE8_51.code ,FILE8_51.DescA ,  file8_52.desca From FILE8_51 left join file8_52 on file8_51.maingroup = file8_52.code "
Generalarray(2) = "Order by FILE8_51.code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "„ÃÊ„⁄…"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub


Sub FillItemFact3(cModel)
Dim ColorTable As New ADODB.Recordset
Dim ScalTable As New ADODB.Recordset
cStrColor = " SELECT FILE1_10.MODEL,  FILE1_10.COLOR, FILE1_10.C_COLOR , code_Color FROM FILE1_10 WHERE MODELFACT0 = " & MyParn(cModel) & " GROUP BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_COLOR , FILE1_10.DESCA , code_Color ORDER BY C_COLOR "
cStrScal = " SELECT FILE1_10.DESCA ,FILE1_10.SCAL, FILE1_10.C_SCAL, FILE1_10.PRICE , FILE1_10.PRICE2 , file1_10.GrModel , file1_10.GrDesca FROM FILE1_10 WHERE MODELFACT0 = " & MyParn(cModel) & " GROUP BY FILE1_10.DESCA ,FILE1_10.SCAL, FILE1_10.C_SCAL, FILE1_10.PRICE, FILE1_10.PRICE2 , file1_10.GrModel , file1_10.GrDesca ORDER BY C_SCAL "
ColorTable.Open cStrColor, confact, adOpenKeyset, adLockReadOnly, adCmdText
ScalTable.Open cStrScal, confact, adOpenKeyset, adLockReadOnly, adCmdText
If ScalTable.RecordCount > 0 Then
With vsScal
    .Cols = 1
    ScalTable.MoveFirst
    xDescA.text = ScalTable!desca & ""
    XGRMODEL.text = ScalTable!GRMODEL & ""
    xGrDesca.text = ScalTable!GrDesca & ""
    Do While True
        .Cols = .Cols + 1
        .TextMatrix(6, .Cols - 1) = .Cols - 1
        .TextMatrix(0, .Cols - 1) = ScalTable!SCAL
        .TextMatrix(1, .Cols - 1) = Format(ScalTable!price, "##0.00")
        .TextMatrix(2, .Cols - 1) = Format(ScalTable!PRICE2, "##0.00")
        .TextMatrix(4, .Cols - 1) = Format(ScalTable!price, "##0.00")
        ScalTable.MoveNext
        If ScalTable.EOF Then Exit Do
    Loop
    .Cols = .Cols + 1
    fixGrd
    End With

    With vsColor
    .Cols = 1
    ColorTable.MoveFirst
    Do While True
        .Cols = .Cols + 1
        .TextMatrix(0, .Cols - 1) = ColorTable!c_Color
        .TextMatrix(1, .Cols - 1) = ColorTable!CODE_COLOR & ""
        ColorTable.MoveNext
        If ColorTable.EOF Then Exit Do
    Loop
    .Cols = .Cols + 1
    Fixgrd2
    End With
End If
End Sub
Sub BACK_FixBarCode2_Fact3(pModel)
        Dim TableFact As New ADODB.Recordset
        Dim cScal As String, cModelFact As String
        TableFact.Open " SELECT * FROM FILE1_10 WHERE MODEL = " & MyParn(pModel) & " ORDER BY ITEM ", con, adOpenStatic, adCmdText
        If TableFact.RecordCount = 0 Then Exit Sub
        TableFact.MoveLast
        TableFact.MoveFirst
        If TableFact.RecordCount > 0 Then
            If GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODELFACT0 = " & MyParn(TableFact!modelfact0), confact3) = "" Then
                cModelFact = "00-" & TableFact!modelfact0
            End If
        End If
        With TableFact
        Do While Not .EOF
            Me.Caption = !Item
            If cModelFact = "" Then cModelFact = UCase(DelZero(!modelfact))
            cScal = !SCAL
            cBarCode = GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODELFACT0 = " & MyParn(cModelFact) & " AND SCAL = " & MyParn(cScal), confact3)
            If cBarCode = "" Then
                If cScal = "---" Then cScal = "-"
                cBarCode = GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODELFACT0 = " & MyParn(cModelFact) & " AND SCAL = " & MyParn(cScal), confact3)
                If cBarCode = "" Then cBarCode = GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODELFACT0 = " & MyParn(cModelFact) & " and SCAL = " & MyParn(cScal), confact3)
            End If
            If cBarCode <> "" Then
                aRet = aGetDesca("SELECT ITEM_BARCODE  ,BARCODE13 , BARCODE_GS1 FROM FILE1_10 WHERE ITEM = " & MyParn(cBarCode), confact3)
                If UBound(aRet) > 0 Then
                If aRet(1) <> "" Then
                    cBarCodE2 = aRet(1)
                    If cBarCodE2 <> "" Then
                        con.Execute " UPDATE FILE1_10 SET BARCODE = " & addstring(cBarCode) & " , BARCODE2 = " & addstring(cBarCodE2) & "  , BARCODE13 = " & addstring(aRet(2)) & ", BARCODE_GS1 = " & addstring(aRet(3)) & " WHERE ITEM = " & Val(!Item)
                        confact3.Execute " UPDATE FILE1_10 SET BARCODE_SHOP = " & Val(!Item) & " WHERE ITEM = " & addstring(cBarCode)
                        If aRet(2) <> "" And Not IsNull(aRet(2)) Then
                            con.Execute " DELETE FROM NEW_ITEM WHERE ITEM = " & !Item
                            For nRow = 2 To nCountBranch
                                sStore = RetZero(nRow, 2)
                                con.Execute "  INSERT INTO NEW_ITEM (ITEM, BRANCH) VALUES        ( " & addvalue(!Item) & " , " & addstring(sStore) & " )  "
                            Next nRow
                            con.Execute " DELETE FROM NEW_ITEM_FR WHERE ITEM = " & !Item
                            For nRow = 61 To nCountBranch_fr
                                sStore = RetZero(nRow, 2)
                                con.Execute "  INSERT INTO NEW_ITEM_FR (ITEM, BRANCH) VALUES        ( " & addvalue(!Item) & " , " & addstring(sStore) & " )  "
                            Next nRow
                        End If
                    End If
                End If
            End If
            Else
                Me.Caption = !Item
            End If
            .MoveNext
        Loop
        End With
End Sub
Private Sub myreplace_ItemFct(pModel, pMosm, pFact)
Dim pItemTable As New ADODB.Recordset
pItemTable.Open "select * from file1_10 where model = " & MyParn(pModel) & " order by item ", confact, adOpenStatic, adLockReadOnly, adCmdText
With pItemTable
Dim aInsert(27, 1)
Dim sitem As String, aRet As Variant
Dim cBarCode As String
cModel = pFact & pMosm & RetZero(pModel, 10)
cModelNo = pFact & pMosm & RetZero(pModel, 10)
If GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(cModel), con) <> "" Then Exit Sub
Do While Not pItemTable.EOF
    sitem = Val(GetDesca("SELECT MAX(ITEM) FROM FILE1_10", con)) + 1
    If Val(sitem) < 100 Then sitem = 101

    aInsert(0, 0) = "Model"
    aInsert(0, 1) = addstring(cModel)

    aInsert(1, 0) = "ModelNO"
    aInsert(1, 1) = addstring(cModelNo)

    aInsert(2, 0) = "Desca"
    aInsert(2, 1) = addstring(!desca)

    aInsert(3, 0) = "FACT"
    aInsert(3, 1) = addstring(pFact)

    aInsert(4, 0) = "MOSM"
    aInsert(4, 1) = addstring(pMosm)

    aInsert(5, 0) = "MODELFACT"
    aInsert(5, 1) = addstring(RetZero(pModel, 10))

    aInsert(6, 0) = "MODELFACT0"
    aInsert(6, 1) = addstring(DelZero(pModel))

    aInsert(7, 0) = "SUPP"
    aInsert(7, 1) = addstring("")

    aInsert(8, 0) = "[section]"
    aInsert(8, 1) = addvalue(2)

    aInsert(9, 0) = "Code"
    aInsert(9, 1) = addstring(pFact)

    aInsert(10, 0) = "[group]"
    aInsert(10, 1) = addstring("01")


    aInsert(11, 0) = "[grmodel]"
    aInsert(11, 1) = addstring(XGRMODEL.text)

    aInsert(12, 0) = "[grdesca]"
    aInsert(12, 1) = addstring(xGrDesca.text)

    aInsert(13, 0) = "item"
    aInsert(13, 1) = addvalue(sitem)

    aInsert(14, 0) = "Cost"
    If !ISDOZ Then
        aInsert(14, 1) = Val(!price / 12 & "")
    Else
        aInsert(14, 1) = Val(!price & "")
    End If
    
    aInsert(15, 0) = "Price"
    aInsert(15, 1) = Val(!PRICE2 & "")

    aInsert(16, 0) = "Cost2"
    If !ISDOZ Then
        aInsert(16, 1) = Val(!price / 12 & "")
    Else
        aInsert(16, 1) = Val(!price & "")
    End If

    aInsert(17, 0) = "Scal"
    aInsert(17, 1) = addstring(!SCAL)

    aInsert(18, 0) = "C_Scal"
    aInsert(18, 1) = addvalue(!C_SCAL)

    aInsert(19, 0) = "CODE_Scal"
    aInsert(19, 1) = addvalue(!C_SCAL)

    aInsert(20, 0) = "Color"
    aInsert(20, 1) = addstring(!color)

    aInsert(21, 0) = "C_Color"
    aInsert(21, 1) = 1

    aInsert(22, 0) = "Code_Color"
    aInsert(22, 1) = 1

    aInsert(23, 0) = "barcode"
    aInsert(23, 1) = addstring(!Item)
    
    aInsert(24, 0) = "barcode2"
    aInsert(24, 1) = addstring(!BARCODE2)
    
    aInsert(25, 0) = "barcode13"
    aInsert(25, 1) = addstring(!BARCODE13)
    
    aInsert(26, 0) = "barcode_gs1"
    aInsert(26, 1) = addstring(!BARCODE_GS1)
    
    aInsert(27, 0) = "Color_E"
    aInsert(27, 1) = addstring(!color)
    
    con.Execute CreateInsert(aInsert, "File1_10"), nRec
    If nRec = 1 Then Me.Caption = " Add " & pModel & " " & !SCAL
    
    .MoveNext
Loop
End With
End Sub
Private Sub ScanImage()
On Error GoTo myError
Set twain = New ImgXTwain
twain.OpenTwain Me.hwnd
If twain.QuerySupport(ixtcResolution) Then
     twain.Resolution = 150
End If
twain.Acquire False, Me.hwnd
Exit Sub
myError:
MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub
Private Sub Twain_ImageAcquired(Image As ImgX_Image)
ImgX1.Images.Replace Image, , False
ImgX1.TIFCompression = ixtcJPEGCompression

myReplaceImage IIf(chkEdit.Value = 0, "", xId.Caption)
End Sub
Private Sub Twain_CanCloseTwain()
    ' This event is called after you call Acquire.
    ' It let's you know when it's safe to call CloseTwain.
    twain.CloseTwain
    ' Steps menu
End Sub
Private Function myReplaceImage(sid As String, Optional pFile As String = "") As Boolean
Dim mStream As ADODB.Stream
Set mStream = New ADODB.Stream

mStream.Type = adTypeBinary ' Set the stream type to binary data
mStream.Open

Dim command As ADODB.command
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "MODEL", XMODEL.text)
If sid <> "" Then
    aPrm = AddFlag(aPrm, "ID", sid)
End If

If pFile <> "" Then
    mStream.LoadFromFile pFile
    aPrm = AddFlag(aPrm, "IMAGE", mStream.Read)
Else
    Dim bytes() As Byte
    ImgX1.Export.ToMemoryFile bytes, ixmfJPG
    aPrm = AddFlag(aPrm, "IMAGE", bytes)
End If

Set cm = cmd("sp_add_model_image", con_image, adStoredProc, aPrm)
cm.Execute
ImgX1.Images.Clear

If sid <> "" Then
    If Not openTablePhotos(tbMode.tbFind, sid) Then
        If Not openTablePhotos Then myDefinePhotos
    End If
Else
    If Not openTablePhotos Then myDefinePhotos
End If
End Function
Private Function getModelImage(con As ADODB.Connection, ID As String) As Byte()
Dim cmd As ADODB.command
Dim paramOutput As ADODB.Parameter
'Dim imageData() As Byte

Set cmd = New ADODB.command
With cmd
    .ActiveConnection = con_image ' An open ADODB.Connection
    .CommandText = "[dbo].[sp_GetImageModel]"
    .CommandType = adCmdStoredProc
    
    ' Input parameter
    .Parameters.Append .CreateParameter("@ID", adInteger, adParamInput, , ID)
    
    ' Output parameter: Use adVarBinary or adLongVarBinary
    ' Set size to -1 for VARBINARY(MAX)
    Set paramOutput = .CreateParameter("@Image", adVarBinary, adParamOutput, 8000000)
    .Parameters.Append paramOutput
    
    ' Execute the procedure
    .Execute
End With

' Check if the value is not null, then assign to Byte Array
If Not IsNull(cmd.Parameters("@image").Value) Then
    getModelImage = cmd.Parameters("@image").Value
    ' You now have the binary image data in the 'imageData' array
End If
End Function
'Private Sub opencardPhotos(Optional bRefresh As Boolean = False)
'Set rsPhotos = New adodb.Recordset
'Dim cString As String
'
'cString = "select id,image from MODEL_IMAGES  where model = " & MyParn(xModel.text) & " order by id desc"
'If rsPhotos.Source <> cString Or bRefresh Then
'    Set rsPhotos = cmd(cString, con_image).Execute
'End If
'If rsPhotos.EOF Then
'    myDefinePhotos
'Else
'    myLoadPhotos
'End If
'End Sub
Private Function openTablePhotos(Optional pMode As Integer = tbMode.tblast, Optional pId As String = "") As Boolean
Dim cString As String
Dim cWhere As String
Dim cOrder As String
Dim nRecords As Long, nRecord As Long

If (pMode = tbMode.tbFind Or pMode = tbMode.tbNext Or pMode = tbMode.tbPrevious) And Trim(pId) = "" Then
    Exit Function
End If

Me.MousePointer = vbHourglass

On Error GoTo myError
Set rsPhotos = New ADODB.Recordset
cString = "SELECT TOP 1 MODEL_IMAGES.*" & _
          " FROM MODEL_IMAGES "
If pMode = tbMode.tbFirst Then
    'cOrder = "Order by MODEL_IMAGES.DOC_NO_FLAG"
    cOrder = "Order by MODEL_IMAGES.ID"
ElseIf pMode = tbMode.tblast Then
    'cOrder = "Order by MODEL_IMAGES.DOC_NO_FLAG DESC"
    cOrder = "Order by MODEL_IMAGES.ID DESC"
ElseIf pMode = tbMode.tbFind Then
    cWhere = "MODEL_IMAGES.ID = " & pId
ElseIf pMode = tbMode.tbPrevious Then
    'cWhere = "MODEL_IMAGES.DOC_NO_FLAG < " & MyParn(pDoc_no)
    'cOrder = "Order by MODEL_IMAGES.DOC_NO_FLAG DESC"
    
    cWhere = "MODEL_IMAGES.ID < " & pId
    cOrder = "Order by MODEL_IMAGES.ID DESC"
ElseIf pMode = tbMode.tbNext Then
    'cWhere = "MODEL_IMAGES.DOC_NO_FLAG > " & MyParn(pDoc_no)
    cWhere = "MODEL_IMAGES.ID > " & pId
    cOrder = "Order by MODEL_IMAGES.ID"
End If

If cWhere <> "" Then
    cString = cString & " WHERE " & cWhere
End If

cFilterPhotos = "MODEL = " & MyParn(XMODEL.text)
If cFilterPhotos <> "" Then
    cString = cString & IIf(cWhere = "", " WHERE ", " AND ") & cFilterPhotos
End If

cString = cString & " " & cOrder

Set rsPhotos = cmd(cString, con_image).Execute

If (Not rsPhotos.EOF) Then
    myLoadPhotos
    openTablePhotos = True
End If

Me.MousePointer = vbNormal
Exit Function
myError:
Me.MousePointer = vbNormal
MsgBox Err.Description
Err.Clear
End Function
Private Sub myDefinePhotos()
xId.Caption = ""
ImgX1.Images.Clear
HandleControlsPhotos DefineMode
lblRecords.Caption = "·«  ÊÃœ ’Ê—"
End Sub
Private Sub myLoadPhotos()
On Error GoTo myError
ImgX1.Import.FromMemoryFile (rsPhotos!Image)
xId.Caption = rsPhotos!ID
cmdColor.Tag = rsPhotos!c_Color & ""
If Not IsNull(rsPhotos!c_Color) Then
    cmdColor.Caption = myField("select color from file1_10 where " & _
                               " MODEL = " & MyParn(rsPhotos!MODEL) & _
                               " AND C_COLOR = " & rsPhotos!c_Color, con)
Else
    cmdColor.Caption = cmdColor.TagVariant
End If
HandleControlsPhotos LoadMode
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub HandleControlsPhotos(nMode)
xId.Tag = nMode
Dim nRecord As Long, nRecords As Long
retRecords xId.Caption, nRecords, nRecord
If nMode = LoadMode Then
    lblRecords.Caption = ArbString("’Ê—… " & nRecord & " „‰ " & nRecords)
Else
    lblRecords.Caption = "«÷«›… ’Ê—… " & (nRecords + 1)
End If

cmdPrevious.Enabled = (nMode = LoadMode) And nRecord > 1
cmdNext.Enabled = (nMode = LoadMode) And nRecord < nRecords
cmdLast.Enabled = (nMode = LoadMode) And nRecord < nRecords And nRecords > 2
cmdFirst.Enabled = (nMode = LoadMode) And nRecord > 1 And nRecords > 2
cmdColor.Enabled = nMode = LoadMode
cmdDelPhoto.Enabled = nMode = LoadMode
chkEdit.Enabled = nMode = LoadMode
cmdScan.Enabled = XMODEL.text <> ""
cmdRotate.Enabled = XMODEL.text <> ""
End Sub
Private Function retRecords(pCode, ByRef nRecords As Long, ByRef nRecord As Long) As Variant
Dim cString As String, loctable As New ADODB.Recordset
If pCode <> "" Then
   cString = "SELECT Count(*) AS records,COUNT(CASE WHEN ID <= " & pCode & " THEN 1 END) AS record"
Else
    cString = "SELECT Count(*) AS records,0 as record"
End If

cString = cString & " FROM MODEL_IMAGES"
cString = cString & " WHERE MODEL = " & MyParn(XMODEL.text)

Set loctable = mycmd(cString, con_image)
If Not loctable.EOF Then
    nRecords = loctable!RECORDS
    nRecord = Val(loctable!Record & "")
End If
End Function

