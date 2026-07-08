VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form items 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„ÊœÌ·« "
   ClientHeight    =   9390
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   15000
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
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
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   9390
   ScaleWidth      =   15000
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   825
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   45
      Visible         =   0   'False
      Width           =   1410
      Begin VB.CommandButton cmdCopy 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   67
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
      Height          =   2400
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   4905
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
         Rows            =   7
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
   Begin VB.CommandButton Command1 
      Caption         =   "÷»ÿ ‰ﬁ· „ÊœÌ·« "
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   14805
      RightToLeft     =   -1  'True
      TabIndex        =   64
      Top             =   -360
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Frame FramProg 
      Height          =   780
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   4095
      Width           =   8340
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   555
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   -45
         Visible         =   0   'False
         Width           =   3345
      End
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   510
         Left            =   90
         TabIndex        =   62
         Top             =   180
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   900
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
      Height          =   1995
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   58
      Top             =   7290
      Width           =   14865
      Begin VSFlex7Ctl.VSFlexGrid vsColor 
         Height          =   1560
         Left            =   135
         TabIndex        =   63
         Top             =   315
         Width           =   14640
         _cx             =   25823
         _cy             =   2752
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
      Height          =   780
      Left            =   8415
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   4095
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
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   180
         Visible         =   0   'False
         Width           =   1635
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   510
         Left            =   90
         TabIndex        =   56
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
         Picture         =   "Items.frx":3238
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
      TabIndex        =   51
      Top             =   0
      Width           =   4155
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   90
         Picture         =   "Items.frx":575A
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   270
         Width           =   1185
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
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1905
      End
      Begin VB.Label Label7 
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
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame5 
      Height          =   825
      Left            =   5805
      RightToLeft     =   -1  'True
      TabIndex        =   45
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
         Picture         =   "Items.frx":7F2D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items.frx":A290
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items.frx":C809
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   1275
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items.frx":EC75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   47
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4830
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Items.frx":1150F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   225
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      Height          =   825
      Left            =   11925
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   4095
      Width           =   2985
      Begin Threed.SSCommand calc 
         Height          =   540
         Index           =   1
         Left            =   795
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
         Caption         =   "1/10"
      End
      Begin Threed.SSCommand calc 
         Height          =   540
         Index           =   0
         Left            =   90
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
         Caption         =   "1/12"
      End
      Begin Threed.SSCommand calc 
         Height          =   540
         Index           =   2
         Left            =   1500
         TabIndex        =   42
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
         TabIndex        =   43
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
   Begin VB.Frame Frame3 
      Caption         =   "’Ê—… «·„ÊœÌ· "
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4065
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   45
      Width           =   4290
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
         Height          =   600
         Left            =   2820
         Picture         =   "Items.frx":13ABB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3375
         Width           =   1365
      End
      Begin VB.CommandButton cmdRotate 
         Caption         =   "Rotate"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1455
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   3375
         Width           =   1365
      End
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
         Height          =   600
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3375
         Width           =   1365
      End
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   3030
         Left            =   90
         TabIndex        =   37
         Top             =   270
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   5345
         BackColor       =   16777215
         BorderStyle     =   4
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         AutoZoomType    =   3
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrvb71"
         LicenseRegCode  =   "íß“ªß•≤º∂´≠“±®ππ∂´µßZQEH-AOZOOOZT-EFLF6gI"
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
         TabIndex        =   68
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
         TabIndex        =   69
         Top             =   1755
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Height          =   960
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   900
      Width           =   10545
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
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   180
         Width           =   1455
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
         Left            =   2835
         MaxLength       =   4
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1455
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
         TabIndex        =   54
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
         TabIndex        =   38
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
         Caption         =   "„ﬂ » Ã„·… :"
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   960
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
      Height          =   2220
      Left            =   4365
      TabIndex        =   21
      Top             =   1890
      Width           =   10545
      Begin VB.CommandButton cmdGroup 
         Caption         =   "..."
         Height          =   330
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   180
         Width           =   330
      End
      Begin VB.TextBox xbig 
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
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1260
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox xShowSales 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ „Œ ’— ›Ï «·„»Ì⁄« "
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1710
         Visible         =   0   'False
         Width           =   2355
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
         RightToLeft     =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   900
         Width           =   1815
      End
      Begin VB.TextBox xRedem 
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
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   540
         Width           =   1815
      End
      Begin VB.TextBox xDescA 
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
         Left            =   3960
         MaxLength       =   40
         TabIndex        =   7
         Top             =   540
         Width           =   4785
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
         Left            =   7155
         Locked          =   -1  'True
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   900
         Width           =   1590
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
         RightToLeft     =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1260
         Width           =   1590
      End
      Begin VB.TextBox xFixPrice 
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
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1620
         Visible         =   0   'False
         Width           =   1590
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   5940
         TabIndex        =   6
         Top             =   180
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·Ã„·… :"
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
         TabIndex        =   59
         Top             =   1305
         Visible         =   0   'False
         Width           =   1035
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
         TabIndex        =   32
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
         TabIndex        =   27
         Top             =   945
         Width           =   1605
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Õœ ≈⁄«œ… «·ÿ·» :"
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
         TabIndex        =   26
         Top             =   585
         Width           =   1305
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   900
         Width           =   3165
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
         TabIndex        =   23
         Top             =   1290
         Width           =   1245
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄— »Ì⁄ „ÊÕœ :"
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
         Top             =   1665
         Visible         =   0   'False
         Width           =   1215
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
   Begin VB.Frame Frame9 
      Height          =   780
      Left            =   810
      TabIndex        =   28
      Top             =   4095
      Visible         =   0   'False
      Width           =   6090
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
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   180
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
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
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
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   180
         Width           =   2040
      End
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
End
Attribute VB_Name = "ITEMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents twain As ImgXTwain
Attribute twain.VB_VarHelpID = -1
Dim con As New ADODB.Connection
Dim bAdded As Boolean
Public CardTable As ADODB.Recordset
Public bEdit As Boolean
Public nFlag As Integer
Public sFact As String, sMosm As String, sSupp As String, sModelFact As String, sCode As String, sModel As String, sSection  As String
Public bAddModel As Boolean
Dim oSearchItem As New Search3, sControl As String
Dim sOkaz As String, sRate As String, sBig As String, sRedem As String
Dim ColorTable As New ADODB.Recordset
Dim cFilter As String
Dim nColWidth As Long
Dim nRecordNumber As Integer
Const LoadMode = 1, DefineMode = 2
Sub mydefine()
xModel.Text = ""
xFact.BoundText = ""
xModelFact.Text = ""
xMosm.BoundText = ""
xGroup.BoundText = ""
xSupp.Text = ""
xCode.Text = ""
xCodeDesc.Caption = ""
xFixPrice.Text = ""
xRedem.Text = sRedem
xbig.Text = sBig
xOkaz.Text = sOkaz
xDesca.Text = ""
vsScal.Rows = 1
vsColor.Rows = 1
xItem.Text = ""
imgx1.Images.Clear
xShowSales.Value = 0
myDefineGrd
Handlecontrols DefineMode
End Sub
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    If cFilter <> "" Then
        cFilter = ""
        openCardTable
    End If
    xModel.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    Unload oSearchItem
    myUndo
End If
End Sub
Public Sub myload()
On Error GoTo myerror
xModel.Text = CardTable!Model
xModelFact.Text = CardTable!MODELFACT & ""
xModelNo.Text = CardTable!MODELno & ""
xSupp.Text = CardTable!Supp & ""
xMosm.BoundText = CardTable!Mosm & ""
xFact.BoundText = CardTable!Fact & ""
xOkaz.Text = Myvalue(CardTable!okaz & "")
xRate.Text = Myvalue(CardTable!Rate, "#0.00")
xRedem.Text = Myvalue(CardTable!Redem, "#0.00")
xGroup.BoundText = CardTable!Group & ""
xFixPrice.Text = Myvalue(CardTable!fixprice, "#0.00")
xCode.Text = CardTable!CODE & ""
xCodeDesc.Caption = GetDesca("select desca from file4_10 where code = " & MyParn(CardTable!CODE))
xDesca.Text = CardTable!DESCA & ""
'xShowSales.Value = IIf(CardTable!ShowSales, 1, 0)
'xModelNo.Text = xFact.Text & XMOSM.BoundText & DelZero(xModelFact.Text)
xShowSales.Value = IIf(CardTable!SHOWSALES, 1, 0)
'xSection.BoundText = CardTable!Section & ""
LoadPhoto
myloadgrd
myloadgrd2
'xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub BrtMinus_Click()
On Error GoTo myerror
If validPhoto(retFile(xModel.Text)) Then
    imgx1.Filters.Brightness (-5)
    imgx1.Export.ToFile retFile(xModel.Text), ixfsJPG
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub BrtPlus_Click()
On Error GoTo myerror
If validPhoto(retFile(xModel.Text)) Then
    imgx1.Filters.Brightness (5)
    imgx1.Export.ToFile retFile(xModel.Text), ixfsJPG
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Cmd_Undo_Click()
myUndo
End Sub

Private Sub calc_Click(index As Integer)
With vsScal
Dim nValue As Double
Select Case index
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

Private Sub cmdCopy_Click()
If MsgBox("⁄„· ‰”Œ… Â· «‰  „ √ﬂœ", vbOKCancel) = vbOK Then
    MakeCopy
End If
End Sub

Private Sub CmdDel_Click()
If GetDesca("SELECT FILE1_10.ITEM FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_11.ITEM = FILE1_10.ITEM WHERE FILE1_10.MODEL = " & MyParn(xModel.Text)) <> "" Then
    MsgBox "ÌÊÃœ Õ—ﬂ… ⁄·Ï «·„ÊœÌ·"
    Exit Sub
End If
If MsgBox("”Ì „ «·€«¡ «·„ÊœÌ· »«·ﬂ«„· ", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
    con.BeginTrans
    On Error GoTo myerror
    cString = " DELETE  FROM FILE1_10 WHERE MODEL = " & MyParn(xModel.Text)
    con.Execute cString
    cString = " DELETE  FROM FILE1_10H WHERE MODEL = " & MyParn(xModel.Text)
    con.Execute cString
    con.CommitTrans
End If
mydefine
openCardTable
MsgBox "  „ Õ–› «·„ÊœÌ· "
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub cmdDelPhoto_Click()
Dim fs As New FileSystemObject
On Error GoTo myerror
If MsgBox("Õ–› ’Ê—… «·„ÊœÌ· !! „ √ﬂœ ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
If validPhoto(retFile(xModelNo.Text)) Then fs.DeleteFile retFile(xModelNo.Text)
LoadPhoto
Handlecontrols LoadMode
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGroup_Click()
ReDim aLocal(6)
cCode = xGroup.BoundText
aLocal(0) = "FILE1_50"
aLocal(1) = "Code"
aLocal(2) = "Desca"
aLocal(3) = "ﬂÊœ „Ã„Ê⁄…"
aLocal(4) = "»Ì«‰ „Ã„Ê⁄« "
aLocal(5) = " ”ÃÌ· „Ã„Ê⁄« "
aLocal(6) = 2
FlagFrm2.bEdit = True
FlagFrm2.myPublic = aLocal
FlagFrm2.Show 1
data1.Refresh
xGroup.BoundText = cCode
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub

Private Sub CmdInform_Click()
ModelLookupAll Me, oSearchItem
End Sub
Private Sub cmdRotate_Click()
On Error GoTo myerror
If validPhoto(retFile(xModelNo.Text)) Then
    imgx1.Effects.Rotate 90
    imgx1.Export.ToFile retFile(xModelNo.Text), ixfsJPG
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If myreplace Then
    Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
    If sModel <> "" Then
        Unload Me
    Else
        openCardTable
        myUndo
    End If
End If
End Sub

Private Sub cmdScan_Click()
If xFact.Text <> "" And xMosm.BoundText <> "" And xFact.BoundText <> "" Then
    Dim sModelNo As String
    sModelNo = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.Text, 4)
    If Not xMosm.MatchedWithList Then MsgBox "«·„Ê”„ €Ì— ’ÕÌÕ"
    If Len(xMosm.BoundText) < 3 Then MsgBox "«·„Ê”„ €Ì— ’ÕÌÕ"
    MyCreateFolder App.Path & "\PICT" & turn(xMosm.BoundText, "\") & xMosm.BoundText
    ScanImage
End If
End Sub
Private Sub CmdUndo_Click()
If xModel.Tag = DefineMode And bAddModel Then
    myAddModel
Else
    openCardTable
    myUndo
End If
End Sub
Private Sub cmdFixModel_click()
Dim cMod As String, cMod0 As String, cModel As String, cModelNo  As String, cString As String, cCaption As String
Dim loctable As New ADODB.Recordset
cString = "Select * from file1_10"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
con.BeginTrans
On Error GoTo myerror
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

        cMod = Trim(!MODELFACT & "")
        cMod = RetZero(Replace(cMod, "'", "-"), 4)
        cMod0 = DelZero(cMod)
        cModelNo = !Fact & !Mosm & cMod
        cModel = !Fact & !Mosm & cMod & Trim(!Supp & "")
        
        aInsert(0, 0) = "ModelFact"
        aInsert(0, 1) = addstring(cMod)
        
        aInsert(1, 0) = "ModelFact0"
        aInsert(1, 1) = addstring(cMod0)
        
        aInsert(2, 0) = "MODEL"
        aInsert(2, 1) = addstring(cModel)
        
        aInsert(3, 0) = "MODELNO"
        aInsert(3, 1) = addstring(cModelNo)
        
        con.Execute CreateUpdate(aInsert, "FILE1_10", " where item = " & !Item, Array(-1))
        .MoveNext
    Loop
End With
con.CommitTrans
Set loctable = Nothing
Inform " „ ÷»ÿ »Ì«‰ «·„ÊœÌ·«  »‰Ã«Õ"
Me.FramProg.Visible = False
Me.Caption = cCaption
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
Me.Caption = cCaption
End Sub
Private Sub cmdFixModel2_click()
Dim cString As String, sModel As String, nCode_Scal As Long
Dim loctable As New ADODB.Recordset
loctable.Open "SELECT FILE1_10.* FROM FILE1_10 ORDER BY FILE1_10.MODEL,FILE1_10.C_SCAL", con, adOpenStatic, adLockReadOnly, adCmdText
On Error GoTo myerror
Me.FramProg.Visible = True
sModel = "Dummy123"
con.BeginTrans
On Error GoTo myerror
With loctable
    Dim nRecord As Long, nRecordCount As Long
    nRecordCount = loctable.RecordCount
    Do Until .EOF
        If sModel <> loctable!Model Then
            nCode_Scal = 1
            sModel = loctable!Model
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
myerror:
Err.Clear
MsgBox Err.Description
GoTo lastsub
End Sub

Private Sub Command1_Click()
Dim cString As String, loctable As New ADODB.Recordset
cString = "SELECT * FROM FILE1_10"
cString = cString & turn(cString) & "FILE1_10.MODEL = " & MyParn(xModel.Text)
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
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub Command2_Click()
Dim fs, f, f1, fc, s, photoDir As String, cFile As String, sMosm As String
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
        sMosm = Mid(cFile, 4, 3)
        cFile = retFile(cFile)
        MyCreateFolder App.Path & "\PICT" & turn(sMosm, "\") & sMosm
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


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 19 Then
    cmdSave_Click
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then
    KeyCode = 0
    SendKeys "{tAB}"
End If
End Sub
Private Sub Form_Load()
SetKbLayout Lang_EN
Dim ImgXTmp As ImgX
Set ImgXTmp = New ImgX
ImgXTmp.Import.FromMemoryFile LoadResData("BACKGROUND", "CUSTOM")
Set imgx1.BackgroundImage = ImgXTmp.image
Set ImgXTmp = Nothing

'ImgX1.AutoZoom = True
'ImgX1.AutoZoomType = 1
nColWidth = vsScal.Width / 16
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FACT ORDER BY DESCA"
Set xFact.RowSource = data1
xFact.ListField = "DESCA"
xFact.BoundColumn = "CODE"

'data5.ConnectionString = strCon
'data5.RecordSource = "SELECT * FROM SUPP ORDER BY DESCA "
'Set xFact.RowSource = data1
'xFact.ListField = "DESCA"
'xFact.BoundColumn = "CODE"

data2.ConnectionString = strCon
data2.RecordSource = "SELECT * FROM FILE1_50 ORDER BY DESCA"
Set xGroup.RowSource = data2
xGroup.ListField = "DESCA"
xGroup.BoundColumn = "CODE"

'data3.ConnectionString = strCon
'data3.RecordSource = "SELECT * FROM FILE1_10SC ORDER BY DESCA"
'Set xSection.RowSource = data3
'xSection.ListField = "DESCA"
'xSection.BoundColumn = "CODE"

data4.ConnectionString = strCon
data4.RecordSource = "SELECT * FROM MOSM ORDER BY DATE DESC "
Set xMosm.RowSource = data4
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"

myDefineGrd
'vsScal.RowHidden(vsScal.Rows - 1) = True

mydefine
If sModel <> "" Then
    If bAddModel Then myAddModel
Else
    openCardTable
    myUndo
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set ITEMS = Nothing
closeCon con
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
        con.BeginTrans
        cString = "DELETE  FROM FILE1_10 WHERE MODEL = " & MyParn(xModel.Text) & " AND c_color = " & MyParn(.TextMatrix(.Rows - 1, .Col))
        con.Execute cString, nAffect
        con.CommitTrans
        nCol = .Col
        myloadgrd2
        .Col = nCol
    End If
ElseIf KeyCode = 13 Then
'    CellPos2 KeyCode, Row, Col
End If
End With
Exit Sub
myerror:
con.RollbackTrans
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
    End If
Else
    CalcRate Col, Row
End If
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
If vsScal.Row < 0 Then vsScal.Row = 0
vsScal.Col = vsScal.Cols - 2
End Sub
Private Sub vsScal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CellPos KeyAscii, vsScal.Row, vsScal.Col
End Sub
Private Sub vsScal_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 13 Then CellPos KeyAscii, Row, Col
End Sub
Private Sub vsScal_KeyUp(KeyCode As Integer, Shift As Integer)
With vsScal
    If cmdSave.Enabled = False Then Exit Sub
    If KeyCode = 45 And .Col <> .Cols - 1 And validCol(.Col) Then
        InsertCol vsScal, .Col
        FixCol .Col
   ElseIf KeyCode = 46 And .Col <> .Cols - 1 And .Col <> 0 And .Row = 0 Then
        If Not IsNumeric(.TextMatrix(.Rows - 1, .Col)) Then
            DelCol vsScal, .Col
            FixCol .Col
        Else
            If MsgBox("Õ–› «·„ﬁ«” ø", vbYesNo + vbDefaultButton2) <> vbYes Then
                bAdd = False
                Exit Sub
            End If
            con.BeginTrans
            On Error GoTo myerror
            cString = "DELETE  FROM FILE1_10 WHERE MODEL = " & MyParn(xModel.Text) & " AND CODE_SCAL = " & .TextMatrix(.Rows - 1, .Col)
            con.Execute cString
            con.CommitTrans
            nCol = .Col
            myloadgrd
            .Col = nCol
        End If
    End If
End With
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long, Optional bCalc As Boolean = True)
With vsScal
If cmdSave.Enabled = False Then Exit Sub
If (Not validCol(Col)) Then Exit Sub
KeyCode = 0
If Col < vsScal.Cols - 1 Then vsScal.Col = Col + 1
If vsScal.Col = vsScal.Cols - 1 Then
    If Row = 0 Then CalcCol vsScal.Col, 0
    If validCol(.Cols - 1) Then
        .Cols = .Cols + 1
        CalcCol vsScal.Col, 1
        FixCol .Cols - 1
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
    If .Row = 1 And cmdSave.Enabled Then .Editable = flexEDKbdMouse Else .Editable = flexEDNone
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
End With
End Sub
Private Sub xCmd_Fix_Click()
If MsgBox("”Ì „  ⁄œÌ· ”⁄— «·√Êﬂ«“ÌÊ‰ ·ÌﬂÊ‰ ”⁄— „ÊÕœ »ﬁÌ„… " & Format(xFixPrice.Text, "##0.00"), vbOKCancel) = vbOK Then
    With vsScal
    For i = 1 To .Cols - 1
        .TextMatrix(4, i) = Val(xFixPrice.Text)
    Next i
    End With

    cStr1 = " update FILE1_10 SET file1_10.price2 = " & Val(xFixPrice.Text)
    If xFact.Text <> "" Then cStr1 = cStr1 & " and fact = " & MyParn(xFact.Text)
    If xModelFact.Text <> "" Then cStr1 = cStr1 & " and modelfact = " & MyParn(xModelFact.Text)
    If xMosm.BoundText <> "" Then cStr1 = cStr1 & " and mosm = " & MyParn(xMosm.BoundText)
    mydb.Execute cStr1
End If
myload
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    myload
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub
Private Sub myloadgrd()
Dim loctable As New ADODB.Recordset, i As Integer
loctable.CursorLocation = adUseClient
loctable.Open "select scal,Cost,Price ,PRICE2,Cost2,[SEC], code_scal from file1_10 where model = " & MyParn(xModel.Text) & " group by file1_10.c_scal,scal,Cost,Price ,Cost2,price2,FILE1_10.SEC,file1_10.code_Scal order by file1_10.c_scal ", con, adOpenForwardOnly, adLockReadOnly, adCmdText
With vsScal
.Cols = 1
Do Until loctable.EOF
    .Cols = .Cols + 1
    For i = 0 To 6
        If i <> 4 Then .TextMatrix(i, .Cols - 1) = loctable.Fields(i) & ""
    Next
    loctable.MoveNext
Loop
If cmdSave.Enabled Then .Cols = .Cols + 1
Fixgrd
loctable.Close
Set loctable = Nothing
End With
End Sub
Private Sub myloadgrd2()
Dim loctable As New ADODB.Recordset
loctable.CursorLocation = adUseServer
loctable.Open "select file1_10.C_Color,Color from file1_10 where model = " & MyParn(xModel.Text) & " group by file1_10.C_color,COLOR order by file1_10.C_color", con, adOpenStatic, adLockReadOnly, adCmdText
With vsColor
.Cols = 1
Do Until loctable.EOF
    .Cols = .Cols + 1
    .TextMatrix(0, .Cols - 1) = loctable!c_Color & ""
    .TextMatrix(1, .Cols - 1) = loctable!Color & ""
    loctable.MoveNext
Loop
If cmdSave.Enabled Then .Cols = .Cols + 1
Fixgrd2
End With
loctable.Close
Set loctable = Nothing
End Sub
Private Sub Fixgrd()
With vsScal
.ColWidth(0) = 1300
For i = 1 To .Cols - 1
    .ColWidth(i) = nColWidth
    .ColAlignment(i) = flexAlignRightCenter
Next
.RowHidden(4) = True
.RowHidden(.Rows - 1) = True
End With
FixBackColor
End Sub
Private Sub Fixgrd2()
With vsColor
.ColWidth(0) = 1300
For i = 1 To .Cols - 1
    .ColWidth(i) = nColWidth
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
Dim nColor As Integer, nScal As Integer
Dim lNew As Boolean
Dim aInsert(13, 1)
cModel = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.Text, 4) & retCode(xCode.Text, con)
cModelNo = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.Text, 4)
aInsert(0, 0) = "Model"
aInsert(0, 1) = addstring(cModel)

aInsert(1, 0) = "ModelNO"
aInsert(1, 1) = addstring(cModelNo)

aInsert(2, 0) = "Desca"
aInsert(2, 1) = addstring(xDesca.Text)

aInsert(3, 0) = "FACT"
aInsert(3, 1) = addstring(xFact.BoundText)

aInsert(4, 0) = "MOSM"
aInsert(4, 1) = addstring(xMosm.BoundText)

aInsert(5, 0) = "MODELFACT"
aInsert(5, 1) = addstring(RetZero(xModelFact.Text, 4))

aInsert(6, 0) = "MODELFACT0"
aInsert(6, 1) = addstring(DelZero(xModelFact.Text))

aInsert(7, 0) = "SUPP"
aInsert(7, 1) = addstring(xSupp.Text)

aInsert(8, 0) = "FIXPRICE"
aInsert(8, 1) = Val(xFixPrice.Text)

aInsert(9, 0) = "OKAZ"
aInsert(9, 1) = Val(xOkaz.Text)

aInsert(10, 0) = "RATE"
aInsert(10, 1) = Val(xRate.Text)

aInsert(11, 0) = "REDEM"
aInsert(11, 1) = Val(xRedem.Text)

aInsert(12, 0) = "Code"
aInsert(12, 1) = addstring(xCode.Text)

aInsert(13, 0) = "[group]"
aInsert(13, 1) = addstring(xGroup.BoundText)

con.BeginTrans
If Model_Found(xModel.Text, con) = 0 Then
    con.Execute CreateInsert(aInsert, "FILE1_10H")
Else
    con.Execute CreateUpdate(aInsert, "FILE1_10H", " WHERE MODEL = " & MyParn(xModel.Text))
End If
myreplaceGrd
con.CommitTrans
'openCardTable
myreplace = True
Exit Function
myerror:
'con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function MYVALID() As Boolean
If xFact.BoundText = "" Then
    MsgBox "«·„’‰⁄"
    Exit Function
End If
If xModelFact.Text = "" Then
    MsgBox "—ﬁ„ «·„ÊœÌ·"
    Exit Function
End If
If xMosm.BoundText = "" Then
    MsgBox "«·„Ê”„ €Ì— „”Ã·"
    Exit Function
End If

'If xSection.BoundText = "" Then
'    MsgBox "«·ﬁ”„ €Ì— „”Ã·"
'    Exit Function
'End If

If xDesca.Text = "" Then
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
MYVALID = True
End Function
Private Sub Handlecontrols(nMode)
CmdDel.Enabled = (nMode = LoadMode) And bEdit
cmdSave.Enabled = bEdit
CmdUndo.Enabled = bEdit
'CmdInform.Enabled = (nMode = LoadMode)
xModel.Tag = nMode
xModelFact.Enabled = False
'xSection.Locked = sSection <> ""
'xSection.TabStop = sSection = ""
xModel.Tag = nMode
xModel.Enabled = Not (nMode = LoadMode)
End Sub
Private Sub myDefineGrd()
With vsScal
    .Cols = 1
    .Rows = 7
    .TextMatrix(0, 0) = "«·„ﬁ«”"
    .TextMatrix(1, 0) = "”⁄— «·„’‰⁄"
    .TextMatrix(2, 0) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(3, 0) = "”⁄— √Êﬂ«“ÌÊ‰"
    .TextMatrix(4, 0) = "”⁄— Ã„·…"
    .TextMatrix(5, 0) = "ﬁ”„"
    .TextMatrix(6, 0) = ""
    .FixedCols = 1
    .Cols = .Cols + 1
    .TextMatrix(0, 1) = "---"
    .Cols = .Cols + 1
    Fixgrd
End With
With vsColor
    .Cols = 1
    .Rows = 3
    .TextMatrix(0, 0) = "«·ﬂÊœ"
    .TextMatrix(1, 0) = "«··Ê‰"
    .Cols = .Cols + 1
    .TextMatrix(1, 1) = "---"
    .TextMatrix(0, 1) = 1
    .Cols = .Cols + 1
    vsColor.RowHidden(vsColor.Rows - 1) = True
    Fixgrd2
End With
End Sub

Private Sub xGroup_Validate(Cancel As Boolean)
On Error GoTo myerror
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
If Trim(xDesca.Text) = "" Then xDesca.Text = xGroup.Text
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(xItem) Then
        If cFilter <> "" Then
            cFilter = ""
            openCardTable
        End If
        With CardTable
            Dim cString As String
            cString = "SELECT MODEL FROM FILE1_10"
            cString = cString & turn(cString) & " ITEM = " & Val(xItem.Text)
            xModel.Text = GetDesca(cString)
            If xModel.Text <> "" Then
                CardTable.Find "Model = " & MyParn(xModel.Text), , adSearchForward, adBookmarkFirst
                If Not CardTable.EOF Then myload Else mydefine
            Else
                mydefine
            End If
        End With
    End If
End Sub
Private Sub LoadPhoto()
On Error GoTo myerror
imgx1.Images.Clear
imgx1.Import.FromFile retFile(xModelNo.Text)
Exit Sub
myerror:
Err.Clear
End Sub
Private Sub ScanImage()
On Error GoTo myerror
Set twain = New ImgXTwain
twain.OpenTwain Me.hWnd
'If twain.QuerySupport(ixtcResolution) Then
'     twain.Resolution = 150
'End If
twain.SelectSource
twain.Acquire False, Me.hWnd
Exit Sub
myerror:
MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub
Private Sub Twain_ImageAcquired(image As ImgX_Image)
    addImage image
End Sub
Private Sub Twain_TwainError(ByVal erNum As Long, ByVal erSource As String, ByVal Description As String)
MsgBox "Error Number:  " & erNum & vbCrLf & Description, vbInformation, erSource
End Sub
Private Sub Twain_CanCloseTwain()
    ' This event is called after you call Acquire.
    ' It let's you know when it's safe to call CloseTwain.
    twain.CloseTwain
    ' Steps menu
End Sub
Private Sub addImage(image As ImgX_Image)
Dim csource As String
csource = retFile(xModelNo.Text)
imgx1.Images.Replace image, , False
imgx1.TIFCompression = ixtcJPEGCompression
imgx1.Export.ToFile csource, ixfsJPG
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
xItem.SelStart = 0
xItem.SelLength = Len(xItem.Text)
xItem.BackColor = &H80FFFF
End Sub
Private Sub xModelFact_Validate(Cancel As Boolean)
If xModelFact.Text = "" Then
    Cancel = True
    Exit Sub
End If
xModelFact.Text = RetZero(xModelFact.Text, 4)
xModel.Text = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.Text, 4) & xSupp.Text
xModelNo.Text = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.Text, 4)
openCardTable
CardTable.Find "MODEL = " & MyParn(xModel.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    xModelFact.Enabled = False
    myload
End If
End Sub
Private Sub xModelNo_GotFocus()
xModelNo.SelStart = 0
xModelNo.SelLength = Len(xModelNo.Text)
xModelNo.BackColor = &H80FFFF
End Sub
Private Sub xModel_GotFocus()
xModel.SelStart = 0
xModel.SelLength = Len(xModel.Text)
xModel.BackColor = &H80FFFF
End Sub
Private Sub xModelFACT_GotFocus()
xModelFact.SelStart = 0
xModelFact.SelLength = Len(xModelFact.Text)
xModelFact.BackColor = &H80FFFF
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
xSupp.SelLength = Len(xSupp.Text)
xSupp.BackColor = &H80FFFF
End Sub
Private Sub xFact_GotFocus()
xFact.BackColor = &H80FFFF
End Sub
Private Sub xMosm_GotFocus()
xMosm.BackColor = &H80FFFF
End Sub
Private Sub xOkaz_GotFocus()
xOkaz.SelStart = 0
xOkaz.SelLength = Len(xOkaz.Text)
xOkaz.BackColor = &H80FFFF
End Sub
Private Sub xRedem_GotFocus()
xRedem.SelStart = 0
xRedem.SelLength = Len(xRedem.Text)
xRedem.BackColor = &H80FFFF
End Sub
Private Sub xDesca_GotFocus()
If Trim(xDesca.Text) = Trim(xGroup.Text) Then
    xDesca.Text = xGroup.Text & " "
    xDesca.SelStart = Len(xDesca.Text)
    xDesca.SelLength = 1
Else
    xDesca.SelLength = Len(xDesca.Text)
End If
xDesca.BackColor = &H80FFFF
End Sub
Private Sub xCode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.Text)
xCode.BackColor = &H80FFFF
End Sub
Private Sub xRate_GotFocus()
xRate.SelStart = 0
xRate.SelLength = Len(xRate.Text)
xRate.BackColor = &H80FFFF
End Sub
Private Sub xFixPrice_GotFocus()
xFixPrice.SelStart = 0
xFixPrice.SelLength = Len(xFixPrice.Text)
xFixPrice.BackColor = &H80FFFF
End Sub
Private Sub xGroup_GotFocus()
xGroup.BackColor = &H80FFFF
End Sub
Private Sub xSection_GotFocus()
xSection.BackColor = &H80FFFF
End Sub
Private Sub xitem_LostFocus()
xItem.BackColor = &H80000005
End Sub
Private Sub xModelNo_LostFocus()
xModelNo.BackColor = &H80000005
End Sub
Private Sub XModel_LostFocus()
xModel.BackColor = &H80000005
End Sub
Private Sub XModelFACT_LostFocus()
xModelFact.BackColor = &H80000005
End Sub
Private Sub xSupp_LostFocus()
xSupp.BackColor = &H80000005
End Sub
Private Sub xFact_LostFocus()
xFact.BackColor = &H80000005
End Sub
Private Sub xMosm_LostFocus()
xMosm.BackColor = &H80000005
End Sub
Private Sub xOkaz_LostFocus()
xOkaz.BackColor = &H80000005
End Sub
Private Sub xRedem_LostFocus()
xRedem.BackColor = &H80000005
End Sub
Private Sub xDesca_LostFocus()
xDesca.BackColor = &H80000005
End Sub
Private Sub xCode_LostFocus()
xCode.BackColor = &H80000005
End Sub
Private Sub xRate_LostFocus()
xRate.BackColor = &H80000005
End Sub
Private Sub xFixPrice_LostFocus()
xFixPrice.BackColor = &H80000005
End Sub
Private Sub xGroup_LostFocus()
xGroup.BackColor = &H80000005

End Sub
Private Sub xSection_LostFocus()
xSection.BackColor = &H80000005
End Sub
Private Sub myUndo()
'On Error GoTo myerror
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(xModel.Text) <> "" Then
        CardTable.Find "MODEL = " & MyParn(xModel.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "select FILE1_10H.* from file1_10H"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " order by model desc"

Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Function LastCol() As Integer
With vsScal
For i = 1 To .Cols - 1
    If Val(.TextMatrix(.Rows - 1, i)) > LastCol Then LastCol = Val(.TextMatrix(.Rows - 1, i))
Next
LastCol = LastCol + 1
End With
End Function
Private Sub FixCode()
With vsScal
For i = 1 To .Cols - 1
    If .TextMatrix(.Rows - 1, i) = "" Then .TextMatrix(.Rows - 1, i) = LastCol
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
            End If
         End If
    End If
End If
CalcRate Col, Row
End With
End Sub
Private Sub myAddModel()
'cFilter = "SUBSTRING(MODEL,1,16) = " & MyParn(Mid(sModel, 1, 16))
cFilter = "Model = " & MyParn(sModel)
openCardTable
If Not (CardTable.EOF And CardTable.BOF) Then
    myUndo
    Exit Sub
End If

Dim sModelNo As String
sModelNo = sFact & sMosm & RetZero(sModelFact, 4)
cFilter = "ModelNo = " & MyParn(sModelNo)
openCardTable
myUndo
If (xModel.Text <> sModel) Or xModel.Tag = DefineMode Then
    xModel.Tag = DefineMode
    xFact.BoundText = sFact
    xMosm.BoundText = sMosm
    xModelFact.Text = sModelFact
    xModelNo.Text = sFact & sMosm & RetZero(sModelFact, 4)
    xModel.Text = sModel
    xSupp.Text = sSupp
    xCode.Text = sCode
    If xCode.Text <> "" Then
        Dim aRet As Variant
        aRet = aGetDesca("Select desca,Rate from file4_10 where code = " & MyParn(xCode.Text))
        If UBound(aRet) > 0 Then
            xCodeDesc.Caption = aRet(1)
            xRate.Text = aRet(2) & ""
        End If
    End If
    For i = 1 To vsScal.Cols - 1
        vsScal.TextMatrix(vsScal.Rows - 1, i) = ""
    Next
    FixBackColor
End If
cFilter = "FILE1_10H.MODEL = " & MyParn(sModel)
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
    sBig = Myvalue(loctable!big)
End If
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CalcRate(Col, Row)
With vsScal
If Col < .Cols - 1 And Col > 0 Then
    If Row <> 2 And Trim(.TextMatrix(2, Col)) = "" Then
        If Val(xRate.Text) <> 0 And Val(.TextMatrix(1, Col)) > 0 And (Val(.TextMatrix(3, Col)) = 0 Or Row = 1) Then
            Dim nPrice As Single, nRem As Single
            nPrice = Val(.TextMatrix(1, Col)) * (100 + Val(xRate.Text)) / 100
            nRem = nPrice - Int(nPrice)
            Select Case nRem
                Case Is = 0
                    nPrice = Int(nPrice)
                Case Is > 0
                    nPrice = Int(nPrice) + 1
                Case Is < 0.5
                    nPrice = Int(nPrice) + 0.5
            End Select
            .TextMatrix(2, Col) = Myvalue(nPrice, "#0.00")
        End If
    End If
    
    If Row <> 3 And Trim(.TextMatrix(3, Col)) = "" Then
        If (Val(.TextMatrix(3, Col)) = 0 Or Row = 1) And Val(xOkaz.Text) > 0 Then
            .TextMatrix(3, Col) = Myvalue(Val(.TextMatrix(1, Col)) * (100 + Val(xOkaz.Text)) / 100, "#0.00")
        End If
    End If
    If Row <> 4 And Trim(.TextMatrix(4, Col)) = "" Then
        If Val(xbig.Text) <> 0 And Val(.TextMatrix(1, Col)) > 0 And (Val(.TextMatrix(4, Col)) = 0 Or Row = 1) Then
            nPrice = Val(.TextMatrix(1, Col)) * (100 + Val(xbig.Text)) / 100
            nRem = nPrice - Int(nPrice)
            Select Case nRem»
                Case Is = 0
                    nPrice = Int(nPrice)
                Case Is > 0
                    nPrice = Int(nPrice) + 1
                Case Is < 0.5
                    nPrice = Int(nPrice) + 0.5
            End Select
            .TextMatrix(4, Col) = Myvalue(nPrice, "#0.00")
        End If
    End If
End If
End With
End Sub
Private Sub MakeCopy()
cFilter = ""
xModelNo.Text = ""
With vsScal
For nCol = 1 To vsScal.Cols - 1
    .TextMatrix(.Rows - 1, nCol) = ""
Next
End With

With vsColor
For nCol = 1 To vsScal.Cols - 1
    .TextMatrix(.Rows - 1, nCol) = ""
Next
xModelNo.Text = ""
xModelFact.Text = ""
xModelFact.Enabled = True
FixBackColor
FixBackColor2
End With
End Sub
Private Sub myreplaceGrd()
FixCode
Dim aInsert(26, 1)
Dim sitem As String, aRet As Variant

cModel = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.Text, 4) & retCode(xCode.Text, con)
cModelNo = xFact.BoundText & xMosm.BoundText & RetZero(xModelFact.Text, 4)

aInsert(0, 0) = "Model"
aInsert(0, 1) = addstring(cModel)

aInsert(1, 0) = "ModelNO"
aInsert(1, 1) = addstring(cModelNo)

aInsert(2, 0) = "Desca"
aInsert(2, 1) = addstring(xDesca.Text)

aInsert(3, 0) = "FACT"
aInsert(3, 1) = addstring(xFact.BoundText)

aInsert(4, 0) = "MOSM"
aInsert(4, 1) = addstring(xMosm.BoundText)

aInsert(5, 0) = "MODELFACT"
aInsert(5, 1) = addstring(RetZero(xModelFact.Text, 4))

aInsert(6, 0) = "MODELFACT0"
aInsert(6, 1) = addstring(DelZero(xModelFact.Text))

aInsert(7, 0) = "SUPP"
aInsert(7, 1) = addstring(xSupp.Text)

aInsert(8, 0) = "FIXPRICE"
aInsert(8, 1) = Val(xFixPrice.Text)

aInsert(9, 0) = "OKAZ"
aInsert(9, 1) = Val(xOkaz.Text)

aInsert(10, 0) = "RATE"
aInsert(10, 1) = Val(xRate.Text)

aInsert(11, 0) = "REDEM"
aInsert(11, 1) = Val(xRedem.Text)

aInsert(12, 0) = "Code"
aInsert(12, 1) = addstring(xCode.Text)

aInsert(13, 0) = "[group]"
aInsert(13, 1) = addstring(xGroup.BoundText)

For nScal = 1 To vsScal.Cols - 2
    If validCol(nScal) Then
        For nColor = 1 To vsColor.Cols - 2
            lNew = False
            If Val(vsColor.TextMatrix(0, nColor)) <> 0 And vsScal.TextMatrix(0, nScal) <> "" Then
                sitem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(xModel.Text) & " AND C_COLOR = " & Val(vsColor.TextMatrix(0, nColor)) & " AND code_SCAL = " & Val(vsScal.TextMatrix(vsScal.Rows - 1, nScal)), con)
                If sitem = "" Then
                    sitem = Val(GetDesca("SELECT MAX(ITEM),MAX(c_scal) FROM FILE1_10", con)) + 1
                    If Val(sitem) < 100 Then sitem = 101
                    lNew = True
                End If
                
                aInsert(15, 0) = "item"
                aInsert(15, 1) = addvalue(sitem)
                
                aInsert(16, 0) = "Cost"
                aInsert(16, 1) = Val(vsScal.TextMatrix(1, nScal))
                
                aInsert(17, 0) = "Price"
                aInsert(17, 1) = Val(vsScal.TextMatrix(2, nScal))
                
                aInsert(18, 0) = "Price2"
                aInsert(18, 1) = Val(vsScal.TextMatrix(3, nScal))
                
                aInsert(19, 0) = "Cost2"
                aInsert(19, 1) = Val(vsScal.TextMatrix(4, nScal))
            
                aInsert(20, 0) = "Scal"
                aInsert(20, 1) = addstring(vsScal.TextMatrix(0, nScal))
            
                aInsert(21, 0) = "C_Scal"
                aInsert(21, 1) = nScal
                
                aInsert(22, 0) = "Color"
                aInsert(22, 1) = addstring(vsColor.TextMatrix(1, nColor))
        
                aInsert(23, 0) = "C_Color"
                aInsert(23, 1) = addstring(vsColor.TextMatrix(0, nColor))
        
                aInsert(24, 0) = "SHOWSALES"
                aInsert(24, 1) = Val(xShowSales.Value)
        
                aInsert(25, 0) = "Code_Scal"
                aInsert(25, 1) = Val(vsScal.TextMatrix(vsScal.Rows - 1, nScal))
        
                aInsert(26, 0) = "SEC"
                aInsert(26, 1) = addstring(vsScal.TextMatrix(5, nScal))
        
                If lNew Then
                    con.Execute CreateInsert(aInsert, "File1_10")
                Else
                    con.Execute CreateUpdate(aInsert, "File1_10", " where item = " & sitem)
                End If
            End If
        Next
    End If
Next nScal
End Sub
