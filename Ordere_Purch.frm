VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Ordere_Purch 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9855
   ClientLeft      =   2040
   ClientTop       =   2490
   ClientWidth     =   11430
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
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   13065
   ScaleWidth      =   23910
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   570
      Left            =   7785
      TabIndex        =   83
      Top             =   2025
      Visible         =   0   'False
      Width           =   1140
      _cx             =   2011
      _cy             =   1005
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   14
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
   Begin VB.CommandButton cmd_ex 
      BackColor       =   &H00FFFFC0&
      Caption         =   " «” Ì—«œ „‰ «·«ﬂ”Ì· "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   12330
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   81
      Top             =   1980
      Width           =   1410
   End
   Begin VB.CommandButton cmdpast 
      Caption         =   "PASTE"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   135
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   78
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   135
      UseMaskColor    =   -1  'True
      Width           =   1050
   End
   Begin VB.CommandButton cmdcopy 
      Caption         =   "COPY"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1170
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   77
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   135
      UseMaskColor    =   -1  'True
      Width           =   1050
   End
   Begin VB.Frame Frame15 
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
      Height          =   2670
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   69
      Top             =   0
      Width           =   4110
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   2355
         Left            =   90
         TabIndex        =   70
         Top             =   225
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   4154
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   90
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame12 
      Height          =   645
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   6930
      Width           =   2130
      Begin VB.Label lblTotalQuant 
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·ﬂ„Ì… :"
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
         Left            =   855
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   225
         Width           =   1185
      End
      Begin VB.Label xtotalQuant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1860
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   810
      Width           =   2085
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         TabIndex        =   61
         Top             =   1170
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   1058
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Ordere_Purch.frx":0000
         Alignment       =   4
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   1035
         TabIndex        =   62
         Top             =   180
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1693
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Ordere_Purch.frx":25CC
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   90
         TabIndex        =   63
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1693
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Ordere_Purch.frx":4C44
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   7335
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   56
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
         Picture         =   "Ordere_Purch.frx":7285
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Ordere_Purch.frx":9455
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   57
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
         Picture         =   "Ordere_Purch.frx":B59D
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Ordere_Purch.frx":D765
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   58
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
         Picture         =   "Ordere_Purch.frx":F8B4
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Ordere_Purch.frx":11A94
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   59
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
         Picture         =   "Ordere_Purch.frx":13BEF
         Caption         =   " «Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Ordere_Purch.frx":15DAB
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   54
      Top             =   12690
      Width           =   23910
      _ExtentX        =   42175
      _ExtentY        =   661
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame9 
      Height          =   780
      Left            =   7785
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   0
      Width           =   7350
      Begin VB.CheckBox xprintnoprice 
         Alignment       =   1  'Right Justify
         Caption         =   "»œÊ‰  ”⁄—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   225
         Width           =   870
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   " ›’Ì·Ï «ﬁ”«„"
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
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   315
         Width           =   1365
      End
      Begin Threed.SSCommand CMD_PRINT 
         Height          =   510
         Left            =   5400
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   180
         Width           =   1860
         _ExtentX        =   3281
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
         Picture         =   "Ordere_Purch.frx":17EFA
         Caption         =   "ÿ»«⁄… „” ‰œ «·ÿ·»Ì… "
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   0
         TabIndex        =   75
         Top             =   270
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label16 
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
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   315
         Width           =   570
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1440
      Width           =   1365
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Ordere_Purch.frx":1A309
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
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
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Ordere_Purch.frx":1C882
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   15120
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton CmdDelInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1275
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Ordere_Purch.frx":1EBE5
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2505
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Ordere_Purch.frx":2147F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   3735
         Picture         =   "Ordere_Purch.frx":23A2B
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Ordere_Purch.frx":261FE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   -3330
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   -675
      Visible         =   0   'False
      Width           =   6810
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·„” ‰œ „€·ﬁ"
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
         Left            =   200
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   100
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1815
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   720
      Width           =   11175
      Begin VB.CheckBox xorder 
         Alignment       =   1  'Right Justify
         Caption         =   "ÿ·»Ì… ··„’‰⁄"
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   1305
         Width           =   960
      End
      Begin VB.CheckBox XISEND_ORDER 
         Alignment       =   1  'Right Justify
         Caption         =   "ÿ·»Ì…   „ ≈€·«ﬁÂ«"
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
         Height          =   495
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   1215
         Width           =   1905
      End
      Begin VB.TextBox Xremark 
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
         Height          =   330
         Left            =   4635
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   855
         Width           =   5280
      End
      Begin VB.TextBox xDate 
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
         Height          =   330
         Left            =   1485
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   2400
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
         Height          =   330
         Left            =   8865
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   495
         Width           =   1095
      End
      Begin VB.TextBox xDoc_No 
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
         Height          =   330
         Left            =   8865
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
      End
      Begin VB.TextBox xInv_No 
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
         Height          =   330
         Left            =   4635
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   135
         Width           =   1320
      End
      Begin VB.CheckBox xSupp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„ﬂ » Ã„·…"
         Enabled         =   0   'False
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
         Height          =   375
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   -450
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   1485
         TabIndex        =   4
         Top             =   900
         Width           =   2400
         _ExtentX        =   4233
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
      Begin Threed.SSCommand cmdAddItems 
         Height          =   1005
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1773
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Ordere_Purch.frx":2866A
         Caption         =   " ”ÃÌ· „ÊœÌ·«  "
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   1485
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   180
         Width           =   2400
         _ExtentX        =   4233
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
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   1440
         TabIndex        =   80
         Top             =   1260
         Width           =   1875
         _ExtentX        =   3307
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
         Picture         =   "Ordere_Purch.frx":2AE72
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê”„ :"
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
         Index           =   1
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄ "
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
         Left            =   3930
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   900
         Width           =   465
      End
      Begin VB.Label xsup_desca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1350
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ«  :"
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
         Left            =   10050
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   900
         Width           =   780
      End
      Begin VB.Label xsubcode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8865
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1485
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê—œ :"
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
         Left            =   10050
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   555
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10050
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   210
         Width           =   930
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   540
         Width           =   645
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5985
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   495
         Width           =   2850
      End
      Begin VB.Label xBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   495
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·›« Ê—… :"
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
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   1035
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   375
      Left            =   135
      TabIndex        =   16
      Top             =   6930
      Visible         =   0   'False
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   3780
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
   Begin VB.Frame Frame7 
      Height          =   1455
      Left            =   10260
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   6930
      Width           =   9915
      Begin VB.TextBox xr_Discount 
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
         Height          =   315
         Left            =   3195
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   585
         Width           =   615
      End
      Begin VB.TextBox xTax 
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
         Height          =   330
         Left            =   765
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   585
         Width           =   960
      End
      Begin VB.TextBox xr_tax 
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
         Height          =   330
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   585
         Width           =   465
      End
      Begin VB.TextBox xDiscount 
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
         Height          =   315
         Left            =   3825
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ «·„Ê—œ :"
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label xRateSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   945
         Width           =   1590
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ √’‰«› :"
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
         Left            =   8415
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label xDiscountItem 
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
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   585
         Width           =   1245
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï «·√’‰«› :"
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
         Left            =   8415
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label xtotalOrg 
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
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   225
         Width           =   1245
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "÷—«∆» «·„»Ì⁄«  :"
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   630
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "»⁄œ«·Œ’„ :"
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
         Left            =   1815
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   270
         Width           =   960
      End
      Begin VB.Label xTotalNoTax 
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label xTotal 
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
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   945
         Width           =   1470
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì «·›« Ê—… :"
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   990
         Width           =   1365
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   165
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   630
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ ›« Ê—… :"
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   1185
      End
      Begin VB.Label xTotalItem 
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
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3195
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   1605
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï »⁄œ Œ’„ «·«’‰«› :"
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   2130
      End
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   -720
      Top             =   -270
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
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   -5220
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   -450
      Visible         =   0   'False
      Width           =   5775
      Begin VB.Label xTime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   180
         Width           =   2625
      End
      Begin VB.Label xUserName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   -135
         Width           =   2850
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   4275
      Top             =   -180
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   3780
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4125
      Left            =   135
      TabIndex        =   6
      Top             =   2700
      Width           =   20040
      _cx             =   35348
      _cy             =   7276
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   14
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
   Begin MSAdodcLib.Adodc data22 
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
Attribute VB_Name = "Ordere_Purch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cFileHeader As String, sDoc_no As String, cFile As String, cFile_FR As String, cFileHeader_FR As String
Dim confact As New ADODB.Connection, cDocCopy As String, cCodeCopy As String
Dim confact2 As New ADODB.Connection
Dim nRound As Integer
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim searchitem As New Search3
Dim Search1 As New Search3, Search2 As New Search3, oSearchSup As New Search3
Dim Search_FACT As New Search3_FACT

Dim bedit As Boolean, pGrCust As String
Dim cFileClient, cFieldClient, cCodeDesca As String
Dim formMode, dDateLast As String
Dim bUpdated As Boolean
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional bAddModel As Boolean = False) As Boolean
Dim aInsert(9, 1), nR_Tax As Double, nR_DISC As Double, nRCost As Double
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(xDate.text)

aInsert(2, 0) = "Code"
aInsert(2, 1) = addstring(xCode.text)

aInsert(3, 0) = "Discount"
aInsert(3, 1) = Val(xDiscount.text)

aInsert(4, 0) = "Tax"
aInsert(4, 1) = Val(xTax.text)

aInsert(5, 0) = "Store"
aInsert(5, 1) = addstring(xStore.BoundText)

aInsert(6, 0) = "INV_NO"
aInsert(6, 1) = addstring(xInv_No.text)

aInsert(7, 0) = "mosm"
aInsert(7, 1) = addstring(xMosm.BoundText)

aInsert(8, 0) = "ISEND_ORDER"
aInsert(8, 1) = IIf(XISEND_ORDER.Value, 1, 0)


aInsert(9, 0) = "remark"
aInsert(9, 1) = addstring(Xremark.text)

con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    If myPublic = 0 Or myPublic = 1 Then
        xDoc_No.text = RetZero(Newflag(cFileHeader, "doc_no", con))
    Else
        xDoc_No.text = RetZero(Newflag_PurchBr(cFileHeader, "doc_no", xStore.BoundText, con))
    End If
    aInsert(0, 1) = addstring(xDoc_No.text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
   con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & MyParn(xDoc_No.text))
End If

If bAddModel Then
    myreplaceGrd
End If

con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
'On Error GoTo myerror
If ActiveControl.Name = CmdInform.Name Then
    xDoc_No.text = Search1.grid1.TextMatrix(Search1.grid1.Row, 0)
    Search1.Hide
    myUndo
ElseIf TypeOf ActiveControl Is TextBox Then
    ActiveControl.text = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    Unload oSearchSup
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmd_ex_Click()
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet
    Dim xlwbook As Excel.Workbook
    Dim cDoc_No As String
    Dim cBarCode As String
    Dim cModel As String
    Dim cItem As String, cENA As String
    Dim nQTY As Double, nRec As Double
    Dim nPrice As Double, cGroup As String, cCodeGroup As String
    Dim nDisc As Double, cDesca2 As String
    Dim nNewItem As String, cCode As String, cSex As String
    Dim lAddTrue As Boolean, cModelAll As String
    Dim cStr1 As String, cStr2 As String, nCoutnR As Double, cDocNo As String, nRowDoc As Double
    prog1.Min = 0
    prog1.Value = 0
    prog1.Visible = True
    If MsgBox("≈÷«›… «’‰«› „‰ „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
            For nRow = 2 To 10000
                If Len(xlsheet.Cells(nRow, 1)) > 1 Then
                    nRec = nRow
                End If
                If Len(xlsheet.Cells(nRow, 1)) = 0 Then Exit For
            Next nRow
            prog1.Max = nRec
            i = 0
            For nRow = 2 To nRec
                prog1.Visible = True
                prog1.Value = nRow
                If Len(xlsheet.Cells(nRow, 1)) = 0 Then Exit For
                cDoc_No = xDoc_No.text
                cBarCode = Trim(xlsheet.Cells(nRow, 1))
                nQTY = Val(xlsheet.Cells(nRow, 2))
                    
                cItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & Val(cBarCode) & " OR BARCODE13 = " & MyParn(cBarCode) & " OR BARCODE_GS1 = " & MyParn(cBarCode), con)
                If cItem <> "" Then
                    cStr2 = " ; " & cStr2 & " INSERT INTO " & cFile & " (DOC_NO , ITEM , QUANT ) VALUES     ( " & addstring(cDoc_No) & "  , " & addstring(cItem) & " , " & nQTY & ")"
                End If
                If nCoutnR = 100 Then
                    If cStr2 <> "" Then con.Execute cStr2
                    nCoutnR = 0
                    cStr2 = ""
                Else
                    nCoutnR = nCoutnR + 1
                End If
            Next nRow
            If cStr2 <> "" Then con.Execute cStr2
            con.Execute " UPDATE File6_50 SET   File6_50.MODEL = FILE1_10.MODEL , File6_50.PRICE = FILE1_10.COST FROM File6_50 INNER JOIN FILE1_10 ON File6_50.ITEM = FILE1_10.ITEM  WHERE ( File6_50.PRICE = 0 OR File6_50.PRICE IS NULL ) AND File6_50.DOC_NO = " & MyParn(xDoc_No.text)
            myloadgrd
        End If
    End If

End Sub

Private Sub cmd_excel_Click()
    If xorder.Value = 0 Then
        grid1.ColWidth(16) = 2000
        
        grid1.ColHidden(14) = False
        grid1.ColHidden(15) = False
        grid1.ColHidden(12) = True
        grid1.ColHidden(13) = True
        grid1.ColHidden(8) = True
        grid1.ColHidden(9) = True
        grid1.ColHidden(10) = True
        
        ToFileExel2 grid1, , , , , 1.1, , , , , , Me
    
        grid1.ColWidth(16) = 900
        grid1.ColHidden(14) = True
        grid1.ColHidden(15) = True
        grid1.ColHidden(12) = False
        grid1.ColHidden(13) = False
        grid1.ColHidden(8) = False
        grid1.ColHidden(9) = False
        grid1.ColHidden(10) = False
    Else
        myloadgrd2
        ToFileExel2 grid2, , , , , 1.1, , , , , , Me
    End If
End Sub

Private Sub cmdAddItems_Click()
    Dim lIsNoItem  As Boolean
    myAddModels
End Sub
Private Sub cmdAddItems_GotFocus()
    cmdAddItems.BackColor = vbYellow
End Sub
Private Sub cmdAddItems_LostFocus()
    cmdAddItems.BackColor = &H8000000F
End Sub

Private Sub cmdcopy_Click()
    cDocCopy = xDoc_No.text
    cCodeCopy = xCode.text
    Inform " Copy " & cDocCopy

End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.Execute "Delete  From FILE6_50  where Doc_No = " & MyParn(xDoc_No.text)
    con.Execute "Delete  From FILE6_50H where Doc_No = " & MyParn(xDoc_No.text)
'    con.CommitTrans
    myUndo
    CmdNewInv_Click
End If
Exit Sub
myerror:
'con.RollbackTrans
MsgBox Err.Description
Err.Clearh
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdInform_Click()
    CardLookup
End Sub
Private Sub CmdNewInv_Click()
    myDefine
    xInv_No.SetFocus
End Sub
Private Sub cmdpast_Click()
    If cDocCopy <> "" And xCode.text = cCodeCopy Then
        If grid1.Rows > 3 Then
            MsgBox " ÌÊÃœ „ÊœÌ·«  „”Ã·… ⁄·Ï «·„” ‰œ "
            Exit Sub
        End If
        If xDoc_No.Enabled Then
            If MsgBox("  ﬂ—«— „” ‰œ —ﬁ„  " & cDocCopy, vbYesNo) = vbYes Then
                If Not mysave Then Exit Sub
                con.Execute " INSERT INTO FILE6_50 (doc_no , ITEM, QUANT, PRICE, DISCOUNT, ROW, MODEL, REMARK, r_tax) SELECT " & addstring(xDoc_No.text) & "  , ITEM, QUANT, PRICE, DISCOUNT, ROW, MODEL, REMARK, r_tax FROM  FILE6_50 AS FILE6_50_1 WHERE DOC_NO = " & MyParn(cDocCopy)
                myUndo
            End If
        End If
    End If
End Sub

Private Sub cmdSave_Click()
    mysave
End Sub
Private Sub CmdUndo_Click()
    myUndo
End Sub
Private Sub CMD_PRINT_Click()
    myUndo
    doprint
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
    
    nRound = 2
    
'    openConFACT confact
'    openConFACT2 confact2

    con.CursorLocation = adUseClient
    con.Open strCon
    
    Set data2.Recordset = myRecordSet("SELECT * FROM FILE1_10SC ", con)
    Set xSection.RowSource = data2
    xSection.ListField = "Desca"
    xSection.BoundColumn = "CODE"

    pGrCust = "04"
    cFile = "File6_50"
    cFileHeader = "File6_50H"
    
    cFileClient = "File4_10"
    Me.Caption = "«„— ‘—«¡"
        
    Set data1.Recordset = myRecordSet("SELECT * FROM STORE_BR", con)
    Set xStore.RowSource = data1
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    
    Set data4.Recordset = myRecordSet("SELECT * FROM MOSM ", con)
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    With grid1
        .Cols = 10
        .Rows = 1
    End With
    
    
    Set grid1.DataSource = DATA3
    Set grid2.DataSource = data22
        
    xDoc_No.text = sDoc_no
    myUndo
    If sDoc_no = "" Then myDefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
    Err.Clear
    Set Purchasefrm = Nothing
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If xClosed.Value = 1 Then Exit Sub
With grid1
    If .Col = 9 Then
        con.Execute " update FILE6_50 set discount = " & Val(.TextMatrix(.Row, 9)) & " FROM   " & cFile & " JOIN FILE1_10 ON " & cFile & ".ITEM = FILE1_10.ITEM  where FILE1_10.model = " & MyParn(.TextMatrix(.Row, 13)) & " and doc_no = " & MyParn(xDoc_No.text)
    End If
    myloadgrd
End With
End Sub
Private Sub grid1_DblClick()
    If xClosed.Value <> 0 Then Exit Sub
    If grid1.Col = 8 Then
        myAddModels True
    End If
End Sub
Private Sub grid1_EnterCell()
If xClosed.Value = 1 Then Exit Sub
If grid1.TextMatrix(grid1.Row, 13) <> "" Then LoadPhoto grid1.TextMatrix(grid1.Row, 13) Else ImgX1.Images.Clear
If grid1.Col = 9 Or grid1.Col = 14 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub ImgX1_Click()
If validPhoto(retFile(grid1.TextMatrix(grid1.Row, 13))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(grid1.TextMatrix(grid1.Row, 13))
    oMaxPhoto.cDesca = grid1.TextMatrix(grid1.Row, 4)
    oMaxPhoto.Show 1
End If
End Sub
Private Sub xCode_GotFocus()
    myGotFocus xCode
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then SuppLookupAll Me, oSearchSup
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
xCodeDesca.Caption = ""
xr_Discount.text = ""
xsubcode.Caption = ""
xSupp.Value = 0
xsup_desca.Caption = ""
If xCode.text = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 3)
If xCode.text = "" Then Exit Sub

Dim cString As String, cWhere As String
cString = "select code,desca, SUPP , DISC ,SUBCODE  FROM FILE4_10"
cString = cString & " WHERE CODE = " & MyParn(xCode.text)
aRet = aGetDesca(cString, con)
If UBound(aRet) > 0 Then
    xCode.text = aRet(1)
    xCodeDesca.Caption = aRet(2)
    xSupp.Value = IIf(aRet(3), 1, 0)
    xsup_desca.Caption = IIf(aRet(3), "„ﬂ » Ã„·…", "„’‰⁄")
    xr_Discount.text = Val(aRet(4) & "")
    xRateSup.Caption = Val(aRet(4) & "")
    xsubcode.Caption = aRet(5) & ""
End If
End Sub
Private Sub xdate_GotFocus()
    myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
End Sub

Private Sub xDate_Validate(Cancel As Boolean)
myValidDate xDate
End Sub
Private Function MYVALID() As Boolean
    If Not IsDate(xDate.text) Then
        MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
        Exit Function
    End If
    If xStore.BoundText = "" Then
        MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
        Exit Function
    End If
    
    If xCodeDesca.Caption = "" Then
        MsgBox "·„ Ì „ «œŒ«· «·„Ê—œ"
        Exit Function
    End If
    MYVALID = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
'lblBarCode.Visible = Not CardTable!isbar
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
XISEND_ORDER.Value = IIf(CardTable!ISEND_ORDER, 1, 0)
If xClosed.Value = 1 Then
    If cmd_closed.Caption <> "› Õ „” ‰œ" Then
        cmd_closed.BackColor = &H8080FF
        cmd_closed.Caption = "› Õ „” ‰œ"
        xClosed.Visible = True
    End If
Else
    If cmd_closed.Caption <> "≈€·«ﬁ „” ‰œ" Then
        cmd_closed.BackColor = &H8000000F
        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
        xClosed.Visible = False
    End If
End If


xDoc_No.text = CardTable!doc_no
xInv_No.text = CardTable!INV_NO & ""
Xremark.text = CardTable!remark & ""
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")

xStore.BoundText = CardTable!STORE & ""
xMosm.BoundText = CardTable!MOSM & ""
If myPublic = 2 Or myPublic = 3 Then
    xStore.Locked = True
End If

xCode.text = CardTable!CODE & ""


xCodeDesca.Caption = CardTable!CodeDesca & ""
xSupp.Value = IIf(CardTable!SUPP, 1, 0)
xsup_desca.Caption = IIf(CardTable!SUPP, "„ﬂ » Ã„·…", "„’‰⁄")
xsubcode.Caption = CardTable!SUBCODE & ""

xDiscount.text = CardTable!discount & ""
xTax.text = TurnValue(Val(CardTable!tax & ""), 0, "")
xRateSup.Caption = Val(CardTable!disc & "")

If GetDesca("Select doc_no from addprint where BRANCH = " & MyParn(cBranch) & " AND doc_no = " & MyParn(xDoc_No.text), con) <> "" Then
    Me.StatusBar1.Panels(1).text = "„”Ã· ·Â« »«—ﬂÊœ"
Else
    Me.StatusBar1.Panels(1).text = ""
End If
'xDiscount.Text = TurnValue(Val(CardTable!DISCOUNT & ""), 0, "")
myloadgrd
myloadgrd2
grid1.Row = 1
grid1.Col = 0
grid1_EnterCell
End Sub
Public Sub myloadgrd()
Dim cString As String
With grid1
'                           0               1           2                   3           4               5           6           7           8           9       10              11          12          13                 14              15               16
    cString = "SELECT FILE6_50.ITEM,file1_10.mosm,FACT.DESCA ,file1_10.modelfact0,file1_10.desca ,file1_10.scal,file1_10.color,Quant,FILE6_50.Price, DISCOUNT, total , FILE1_10.PRICE, COST_INV ,FILE1_10.MODELNO,FILE1_10.barcode , FILE1_10.barcode2 , file1_10sc.desca " & _
          " FROM FILE6_50 INNER JOIN FILE1_10 ON FILE6_50.ITEM = FILE1_10.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE left join file1_10sc on file1_10sc.code = file1_10.[section] "
    cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.text)
    
    Set DATA3.Recordset = myRecordSet(cString, con)
    grid1.AddItem ""
'    xCode.Enabled = .Rows <= 2
    xDoc_No.Enabled = .Rows <= 2
End With
Handlecontrols LoadMode
CalcTotals
Fixgrd
End Sub
Private Sub myDefine()
    xRateSup.Caption = ""
    xClosed.Value = 0
    XISEND_ORDER.Value = 0
    
    xMosm.BoundText = cPMosm
    If cmd_closed.Caption <> "≈€·«ﬁ „” ‰œ" Then
        cmd_closed.BackColor = &H8000000F
        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
        xClosed.Visible = False
    End If

    Xremark.text = ""
    cmdAddItems.Enabled = True
    xsubcode.Caption = ""
    xClosed.Visible = False
    xClosed.Value = ssCBUnchecked
    xStore.BoundText = ""
    xCode.Enabled = True
    xInv_No.text = ""
    xDate.text = Format(Date, "dd-mm-yyyy")
    xStore.Locked = False
    If myPublic = 0 Or myPublic = 1 Then
        xDoc_No.text = RetZero(Newflag(cFileHeader, "doc_no", con))
    Else
        xDoc_No.text = ""
'        xdoc_no.Text = RetZero(Newflag(cFileHeader, "doc_no"))
    End If

    xCodeDesca.Caption = ""
    xBalance.Caption = ""
    xCode.text = ""
    xDiscount.text = ""

    
    xtotalOrg.Caption = ""
    xDiscountItem.Caption = ""
    xTotalItem.Caption = ""
    xr_Discount.text = ""
    xTotalNoTax.Caption = ""
    xr_tax.text = ""
    xTax.text = ""
    xTotal.Caption = ""
    xtotalQuant.Caption = ""
    xsup_desca.Caption = ""

'    xTotalDis.Caption = ""
    xUserName.Caption = ""
'    xRate.Text = ""

    grid1.Rows = 1
    grid1.AddItem ""
    ImgX1.Images.Clear
    Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
    cmdSave.Enabled = (xClosed.Value = 0)
    CmdDelInv.Enabled = nMode = LoadMode And (xClosed.Value = 0)
    cmdAddItems.Enabled = (xClosed.Value = 0)

    cmdNewInv.Enabled = nMode = LoadMode
    cmd_closed.Enabled = nMode = LoadMode
    
    cmdFirst.Enabled = (nMode = LoadMode)
    cmdLast.Enabled = (nMode = LoadMode)
    cmdNext.Enabled = (nMode = LoadMode)
    cmdPrevious.Enabled = (nMode = LoadMode)
    xClosed.Enabled = (nMode = LoadMode)
    xDoc_No.Enabled = (nMode = DefineMode)
    cmd_closed.Enabled = (nMode = LoadMode)
    xDoc_No.Tag = nMode
    LoadPict xClosed.Value
End Sub
Private Sub xDiscount_Validate(Cancel As Boolean)
    CalcTotals
End Sub
Private Sub xDoc_No_GotFocus()
    myGotFocus xDoc_No
End Sub
Private Sub xDoc_No_LostFocus()
    myLostFocus xDoc_No
    If xDoc_No.text = "" Then Exit Sub
    xDoc_No.text = RetZero(xDoc_No.text)
    myUndo
End Sub
Private Function CalcTotals()
Dim nTotalOrg As Double, nTotalItem As Double, nDiscountItem As Double
Dim nTotalQuant As Double

With grid1
For i = 1 To grid1.Rows - 2
    nTotalOrg = nTotalOrg + (Val(.TextMatrix(i, 7)) * Val(.TextMatrix(i, 8)))
    nTotalQuant = nTotalQuant + Val(.TextMatrix(i, 7))
    nDiscountItem = nDiscountItem + (Val(.TextMatrix(i, 7)) * Val(.TextMatrix(i, 8)) * Val(.TextMatrix(i, 9)) / 100)
Next
End With
xtotalQuant.Caption = Format(nTotalQuant, "#")

xtotalOrg.Caption = Format(nTotalOrg, "#0.00")
xDiscountItem.Caption = Format(nDiscountItem, "#0.00")
xTotalItem.Caption = Format(Val(xtotalOrg.Caption) - Val(xDiscountItem.Caption), "#0.00")

If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xr_Discount.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xr_Discount.text = Myvalue(Round((Val(xDiscount.text) / Val(xTotalItem.Caption)) * 100, nRound))
    End If
Else
    xr_Discount.text = ""
End If

xTotalNoTax.Caption = Format(Val(xTotalItem.Caption) - Val(xDiscount.text), "#0.00")

If Val(xTotalNoTax.Caption) <> 0 Then
    If Round(Val(xr_tax.text), nRound) <> Round(Val(xTax.text) / Val(xTotalNoTax.Caption) * 100, nRound) Then
        xr_tax.text = Myvalue(Round((Val(xTax.text) / Val(xTotalNoTax.Caption)) * 100, nRound))
    End If
Else
    xr_tax.text = ""
End If
xTotal.Caption = Format(Val(xTotalNoTax.Caption) + Val(xTax.text))
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(7, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT     top 200 DOC_NO,[DATE] AS DATE1,DATE, FILE4_10.Desca , inv_no , ' ', STORE_BR.desca , ISEND_ORDER " & _
                  " FROM  (" & cFileHeader & " inner JOIN file4_10 ON " & cFileHeader & ".CODE " & " = FILE4_10.CODE inner join STORE_BR on STORE_BR.code = " & cFileHeader & ".store )"
If cBranch = "00" And Not bOpt5 Then Generalarray(1) = Generalarray(1) & " where store in (select store from USERSHOP where code = " & nusercode & " ) "
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

GrdArray(7, 0) = "„€·ﬁ…"
GrdArray(7, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search1
Search1.Caption = "«” ⁄·«„"
Search1.Show 1
End Sub
Private Sub xInv_No_GotFocus()
myGotFocus xInv_No
End Sub
Private Sub xInv_No_LostFocus()
myLostFocus xInv_No
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xInv_No
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xInv_No
End Sub
Private Sub xRate_LostFocus()
If Val(xRate.text) <> 0 Then
    xTax.text = Format(Val(xTotalDis.Caption) * (Val(xRate.text) / 100), "Fixed")
    CalcTotals
End If
End Sub
Private Sub Fixgrd()
With grid1

'                   0            1          2           3       4           5           6           7           8           9       10              11              12          13              14              15
.FormatString = "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "Œ’„ ’‰›|" & "«·≈Ã„«·Ì|" & "”⁄— „” Â·ﬂ|" & " ﬂ·›…|" & "«·„ÊœÌ·.|" & "»«—ﬂÊœ „’‰⁄|" & "»«—ﬂÊœ „’‰⁄"
.WordWrap = True
.RowHeight(0) = 600
.ColWidth(0) = 1000
.ColWidth(1) = 800
.ColWidth(2) = 2200
.ColWidth(3) = 2000
.ColWidth(4) = 3500
.ColWidth(5) = 700
.ColWidth(6) = 1000
.ColWidth(7) = 600
.ColWidth(8) = 900
.ColWidth(9) = 800
.ColWidth(10) = 1000
.ColWidth(11) = 1000
.ColWidth(12) = 900
.ColWidth(13) = 1000
.ColWidth(14) = 1400
.ColWidth(15) = 900
.ColHidden(16) = True
.ColHidden(13) = True

'.ColHidden(8) = (cBranch <> "00" And cBranch < "60") Or Not bOpt10
'.ColHidden(9) = (cBranch <> "00" And cBranch < "60") Or Not bOpt10
'.ColHidden(10) = (cBranch <> "00" And cBranch < "60") Or Not bOpt10
'.ColHidden(12) = (cBranch <> "00" And cBranch < "60") Or Not bOpt10

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub xr_Discount_Validate(Cancel As Boolean)
If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xr_Discount.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xDiscount.text = Round((Val(xr_Discount.text) * Val(xTotalItem.Caption)) / 100, nRound)
    End If
Else
    xDiscount.text = ""
End If
CalcTotals
End Sub
Private Sub xr_tax_Validate(Cancel As Boolean)
If Val(xTotalNoTax.Caption) <> 0 Then
    If Round(Val(xr_tax.text), nRound) <> Round(Val(xTax.text) / Val(xTotalNoTax.Caption) * 100, nRound) Then
        xTax.text = Round((Val(xr_tax.text) * Val(xTotalNoTax.Caption)) / 100, nRound)
    End If
Else
    xTax.text = ""
End If
CalcTotals
End Sub
Private Sub xRateSup_Click()
xr_Discount.text = Val(xRateSup.Caption)
xr_Discount_Validate False
End Sub
Private Sub xSection_GotFocus()
    myGotFocus xSection
End Sub
Private Sub xSection_LostFocus()
    myLostFocus xSection
End Sub
Private Sub xTax_GotFocus()
myGotFocus xTax
End Sub
Private Sub xTax_LostFocus()
myLostFocus xTax
CalcTotals
End Sub
Private Sub doprint()
Dim aHeader(2)
'If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim lPrice As Boolean
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
For i = 1 To grid1.Rows - 2
    If xSection.BoundText = "" Or xSection.text = grid1.TextMatrix(i, 16) Then
        temptable.AddNew
        temptable!str10 = Me.Caption & "  " & xStore.text
        temptable!str6 = xDoc_No.text
        temptable!STR7 = xCodeDesca.Caption
        temptable!str5 = TurnValue(xInv_No.text)
        temptable!Date1 = xDate.text
        temptable!str1 = TurnValue(grid1.TextMatrix(i, 0))
        temptable!str13 = TurnValue(grid1.TextMatrix(i, 3))
        temptable!str2 = TurnValue(grid1.TextMatrix(i, 2), "", Null)
        temptable!str3 = TurnValue(grid1.TextMatrix(i, 4))
        temptable!str4 = TurnValue(grid1.TextMatrix(i, 5))
        temptable!str8 = TurnValue(grid1.TextMatrix(i, 6))
        temptable!val3 = Val(grid1.TextMatrix(i, 7))
        temptable!VAL1 = Val(grid1.TextMatrix(i, 8))
        temptable!VAL4 = Val(grid1.TextMatrix(i, 11))
        temptable!val15 = Val(grid1.TextMatrix(i, 10))
        temptable!Val10 = Val(grid1.TextMatrix(i, 9))
        temptable!val2 = Val(grid1.TextMatrix(i, 12))
        temptable!str15 = TurnValue(grid1.TextMatrix(i, 13))
        temptable!str16 = TurnValue(grid1.TextMatrix(i, 13))
        temptable!str14 = TurnValue(grid1.TextMatrix(i, 16))
        
        temptable!str16 = " ÕÊÌ· ≈·Ï : " & xStore.text
        If xSection.BoundText = "" Then
            temptable!val5 = Val(xDiscount.text)
            temptable!VAL6 = Val(xTax.text)
            temptable!val17 = Val(xTotal.Caption)
        End If
        temptable!STR19 = Firsttitle
        temptable!VAL14 = i
        temptable.Update
    End If
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If xprintnoprice.Value = 0 Then
    If Check1.Value = 0 Then
        Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_ORDFER.rpt"
    Else
        Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_ORDFER2.rpt"
    End If
Else
    If Check1.Value = 0 Then
        Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_ORDFER_N.rpt"
    Else
        Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_ORDFER2_N.rpt"
    End If
End If
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Function mysave(Optional bMsg As Boolean = True) As Boolean
If Not MYVALID Then Exit Function
If Not myreplace Then Exit Function
If bMsg Then Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ", , 300
mysave = True
myUndo
End Function
Sub myproc2(nDoc_no)
bNoMsgExit = True
CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
    Unload Me
End If
End Sub
Private Sub cmd_closed_Click()
Dim cString As String
con.BeginTrans
On Error GoTo myerror
cString = " update " & cFileHeader & " set closed = " & IIf(xClosed.Value = 1, "0", "1")
cString = cString & turn(cString) & "doc_no = " & MyParn(xDoc_No.text)
con.Execute cString
con.CommitTrans
myUndo
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub cmd_CLOSEDDATE_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sCaption = "ÿ·»Ì«  "
oClosefrm.nMode = 0
oClosefrm.Show 1
myUndo
End Sub
Private Sub cmd_open_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sCaption = "ÿ·»Ì«  "
oClosefrm.nMode = 1
oClosefrm.Show 1
myUndo
End Sub
Private Sub CardLookupSupp()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(5, 1)
If xCode.text = "" Then Exit Sub
Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,[DATE] AS DATE1,DATE, " & cFileClient & ".Desca , inv_no , mosm " & _
                  " FROM  (" & cFileHeader & " left JOIN " & cFileClient & " ON " & cFileHeader & ".CODE " & " = " & cFileClient & ".CODE ) WHERE FILE4_10.CODE = " & MyParn(xCode.text)

Generalarray(2) = "Order by Date"
Generalarray(3) = 6000
Generalarray(5) = True


listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
listarray(0, 1) = "(inv_no Like '%cFilter%' or Doc_No Like '%cFilter%' or  " & cFileClient & ".DESCA LIKE '%cFilter%' OR " & _
                  "##date##)"


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
GrdArray(5, 1) = 1000
searchArray = Array(Generalarray, listarray, GrdArray)
Search1.Caption = "«” ⁄·«„"
Search1.Show 1
End Sub
Private Function myreplaceGrd() As Boolean
Dim loctable As ADODB.Recordset, cString As String, nItem As Double, cQDem As String
cString = " DELETE FROM " & cFile & " FROM   FILE1_10 INNER JOIN " & cFile & " ON FILE1_10.ITEM = " & cFile & ".ITEM"
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.text)
cString = cString & turn(cString) & "FILE1_10.MODEL = " & MyParn(addModelOrder.XMODEL.Caption)
con.Execute cString, nDel
With addModelOrder.grid1
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                nItem = addModelOrder.grid2.TextMatrix(nRow, nCol)
                Set loctable = ItemFind(addModelOrder.grid2.TextMatrix(nRow, nCol), con)
                If Not (loctable.EOF And loctable.BOF) Then
                cString = "Insert into " & cFile & " (doc_no,item,MODEL,REMARK,discount,price,Quant)" & _
                           "Values(" & _
                           addstring(xDoc_No.text) & "," & _
                           addvalue(loctable!Item) & "," & _
                           addstring(loctable!MODEL) & "," & _
                           addstring(addModelOrder.Xremark.text) & "," & _
                           Val(addModelOrder.xDiscount.text) & "," & _
                           Val(loctable!cost & "") & "," & _
                           Val(.TextMatrix(nRow, nCol)) & _
                           ")"
                    con.Execute cString
                End If
                Set loctable = Nothing
            End If
        Next
    Next
End With
End Function
Sub myProcAddModel()
If Not MYVALID Then Exit Sub
If myreplace(True) Then
    bUpdated = True
    Inform " „  «÷«›… «·„ÊœÌ·«  »‰Ã«Õ", , 500
    If xDoc_No.Tag = DefineMode Then
        myUndo
    Else
        myloadgrd
    End If
End If
grid1.ShowCell grid1.Rows - 1, 0
grid1.Select grid1.Rows - 1, 0
End Sub
Private Sub LoadPict(pValue As Long)
On Error Resume Next
cmd_closed.Picture = LoadPicture(App.Path & "\SYSIMG\" & IIf(pValue = 0, "CLOSE_S.jpg", "OPEN_S.jpg"))
cmd_closed.PictureDisabled = LoadPicture(App.Path & "\SYSIMG\" & IIf(pValue = 0, "CLOSE_S_D.jpg", "OPEN_S_D.jpg"))
Err.Clear
End Sub
Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
ImgX1.Images.Clear
ImgX1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub
Private Sub myAddModels(Optional bAddModel As Boolean = False)
If Not MYVALID Then Exit Sub
Set addModelOrder.myForm = Me
addModelOrder.nColItem = 0
addModelOrder.nColQuant = 7
addModelOrder.SMosm = xMosm.BoundText
addModelOrder.sStore = xStore.BoundText
addModelOrder.bshowBal = True
'addModelOrder.sSection = loctable!Section & ""
addModelOrder.SModelFact = ""
addModelOrder.sModel = ""


addModelOrder.Sdiscount = Val(grid1.TextMatrix(grid1.Row, 9))
If xSupp.Value = 1 Then addModelOrder.sSupp = xCode.text Else addModelOrder.SFact = xsubcode.Caption
If IsNumeric(grid1.TextMatrix(grid1.Row, 0)) And bAddModel Then
'    Dim loctable As ADODB.Recordset
    Set loctable = ItemFind(grid1.TextMatrix(grid1.Row, 0), con)
    If Not (loctable.EOF And loctable.BOF) Then
        addModelOrder.SModelFact = loctable!modelfact0 & ""
        addModelOrder.SFact = loctable!Fact & ""
        addModelOrder.sSupp = loctable!SUPP & ""
        addModelOrder.SMosm = loctable!MOSM & ""
        addModelOrder.sModel = loctable!MODEL & ""
        addModelOrder.sSection = loctable!Section & ""
    End If
    loctable.Close
    Set loctable = Nothing
End If
bUpdated = False
addModelOrder.Show 1
If bUpdated Then
    If Not MYVALID Then Exit Sub
    xr_Discount.text = Val(xRateSup.Caption)
    xr_Discount_Validate False
    myreplace
End If
End Sub
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
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
    If OldRow <> NewRow Then
        myLoadPicture .TextMatrix(NewRow, 13)
    End If
End With

End Sub
Private Sub myUndo()
If IsNumeric(xDoc_No.text) Then
    openCardTable xDoc_No.text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then myDefine Else myload
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function openCardTable(Optional pDoc_No As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset

cString = "SELECT TOP 1 " & cFileHeader & ".* , FILE4_10.disc , FILE4_10.DESCA AS CodeDesca , FILE4_10.SUPP , FILE4_10.SUBCODE from " & cFileHeader & " inner join file4_10 on file4_10.code = " & cFileHeader & ".code "
If cBranch = "00" And Not bOpt5 Then
    cString = cString
End If
If pSign = "=" Then
    If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addstring(pDoc_No)
Else
    If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addstring(pDoc_No)
End If
If Not bOpt5 And cBranch = "00" Then cFilter = cFilter & turn(cFilter, " and ") & " STORE IN (SELECT STORE FROM USERSHOP WHERE CODE =  " & nusercode & " ) "
' «·«Œ Ì«—« 
If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by doc_no desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by doc_no ASC"
End If
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function
Private Sub CmdNext_Click()
openCardTable xDoc_No.text, ">"
If CardTable.EOF Then openCardTable xDoc_No.text
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xDoc_No.text, "<"
If CardTable.EOF Then openCardTable xDoc_No.text
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub
Public Sub myloadgrd2()
Dim cString As String
With grid2
'                           0               1           2                   3           4               5           6           7           8           9       10              11          12          13                 14              15               16
    cString = "SELECT FILE1_10.barcode ,FILE1_10.barcode2 ,FILE6_50.Quant  " & _
          " FROM FILE6_50 INNER JOIN FILE1_10 ON FILE6_50.ITEM = FILE1_10.ITEM "
    cString = cString & " WHERE DOC_NO = " & MyParn(xDoc_No.text)
    If xSection.BoundText <> "" Then cString = cString & " AND FILE1_10.[SECTION] = " & Val(xSection.BoundText)
    Set data22.Recordset = myRecordSet(cString, con)
    grid2.AddItem ""
End With
Fixgrd2
End Sub
Private Sub Fixgrd2()
With grid2
'                   0            1          2
.FormatString = "»«—ﬂÊœ „’‰⁄|" & "»«—ﬂÊœ J|" & "«·ﬂ„Ì…|"

.ColWidth(0) = 2000
.ColWidth(1) = 2000
.ColWidth(2) = 2000
End With
End Sub

