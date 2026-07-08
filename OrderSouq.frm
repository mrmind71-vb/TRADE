VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form OrderSouq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·»Ì«  ⁄„·«¡"
   ClientHeight    =   9990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15315
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
   ScaleHeight     =   9990
   ScaleWidth      =   15315
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_ex2 
      BackColor       =   &H00FFFFC0&
      Caption         =   " «” Ì—«œ „‰ «·«ﬂ”Ì·  ⁄„·«¡ Ã„·…"
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
      Left            =   12015
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   90
      Width           =   1995
   End
   Begin VSFlex7Ctl.VSFlexGrid GRID3 
      Height          =   420
      Left            =   19485
      TabIndex        =   48
      Top             =   720
      Visible         =   0   'False
      Width           =   1995
      _cx             =   3519
      _cy             =   741
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
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
      RightToLeft     =   0   'False
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
   Begin VB.CommandButton CMD_TOEXECL 
      BackColor       =   &H00FFFFC0&
      Caption         =   " ÕÊÌ· «ﬂ”Ì·"
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
      Left            =   19485
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   1170
      Width           =   1995
   End
   Begin VB.Frame Frame11 
      Height          =   3210
      Left            =   14085
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   8730
      Width           =   5325
      Begin ImgXCtrl6.ImgXCtrl imgx1 
         Height          =   2940
         Left            =   90
         TabIndex        =   44
         Top             =   180
         Width           =   5145
         _ExtentX        =   9075
         _ExtentY        =   5186
         BackColor       =   16777215
         BorderStyle     =   0
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "amr123"
         LicenseRegCode  =   "íß“µª∫≥ºΩ∫æ“±æß´¥Ω∏•OANH-FAZOHIZB-EFTP6gI"
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   5640
      Left            =   19485
      TabIndex        =   42
      Top             =   1665
      Width           =   1995
      _cx             =   3519
      _cy             =   9948
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
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
      RightToLeft     =   0   'False
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
   Begin VB.Frame Frame7 
      Caption         =   " —ÕÌ· «·ÿ·»Ì… ≈·Ï „” ‰œ „»Ì⁄« "
      Height          =   960
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   630
      Visible         =   0   'False
      Width           =   3345
      Begin VB.TextBox xdoc_sales 
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
         Left            =   45
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox xdate_sales 
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
         Left            =   45
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„” ‰œ «·„»Ì⁄«  "
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
         Left            =   1500
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   495
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «· —ÕÌ· "
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
         Left            =   1545
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   90
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmd_ex 
      BackColor       =   &H00FFFFC0&
      Caption         =   " «” Ì—«œ „‰ «·«ﬂ”Ì·  AMAZON"
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
      Left            =   9900
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   90
      Width           =   1995
   End
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   -45
      Width           =   6180
      Begin VB.CommandButton cmd_TO_XLS 
         BackColor       =   &H00FFFFC0&
         Caption         =   " ÕÊÌ· «·ÿ·»Ì… ≈·Ï «ﬂ”Ì· ··„Ê—œÌ‰"
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
         Left            =   3195
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   135
         Width           =   2895
      End
      Begin Threed.SSCommand CMDPRINT 
         Height          =   465
         Left            =   45
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   135
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   820
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
         Picture         =   "OrderSouq.frx":0000
         Caption         =   "ÿ»«⁄… ÿ·»Ì«  ··„Ê—œÌ‰ ·ÿ·»Ì«  «·ÌÊ„"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame14 
      Height          =   870
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   9360
      Width           =   4155
      Begin VB.Label XTIME2 
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   495
         Width           =   2220
      End
      Begin VB.Label xUserName2 
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   495
         Width           =   1815
      End
      Begin VB.Label XTIME 
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   135
         Width           =   2220
      End
      Begin VB.Label XUSERNAME 
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   135
         Width           =   1815
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1680
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   0
      Width           =   2085
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
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   90
         Visible         =   0   'False
         Width           =   1410
      End
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   90
         TabIndex        =   21
         Top             =   1125
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   900
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
         Picture         =   "OrderSouq.frx":240F
         Alignment       =   4
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   915
         Left            =   1035
         TabIndex        =   22
         Top             =   180
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1614
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
         Picture         =   "OrderSouq.frx":49DB
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   915
         Left            =   90
         TabIndex        =   23
         Top             =   180
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1614
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
         Picture         =   "OrderSouq.frx":7053
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   8685
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   15
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
         Picture         =   "OrderSouq.frx":9694
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "OrderSouq.frx":B864
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   16
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
         Picture         =   "OrderSouq.frx":D9AC
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "OrderSouq.frx":FB74
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   17
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
         Picture         =   "OrderSouq.frx":11CC3
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "OrderSouq.frx":13EA3
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   18
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
         Picture         =   "OrderSouq.frx":15FFE
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "OrderSouq.frx":181BA
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   630
      Width           =   1365
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   375
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "OrderSouq.frx":1A309
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   585
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
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "OrderSouq.frx":1C882
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   14085
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   5325
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "OrderSouq.frx":1EBE5
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   510
         Left            =   1320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "OrderSouq.frx":21051
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   510
         Left            =   2655
         MaskColor       =   &H00FFFFFF&
         Picture         =   "OrderSouq.frx":238EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   3960
         Picture         =   "OrderSouq.frx":25E97
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   7065
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   630
      Width           =   12390
      Begin VB.CheckBox xisPosted 
         Alignment       =   1  'Right Justify
         Caption         =   "ÿ·»Ì…  „ ≈€·««ﬁÂ«"
         Height          =   240
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   32
         Top             =   540
         Width           =   6045
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
         Left            =   9090
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1950
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
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xcode 
         Height          =   330
         Left            =   7470
         TabIndex        =   50
         Top             =   540
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "«·⁄„Ì· :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   495
         Width           =   930
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   45
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   90
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·»Ì«‰ : "
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
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   495
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   135
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -300
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "data1"
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
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   1575
      Top             =   135
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "data1"
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
      Left            =   135
      TabIndex        =   2
      Top             =   1665
      Width           =   19275
      _cx             =   33999
      _cy             =   12330
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   12648384
      ForeColorSel    =   128
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      RightToLeft     =   0   'False
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "data1"
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   19
      Top             =   9570
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   2835
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   420
      Left            =   3465
      TabIndex        =   34
      Top             =   8820
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label xpict 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   15750
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   11610
      Width           =   2715
   End
   Begin VB.Label XBARCODE 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   11655
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   8865
      Width           =   2355
   End
End
Attribute VB_Name = "OrderSouq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Public sDoc_no As String
Dim bAct As Boolean
Dim cFilter As String
Dim con As New ADODB.Connection
Dim oSearchDoc As New Search3, oSearchItem As New Search3
Dim CardTable As ADODB.Recordset
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Sub myProc()
On Error GoTo myerror
xdoc_no.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
myUndo
Unload oSearchDoc
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Unload Search
End Sub


Private Sub CMD_EX2_Click()
    'On Error GoTo myerror
    AddFrom_Order2
    myUndo
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear

End Sub


Private Sub cmd_TO_XLS_Click()
    Dim aRow As Variant
'    On Error GoTo myerror
    With grid1
    .ColHidden(10) = True
    .ColHidden(11) = True
    .ColHidden(12) = True
    .ColHidden(13) = True
    .ColHidden(14) = True
    .ColHidden(1) = False
    
'    .ColHidden(18) = False
'    .ColHidden(17) = False
'    .ColHidden(16) = False
    
    .ColHidden(.Cols - 2) = False
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
'    .ColHidden(18) = True
'    .ColHidden(17) = True
'    .ColHidden(16) = True
    .ColHidden(10) = False
    .ColHidden(11) = False
    .ColHidden(12) = False
    .ColHidden(13) = False
    .ColHidden(14) = False
    .ColHidden(1) = True
    .ColHidden(.Cols - 2) = True
    
    End With
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub CMD_TOEXECL_Click()
    On Error GoTo myerror
    ToFileExel2 GRID2, , , , , 1.1, , , , , , Me
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub cmdDelinv_Click()
On erorr GoTo myerror
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    con.BeginTrans
    On Error GoTo myerror
    con.Execute " Delete  From FILE6_51 where Doc_No = " & MyParn(xdoc_no.text)
    con.Execute " Delete  From FILE6_51H where Doc_No = " & MyParn(xdoc_no.text)
    con.CommitTrans
    
    AddLod_Data cusername, 2, " Õ–› " & Me.Caption, con, xdoc_no.text, xDate.text, , xCode.text
    
    myUndo
    
    Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
    myDefine
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub CmdExit_Click()
     Unload Me
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(6, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),file3_10.desca , DOC_SALES , DATE_SALES , ISPOSTED " & _
                  " FROM FILE6_51H left join file3_10 on file3_10.code = file6_51H.code "
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If

Generalarray(2) = "Order by Date DESC , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·⁄„Ì·"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "—ﬁ„ „” ‰œ „»Ì⁄«  "
GrdArray(4, 1) = 1000

GrdArray(5, 0) = " «—ÌŒ „” ‰œ „»Ì⁄«  "
GrdArray(5, 1) = 1500

GrdArray(6, 0) = " „€·ﬁ"
GrdArray(6, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
Load oSearchDoc
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub CmdNewInv_Click()
    myDefine
    On Error Resume Next
    xStore1.SetFocus
    Err.Clear
End Sub
Private Sub cmdPrint_Click()
    doprint
End Sub
Private Sub cmdSave_Click()

con.Execute " UPDATE FILE6_51H SET CODE = " & addstring(xCode.BoundText) & " , DATE = " & addDate(xDate.text) & " ,  isPosted = " & IIf(xisPosted.Value, 1, 0) & " WHERE DOC_NO = " & MyParn(xdoc_no.text)

AddLod_Data cusername, 1, " Õ›Ÿ " & Me.Caption, con, xdoc_no.text, xDate.text, , xCode.text

''foundOther
'If Not MYVALID Then Exit Sub
'If Not myreplace Then Exit Sub
'cString = " update FILE6_51h set ISclosed = 1 where doc_no = " & MyParn(xDoc_No.Text)
'con.Execute cString
'
'cString = " update FILE6_51 set store_2 = " & addstring(xStore2.BoundText) & " where doc_no = " & MyParn(xDoc_No.Text)
'con.Execute cString
'
'Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
'openCardTable
'myUndo
End Sub
Private Sub CmdUndo_Click()
    myUndo
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    'KeyCode = 0
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Me.Caption = sCaption
openCon con


Set DATA2.Recordset = myRecordSet("SELECT * FROM FILE3_10 order by desca  ", con)
Set xCode.RowSource = DATA2
xCode.ListField = "Desca"
xCode.BoundColumn = "CODE"

'Frame9.Visible = (cBranch = "00")
cmdSave.Visible = (cBranch = "00")
CmdDelInv.Visible = (cBranch = "00")
cmdNewInv.Visible = (cBranch = "00")

Frame4.Visible = bedit
cmd_open.Visible = bopt1
cmd_CLOSEDDATE.Visible = bopt1

CmdInform.Visible = bopt2
Frame6.Visible = bopt2

'Set CardTable = New ADODB.Recordset
'data1.ConnectionString = strCon

Set grid1.DataSource = data10

myUndo
myDefine
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload oSearchItem
Unload oSearchDoc
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
'cString = " update FILE6_51h set ISclosed = 1 where doc_no = " & MyParn(xDoc_No.Text)
'con.Execute cString

CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
Set Transfrm = Nothing
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If xClosed.Value = 0 And xdoc_sales.text = "" Then
        With grid1
            con.Execute " UPDATE FILE6_51 SET QUANT_SALES = " & Val(.TextMatrix(Row, 14)) & " WHERE ID = " & Val(.TextMatrix(Row, .Cols - 1))
        End With
    End If
End Sub

Private Sub xdate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.text)
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then CmdInform_Click
End Sub
Private Function MYVALID(Optional bCheckRows As Boolean = True) As Boolean
If xdoc_no.text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

MYVALID = True
End Function
Private Sub myload()
xClosed.Value = IIf(CardTable!ISCLOSED, 1, 0)
If xClosed.Value = 1 Then
    If cmd_closed.Caption <> "› Õ „” ‰œ" Then
        cmd_closed.BackColor = &H8080FF
        cmd_closed.Caption = "› Õ „” ‰œ"
    End If
Else
    If cmd_closed.Caption <> "≈€·«ﬁ „” ‰œ" Then
        cmd_closed.BackColor = &H8000000F
        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
    End If
End If

xisPosted.Value = IIf(CardTable!isPosted, 1, 0)
xusername.Caption = CardTable!UserName & ""
xUserName2.Caption = CardTable!USERNAME2 & ""
xtime.Caption = CardTable!Time & ""
XTIME2.Caption = CardTable!TIME2 & ""

xdoc_no.text = CardTable!doc_no
xCode.BoundText = CardTable!CODE & ""
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
xNotes.text = CardTable!NOTES & ""
xdoc_no.Enabled = False

xdoc_sales.text = CardTable!doc_sales & ""
xdate_sales.text = CardTable!DATE_sales & ""

myloadgrd

Handlecontrols LoadMode
StatusBar1.Panels(1).text = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
StatusBar1.Panels(2).text = "⁄œœ «·”Ã·«  : " & CardTable.RecordCount
grid1.Select grid1.Rows - 1, 0
grid1.ShowCell grid1.Rows - 1, 0
CalcTotals
End Sub
Private Sub myDefine()

cmd_closed.BackColor = &H8000000F
cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
xClosed.Visible = False
xisPosted.Value = False
xdoc_no.text = IncRec(GetDesca("SELECT MAX(DOC_NO) FROM FILE6_51H WHERE SUBSTRING(DOC_NO,1,1)= '0' ", con))
If xdoc_no.text = "" Then xdoc_no.text = "000001"
xCode.BoundText = ""
xusername.Caption = ""
xUserName2.Caption = ""
xtime.Caption = ""
XTIME2.Caption = ""
xClosed.Value = 0
xdoc_no.Enabled = True
xDate.text = Format(Date, "DD-MM-YYYY")
xNotes.text = ""
StatusBar1.Panels(1).text = ""

xdoc_sales.text = ""
xdate_sales.text = ""


StatusBar1.Panels(3).text = ""
StatusBar1.Panels(4).text = ""

'If cmdSave.Enabled Then
'    StatusBar1.Panels(1).Text = "«÷«›… ”Ã· —ﬁ„ " & (CardTable.RecordCount + 1)
'    StatusBar1.Panels(2).Text = "⁄œœ «·”Ã·«  : " & CardTable.RecordCount
'ElseIf CardTable.EOF And CardTable.BOF Then
'    StatusBar1.Panels(1).Text = ""
'    StatusBar1.Panels(2).Text = "·«  ÊÃœ ”Ã·« "
'End If
grid1.Rows = 1
Handlecontrols DefineMode
Fixgrd

GRID2.Cols = 3
GRID2.Rows = 1
GRID2.ColWidth(0) = 1600
GRID2.ColWidth(1) = 8000
GRID2.ColWidth(2) = 1500

End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode
cmdSave.Enabled = (xClosed.Value = 0) And (xdoc_sales.text = "")
CmdDelInv.Enabled = nMode = LoadMode And (xClosed.Value = 0) And (xdoc_sales.text = "")
cmdPrint.Enabled = True '(xdoc_sales.Text = "")

cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And sDoc_no = ""
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And sDoc_no = ""
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2 And sDoc_no = ""
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2 And sDoc_no = ""
xdoc_no.Enabled = (nMode = DefineMode)
xdoc_no.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
If xdoc_no.text = "" Then Exit Sub
xdoc_no.text = RetZero(xdoc_no.text, 6)
If CardTable.BOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function CalcTotals()
Dim nTotalQuant As Double
With grid1
For i = 1 To grid1.Rows - 1
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 9))
Next
StatusBar1.Panels(3) = ""
StatusBar1.Panels(4) = ""
If grid1.Rows > 1 Then StatusBar1.Panels(3) = "⁄œœ «·√’‰«› : " & grid1.Rows - 1
If nTotalQuant > 0 Then StatusBar1.Panels(4) = "≈Ã„«·Ì «·ﬂ„Ì… : " & nTotalQuant
End With
End Function
Private Sub doprint()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
cStr1 = " SELECT FILE4_10.CODE, FILE4_10.DESCA, FACT.DESCA AS fact_name, FILE1_10.MODELFACT0, FILE1_10.desca AS desca_item, FILE1_10.COLOR, FILE1_10.SCAL, FILE1_10.C_SCAL, FILE6_51.QUANT, FILE1_10.COST , (SELECT SUM([IN]-[OUT]) FROM FILE1_11 WHERE FILE6_51.ITEM = FILE1_11.ITEM ) AS BAL , FILE1_10.ENA " & _
        " FROM   FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE INNER JOIN FILE6_51 ON FILE1_10.ITEM = FILE6_51.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE INNER JOIN FILE6_51H ON FILE6_51.DOC_NO = FILE6_51H.DOC_NO where FILE6_51h.DOC_NO = " & MyParn(xdoc_no.text)
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
Do While Not .EOF
    If Val(!Quant - Val(!bal & "")) > 0 Then
        temptable.AddNew
        temptable!str6 = !Desca
        temptable!str3 = !CODE
        temptable!Date1 = xDate.text
        temptable!str4 = !fact_name
        temptable!str10 = !modelfact0
        temptable!str2 = !Desca
        temptable!str9 = !ENA
        temptable!str8 = xdoc_no.text
        
        temptable!str2 = !desca_item
        temptable!val2 = !cost
        temptable!STR7 = !color
        temptable!str5 = !scal
        temptable!val5 = !C_SCAL
        temptable!VAL1 = !Quant
        temptable.Update
    End If
    .MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\Order_Supler.rpt"
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub Fixgrd()
With grid1

'                  0            1           2             3         4         5             6                   7         8         9         10             11         12
.FormatString = "Supler |" & "PO No|" & "productId|" & "ASIN|" & "BRAND|" & "SKU|" & "productDescription|" & "Color|" & "Size|" & "Qty|" & "Price PO.|" & "Total|" & "PRICE|"

.Cols = 15
.WordWrap = True
.RowHeight(0) = 800

.ColWidth(0) = 2300
.ColWidth(1) = 1500
.ColWidth(2) = 1600
.ColWidth(3) = 1600
.ColWidth(4) = 2000
.ColWidth(5) = 1500
.ColWidth(6) = 4000
.ColWidth(7) = 1200
.ColWidth(8) = 1200
.ColWidth(9) = 700
.ColWidth(10) = 1000
.ColWidth(11) = 1000
.ColWidth(12) = 1200
.ColWidth(13) = 700

.ColHidden(1) = True

.ColHidden(.Cols - 1) = True
.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 1

For nRow = 1 To .Rows - 1
    nRate = 0
    nDisc = Val(.TextMatrix(nRow, 12)) - Val(.TextMatrix(nRow, 10))
    If Val(.TextMatrix(nRow, 12)) > 0 Then nRate = Round((nDisc / Val(.TextMatrix(nRow, 12))) * 100, 2)
    .TextMatrix(nRow, 13) = nRate
Next nRow
End With
End Sub
Private Sub myloadgrd()
Dim cString As String
'                    0                      1               2               3           4                 5                     6             7               8               9                   10                  11           12           13                          14
cString = "SELECT FILE4_10.DESCA , FILE6_51.Doc_No ,  file1_10.BARCODE,file1_10.ENA, FACT.desca, FILE1_10.MODELFACT0  , FILE1_10.DESCA2 , file1_10.COLOR , file1_10.SCAL  , FILE6_51.QUANT , FILE6_51.PRICE ,  FILE6_51.TOTAL , ROUND((file1_10.PRICE / 1.14),2),0 , FILE6_51.ID " & _
          " from (FILE6_51 inner join file1_10 on file1_10.item = FILE6_51.item ) INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT INNER JOIN FILE4_10 ON FILE4_10.CODE = FILE1_10.CODE "
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xdoc_no.text)
cString = cString & " ORDER BY FILE4_10.DESCA , FILE1_10.FACT , FILE1_10.MODEL , FILE1_10.COLOR , FILE1_10.C_SCAL "
Set data10.Recordset = myRecordSet(cString, con)
Fixgrd
End Sub
Private Sub myloadgrd_bal()
Dim cString As String
Dim cFiledBal As String
cFiledBal = "(SELECT SUM([IN]-[OUT]) FROM FILE1_11 WHERE FILE1_11.ITEM = FILE6_51.ITEM)"

    '                   0                   1               2               3               4                   5                   6           7               8               9               10                  11              12                  13                  14                  15                  16              17              18          19                  20
cString = "SELECT FILE4_10.DESCA , FILE6_51.Doc_No ,  file1_10.BARCODE,file1_10.ENA, FACT.desca, FILE1_10.MODELFACT0  , FILE1_10.DESCA2 , file1_10.COLOR , file1_10.SCAL  , FILE6_51.QUANT , FILE6_51.PRICE ,  FILE6_51.TOTAL , file1_10.PRICE , " & cFiledBal & " , FILE6_51.QUANT_SALES , QUANT_Confirmed   ,FILE6_51.ITEM, FILE6_51.PRICE,file1_10.PRICE , file1_10.pict , FILE6_51.ID " & _
          " from (FILE6_51 inner join file1_10 on file1_10.item = FILE6_51.item ) INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT INNER JOIN FILE4_10 ON FILE4_10.CODE = FILE1_10.CODE "
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xdoc_no.text)
cString = cString & " ORDER BY FILE4_10.DESCA , FILE1_10.FACT , FILE1_10.MODEL , FILE1_10.COLOR , FILE1_10.C_SCAL "
Set data10.Recordset = myRecordSet(cString, con)
data10.Refresh
End Sub

Private Sub old_openCardTable()
Set CardTable = New ADODB.Recordset
Dim cString As String
cFilter = ""
cString = "SELECT * FROM FILE6_51H"
If sDoc_no <> "" Then cFilter = cFilter & turn(cFilter, " AND ") & "DOC_NO = " & MyParn(sDoc_no)
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY DATE , DOC_NO"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Function openCardTable(Optional pDoc_No As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset
cFilter = ""
cFilterLook = ""
cString = "SELECT TOP 1 * from FILE6_51H where doc_no is not null "
If sDoc_no <> "" Then
    cFilter = cFilter & turn(cFilter, " AND ") & " DOC_NO = " & MyParn(sDoc_no)
End If
If pSign = "=" Then
    If pDoc_No <> "" Then cWhere = " AND DOC_NO  " & pSign & addstring(pDoc_No)
Else
    If pDoc_No <> "" Then cWhere = " AND DOC_NO  " & pSign & addstring(pDoc_No)
End If
If cFilter <> "" Then cWhere = cWhere & cFilter
If cWhere <> "" Then cString = cString & cWhere
If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by doc_no desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by doc_no ASC"
End If
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function
Private Sub cmd_closed_Click()
Dim cString As String
con.BeginTrans

'AddLod_Data cusername, , " › Õ „” ‰œ ·· ⁄œÌ· " & Me.Caption, con, xdoc_no.text, xDate.text, , xCode.BoundText

On Error GoTo myerror
cString = " update FILE6_51h set ISclosed = " & IIf(xClosed.Value = 1, "0", "1")
cString = cString & turn(cString) & "doc_no = " & MyParn(xdoc_no.text)
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
oClosefrm.sFile = "FILE6_51h"
oClosefrm.sCaption = DocTitle
oClosefrm.sFieldClose = "isclosed"
oClosefrm.nMode = 0
oClosefrm.Show 1

myUndo
End Sub
Private Sub cmd_open_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = "FILE6_51h"
oClosefrm.sCaption = DocTitle
oClosefrm.sFieldClose = "isclosed"
oClosefrm.nMode = 1
oClosefrm.Show 1

myUndo
End Sub
Private Sub cmd_ex_Click()
    On Error GoTo myerror
    AddFrom_Order
    myUndo
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Sub AddFrom_Order()
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
    If MsgBox("≈÷«›… »Ì«‰«  ÿ·»Ì… «„«“Ê‰ „‰ „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(2)
            
            For nRow = 4 To 10000
                If Len(xlsheet.Cells(nRow, 2)) > 1 Then
                    nRec = nRow
                End If
                cDocNo = xlsheet.Cells(nRow, 1)
                nRowDoc = GRID3.FindRow(cDocNo, 0, 0)
                If nRowDoc < 0 And cDocNo <> "" Then
                    GRID3.AddItem ""
                    GRID3.TextMatrix(GRID3.Rows - 1, 0) = cDocNo
                    con.Execute " INSERT INTO FILE6_51H  (DOC_NO,store, DATE , [NOTES])    VALUES     ( " & addstring(cDocNo) & "  , " & addstring(xlsheet.Cells(nRow, 22)) & " , " & addDate(DateValue(xlsheet.Cells(nRow, 25))) & " , " & addstring(xlsheet.Cells(nRow, 19)) & ")"
                    Inform " New Order " & xlsheet.Cells(nRow, 1)
                End If
                If Len(xlsheet.Cells(nRow, 2)) = 0 Then Exit For
            Next nRow
            prog1.Max = nRec
            i = 0
            For nRow = 4 To nRec
                prog1.Visible = True
                prog1.Value = nRow
                If Len(xlsheet.Cells(nRow, 1)) = 0 Then Exit For
                cDoc_No = Trim(xlsheet.Cells(nRow, 1))
                cBarCode = Trim(xlsheet.Cells(nRow, 2))
                nQTY = Val(xlsheet.Cells(nRow, 10))
                nPrice = Val(xlsheet.Cells(nRow, 9))
                nTotal = Val(xlsheet.Cells(nRow, 10)) * Val(xlsheet.Cells(nRow, 9))
                cENA = xlsheet.Cells(nRow, 4)
                cDesca2 = xlsheet.Cells(nRow, 6)
                
                cStr1 = " ; " & cStr1 & " UPDATE FILE1_10 SET DESCA2 = " & addstring(cDesca2) & " , ENA = " & MyParn(cENA) & " WHERE BARCODE_GS1 = " & addstring(cBarCode)
                cItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE_GS1 = " & MyParn(cBarCode), con)
                If cItem <> "" Then
                    If GetDesca("SELECT DOC_NO FROM FILE6_51 WHERE DOC_NO = " & MyParn(cDoc_No) & " AND ITEM = " & Val(cItem), con) = "" Then
                        cStr2 = " ; " & cStr2 & " INSERT INTO FILE6_51 (DOC_NO , ITEM , QUANT , PRICE , ROW ) VALUES     ( " & addstring(cDoc_No) & "  , " & addstring(cItem) & " , " & nQTY & "," & nPrice & " , " & nRow & ")"
                    End If
                Else
                    GRID2.AddItem ""
                    GRID2.TextMatrix(GRID2.Rows - 1, 0) = cBarCode
                    GRID2.TextMatrix(GRID2.Rows - 1, 1) = cDesca2
                    GRID2.TextMatrix(GRID2.Rows - 1, 2) = cENA
                    xbarcode.Caption = cBarCode
'                   Inform " Barcode Not Found " & cBarCode
                End If
                If nCoutnR = 100 Then
                    If cStr1 <> "" Then con.Execute cStr1
                    If cStr2 <> "" Then con.Execute cStr2
                    nCoutnR = 0
                    cStr1 = ""
                    cStr2 = ""
                Else
                    nCoutnR = nCoutnR + 1
                End If
            Next nRow
            If cStr1 <> "" Then con.Execute cStr1
            If cStr2 <> "" Then con.Execute cStr2
        End If
    End If
End Sub


Sub AddFrom_Confirmed()
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet
    Dim xlwbook As Excel.Workbook
    Dim cDoc_No As String
    Dim cBarCode As String
    Dim cModel As String
    Dim cItem As String, cENA As String
    Dim nQTY As Double
    Dim nPrice As Double, cGroup As String, cCodeGroup As String
    Dim nDisc As Double, cDesca2 As String
    Dim nNewItem As String, cCode As String, cSex As String
    Dim lAddTrue As Boolean, cModelAll As String
    prog1.Min = 0
    prog1.Value = 0
    prog1.Visible = True
    If MsgBox(" ÕœÌÀ «·ﬂ„Ì«  «·„Ê—œ… „‰ „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
            For nRow = 2 To 10000
                nRec = nRow
                If Len(xlsheet.Cells(nRow, 1)) = 0 Then Exit For
            Next nRow
            prog1.Max = nRec
            i = 0
            For nRow = 2 To nRec
                prog1.Visible = True
                prog1.Value = nRow
                If Len(xlsheet.Cells(nRow, 1)) = 0 Then Exit For
                cItem = Trim(xlsheet.Cells(nRow, 12))
                cDoc_No = Trim(xlsheet.Cells(nRow, 2))
                nQTY_Confirmed = Val(xlsheet.Cells(nRow, 11))
                If cItem <> "" And nQTY_Confirmed > 0 Then
                    con.Execute " UPDATE FILE6_51 SET QUANT_Confirmed =  " & (nQTY_Confirmed) & " WHERE DOC_NO = " & MyParn(cDoc_No) & " AND ITEM = " & Val(cItem)
                End If
            Next nRow
        End If
        Inform "  „   ÕœÌÀ «·ﬂ„Ì«  „‰ «·„Ê—œ "
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
If validPhoto(retFile(grid1.TextMatrix(grid1.Row, 19))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(grid1.TextMatrix(grid1.Row, 19))
    oMaxPhoto.cDesca = grid1.TextMatrix(grid1.Row, 6)
    oMaxPhoto.Show 1
End If
End Sub
Sub AddFrom_Order2()
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
    If MsgBox("≈÷«›… »Ì«‰«  ÿ·»Ì… ⁄„·«¡ „‰ „·› Excel", vbYesNo) = vbYes Then
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
            If xdoc_no.Enabled Then con.Execute " INSERT INTO FILE6_51H  (DOC_NO, DATE ) VALUES     ( " & addstring(xdoc_no.text) & " , " & addDate(xDate.text) & ")"
            For nRow = 2 To nRec
                prog1.Visible = True
                prog1.Value = nRow
                If Len(xlsheet.Cells(nRow, 1)) = 0 Then Exit For
                cDoc_No = xdoc_no.text
                cBarCode = Trim(xlsheet.Cells(nRow, 1))
                nQTY = Val(xlsheet.Cells(nRow, 2))
                nPrice = Val(xlsheet.Cells(nRow, 3))
                nTotal = Val(xlsheet.Cells(nRow, 4))
                    
                cItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE13 = " & MyParn(cBarCode) & " OR BARCODE_GS1 = " & MyParn(cBarCode), con)
                If cItem <> "" Then
                    cStr2 = " ; " & cStr2 & " INSERT INTO FILE6_51 (DOC_NO , ITEM , QUANT , PRICE , ROW ) VALUES     ( " & addstring(cDoc_No) & "  , " & addstring(cItem) & " , " & nQTY & "," & nPrice & " , " & nRow & ")"
                Else
                    GRID2.AddItem ""
                    GRID2.TextMatrix(GRID2.Rows - 1, 0) = cBarCode
                    xbarcode.Caption = cBarCode
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
        End If
    End If
End Sub

Private Sub myUndo()
If (xdoc_no.text) <> "" Then
    openCardTable xdoc_no.text
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

Private Sub CmdNext_Click()
openCardTable xdoc_no.text, ">"
If CardTable.EOF Then openCardTable xdoc_no.text
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xdoc_no.text, "<"
If CardTable.EOF Then openCardTable xdoc_no.text
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


