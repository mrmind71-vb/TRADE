VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form VsTitemShop_br 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄… «·„ÊœÌ·« "
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15120
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
   ScaleHeight     =   11040
   ScaleWidth      =   15120
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton del_dem 
      Caption         =   "Õ–› ÿ·»Ì«  ⁄·Ï „ÊœÌ·«  »œÊ‰ ÿ·»Ì…"
      Height          =   330
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   3195
      Width           =   2715
   End
   Begin VB.CommandButton CMD_ALL_DEL 
      BackColor       =   &H008080FF&
      Caption         =   "Õ–› ﬂ· ÿ·»Ì« "
      Height          =   330
      Left            =   2790
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   3195
      Width           =   1275
   End
   Begin VB.Frame Frame2 
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
      Height          =   3210
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   -45
      Width           =   5010
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   2895
         Left            =   90
         TabIndex        =   63
         Top             =   225
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   5106
         BackColor       =   16777215
         BorderStyle     =   4
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         AutoZoomType    =   3
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "«Œ›«¡ „«  „ ÿ»«⁄ …"
      Height          =   330
      Left            =   45
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   1980
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Frame Frame6 
      Height          =   1815
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   1350
      Width           =   4200
      Begin VB.CheckBox XALLNEW 
         Alignment       =   1  'Right Justify
         Caption         =   "≈⁄«œ… «—”«· ﬂ· «·ÿ·»Ì« "
         Height          =   240
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   675
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   " Ê“Ì⁄ «·ÿ·»Ì… ··›—Ê⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   1395
         Width           =   1725
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   " ›’Ì·Ï „ﬁ«”« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1035
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo XDATEDEM1 
         Height          =   315
         Left            =   90
         TabIndex        =   47
         Top             =   990
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XDATEDEM2 
         Height          =   315
         Left            =   90
         TabIndex        =   49
         Top             =   1395
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Threed.SSCommand CMD_PRINTORDER 
         Height          =   420
         Left            =   0
         TabIndex        =   60
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   741
         _Version        =   196610
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÿ»«⁄… «·ÿ·»Ì…"
         ButtonStyle     =   2
      End
      Begin Threed.SSCommand CMD_SEND 
         Height          =   420
         Left            =   2115
         TabIndex        =   61
         Top             =   180
         Visible         =   0   'False
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   741
         _Version        =   196610
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " —ÕÌ· «·ÿ·»Ì… ··„’‰⁄"
         ButtonStyle     =   2
      End
   End
   Begin VB.Frame Frame11 
      Height          =   555
      Left            =   4455
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   2475
      Width           =   3570
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   42
         Top             =   135
         Visible         =   0   'False
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame5 
      Height          =   600
      Left            =   9615
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   2475
      Visible         =   0   'False
      Width           =   3795
      Begin VB.Label xCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "⁄œœ «·„ÊœÌ·« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "‰”»… «·„»Ì⁄« "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label xRate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   180
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1230
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   90
      Width           =   4200
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2610
         Picture         =   "VsTitemShop_br.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1455
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   90
         Picture         =   "VsTitemShop_br.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton cmd_Print 
         Height          =   465
         Left            =   1305
         Picture         =   "VsTitemShop_br.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdBarCode 
         Height          =   465
         Left            =   2610
         Picture         =   "VsTitemShop_br.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   675
         Width           =   1455
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   90
         TabIndex        =   35
         Top             =   675
         Width           =   2460
         _ExtentX        =   4339
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
         Picture         =   "VsTitemShop_br.frx":92B1
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " ÕœÌœ „ÊœÌ·« "
      Height          =   600
      Left            =   13410
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2475
      Width           =   6855
      Begin VB.TextBox XBAL0 
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   27
         Top             =   180
         Width           =   600
      End
      Begin VB.TextBox X1 
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
         Left            =   4950
         MaxLength       =   15
         TabIndex        =   25
         Top             =   180
         Width           =   600
      End
      Begin VB.TextBox X2 
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
         Left            =   4230
         MaxLength       =   15
         TabIndex        =   24
         Top             =   180
         Width           =   600
      End
      Begin Threed.SSCommand CMD_RATE 
         Height          =   420
         Left            =   45
         TabIndex        =   28
         Top             =   135
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   741
         _Version        =   196610
         ForeColor       =   64
         Caption         =   "⁄—÷ „ÊœÌ·« "
         ButtonStyle     =   2
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—’Ìœ  «ﬁ· „‰ "
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
         Height          =   195
         Left            =   2505
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "‰”»… „»Ì⁄«  "
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
         Height          =   195
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   270
         Width           =   1065
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2475
      Left            =   9615
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   10650
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
         Left            =   8055
         TabIndex        =   55
         Top             =   2070
         Width           =   1050
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "„ÊœÌ·«  „”„ÊÕ ·Â« »ÿ·»Ì« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   2220
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "„ÊœÌ·«  ·Â« ÿ·»Ì«  ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1860
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "ÿ·»Ì«  ÃœÌœ… ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1860
      End
      Begin VB.TextBox xdate1 
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
         Left            =   7695
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   225
         Width           =   1410
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
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox xBarCode 
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
         Left            =   7560
         MaxLength       =   15
         TabIndex        =   17
         Top             =   1692
         Width           =   1545
      End
      Begin VB.TextBox xModelFact 
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
         Left            =   3195
         MaxLength       =   15
         TabIndex        =   10
         Top             =   952
         Width           =   1545
      End
      Begin VB.TextBox xDesca 
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
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1353
         Width           =   4590
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   5985
         TabIndex        =   3
         Top             =   1329
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   5985
         TabIndex        =   4
         Top             =   966
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   2025
         TabIndex        =   5
         Top             =   180
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   5985
         TabIndex        =   13
         Top             =   603
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   2025
         TabIndex        =   15
         Top             =   570
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Threed.SSCheck XNOSAL 
         Height          =   285
         Left            =   180
         TabIndex        =   29
         Top             =   2070
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   503
         _Version        =   196610
         BackColor       =   12648447
         Caption         =   "«’‰«› ·Ì” ·Â« „»Ì⁄« "
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   1305
         TabIndex        =   43
         Top             =   1710
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Threed.SSCheck XISNOMOSM2 
         Height          =   285
         Left            =   7065
         TabIndex        =   56
         Top             =   2115
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   196610
         BackColor       =   12648447
         Caption         =   "„« ⁄œ«"
      End
      Begin Threed.SSCheck XISSAL 
         Height          =   285
         Left            =   2295
         TabIndex        =   58
         Top             =   2070
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   503
         _Version        =   196610
         BackColor       =   12648447
         Caption         =   "«’‰«› ·Â« „»Ì⁄«  ›ﬁÿ"
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„Ê”„ œ«Œ·Ï :"
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
         Height          =   195
         Left            =   9165
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   2115
         Width           =   1155
      End
      Begin VB.Label xmodelno 
         Alignment       =   1  'Right Justify
         Caption         =   "Label7"
         Height          =   240
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
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
         Index           =   2
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1755
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… :"
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
         Left            =   9165
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "»«—ﬂÊœ :"
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
         Height          =   195
         Left            =   9180
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1770
         Width           =   585
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
         TabIndex        =   16
         Top             =   641
         Width           =   1230
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
         Left            =   9180
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ÊœÌ· :"
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
         Height          =   195
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1057
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ’‰› :"
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
         Index           =   4
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1428
         Width           =   1155
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
         Index           =   1
         Left            =   9180
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1395
         Width           =   1320
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
         Left            =   9180
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   975
         Width           =   1230
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
         Index           =   3
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1230
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10710
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   1485
      Top             =   45
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
      Bindings        =   "VsTitemShop_br.frx":BAAC
      Height          =   6285
      Left            =   45
      TabIndex        =   2
      Top             =   3555
      Width           =   20220
      _cx             =   35666
      _cy             =   11086
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777088
      ForeColorSel    =   4194368
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   2
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   1188
      Top             =   45
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
      Left            =   891
      Top             =   45
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
      Left            =   297
      Top             =   45
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
      Top             =   45
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
   Begin MSAdodcLib.Adodc DATA6 
      Height          =   330
      Left            =   594
      Top             =   45
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
   Begin MSAdodcLib.Adodc DATA8 
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
      Caption         =   "DATA8"
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
   Begin MSAdodcLib.Adodc DATA9 
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
   Begin MSAdodcLib.Adodc DATA10 
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
End
Attribute VB_Name = "VsTitemShop_br"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cSubBranch As String
Dim confact As New ADODB.Connection
Dim oSearch As New Search3
Public cMyModel As String
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset, clistsec As String
Private Sub CMD_EX2_Click()
    With grid1
    .ColHidden(5) = True
    .ColHidden(6) = True
    .ColHidden(7) = True
    .ColHidden(8) = True
    .ColHidden(9) = True
    .ColHidden(11) = True
    .ColHidden(12) = True
    .ColHidden(13) = True
    .ColHidden(14) = True
    .ColHidden(15) = True
    .ColHidden(16) = True
    .ColHidden(17) = True
    .ColHidden(18) = True
    .ColHidden(19) = True
    
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
    
    .ColHidden(5) = False
    .ColHidden(6) = False
    .ColHidden(7) = False
    .ColHidden(8) = False
    .ColHidden(9) = False
    .ColHidden(11) = False
    .ColHidden(12) = False
    .ColHidden(13) = False
    .ColHidden(14) = False
    .ColHidden(15) = False
    .ColHidden(16) = False
    .ColHidden(17) = False
    .ColHidden(18) = False
    .ColHidden(19) = False
    
    End With
End Sub
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    grid1.ColHidden(17) = True
    cHead1 = "«—’œ… Ê ÿ·»Ì…  ﬂ—«— «·„ÊœÌ·«  „‰ «·„’‰⁄ "
    cHead2 = " Õ Ï   «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY")
    Load PrintGrd
    PrintGrd.DOPRINT Me.grid1, 0.85, -2, cHead1, cHead2, , False, False, 8
    PrintGrd.Show 1
    grid1.ColHidden(17) = False
End Sub

Private Sub Cmd_Print2_Click()
'    With grid1
'    .ColHidden(5) = True
'    .ColHidden(6) = True
'    .ColHidden(7) = True
'    .ColHidden(8) = True
'    .ColHidden(9) = True
'    .ColHidden(11) = True
'    .ColHidden(12) = True
'    .ColHidden(13) = True
'    .ColHidden(14) = True
'    .ColHidden(15) = True
'    .ColHidden(16) = True
'    .ColHidden(17) = True
'    .ColHidden(18) = True
'    .ColHidden(19) = True
'
'    Dim cHead1 As String
'    Dim cHead2 As String
'    For nRow = 2 To .Rows - 1
'        If Val(.TextMatrix(nRow, 20)) = 0 Then .RowHidden(nRow) = True
'    Next nRow
'    cHead1 = "«—’œ… Ê ÿ·»Ì…  ﬂ—«— «·„ÊœÌ·«  „‰ «·„’‰⁄ "
'    cHead2 = " Õ Ï   «—ÌŒ " & Format(xdate1.Text, "DD-MM-YYYY")
'    Load PrintGrd
'    PrintGrd.doprint Me.grid1, 1.3, -2, cHead1, cHead2, , False, False, 10
'    PrintGrd.Show 1
'
'
'    .ColHidden(5) = False
'    .ColHidden(6) = False
'    .ColHidden(7) = False
'    .ColHidden(8) = False
'    .ColHidden(9) = False
'    .ColHidden(11) = False
'    .ColHidden(12) = False
'    .ColHidden(13) = False
'    .ColHidden(14) = False
'    .ColHidden(15) = False
'    .ColHidden(16) = False
'    .ColHidden(17) = False
'    .ColHidden(18) = False
'    .ColHidden(19) = False
'    For nRow = 2 To .Rows - 1
'        .RowHidden(nRow) = False
'    Next nRow
'
'    End With
'
End Sub
Private Sub CMD_PRINTORDER_Click()
    If Check1.Value = 0 Then
        doprint1
    Else
        PrintScalDem
    End If
End Sub
Private Sub CMD_RATE_Click()
With grid1
    For nRow = 2 To .Rows - 1
        .RowHidden(nRow) = False
        If X1.text <> "" Then
            If Val(.TextMatrix(nRow, 11)) < Val(X1.text) Then .RowHidden(nRow) = True
        End If
    
        If X2.text <> "" Then
            If Val(.TextMatrix(nRow, 11)) > Val(X2.text) Then .RowHidden(nRow) = True
        End If
    
        If XBAL0.text <> "" Then
            If Val(.TextMatrix(nRow, 10)) > Val(XBAL0.text) Then .RowHidden(nRow) = True
        End If
    Next nRow
End With
End Sub

Private Sub cmdBarCode_Click()
con.BeginTrans
On Error GoTo myerror
If xModelNo.Caption = "" Then Exit Sub
cString = "INSERT INTO ADDPRINT(ITEM,BRANCH,QUANT,ISPRINT)"
cString = cString & _
          " SELECT FILE1_10.ITEM," & addstring(cBranch) & ",SUM([IN]-[OUT]),1 FROM FR1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FR1_11.ITEM WHERE STORE = " & MyParn(xStore.BoundText) & " AND FILE1_10.MODELNO = " & MyParn(xModelNo.Caption)
cString = cString & " GROUP BY FILE1_10.ITEM "
con.Execute cString
con.CommitTrans
Inform " „ «· ÕÊÌ· »‰Ã«Õ"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear

End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    If xStore.BoundText = "" Then Exit Sub
    If Not IsDate(xDate1.text) Then
        xDate1.text = Format(GetDesca("select min(date) from FR1_11", con))
    End If
    myload
End Sub
Private Sub Command1_Click()
    If MsgBox(" ≈Œ›«¡ «·ÿ·»Ì«  «· Ï  „ ÿ»«⁄ Â« ", vbYesNo + vbDefaultButton2) = vbYes Then
        con.Execute " UPDATE model_dem SET ISPRINTED = 1 WHERE ISPRINTED2 = 1 "
        con.Execute " UPDATE model_dem SET ISPRINTED2 = 0 "
    Else
        con.Execute " UPDATE model_dem SET ISPRINTED2 = 0 "
    End If
End Sub

Private Sub Form_Load()
    openCon con
    xDate1.text = Format(GetDesca("select min(date) from FR1_11", con))
    xDate2.text = Format(Date, "dd-mm-yyyy")
    XMOSM2.AddItem ""
    XMOSM2.AddItem "S"
    XMOSM2.AddItem "W"
    XMOSM2.AddItem "M"
    
    If cBranch = "00" Then
'        con.Execute " UPDATE model_dem  SET [DATE] = CONVERT(varchar,model_dem.DATE , 101)   "
'        con.Execute " update model_dem set Branch = brach where brach is not null "
    End If
    
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
    Set XSECTION.RowSource = data1
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
    DATA11.ConnectionString = strCon
    DATA11.RecordSource = "Select DATE   AS DATE From model_dem GROUP BY DATE  order by DATE  DESC "
    Set XDATEDEM1.RowSource = DATA11
    XDATEDEM1.ListField = "DATE"
    XDATEDEM1.BoundColumn = "DATE"
    
    Set XDATEDEM2.RowSource = DATA11
    XDATEDEM2.ListField = "DATE"
    XDATEDEM2.BoundColumn = "DATE"
    
    data10.ConnectionString = strCon
    data10.RecordSource = "Select Code,DescA From branch_FR  order by CODE "
    Set xStore.RowSource = data10
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    
    XALLNEW.Visible = bopt2
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File4_10 order by Desca"
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    DATA3.ConnectionString = strCon
    DATA3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = DATA3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    DATA4.ConnectionString = strCon
    DATA4.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC "
    Set xMosm.RowSource = DATA4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    DATA5.ConnectionString = strCon
    DATA5.RecordSource = "Select code ,desca From fact ORDER BY DESCA "
    Set xFact.RowSource = DATA5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    grid1.Rows = 2
    grid1.Cols = 19
    FIXGRID
    CMD_SEND.Enabled = (cBranch <> "00")

End Sub
Private Sub myload()
    Dim cWhere As String
    Dim cStrH As String, cStore As String
    If xStore.BoundText <> "" Then cStore = " AND FR1_11.STORE = " & MyParn(xStore.BoundText)
    If Check2.Value <> 0 Then
        cFDem1 = " , ( SELECT sum(QUANT)    FROM model_dem WHERE [ISPRINTED] = 0 AND STORE = " & MyParn(xStore.BoundText) & " AND MODELNO = FILE1_10.MODELNO ) "
        cFDem2 = " , ( SELECT TOP 1 DATE    FROM model_dem WHERE [ISPRINTED] = 0 AND STORE = " & MyParn(xStore.BoundText) & " AND MODELNO = FILE1_10.MODELNO ) "
        cFDem3 = " , ( SELECT TOP 1 REMARK  FROM model_dem WHERE [ISPRINTED] = 0 AND STORE = " & MyParn(xStore.BoundText) & " AND MODELNO = FILE1_10.MODELNO ) "
        cFDem4 = " , ( SELECT TOP 1 DATE    FROM FR1_11 AS FR1_11_2 INNER JOIN FILE1_10 AS FILE1_10_2 ON FILE1_10_2.ITEM = FR1_11_2.ITEM WHERE FILE1_10_2.MODELNO = FILE1_10.MODELNO AND FR1_11_2.STORE = " & MyParn(xStore.BoundText) & " AND ( TYPE = '2' OR TYPE = 'T' )  ORDER BY DATE DESC ) "
        CFDEM5 = " , ( SELECT TOP 1 branch  FROM model_dem WHERE [ISPRINTED] = 0 AND STORE = " & MyParn(xStore.BoundText) & " AND MODELNO = FILE1_10.MODELNO ) "
    Else
        cFDem1 = " , ( SELECT sum(QUANT)    FROM model_dem WHERE STORE = " & MyParn(xStore.BoundText) & " AND MODELNO = FILE1_10.MODELNO ) "
        cFDem2 = " , ( SELECT TOP 1 DATE    FROM model_dem WHERE STORE = " & MyParn(xStore.BoundText) & " AND MODELNO = FILE1_10.MODELNO ) "
        cFDem3 = " , ( SELECT TOP 1 REMARK  FROM model_dem WHERE STORE = " & MyParn(xStore.BoundText) & " AND MODELNO = FILE1_10.MODELNO ) "
        cFDem4 = " , ( SELECT TOP 1 DATE    FROM FR1_11 AS FR1_11_2 INNER JOIN FILE1_10 AS FILE1_10_2 ON FILE1_10_2.ITEM = FR1_11_2.ITEM WHERE FILE1_10_2.MODELNO = FILE1_10.MODELNO AND FR1_11_2.STORE = " & MyParn(xStore.BoundText) & " AND ( TYPE = '2' OR TYPE = 'T' ) ORDER BY DATE DESC ) "
        CFDEM5 = " , ( SELECT TOP 1 branch  FROM model_dem WHERE STORE = " & MyParn(xStore.BoundText) & " AND MODELNO = FILE1_10.MODELNO ) "
    End If
    
    cField4 = myiif("FR1_11.date < " & DateSq(xDate1) & cStore, "([IN] - [OUT]) ") & " AS f_sal  , "
    cField5 = myiif("type = 'z' and FR1_11.date >= " & DateSq(xDate1) & cStore, "([in] - [out]) ") & " AS t_comp , "
    cField6 = myiif("(type = '2' OR type = 'F' OR type = 'T'  ) and FR1_11.date >= " & DateSq(xDate1) & cStore, "([IN] ) ") & " AS t_purch, "
    cField7 = myiif("(type = '7' OR type = 'F' OR type = 'T'  ) and FR1_11.date >= " & DateSq(xDate1) & cStore, "(OUT  ) ") & " AS tR_purch, "
    cField8 = myiif("type = '6' and FR1_11.date >= " & DateSq(xDate1) & cStore, "([out] - [in]) ") & " AS t_sal  , "
    '                   0                   1                  2               3            4                              5
    cString = " SELECT  FACT.desca , file1_10.mosm, file1_10.modelfact0 , MODELNO   ,file1_10.desca ,   " & _
            cField4 & cField5 & cField6 & cField7 & cField8 & _
            " sum([in] - out) as bal ,  " & _
            " ' ' as rate , min(file1_10.price) , min(file1_10.cost) , OKAZ , MIN(PRICE2), sum((FR1_11.[in] - FR1_11.out) * file1_10.cost ) as costbal  " & cFDem1 & cFDem2 & cFDem3 & cFDem4 & CFDEM5 & _
            " FROM FILE1_10 INNER JOIN FR1_11 ON FILE1_10.ITEM = FR1_11.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE WHERE FILE1_10.MODELNO IS NOT NULL "

    If XMOSM2.text <> "" Then
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 <> " & MyParn(XMOSM2.text)
    End If
    If xStore.BoundText <> "" Then cStr2 = cStr2 & " AND FR1_11.STORE = " & MyParn(xStore.BoundText)
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If Check5.Value <> 0 Then cStr2 = cStr2 & " AND  ( FILE1_10.[ISNODEM] = 0  AND FILE1_10.[ISSTOP] = 0) "
    If xbarcode.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.item = " & xbarcode.text
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XSECTION.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & XSECTION.BoundText
    If XDESCA.text <> "" Then cStr2 = cStr2 & " AND file1_10.DESCA LIKE ('%" & XDESCA.text & "%')   "
    If IsDate(xDate2.text) Then cStr2 = cStr2 & " AND FR1_11.[DATE] <= " & DateSq(xDate2.text)
    If Check4.Value <> 0 Then
        If Check2.Value = 0 Then
            If xStore.BoundText = "" Then
                cStr2 = cStr2 & " AND FILE1_10.ITEM IN ( SELECT ITEM FROM model_dem ) "
            Else
                cStr2 = cStr2 & " AND FILE1_10.ITEM IN ( SELECT ITEM FROM model_dem WHERE STORE = " & MyParn(xStore.BoundText) & " ) "
            End If
        Else
            If xStore.BoundText = "" Then
                cStr2 = cStr2 & " AND FILE1_10.ITEM IN ( SELECT ITEM FROM model_dem WHERE [ISPRINTED] = 0  ) "
            Else
                cStr2 = cStr2 & " AND FILE1_10.ITEM IN ( SELECT ITEM FROM model_dem WHERE [ISPRINTED] = 0  AND STORE = " & MyParn(xStore.BoundText) & " ) "
            End If
        End If
    End If
    cString = cString & cStr2 & " GROUP BY file1_10.MODELNO  , FILE1_10.REDEM , FACT.desca , file1_10.mosm, file1_10.modelfact0 , file1_10.desca  , file1_10.[section] , OKAZ "
    cStrH = ""
'    If XBAL.Value = -1 Then cStrH = " HAVING      (SUM(FR1_11.[IN] - FR1_11.OUT) <> 0)"
    If XNOSAL.Value = -1 Then
        cStrH = " HAVING      (SUM(FR1_11.[IN] - FR1_11.OUT) <> 0)"
        cStrH = cStrH & " AND SUM(CASE WHEN TYPE = '6' THEN FR1_11.OUT ELSE 0 END )  = 0"
    End If
    If XISSAL.Value = -1 Then
        cStrH = " HAVING      (SUM(FR1_11.[IN] - FR1_11.OUT) <> 0)"
        cStrH = cStrH & " AND SUM(CASE WHEN TYPE = '6' THEN FR1_11.OUT ELSE 0 END )  <>  0"
    End If
    
    DATA6.RecordSource = cString & cStrH
    DATA6.Refresh
    FIXGRID
End Sub
Sub FIXGRID()
    Dim CodeModelTable As New ADODB.Recordset
    Dim nTIn As Double
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 2
    .FrozenCols = 5
    .Cols = 23
    .RowHidden(1) = True
    .RowHeight(0) = 900
    .WordWrap = True
    .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
    .RowHidden(1) = True
    
    .TextMatrix(0, 0) = "«·„’‰⁄"
    .TextMatrix(0, 1) = "„Ê”„"
    .TextMatrix(0, 2) = "—ﬁ„ „ÊœÌ·"
    .TextMatrix(0, 3) = "«·’‰›"
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = "—’Ìœ √Ê· "
    .TextMatrix(0, 6) = " ”ÊÌ… Ã—œ"
    .TextMatrix(0, 7) = "Ê«—œ "
    .TextMatrix(0, 8) = "’«œ— "
    .TextMatrix(0, 9) = "„»Ì⁄« "
    .TextMatrix(0, 10) = "—’Ìœ"
    .TextMatrix(0, 11) = "‰”»… »Ì⁄"
    .TextMatrix(0, 12) = " ”⁄— »Ì⁄"
    
    .TextMatrix(0, 13) = "”⁄— „’‰⁄"
    .TextMatrix(0, 14) = "‰”»… √Êﬂ«“ÊÌ‰"
    .TextMatrix(0, 15) = "”⁄— √Êﬂ«“ÊÌ‰"
    .TextMatrix(0, 16) = " ﬁÌ„… —’Ìœ"
    
    .TextMatrix(0, 17) = "ÿ·»Ì…"
    .TextMatrix(0, 18) = " «—ÌŒ"
    
    .TextMatrix(0, 19) = "„·«ÕŸ« "
        
    .TextMatrix(0, 20) = "«Œ— Ê«—œ"
    .TextMatrix(0, 21) = "≈Œ›«¡"
    .TextMatrix(0, 22) = "›—⁄"
        
    .ColWidth(0) = 1300
    .ColWidth(1) = 500
    .ColWidth(2) = 1400
    .ColWidth(3) = 0
    .ColWidth(4) = 2500
    .ColWidth(5) = 700
    .ColWidth(6) = 700
    .ColWidth(7) = 700
    .ColWidth(8) = 700
    .ColWidth(9) = 700
    .ColWidth(10) = 700
    .ColWidth(11) = 800
    .ColWidth(12) = 800
    .ColWidth(13) = 0
    .ColWidth(14) = 800
    .ColWidth(15) = 800
    .ColWidth(16) = 0
    .ColWidth(17) = 900
    .ColWidth(18) = 1300
    .ColWidth(19) = 0
    .ColWidth(20) = 1300
    .ColWidth(21) = 600
    .ColWidth(22) = 500
    .ColHidden(19) = True
    .ColComboList(17) = "..."
    
    .ColDataType(18) = flexDTDate
    .ColDataType(20) = flexDTDate
    
    .ColDataType(19) = flexDTString
    
    .ColDataType(21) = flexDTBoolean
    
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = 4
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString
    
    If cBranch <> "00" Then
        .ColHidden(13) = True
        .ColHidden(16) = True
    End If
    
    For nCol = 6 To .Cols - 4
        .ColDataType(nCol) = flexDTDouble
    Next nCol
    
    .Editable = flexEDNone
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    .Subtotal flexSTClear
    If .Rows > 2 Then
        For I = 2 To .Rows - 1
            nRate = 0
            nTIn = Val(.TextMatrix(I, 5)) + Val(.TextMatrix(I, 6)) + Val(.TextMatrix(I, 7))
            If nTIn > 0 Then nRat = Format(Val(.TextMatrix(I, 9)) / nTIn * 100, "#0.00")
            .TextMatrix(I, 11) = Format(nRat, "##0.00")
        Next I
    End If
    
    .Subtotal flexSTSum, -1, 5, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 6, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 7, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 8, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 9, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 10, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 16, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 17, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    
    If .Rows > 2 Then
        nRate = 0
        nTIn = Val(.TextMatrix(2, 5)) + Val(.TextMatrix(2, 6)) + Val(.TextMatrix(2, 7))
        If nTIn > 0 Then nRat = Format(Val(.TextMatrix(2, 9)) / nTIn * 100, "#0.00")
        .TextMatrix(2, 11) = Format(nRat, "##0.00")
        XCOUNT.Caption = .Rows - 3
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    Else
        XCOUNT.Caption = ""
    End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
    If .Col = 21 Then
        con.Execute " UPDATE model_dem set [ISPRINTED] =  " & Val(.TextMatrix(.Row, 21)) & " WHERE MODELNO = " & MyParn(xModelNo.Caption), nRec
    End If
End With
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If xStore.BoundText <> "" Then
        cString = "SELECT  (SELECT SUM(QUANT) FROM model_dem WHERE ITEM =   FILE1_10.ITEM AND STORE = " & MyParn(xStore.BoundText) & " ) AS bal , FILE1_10.COLOR,FILE1_10.SCAL" & _
                  " FROM FILE1_10 "
        cString = cString & " where FILE1_10.modelno = " & MyParn(grid1.TextMatrix(grid1.Row, 3))
        
        cString2 = "SELECT  SUM([IN]-[OUT]) AS bal , FILE1_10.COLOR,FILE1_10.SCAL FROM FILE1_10 INNER JOIN FR1_11 ON FR1_11.ITEM = FILE1_10.ITEM "
        cString2 = cString2 & " WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 3)) & " AND STORE = " & MyParn(xStore.BoundText) & _
                " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL "
        grdModel_3.pString = cString
        grdModel_3.pString_BAL = cString2
        grdModel_3.pFieldValue = "bal"
        grdModel_3.pModelNo = grid1.TextMatrix(grid1.Row, 3)
        grdModel_3.pstore = xStore.BoundText
        grdModel_3.Show 1
        grid1.TextMatrix(grid1.Row, 17) = Val(GetDesca("select sum(quant) from model_dem where STORE = " & MyParn(xStore.BoundText) & " AND modelno = " & MyParn(xModelNo.Caption), con) & "")
    End If
End Sub

Private Sub grid1_DblClick()
    If grid1.Col = 2 And cBranch = "00" Then
        If Val(GetDesca("SELECT COUNT (CODE) FROM ModelSupp WHERE MODELNO = " & MyParn(xModelNo.Caption), con) & "") > 1 Then
            FrmListSupp.Show 1
        Else
            cMyModel = GetDesca("SELECT MODELNO FROM ModelSupp WHERE MODELNO = " & MyParn(xModelNo.Caption), con)
        End If
        Load ModelMove_FR
        ModelMove_FR_FR.xDate.text = xDate1.text
        ModelMove_FR.XMODEL.text = cMyModel
        ModelMove_FR.xStore.BoundText = xStore.BoundText
        ModelMove_FR.Show
    End If
    
    If grid1.Col = 2 Then
        If Val(GetDesca("SELECT COUNT (CODE) FROM ModelSupp WHERE MODELNO = " & MyParn(xModelNo.Caption), con) & "") > 1 Then
            FrmListSupp.Show 1
        Else
            cMyModel = GetDesca("SELECT MODELNO FROM ModelSupp WHERE MODELNO = " & MyParn(xModelNo.Caption), con)
        End If
        Load ModelMove_FR
        ModelMove_FR.xDate.text = xDate1.text
        ModelMove_FR.XMODEL.text = cMyModel
        ModelMove_FR.xStore.BoundText = xStore.BoundText
        ModelMove_FR.Show
    End If
    
    If grid1.Col = 10 Then
        cString = "SELECT SUM([IN]-[OUT]) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL" & _
                  " FROM FR1_11 INNER JOIN FILE1_10 ON FR1_11.ITEM = FILE1_10.ITEM "
        If IsDate(xDate1.text) Then cString = cString & turn(cString) & "FR1_11.DATE <= " & DateSq(xDate2.text)
        cString = cString & turn(cString) & "FR1_11.STORE = " & MyParn(xStore.BoundText)
        cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(grid1.TextMatrix(grid1.Row, 3))
        cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL"
        grdModelfrm.pString = cString
        grdModelfrm.pFieldValue = "sumofQuant"
        grdModelfrm.pModelNo = grid1.TextMatrix(grid1.Row, 3)
        grdModelfrm.Show 1
    End If
End Sub
Private Sub grid1_EnterCell()
    xModelNo.Caption = grid1.TextMatrix(grid1.Row, 3)
    LoadPhoto grid1.TextMatrix(grid1.Row, 3)
    grid1.Editable = flexEDNone
    If grid1.Col = 17 Or grid1.Col = 21 Then
        grid1.Editable = flexEDKbdMouse
    Else
        grid1.Editable = flexEDNone
    End If
End Sub

Private Sub ImgX1_Click()
If validPhoto(retFile(grid1.TextMatrix(grid1.Row, 3))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(grid1.TextMatrix(grid1.Row, 3))
    oMaxPhoto.cDesca = grid1.TextMatrix(grid1.Row, 4)
    oMaxPhoto.Show 1
End If
End Sub

Private Sub XSTORE_LostFocus()
    If xStore.BoundText <> "" Then
        DATA11.ConnectionString = strCon
        DATA11.RecordSource = "Select DATE From model_dem WHERE STORE = " & MyParn(xStore.BoundText) & " GROUP BY DATE   order by DATE  DESC "
        Set XDATEDEM1.RowSource = DATA11
        XDATEDEM1.ListField = "DATE"
        XDATEDEM1.BoundColumn = "DATE"
        
        Set XDATEDEM2.RowSource = DATA11
        XDATEDEM2.ListField = "DATE"
        XDATEDEM2.BoundColumn = "DATE"
        
        DATA11.Refresh
    End If
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
Private Sub doprint1()
Dim aHeader(2)
Dim temptable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

With grid1
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 17)) > 0 Then
            temptable.AddNew
            temptable!str15 = .TextMatrix(nRow, 0)
            temptable!str16 = .TextMatrix(nRow, 1)
            temptable!str1 = .TextMatrix(nRow, 2)
            temptable!str2 = .TextMatrix(nRow, 4)
            temptable!str3 = TurnValue(.TextMatrix(nRow, 19))
            temptable!VAL4 = Val(.TextMatrix(nRow, 17))
            If cBranch = "00" Then
                temptable!VAL1 = Val(.TextMatrix(nRow, 13))
                temptable!val5 = Val(.TextMatrix(nRow, 13)) * Val(.TextMatrix(nRow, 17))
            End If
            temptable!val2 = Val(.TextMatrix(nRow, 10))
            temptable!STR7 = "ÿ·»Ì…  » «—ÌŒ " & Format(Date, "DD-MM-YYYY") & "  " & xStore.text
            temptable.Update
        End If
    Next nRow
End With
temptable.Requery
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If cBranch = "00" Then
    Main.REPORT1.ReportFileName = App.Path & "\Reports\R_ORDER2.rpt"
Else
    Main.REPORT1.ReportFileName = App.Path & "\Reports\R_ORDER.rpt"
End If
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub doprint2()
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
    
    cStr1 = "SELECT model_dem.ITEM , model_dem.QUANT , FILE1_10.DESCA , FILE1_10.COLOR , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.MOSM , FILE1_10.MODELFACT0 , FILE1_10.MODEL , FILE1_10.COST , FILE1_10.PRICE , FACT.DESCA AS FACTDESCA FROM model_dem INNER JOIN FILE1_10 ON model_dem.ITEM = FILE1_10.ITEM INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT WHERE FILE1_10.MODELNO IS NOT NULL "
    
    If xStore.BoundText <> "" Then cStr2 = cStr2 & " AND model_dem.STORE = " & MyParn(xStore.BoundText)
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xbarcode.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.item = " & xbarcode.text
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XSECTION.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & XSECTION.BoundText
    If Check2.Value <> 0 Then cStr2 = cStr2 & " AND [ISPRINTED] = 0  "
    If IsDate(XDATEDEM1.text) Then cStr2 = cStr2 & " AND DATE  >= " & DateSq(XDATEDEM1.text)
    If IsDate(XDATEDEM2.text) Then cStr2 = cStr2 & " AND DATE  <= " & DateSq(XDATEDEM2.text)
    
    If XDESCA.text <> "" Then cStr2 = cStr2 & " AND file1_10.DESCA LIKE ('%" & XDESCA.text & "%')   "
    cStr1 = cStr1 & cStr2
    sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText

    With sourcetable
    Do While Not .EOF
        temptable.AddNew
        temptable!str11 = !MODEL
        temptable!str15 = !FACTDESCA
        temptable!str16 = !MOSM
        temptable!str1 = !modelfact0
        temptable!str2 = !DESCA
        
        temptable!str12 = Trim(!scal)
        temptable!str13 = Trim(!color)
        temptable!val12 = !C_SCAL
        
        temptable!VAL4 = !Quant
        If cBranch = "00" Then
            temptable!VAL1 = !cost
            temptable!val5 = !cost * !Quant
        Else
            temptable!VAL1 = !price
        End If
        temptable!STR7 = "ÿ·»Ì…  ··› —… „‰ " & Format(XDATEDEM1.text, "DD-MM-YYYY") & " ≈·Ï " & Format(XDATEDEM2.text, "DD-MM-YYYY") & " " & xStore.text
        temptable.Update
        
        If xStore.BoundText <> "" Then
            con.Execute " UPDATE model_dem SET ISPRINTED2 = 1 WHERE ITEM   = " & !Item & " AND  STORE = " & MyParn(xStore.BoundText)
        End If
        
        .MoveNext
    Loop
    End With

If Check3.Value <> 0 And xStore.BoundText = "" Then
    cStr1 = "SELECT model_dem.STORE , FILE0_40.DESCA AS STORENAME , model_dem.ITEM , model_dem.QUANT , FILE1_10.DESCA , FILE1_10.COLOR , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.MOSM , FILE1_10.MODELFACT0 , FILE1_10.MODEL , FILE1_10.COST , FILE1_10.PRICE , FACT.DESCA AS FACTDESCA FROM model_dem INNER JOIN FILE1_10 ON model_dem.ITEM = FILE1_10.ITEM INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT INNER JOIN FILE0_40 ON FILE0_40.CODE = model_dem.STORE WHERE FILE1_10.MODELNO IS NOT NULL "
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xbarcode.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.item = " & xbarcode.text
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XSECTION.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & XSECTION.BoundText
    If Check2.Value <> 0 Then cStr2 = cStr2 & " AND [ISPRINTED] = 0  "
    If IsDate(XDATEDEM1.text) Then cStr2 = cStr2 & " AND DATE  >= " & DateSq(XDATEDEM1.text)
    If IsDate(XDATEDEM2.text) Then cStr2 = cStr2 & " AND DATE  <= " & DateSq(XDATEDEM2.text)
    
    If XDESCA.text <> "" Then cStr2 = cStr2 & " AND file1_10.DESCA LIKE ('%" & XDESCA.text & "%')   "
    cStr1 = cStr1 & cStr2
    sourcetable.Close
    
    sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText

    With sourcetable
    Do While Not .EOF
        temptable.AddNew
        temptable!str14 = !STORE
        temptable!str17 = !STORENAME
        temptable!str11 = !MODEL
        temptable!str15 = !FACTDESCA
        temptable!str16 = !MOSM
        temptable!str1 = !modelfact0
        temptable!str2 = !DESCA
        
        temptable!str12 = Trim(!scal)
        temptable!str13 = Trim(!color)
        temptable!val12 = !C_SCAL
        
        temptable!VAL4 = !Quant
        If cBranch = "00" Then
            temptable!VAL1 = !cost
            temptable!val5 = !cost * !Quant
        Else
            temptable!VAL1 = !price
        End If
        temptable!STR7 = "ÿ·»Ì…  ··› —… „‰ " & Format(XDATEDEM1.text, "DD-MM-YYYY") & " ≈·Ï " & Format(XDATEDEM2.text, "DD-MM-YYYY") & " " & xStore.text
        temptable.Update
        
        If xStore.BoundText = "" Then
            con.Execute " UPDATE model_dem SET ISPRINTED2 = 1 WHERE ITEM   = " & !Item & " AND STORE = " & MyParn(!STORE)
        End If
        .MoveNext
    Loop
    End With

End If

temptable.Requery
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\R_ORDERS.rpt"
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub

Private Sub CMD_SEND_Click()
'On Error GoTo myerror

cSubBranch = GetDesca("SELECT BRANCH FROM ADDRESS", con)
strConfact = LoadConStringsFACT
openConFACT confact
Inform "  „ «·≈ ’«· »«·”Ì—›— "

If XALLNEW.Value <> 0 Then con.Execute " UPDATE model_dem SET ISNEW = 1 "
SendOrder
Inform " „  —ÕÌ· «·ÿ·»Ì… ··≈œ«—…"
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Function openConFACT(ByRef pCon As ADODB.Connection, Optional ByVal pString As String = "") As String
On Error GoTo myerror
Dim cString As String
If pString = "" Then cString = strConfact Else cString = pString
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.Open cString
openConFACT = "ok"
Exit Function
myerror:
openConFACT = Err.Description
Err.Clear
End Function
Private Function LoadConStringsFACT()
Exit Function
Dim cServerName As String, cUserId As String, cPassword As String
cServerName = GetDesca("SELECT IPSERVER FROM branch ", con)
If UCase(cServerName) = "ONLINE" Then lServerOnLineShop = True
If lServerOnLineShop Then
'    cUserId = decrypt("062F8C407C77", "dr")
'    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
'    cServerName = "junior-sql.database.windows.net"
Else
    cUserId = decrypt(RetSetting("userId", App.Path & "\conf.txt"), "dr")
    cPassword = decrypt(RetSetting("Password", App.Path & "\conf.txt"), "dr")
End If
LoadConStringsFACT = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
            & "catalog=JUNIOR_SHOP;user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Private Function SendOrder() As Long
'On Error GoTo myerror
Dim cFile As String
cString = "SELECT * FROM model_dem WHERE ISNEW = 1 "
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    confact.Execute " DELETE FROM model_dem where item = " & Val(loctable!Item)
    
'    Me.Caption = loctable!Item
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If Val(loctable!Quant) <> 0 Then
        aInsert = AddFlag(Empty, "MODELNO", addstring(loctable!MODELNO))
        aInsert = AddFlag(aInsert, "REMARK", addstring(loctable!remark))
        aInsert = AddFlag(aInsert, "DATE", "GETDATE()")
        aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant))
        aInsert = AddFlag(aInsert, "ITEM", Val(loctable!Item))
        aInsert = AddFlag(aInsert, "STORE", addstring(loctable!STORE))
        confact.Execute addInsert(aInsert, "model_dem"), nAffect
        con.Execute " UPDATE model_dem SET ISNEW = 0 WHERE ITEM = " & MyParn(loctable!Item)
    Else
        con.Execute " DELETE FROM model_dem WHERE ITEM = " & Val(loctable!Item)
    End If
    loctable.MoveNext
    sendItems = sendItems + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendItems = -1
GoTo lastsub
End Function

Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
imgx1.Images.Clear
imgx1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub


Private Sub myLoadPicture(pModel As String)
On Error Resume Next
    If Not lServerPict Then Exit Sub
    imgx1.Images.Clear
    If pModel = "" Then Exit Sub
    If conPict Is Nothing Then Exit Sub
    Dim loctable As Recordset
    Dim bytes() As Byte, nSize As Long
    imgx1.Images.Clear
    Set loctable = myRecordSet("select PICT,size from PICT where MODELNO = " & MyParn(pModel), conPict)
    If loctable.EOF Then Exit Sub
    nSize = CLng(loctable("size").Value)
    bytes = loctable("PICT").GetChunk(nSize)
    imgx1.Import.FromMemoryFile bytes, ixmfJPG
    Err.Clear
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
    If OldRow <> NewRow Then
        myLoadPicture xModelNo.Caption
    End If
End With

End Sub
Private Sub CMD_ALL_DEL_Click()
If xStore.BoundText = "" Then Exit Sub
If MsgBox("Õ–› «·ÿ·»Ì«  ··„ÊœÌ·«  «·„Õœœ… «·Œ«’… »«·›—⁄ ", vbYesNo + vbDefaultButton2) = vbYes Then
    With grid1
        For nRow = 3 To .Rows - 1
            con.Execute "DELETE FROM MODEL_DEM WHERE MODELNO = " & MyParn(.TextMatrix(nRow, 3)) & " AND STORE = " & MyParn(xStore.BoundText)
        Next nRow
        Inform "  „ Õ–› ÿ·»Ì«  «·›—⁄  ··„ÊœÌ·«  «·„Õœœ… "
    End With
End If
End Sub
Private Sub del_dem_Click()
    con.Execute " delete from model_dem where item in (select item from file1_10 where ISNODEM = 1 ) ", nRec
    Inform "  „ Õ–› ··›—Ê⁄ ⁄œœ " & nRec
End Sub
Sub PrintScalDem()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset, nBalance As Single
Dim ModelsTable As New ADODB.Recordset
Dim ScalTable As New ADODB.Recordset
Dim BranchTable As New ADODB.Recordset

Dim aHeader(2)
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

    cStr1 = "  SELECT  FILE1_10.MODEL, FILE1_10.desca  , FILE1_10.COLOR , FILE1_10.MODELFACT0 " & _
                  " FROM  FILE1_10 INNER JOIN MODEL_DEM ON FILE1_10.ITEM = MODEL_DEM.ITEM WHERE FILE1_10.MODEL IS NOT NULL  "
    If xStore.BoundText <> "" Then cStr1 = cStr1 & " AND MODEL_DEM.STORE = " & MyParn(xStore.BoundText)
    If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XISNOMOSM2.Value = 0 Then
        If XMOSM2.text <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM2 = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr1 = cStr1 & " AND ( FILE1_10.MOSM2 IS NULL OR FILE1_10.MOSM2 <> " & MyParn(XMOSM2.text) & " ) "
    End If
    If xModelFact.text <> "" Then cStr1 = cStr1 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xbarcode.text <> "" Then cStr1 = cStr1 & " AND FILE1_10.item = " & xbarcode.text
    If xSupp.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XSECTION.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[Section] = " & XSECTION.BoundText
    If Check2.Value <> 0 Then cStr1 = cStr1 & " AND [ISPRINTED] = 0  "
    If IsDate(XDATEDEM1.text) Then cStr1 = cStr1 & " AND DATE  >= " & DateSq(XDATEDEM1.text)
    If IsDate(XDATEDEM2.text) Then cStr1 = cStr1 & " AND DATE  <= " & DateSq(XDATEDEM2.text)
    If XDESCA.text <> "" Then cStr1 = cStr1 & " AND file1_10.DESCA LIKE ('%" & XDESCA.text & "%')   "
    cStr1 = cStr1 & " GROUP BY FILE1_10.MODEL, FILE1_10.desca  , FILE1_10.COLOR , FILE1_10.MODELFACT0  "
    ModelsTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    Dim cModel As String, cColor As String, cModelFact As String
    Dim TModelDem As New Recordset
    Dim nTotal As Double
    Dim nCol As Double
    If ModelsTable.RecordCount = 0 Then
        Exit Sub
    End If
    With temptable
     
    prog1.Max = ModelsTable.RecordCount
    prog1.Min = 0
    prog1.Visible = True
    I = 0
    ModelsTable.MoveFirst
    Do While Not ModelsTable.EOF
        I = I + 1
        prog1.Value = I
        cModel = ModelsTable!MODEL
        cModelFact = ModelsTable!modelfact0
        cColor = ModelsTable!color
    
        If ScalTable.State = adStateOpen Then ScalTable.Close
        cStr1 = "SELECT C_SCAL , SCAL FROM FILE1_10 WHERE MODEL =  " & MyParn(cModel) & " AND COLOR = " & MyParn(cColor) & " GROUP BY FILE1_10.C_SCAL , SCAL ORDER BY C_SCAL "
        ScalTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    
        cStr1 = "SELECT STORE , STORE_BR.DESCA FROM MODEL_DEM INNER JOIN STORE_BR ON STORE_BR.CODE = MODEL_DEM.STORE INNER JOIN FILE1_10 ON FILE1_10.ITEM = MODEL_DEM.ITEM WHERE FILE1_10.MODEL =  " & MyParn(cModel) & " AND FILE1_10.COLOR = " & MyParn(cColor) & " GROUP BY STORE , STORE_BR.DESCA  ORDER BY STORE  "
        If BranchTable.State = adStateOpen Then BranchTable.Close
        BranchTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
        
        If BranchTable.RecordCount > 0 Then
            Do While Not BranchTable.EOF
                Me.Caption = cModel
                .AddNew
                !STR20 = cModel & " " & cColor
                !str17 = cModelFact
                !STR19 = cColor
                !str18 = ArbString(ModelsTable!DESCA)
                nTotal = 0
                ScalTable.MoveFirst
                nCol = 2
                Do While Not ScalTable.EOF
                    nCol = nCol + 1
                    .Fields("STR" & nCol) = ScalTable!scal
                    nItem = Val(GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(cModel) & " AND COLOR = " & MyParn(cColor) & " AND SCAL = " & MyParn(ScalTable!scal), con) & "")
                    If nItem > 0 Then
                        .Fields("VAL" & nCol) = Val(GetDesca("SELECT QUANT FROM MODEL_DEM WHERE ITEM = " & nItem & " AND STORE = " & MyParn(BranchTable!STORE), con) & "")
                        nTotal = nTotal + Val((.Fields("VAL" & nCol)) & "")
                    End If
                    ScalTable.MoveNext
                Loop
                !str16 = BranchTable!DESCA
                !VAL13 = nTotal
                .Update
                BranchTable.MoveNext
            Loop
        End If
        ModelsTable.MoveNext
    Loop

    contemp.BeginTrans
    contemp.CommitTrans

    If temptable.EOF And temptable.BOF Then
        MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
    Else
        Main.REPORT1.ReportFileName = App.Path & "\Reports\PrintModelScal.rpt"
        Main.REPORT1.DataFiles(0) = tempFile
        Main.REPORT1.Action = 1
    End If
    End With
    temptable.Close
    ScalTable.Close
    BranchTable.Close
    Set temptable = Nothing
    Set BranchTable = Nothing
End Sub


