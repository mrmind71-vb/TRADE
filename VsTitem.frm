VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form VsTItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄… «·„ÊœÌ·« "
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20400
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
   ScaleWidth      =   20400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame10 
      Height          =   1230
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   103
      Top             =   3060
      Width           =   4110
      Begin VB.CheckBox chkStoped 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "≈ŸÂ«— ›—Ê⁄ „ Êﬁ›…"
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
         Height          =   240
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   225
         Width           =   1815
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "ﬂ· «·«’‰«› »Ì⁄ «Êﬂ«“ÊÌ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   855
         Width           =   2265
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "⁄—÷ «Êﬂ«“ÊÌ‰ «·⁄—Ê÷"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   540
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "≈Ã„«·Ï „Ã„Ê⁄… „ÊœÌ·"
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
         Height          =   285
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   180
         Width           =   2175
      End
   End
   Begin Threed.SSCommand cmd_okaz 
      Height          =   510
      Left            =   4320
      TabIndex        =   52
      Top             =   3690
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   900
      _Version        =   196610
      BackColor       =   13822956
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ”⁄Ì— ‰”»… «·«Êﬂ«“ÊÌ‰"
      PictureAlignment=   1
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
      Height          =   4335
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   -45
      Width           =   4695
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   3750
         Left            =   90
         TabIndex        =   66
         Top             =   225
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   6615
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
   Begin VB.Frame Frame6 
      Caption         =   " ⁄œÌ· ≈·Ï "
      Height          =   3030
      Left            =   8640
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   45
      Width           =   4200
      Begin VB.CommandButton cmd_save9 
         Caption         =   " ⁄œÌ·"
         Height          =   330
         Left            =   2430
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "⁄—÷"
         Top             =   2025
         Width           =   690
      End
      Begin VB.CheckBox xEdit_STOP 
         Alignment       =   1  'Right Justify
         Caption         =   "„ Êﬁ›"
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
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   2115
         Width           =   1095
      End
      Begin VB.CheckBox XNOEDIT 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         Caption         =   "»œÊ‰  ⁄œÌ· ›Ï «·›—Ê⁄"
         Height          =   240
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   2610
         Value           =   1  'Checked
         Width           =   1860
      End
      Begin VB.CommandButton cmd_save8 
         Caption         =   " ⁄œÌ·"
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "⁄—÷"
         Top             =   2655
         Width           =   690
      End
      Begin VB.CheckBox xEdit_Outlet 
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
         Height          =   195
         Left            =   945
         RightToLeft     =   -1  'True
         TabIndex        =   98
         Top             =   2655
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save7 
         Caption         =   " ⁄œÌ·"
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   96
         ToolTipText     =   "⁄—÷"
         Top             =   2340
         Width           =   690
      End
      Begin VB.CheckBox xEdit_ISONLINE 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ê‰ ·«Ì‰ "
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
         Left            =   945
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   2340
         Width           =   1095
      End
      Begin VB.CommandButton cmd_save6 
         Caption         =   " ⁄œÌ·"
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   93
         ToolTipText     =   "⁄—÷"
         Top             =   2025
         Width           =   690
      End
      Begin VB.CheckBox xEdit_ISNODEM 
         Alignment       =   1  'Right Justify
         Caption         =   "»œÊ‰ ÿ·»Ì…"
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
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   2025
         Width           =   1140
      End
      Begin VB.CommandButton cmd_save5 
         Caption         =   " ⁄œÌ·"
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   76
         ToolTipText     =   "⁄—÷"
         Top             =   1620
         Width           =   690
      End
      Begin VB.CommandButton cmd_save4 
         Caption         =   " ⁄œÌ·"
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "⁄—÷"
         Top             =   1260
         Width           =   690
      End
      Begin VB.ComboBox XM 
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
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   900
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   " ⁄œÌ·"
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "⁄—÷"
         Top             =   855
         Width           =   690
      End
      Begin VB.CommandButton cmd_save2 
         Caption         =   " ⁄œÌ·"
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "⁄—÷"
         Top             =   495
         Width           =   690
      End
      Begin VB.CommandButton cmd_save 
         Caption         =   " ⁄œÌ·"
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   690
      End
      Begin MSDataListLib.DataCombo xage 
         Height          =   315
         Left            =   900
         TabIndex        =   70
         Top             =   1665
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xsec 
         Height          =   315
         Left            =   900
         TabIndex        =   72
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XGR 
         Height          =   315
         Left            =   900
         TabIndex        =   73
         Top             =   540
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xsex 
         Height          =   315
         Left            =   900
         TabIndex        =   74
         Top             =   1305
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·›∆… «·⁄„—Ì… : "
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
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   1665
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "«·‰Ê⁄ : "
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
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   1350
         Width           =   825
      End
      Begin VB.Label Label2 
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
         Height          =   240
         Index           =   8
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… : "
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
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   720
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ : "
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
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   300
         Width           =   825
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1320
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   2880
      Width           =   3885
      Begin VB.CommandButton cmdBarCode 
         Height          =   465
         Left            =   2520
         Picture         =   "VsTitem.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   810
         Width           =   1275
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   465
         Left            =   1305
         Picture         =   "VsTitem.frx":2529
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   315
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   90
         Picture         =   "VsTitem.frx":4953
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   315
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2520
         Picture         =   "VsTitem.frx":6DBF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "⁄—÷"
         Top             =   315
         Width           =   1275
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   90
         TabIndex        =   48
         Top             =   135
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   90
         TabIndex        =   51
         Top             =   810
         Width           =   2370
         _ExtentX        =   4180
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
         Picture         =   "VsTitem.frx":92B1
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   0
      Width           =   3840
      Begin VB.CommandButton CMD_TOT 
         BackColor       =   &H00FFFF80&
         Caption         =   "«ŸÂ«—"
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   540
         Width           =   645
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "‰ «·„»Ì⁄«  :"
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
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   960
      End
      Begin VB.Label xbalcode 
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
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2070
         Width           =   1725
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ”«» «·„Ê—œ :"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2055
         Width           =   1185
      End
      Begin VB.Label xcash 
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
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2430
         Width           =   1725
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "œ›⁄… „” Õﬁ… :"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   2430
         Width           =   1590
      End
      Begin VB.Label XTSAL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1680
         Width           =   1725
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·„»Ì⁄«  :"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1680
         Width           =   1590
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   600
      End
      Begin VB.Label xTBal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   930
         Width           =   1725
      End
      Begin VB.Label xTSalCOST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1305
         Width           =   1725
      End
      Begin VB.Label xTPurch 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   555
         Width           =   1725
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   " ﬂ·›… „»Ì⁄«  :"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1305
         Width           =   1590
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… «·—’Ìœ :"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   930
         Width           =   1590
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬁÌ„… „‘ —Ì«  :"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   555
         Width           =   1590
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·„ÊœÌ·«  :"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   180
         Width           =   1590
      End
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   645
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4155
      Left            =   12870
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   7710
      Begin VB.CheckBox XISONLINE 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ê‰ ·«Ì‰ "
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   97
         Top             =   2340
         Width           =   1320
      End
      Begin VB.CheckBox BARCODE2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "»«—ﬂÊœ „ﬂ——"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   3240
         Width           =   1140
      End
      Begin VB.Frame Frame8 
         Height          =   600
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   3465
         Width           =   3570
         Begin VB.OptionButton xall 
            Alignment       =   1  'Right Justify
            Caption         =   "«·ﬂ·"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   225
            Value           =   -1  'True
            Width           =   555
         End
         Begin VB.OptionButton xall 
            Alignment       =   1  'Right Justify
            Caption         =   " ÊﬂÌ·«  ›ﬁÿ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   990
            RightToLeft     =   -1  'True
            TabIndex        =   89
            Top             =   225
            Width           =   1140
         End
         Begin VB.OptionButton xall 
            Alignment       =   1  'Right Justify
            Caption         =   "›—Ê⁄ ›ﬁÿ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   2385
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   225
            Width           =   960
         End
      End
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
         Height          =   240
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   2340
         Width           =   1005
      End
      Begin VB.Frame Frame7 
         Height          =   600
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   3465
         Width           =   3795
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
            TabIndex        =   85
            Top             =   180
            Value           =   -1  'True
            Width           =   780
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
            TabIndex        =   84
            Top             =   180
            Width           =   825
         End
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
            TabIndex        =   83
            Top             =   180
            Width           =   1185
         End
      End
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   "„”„ÊÕ ·Â« ÿ·»Ì« "
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   3150
         Width           =   2175
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊœÌ·«   «„«‰« "
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   2340
         Width           =   1500
      End
      Begin VB.ComboBox xmosm2 
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
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   2790
         Width           =   1050
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "„« ⁄œ«"
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
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   2880
         Width           =   780
      End
      Begin VB.CheckBox XISSTOP 
         Alignment       =   1  'Right Justify
         Caption         =   "«’‰«› €Ì— „ Êﬁ›…"
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1260
         Width           =   1770
      End
      Begin VB.TextBox XOKAZ 
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
         MaxLength       =   15
         TabIndex        =   46
         Top             =   855
         Width           =   1545
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   135
         Width           =   2805
         Begin Threed.SSCommand CMD_RATE 
            Height          =   375
            Left            =   1530
            TabIndex        =   44
            Top             =   135
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   661
            _Version        =   196610
            ForeColor       =   64
            Caption         =   "‰”»… „»Ì⁄« "
            ButtonStyle     =   2
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
            Left            =   45
            MaxLength       =   15
            TabIndex        =   43
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
            Left            =   765
            MaxLength       =   15
            TabIndex        =   42
            Top             =   180
            Width           =   600
         End
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
         Left            =   4770
         MaxLength       =   15
         TabIndex        =   15
         Top             =   1620
         Width           =   1635
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
         Left            =   135
         MaxLength       =   15
         TabIndex        =   10
         Top             =   1575
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1935
         Width           =   3915
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
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   2970
         TabIndex        =   3
         Top             =   1260
         Width           =   3435
         _ExtentX        =   6059
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
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   2970
         TabIndex        =   4
         Top             =   900
         Width           =   3435
         _ExtentX        =   6059
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
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   4770
         TabIndex        =   5
         Top             =   1980
         Width           =   1635
         _ExtentX        =   2884
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
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   2970
         TabIndex        =   13
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
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
      Begin Threed.SSCheck XRET 
         Height          =   330
         Left            =   2970
         TabIndex        =   45
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   12648447
         Caption         =   "√Õ›«¡ «·„— Ã⁄ »«·ﬂ«„·"
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   4770
         TabIndex        =   53
         Top             =   2385
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xmodelsex 
         Height          =   315
         Left            =   3780
         TabIndex        =   77
         Top             =   2745
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xmodelage 
         Height          =   315
         Left            =   3780
         TabIndex        =   78
         Top             =   3150
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "«·›∆… «·⁄„—Ì… : "
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
         Index           =   12
         Left            =   6435
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   3195
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "«·‰Ê⁄ : "
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
         Index           =   11
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   2790
         Width           =   825
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
         Index           =   6
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   2835
         Width           =   1140
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   2415
         Width           =   825
      End
      Begin VB.Label Label2 
         Caption         =   "‰”»… «Êﬂ«“ÊÌ‰"
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
         Index           =   2
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label xModelNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1935
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "»«—ﬂÊœ :"
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
         Height          =   270
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1725
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   600
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ÊœÌ· :"
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
         Height          =   270
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1620
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»Ì«‰ :"
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
         Index           =   4
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2025
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… :"
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
         Index           =   1
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1395
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„’‰⁄ :"
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1020
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ê”„ :"
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
         Index           =   3
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   2025
         Width           =   870
      End
      Begin VB.Label Label1 
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   870
      End
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
   Begin MSAdodcLib.Adodc data6 
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   195
      Left            =   11790
      TabIndex        =   33
      Top             =   4095
      Visible         =   0   'False
      Width           =   9645
      _cx             =   17013
      _cy             =   344
      _ConvInfo       =   1
      Appearance      =   0
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
   Begin VSFlex7Ctl.VSFlexGrid GRID1 
      Height          =   5505
      Left            =   135
      TabIndex        =   40
      Top             =   4365
      Width           =   20175
      _cx             =   35586
      _cy             =   9710
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   16777152
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   20
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
   Begin MSAdodcLib.Adodc data20 
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
   Begin MSAdodcLib.Adodc data11 
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
   Begin MSAdodcLib.Adodc data12 
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
Attribute VB_Name = "VsTItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3, StrMosm2 As String
Public cMyModel As String
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï  ⁄«„·«  «·„ÊœÌ·«  "
    cHead2 = " Õ Ï   «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY")
    Load PrintGrd
    PrintGrd.DOPRINT Me.grid1, 0.75, -2, cHead1, cHead2, , False, True, 6
    PrintGrd.Show 1
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
    If MsgBox(" ÕœÌœ ﬂ· «·«’‰«› ›Ï «· ﬁ—Ì— »”⁄— »Ì⁄ «Êﬂ«“ÊÌ‰ ", vbYesNo + vbDefaultButton2) = vbYes Then
        With grid1
            prog1.Value = 0
            prog1.Min = 0
            prog1.Max = .Rows
            For nRow = 1 To .Rows - 1
                con.Execute " update file1_10 set [ISDATE] = getdate()  ,[ISOKAZITEM ] = 1 where modelno = " & MyParn(.TextMatrix(nRow, 30))
                .TextMatrix(nRow, 15) = True
                prog1.Value = nRow
            Next nRow
        End With
    End If
Else
    If MsgBox(" ÕœÌœ ﬂ· «·«’‰«› ›Ï «· ﬁ—Ì— »”⁄— »Ì⁄ „” Â·ﬂ", vbYesNo + vbDefaultButton2) = vbYes Then
        With grid1
            prog1.Value = 0
            prog1.Min = 0
            prog1.Max = .Rows
            For nRow = 1 To .Rows - 1
                con.Execute " update file1_10 set [ISDATE] = GETDATE() ,[ISOKAZITEM ] = 0 where modelno = " & MyParn(.TextMatrix(nRow, 30))
                .TextMatrix(nRow, 15) = False
                prog1.Value = nRow
            Next nRow
        End With
    End If
End If
End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub cmd_nopict_Click()
Dim fs As New FileSystemObject
With grid1
    For nRow = 2 To .Rows - 1
        If fs.FileExists(retFile(.TextMatrix(nRow, 30))) Then .RowHidden(nRow) = True
    Next nRow
End With
End Sub

Private Sub cmd_okaz_Click()
    CalcPkazFRm.Show 1
End Sub

Private Sub CMD_RATE_Click()
With grid1
    For nRow = 2 To .Rows - 1
        .RowHidden(nRow) = False
        If X1.text <> "" Then
            If Val(.TextMatrix(nRow, 10)) < Val(X1.text) Then .RowHidden(nRow) = True
        End If
        If X2.text <> "" Then
            If Val(.TextMatrix(nRow, 10)) > Val(X2.text) Then .RowHidden(nRow) = True
        End If
    Next nRow
End With
End Sub
Private Sub CMD_SAVE_Click()
    SubEditModel "[SECTION]", XSEC.BoundText
End Sub

Private Sub cmd_save2_Click()
Dim cStrUpdate As String
Dim cStrDEl As String
Dim cStrUpdate_fr As String
Dim cStrDEl_fr As String
Dim I As Double
Dim BranchTrans As New ADODB.Recordset
BranchTrans.Open "SELECT CODE FROM BRANCH WHERE IPSQL IS NOT NULL ", con, adOpenStatic, adLockReadOnly, adCmdText
With grid1
    prog1.Value = 0
    prog1.Min = 0
    prog1.Max = .Rows
    For nRow = 1 To .Rows - 1
        I = I + 1
        con.Execute " update file1_10 set [ISDATE] = GETDATE() ,[GROUP] = " & addstring(XGR.BoundText) & " where modelno = " & MyParn(.TextMatrix(nRow, 30))
        
'        cStrDEl = cStrDEl & " DELETE FROM NEW_ITEM FROM            NEW_ITEM INNER JOIN                          FILE1_10 ON NEW_ITEM.ITEM = FILE1_10.ITEM  where modelno = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
'        BranchTrans.MoveFirst
'        Do While Not BranchTrans.EOF
'            cBr = RetZero(sbr, 2)
'            cStrUpdate = cStrUpdate & "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(branch!CODE) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
'            BranchTrans.MoveNext
'        Loop
'
'        cStrDEl_fr = cStrDEl_fr & " DELETE FROM NEW_ITEM_fr FROM            NEW_ITEM_fr INNER JOIN  FILE1_10 ON NEW_ITEM_fr.ITEM = FILE1_10.ITEM  where modelno = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
'        For sbr = 61 To nCountBranch_fr
'            cBr = RetZero(sbr, 2)
'            cStrUpdate_fr = cStrUpdate_fr & "  INSERT INTO NEW_ITEM_fr (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
'        Next sbr
'        If i = 100 Then
'            con.Execute cStrDEl
'            con.Execute cStrUpdate
'            cStrUpdate = ""
'            cStrDEl = ""
'
'            con.Execute cStrDEl_fr
'            con.Execute cStrUpdate_fr
'            cStrUpdate_fr = ""
'            cStrDEl_fr = ""
'
'            i = 0
'        End If
        prog1.Value = nRow
    Next nRow
'    If cStrDEl <> "" Then con.Execute cStrDEl
'    If cStrUpdate <> "" Then con.Execute cStrUpdate
'
'    If cStrDEl_fr <> "" Then con.Execute cStrDEl_fr
'    If cStrUpdate_fr <> "" Then con.Execute cStrUpdate_fr

End With

End Sub

Private Sub cmd_save4_Click()
    SubEditModel "MODELSEX", xsex.BoundText

End Sub

Private Sub cmd_save5_Click()
    SubEditModel "MODELAGE", xage.BoundText
End Sub


Private Sub cmd_save9_Click()
    SubEditModel "ISSTOP", xEdit_STOP.Value

End Sub

Private Sub CMD_TOT_Click()
    If xMosm.BoundText <> "" Or xSupp.BoundText <> "" Or xFact.BoundText <> "" Or xGroup.BoundText <> "" And xall(0).Value <> 0 Then
        CalcTot
    End If

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    
    XCOUNT.Caption = ""
    xTPurch.Caption = ""
    xTSalCOST.Caption = ""
    XTSAL.Caption = ""
    xTBal.Caption = ""
    xRate.Caption = ""
    xbalcode.Caption = ""
    xcash.Caption = ""
        
    myload
    
    XCOUNT.Caption = grid1.Rows - 3
End Sub
Private Sub cmdBarCode_Click()
con.BeginTrans
On Error GoTo myerror
If xModelNo.Caption = "" Then Exit Sub
cString = "INSERT INTO ADDPRINT(ITEM,BRANCH,QUANT,ISPRINT)"
cString = cString & _
          " SELECT FILE1_10.ITEM, " & addstring(cBranch) & " ,SUM([IN]-[OUT]),1 FROM FILE1_11_ALL INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11_ALL.ITEM WHERE FILE1_10.MODELNO = " & MyParn(xModelNo.Caption)
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

Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    grid1.ColHidden(12) = True
    grid1.ColHidden(13) = True
    grid1.ColHidden(15) = True
    cHead1 = "»Ì«‰ ≈Ã„«·Ï  ⁄«„·«  «·„ÊœÌ·«  "
    cHead2 = " Õ Ï   «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY")
    
    Load PrintGrd
    PrintGrd.DOPRINT Me.grid1, 0.9, 2, cHead1, cHead2, , False, , 7
    PrintGrd.Show 1
    grid1.ColHidden(12) = False
    grid1.ColHidden(13) = False
    grid1.ColHidden(15) = False

End Sub
Private Sub Command1_Click()
Dim BranchTrans As New ADODB.Recordset
Dim cStrUpdate As String
Dim cStrDEl As String

Dim cStrUpdate_fr As String
Dim cStrDEl_fr As String

BranchTrans.Open "SELECT CODE FROM BRANCH WHERE IPSQL IS NOT NULL ", con, adOpenStatic, adLockReadOnly, adCmdText
Dim I As Double
With grid1
    prog1.Value = 0
    prog1.Min = 0
    prog1.Max = .Rows
    For nRow = 1 To .Rows - 1
        I = I + 1
        con.Execute " update file1_10 set [ISDATE] = GETDATE() ,[MOSM2] = " & addstring(XM.text) & " where modelno = " & MyParn(.TextMatrix(nRow, 30))
        
        cStrDEl = cStrDEl & " DELETE FROM NEW_ITEM FROM            NEW_ITEM INNER JOIN                          FILE1_10 ON NEW_ITEM.ITEM = FILE1_10.ITEM  where modelno = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
        BranchTrans.MoveFirst
        Do While Not BranchTrans.EOF
            cBr = RetZero(sbr, 2)
            cStrUpdate = cStrUpdate & "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(BranchTrans!CODE) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
            BranchTrans.MoveNext
        Loop
        
        cStrDEl_fr = cStrDEl_fr & " DELETE FROM NEW_ITEM_fr FROM            NEW_ITEM_fr INNER JOIN  FILE1_10 ON NEW_ITEM_fr.ITEM = FILE1_10.ITEM  where modelno = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
        For sbr = 61 To nCountBranch_fr
            cBr = RetZero(sbr, 2)
            cStrUpdate_fr = cStrUpdate_fr & "  INSERT INTO NEW_ITEM_fr (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
        Next sbr
        If I = 100 Then
            con.Execute cStrDEl
            con.Execute cStrUpdate
            cStrUpdate = ""
            cStrDEl = ""
            
            con.Execute cStrDEl_fr
            con.Execute cStrUpdate_fr
            cStrUpdate_fr = ""
            cStrDEl_fr = ""
            I = 0
        End If
        prog1.Value = nRow
    Next nRow
    If cStrDEl <> "" Then con.Execute cStrDEl
    If cStrUpdate <> "" Then con.Execute cStrUpdate

    If cStrDEl_fr <> "" Then con.Execute cStrDEl_fr
    If cStrUpdate_fr <> "" Then con.Execute cStrUpdate_fr
End With
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Load()
    openCon con
    
    StoreTable.Open " SELECT * FROM FILE0_40 WHERE ISSTOP = 0 ", con, adOpenStatic, adLockReadOnly, adCmdText
'   xDate1.Text = Format(Date, "dd-mm-yyyy")
    StrMosm2 = "|# ; " & "|#S;S" & "|#W;W" & "|#M;M"
    
    Frame6.Visible = (cBranch = "00")
    cmd_okaz.Visible = (cBranch = "00") And bopt2
    
    XMOSM2.AddItem " "
    XMOSM2.AddItem "S"
    XMOSM2.AddItem "W"
    XMOSM2.AddItem "M"
    
    XM.AddItem " "
    XM.AddItem "S"
    XM.AddItem "W"
    XM.AddItem "M"
    
    Dim ImgXTmp As ImgX
    Set ImgXTmp = New ImgX
'   ImgXTmp.Import.FromMemoryFile LoadResData("BACKGROUND", "CUSTOM")
'   Set ImgX1.BackgroundImage = ImgXTmp.image
    Set ImgXTmp = Nothing
    
    
    Set data1.Recordset = myRecordSet("Select Code,DescA From File1_10SC order by Desca", con)
    Set XSECTION.RowSource = data1
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
    Set XSEC.RowSource = data1
    XSEC.ListField = "Desca"
    XSEC.BoundColumn = "Code"
    
    Set DATA2.Recordset = myRecordSet("Select Code,DescA From File4_10 order by Desca", con)
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    Set DATA3.Recordset = myRecordSet("Select Code,DescA From File1_50 ORDER BY DESCA", con)
    Set xGroup.RowSource = DATA3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set XGR.RowSource = DATA3
    XGR.ListField = "Desca"
    XGR.BoundColumn = "Code"
    
    Set DATA4.Recordset = myRecordSet("Select mosm ,descA From mosm ORDER BY date DESC ", con)
    Set xMosm.RowSource = DATA4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    
    Set DATA5.Recordset = myRecordSet("Select code,desca From fact ORDER BY code ", con)
    Set xFact.RowSource = DATA5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    
    Set DATA11.Recordset = myRecordSet("SELECT * FROM FILE0_82 ORDER BY DESCA", con)
    Set xmodelage.RowSource = DATA11
    xmodelage.ListField = "DESCA"
    xmodelage.BoundColumn = "CODE"
    
    Set xage.RowSource = DATA11
    xage.ListField = "DESCA"
    xage.BoundColumn = "CODE"
    
    Set data12.Recordset = myRecordSet("SELECT * FROM FILE0_81 ORDER BY DESCA", con)
    Set xmodelsex.RowSource = data12
    xmodelsex.ListField = "DESCA"
    xmodelsex.BoundColumn = "CODE"
    
    Set xsex.RowSource = data12
    xsex.ListField = "DESCA"
    xsex.BoundColumn = "CODE"
    
    
    Set grid1.DataSource = DATA6
    
'    Set grid2.DataSource = data20
    chkStoped.Visible = lSupperVisor
    grid1.Rows = 2
    
'    grid2.Rows = 1
    
    Fixgrd
    Fixgrd2
    Frame2.Visible = bOpt10
End Sub
Private Sub myload()
    Dim cWhere As String, cString As String
    Dim cFOkaz As String
     
    cField = myiif("type = 'z' ", "([out] - [in]) ") & " AS t_comp"
    
'   cField = " Count (DISTINCT file1_10.item) "
    
    cField = cField & "," & _
            myiif("type = '2' ", "([IN] - [OUT]) ") & " AS t_purch "
    
    cField = cField & "," & _
             myiif("type = '7' ", "([out] - [in]) ") & " AS t_ret"

    cField = cField & "," & _
             myiif("type = '6' ", "([out] - [in]) ") & " AS t_sal"



'    Dim aRet As Variant
'    aRet = GetRows("Select code,Desca from file0_40 WHERE ISSTOP = 0 ")
'    For i = 0 To UBound(aRet)
'        cField3 = cField3 & turn(cField3, ",") & _
'                   myiif("FILE1_11_ALL.STORE = " & MyParn(retFlag(aRet(i), "CODE")), "([IN]-[OUT]) ") & " AS [" & retFlag(aRet(i), "CODE") & "]"
'    Next
'
    cField2 = "DATEDIFF(dd , " & myiif2("type = '2' ", "[date]", "MIN") & " , GETDATE()) AS DATEDIFF"
    cField2 = cField2 & " , " & myiif2("type = '2' ", "[date]", "MIN")
    cField2 = cField2 & " , " & myiif2("type = '2' ", "[date]", "MAX")
    
    If Check2.Value = 0 Then
        cFOkaz = " MAX(OKAZ) , MIN(PRICE2) "
    Else
        cFOkaz = " MAX(OKAZ_2) , MIN(PRICE_2) "
    End If
    '                   0                   1               2           3                   4
    cString = " SELECT file1_10.GrModel , fact.desca , file1_10.mosm, file1_10.MODELFACT0, file1_10.desca," & _
            cField & _
            ", sum([in] - out) as [«·—’Ìœ] ,  " & _
            " ' ' as rate , min(file1_10.price) , min(file1_10.COSTITEM) , sum((FILE1_11_ALL.[in] - FILE1_11_ALL.out) * file1_10.COSTITEM ) as costbal ," & cField2 & _
            ", ISOKAZITEM AS N15 , " & cFOkaz & " , FILE1_10.ISSTOP , MOSM2 , ISNODEM, FILE1_10.[Section] , FILE1_10.[GROUP] , FILE1_10.[MODELSEX]  , FILE1_10.[MODELAGE] ,Outlet  , ISONEST , ISONLINE , ModelNo " & _
            " FROM FILE1_10 INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE " & _
            " INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM " & _
            " INNER JOIN STORE_BR ON FILE1_11_ALL.STORE = STORE_BR.CODE" & _
            " WHERE file1_10.ISNOITEM  = 0  "
            
    If chkStoped.Value = 0 Then cStr2 = cStr2 & " AND STORE_BR.ISSTOP = 0"
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    
    If Len(xModelFact.text) = 6 Then
        If xModelFact.text <> "" Then cStr2 = cStr2 & " AND SUBSTRING(FILE1_10.MODELFACT0,1,6) = " & MyParn(xModelFact.text)
    Else
        If xModelFact.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    End If
    If xbarcode.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.item = " & Val(xbarcode.text)
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If XSECTION.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Section] = " & Val(XSECTION.BoundText)
    If xmodelsex.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[ModelSex] = " & Val(xmodelsex.BoundText)
    If xmodelage.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[Modelage] = " & Val(xmodelage.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xisstop.Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.[ISSTOP] = 0 "
    If XOutlet.Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.[Outlet] = 1 "
    If XISONLINE.Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.[ISONLINE] = 1 "

    If Check2.Value = 0 Then
        If XOKAZ.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.[OKAZ] = " & Val(XOKAZ.text)
    Else
        If XOKAZ.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.[OKAZ_2] = " & Val(XOKAZ.text)
    End If
    If Trim(XDESCA.text) <> "" Then cString = cString & turn(cString) & MyParnAnd(XDESCA.text, "file1_10.desca")
    If IsDate(xDate1.text) Then cStr2 = cStr2 & " AND FILE1_11_ALL.[DATE] <= " & DateSq(xDate1.text)
    If Check3.Value = 0 Then
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND ( FILE1_10.MOSM2 IS NULL OR FILE1_10.MOSM2 <> " & MyParn(XMOSM2.text) & " ) "
    End If
    If Check4.Value <> 0 Then
        cStr2 = cStr2 & " AND FILE1_10.ITEM IN (SELECT ITEM FROM Q_ONEST) "
    End If
    If Check5.Value <> 0 Then
        cStr2 = cStr2 & " AND FILE1_10.ISNODEM = 0 "
    End If
    If xonest(0).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 0 "
    If xonest(1).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_10.ISONEST = 1 "
    
    If xall(0).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_11_ALL.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cStr2 = cStr2 & " AND FILE1_11_ALL.ISBRANCH = 0  "
    
    If BARCODE2.Value <> 0 Then cStr2 = cStr2 & " AND MODELNO IN (SELECT MODELNO FROM D_MODELNO2 )"
    cString = cString & cStr2 & " GROUP BY ISOKAZITEM,Outlet ,file1_10.ModelNo, FACT.desca , file1_10.mosm, FILE1_10.[Section], file1_10.MODELFACT0, file1_10.desca , FILE1_10.FACT , FILE1_10.[GROUP] , FILE1_10.ISSTOP , MOSM2 , ISNODEM , FILE1_10.[MODELSEX]  , FILE1_10.[MODELAGE] , ISONEST  , GrModel , ISONLINE "
    If Check6.Value <> 0 Then cString = cString & " ORDER BY GrModel , FILE1_10.MODELNO "
    Set DATA6.Recordset = myRecordSet(cString, con)
    
    Fixgrd
    

End Sub
Sub Fixgrd()
    With grid1
    .Cols = 31
    .ExplorerBar = flexExSortShow
    .AllowUserResizing = flexResizeColumns
        
    .FixedRows = 2
    .RowHidden(1) = True
    .RowHeight(0) = 700
    .WordWrap = True
    .ColHidden(0) = True
    .RowHidden(1) = True
    .FrozenCols = 5
    
    .TextMatrix(0, 0) = "„ÊœÌ·"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "„ÊœÌ· "
    .TextMatrix(0, 4) = "«·’‰›"

    .TextMatrix(0, 5) = "Ã—œ"
    .TextMatrix(0, 6) = "„‘ —Ì« "
    .TextMatrix(0, 7) = "„— Ã⁄« "
    .TextMatrix(0, 8) = "„»Ì⁄« "
    .TextMatrix(0, 9) = "—’Ìœ"

    .TextMatrix(0, 10) = "‰”»… »Ì⁄"
    .TextMatrix(0, 11) = " ”⁄— »Ì⁄"
    .TextMatrix(0, 12) = "”⁄—  ﬂ·›…"
    .TextMatrix(0, 13) = " ﬁÌ„… «· ﬂ·›…"
'
    .TextMatrix(0, 14) = "⁄„— „ÊœÌ·"
    
    .TextMatrix(0, 15) = "«Ê·  «—ÌŒ"
    .TextMatrix(0, 16) = "«Œ—  «—ÌŒ"
    
    .TextMatrix(0, 17) = "»Ì⁄ √Êﬂ«“ÊÌ‰"
    If Check2.Value = 0 Then
        .TextMatrix(0, 18) = "‰”»… √Êﬂ«“ÊÌ‰"
        .TextMatrix(0, 19) = "”⁄— √Êﬂ«“ÊÌ‰"
    Else
        .TextMatrix(0, 18) = "‰”»… √Êﬂ«“ÊÌ‰ ⁄—Ê÷"
        .TextMatrix(0, 19) = "”⁄— √Êﬂ«“ÊÌ‰ ⁄—Ê÷"
        .ColHidden(17) = True
    End If
    
    .TextMatrix(0, 20) = "’‰› „ Êﬁ›"
    .TextMatrix(0, 21) = "„Ê”„ œ«Œ·Ï"
    .TextMatrix(0, 22) = "»œÊ‰ ÿ·»Ì…"
    .TextMatrix(0, 23) = "«·ﬁ”„"
    .TextMatrix(0, 24) = "«·„Ã„Ê⁄…"
    .TextMatrix(0, 25) = "«·‰Ê⁄"
    .TextMatrix(0, 26) = "«·›∆… «·⁄„—Ì…"
    .TextMatrix(0, 27) = "Outlet "
    .TextMatrix(0, 28) = "«„«‰« "
    .TextMatrix(0, 29) = "«Ê‰·«Ì‰"
    
    
    .ColHidden(30) = True
    .ColHidden(12) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    
    .ColWidth(0) = 0
    .ColWidth(1) = 1500
    .ColWidth(2) = 500
    .ColWidth(3) = 1500
    .ColWidth(4) = 3600
    .ColWidth(5) = 500
    .ColWidth(6) = 700
    .ColWidth(7) = 700
    .ColWidth(8) = 700
    .ColWidth(9) = 700
    
    .ColWidth(10) = 700
    .ColWidth(11) = 700
    .ColWidth(12) = 700
    .ColWidth(13) = 900
    
    .ColWidth(14) = 550
    .ColWidth(15) = 1300
    .ColWidth(16) = 1300
    .ColWidth(17) = 700
    .ColWidth(18) = 700
    .ColWidth(19) = 700
    
    .ColWidth(20) = 700
    .ColWidth(21) = 700
    .ColWidth(22) = 700
    
    .ColWidth(23) = 1000
    .ColWidth(24) = 1200
    .ColWidth(25) = 1000
    .ColWidth(26) = 1200
    .ColWidth(27) = 700
    .ColWidth(28) = 700
    .ColWidth(29) = 700
    
    
    .ColDataType(20) = flexDTBoolean
    .ColDataType(22) = flexDTBoolean
    .ColDataType(27) = flexDTBoolean
    .ColDataType(28) = flexDTBoolean
    .ColDataType(29) = flexDTBoolean
    
    .ColComboList(21) = StrMosm2
    .ColComboList(23) = StrList("SELECT CODE , DESCA FROM FILE1_10SC  ORDER BY DESCA ", con)
    If XSEC.BoundText <> "" Then
        .ColComboList(24) = StrList("SELECT CODE , DESCA FROM FILE1_50 WHERE [GROUP] = " & Val(XSEC.BoundText) & "  ORDER BY DESCA ", con)
    Else
        .ColComboList(24) = StrList("SELECT CODE , DESCA FROM FILE1_50 ORDER BY DESCA ", con)
    End If
    
    .ColComboList(25) = StrList("SELECT CODE , DESCA FROM FILE0_81  ORDER BY DESCA ", con)
    .ColComboList(26) = StrList("SELECT CODE , DESCA FROM FILE0_82 ORDER BY DESCA ", con)
    
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = 4
    
'    .TextMatrix(0, .Cols - 7) = "’‰› „ Êﬁ›"
'    .TextMatrix(0, .Cols - 6) = "„Ê”„ œ«Œ·Ï"
'    .TextMatrix(0, .Cols - 5) = "»œÊ‰ ÿ·»Ì…"
'    .TextMatrix(0, .Cols - 4) = "«·ﬁ”„"
'    .TextMatrix(0, .Cols - 3) = "«·„Ã„Ê⁄…"
'    .TextMatrix(0, .Cols - 2) = "«·‰Ê⁄"
'    .TextMatrix(0, .Cols - 1) = "«·›∆… «·⁄„—Ì…"
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString
    
    For nCol = 5 To 19
        .ColDataType(nCol) = flexDTDouble
        .ColWidth(nCol) = 800
    Next nCol
'    .ColDataType(15) = flexDTBoolean
    .ColDataType(17) = flexDTString
    
    .ColWidth(15) = 1300
    .ColWidth(16) = 1300
    
    
    .Editable = flexEDNone
    .ColDataType(13) = flexDTDouble
    .ColDataType(17) = flexDTBoolean
    .ColFormat(13) = "#0.00"
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    .Subtotal flexSTClear
    .WordWrap = True
    .Subtotal flexSTSum, -1, 5, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 6, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 7, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 8, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 9, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 13, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    
    If Check6.Value <> 0 Then
        .Subtotal flexSTCount, 0, 4, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
        .Subtotal flexSTSum, 0, 5, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
        .Subtotal flexSTSum, 0, 6, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
        .Subtotal flexSTSum, 0, 7, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
        .Subtotal flexSTSum, 0, 8, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
        .Subtotal flexSTSum, 0, 9, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
        .Subtotal flexSTSum, 0, 13, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    End If
    
    If .Rows > 2 Then
        For I = 2 To .Rows - 1
            If (Val(.TextMatrix(I, 6)) = Val(.TextMatrix(I, 7))) Then
                .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = &HC0C0FF
            End If
            If Val(.TextMatrix(I, 22)) <> 0 Then
                .Cell(flexcpBackColor, I, 0, I, .Cols - 1) = &H8080&
            End If
            nRat = 0
            If (Val(.TextMatrix(I, 6)) - Val(.TextMatrix(I, 7))) > 0 Then nRat = Val(.TextMatrix(I, 8)) / (Val(.TextMatrix(I, 6)) - Val(.TextMatrix(I, 7))) * 100
            .TextMatrix(I, 10) = Format(nRat, "##0.00")
        Next I
    End If
    If .Rows > 2 Then
        nRat = 0
        If (Val(.TextMatrix(2, 6)) - Val(.TextMatrix(2, 7))) > 0 Then nRat = Val(.TextMatrix(2, 8)) / (Val(.TextMatrix(2, 6)) - Val(.TextMatrix(2, 7))) * 100
        .TextMatrix(2, 10) = Format(nRat, "##0.00")
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    End If

End With
End Sub
Sub Fixgrd2()
    With GRID2
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .Cols = 8
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    .TextMatrix(0, 2) = "„‘ —Ì« "
    .TextMatrix(0, 3) = "„— Ã⁄« "
    .TextMatrix(0, 4) = "’«›Ï"
    .TextMatrix(0, 5) = "„»Ì⁄« "
    .TextMatrix(0, 6) = "—’Ìœ"
    .TextMatrix(0, 7) = "‰”»… »Ì⁄"
    .ColWidth(0) = 1000
    .ColWidth(1) = 3000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 4)) <> 0 Then .TextMatrix(nRow, 7) = Format(Val(.TextMatrix(nRow, 5)) / Val(.TextMatrix(nRow, 4)), "#0.00%")
        If Val(.TextMatrix(nRow, 4)) = Val(.TextMatrix(nRow, 5)) Then
            If Val(.TextMatrix(nRow, 6)) = 0 Then .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = &HC0C0FF
        End If
    Next nRow
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub

Private Sub grid1_DblClick()
On Error GoTo myerror
    If grid1.Row <= 1 Then Exit Sub
    Dim cString As String
    If grid1.Col = 4 Then
        items.sModel = GetDesca("Select Model FROM FILE1_10H WHERE MODELNO = " & MyParn(xModelNo.Caption), con)
        items.Show 1
    ElseIf grid1.Col <= 3 Then
        Load ModelMove
        ModelMove.XMODEL.text = xModelNo.Caption
        ModelMove.Show
    ElseIf grid1.Col = 6 Or grid1.Col = 7 Then
        InvModel.pModelNo = xModelNo.Caption
        InvModel.Show
    ElseIf grid1.Col = 8 And Val(grid1.TextMatrix(grid1.Row, 8)) <> 0 Then
        cString = "SELECT SUM(QUANT) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL" & _
                  " FROM FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM" & _
                  " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO"
        If IsDate(xDate1.text) Then cString = cString & turn(cString) & "FILE6_20H.DATE <= " & DateSq(xDate1.text)
        cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL"
        grdModelfrm.pString = cString
        grdModelfrm.pFieldValue = "sumofQuant"
        grdModelfrm.pModelNo = grid1.TextMatrix(grid1.Row, 30)
        grdModelfrm.Show 1
    ElseIf grid1.Col = 9 Then
        cString = "SELECT SUM([IN]-[OUT]) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL" & _
                  " FROM FILE1_11_ALL INNER JOIN FILE1_10 ON FILE1_11_ALL.ITEM = FILE1_10.ITEM"
        If IsDate(xDate1.text) Then cString = cString & turn(cString) & "FILE1_11_ALL.DATE <= " & DateSq(xDate1.text)
        cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL"
        grdModelfrm.pString = cString
        grdModelfrm.pFieldValue = "sumofQuant"
        grdModelfrm.pModelNo = grid1.TextMatrix(grid1.Row, 30)
        grdModelfrm.Show 1
'    ElseIf Grid1.Col = 9 Then
'        VsModelNoLook.sModelNo = Grid1.TextMatrix(Grid1.Row, 0)
'        VsModelNoLook.SFact = Grid1.TextMatrix(Grid1.Row, 1)
'        VsModelNoLook.SModelFact = Grid1.TextMatrix(Grid1.Row, 3)
'        VsModelNoLook.SMosm = Grid1.TextMatrix(Grid1.Row, 2)
'        VsModelNoLook.sdesca = Grid1.TextMatrix(Grid1.Row, 4)
'        VsModelNoLook.Show 1
'    ElseIf GRID1.Col = 11 Or GRID1.Col = 12 Then
'        Ordersfrm.sModelNo = xModelNo.Caption
'        Ordersfrm.Show
    ElseIf grid1.Col >= 18 And Val(grid1.TextMatrix(grid1.Row, grid1.Col)) <> 0 Then
        cString = "SELECT SUM([IN]-[OUT]) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL" & _
                  " FROM FILE1_11_ALL INNER JOIN FILE1_10 ON FILE1_11_ALL.ITEM = FILE1_10.ITEM WHERE STORE = " & MyParn(grid1.TextMatrix(1, grid1.Col))
        If IsDate(xDate1.text) Then cString = cString & turn(cString) & "FILE1_11_ALL.DATE <= " & DateSq(xDate1.text)
        cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL"
        grdModelfrm.pString = cString
        grdModelfrm.pFieldValue = "sumofQuant"
        grdModelfrm.pModelNo = grid1.TextMatrix(grid1.Row, 30)
        grdModelfrm.Show 1
    End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub grid1_EnterCell()
    xModelNo.Caption = grid1.TextMatrix(grid1.Row, 30)
'   If xmodelno.Caption <> "" Then MYLOAD2
    If grid1.Rows > 2 Then
        If grid1.TextMatrix(2, 16) <> "" Then Exit Sub
    End If
    If grid1.Col >= 20 Or ((grid1.Col = 17 Or grid1.Col = 18 Or grid1.Col = 19) And bopt2) Then
        grid1.Editable = flexEDKbdMouse
    Else
        grid1.Editable = flexEDNone
    End If
'    grid1.ToolTipText = GetDesca("SELECT DESCA FROM FACT WHERE CODE = " & MyParn(grid1.TextMatrix(grid1.Row, 1)))
End Sub
Sub CalcTot()
    Dim cWher As String
    
    If Check3.Value = 0 Then
        If XMOSM2.text <> "" Then cWher = cWher & " AND FILE1_10.MOSM2 = " & MyParn(XMOSM2.text)
    Else
        If XMOSM2.text <> "" Then cWher = cWher & " AND ( FILE1_10.MOSM2 IS NULL OR FILE1_10.MOSM2 <> " & MyParn(XMOSM2.text) & " ) "
    End If
    If Check5.Value <> 0 Then
        cWher = cWher & " AND FILE1_10.ISNODEM = 0 "
    End If
    If Check2.Value = 0 Then
        If XOKAZ.text <> "" Then cWher = cWher & " AND FILE1_10.[OKAZ] = " & Val(XOKAZ.text)
    Else
        If XOKAZ.text <> "" Then cWher = cWher & " AND FILE1_10.[OKAZ_2] = " & Val(XOKAZ.text)
    End If
    If Trim(XDESCA.text) <> "" Then cWher = cWher & " and " & MyParnAnd(XDESCA.text, "file1_10.desca")
    If xmodelsex.BoundText <> "" Then cWher = cWher & " AND FILE1_10.[ModelSex] = " & Val(xmodelsex.BoundText)
    If xmodelage.BoundText <> "" Then cWher = cWher & " AND FILE1_10.[Modelage] = " & Val(xmodelage.BoundText)
    If XSECTION.BoundText <> "" Then cWher = cWher & " AND FILE1_10.[Section] = " & Val(XSECTION.BoundText)
    If xbarcode.text <> "" Then cWher = cWher & " AND FILE1_10.item = " & Val(xbarcode.text)
    If xMosm.BoundText <> "" Then cWher = cWher & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.text <> "" Then cWher = cWher & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xSupp.BoundText <> "" Then cWher = cWher & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xFact.BoundText <> "" Then cWher = cWher & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xisstop.Value <> 0 Then cWher = cWher & " AND FILE1_10.[ISSTOP] = 0 "
    If xGroup.BoundText <> "" Then cWher = cWher & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    
    If IsDate(xDate1.text) Then cWher = cWher & " AND [DATE] <= " & DateSq(xDate1.text)
    Dim cF1 As String
    cString = " SELECT COUNT(DISTINCT modelno) " & _
            " , ( SELECT SUM((FILE1_11.[IN] -FILE1_11.OUT) * FILE1_10.COST ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE (TYPE = '2' OR TYPE = '7' ) " & cWher & " )  " & _
            " , ( SELECT SUM(FILE6_20.TOTAL ) FROM FILE6_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_20.ITEM INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20.DOC_NO IS NOT NULL " & cWher & " )  " & _
            " , ( SELECT SUM(FILE6_20.QUANT * FILE1_10.COST ) FROM FILE6_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_20.ITEM INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO  WHERE FILE6_20.DOC_NO IS NOT NULL " & cWher & " ) " & _
            " , ( SELECT SUM((FILE1_11.[IN] -FILE1_11.OUT) * FILE1_10.COST ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE FILE1_10.ITEM IS NOT NULL " & cWher & " ) "
    cWher = ""
    If xMosm.BoundText <> "" Then cWher = cWher & " AND MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cWher = cWher & " AND CODE = " & MyParn(xSupp.BoundText)
    If xFact.BoundText <> "" Then cWher = cWher & " AND FACT = " & MyParn(xFact.BoundText)
    If xisstop.Value <> 0 Then cWher = cWher & " AND [ISSTOP] = 0 "
    If xGroup.BoundText <> "" Then cWher = cWher & " AND [GROUP] = " & MyParn(xGroup.BoundText)
            
    cString = cString & " FROM  FILE1_10  AS FILE1_10_0 WHERE FILE1_10_0.ITEM IS NOT NULL " & cWher
    aRet = aGetDesca(cString, con)
    If UBound(aRet) > 0 Then
        XCOUNT.Caption = Format(aRet(1), "#0")
        xTPurch.Caption = Format(aRet(2), "#0")
        xTSalCOST.Caption = Format(aRet(4), "#0")
        XTSAL.Caption = Format(aRet(3), "#0")
        xTBal.Caption = Format(aRet(5), "#0")
        If Val(xTPurch.Caption) <> 0 Then xRate.Caption = Format(Val(xTSalCOST.Caption) / Val(xTPurch.Caption) * 100, "#0")
    End If
    If xSupp.BoundText <> "" Then
'        xbalcode.Caption = Round(BAL_SUPPL(xSupp.BoundText, XMOSM.BoundText), "#0.00")
        If IsDate(xDate1.text) Then
            xbalcode.Caption = Round(Val(GetDesca("SELECT SUM( SAL - PAY ) FROM FILE4_11 WHERE CODE = " & MyParn(xSupp.BoundText) & " AND MOSM = " & MyParn(xMosm.BoundText) & " AND DATE <= " & DateSq(xDate1.text), con) & ""), 2)
        Else
            xbalcode.Caption = Round(Val(GetDesca("SELECT SUM( SAL - PAY ) FROM FILE4_11 WHERE CODE = " & MyParn(xSupp.BoundText) & " AND MOSM = " & MyParn(xMosm.BoundText), con) & ""), 2)
        End If
    End If
End Sub

Private Sub Grid2_dblClick()
    If bopt2 Then
        items.sModel = GetDesca("Select Model FROM FILE1_10H WHERE CODE = " & MyParn(GRID2.TextMatrix(GRID2.Row, 30)) & " AND MODELNO = " & MyParn(xModelNo.Caption), con)
        items.Show 1
    End If
End Sub

Private Sub ImgX1_Click()
If validPhoto(retFile(grid1.TextMatrix(grid1.Row, 30))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(grid1.TextMatrix(grid1.Row, 30))
    oMaxPhoto.cDesca = grid1.TextMatrix(grid1.Row, 4)
    oMaxPhoto.Show 1
End If
End Sub
Sub MYLOAD2()
'    Dim cString As String
'    cField2 = myiif("type = '2' ", "[IN]  ") & " AS t_purch, "
'    cField3 = myiif("type = '7' ", "[OUT] ") & " AS t_ret  , "
'    cField4 = myiif("type = '2' OR type = '7'", "([IN] - [OUT]) ") & " AS t_NET   , "
'    cField5 = myiif("type = '6' ", "[out]  ") & " AS t_sal  , "
'
'    '                   0                   1               2           3                   4
'    cString = " SELECT file4_10.CODE , FILE4_10.DESCA , " & _
'            cField2 & cField3 & cField4 & cField5 & _
'            " sum([in] - out) as bal   " & _
'            " FROM   FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE WHERE FILE1_10.MODELNO = " & MyParn(xmodelno.Caption)
'    cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA  "
'
'    Set data20.Recordset = myRecordSet(cString, con)
'    Fixgrd2
End Sub
Sub myload20()
    Dim cField1 As String, cField2 As String, cField3 As String, cField4 As String, cField As String
    cField1 = myiif("type = '2' ", "[IN]  ") & " AS t_purch"
    cField2 = myiif("type = '7' ", "[OUT] ") & " AS t_ret"
    cField3 = myiif("type = '2' OR type = '7'", "([IN] - [OUT]) ") & " AS t_NET"
    cField4 = myiif("type = '6' ", "[out]  ") & " AS t_sal "
    cField = cField1 & "," & cField2 & "," & cField3 & "," & cField4
    '                   0                   1               2           3                   4
    cString = " SELECT file4_10.CODE , FILE4_10.DESCA" & turn(cField, ",") & cField & _
            " FROM  (FILE1_11_ALL INNER  JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11_ALL.ITEM) INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE "
    cString = cString & turn(cString) & " FILE1_10.MODELNO = " & MyParn(xModelNo.Caption)
    cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA  "

    data20.RecordSource = cString
    data20.Refresh
    Fixgrd2
End Sub
Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
imgx1.Images.Clear
imgx1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub

Private Sub Option1_Click()

End Sub

Private Sub xFact_Validate(Cancel As Boolean)
If Not xFact.MatchedWithList Then
    cFact = xFact.BoundText
    xFact.BoundText = cFact
    If xFact.MatchedWithList Then Exit Sub
End If

End Sub

Private Sub XRET_Click(Value As Integer)
With grid1
    If XRET.Value Then
        For nRow = 2 To .Rows - 1
            If Val(.TextMatrix(nRow, 6)) = Val(.TextMatrix(nRow, 7)) Then .RowHidden(nRow) = True
        Next nRow
    Else
        For nRow = 2 To .Rows - 1
            .RowHidden(nRow) = False
        Next nRow
    End If
End With
End Sub

Private Sub xsec_LostFocus()
    DATA3.ConnectionString = strCon
    If XSEC.BoundText = "" Then
        DATA3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Else
        DATA3.RecordSource = "Select Code,DescA From File1_50 where [group] = " & Val(XSEC.BoundText) & " ORDER BY DESCA"
    End If
    Set XGR.RowSource = DATA3
    XGR.ListField = "Desca"
    XGR.BoundColumn = "Code"
    DATA3.Refresh
End Sub

Private Sub xSection_LostFocus()
    DATA3.ConnectionString = strCon
    If XSECTION.BoundText = "" Then
        DATA3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Else
        DATA3.RecordSource = "Select Code,DescA From File1_50 where [group] = " & Val(XSECTION.BoundText) & " ORDER BY DESCA"
    End If
    Set xGroup.RowSource = DATA3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    DATA3.Refresh

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
Private Sub xsupp_Validate(Cancel As Boolean)
If xSupp.BoundText = "" Then Exit Sub
If Not xSupp.MatchedWithList Then
    cFact = GetDesca("SELECT CODE FROM FILE4_10 WHERE SUBCODE = " & MyParn(xSupp.BoundText), con)
    xSupp.BoundText = cFact
    If xSupp.MatchedWithList Then Exit Sub
End If
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'On Error GoTo myerror
Dim cBr As String, nRateOkaz As Double
With grid1
    If .Col >= 20 Then
        If XNOEDIT.Value = 0 Then
            For sbr = 3 To nCountBranch
                cBr = RetZero(sbr, 2)
                con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
            Next sbr
            
            For sbr = 61 To nCountBranch_fr
                cBr = RetZero(sbr, 2)
                con.Execute "  INSERT INTO NEW_ITEM_fr (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
            Next sbr
        End If
        If .Col = 20 Then
            con.Execute " UPDATE FILE1_10 SET [isstop] = " & Val(.TextMatrix(.Row, 20)) & "   WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        End If
        
        If .Col = 21 Then
            con.Execute " UPDATE FILE1_10 SET [ISDATE] = GETDATE() ,[MOSM2] = " & addstring(.TextMatrix(.Row, 21)) & "   WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        End If
        
        If .Col = 22 Then
            con.Execute " UPDATE FILE1_10 SET [ISDATE] = GETDATE() ,[ISNODEM] = " & Val(.TextMatrix(.Row, 22)) & "       WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        End If
    
        If .Col = 23 Then
            con.Execute " UPDATE FILE1_10 SET [ISDATE] = GETDATE() ,[SECTION] = " & Val(.TextMatrix(.Row, 23)) & "       WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        End If
        If .Col = 24 Then
            con.Execute " UPDATE FILE1_10 SET [ISDATE] = GETDATE() ,[GROUP] = " & addstring(.TextMatrix(.Row, 24)) & "   WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        End If
    
        If .Col = 25 Then
            con.Execute " UPDATE FILE1_10 SET [ISDATE] = GETDATE() ,[MODELSEX] = " & Val(.TextMatrix(.Row, 25)) & "      WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        End If
    
        If .Col = 26 Then
            con.Execute " UPDATE FILE1_10 SET [ISDATE] = GETDATE() ,[MODELAGE] = " & Val(.TextMatrix(.Row, 26)) & "   WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        End If
    
        If .Col = 27 Then
            con.Execute " UPDATE FILE1_10 SET [Outlet] = " & Val(.TextMatrix(.Row, 27)) & "       WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30)), nRec
        End If
    
        If .Col = 28 Then
            con.Execute " UPDATE FILE1_10 SET [ISONEST] = " & Val(.TextMatrix(.Row, 28)) & "       WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30)), nRec
        End If
    
        If .Col = 29 Then
            con.Execute " UPDATE FILE1_10 SET [ISONLINE] = " & Val(.TextMatrix(.Row, 29)) & "       WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30)), nRec
        End If
    
    End If
    If .Col = 17 Then
        If Val(.TextMatrix(.Row, 17)) <> 0 Then
            con.Execute " UPDATE FILE1_10 SET [ISDATE] = GETDATE() ,ISOKAZITEM = 1 WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        Else
            con.Execute " UPDATE FILE1_10 SET [ISDATE] = GETDATE() ,ISOKAZITEM = 0 WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
        End If
        If XNOEDIT.Value = 0 Then
            For sbr = 3 To nCountBranch
                cBr = RetZero(sbr, 2)
                con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
            Next sbr
    
            For sbr = 61 To nCountBranch_fr
                cBr = RetZero(sbr, 2)
                con.Execute "  INSERT INTO NEW_ITEM_fr (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
            Next sbr
        End If
    End If
    
    If Col = 18 Then
        If Not .IsSubtotal(Row) Then
            If .TextMatrix(Row, 16) <> "" Then
                DoFixPrice2 grid1.TextMatrix(grid1.Row, 30), Val(grid1.TextMatrix(grid1.Row, 18)), grid1.Row
            End If
        End If
        If .IsSubtotal(Row) Then
            If MsgBox("‰”»… √Êﬂ«“ÌÊ‰ ·ﬂ· «·„ÊœÌ·« ", vbYesNo + vbDefaultButton2) = vbYes Then
                prog1.Max = .Rows - 1
                prog1.Min = 0
                prog1.Value = 0
                For nRow = 3 To .Rows - 1
                    If .TextMatrix(nRow, 18) <> Val(.TextMatrix(Row, 18)) Then
                       .TextMatrix(nRow, 18) = Val(.TextMatrix(Row, 18))
                        DoFixPrice2 grid1.TextMatrix(nRow, 30), Val(.TextMatrix(Row, 18)), nRow
                        .Select nRow, 0, nRow, .Cols - 1
                    End If
                    prog1.Value = nRow
                    XCOUNT.Caption = nRow
                Next nRow
            End If
        End If
    End If
    
    If Col = 19 Then
        If Not .IsSubtotal(Row) Then
            If .TextMatrix(Row, 16) <> "" Then
                nRateOkaz = 0
                nRateOkaz = Round(((Val(grid1.TextMatrix(grid1.Row, 11)) - Val(grid1.TextMatrix(grid1.Row, 19))) / Val(grid1.TextMatrix(grid1.Row, 11)) * 100), 2)
                If Check2.Value = 0 Then
                    con.Execute " update   file1_10 set OKAZ = " & nRateOkaz & " , price2 = " & Val(grid1.TextMatrix(grid1.Row, 19)) & " where model = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
                Else
                    nRateOkaz = Round(((Val(grid1.TextMatrix(grid1.Row, 11)) - Val(grid1.TextMatrix(grid1.Row, 19))) / Val(grid1.TextMatrix(grid1.Row, 11)) * 100), 2)
                    con.Execute " update   file1_10 set OKAZ_2 = " & nRateOkaz & " ,price_2 = " & Val(grid1.TextMatrix(grid1.Row, 19)) & " where model = " & MyParn(grid1.TextMatrix(grid1.Row, 30))
                End If
                .TextMatrix(Row, 17) = nRateOkaz
            End If
        End If
    End If
    cStr2 = ""
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND MOSM = " & MyParn(xMosm.BoundText)
    If Len(xModelFact.text) = 6 Then
        If xModelFact.text <> "" Then cStr2 = cStr2 & " AND SUBSTRING(FILE1_10.MODELFACT0,1,6) = " & MyParn(xModelFact.text)
    Else
        If xModelFact.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    End If
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FACT = " & MyParn(xFact.BoundText)
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub DoFixPrice2(cModel, nOkaz, nRow)
    Dim SubItemTable As New ADODB.Recordset
    Dim nRem As Double, nPrice As Double
    Dim nFPrice As Double
    cStr1 = " SELECT * FROM FILE1_10  WHERE MODELNO = " & MyParn(cModel)
    SubItemTable.Open cStr1, con, adOpenForwardOnly, adLockPessimistic, adCmdText


'   SubItemTable.Open cStr1, con, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    SubItemTable.MoveFirst
        Do While Not SubItemTable.EOF
            If nOkaz > 0 Then
                nPrice = Val(SubItemTable!price & "") * (100 - nOkaz) / 100
                nRem = nPrice - Int(nPrice)
                Select Case nRem
                    Case Is = 0
                        nPrice = Int(nPrice)
                    Case Is <= 0.5
                        nPrice = Int(nPrice) + 0.5
                    Case Is < 1
                        nPrice = Int(nPrice) + 1
                End Select
                If nPrice > SubItemTable!price Then nPrice = SubItemTable!price
            Else
                nPrice = 0
            End If
            If nFPrice = 0 Then nFPrice = nPrice
            If Check2.Value = 0 Then
                SubItemTable!PRICE2 = nPrice
                SubItemTable!okaz = nOkaz
            Else
                SubItemTable!PRICE_2 = nPrice
                SubItemTable!okaz_2 = nOkaz
            End If
            SubItemTable!IsDate = Date
            SubItemTable!IsDate1 = Date
            SubItemTable!IsDate2 = Date
            SubItemTable!IsNEW1 = 1
            SubItemTable!ISNEW2 = 1
            SubItemTable!ISNEW3 = 1
            SubItemTable!ISNEW4 = 1
            SubItemTable!ISNEW5 = 1
            SubItemTable.Update
            SubItemTable.MoveNext
        Loop
        grid1.TextMatrix(nRow, 17) = Format(nFPrice, "#0.00")
        For sbr = 3 To nCountBranch
            cBr = RetZero(sbr, 2)
            con.Execute " INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(cModel)
        Next sbr

        For sbr = 61 To nCountBranch_fr
            cBr = RetZero(sbr, 2)
            con.Execute " INSERT INTO NEW_ITEM_fr (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(cModel)
        Next sbr

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
        myLoadPicture .TextMatrix(NewRow, 30)
    End If
End With

End Sub
Private Sub SubEditModel(pFiled, pText)
Dim cStrUpdate As String
Dim cStrDEl As String
Dim cStrUpdate_fr As String
Dim cStrDEl_fr As String
Dim I As Double
Dim BranchTrans As New ADODB.Recordset
BranchTrans.Open "SELECT CODE FROM BRANCH WHERE IPSQL IS NOT NULL ", con, adOpenStatic, adLockReadOnly, adCmdText

With grid1
    prog1.Value = 0
    prog1.Min = 0
    prog1.Max = .Rows
    For nRow = 1 To .Rows - 1
        I = I + 1
        If .TextMatrix(nRow, 30) <> "" Then
            con.Execute " update file1_10 SET " & pFiled & " = " & addvalue(pText) & " where modelno = " & MyParn(.TextMatrix(nRow, 30))
'            cStrDEl = cStrDEl & " DELETE FROM NEW_ITEM FROM NEW_ITEM INNER JOIN  FILE1_10 ON NEW_ITEM.ITEM = FILE1_10.ITEM  where modelno = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
'            BranchTrans.MoveFirst
'            Do While Not BranchTrans.EOF
'                cBr = RetZero(sbr, 2)
'                cStrUpdate = cStrUpdate & "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(BranchTrans!CODE) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
'                BranchTrans.MoveNext
'            Loop
'
'            cStrDEl_fr = cStrDEl_fr & " DELETE FROM NEW_ITEM_fr FROM NEW_ITEM_Fr INNER JOIN  FILE1_10 ON NEW_ITEM_fr.ITEM = FILE1_10.ITEM  where modelno = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
'            For sbr = 61 To nCountBranch_fr
'                cBr = RetZero(sbr, 2)
'                cStrUpdate_fr = cStrUpdate_fr & "  INSERT INTO NEW_ITEM_fr (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(.TextMatrix(nRow, 30)) & " ; "
'            Next sbr
'            If i = 100 Then
''                con.ConnectionTimeout = 600
'                con.Execute cStrDEl
''               con.ConnectionTimeout = 600
'                con.Execute cStrUpdate
'                cStrUpdate = ""
'                cStrDEl = ""
'
''                con.ConnectionTimeout = 600
'                con.Execute cStrDEl_fr
''                con.ConnectionTimeout = 600
'                con.Execute cStrUpdate_fr
'                cStrUpdate_fr = ""
'                cStrDEl_fr = ""
'
'                i = 0
'            End If
        End If
        prog1.Value = nRow
    Next nRow
'    con.ConnectionTimeout = 600
'    If cStrDEl <> "" Then con.Execute cStrDEl
'    If cStrUpdate <> "" Then con.Execute cStrUpdate
'
''    con.ConnectionTimeout = 600
'    If cStrDEl_fr <> "" Then con.Execute cStrDEl_fr
'    If cStrUpdate_fr <> "" Then con.Execute cStrUpdate_fr
End With
End Sub
Private Sub cmd_save6_Click()
    SubEditModel "ISNODEM", xEdit_ISNODEM.Value
End Sub
Private Sub cmd_save7_Click()
    SubEditModel "ISONLINE", xEdit_ISONLINE.Value
End Sub
Private Sub cmd_save8_Click()
    SubEditModel "Outlet", xEdit_Outlet.Value
End Sub

