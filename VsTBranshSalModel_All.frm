VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form VsTfrBranshSalModel_ALL 
   Caption         =   "≈Ã„«·Ï „‘ —Ì«  & —„»Ì⁄«  & —’Ìœ «·„ÊœÌ·«  ··›—Ê⁄ Ê ·· ÊﬂÌ·« "
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   10890
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
   ScaleWidth      =   10890
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1410
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   1350
      Width           =   3795
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   945
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo xage 
         Height          =   315
         Left            =   180
         TabIndex        =   38
         Top             =   585
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
         Left            =   180
         TabIndex        =   39
         Top             =   180
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         Text            =   ""
         RightToLeft     =   -1  'True
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
         Index           =   3
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   990
         Width           =   1140
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
         Index           =   2
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   630
         Width           =   1005
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
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   270
         Width           =   555
      End
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
      Height          =   3885
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   0
      Width           =   4920
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   3525
         Left            =   90
         TabIndex        =   33
         Top             =   270
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   6218
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
   Begin VB.Frame Frame5 
      Height          =   1050
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2835
      Width           =   3840
      Begin VB.CommandButton Command2 
         Caption         =   "ÿ»«⁄…  Ê“Ì⁄ Ê«—œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         TabIndex        =   44
         Top             =   180
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄…  Ê“Ì⁄ ’«œ—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1980
         TabIndex        =   28
         Top             =   135
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   90
         TabIndex        =   29
         Top             =   630
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16761024
         Text            =   ""
         RightToLeft     =   -1  'True
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1365
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   -45
      Width           =   3840
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "VsTBranshSalModel_All.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "VsTBranshSalModel_All.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2475
         Picture         =   "VsTBranshSalModel_All.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   22
         Top             =   630
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
         Picture         =   "VsTBranshSalModel_All.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   45
         TabIndex        =   25
         Top             =   1125
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   8895
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   -45
      Width           =   11100
      Begin VB.CheckBox xisnodem 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
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
         Left            =   4455
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1755
         Width           =   1500
      End
      Begin VB.CheckBox XOutlet 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1755
         Width           =   1500
      End
      Begin VB.CheckBox Check2_3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "„»Ì⁄« "
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
         Height          =   375
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   585
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox Check2_2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ê«—œ"
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
         Height          =   375
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   585
         UseMaskColor    =   -1  'True
         Value           =   1  'Checked
         Width           =   690
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "„ÊœÌ·«  ·Â« „‘ —Ì«  ›ﬁÿ"
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
         Height          =   375
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1350
         Width           =   2130
      End
      Begin VB.TextBox xquant 
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
         Left            =   8700
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1350
         Width           =   735
      End
      Begin VB.CheckBox Check2_1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "—’Ìœ"
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
         Height          =   375
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   585
         Value           =   1  'Checked
         Width           =   690
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Caption         =   "≈Ã„«·Ï „Ã„Ê⁄… „ÊœÌ·"
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
         Height          =   375
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   180
         Width           =   2400
      End
      Begin VB.TextBox xDate2 
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
         TabIndex        =   23
         Top             =   225
         Width           =   1680
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
         Left            =   7980
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   1455
      End
      Begin VB.TextBox xDescItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EAEAEA&
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   990
         Width           =   2925
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
         TabIndex        =   4
         Top             =   1395
         Width           =   4590
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
         TabIndex        =   3
         Top             =   990
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   6000
         TabIndex        =   7
         Top             =   980
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
         TabIndex        =   8
         Top             =   610
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
         Left            =   2700
         TabIndex        =   9
         Top             =   180
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   2700
         TabIndex        =   10
         Top             =   585
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "—’Ìœ  «ﬂ»— „‰ "
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
         Left            =   9585
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1395
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›‰—…"
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
         TabIndex        =   17
         Top             =   270
         Width           =   450
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
         TabIndex        =   16
         Top             =   225
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
         TabIndex        =   15
         Top             =   615
         Width           =   1320
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
         TabIndex        =   14
         Top             =   990
         Width           =   1410
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
         TabIndex        =   13
         Top             =   1440
         Width           =   1155
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
         Top             =   990
         Width           =   915
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
         Index           =   1
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   585
         Width           =   1230
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   10890
      _ExtentX        =   19209
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTBranshSalModel_All.frx":9583
      Height          =   6345
      Left            =   135
      TabIndex        =   1
      Top             =   3960
      Width           =   19860
      _cx             =   35031
      _cy             =   11192
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
      Begin ComctlLib.ListView ListView1 
         Height          =   30
         Left            =   4140
         TabIndex        =   24
         Top             =   3510
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   53
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   327682
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
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
   Begin MSAdodcLib.Adodc DATA6 
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
   Begin VSFlex7Ctl.VSFlexGrid GRID2 
      Bindings        =   "VsTBranshSalModel_All.frx":9597
      Height          =   1890
      Left            =   8910
      TabIndex        =   26
      Top             =   2070
      Width           =   11085
      _cx             =   19553
      _cy             =   3334
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
End
Attribute VB_Name = "VsTfrBranshSalModel_ALL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StoreTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = Me.Caption
    
    cHead2 = " Õ Ï  «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY")
    Dim cHead3 As String
    cHead3 = XSECTION.text & "  " & xGroup.text & "  " & xFact.text
    
    Load PrintGrd
    PrintGrd.DOPRINT grid1, 0.9, , cHead1, cHead2, cHead3, , True, 9, , Array(1, 2)
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
    Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub Command1_Click()
    doprint2
End Sub
Private Sub Command2_Click()
    doprint3
End Sub
Private Sub Form_Load()
    openCon con
    
    StoreTable.Open "select STORE_BR.CODE , STORE_BR.DESCA FROM STORE_BR WHERE  " & cUserStore, con, adOpenStatic, adLockReadOnly, adCmdText
    
    Set DATA7.Recordset = myRecordSet("Select Code,DescA From STORE_BR WHERE " & cUserStore, con)
    Set xStore.RowSource = DATA7
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    
    
    XMOSM2.AddItem " "
    XMOSM2.AddItem "S"
    XMOSM2.AddItem "W"
    XMOSM2.AddItem "M"
    
    Set DATA11.Recordset = myRecordSet("Select Code,DescA From File0_82  order by Desca", con)
    Set xage.RowSource = DATA11
    xage.ListField = "Desca"
    xage.BoundColumn = "Code"
    
    Set data12.Recordset = myRecordSet("Select Code,DescA From File0_81  order by Desca", con)
    Set xsex.RowSource = data12
    xsex.ListField = "Desca"
    xsex.BoundColumn = "Code"
    
    xDate1.text = Format("1-1-" & Year(Date), "dd-mm-yyyy")
    xDate2.text = Format(Date, "dd-mm-yyyy")
    
    Set data1.Recordset = myRecordSet("Select Code,DescA From File1_10SC order by Desca", con)
    Set XSECTION.RowSource = data1
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
    Set DATA3.Recordset = myRecordSet("Select Code,DescA From File1_50 ORDER BY DESCA", con)
    Set xGroup.RowSource = DATA3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set DATA4.Recordset = myRecordSet("Select mosm ,descA From mosm ORDER BY date DESC ", con)
    Set xMosm.RowSource = DATA4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    Set DATA5.Recordset = myRecordSet("Select code ,desca From fact ORDER BY code ", con)
    Set xFact.RowSource = DATA5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA6
    grid1.Rows = 2
    grid1.FixedRows = 2
    grid1.Cols = 7
    
    Set GRID2.DataSource = data10
    data10.ConnectionString = strCon
    GRID2.Rows = 1
    GRID2.Cols = 7

    MYLOAD2
    '
'    FixGrid
End Sub
Private Sub myload()
Dim I As Double
Dim cStrall  As String
grid1.Rows = 0
grid1.Rows = 2
grid1.FixedRows = 2
grid1.Cols = 6

Dim lOk As Boolean
    StoreTable.MoveFirst
    Do Until StoreTable.EOF
        cWhere = " ( FILE1_11_ALL.TYPE = '2' OR FILE1_11_ALL.TYPE = '7'  OR FILE1_11_ALL.TYPE = 'F'  OR FILE1_11_ALL.TYPE = 'T' ) AND STORE = " & MyParn(StoreTable!CODE)
        cWhere = cWhere & " AND ( DATE >= " & DateSq(xDate1.text) & " AND DATE <= " & DateSq(xDate2.text) & " ) "
        cField2 = cField2 & " , " & myiif(cWhere, "[IN] - [OUT]") & " AS " & addstring(StoreTable!CODE)
        
        cWhere = " FILE1_11_ALL.TYPE = '6' AND STORE = " & MyParn(StoreTable!CODE)
        cWhere = cWhere & " AND ( DATE >= " & DateSq(xDate1.text) & " AND DATE <= " & DateSq(xDate2.text) & " ) "
        cField2 = cField2 & " , " & myiif(cWhere, "[OUT] ")
        
        cWhere = " STORE = " & MyParn(StoreTable!CODE)
        cWhere = cWhere & " AND ( DATE <= " & DateSq(xDate2.text) & " ) "
        cField2 = cField2 & " , " & myiif(cWhere, "[IN] - [OUT]")
        
        
        cField2 = cField2 & " , ' ' "
        StoreTable.MoveNext
    Loop

cWhere = "  (FILE1_11_ALL.TYPE = '2' OR FILE1_11_ALL.TYPE = '7'   ) "
cWhere = cWhere & " AND ( FILE1_11_ALL.DATE >= " & DateSq(xDate1.text) & " AND DATE <= " & DateSq(xDate2.text) & " ) "
cField3 = myiif(cWhere, "[IN]-[OUT]")

cWhere = "  (FILE1_11_ALL.TYPE = '6'   ) "
cWhere = cWhere & " AND ( FILE1_11_ALL.DATE >= " & DateSq(xDate1.text) & " AND DATE <= " & DateSq(xDate2.text) & " ) "
cWhere = cWhere & " AND ( DATE <= " & DateSq(xDate2.text) & " ) "
cField4 = myiif(cWhere, "[OUT]")

cField12 = "DATEDIFF(dd , " & myiif2("type = '2' ", "FILE1_11_ALL.date", "MIN") & " , GETDATE()) AS DATEDIFF"

With grid1
'                           0               1                 2                3               4                        5
If Check1.Value = 0 Then
    cStrall = " SELECT file1_10.ModelNo, file1_10.mosm, FACT.desca , file1_10.modelfact0 , file1_10.desca  , MIN(FILE1_10.PRICE)  " & _
                cField2 & " ,  " & cField3 & " ,  " & cField4 & " , sum([in] - [OUT]) as bal , ' ' , FILE1_10.Okaz_2 ,  " & cField12 & " , FILE1_10.ISNODEM as [»œÊ‰ ÿ·»Ì…]  , file1_10.ModelNo " & _
            " FROM  FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE  WHERE file1_10.ISNOITEM = 0 "
    If xMosm.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "" Then cStrall = cStrall & " AND FILE1_10.MOSM2 = " & MyParn(XMOSM2.text)
    If xModelFact.text <> "" Then cStrall = cStrall & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xGroup.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XOutlet.Value <> 0 Then cStrall = cStrall & " AND FILE1_10.Outlet = 1 "
    If xage.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If xisnodem.Value <> 0 Then cStrall = cStrall & " AND FILE1_10.ISNODEM =  1 "
    If XSECTION.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.[Section] = " & XSECTION.BoundText
'   If Check3.Value <> 0 Then cStrall = cStrall & " AND FILE1_10.[Section] = " & XSECTION.BoundText
    
    If XDESCA.text <> "" Then cStrall = cStrall & " AND file1_10.DESCA LIKE ('%" & XDESCA.text & "%')   "
    cStrall = cStrall & " GROUP BY file1_10.ModelNo, FACT.desca , file1_10.mosm, file1_10.desca  , modelfact0 , FILE1_10.Okaz_2 , FILE1_10.ISNODEM "
    If Val(xquant.text) > 0 Then cStrall = cStrall & " having sum([in] - [OUT]) >=  " & Val(xquant.text)
    cStrall = cStrall & " ORDER BY file1_10.ModelNo "
Else
    cStrall = " SELECT GrModel  , file1_10.mosm, FACT.desca , file1_10.modelfact0 , file1_10.desca  , MIN(FILE1_10.PRICE)  " & _
                cField2 & " ,  " & cField3 & " ,  " & cField4 & " , sum([in] - [OUT]) as bal , ' ' , FILE1_10.Okaz_2 ,  " & cField12 & " , FILE1_10.ISNODEM as [»œÊ‰ ÿ·»Ì…] , file1_10.ModelNo " & _
            " FROM  FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE  WHERE file1_10.ISNOITEM = 0 "
    If xMosm.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.text <> "" Then cStrall = cStrall & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xGroup.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XMOSM2.text <> "" Then cStrall = cStrall & " AND FILE1_10.MOSM2 = " & MyParn(XMOSM2.text)
    If XOutlet.Value <> 0 Then cStrall = cStrall & " AND FILE1_10.Outlet = 1 "
    If xage.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.MODELAGE = " & MyParn(xage.BoundText)
    If xsex.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.MODELSEX = " & MyParn(xsex.BoundText)
    If XSECTION.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.[Section] = " & XSECTION.BoundText
    If XDESCA.text <> "" Then cStrall = cStrall & " AND file1_10.DESCA LIKE ('%" & XDESCA.text & "%')   "
    cStrall = cStrall & " GROUP BY GrModel, FACT.desca , file1_10.mosm, file1_10.desca  , modelfact0 , FILE1_10.Okaz_2 , FILE1_10.ISNODEM , file1_10.ModelNo"
    If Val(xquant.text) > 0 Then cStrall = cStrall & " having sum([in] - [OUT]) >=  " & Val(xquant.text)
    cStrall = cStrall & " ORDER BY GrModel  , FILE1_10.MODELFACT0  "
End If
    con.CommandTimeout = 2000
    Set DATA6.Recordset = myRecordSet(cStrall, con)
End With
FIXGRID
'If grid1.Rows > 0 Then grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FIXGRID()
Dim nColS1 As Double
Dim nColS2 As Double
With grid1
'    .FixedRows = 2
    .RowHeight(0) = 700
    .RowHeight(1) = 700
    .WordWrap = True
    .MergeCells = flexMergeFixedOnly
    
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    
    .MergeRow(0) = True
    .MergeRow(1) = True
    .ColHidden(.Cols - 1) = True
    If Check1.Value = 0 Then
        .TextMatrix(0, 0) = "„ÊœÌ·"
        .TextMatrix(0, 1) = "«·„Ê”„"
        .TextMatrix(0, 2) = "«·„’‰⁄"
        .TextMatrix(0, 3) = "—ﬁ„ „ÊœÌ·"
        .TextMatrix(1, 0) = "„ÊœÌ·"
        .TextMatrix(1, 1) = "«·„Ê”„"
        .TextMatrix(1, 2) = "«·„’‰⁄"
        .TextMatrix(1, 3) = "—ﬁ„ „ÊœÌ·"
        .ColWidth(0) = 0
        .ColWidth(1) = 700
        .ColWidth(2) = 1800
        .ColWidth(3) = 1200
    Else
        .TextMatrix(0, 0) = "«·„Ã„Ê⁄…"
        .TextMatrix(0, 1) = "«·„Ê”„"
        .TextMatrix(0, 2) = "«·„’‰⁄"
        .TextMatrix(0, 3) = "—ﬁ„ „ÊœÌ·"
        .TextMatrix(1, 0) = "«·„Ã„Ê⁄…"
        .TextMatrix(1, 1) = "«·„Ê”„"
        .TextMatrix(1, 2) = "«·„’‰⁄"
        .TextMatrix(1, 3) = "—ﬁ„ „ÊœÌ·"
        .ColWidth(0) = 1000
        .ColWidth(1) = 0
        .ColWidth(2) = 0
        .ColWidth(3) = 1200
    End If
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = "”⁄— „” Â·ﬂ"
    
    .TextMatrix(1, 4) = "«·’‰›"
    .TextMatrix(1, 5) = "”⁄— „” Â·ﬂ"
    .FrozenCols = 6
    .ColWidth(4) = 3000
    .ColWidth(5) = 700
    
    .ColDataType(5) = flexDTDouble
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    For nCol = 6 To .Cols - 8 Step 4
        .TextMatrix(1, nCol) = "Ê«—œ"
        .TextMatrix(1, nCol + 1) = "„»Ì⁄« "
        .TextMatrix(1, nCol + 2) = "—’Ìœ"
        .TextMatrix(1, nCol + 3) = "‰”»…"
        
        .TextMatrix(0, nCol) = GetDesca("SELECT DESCA FROM dbo.STORE_BR WHERE CODE = " & MyParn(.TextMatrix(0, nCol)), con)
        .TextMatrix(0, nCol + 1) = .TextMatrix(0, nCol)
        .TextMatrix(0, nCol + 2) = .TextMatrix(0, nCol)
        .TextMatrix(0, nCol + 3) = .TextMatrix(0, nCol)
        
        .ColWidth(nCol) = 800
        .ColWidth(nCol + 1) = 800
        .ColWidth(nCol + 2) = 800
        .ColWidth(nCol + 3) = 800
        
'        If Check2_1.Value <> 0 Then
'            .ColHidden(nCol) = True
'            .ColHidden(nCol + 1) = True
'            .ColHidden(nCol + 3) = True
'        End If
        
        .ColHidden(nCol + 2) = (Check2_1.Value = 0)
        .ColHidden(nCol) = (Check2_2.Value = 0)
        .ColHidden(nCol + 1) = (Check2_3.Value = 0)
        If (Check2_1.Value = 0) Or (Check2_2.Value = 0) Or (Check2_3.Value = 0) Then
            .ColHidden(nCol + 3) = True
        End If
        .ColDataType(nCol) = flexDTDouble
        .ColDataType(nCol + 1) = flexDTDouble
        .ColDataType(nCol + 2) = flexDTDouble
        .ColDataType(nCol + 3) = flexDTDouble
    Next nCol
    
    .TextMatrix(0, .Cols - 8) = "≈Ã„«·Ï «·›—Ê⁄"
    .TextMatrix(0, .Cols - 7) = "≈Ã„«·Ï «·›—Ê⁄"
    .TextMatrix(0, .Cols - 6) = "≈Ã„«·Ï «·›—Ê⁄"
    .TextMatrix(0, .Cols - 5) = "≈Ã„«·Ï «·›—Ê⁄"
    .TextMatrix(0, .Cols - 4) = "≈Ã„«·Ï «·›—Ê⁄"
    .TextMatrix(0, .Cols - 3) = "≈Ã„«·Ï «·›—Ê⁄"
    .TextMatrix(0, .Cols - 2) = "»œÊ‰ ÿ·»Ì…"
    .TextMatrix(1, .Cols - 2) = "»œÊ‰ ÿ·»Ì…"
    .MergeCol(.Cols - 2) = True
    
    .TextMatrix(1, .Cols - 8) = "Ã „‘ —Ì« "
    .TextMatrix(1, .Cols - 7) = "Ã „»Ì⁄« "
    .TextMatrix(1, .Cols - 6) = "Ã «·—’Ìœ"
    .TextMatrix(1, .Cols - 5) = "Ã ‰”»…"
    .TextMatrix(1, .Cols - 4) = "‰”»… √Êﬂ«“ÊÌ‰2"
    .TextMatrix(1, .Cols - 3) = "⁄„— «·„ÊœÌ·"
    
    .ColWidth(.Cols - 2) = 900
    .ColWidth(.Cols - 4) = 900
    .ColWidth(.Cols - 5) = 900
    .ColWidth(.Cols - 6) = 900
    .ColWidth(.Cols - 7) = 900
    .ColWidth(.Cols - 8) = 900
    .ColHidden(.Cols - 1) = True
'    If Check2.Value <> 0 Then
'        .ColHidden(.Cols - 7) = True
'        .ColHidden(.Cols - 6) = True
'        .ColHidden(.Cols - 4) = True
'   End If
    
    .ColHidden(.Cols - 8) = (Check2_2.Value = 0)
    .ColHidden(.Cols - 7) = (Check2_3.Value = 0)
    .ColHidden(.Cols - 5) = (Check2_1.Value = 0)
    If (Check2_1.Value = 0) Or (Check2_2.Value = 0) Or (Check2_3.Value = 0) Then
        .ColHidden(.Cols - 4) = True
    End If
    
    
    .MergeCol(.Cols - 3) = True
    .MergeCol(.Cols - 4) = True
    .MergeCol(.Cols - 5) = True
    .MergeCol(.Cols - 6) = True
    .MergeCol(.Cols - 7) = True
    .MergeCol(.Cols - 8) = True
    
    For nRow = 2 To .Rows - 1
        For nCol = 5 To .Cols - 5
            .TextMatrix(nRow, nCol) = Round(Val(.TextMatrix(nRow, nCol)), 0)
        Next nCol
        For nCol = 6 To .Cols - 7 Step 4
            nRate = 0
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, nCol + 1)) / Val(.TextMatrix(nRow, nCol)) * 100, 2)
            .TextMatrix(nRow, nCol + 3) = nRate
        Next nCol
        If Val(.TextMatrix(nRow, .Cols - 2)) <> 0 Then .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = &H8080&
    Next nRow
    .SubtotalPosition = flexSTAbove
    If .Rows > 2 Then
    
    For I = 6 To .Cols - 5
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
        If Check1.Value <> 0 Then .Subtotal flexSTSum, 0, I, "#0", &H8080&, , True, "  "
    Next I
    
    If Check1.Value <> 0 Then .Subtotal flexSTCount, 0, 3, , &H8080&, , True, "  "

    For nRow = 2 To .Rows - 1
        If .IsSubtotal(nRow) Then
'            For nCol = 6 To .Cols - 1
'                .TextMatrix(nRow, nCol) = Round(Val(.TextMatrix(nRow, nCol)), 0)
'            Next
            
            For nCol = 6 To .Cols - 9 Step 4
                nRate = 0
                If Val(.TextMatrix(nRow, nCol)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, nCol + 1)) / Val(.TextMatrix(nRow, nCol)) * 100, 2)
                .TextMatrix(nRow, nCol + 3) = nRate
            Next nCol
            If Val(.TextMatrix(nRow, .Cols - 8)) <> 0 Then nRate = Round((Val(.TextMatrix(nRow, .Cols - 7)) / Val(.TextMatrix(nRow, .Cols - 8))) * 100, 2)
            .TextMatrix(nRow, .Cols - 5) = nRate
        
        End If
    Next nRow
    nRow = 2

    For nCol = 6 To .Cols - 7 Step 4
        nRate = 0
        If Val(.TextMatrix(nRow, nCol)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, nCol + 1)) / Val(.TextMatrix(nRow, nCol)) * 100, 2)
        .TextMatrix(nRow, nCol + 3) = nRate
    Next nCol
    
    nRow = 2
    nCol = .Cols - 5
    nRate = 0
    If .Rows > 2 Then
        If Val(.TextMatrix(nRow, .Cols - 8)) <> 0 Then nRate = Round(Val(.TextMatrix(nRow, .Cols - 7)) / Val(.TextMatrix(nRow, .Cols - 8)) * 100, 2)
        .TextMatrix(nRow, nCol) = nRate
    End If
    End If
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    StoreTable.Close
    Set StoreTable = Nothing
    closeCon con
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'            Dim cModelNo As String
'            With grid1
'                If Check1.Value = 0 Then
'                    If .Col = .Cols - 1 Then
'                        For sbr = 2 To nCountBranch
'                            cBr = RetZero(sbr, 2)
'                            con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(grid1.TextMatrix(grid1.Row, 0)), nRec
'                        Next sbr
'                        If Not .IsSubtotal(Row) Then
'                            If .TextMatrix(Row, .Cols - 1) <> "" Then
'                                DoFixPrice2 grid1.TextMatrix(grid1.Row, 0), Val(grid1.TextMatrix(grid1.Row, .Cols - 1)), grid1.Row
'                            End If
'                        End If
'                    End If
'                Else
'                    If .Col = .Cols - 1 Then
'                        cModelNo = GetDesca("SELECT MODELNO FROM FILE1_10 WHERE modelfact0 = " & MyParn(.TextMatrix(.Row, 3)))
'                        For sbr = 2 To nCountBranch
'                            cBr = RetZero(sbr, 2)
'                            con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBr) & " FROM  FILE1_10  WHERE MODELNO = " & MyParn(cModelNo), nRec
'                        Next sbr
'                        If Not .IsSubtotal(Row) Then
'                            If .TextMatrix(Row, .Cols - 1) <> "" Then
'                                DoFixPrice2 cModelNo, Val(grid1.TextMatrix(grid1.Row, .Cols - 1)), grid1.Row
'                            End If
'                        End If
'                    End If
'                End If
'            End With
End Sub

Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
    If OldRow <> NewRow Then
        myLoadPicture .TextMatrix(NewRow, 0)
    End If
End With
End Sub

Private Sub grid1_DblClick()
    If grid1.Row > 2 Then
        If grid1.Col <= 3 Then
            Load ModelMove
            ModelMove.XMODEL.text = grid1.TextMatrix(grid1.Row, 0)
            If grid1.Col > 6 Then ModelMove.xStore.BoundText = grid1.TextMatrix(1, grid1.Col)
            ModelMove.Show
        Else
            VsTfrBranshSalModel_SCAL.sModel = grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            VsTfrBranshSalModel_SCAL.sDate1 = xDate1.text
            VsTfrBranshSalModel_SCAL.sDate2 = xDate2.text
            VsTfrBranshSalModel_SCAL.Show
            MYLOAD2
        End If
    End If
End Sub
Private Sub grid1_EnterCell()
With grid1
    If .Col = .Cols - 1 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Sub DoFixPrice2(cModel, nOkaz, nRow)
    Dim SubItemTable As New ADODB.Recordset
    Dim nRem As Double, nPrice As Double
    Dim nFPrice As Double
    cStr1 = " SELECT * FROM FILE1_10  WHERE MODELNO = " & MyParn(cModel)
    SubItemTable.Open cStr1, con, adOpenForwardOnly, adLockPessimistic, adCmdText
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
                nPrice = SubItemTable!price
            End If
            If nFPrice = 0 Then nFPrice = nPrice
            SubItemTable!PRICE_2 = nPrice
            SubItemTable!okaz_2 = nOkaz
            SubItemTable!IsDate = Date
            SubItemTable!IsDate1 = Date
            SubItemTable!IsDate2 = Date
            SubItemTable.Update
            SubItemTable.MoveNext
        Loop
End Sub
Private Sub doprint2()
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
    cStr1 = "SELECT STORE_BR_1.DESCA AS NAMESTORE2 ,  STORE_BR.DESCA AS NAMESTORE1 ,SUBTRANS.STORE1 , SUBTRANS.STORE2, SUBTRANS.ITEM , SUBTRANS.QUANT , FILE1_10.DESCA , FILE1_10.COLOR , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.MOSM , FILE1_10.MODELFACT0 , FILE1_10.MODEL , FACT.DESCA AS FACTDESCA , FILE1_10.ITEM FROM    SUBTRANS INNER JOIN FILE1_10 ON SUBTRANS.ITEM = FILE1_10.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE INNER JOIN STORE_BR ON SUBTRANS.STORE1 = STORE_BR.CODE INNER JOIN STORE_BR AS STORE_BR_1 ON SUBTRANS.STORE2 = STORE_BR_1.CODE WHERE FILE1_10.MODELNO IS NOT NULL "
    If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.text <> "" Then cStr1 = cStr1 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xGroup.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XSECTION.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[Section] = " & XSECTION.BoundText
    If xStore.BoundText <> "" Then cStr1 = cStr1 & " AND SUBTRANS.[STORE1] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cStr1 = cStr1 & " AND SUBTRANS.[STORE1] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XDESCA.text <> "" Then cStr1 = cStr1 & " AND file1_10.DESCA LIKE ('%" & XDESCA.text & "%')   "
    sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText

    With sourcetable
    Do While Not .EOF
        temptable.AddNew
        
        temptable!str3 = !STORE1
        temptable!str4 = !NAMESTORE1
        temptable!str5 = !STORE2 & " " & !NAMESTORE2
        
        
        temptable!str1 = !FACTDESCA
        temptable!str2 = !modelfact0
        temptable!str17 = !DESCA
        temptable!str15 = !MODEL
        
        temptable!str12 = Trim(!scal)
        temptable!str13 = Trim(!color)
        temptable!val12 = !C_SCAL
        temptable!VAL13 = !Item
        
        temptable!VAL1 = !Quant
        temptable!STR7 = " Ê“Ì⁄…  ÕÊÌ·«  "
        temptable.Update
        .MoveNext
    Loop
    End With

temptable.Requery
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\R_SUBTRANS.rpt"
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Sub FixGrid2()
'               0       1           2       3           4       5       6   7
'   " SELECT  STORE, STOREDESCA, FACT, FACTDESCA, SECTION, SECDESCA, MOSM, TQUANT FROM  QTRANS ORDER BY STORE, STOREDESCA, FACT, FACTDESCA, SECTION, SECDESCA, MOSM "
    With GRID2
        .Cols = 10
        .TextMatrix(0, 1) = "„‰ ›—⁄"
        .TextMatrix(0, 3) = "«·„’‰⁄"
        .TextMatrix(0, 5) = "«·ﬁ”„"
        .TextMatrix(0, 6) = "«·„Ê”„"
        .TextMatrix(0, 7) = "Ã. ﬂ„Ì…"
        
        .TextMatrix(0, 8) = "ÿ»«⁄…"
        .TextMatrix(0, 9) = "Õ–›"
        
        .ColWidth(1) = 1100
        .ColWidth(3) = 1100
        .ColWidth(5) = 1100
        .ColWidth(6) = 1000
        .ColWidth(7) = 800
        
        .ColWidth(8) = 600
        .ColWidth(9) = 600
        
        .ColHidden(0) = True
        .ColHidden(2) = True
        .ColHidden(4) = True
        
        .ColComboList(8) = "..."
        .ColComboList(9) = "..."
    End With
End Sub
Private Sub GRID2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With GRID2
    If .Col = 8 Then
        xFact.BoundText = .TextMatrix(.Row, 2)
        XSECTION.BoundText = .TextMatrix(.Row, 4)
        xStore.BoundText = .TextMatrix(.Row, 0)
        xMosm.BoundText = .TextMatrix(.Row, 6)
        doprint2
    End If
    If .Col = 9 Then
        If MsgBox(" Õ–› «· Ê“Ì⁄…  ", vbYesNo + vbDefaultButton2) = vbYes Then
            con.Execute " DELETE FROM SUBTRANS FROM FILE1_10 INNER JOIN SUBTRANS ON FILE1_10.ITEM = SUBTRANS.ITEM  WHERE STORE1 = " & MyParn(.TextMatrix(.Row, 0)) & " AND FILE1_10.FACT = " & MyParn(.TextMatrix(.Row, 2)) & " AND FILE1_10.FACT = " & MyParn(.TextMatrix(.Row, 2)) & " AND FILE1_10.MOSM = " & MyParn(.TextMatrix(.Row, 6)), nRec
            Inform " Õ–› ⁄œœ  " & nRec
            MYLOAD2
        End If
    End If
    
End With
End Sub
Private Sub grid2_EnterCell()
With GRID2
    If .Col >= 8 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Sub MYLOAD2()
    If bOpt5 Then
        data10.RecordSource = " SELECT  STORE, STOREDESCA, FACT, FACTDESCA, SECTION, SECDESCA, MOSM, TQUANT FROM  QTRANS ORDER BY STORE, STOREDESCA, FACT, FACTDESCA, SECTION, SECDESCA, MOSM "
    Else
        data10.RecordSource = " SELECT  STORE, STOREDESCA, FACT, FACTDESCA, SECTION, SECDESCA, MOSM, TQUANT FROM  QTRANS WHERE STORE IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " )  ORDER BY STORE, STOREDESCA, FACT, FACTDESCA, SECTION, SECDESCA, MOSM "
    End If
    data10.Refresh
    FixGrid2
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
Private Sub doprint3()
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
    cStr1 = "SELECT STORE_BR_1.DESCA AS NAMESTORE2 ,  STORE_BR.DESCA AS NAMESTORE1 ,SUBTRANS.STORE1 , SUBTRANS.STORE2, SUBTRANS.ITEM , SUBTRANS.QUANT , FILE1_10.DESCA , FILE1_10.COLOR , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.MOSM , FILE1_10.MODELFACT0 , FILE1_10.MODEL , FACT.DESCA AS FACTDESCA FROM    SUBTRANS INNER JOIN FILE1_10 ON SUBTRANS.ITEM = FILE1_10.ITEM INNER JOIN FACT ON FILE1_10.code = FACT.CODE INNER JOIN STORE_BR ON SUBTRANS.STORE1 = STORE_BR.CODE INNER JOIN STORE_BR AS STORE_BR_1 ON SUBTRANS.STORE2 = STORE_BR_1.CODE WHERE FILE1_10.MODELNO IS NOT NULL "
    If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.text <> "" Then cStr1 = cStr1 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xGroup.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XSECTION.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[Section] = " & XSECTION.BoundText
    If xStore.BoundText <> "" Then cStr1 = cStr1 & " AND SUBTRANS.[STORE2] = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cStr1 = cStr1 & " AND SUBTRANS.[STORE2] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    If XDESCA.text <> "" Then cStr1 = cStr1 & " AND file1_10.DESCA LIKE ('%" & XDESCA.text & "%')   "
    sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText

    With sourcetable
    Do While Not .EOF
        temptable.AddNew
        
        temptable!str3 = !STORE1
        temptable!str4 = !NAMESTORE1
        temptable!str5 = !STORE2 & " " & !NAMESTORE2
        temptable!str1 = !FACTDESCA
        temptable!str2 = !modelfact0
        temptable!str17 = !DESCA
        temptable!str15 = !MODEL
        temptable!str12 = Trim(!scal)
        temptable!str13 = Trim(!color)
        temptable!val12 = !C_SCAL
        
        temptable!VAL1 = !Quant
        temptable!STR7 = " Ê“Ì⁄…  ÕÊÌ·«  "
        temptable.Update
        .MoveNext
    Loop
    End With

temptable.Requery
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\R_SUBTRANS.rpt"
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub

