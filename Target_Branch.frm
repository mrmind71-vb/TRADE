VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Target_Branch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„»Ì⁄«  «·„” Âœ›… «·‘Â—Ì… ··›—Ê⁄"
   ClientHeight    =   13080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15600
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
   ScaleHeight     =   13080
   ScaleWidth      =   15600
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox XTEXT1 
      Alignment       =   1  'Right Justify
      Height          =   1050
      Left            =   15075
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   7290
      Width           =   5460
   End
   Begin VB.TextBox XTEXT3 
      Alignment       =   1  'Right Justify
      Height          =   1140
      Left            =   15075
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   8415
      Width           =   5460
   End
   Begin VB.CommandButton CMD_ZONE 
      Caption         =   " ÕœÌÀ «·„‰«ÿﬁ «·Õ«·Ì…"
      Height          =   555
      Left            =   7380
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   1170
      Width           =   1860
   End
   Begin Threed.SSCommand CMD_AREA 
      Height          =   510
      Left            =   18045
      TabIndex        =   41
      Top             =   1350
      Visible         =   0   'False
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   900
      _Version        =   196610
      BackColor       =   13822956
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "⁄—÷ ≈Ã„«·Ï „‰«ÿﬁ"
      PictureAlignment=   1
   End
   Begin ComctlLib.ProgressBar prog1 
      Height          =   195
      Left            =   2205
      TabIndex        =   36
      Top             =   1620
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
   Begin Threed.SSCommand cmd_excel 
      Height          =   555
      Left            =   4545
      TabIndex        =   35
      Top             =   990
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   979
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
      Picture         =   "Target_Branch.frx":0000
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand CMD_ADDBRANCH 
      Height          =   600
      Left            =   2205
      TabIndex        =   34
      Top             =   990
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   1058
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Target_Branch.frx":27FB
      Caption         =   "≈÷«›… «·›—Ê⁄ "
      Alignment       =   1
      PictureAlignment=   3
   End
   Begin VB.Frame FRM_CLOSED 
      Height          =   1815
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   0
      Width           =   2085
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·„” ‰œ „€·ﬁ"
         Enabled         =   0   'False
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
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   135
         Visible         =   0   'False
         Width           =   1410
      End
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         TabIndex        =   29
         Top             =   1125
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
         Picture         =   "Target_Branch.frx":5003
         Alignment       =   4
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   1035
         TabIndex        =   30
         Top             =   135
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
         Picture         =   "Target_Branch.frx":75CF
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   90
         TabIndex        =   31
         Top             =   135
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
         Picture         =   "Target_Branch.frx":9C47
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame Frame9 
      Height          =   870
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   0
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
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   495
         Width           =   1815
      End
      Begin VB.Label XTIME1 
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
         TabIndex        =   24
         Top             =   135
         Width           =   2220
      End
      Begin VB.Label xUserName 
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
         TabIndex        =   23
         Top             =   135
         Width           =   1815
      End
      Begin VB.Label xusercode 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   -270
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   9540
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   5505
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Target_Branch.frx":C288
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Target_Branch.frx":C3D2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   2790
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Target_Branch.frx":EC6C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4095
         Picture         =   "Target_Branch.frx":11218
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   7650
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   1500
      Begin VB.CommandButton CmdUndo 
         Height          =   375
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Target_Branch.frx":139EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   540
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdSave 
         Height          =   375
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Target_Branch.frx":15F64
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   15600
      TabIndex        =   6
      Top             =   12225
      Width           =   15600
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   0
         Width           =   3975
         Begin VB.CommandButton cmdNext 
            Height          =   375
            Left            =   2985
            Picture         =   "Target_Branch.frx":163A6
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "«· «·Ì"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdPrevious 
            Height          =   375
            Left            =   2070
            Picture         =   "Target_Branch.frx":18966
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "«·”«»ﬁ"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdLast 
            Height          =   375
            Left            =   915
            Picture         =   "Target_Branch.frx":1AF39
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "√ŒÌ—"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdFirst 
            Height          =   375
            Left            =   45
            Picture         =   "Target_Branch.frx":1D613
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "√Ê·"
            Top             =   135
            Width           =   870
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   12780
      Width           =   15600
      _ExtentX        =   27517
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:05 „"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   9540
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   585
      Width           =   5505
      Begin VB.TextBox XRATEDAY 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   135
         Width           =   1290
      End
      Begin VB.TextBox xDate2 
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
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   540
         Width           =   1290
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
         Height          =   315
         Left            =   2835
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1290
      End
      Begin VB.TextBox xDate1 
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
         Left            =   2835
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1290
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·ÌÊ„ "
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
         Left            =   1515
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ : "
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
         Left            =   1470
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   585
         Width           =   870
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   600
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
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
         Left            =   4215
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   930
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5415
      Left            =   45
      TabIndex        =   7
      Top             =   1845
      Width           =   15000
      _cx             =   26458
      _cy             =   9551
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777088
      ForeColorSel    =   4210752
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   375
      Left            =   6435
      Top             =   135
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Height          =   2265
      Left            =   45
      TabIndex        =   39
      Top             =   7290
      Width           =   15000
      _cx             =   26458
      _cy             =   3995
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777088
      ForeColorSel    =   4210752
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
   Begin VSFlex7Ctl.VSFlexGrid grid3 
      Height          =   5370
      Left            =   15120
      TabIndex        =   40
      Top             =   1890
      Width           =   5460
      _cx             =   9631
      _cy             =   9472
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777088
      ForeColorSel    =   4210752
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   375
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
Attribute VB_Name = "Target_Branch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte, cStrBox As String
Public con As New ADODB.Connection
Public bedit As Boolean
Dim CardTable As ADODB.Recordset
Dim dLastdate As String, defBox As String
Dim DocField As String, dDateLast As String
Dim formMode, cFileHeader As String
Dim lCellButton As Boolean
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(5, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xdoc_no.text)

aInsert(1, 0) = "Date1"
aInsert(1, 1) = addDate(xDate1.text)

aInsert(2, 0) = "Date2"
aInsert(2, 1) = addDate(xdate2.text)

On Error GoTo myerror
con.BeginTrans
If xdoc_no.Enabled Then
    
    aInsert(3, 0) = "username"
    aInsert(3, 1) = addstring(cusername)
    aInsert(4, 0) = "time"
    aInsert(4, 1) = "getdate()"
    
    xdoc_no.text = RetZero(Val(Newflag("FILE6_80H", "doc_no", con)))
    aInsert(0, 1) = addstring(xdoc_no.text)
    con.Execute CreateInsert(aInsert, "FILE6_80H")
Else
    aInsert(3, 0) = "username2"
    aInsert(3, 1) = addstring(cusername)
    aInsert(4, 0) = "time2"
    aInsert(4, 1) = "getdate()"
    
    con.Execute CreateUpdate(aInsert, "FILE6_80H", " where doc_no = " & addstring(xdoc_no.text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd()
Dim aInsert(3, 1)
With grid1
    For i = 1 To .Rows - 1
        If Not .IsSubtotal(i) Then
            aInsert(0, 0) = "doc_no"
            aInsert(0, 1) = addstring(xdoc_no.text)
                   
            aInsert(1, 0) = "branch"
            aInsert(1, 1) = addstring(.TextMatrix(i, 1))
            
            aInsert(2, 0) = "VALUE_TARGET"
            aInsert(2, 1) = Val(grid1.TextMatrix(i, 2))
            
            aInsert(3, 0) = "[Row]"
            aInsert(3, 1) = i
    
            If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
                con.Execute CreateInsert(aInsert, "FILE6_80")
            Else
                con.Execute CreateUpdate(aInsert, "FILE6_80", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
            End If
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
    Unload Search3
End If
End Sub

Private Sub CMD_ADDBRANCH_Click()
    cmdSave_Click
    
    con.Execute " delete FROM FILE6_80 WHERE DOC_NO = " & MyParn(xdoc_no.text) & " and branch is null  "
    con.Execute " INSERT INTO FILE6_80  (BRANCH  , DOC_NO, [VALUE_TARGET] , GR_MAN ) SELECT CODE, " & addstring(xdoc_no.text) & " , 0 , BRANCH_GR FROM STORE_BR2 WHERE  branch_gr is not null and (CODE NOT IN  (SELECT BRANCH FROM FILE6_80 AS FILE6_80_1 WHERE DOC_NO = " & MyParn(xdoc_no.text) & " ))"
    
    AddLod_Data cusername, 0, " ≈÷«›… «·›—Ê⁄ " & Me.Caption, con, xdoc_no.text, xDate1.text, , xdate2.text
    
    myloadgrd ""
    myloadgrd3
End Sub

Private Sub CMD_AREA_Click()
    myloadgrd3

End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub CMD_ZONE_Click()
    If bopt2 Then
        If MsgBox("  ÕœÌÀ «·›—Ê⁄ Õ”» «·„‰«ÿﬁ  «·Õ«·Ì…  ", vbYesNo + vbDefaultButton2) = vbYes Then
            con.Execute " UPDATE FILE6_80 SET GR_MAN =  STORE_BR2.BRANCH_GR  FROM FILE6_80 INNER JOIN STORE_BR2 ON STORE_BR2.CODE = FILE6_80.BRANCH  WHERE DOC_NO = " & MyParn(xdoc_no.text)
            myloadgrd ""
            myloadgrd3

        End If
    End If
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE6_80 where Doc_No = " & MyParn(xdoc_no.text)
    con.Execute "Delete  From FILE6_80H where Doc_No = " & MyParn(xdoc_no.text)
    con.CommitTrans
    
    AddLod_Data cusername, 2, " Õ–› „” ‰œ " & Me.Caption, con, xdoc_no.text, xDate1.text, , xdate2.text
    
    CardTable.Requery
    If CardTable.EOF And CardTable.EOF Then
        myDefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xdoc_no.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
    CardTable.MoveFirst
    myload
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT FILE6_80H.Doc_No, Convert(varChar(10),FILE6_80H.Date1,111),Convert(varChar(10),FILE6_80H.Date2,111) " & _
          " FROM FILE6_80H  "
Generalarray(1) = cString
Generalarray(2) = " order by FILE6_80H.Doc_No"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = " «—ÌŒ «·„” ‰œ"
listarray(0, 1) = "##FILE6_80H.Date1## "

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "„‰  «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "≈·Ï  «—ÌŒ «·„” ‰œ"
GrdArray(2, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„ "
Search3.Show 1
End Sub
Private Sub CmdInform_Click()
    CardLookup
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
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
Private Sub CmdNewInv_Click()
    myDefine
    xdoc_no.SetFocus
End Sub
Private Sub cmdSave_Click()
    If Not MYVALID Then Exit Sub
    If Not myreplace Then Exit Sub
    Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
    
    AddLod_Data cusername, 1, " Õ›Ÿ „” ‰œ " & Me.Caption, con, xdoc_no.text, xDate1.text, , xdate2.text
    
    CardTable.Requery
    CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
    myload
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    myDefine
Else
    If xdoc_no.Enabled Then CardTable.MoveLast Else CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
    myload
End If
End Sub

Private Sub Command1_Click()
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Me.Caption = DocTitle
openCon con

CMD_ADDBRANCH.Visible = bSupermode Or lSupperVisor

CMD_ZONE.Visible = bSupermode
cStrBox = StrBox
FRM_CLOSED.Visible = Not lIsBranchStore
FRM_CLOSED.Visible = lSupperVisor Or (bopt2 And (cBranch = "00" Or cBranch > "60"))

CmdSave.Visible = lSupperVisor
CmdDelInv.Visible = lSupperVisor

XTEXT1.Visible = lSupperVisor 'Or bSupermode
XTEXT3.Visible = lSupperVisor 'Or bSupermode


Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM FILE6_80H  ORDER BY DOC_NO", con, adOpenStatic, adLockOptimistic, adCmdText
cFileHeader = "FILE6_80H"
Set grid1.DataSource = data1
data1.ConnectionString = strCon

Set GRID3.DataSource = data3
data3.ConnectionString = strCon

StatusBar1.Panels.Add 1
StatusBar1.Panels(1).Width = 2000


GRID3.Rows = 1
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    Fixgrd
    myDefine
End If
FixGrid3
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
    If .Col = 10 Then con.Execute " UPDATE FILE6_80 SET FILE6_80.GR_MAN   = " & addstring(.TextMatrix(.Row, 10)) & "  WHERE ID = " & Val(.TextMatrix(.Row, .Cols - 1))
End With
End Sub
Private Sub grid1_DblClick()
With grid1
    VsTSalesday.XSTORE.BoundText = .TextMatrix(.Row, 1)
    VsTSalesday.xDate1.text = xDate1.text
    VsTSalesday.xdate2.text = xdate2.text
    VsTSalesday.Show
End With
End Sub

Private Sub grid1_EnterCell()
If grid1.Col = 2 Or grid1.Col = 10 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
    
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.SetFocus
    grid1.Select 1, 0
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from FILE6_80 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 46 And Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        grid1.RemoveItem Row
    End If
End If
End Sub
Private Function MYVALID() As Boolean
If Trim(xdoc_no.text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If
If Not IsDate(xDate1.text) Then
    MsgBox " «—ÌŒ «Ê· €Ì— ”·Ì„"
    Exit Function
End If

If Not IsDate(xdate2.text) Then
    MsgBox " «—ÌŒ «Œ— €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub myload()
    If CardTable!ISCLOSED Then
        cmd_closed.BackColor = &H8080FF
        cmd_closed.Caption = "› Õ „” ‰œ"
        xClosed.Value = ssCBChecked
        xClosed.Visible = True
    Else
        cmd_closed.BackColor = vbGreen
        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
        xClosed.Value = ssCBUnchecked
        xClosed.Visible = False
    End If
    
    xusername.Caption = CardTable!UserName & ""
    xUserName2.Caption = CardTable!USERNAME2 & ""
    XTIME1.Caption = CardTable!Time & ""
    XTIME2.Caption = CardTable!TIME2 & ""

    xdoc_no.text = CardTable!doc_no
    xDate1.text = Format(CardTable!Date1, "dd-mm-yyyy")
    xdate2.text = Format(CardTable!date2, "dd-mm-yyyy")
    Handlecontrols LoadMode
    myloadgrd ""
    myloadgrd3
End Sub
Private Sub myloadgrd(pGr)
Dim cFiled1 As String, cFiled2 As String
Dim dDate11 As Date, dDate22 As Date
dDate11 = DateAdd("yyyy", -1, DateValue(xDate1.text))
If IsDate(xdate2.text) And Date <= DateValue(xdate2.text) Then
    dDate22 = DateAdd("yyyy", -1, DateValue(Date))
Else
    dDate22 = DateAdd("yyyy", -1, DateValue(xdate2.text))
End If
cFiled1 = ",(SELECT  SUM(T_TOTAL - DISCOUNT)   FROM T_SALESDOC_ALL WHERE T_SALESDOC_ALL.STORE =  FILE6_80.BRANCH AND DATE >= " & DateSq(xDate1.text) & " AND DATE <= " & DateSq(xdate2.text) & " )"
cFiled2 = ",(SELECT  SUM(T_TOTAL - DISCOUNT)   FROM T_SALESDOC_ALL WHERE T_SALESDOC_ALL.STORE =  FILE6_80.BRANCH AND DATE >= " & DateSq(dDate11) & " AND DATE <= " & DateSq(dDate22) & " )"

'    .FormatString = "«·›—⁄|" & "«·„” Âœ› «·‘Â—Ï|"
'    .TextMatrix(0, 0) = "«·›—⁄"
'    .TextMatrix(0, 1) = "«·„” Âœ› «·‘Â—Ï"
'    .TextMatrix(0, 2) = "«·„»Ì⁄« "
'    .TextMatrix(0, 3) = "«·›—ﬁ"
'    .TextMatrix(0, 4) = "‰”»… "
'    .TextMatrix(0, 5) = "›—ﬁ «·‰”»… "
'    .TextMatrix(0, 6) = "„»Ì⁄«  «·⁄«„ «·”«»ﬁ"
'    .TextMatrix(0, 7) = "‰”»… „»Ì⁄«  «·› —… ··› —… „‰ «·⁄«„ «·”«»ﬁ"

With grid1
'                           0                   1                       2         3   4   5         6          7
    cString = "SELECT STORE_BR2.FLAG , FILE6_80.[BRANCH], FILE6_80.value_TARGET  " & cFiled1 & " , 0 , 0 , 0  " & cFiled2 & " , 0 , 0, 0 " & _
              " ,FILE6_80.gr_man, FILE6_80.ID " & _
              " FROM FILE6_80 INNER JOIN STORE_BR2 ON STORE_BR2.CODE  = FILE6_80.BRANCH " & _
              " where FILE6_80.doc_no = " & MyParn(xdoc_no.text)
    If Not bOpt5 Then cString = cString & " AND FILE6_80.[BRANCH] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
'    If pGr <> "" Then cString = cString & " AND STORE_BR2.BRANCH_GR =  " & Val(pGr)
    If pGr <> "" Then cString = cString & " AND FILE6_80.GR_MAN =  " & Val(pGr)
    cString = cString & " Order by STORE_BR2.FLAG,FILE6_80.[BRANCH]"
    data1.RecordSource = cString
    data1.Refresh

    XTEXT1.text = cString


End With
CalcTotals
Fixgrd
End Sub
Sub myloadgrd3()
'    On Error GoTo myerror
''''    cString = " SELECT STORE_BR2.BRANCH_GR , BRANCH_GR.DESCA, SUM(FILE6_80.value_TARGET  ) " & _
''''                " ,(SELECT  SUM(T_TOTAL - DISCOUNT)   FROM T_SALESDOC_ALL INNER JOIN STORE_BR2 AS STORE_BR2_2 ON STORE_BR2_2.CODE = T_SALESDOC_ALL.STORE WHERE STORE_BR2_2.BRANCH_GR =  STORE_BR2.BRANCH_GR AND DATE >= " & DateSq(xDate1.text) & "  AND DATE <= " & DateSq(xdate2.text) & ")" & _
''''                " FROM FILE6_80 INNER JOIN STORE_BR2 ON FILE6_80.BRANCH = STORE_BR2.CODE INNER JOIN BRANCH_GR ON BRANCH_GR.CODE = FILE6_80.GR_MAN  where FILE6_80.doc_no = " & MyParn(xdoc_no.text)
''''    If Not bOpt5 Then cString = cString & " AND FILE6_80.[BRANCH] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
''''    cString = cString & " GROUP BY STORE_BR2.BRANCH_GR , BRANCH_GR.DESCA Order by STORE_BR2.BRANCH_GR"
''''    data3.RecordSource = cString
    
    '                           0                       1               2
    cString = " SELECT FILE6_80.GR_MAN , BRANCH_GR.DESCA, SUM(FILE6_80.value_TARGET  ) " & _
                " ,(SELECT  SUM(T_TOTAL - DISCOUNT)   FROM T_SALESDOC_ALL INNER JOIN FILE6_80 AS FILE6_80_2 ON FILE6_80_2.BRANCH = T_SALESDOC_ALL.STORE  WHERE  FILE6_80_2.GR_MAN =  FILE6_80.GR_MAN  AND FILE6_80_2.DOC_NO = " & MyParn(xdoc_no.text) & "  AND DATE >= " & DateSq(xDate1.text) & "  AND DATE <= " & DateSq(xdate2.text) & ")" & _
                " FROM FILE6_80 INNER JOIN STORE_BR2 ON FILE6_80.BRANCH = STORE_BR2.CODE INNER JOIN BRANCH_GR ON BRANCH_GR.CODE = FILE6_80.GR_MAN  where FILE6_80.doc_no = " & MyParn(xdoc_no.text)
    If Not bOpt5 Then cString = cString & " AND FILE6_80.[BRANCH] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    cString = cString & " GROUP BY FILE6_80.GR_MAN , BRANCH_GR.DESCA   Order by  FILE6_80.GR_MAN"
    data3.RecordSource = cString
    XTEXT3.text = cString
    
    data3.Refresh
    FixGrid3
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub myDefine()
    xusername.Caption = ""
    xUserName2.Caption = ""
    XTIME1.Caption = ""
    XTIME2.Caption = ""
    cmd_closed.BackColor = &H8080FF
    cmd_closed.Caption = "› Õ „” ‰œ"
    xClosed.Visible = False
    xClosed.Value = ssCBUnchecked

    xdoc_no.text = RetZero(Val(Newflag("FILE6_80H", "doc_no", con)))
    xDate1.text = "1-" & Month(Date) & " -" & Year(Date)
    xdate2.text = DateAdd("D", -1, DateAdd("M", 1, DateValue(xDate1.text)))
       
    XTEXT1.text = ""
    XTEXT3.text = ""

grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = defBox

GRID3.Rows = 1

Handlecontrols DefineMode
CalcTotals
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
CmdDelInv.Enabled = (nMode = LoadMode) And (xClosed.Value = 0)
cmdPrevious.Enabled = (nMode = LoadMode)
xdoc_no.Enabled = (nMode = DefineMode)
CmdSave.Enabled = (xClosed.Value = 0)
End Sub
Private Sub GRID3_DblClick()
    myloadgrd (GRID3.TextMatrix(GRID3.Row, 0))
End Sub

Private Sub xDate1_LostFocus()
    If IsDate(xDate1.text) Then xdate2.text = DateAdd("D", -1, DateAdd("M", 1, DateValue(xDate1.text)))
End Sub

Private Sub xDoc_No_LostFocus()
If Trim(xdoc_no.text) = "" Then Exit Sub
xdoc_no.text = RetZero(xdoc_no.text)
CardTable.Find "Doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function CalcTotals()
Dim nTotal As Double
With grid1
For i = 1 To grid1.Rows - 2
    nTotal = nTotal + Round(Val(grid1.TextMatrix(i, 1)), 2)
Next
StatusBar1.Panels(1).text = "≈Ã„«·Ì «· «—ÃÌ  : " & Format(nTotal, "Fixed")
End With
End Function
Private Sub xDoc_No_Validate(Cancel As Boolean)
If xdoc_no.text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
Dim nDay As Double, nRateDay As Double, nMaxDay As Double
nDay = Day(Date)
If IsDate(xdate2.text) And Date <= DateValue(xdate2.text) Then
    nMaxDay = Day(xdate2.text)
    nRateDay = Round((nDay / nMaxDay) * 100, 2)
    XRATEDAY.text = nRateDay
Else
    nRateDay = 100
    XRATEDAY.text = nRateDay

End If


With GRID2
    .ExplorerBar = flexExSortShow
    .Cols = 10
    .Rows = 1
    .RowHeight(0) = 700
    .WordWrap = True
    .ColHidden(0) = True
    .TextMatrix(0, 1) = ""
    .TextMatrix(0, 2) = "«·„” Âœ› «·‘Â—Ï"
    .TextMatrix(0, 3) = "«·„»Ì⁄« "
    .TextMatrix(0, 4) = "«·›—ﬁ"
    .TextMatrix(0, 5) = "‰”»… "
    .TextMatrix(0, 6) = "›—ﬁ «·‰”»… "
    .TextMatrix(0, 7) = "„»Ì⁄«  «·⁄«„ «·”«»ﬁ"
    .TextMatrix(0, 8) = "‰”»… „»Ì⁄«  «·› —… ··› —… „‰ «·⁄«„ «·”«»ﬁ"
    
    .ColHidden(9) = True
    .ColWidth(1) = 2000
    .ColWidth(2) = 1500
    
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500

End With
With grid1
    .ExplorerBar = flexExSortShow
    .Cols = 13
    .RowHeight(0) = 700
    .WordWrap = True
    .ColHidden(0) = True
    .TextMatrix(0, 1) = "«·›—⁄"
    .TextMatrix(0, 2) = "«·„” Âœ› «·‘Â—Ï"
    .TextMatrix(0, 3) = "«·„»Ì⁄« "
    .TextMatrix(0, 4) = "«·›—ﬁ"
    .TextMatrix(0, 5) = "‰”»… "
    .TextMatrix(0, 6) = "›—ﬁ «·‰”»… "
    .TextMatrix(0, 7) = "„»Ì⁄«  «·⁄«„ «·”«»ﬁ"
    .TextMatrix(0, 8) = "‰”»… „»Ì⁄«  «·› —… ··› —… „‰ «·⁄«„ «·”«»ﬁ"
    .TextMatrix(0, 9) = "„»Ì⁄«  ÌÊ„Ï „” Âœ›"
    .TextMatrix(0, 10) = "„ Ê”ÿ „»Ì⁄«  ÌÊ„Ï"
    .ColComboList(11) = StrList("select code , desca from BRANCH_GR", con)
    .ColComboList(1) = StrList("SELECT CODE , DESCA FROM STORE_BR2 ", con)
    .ColWidth(1) = 1600
    .ColWidth(2) = 1200
    
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1700
    .ColHidden(12) = True
    If .Rows > 1 Then
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 2, "#,###,##0", vbYellow, "", True, " "
        .Subtotal flexSTSum, -1, 3, "#,###,##0", vbYellow, "", True, " "
        .Subtotal flexSTSum, -1, 4, "#,###,##0", vbYellow, "", True, " "
        .Subtotal flexSTSum, -1, 7, "#,###,##0", vbYellow, "", True, " "
        .TextMatrix(.Rows - 1, 1) = "„»Ì⁄«  «·‘—ﬂ…"
        
        If bOpt5 Then
            .Subtotal flexSTSum, 0, 2, "#,###,##0", vbYellow, "", True
            .Subtotal flexSTSum, 0, 3, "#,###,##0", vbYellow, "", True
            .Subtotal flexSTSum, 0, 4, "#,###,##0", vbYellow, "", True
            .Subtotal flexSTSum, 0, 7, "#,###,##0", vbYellow, "", True
        End If
                        
        For nRow = 1 To .Rows - 1
            If .ValueMatrix(nRow, 2) > 0 Then
                .TextMatrix(nRow, 4) = Round(.ValueMatrix(nRow, 2) - .ValueMatrix(nRow, 3), 2)
                .TextMatrix(nRow, 5) = Round((.ValueMatrix(nRow, 3) / .ValueMatrix(nRow, 2)) * 100, 2)
                .TextMatrix(nRow, 6) = Round(Val(.TextMatrix(nRow, 5)) - nRateDay, 2)
                .TextMatrix(nRow, 7) = Round(.ValueMatrix(nRow, 7), 2)
                If .ValueMatrix(nRow, 7) > 0 Then .TextMatrix(nRow, 8) = Round(((.ValueMatrix(nRow, 3) / .ValueMatrix(nRow, 7)) * 100) - 100, 2)
                
                
                
                If DateValue(Date) <= DateValue(xdate2.text) And DateValue(Date) >= DateValue(xDate1.text) Then
                    If (nMaxDay - nDay) > 0 Then .TextMatrix(nRow, 9) = Round(Val(.TextMatrix(nRow, 4)) / (nMaxDay - nDay), 2)
                End If
            End If
            If DateValue(Date) <= DateValue(xdate2.text) And DateValue(Date) >= DateValue(xDate1.text) Then
                If (nMaxDay - nDay) > 0 Then .TextMatrix(nRow, 10) = Round(Val(.TextMatrix(nRow, 3)) / (nDay), 2)
            End If
        
        
        
        Next nRow
                
        If bOpt5 Then
            For nRow = 1 To .Rows - 2
                If .IsSubtotal(nRow) Then
                    If .TextMatrix(nRow, 0) = "Total 1" Then
                        GRID2.AddItem ""
                        GRID2.TextMatrix(GRID2.Rows - 1, 1) = "≈Ã„«·Ï ›—Ê⁄"
                        For nCol = 2 To .Cols - 4
                            GRID2.TextMatrix(GRID2.Rows - 1, nCol) = .TextMatrix(nRow, nCol)
                        Next nCol
'                        .RowHidden(nRow) = True
                    End If
                
                    If .TextMatrix(nRow, 0) = "Total 2" Then
                        GRID2.AddItem ""
                        GRID2.TextMatrix(GRID2.Rows - 1, 1) = "≈Ã„«·Ï «„«‰«  "
                        For nCol = 2 To .Cols - 4
                            GRID2.TextMatrix(GRID2.Rows - 1, nCol) = .TextMatrix(nRow, nCol)
                        Next nCol
'                        .RowHidden(nRow) = True
                    End If
                
                    If .TextMatrix(nRow, 0) = "Total 3" Then
                        GRID2.AddItem ""
                        GRID2.TextMatrix(GRID2.Rows - 1, 1) = "≈Ã„«·Ï  ÊﬂÌ·«  "
                        For nCol = 2 To .Cols - 4
                            GRID2.TextMatrix(GRID2.Rows - 1, nCol) = .TextMatrix(nRow, nCol)
                        Next nCol
'                        .RowHidden(nRow) = True
                    End If
                End If
            Next nRow
        End If
        
        GRID2.AddItem ""
        GRID2.TextMatrix(GRID2.Rows - 1, 1) = "«·≈Ã„«·Ï "
        For nCol = 2 To .Cols - 4
            GRID2.TextMatrix(GRID2.Rows - 1, nCol) = .TextMatrix(.Rows - 1, nCol)
        Next nCol
'        .RowHidden(.Rows - 1) = True
    End If
    For i = 1 To grid1.Cols - 2
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColHidden(.Cols - 1) = True
    .ColFormat(2) = "#,###,##0"
    .ColFormat(3) = "#,###,##0"
    .ColFormat(4) = "#,###,##0"
    .ColFormat(7) = "#,###,##0"
    For nRow = 1 To .Rows - 1
        If nRow >= .Rows Then Exit For
        If .IsSubtotal(nRow) Then .RemoveItem nRow
    Next nRow
    If .IsSubtotal(.Rows - 1) Then .RemoveItem .Rows - 1
    GRID2.IsSubtotal(GRID2.Rows - 1) = True
End With
With GRID2
    For i = 1 To GRID2.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColHidden(.Cols - 1) = True
    .ColFormat(2) = "#,###,##0"
    .ColFormat(3) = "#,###,##0"
    .ColFormat(4) = "#,###,##0"
    .ColFormat(7) = "#,###,##0"

End With
End Sub
Private Sub cmd_CLOSEDDATE_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sFieldClose = "isclosed"
oClosefrm.sCaption = Me.Caption
oClosefrm.nMode = 0
oClosefrm.Show 1

CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub cmd_open_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sFieldClose = "isclosed"
oClosefrm.sCaption = Me.Caption
oClosefrm.nMode = 1
oClosefrm.Show 1

CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload

End Sub
Private Sub cmd_closed_Click()
If CardTable!ISCLOSED Then
    con.Execute " update " & cFileHeader & " set isclosed = 0 where doc_no = " & MyParn(xdoc_no.text)
    AddLod_Data cusername, 0, " ›«Õ «·„” ‰œ ·· ⁄œÌ· " & Me.Caption, con, xdoc_no.text, xDate1.text, , xdate2.text

Else
    con.Execute " update " & cFileHeader & " set isclosed = 1 where doc_no = " & MyParn(xdoc_no.text)
End If
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload

End Sub

Sub FixGrid3()
With GRID3
    .ExplorerBar = flexExSortShow
    .Cols = 5
    .RowHeight(0) = 700
    .WordWrap = True
    .ColHidden(0) = True
    .TextMatrix(0, 1) = "«·„‰ÿﬁ…"
    .TextMatrix(0, 2) = "«·„” Âœ›"
    .TextMatrix(0, 3) = "«·„»Ì⁄« "
    
    .TextMatrix(0, 4) = "‰”»… "
    .ColWidth(1) = 1800
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 800
    .ColFormat(2) = "#,###,##0"
    .ColFormat(3) = "#,###,##0"
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 2, "#,###,##0", , vbRed, True, " "
    .Subtotal flexSTSum, -1, 3, "#,###,##0", , vbRed, True, " "
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 2)) <> 0 Then
            nRate = Round((Val(.TextMatrix(nRow, 3)) / Val(.TextMatrix(nRow, 2))) * 100, 2)
            .TextMatrix(nRow, 4) = nRate
        End If
    Next nRow
End With

End Sub

Private Sub XTEXT1_GotFocus()
XTEXT1.SelStart = 0
XTEXT1.SelLength = Len(XTEXT1.text)
End Sub

Private Sub XTEXT3_GotFocus()
    XTEXT3.SelStart = 0
    XTEXT3.SelLength = Len(XTEXT1.text)
End Sub

