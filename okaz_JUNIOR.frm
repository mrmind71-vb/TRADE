VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form okaz_JUNIOR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " €Ì— «”⁄«— „ÊœÌ·«  - Œ’„ Ê «÷«›…"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18030
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
   ScaleHeight     =   9855
   ScaleWidth      =   18030
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame13 
      Height          =   690
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   60
      Top             =   8685
      Width           =   2490
      Begin VB.Label Label15 
         Caption         =   "⁄œœ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   225
         Width           =   600
      End
      Begin VB.Label xqty 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   180
         Width           =   1500
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "ﬁÌ„… «·»÷«⁄… "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4950
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   45
      Width           =   2085
      Begin VB.TextBox XTCOST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   225
         Width           =   1905
      End
   End
   Begin VB.Frame Frame11 
      Height          =   1230
      Left            =   5175
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   810
      Width           =   2490
      Begin VB.TextBox XBARCODE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   630
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo XSEC 
         Height          =   315
         Left            =   90
         TabIndex        =   51
         Top             =   225
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "»«—ﬂÊœ : "
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
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ”„ :"
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
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   225
         Width           =   525
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   10170
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8685
      Width           =   5010
      Begin Threed.SSCommand cmdFix 
         Height          =   510
         Left            =   1665
         TabIndex        =   15
         Top             =   135
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   900
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
         Picture         =   "okaz_JUNIOR.frx":0000
         Caption         =   "≈⁄«œ… ÷»ÿ «·—’Ìœ"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand cmdPost 
         Height          =   510
         Left            =   45
         TabIndex        =   16
         Top             =   135
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   900
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
         Picture         =   "okaz_JUNIOR.frx":23BA
         Caption         =   "  —ÕÌ· «·„” ‰œ"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand cmdUnpost 
         Height          =   510
         Left            =   3555
         TabIndex        =   17
         Top             =   135
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   900
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
         Picture         =   "okaz_JUNIOR.frx":47BC
         Caption         =   " «·€«¡  —ÕÌ·"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1230
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   810
      Width           =   1635
      Begin VB.CommandButton cmdAddItems 
         Caption         =   "«÷«›… «’‰«› «·„Ê—œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame Frame6 
      Height          =   780
      Left            =   7065
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   1725
      Begin Threed.SSCommand CMD_PRINT 
         Height          =   600
         Left            =   45
         TabIndex        =   28
         Top             =   135
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   1058
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
         Picture         =   "okaz_JUNIOR.frx":6BDF
         Caption         =   "ÿ»«⁄… «·„” ‰œ"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.CheckBox xClosed 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   195
      Left            =   -585
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   -135
      Visible         =   0   'False
      Width           =   1860
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   27
      Top             =   9435
      Width           =   18030
      _ExtentX        =   31803
      _ExtentY        =   741
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   12347
            MinWidth        =   12347
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   810
      Width           =   1410
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "okaz_JUNIOR.frx":8FEE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   1320
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
         Picture         =   "okaz_JUNIOR.frx":B567
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   8820
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   6360
      Begin VB.CommandButton CmdDelInv 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   1620
         MaskColor       =   &H00FFFFFF&
         Picture         =   "okaz_JUNIOR.frx":D8CA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   3195
         MaskColor       =   &H00FFFFFF&
         Picture         =   "okaz_JUNIOR.frx":10164
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   4770
         Picture         =   "okaz_JUNIOR.frx":12710
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   135
         Width           =   1545
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "okaz_JUNIOR.frx":14EE3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1770
      Left            =   7695
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   720
      Width           =   7485
      Begin VB.TextBox xRate3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   405
         MaxLength       =   3
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   1305
         Width           =   690
      End
      Begin VB.TextBox xRate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   405
         MaxLength       =   3
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   900
         Width           =   690
      End
      Begin VB.TextBox xSup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4950
         MaxLength       =   3
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1365
      End
      Begin VB.TextBox xrate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   2
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   900
         Width           =   690
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4950
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Width           =   1950
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   4365
         TabIndex        =   3
         Top             =   900
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
         Caption         =   "Œ’„ „” Â·ﬂ : "
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
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1350
         Width           =   1185
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ «·„Ê—œ"
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
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   540
         Width           =   915
      End
      Begin VB.Label XRDISC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   540
         Width           =   555
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "‰”»… «·„” Â·ﬂ : "
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
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   945
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   945
         Width           =   195
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "%"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   945
         Width           =   195
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   540
         Width           =   2850
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Œ’„ :"
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
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   945
         Width           =   495
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         Height          =   270
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   945
         Width           =   645
      End
      Begin VB.Label Label4 
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
         Height          =   240
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   585
         Width           =   705
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
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label1 
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
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   1935
      Top             =   135
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   3420
      Top             =   90
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
      Height          =   5865
      Left            =   90
      TabIndex        =   5
      Top             =   2610
      Width           =   17835
      _cx             =   31459
      _cy             =   10345
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
   Begin VB.Frame Frame10 
      Caption         =   "»ÕÀ ⁄‰ »«—ﬂÊœ ›Ï «·„” ‰œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1350
      Width           =   2085
      Begin VB.TextBox xfilter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   225
         Width           =   1905
      End
   End
   Begin VB.Frame Frame7 
      Height          =   375
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   9045
      Width           =   3120
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   45
         TabIndex        =   26
         Top             =   135
         Visible         =   0   'False
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame4 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   8460
      Width           =   3165
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         TabIndex        =   36
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
         Picture         =   "okaz_JUNIOR.frx":1734F
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "okaz_JUNIOR.frx":1951F
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   810
         TabIndex        =   37
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
         Picture         =   "okaz_JUNIOR.frx":1B667
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "okaz_JUNIOR.frx":1D82F
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1575
         TabIndex        =   38
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
         Picture         =   "okaz_JUNIOR.frx":1F97E
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "okaz_JUNIOR.frx":21B5E
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2340
         TabIndex        =   39
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
         Picture         =   "okaz_JUNIOR.frx":23CB9
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "okaz_JUNIOR.frx":25E75
      End
   End
   Begin VB.Frame Frame9 
      Height          =   690
      Left            =   5850
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   8685
      Width           =   4290
      Begin VB.CheckBox XISPOST 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   " —ÕÌ· ··„Ê—œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   135
         Width           =   1185
      End
      Begin VB.Label Label3 
         Caption         =   "≈Ã„«·Ì «·Œ’„ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   420
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   180
         Width           =   1635
      End
   End
   Begin MSAdodcLib.Adodc DATA3 
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
   Begin Threed.SSCommand CMD_PrintBar 
      Height          =   555
      Left            =   180
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
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
      Picture         =   "okaz_JUNIOR.frx":27FC4
      Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
      Alignment       =   1
      PictureAlignment=   3
   End
End
Attribute VB_Name = "okaz_JUNIOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchDoc As New Search3, oSearchItem As New Search3, oSearchSup As New Search3
Dim CardTable As ADODB.Recordset
Dim cFilter As String
Dim cDefStore As String
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional nRow As Long = -1) As Boolean
Dim aInsert As Variant
CalcTotals
aInsert = AddFlag(Empty, "[date]", addDate(xDate.text))
aInsert = AddFlag(aInsert, "[sup]", addstring(xSup.text))
aInsert = AddFlag(aInsert, "[rate]", Val(xRate.text))
aInsert = AddFlag(aInsert, "[rate2]", Val(xRate2.text))
aInsert = AddFlag(aInsert, "[mosm]", addstring(xMosm.BoundText))
aInsert = AddFlag(aInsert, "[TOTAL]", Val(xTotal.Caption))
aInsert = AddFlag(aInsert, "[ISPOST]", IIf(XISPOST.Value = 0, 0, 1))
On Error GoTo myerror
con.BeginTrans
If xdoc_no.Tag = DefineMode Then
    xdoc_no.text = RetZero(Newflag("FILE0_70H", "DOC_NO", con))
    aInsert = AddFlag(aInsert, "doc_no", addstring(xdoc_no.text))
    con.Execute addInsert(aInsert, "FILE0_70H")
Else
    con.Execute addUpdate(aInsert, "FILE0_70H", "doc_no = " & addstring(xdoc_no.text))
End If
myreplaceGrd nRow
con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
ElseIf ActiveControl.Name = xbarcode.Name Then
    xbarcode.text = GetDesca("SELECT ITEM FROM FILE1_10 WHERE  MODEL = " & MyParn(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)), con)
    Unload oSearchItem
    
ElseIf ActiveControl.Name = CmdInform.Name Then
    xdoc_no.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
    myUndo
ElseIf ActiveControl.Name = xSup.Name Then
    xSup.text = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    Unload oSearchSup
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Unload Search
End Sub

Private Sub CMD_PRINT_Click()
    Dim cHeader1 As String, cHeader2 As String, cHeader3 As String, cHeader4 As String
    With grid1
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 14, "#0", vbYellow, , True, "  "
    .TextMatrix(.Rows - 1, 6) = Val(xqty.Caption)
    For nRow = 1 To .Rows - 1
        .RowHidden(nRow) = False
        If Val(.TextMatrix(nRow, 14)) = 0 Then .RowHidden(nRow) = True
    Next nRow
    cHeader1 = "„” ‰œ «Êﬂ«“ÌÊ‰ „Ê—œ —ﬁ„ : " & xdoc_no.text & Space(5) & "» «—ÌŒ : " & xDate.text
    cHeader2 = "«·„Ê—œ : " & xDesca.Caption & Space(5) & "„Ê”„ : " & xMosm.text
    cHeader3 = "‰”»… Œ’„ : " & xRate.text & "‰”»… „” Â·ﬂ : " & xRate2.text
    cHeader4 = "≈Ã„«·Ì «·Œ’„  : " & xTotal.Caption
    PrintGrdNew.doprint Me.grid1, 1.1, -1, cHeader1, cHeader2, cHeader3, cHeader4, False, True, 11
    PrintGrdNew.Show 1
    
    grid1.Rows = 1
    openCardTable
    myUndo
    
    End With
End Sub

Private Sub CMD_PrintBar_Click()
con.BeginTrans
On Error GoTo myerror
con.Execute "DELETE FROM ADDPRINT "
cString = "INSERT INTO ADDPRINT(DOC_NO,ITEM,QUANT,ISPRINT)"

cString = cString & _
          " SELECT DOC_NO,ITEM,QUANT,1 FROM FILE0_70"
cString = cString & turn(cString) & " DOC_NO = " & MyParn(xdoc_no.text)
con.Execute cString
 
con.CommitTrans
Inform " „ «· ÕÊÌ· »‰Ã«Õ"
BarCodeFRM.Show
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear

End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE0_70 where Doc_No = " & MyParn(xdoc_no.text)
    con.Execute "Delete  From FILE0_70H where Doc_No = " & MyParn(xdoc_no.text)
    con.CommitTrans
    openCardTable
    CmdNewInv_Click
    Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
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
Private Sub cmdFix_Click()
If Not MYVALID Then Exit Sub
For i = 1 To grid1.Rows - 1
'    grid1.TextMatrix(I, 4) = Val(GetDesca("select sum([in] - [out]) from file1_11 inner join file1_10 on file1_10.item = file1_11.item where ( TYPE = '2' OR TYPE = '7' ) AND file1_10.ITEM = " & Val(grid1.TextMatrix(I, 5)) & " And Date < " & DateSq(xdate.Text)) & "")
'    grid1.TextMatrix(I, 5) = Val(GetDesca("select sum( [out]) from file1_11 inner join file1_10 on file1_10.item = file1_11.item where TYPE = '6' AND file1_10.model = " & MyParn(grid1.TextMatrix(I, 0)) & " And Date < " & DateSq(xdate.Text)) & "")
    grid1.TextMatrix(i, 6) = Val(GetDesca("select sum([in] - [out]) from file1_11 where ITEM = " & Val(grid1.TextMatrix(i, 5)) & " And Date < " & DateSq(xDate.text)) & "")
Next
CalcTotals
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ ÷»ÿ «·„” ‰œ »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO, CONVERT(VARCHAR(10),[DATE],111),FILE4_10.DESCA " & _
                  " FROM FILE0_70H INNER JOIN FILE4_10 ON FILE0_70H.SUP = FILE4_10.CODE"
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If

Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·„Ê—œ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]## OR %%FILE4_10.DESCA%%)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·„Ê—œ"
GrdArray(2, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Load oSearchDoc
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
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
Private Sub cmdPost_Click()
Dim nRDisc  As Double
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
On Error GoTo myerror
con.BeginTrans
cString = " UPDATE FILE0_70H SET FILE0_70H.closed = 1 WHERE FILE0_70H.DOC_NO = " & MyParn(xdoc_no.text)
con.Execute cString
nRDisc = ((100 - Val(XRDISC.Caption)) / 100)

Dim aInsert As Variant
For i = 1 To grid1.Rows - 1
    If Val(grid1.TextMatrix(i, 9)) > 0 Then
        aInsert = AddFlag(Empty, "COST2", Val(grid1.TextMatrix(i, 9)))
        aInsert = AddFlag(aInsert, "COSTNEW", Val(grid1.TextMatrix(i, 9)))
        aInsert = AddFlag(aInsert, "COSTITEM", Val(grid1.TextMatrix(i, 9)) * nRDisc)
    Else
        aInsert = AddFlag(Empty, "COST2", Val(grid1.TextMatrix(i, 7)))
        aInsert = AddFlag(aInsert, "COSTNEW", Val(grid1.TextMatrix(i, 7)))
    End If
    If Val(grid1.TextMatrix(i, 10)) > 0 Then
        aInsert = AddFlag(aInsert, "price2", Val(grid1.TextMatrix(i, 10)))
    Else
        aInsert = AddFlag(aInsert, "price2", Val(grid1.TextMatrix(i, 8)))
    End If
    con.Execute addUpdate(aInsert, "file1_10", "model = " & MyParn(grid1.TextMatrix(i, 0)))
Next
con.CommitTrans
openCardTable
myUndo
Inform " „  —ÕÌ· «·„” ‰œ »‰Ã«Õ"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
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
    On Error Resume Next
    xdoc_no.SetFocus
    Err.Clear
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
    openCardTable
    myUndo
End Sub
Private Sub cmdunPost_Click()
On Error GoTo myerror
con.BeginTrans
cString = "UPDATE FILE0_70H SET FILE0_70H.closed = 0 WHERE FILE0_70H.DOC_NO = " & MyParn(xdoc_no.text)
con.Execute cString
con.CommitTrans
openCardTable
myUndo
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdAddItems_Click()
    AddItems ""
    CalcTotals
    If Not MYVALID Then Exit Sub
    If Not myreplace Then Exit Sub
    Inform " „ «÷«›…  ﬂ·›… «·„” ‰œ »‰Ã«Õ"
    openCardTable
    myUndo
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
bedit = True

openCon con

con.Execute " UPDATE FILE0_70H SET DOC_NO = '0' + FILE0_70H.DOC_NO WHERE LEN(DOC_NO) = 5"
con.Execute " UPDATE FILE0_70 SET DOC_NO = '0' + FILE0_70.DOC_NO WHERE LEN(DOC_NO) = 5"

con.Execute " UPDATE FILE0_70H SET ISPOST = 1 WHERE FILE0_70H.TOTAL > 0 "

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select mosm ,descA From mosm ORDER BY date "
Set xMosm.RowSource = DATA2
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"

data3.ConnectionString = strCon
data3.RecordSource = "Select CODE , DESCA From FILE1_10SC "
Set XSEC.RowSource = data3
XSEC.ListField = "Desca"
XSEC.BoundColumn = "CODE"

Set grid1.DataSource = data1
data1.ConnectionString = strCon

openCardTable
CmdNewInv_Click
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload oSearchDoc
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim bNew As Boolean
'If grid1.Col = 0 Then
'    GrdDesc grid1.Row
'End If
CalcTotals
If Not validRow(Row) Then Exit Sub
'If Row = grid1.Rows - 1 Then
'    myaddItem
'    bNew = True
'End If

With grid1
If myreplace(Row) Then
    If xdoc_no.Tag = DefineMode Then xdoc_no.Tag = LoadMode
End If
myloadgrd
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
'    If Not validRow(OldRow) Then grid1.RemoveItem OldRow
'End If
End Sub
Private Sub grid1_EnterCell()
With grid1
If (grid1.Col = 9 Or grid1.Col = 10 Or grid1.Col = 6) And xClosed.Value = 0 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub
Sub Grid1_GotFocus()
With grid1
If grid1.Row <= 1 And grid1.Rows > 1 Then
    .Select 1, 0, 1, 0
    .ShowCell 1, 0
End If
End With
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    'If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 Then grid1.AddItem "", grid1.Row
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
'If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then grid1.RemoveItem grid1.Row
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
'    If Not IsNumeric(grid1.EditText) Then
'        Inform "»«—ﬂÊœ «·’‰› €Ì— ”·Ì„"
'        Cancel = True
'        Exit Sub
'    End If
'    If Not validItem(grid1.EditText, con) Then
'        Inform "»«—ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
'        Cancel = True
'    End If
End If
'If Col = 8 Then
'    If Val(grid1.EditText) >= Val(grid1.TextMatrix(Row, 9)) Then
'        MsgBox "«·—’Ìœ ·« Ì”„Õ"
'    End If
'End If
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub xbarcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And xbarcode.text <> "" Then
    
    AddItems xbarcode.text
    CalcTotals
    If Not MYVALID Then Exit Sub
    If Not myreplace Then Exit Sub
    Inform " „ «÷«›…  ﬂ·›… «·„” ‰œ »‰Ã«Õ"
    openCardTable
    myUndo
    xbarcode.text = ""
    On Error Resume Next
    xbarcode.SetFocus
    Err.Clear
End If
End Sub

Private Sub xbarcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    ModelLookupAll Me, oSearchItem, xSup.text, xMosm.BoundText
'    XBARCODE_KeyPress (13)

    AddItems xbarcode.text
    CalcTotals
    If Not MYVALID Then Exit Sub
    If Not myreplace Then Exit Sub
    Inform " „ «÷«›…  ﬂ·›… «·„” ‰œ »‰Ã«Õ"
    openCardTable
    myUndo

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

Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CmdInform_Click
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If xdoc_no.text = "" Then
    If Not bIsgMsg Then Inform "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    If Not bIsMsg Then Inform "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If Not xMosm.MatchedWithList Then
    If Not bIsMsg Then Inform "·„ Ì „ «œŒ«· «·„Ê”„"
    Exit Function
End If

If xSup.text = "" Then
    If Not bIsMsg Then Inform "·„ Ì „ «œŒ«· ﬂÊœ «·„Ê—œ"
    Exit Function
End If

If IsEmpty(GetField("select code from file4_10 where code = " & MyParn(xSup.text), con)) Then
    If Not bIsMsg Then Inform "ﬂÊœ «·„Ê—œ €Ì— „ÊÃÊœ"
    Exit Function
End If

With grid1
'For i = 1 To .Rows - 2
'    If Not ValidInt(.TextMatrix(i, 0)) Then
'        If Not validItem(.TextMatrix(i, 0), con) Then
'            MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
'            Exit Function
'        End If
'    End If
'Next
End With
MYVALID = True
End Function
Private Sub myload()
xdoc_no.text = CardTable!doc_no
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
xSup.text = CardTable!sup & ""
xDesca.Caption = CardTable!DESCA & ""
XRDISC.Caption = CardTable!RDISC & ""
xClosed.Value = IIf(CardTable!CLOSED, 1, 0)
XISPOST.Value = IIf(CardTable!ISPOST, 1, 0)
'xrate.Text = Myvalue(CardTable!Rate, "fixed")
'xRate2.Text = Myvalue(CardTable!Rate2, "fixed")

xMosm.BoundText = CardTable!MOSM & ""
xdoc_no.Enabled = False
myloadgrd
Handlecontrols LoadMode
CalcTotals
If XISPOST.Value = 0 Then xTotal.Caption = ""
End Sub
Private Sub myDefine()
Frame5.Visible = True
xdoc_no.Enabled = True
xdoc_no.text = RetZero(Newflag("FILE0_70H", "DOC_NO", con))
xDate.text = Format(Date, "dd-mm-yyyy")
XRDISC.Caption = ""
xMosm.BoundText = cPMosm
xSup.text = ""
xRate3.text = ""
xqty.Caption = ""
XISPOST.Value = 0
xRate.text = ""
xRate2.text = ""
xDesca.Caption = ""
'xMosm.BoundText = ""
xClosed.Value = 0
xTotal.Caption = ""
'StatusBar1.Panels(1).Text = ""
'xTotal.Caption = ""
grid1.Rows = 1
Handlecontrols DefineMode
Fixgrd
End Sub
Private Sub Handlecontrols(nMode)
Dim bClosed As Boolean
cmdNewInv.Enabled = nMode = LoadMode And bedit
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
bClosed = xClosed.Value = 1
cmdPost.Enabled = (Not bClosed) And bedit And nMode = LoadMode
cmdUnpost.Enabled = bClosed And bedit And nMode = LoadMode
cmdFix.Enabled = (Not bClosed) And bedit And nMode = LoadMode
cmdSave.Enabled = (nMode = DefineMode Or Not bClosed)
CmdDelInv.Enabled = (Not bClosed) And nMode = LoadMod
xdoc_no.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xdoc_no
If xdoc_no.text = "" Then Exit Sub
xdoc_no.text = RetZero(xdoc_no.text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
'    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
'        On Error GoTo myerror
'        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
'            con.BeginTrans
'            con.Execute "delete from FILE0_70 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
'            con.CommitTrans
'        End If
'        grid1.RemoveItem grid1.Row
'    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
'ElseIf KeyCode = 45 And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And validRow(grid1.Row) Then
'    grid1.AddItem "", grid1.Row
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
End Sub
Private Function CalcTotals()
Dim nCount As Double, nQTY As Double
Dim nTotal As Double, nRDisc As Double
nRDisc = ((100 - Val(XRDISC.Caption)) / 100)
With grid1
For i = 1 To grid1.Rows - 1
    If Val(.TextMatrix(i, 14)) <> 0 Then
        nQTY = nQTY + Val(.TextMatrix(i, 6))
    End If
    .TextMatrix(i, 11) = Val(.TextMatrix(i, 6)) * Val(.TextMatrix(i, 7))
    .TextMatrix(i, 12) = Val(.TextMatrix(i, 6)) * Val(.TextMatrix(i, 9))
    .TextMatrix(i, 13) = Val(.TextMatrix(i, 7)) - Val(.TextMatrix(i, 9))
    .TextMatrix(i, 14) = (Val(.TextMatrix(i, 11)) - Val(.TextMatrix(i, 12))) * nRDisc
    nTotal = nTotal + Val(.TextMatrix(i, 14))
Next
xTotal.Caption = Myvalue(nTotal, "Fixed")
xqty.Caption = nQTY
End With
End Function
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function

Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = nValue Then
            FoundOtheritem = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub Fixgrd()
Dim nTCost As Double


'                   0                   1           2                   3                   4               5               6               7               8               9               10                  11              12                      13                              14                                  15
'cString = "SELECT FILE0_70.item ,  fact.desca, file1_10.modelfact0 ,file1_10.desca,FILE1_10.SCAL , FILE1_10.COLOR  ,FILE0_70.Quant , FILE0_70.COST , FILE0_70.COST2 , FILE0_70.DISCOUNT  , FILE0_70.PRICE , FILE0_70.PRICE2 , FILE0_70.COST_TOTAL,FILE0_70.COST_TOTAL2, ( FILE0_70.COST_TOTAL2 - FILE0_70.COST_TOTAL ) ,FILE0_70.ID   " & _

With grid1
.ExplorerBar = flexExSortShow
.FormatString = "»«—ﬂÊœ|" & "„’‰⁄|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "·Ê‰|" & "«·—’Ìœ|" & "”⁄— „’‰⁄|" & "”⁄— „” Â·ﬂ|" & "‰”»… Œ’„|" & "”⁄— „’‰⁄ ÃœÌœ|" & "”⁄— „” Â·ﬂ ÃœÌœ|" & "≈Ã„«·Ì  ›—ﬁ «· ﬂ·›… |"
.WordWrap = True

.RowHeight(0) = 1000

.ColWidth(0) = 1000
.ColHidden(0) = True
.ColWidth(1) = 2000
.ColWidth(2) = 1100
.ColWidth(3) = 2500
.ColWidth(4) = 1500
.ColWidth(5) = 0
.ColWidth(6) = 800
.ColWidth(7) = 800
.ColWidth(8) = 800
.ColWidth(9) = 800
.ColWidth(10) = 800
.ColWidth(11) = 1000
.ColWidth(12) = 1000
.ColWidth(13) = 800
.ColHidden(.Cols - 1) = True
'.ColComboList(2) = StrList("SELECT CODE , DESCA FROM FACT ")
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
For nRow = 1 To .Rows - 1
    nTCost = nTCost + Val(.TextMatrix(nRow, 11))
Next nRow
End With
XTCOST.text = Format(nTCost, "#0.00")
End Sub
Private Sub myreplaceGrd(nRow)
Dim aInsert As Variant
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 1, nRow)
        aInsert = AddFlag(Empty, "[doc_no]", addstring(xdoc_no.text))
        aInsert = AddFlag(aInsert, "[model]", addstring(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "[ITEM]", addstring(grid1.TextMatrix(i, 5)))
'        aInsert = AddFlag(aInsert, "[t_in]", Val(grid1.TextMatrix(I, 4)))
'        aInsert = AddFlag(aInsert, "[t_out]", Val(grid1.TextMatrix(I, 5)))
        aInsert = AddFlag(aInsert, "[quant]", Val(grid1.TextMatrix(i, 6)))
        aInsert = AddFlag(aInsert, "[cost]", Val(grid1.TextMatrix(i, 7)))
        aInsert = AddFlag(aInsert, "[price]", Val(grid1.TextMatrix(i, 8)))
        If Val(grid1.TextMatrix(i, 9)) = 0 Then grid1.TextMatrix(i, 9) = Val(grid1.TextMatrix(i, 7))
        aInsert = AddFlag(aInsert, "[cost2]", Val(grid1.TextMatrix(i, 9)))
        aInsert = AddFlag(aInsert, "[price2]", Val(grid1.TextMatrix(i, 10)))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FILE0_70")
        Else
            con.Execute addUpdate(aInsert, "FILE0_70", "ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub myloadgrd()
Dim cString As String
xqty.Caption = ""

'                   0                   1           2                   3                   4               5               6               7               8               9               10                  11              12                      13                              14                                  15
cString = "SELECT FILE0_70.item ,  fact.desca, file1_10.modelfact0 ,file1_10.desca,FILE1_10.SCAL , FILE1_10.COLOR  ,FILE0_70.Quant , FILE0_70.COST , FILE0_70.PRICE , FILE0_70.DISCOUNT  , FILE0_70.COST2 , FILE0_70.PRICE2 , FILE0_70.COST_TOTAL,FILE0_70.COST_TOTAL2, ( FILE0_70.COST_TOTAL2 - FILE0_70.COST_TOTAL ) ,FILE0_70.ID   " & _
          " from FILE0_70 inner join file1_10 on file1_10.ITEM = FILE0_70.ITEM LEFT join fact on fact.code = file1_10.fact  "
cString = cString & turn(cString) & " DOC_NO = " & MyParn(xdoc_no.text)
cString = cString & " ORDER BY ID"
Frame5.Visible = True
data1.RecordSource = cString
data1.Refresh
Fixgrd
CalcTotals
End Sub
Private Sub openCardTable()
Set CardTable = New ADODB.Recordset
Dim cString As String
cString = "SELECT FILE0_70H.*,FILE4_10.DESCA AS DESCA , FILE4_10.DISC AS RDISC FROM FILE0_70H INNER JOIN FILE4_10 ON FILE0_70H.SUP = FILE4_10.CODE"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY DOC_NO"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    myDefine
Else
    If Trim(xdoc_no.text) <> "" Then
        CardTable.Find "DOC_NO = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub xfilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter.text, 0
    End If
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not ValidInt(.TextMatrix(Row, 0)) Then Exit Function
If Not validItem(.TextMatrix(Row, 0), con) Then Exit Function
validRow = True
End With
End Function
Private Sub myaddItem()
grid1.AddItem ""
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    grid1.Col = Col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Row = Row + 1
    grid1.Col = 1
    grid1.ShowCell Row + 1, 1
End If
End Sub

Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xfilter_GotFocus()
myGotFocus xfilter
End Sub
Private Sub xfilter_LostFocus()
myLostFocus xfilter
End Sub
Private Sub xRate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    updateCost
    CalcTotals
    If Not MYVALID Then Exit Sub
    If Not myreplace Then Exit Sub
    Inform " „ «÷«›…  ﬂ·›… «·„” ‰œ »‰Ã«Õ"
    openCardTable
    myUndo
End If
End Sub

Private Sub xRate2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    updateCost
    CalcTotals
    If Not MYVALID Then Exit Sub
    If Not myreplace Then Exit Sub
    Inform " „ «÷«›…  ﬂ·›… «·„” ‰œ »‰Ã«Õ"
    openCardTable
    myUndo
End If
End Sub
Private Sub xRate3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    updateCost
    CalcTotals
    If Not MYVALID Then Exit Sub
    If Not myreplace Then Exit Sub
    Inform " „ «÷«›…  ﬂ·›… «·„” ‰œ »‰Ã«Õ"
    openCardTable
    myUndo
End If
End Sub

Private Sub xSup_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    SuppLookupAll Me, oSearchSup
End If
End Sub

Private Sub xSup_LostFocus()
xDesca.Caption = ""

If Trim(xSup.text) = "" Then Exit Sub
xSup.text = RetZero(xSup, 3)
xDesca.Caption = GetField("Select desca from file4_10 where code = " & MyParn(xSup.text), con) & ""
XRDISC.Caption = GetDesca("Select DISC from file4_10 where code = " & MyParn(xSup.text), con) & ""
End Sub
Private Sub AddItems(pBARCODE As String)
Dim loctable As New ADODB.Recordset, nRecordCount As Integer
Dim cModel As String
cModel = "---"
If pBARCODE <> "" Then
    cModel = GetDesca(" SELECT MODEL FROM FILE1_10 WHERE ITEM =   " & Val(pBARCODE), con)
End If
cField1 = myiif("(type = '7' or type = '2' ) ", "([in] - [out]) ") & " AS t_in "

cField2 = myiif("type = '6' ", "[out] ") & " AS t_sal"

cString = "SELECT FILE1_10.ITEM ,FILE1_10.COST , FILE1_10.PRICE ,file1_10.mosm ,fact.desca as factDesca,file1_10.code, file1_10.modelfact0 ,file1_10.desca , FILE1_10.SCAL,FILE1_10.ITEM,FILE1_10.COST2,FILE1_10.PRICE2,FILE1_10.RATE,Sum(FILE1_11.[IN] - FILE1_11.[out]) as Balance , " & cField1 & " , " & cField2 & _
          " FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT"
cString = cString & turn(cString) & " FILE1_11.DATE < " & DateSq(xDate.text)
cString = cString & turn(cString) & " FILE1_10.code = " & MyParn(xSup.text)
cString = cString & turn(cString) & " FILE1_10.MOSM = " & MyParn(xMosm.BoundText)

If XSEC.BoundText <> "" Then cString = cString & turn(cString) & " FILE1_10.[SECTION] = " & Val(XSEC.BoundText)
If cModel <> "---" Then cString = cString & turn(cString) & " FILE1_10.[MODEL] = " & addstring(cModel)

cString = cString & " GROUP BY FILE1_10.ITEM,FILE1_10.SCAL,FILE1_10.COST , FILE1_10.PRICE ,FILE1_10.model ,file1_10.mosm ,fact.desca,file1_10.code , file1_10.modelfact0 ,file1_10.desca , FILE1_10.COST2,FILE1_10.PRICE2,FILE1_10.RATE HAVING Sum(FILE1_11.[IN] - FILE1_11.[out]) <> 0 "
cString = cString & " ORDER BY FILE1_10.MODEL,FILE1_10.SCAL"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (loctable.EOF And loctable.BOF) Then
    nRecordCount = loctable.RecordCount
End If
With grid1
    prog1.Visible = True
    prog1.Value = 0
    Do Until loctable.EOF
        If loctable!balance <> 0 Then
            If grid1.FindRow(loctable!Item, , 5) = -1 Then
                grid1.AddItem ""
                i = i + 1
                prog1.Value = Round(i / (nRecordCount), 2) * 100
                grid1.TextMatrix(.Rows - 1, 0) = loctable!Item
                grid1.TextMatrix(.Rows - 1, 1) = loctable!FACTDESCA
                grid1.TextMatrix(.Rows - 1, 2) = loctable!MODELFACT0 & ""
                grid1.TextMatrix(.Rows - 1, 3) = loctable!DESCA & ""
                grid1.TextMatrix(.Rows - 1, 4) = loctable!scal & ""
                grid1.TextMatrix(.Rows - 1, 5) = loctable!Color & ""
                grid1.TextMatrix(.Rows - 1, 6) = Val(loctable!balance & "")
                grid1.TextMatrix(.Rows - 1, 7) = Val(loctable!cost & "")
                grid1.TextMatrix(.Rows - 1, 8) = Val(loctable!price & "")
            End If
        End If
        loctable.MoveNext
    Loop
    prog1.Visible = False
End With
CalcTotals
End Sub
Private Sub updateCost()
With grid1
    prog1.Visible = True
    prog1.Value = 0
    For i = 1 To .Rows - 1
        prog1.Value = Round(i / (.Rows - 1), 2) * 100
        If IsNumeric(xRate.text) <> 0 Then
             grid1.TextMatrix(i, 9) = Round(Val(grid1.TextMatrix(i, 7)) * (1 - (Val(xRate.text) / 100)), 2)
        End If
        If IsNumeric(xRate2.text) Then
           grid1.TextMatrix(i, 10) = Round(Val(grid1.TextMatrix(i, 9)) * (1 + (Val(xRate2.text) / 100)), 0)
        End If
        If IsNumeric(xRate3.text) Then
            nRate = ((100 - Val(xRate3.text)) / 100)
            grid1.TextMatrix(i, 10) = Round(Val(grid1.TextMatrix(i, 8)) * nRate, 0)
        End If
    Next
    prog1.Visible = False
End With
CalcTotals
End Sub
