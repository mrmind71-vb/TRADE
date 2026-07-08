VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form purchasefrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14925
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
   ScaleWidth      =   14925
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame11 
      Height          =   2535
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   75
      Top             =   6930
      Width           =   4785
      Begin ImgXCtrl6.ImgXCtrl imgx1 
         Height          =   2175
         Left            =   90
         TabIndex        =   76
         Top             =   225
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   3836
         BackColor       =   16777215
         BorderStyle     =   0
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrvb71"
         LicenseRegCode  =   "íß“ªß•≤º∂´≠“±®ππ∂´µßZQEH-AOZOOOZT-EFLF6gI"
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Frame10"
      Height          =   915
      Left            =   -2295
      RightToLeft     =   -1  'True
      TabIndex        =   70
      Top             =   -585
      Visible         =   0   'False
      Width           =   3660
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   495
         TabIndex        =   71
         Top             =   135
         Visible         =   0   'False
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   45
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin VB.CommandButton CMD_COPY 
      Caption         =   " ﬂ—«— «·›« Ê—…"
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
      Left            =   -1080
      MaskColor       =   &H00FFFFFF&
      Picture         =   "purchase2.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   -270
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Height          =   1860
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   90
      Visible         =   0   'False
      Width           =   2085
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         TabIndex        =   62
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
         Picture         =   "purchase2.frx":2579
         Caption         =   "   ≈€·«ﬁ «·„” ‰œ  "
         Alignment       =   4
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   1035
         TabIndex        =   63
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
         Picture         =   "purchase2.frx":4B45
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   90
         TabIndex        =   64
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
         Picture         =   "purchase2.frx":761A
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame Frame9 
      Height          =   780
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   0
      Width           =   5415
      Begin Threed.SSCommand CMD_PRINT 
         Height          =   510
         Left            =   3810
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
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
         Picture         =   "purchase2.frx":A14C
         Caption         =   "ÿ»«⁄… «·›« Ê—…"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand CMD_PrintBar 
         Height          =   510
         Left            =   90
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   180
         Width           =   1815
         _ExtentX        =   3201
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
         Picture         =   "purchase2.frx":C55B
         Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand cmdBarCode 
         Height          =   510
         Left            =   1905
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   180
         Width           =   1860
         _ExtentX        =   3281
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
         Picture         =   "purchase2.frx":EEDD
         Caption         =   " ÕÊÌ· ··»«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   765
      Width           =   1365
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase2.frx":11876
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   59
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
         Picture         =   "purchase2.frx":13DEF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   9765
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton CmdDelInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1275
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase2.frx":16152
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   56
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
         Picture         =   "purchase2.frx":189EC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   3735
         Picture         =   "purchase2.frx":1AF98
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase2.frx":1D76B
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.TextBox XOLDDOC 
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
      Left            =   -765
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   315
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   90
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
         Left            =   5265
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   225
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   3555
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   720
      Width           =   11265
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
         TabIndex        =   4
         Top             =   540
         Width           =   2445
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
         TabIndex        =   3
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
         MaxLength       =   6
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
         Left            =   4995
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   -225
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   1485
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   900
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   1485
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   2445
         _ExtentX        =   4313
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
         TabIndex        =   6
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
         Picture         =   "purchase2.frx":1FBD7
         Caption         =   " ”ÃÌ· „ÊœÌ·«  "
         Alignment       =   8
         PictureAlignment=   6
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
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   855
         Width           =   3615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰Ê⁄ «·„Ê—œ :"
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
         TabIndex        =   73
         Top             =   900
         Width           =   975
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
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   855
         Width           =   1320
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
         TabIndex        =   49
         Top             =   555
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Œ“‰ :"
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
         Index           =   0
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   900
         Width           =   570
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         Left            =   6345
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   495
         Width           =   2490
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
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   44
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
         Left            =   6390
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   180
         Width           =   1035
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   225
         Width           =   555
      End
   End
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   12420
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   8325
      Width           =   2400
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   45
         Picture         =   "purchase2.frx":223DF
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   180
         Width           =   510
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   1845
         Picture         =   "purchase2.frx":24ABE
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   180
         Width           =   510
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   600
         Picture         =   "purchase2.frx":27198
         Style           =   1  'Graphical
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   180
         Width           =   555
      End
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   1200
         Picture         =   "purchase2.frx":2976B
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   180
         Width           =   600
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   420
      Left            =   4905
      TabIndex        =   19
      Top             =   9045
      Visible         =   0   'False
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   45
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
   Begin VB.Frame Frame7 
      Height          =   1455
      Left            =   4905
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   6885
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
         TabIndex        =   27
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   585
         Width           =   975
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   225
         Width           =   1455
      End
      Begin VB.Label xtotalQuant 
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
         TabIndex        =   18
         Top             =   945
         Width           =   1245
      End
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
         Left            =   8415
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   945
         Width           =   1365
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   990
         Width           =   1365
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "%"
         Height          =   165
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   810
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   270
         Width           =   2130
      End
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   675
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
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   -180
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   45
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
         TabIndex        =   69
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
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   180
         Width           =   2850
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   -1710
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   -135
      Visible         =   0   'False
      Width           =   2130
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   585
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
      Left            =   -225
      Top             =   -45
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
      Height          =   4800
      Left            =   45
      TabIndex        =   7
      Top             =   2070
      Width           =   14775
      _cx             =   26061
      _cy             =   8467
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
      ScrollBars      =   2
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
   Begin VB.CheckBox xIsClosed 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   285
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   65
      Top             =   2160
      Visible         =   0   'False
      Width           =   1005
   End
End
Attribute VB_Name = "Purchasefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cFileHeader As String, sDoc_No As String
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim searchitem As New Search3
Dim Search1 As New Search3, Search2 As New Search3, bMarket As Boolean
Dim bEdit As Boolean
Dim cFile As String, cFileClient, cFieldClient, cCodeDesca As String
Dim formMode, dDateLast As String
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional bAddModel As Boolean = False) As Boolean
Dim aInsert(7, 1)

aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(xDate.Text)

aInsert(2, 0) = "Code"
aInsert(2, 1) = addstring(xCode.Text)

aInsert(3, 0) = "R_Discount"
aInsert(3, 1) = Val(xr_Discount.Text)

aInsert(4, 0) = "Tax"
aInsert(4, 1) = Val(xTax.Text)

aInsert(5, 0) = "Store"
aInsert(5, 1) = addstring(xStore.BoundText)

aInsert(6, 0) = "MOSM"
aInsert(6, 1) = addstring(xMosm.BoundText)

aInsert(7, 0) = "INV_NO"
aInsert(7, 1) = addstring(xInv_No.Text)

'aInsert(8, 0) = "usename"
'aInsert(8, 1) = addstring(xUserName.caption)

con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = RetZero(Newflag(cFileHeader, "doc_no"))
   aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
   con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & MyParn(xDoc_No.Text))
End If
con.CommitTrans
If bAddModel Then myreplaceGrd
myreplace = True
Exit Function
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
On Error GoTo myError
If ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "DOC_NO = " & MyParn(Search1.grid1.TextMatrix(Search1.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    Search1.Hide
    myload
ElseIf ActiveControl.Name = CMD_COPY.Name Then
    XOLDDOC.Text = Search1.grid1.TextMatrix(Search1.grid1.Row, 0)
    Search1.Hide
ElseIf TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
Exit Sub
myError:
End Sub
Private Sub CMD_COPY_Click()
    If xCode.Text <> "" And xDoc_No.Text = "" Then
        CardLookupSupp
        If XOLDDOC.Text <> "" Then
            If xDoc_No.Text = "" Then
                If MsgBox(XOLDDOC.Text & "”Ê› Ì „ ⁄„· ‰”Œ… „‰ ›« Ê—… «·„‘ —Ì«  —ﬁ„ ", vbYesNo + vbDefaultButton2) = vbYes Then
                    myreplaceH
                    
                    cStr1 = "INSERT INTO FILE7_20 (ITEM, QUANT, PRICE, MODEL, DOC_NO )  SELECT ITEM, QUANT, PRICE, MODEL , " & addstring(xDoc_No.Text) & "  FROM   FILE7_20 AS FILE7_20_1 WHERE FILE7_20_1.DOC_NO = " & MyParn(XOLDDOC.Text)
                    con.Execute cStr1
                    myload
                End If
            End If
        End If
    End If
End Sub
Private Sub CMD_FIX_Click()
FIXINV.Show 1
End Sub
Private Sub CMD_PrintBar_Click()
    barcodefrm.Show 1
End Sub
Private Sub CMD_TOT_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT FACT.DESCA DOC_NO,[DATE] AS DATE1,DATE, " & cFileClient & ".Desca , inv_no , mosm " & _
                  " FROM  (" & cFileHeader & " left JOIN " & cFileClient & " ON " & cFileHeader & ".CODE " & " = " & cFileClient & ".CODE )"

Generalarray(2) = "Order by Date"
Generalarray(3) = 6000
Generalarray(5) = True


listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
listarray(0, 1) = "(inv_no Like '%cFilter%' Doc_No Like '%cFilter%' or  " & cFileClient & ".DESCA LIKE '%cFilter%' OR " & _
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
Load Search1
Search1.Caption = "«” ⁄·«„"
Search1.Show 1

End Sub
Private Sub cmdAddItems_Click()
If Not MYVALID Then Exit Sub
Set addModelFrm.myForm = Me
addModelFrm.nColItem = 0
addModelFrm.nColQuant = 7
addModelFrm.sStore = xStore.BoundText
addModelFrm.bshowBal = True
addModelFrm.sMosm = xMosm.BoundText
addModelFrm.sSection = xSection.BoundText
addModelFrm.sMosm = xMosm.Text
If xSupp.Value = 1 Then addModelFrm.sSupp = xCode.Text Else addModelFrm.sFact = xsubcode.Caption
If IsNumeric(grid1.TextMatrix(grid1.Row, 0)) Then
    Dim loctable As ADODB.Recordset
    Set loctable = itemFind(grid1.TextMatrix(grid1.Row, 0), con)
    If Not (loctable.EOF And loctable.BOF) Then
        addModelFrm.sModelFact = loctable!modelfact0 & ""
        addModelFrm.sFact = loctable!Fact & ""
        addModelFrm.sSupp = loctable!SUPP & ""
        addModelFrm.sMosm = loctable!Mosm & ""
        addModelFrm.sModel = loctable!Model & ""
        addModelFrm.sSection = loctable!Section & ""
    End If
    loctable.Close
    Set loctable = Nothing
End If
addModelFrm.Show 1
End Sub
Private Sub cmdAddItems_GotFocus()
cmdAddItems.BackColor = vbYellow
End Sub
Private Sub cmdAddItems_LostFocus()
cmdAddItems.BackColor = &H8000000F
End Sub

Private Sub cmdBarCode_Click()
con.BeginTrans
On Error GoTo myError
con.Execute "DELETE FROM ADDPRINT WHERE DOC_NO = " & MyParn(xDoc_No.Text)

cString = "INSERT INTO ADDPRINT(DOC_NO,ITEM,QUANT,ISPRINT)"
cString = cString & _
          " SELECT DOC_NO,ITEM,QUANT,1 FROM FILE7_20"
cString = cString & turn(cString) & " DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute cString
con.CommitTrans
Inform " „ «· ÕÊÌ· »‰Ã«Õ"
CMD_PrintBar_Click
Exit Sub
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myError
    con.BeginTrans
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    openCardTable
    CmdNewInv_Click
End If
Exit Sub
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub CmdFirst_Click()
    CardTable.MoveFirst
    myload
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
    mydefine
    xInv_No.SetFocus
End Sub
Private Sub cmdSave_Click()
    mysave
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub CMD_PRINT_Click()
doprint
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
    SetKbLayout Lang_EN
    con.CursorLocation = adUseClient
    con.Open strCon
    'Frame8.Visible = lSupperVisor
    bEdit = True
    Select Case myPublic
    Case 0
        cFile = "File7_20"
        cFileHeader = "File7_20H"
        cFileClient = "File4_10"
        cMoveName = "„‘ —Ì« "
        Me.Caption = "›« Ê—… „‘ —Ì« "
    Case 1
        CMD_FIX.Enabled = False
        cFile = "FILE7_10"
        cFileHeader = "FILE7_10H"
        cFileClient = "File4_10"
        cMoveName = "„—œÊœ „‘ —Ì« "
        Me.Caption = "›« Ê—… „—œÊœ „‘ —Ì« "
        Me.BackColor = vbWhite
        Me.Frame1.BackColor = vbWhite
        Me.Frame2.BackColor = vbWhite
        Me.Frame3.BackColor = vbWhite
        Me.Frame4.BackColor = vbWhite
        Me.Frame5.BackColor = vbWhite
        Me.Frame6.BackColor = vbWhite
        Me.Frame7.BackColor = vbWhite
        Me.Frame8.BackColor = vbWhite
        Me.Frame9.BackColor = vbWhite
    End Select
        
    data1.ConnectionString = strCon
    data1.RecordSource = "SELECT * FROM FILE0_40"
    Set xStore.RowSource = data1
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
'    xStore.BoundText = retDef("FILE0_40")
    xStore.BoundText = "1"
    
    data4.ConnectionString = strCon
    data4.RecordSource = "SELECT * FROM MOSM"
    Set xMosm.RowSource = data4
    xMosm.ListField = "DESCA"
    xMosm.BoundColumn = "MOSM"
    cPMosm = cPMosm
    
    data2.ConnectionString = strCon
    data2.RecordSource = "SELECT * FROM file1_10sc order by desca "
    Set xSection.RowSource = data2
    xSection.ListField = "DESCA"
    xSection.BoundColumn = "code"
    xSection.BoundText = retDef("FILE1_10SC")
    
    
    With grid1
        .Cols = 10
        .Rows = 1
    End With
    Set grid1.DataSource = DATA3
    DATA3.ConnectionString = strCon
        
'    If Not (CardTable.EOF And CardTable.BOF) Then
'        CardTable.MoveLast
'        myload
'    Else
        mydefine
        Fixgrd
        'xDoc_No.Text = ""
'    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
    Err.Clear
End Sub
Private Sub grid1_DBLClick()
    If grid1.Col >= 7 Then
        cmdAddItems_Click
    Else
        If IsNumeric(grid1.TextMatrix(grid1.Row, 0)) Then
            ITEMS.bEdit = True
            ITEMS.bAddModel = True
            ITEMS.sModel = GetDesca("select model from file1_10 where item = " & MyParn(grid1.TextMatrix(grid1.Row, 0)))
            ITEMS.Show 1
        End If
    End If
End Sub
Private Sub grid1_EnterCell()
If grid1.TextMatrix(grid1.Row, 13) <> "" Then LoadPhoto grid1.TextMatrix(grid1.Row, 13) Else imgx1.Images.Clear
'If grid1.Col = 0 Then grid1.Editable = flexEDKbdMouse Else grid1.Editable = flexEDNone
End Sub

Private Sub grid1_LostFocus()
imgx1.Images.Clear
End Sub

Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then SuppLookupAll Me, Search3
End Sub
Private Sub xCode_LostFocus()
myLostFocus xCode
xCodeDesca.Caption = ""
xr_Discount.Text = ""
xsubcode.Caption = ""
xSupp.Value = 0
xsup_desca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 3)
If xCode.Text = "" Then Exit Sub

Dim cString As String, cwhere As String
cString = "select code,desca, SUPP , DISC ,SUBCODE  FROM FILE4_10"
cwhere = cwhere & turn(cwhere, " OR ") & " CODE = " & MyParn(xCode.Text)
cwhere = cwhere & turn(cwhere, " OR ") & " SUBCODE = " & MyParn(xCode.Text)
cString = cString & turn(cwhere) & cwhere
aret = aGetDesca(cString)
If UBound(aret) > 0 Then
    xCode.Text = aret(1)
    xCodeDesca.Caption = aret(2)
    xSupp.Value = IIf(aret(3), 1, 0)
    xsup_desca.Caption = IIf(aret(3), "„ﬂ » Ã„·…", "„’‰⁄")
    xr_Discount.Text = Val(aret(4) & "")
    xsubcode.Caption = aret(5) & ""
End If
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
End Sub

Private Sub xDate_Validate(Cancel As Boolean)
myValidDate xDate
End Sub

Private Function MYVALID() As Boolean
    If xDoc_No.Text = "" Then
        MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
        Exit Function
    End If
    
    If Not IsDate(xDate.Text) Then
        MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
        Exit Function
    End If
    
    
    If xStore.BoundText = "" Then
        MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
        Exit Function
    End If
    
    If xMosm.BoundText = "" Then
        MsgBox "·„ Ì „  ÕœÌœ «·„Ê”„ "
        Exit Function
    End If
    
    If xCodeDesca.Caption = "" Then
        MsgBox "·„ Ì „ «œŒ«· «·„Ê—œ"
        Exit Function
    End If
    
    MYVALID = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
    'xClosed.Value = CardTable!CLOSED
xDoc_No.Text = CardTable!doc_no
xInv_No.Text = CardTable!INV_NO & ""
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!store & ""
xMosm.BoundText = Trim(CardTable!Mosm)
xCode.Text = CardTable!CODE & ""
xCodeDesca.Caption = CardTable!CODEDESCA & ""
xSupp.Value = IIf(CardTable!SUPP, 1, 0)
xsup_desca.Caption = IIf(CardTable!SUPP, "„ﬂ » Ã„·…", "„’‰⁄")
xr_Discount.Text = CardTable!r_Discount
xsubcode.Caption = CardTable!subCode & ""
xTax.Text = TurnValue(Val(CardTable!tax & ""), 0, "")
'xDiscount.Text = TurnValue(Val(CardTable!DISCOUNT & ""), 0, "")
myloadgrd
End Sub
Public Sub myloadgrd()
With grid1
'                                      0        1           2           3                   4               5               6           7                   8       9       10          11              12      13
    cString = "SELECT " & cFile & ".ITEM,file1_10.mosm,FILE1_10.fact,file1_10.modelfact0,file1_10.desca ,file1_10.scal,file1_10.color,Quant," & cFile & ".Price, DISCOUNT, total   , FILE1_10.PRICE, ' ' as r,FILE1_10.MODELNO " & _
          " FROM " & cFile & " LEFT JOIN FILE1_10 ON " & cFile & ".ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " order by ROW  "
    DATA3.RecordSource = cString
    DATA3.Refresh
    grid1.AddItem ""
    If .Rows > 2 Then
        xCode.Enabled = False
        xDoc_No.Enabled = False
    Else
        xCode.Enabled = True
        xDoc_No.Enabled = True
    End If
End With
Handlecontrols LoadMode
CalcTotals
Fixgrd
End Sub
Private Sub mydefine()
    xClosed.Value = 0
    xsup_desca.Caption = ""
    cmdAddItems.Enabled = True
    xsubcode.Caption = ""
    xClosed.Visible = False
    xClosed.Value = ssCBUnchecked
    xStore.BoundText = "1"
    xCode.Enabled = True
    xInv_No.Text = ""
    xDate.Text = Format(Date, "dd-mm-yyyy")
    xDoc_No.Text = RetZero(Newflag(cFileHeader, "doc_no"))
    xCodeDesca.Caption = ""
    xBalance.Caption = ""
    xCode.Text = ""
    xDiscount.Text = ""
    
    xtotalOrg.Caption = ""
    xDiscountItem.Caption = ""
    xTotalItem.Caption = ""
    xr_Discount.Text = ""
    xTotalNoTax.Caption = ""
    xr_tax.Text = ""
    xTax.Text = ""
    xTotal.Caption = ""
    xtotalQuant.Caption = ""
'

'    xTotalDis.Caption = ""
    xUserName.Caption = ""

'    xRate.Text = ""
    xMosm.BoundText = cPMosm
    grid1.Rows = 1
    Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
    cmdNewInv.Enabled = nMode = LoadMode And bEdit
    cmdSave.Enabled = (bEdit) And (xClosed.Value = 0) And (xClosed.Value = 0)
    CmdDelInv.Enabled = nMode = LoadMode And bEdit And (xClosed.Value = 0)
    cmdAddItems.Enabled = (bEdit) And (xClosed.Value = 0) And (xClosed.Value = 0)
    cmdFirst.Enabled = (nMode = LoadMode)
    cmdLast.Enabled = (nMode = LoadMode)
    cmdNext.Enabled = (nMode = LoadMode)
    cmdPrevious.Enabled = (nMode = LoadMode)
    xClosed.Enabled = (nMode = LoadMode)
    xDoc_No.Enabled = (nMode = DefineMode)
    cmd_closed.Enabled = (nMode = LoadMode)
    xDoc_No.Tag = nMode
    If xClosed.Value = 1 Then
        cmd_closed.BackColor = &H8000000F
        cmd_closed.Caption = "› Õ „” ‰œ"
        xClosed.Value = ssCBChecked
        xClosed.Visible = True
    Else
        cmd_closed.BackColor = &H8080FF
        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
        xClosed.Value = ssCBUnchecked
        xClosed.Visible = False
    End If
    LoadPict xClosed.Value
End Sub
Private Sub xDiscount_Validate(Cancel As Boolean)
If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xr_Discount.Text), nRound) <> Round(Val(xDiscount.Text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xr_Discount.Text = Myvalue(Round((Val(xDiscount.Text) / Val(xTotalItem.Caption)) * 100, nRound))
    End If
Else
    xr_Discount.Text = ""
End If
CalcTotals
End Sub

Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload True
End Sub
Private Function CalcTotals()
Dim nTotalOrg As Double, nTotalItem As Double, nDiscountItem As Double
Dim nTotalQuant As Integer

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
    If Round(Val(xr_Discount.Text), nRound) <> Round(Val(xDiscount.Text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xDiscount.Text = Round((Val(xr_Discount.Text) * Val(xTotalItem.Caption)) / 100, nRound)
    End If
Else
    xDiscount.Text = ""
End If
xTotalNoTax.Caption = Format(Val(xTotalItem.Caption) - Val(xDiscount.Text), "#0.00")

If Val(xTotalNoTax.Caption) <> 0 Then
    If Round(Val(xr_tax.Text), nRound) <> Round(Val(xTax.Text) / Val(xTotalNoTax.Caption) * 100, nRound) Then
        xTax.Text = Round((Val(xr_tax.Text) * Val(xTotalNoTax.Caption)) / 100, nRound)
    End If
Else
    xTax.Text = ""
End If
xTotal.Caption = Format(Val(xTotalNoTax.Caption) + Val(xTax.Text))
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,[DATE] AS DATE1,DATE, " & cFileClient & ".Desca , inv_no , mosm " & _
                  " FROM  (" & cFileHeader & " left JOIN " & cFileClient & " ON " & cFileHeader & ".CODE " & " = " & cFileClient & ".CODE )"

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
If Val(xRate.Text) <> 0 Then
    xTax.Text = Format(Val(xTotalDis.Caption) * (Val(xRate.Text) / 100), "Fixed")
    CalcTotals
End If
End Sub
Private Sub Fixgrd()
With grid1
'                   0            1      2               3           4           5       6           7           8           9       10          11              12
.FormatString = "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "Œ’„ ’‰›|" & "«·≈Ã„«·Ì|" & "”⁄— „” Â·ﬂ|" & "„.|"
.WordWrap = True
.RowHeight(0) = 600
.ColWidth(0) = 800
.ColWidth(1) = 1000
.ColWidth(2) = 1800
.ColWidth(3) = 1000
.ColWidth(4) = 3500
.ColWidth(5) = 700
.ColWidth(6) = 1200
.ColWidth(7) = 800
.ColWidth(8) = 800
.ColWidth(9) = 0
.ColWidth(10) = 1000
.ColWidth(11) = 1000
.ColHidden(12) = True
.ColHidden(13) = True
.ColWidth(13) = 1000
'.ColComboList(0) = "..."
.ColComboList(2) = StrList("SELECT CODE , DESCA FROM FACT ")
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 12) = nRow
Next nRow
End With
End Sub
Private Sub xr_Discount_Validate(Cancel As Boolean)
CalcTotals
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
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourceTable As New ADODB.Recordset
Dim lPrice As Boolean
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
'If MsgBox("⁄—÷ ”⁄— «·„” Â·ﬂ", vbYesNo + vbDefaultButton2) = vbYes Then lPrice = True
For i = 1 To grid1.Rows - 2
        
    temptable.AddNew
    temptable!str10 = Me.Caption
    
    temptable!str6 = xDoc_No.Text
    temptable!Date1 = xDate.Text
    
    temptable!str7 = xCodeDesca.Caption
    temptable!str5 = TurnValue(xInv_No.Text)
    
    temptable!str1 = TurnValue(grid1.TextMatrix(i, 0))
    temptable!str13 = TurnValue(grid1.TextMatrix(i, 3))
    
    temptable!str2 = GetDesca("select desca from fact where code = " & MyParn(grid1.TextMatrix(i, 2)))
    temptable!str3 = TurnValue(grid1.TextMatrix(i, 4))
    temptable!str4 = TurnValue(grid1.TextMatrix(i, 5))
    temptable!str8 = TurnValue(grid1.TextMatrix(i, 6))
    
    temptable!val3 = Val(grid1.TextMatrix(i, 7))
    temptable!val1 = Val(grid1.TextMatrix(i, 8))
    temptable!VAL4 = Val(grid1.TextMatrix(i, 10))
'    If lPrice Then
        temptable!val2 = Val(grid1.TextMatrix(i, 11))
'    End If
    temptable!val5 = Val(xDiscount.Text)
    temptable!Val6 = Val(xTax.Text)
    temptable!VAL7 = Val(xTotalInv.Caption)
    temptable!VAL14 = i
    
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\Print_Purch.rpt"
main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Function mysave(Optional bMsg As Boolean = True) As Boolean
If Not MYVALID Then Exit Function
CalcTotals
If Not myreplace Then Exit Function
If bMsg Then Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
mysave = True
openCardTable
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

If CardTable!ISCLOSED Then
    con.Execute " update " & cFileHeader & " set isclosed = 0 where doc_no = " & MyParn(xDoc_No.Text)
Else
    con.Execute " update " & cFileHeader & " set isclosed = 1 where doc_no = " & MyParn(xDoc_No.Text)
End If
openCardTable
myUndo
End Sub

Private Sub cmd_CLOSEDDATE_Click()
    
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  „‰  «—ÌŒ", "1-1"), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  Õ Ï  «—ÌŒ", Format(Date, "DD-MM-YYYY")), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ ≈€·«ﬁ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update " & cFileHeader & " set isclosed = 1 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ ≈€·«ﬁ „” ‰œ«  «·› —…" & nRec
            openCardTable
            myUndo
        End If
    End If
End Sub
Private Sub cmd_open_Click()
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "› Õ  „” ‰œ«  „‰  «—ÌŒ", ""), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "› Õ „” ‰œ«  Õ Ï  «—ÌŒ", ""), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ › Õ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update " & cFileHeader & " set isclosed = 0 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ › Õ „” ‰œ«  «·› —…" & nRec
            openCardTable
            myUndo
        End If
    End If
End Sub
Private Sub CardLookupSupp()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(5, 1)
If xCode.Text = "" Then Exit Sub
Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,[DATE] AS DATE1,DATE, " & cFileClient & ".Desca , inv_no , mosm " & _
                  " FROM  (" & cFileHeader & " left JOIN " & cFileClient & " ON " & cFileHeader & ".CODE " & " = " & cFileClient & ".CODE ) WHERE FILE4_10.CODE = " & MyParn(xCode.Text)

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
Load Search1
Search1.Caption = "«” ⁄·«„"
Search1.Show 1
End Sub



Private Function myreplaceH()
Dim aInsert(5, 1)

aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(Purchasefrm.xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(Purchasefrm.xDate.Text)

aInsert(2, 0) = "Code"
aInsert(2, 1) = addstring(Purchasefrm.xCode.Text)

aInsert(3, 0) = "Store"
aInsert(3, 1) = addstring(Purchasefrm.xStore.BoundText)

aInsert(4, 0) = "MOSM"
aInsert(4, 1) = addstring(Purchasefrm.xMosm.BoundText)

aInsert(5, 0) = "INV_NO"
aInsert(5, 1) = addstring(Purchasefrm.xInv_No.Text)
con.BeginTrans
If Purchasefrm.xDoc_No.Text = "" Then
    xDoc_No.Text = RetZero(Newflag(Purchasefrm.cFileHeader, "doc_no"))
    aInsert(0, 1) = addstring(Purchasefrm.xDoc_No.Text)
    con.Execute CreateInsert(aInsert, Purchasefrm.cFileHeader)
End If
con.CommitTrans
Exit Function
myError:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function myreplaceGrd() As Boolean
Dim loctable As ADODB.Recordset, cString As String
cString = " DELETE FROM " & cFile & " FROM   FILE1_10 INNER JOIN " & cFile & "  ON FILE1_10.ITEM = " & cFile & ".ITEM"
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.Text)
cString = cString & turn(cString) & "FILE1_10.MODEL = " & MyParn(addModelFrm.xModel.Caption)
con.Execute cString, nDel

With addModelFrm.grid1
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                Set loctable = itemFind(addModelFrm.Grid2.TextMatrix(nRow, nCol), con)
                If Not (loctable.EOF And loctable.BOF) Then
                    cString = "Insert into " & cFile & " (doc_no,item,MODEL,price,Quant)" & _
                               "Values(" & _
                               addstring(xDoc_No.Text) & "," & _
                               addvalue(loctable!Item) & "," & _
                               addstring(loctable!Model) & "," & _
                               Val(loctable!COST) & "," & _
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
'    Inform " „  «÷«›… «·„ÊœÌ·«  »‰Ã«Õ"
End If
myloadgrd
grid1.ShowCell grid1.Rows - 1, 2 + 1
grid1.Select grid1.Rows - 1, 2 + 1
End Sub
Private Sub LoadPict(pValue As Long)
'On Error Resume Next
cmd_closed.Picture = LoadPicture(App.Path & "\SYSIMG\" & IIf(pValue = 0, "CLOSE_S.jpg", "OPEN_S.jpg"))
cmd_closed.PictureDisabled = LoadPicture(App.Path & "\SYSIMG\" & IIf(pValue = 0, "CLOSE_S_D.jpg", "OPEN_S_D.jpg"))
Err.Clear
End Sub

Private Sub OPENCARDTA()
    Set CardTable = New ADODB.Recordset
    CardTable.Open "SELECT * FROM " & cFileHeader & "  ORDER BY DOC_NO", con, adOpenKeyset, adLockReadOnly, adCmdText
    
    If sDoc_No <> "" And Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "doc_no = " & MyParn(sDoc_No), , adSearchForward, adBookmarkFirst
        If Not CardTable.EOF Then
            myload
            Exit Sub
        End If
    End If

End Sub
Private Sub myUndo()
'On Error GoTo myError
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xDoc_No.Text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
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
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT " & cFileHeader & ".*,FILE4_10.DESCA AS CODEDESCA,FILE4_10.SUPP,FILE4_10.DISC,FILE4_10.SUBCODE FROM " & _
           cFileHeader & " INNER JOIN FILE4_10 ON " & cFileHeader & ".CODE = FILE4_10.CODE"
If sDoc_No <> "" Then cString = cString & turn(cString) & "DOC_NO = " & MyParn(sDoc_No)
cString = cString & " ORDER BY DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myError
imgx1.Images.Clear
imgx1.Import.FromFile retFile(sModelNo)
Exit Sub
myError:
Err.Clear
End Sub

