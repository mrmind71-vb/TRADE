VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form trans_FR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕÊÌ·«   ÊﬂÌ·«  "
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   9990
   ScaleWidth      =   15315
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Caption         =   " ”ÃÌ· »«·»«—ﬂÊœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   135
      Width           =   1860
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "»œÊ‰  ÕœÌÀ"
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
      Left            =   5670
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   1665
      Width           =   1275
   End
   Begin VB.Frame Frame7 
      Height          =   600
      Left            =   2295
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   0
      Width           =   3300
      Begin VB.Label XTRANS_DOC 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   45
         TabIndex        =   51
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "„” ‰œ  ÕÊÌ· „‰ «·›—⁄ "
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
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   180
         Width           =   1725
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   510
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   8955
      Width           =   2265
   End
   Begin VB.CommandButton Command1 
      Caption         =   "send items"
      Height          =   375
      Left            =   5670
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   45
      Width           =   1275
   End
   Begin VB.CommandButton CMD_ACCEPT 
      BackColor       =   &H000080FF&
      Caption         =   "≈” ·«„ «· ÕÊÌ· „‰ «·—∆Ì”ÌÏ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   180
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   945
      Width           =   2085
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Height          =   690
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   2025
      Width           =   15000
      Begin VB.Label xTimeSend 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   12195
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   225
         Visible         =   0   'False
         Width           =   2715
      End
      Begin VB.Label xTimeAccept 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   225
         Width           =   2580
      End
      Begin VB.Label xUserSend 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8685
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   225
         Width           =   2310
      End
      Begin VB.Label xUserAccept 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   225
         Width           =   2085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "User Send"
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
         Left            =   7605
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "User accept"
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time Send "
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
         Left            =   11070
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   180
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Time accept"
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
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   180
         Width           =   1245
      End
   End
   Begin VB.CommandButton TRANS_MODEL 
      Caption         =   " ÕÊÌ· „ÊœÌ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   90
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Frame Frame4 
      Height          =   510
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   1575
      Visible         =   0   'False
      Width           =   3300
      Begin VB.Label xCount 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label xTq 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   8910
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   25
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
         Picture         =   "TRANS_FR.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TRANS_FR.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   26
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
         Picture         =   "TRANS_FR.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TRANS_FR.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   27
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
         Picture         =   "TRANS_FR.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TRANS_FR.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   28
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
         Picture         =   "TRANS_FR.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TRANS_FR.frx":EB26
      End
   End
   Begin VB.CheckBox xsent 
      Alignment       =   1  'Right Justify
      Caption         =   "Sent"
      Enabled         =   0   'False
      Height          =   285
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1935
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox xReceived 
      Alignment       =   1  'Right Justify
      Caption         =   "xReceived"
      Enabled         =   0   'False
      Height          =   285
      Left            =   4365
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1755
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   405
      Width           =   1365
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TRANS_FR.frx":10C75
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Picture         =   "TRANS_FR.frx":131EE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   9810
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   5325
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TRANS_FR.frx":15551
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
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
         Picture         =   "TRANS_FR.frx":179BD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "TRANS_FR.frx":1A257
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Height          =   510
         Left            =   3960
         Picture         =   "TRANS_FR.frx":1C803
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1095
      Left            =   2295
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   540
      Width           =   3300
      Begin VB.CommandButton cmdPrint 
         Caption         =   "ÿ»«⁄… «·„” ‰œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   2025
         Picture         =   "TRANS_FR.frx":1EFD6
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   135
         Width           =   1230
      End
      Begin Threed.SSCommand CMD_PrintBar 
         Height          =   465
         Left            =   45
         TabIndex        =   11
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   820
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
         Picture         =   "TRANS_FR.frx":21400
         Caption         =   " ÿ»«⁄… »«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand CMD_BAR 
         Height          =   465
         Left            =   45
         TabIndex        =   12
         Top             =   585
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   820
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
         Picture         =   "TRANS_FR.frx":23D82
         Caption         =   "  ÕÊÌ· ··»«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   7020
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   630
      Width           =   8115
      Begin VB.CommandButton CMD_SEND 
         BackColor       =   &H00FFFF00&
         Caption         =   "Send"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   945
         Width           =   555
      End
      Begin VB.TextBox XNotes_Doc 
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
         Left            =   720
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   945
         Width           =   6180
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
         Left            =   5400
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1500
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
         Left            =   1350
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo xStore1 
         Height          =   315
         Left            =   4140
         TabIndex        =   2
         Top             =   570
         Width           =   2760
         _ExtentX        =   4868
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
      Begin MSDataListLib.DataCombo xStore2 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   570
         Width           =   2715
         _ExtentX        =   4789
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
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   990
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "≈·Ì „Œ“‰ :"
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   495
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Œ“‰ :"
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
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   495
         Width           =   855
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   135
         Width           =   645
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
         Left            =   6975
         RightToLeft     =   -1  'True
         TabIndex        =   6
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
      Left            =   1620
      Top             =   495
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
      Height          =   6180
      Left            =   135
      TabIndex        =   4
      Top             =   2700
      Width           =   15000
      _cx             =   26458
      _cy             =   10901
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
      TabIndex        =   29
      Top             =   9570
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
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
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
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
End
Attribute VB_Name = "trans_FR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean, sFlagTRans As String
Public sDoc_no As String
Dim bAct As Boolean
Dim cFilter As String
Dim con As New ADODB.Connection
Dim oSearchDoc As New Search3, oSearchItem As New Search3
Dim CardTable As ADODB.Recordset
Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
'5797
Private Function myreplace(Optional bModel As Boolean, Optional nRow As Long = -1) As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "[DATE]", DateSq(xDate.text))
aInsert = AddFlag(aInsert, "[STORE1]", addstring(xStore1.BoundText))
aInsert = AddFlag(aInsert, "[STORE2]", addstring(xStore2.BoundText))
aInsert = AddFlag(aInsert, "[isdate]", "getdate()")
aInsert = AddFlag(aInsert, "[USERSEND]", addstring(cusername))
aInsert = AddFlag(aInsert, "[Notes_Doc]", addstring(XNotes_Doc.text))
aInsert = AddFlag(aInsert, "[ISNEW1]", 1)
aInsert = AddFlag(aInsert, "[ISNEW2]", 1)
con.BeginTrans
'On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    xDoc_No.text = RetZero(Newflag("FR1_60H", "DOC_NO", con))
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(xDoc_No.text))
    con.Execute addInsert(aInsert, "FR1_60h")
Else
    con.Execute addUpdate(aInsert, "FR1_60h", "doc_no = " & addstring(xDoc_No.text))
End If
'For nbr = 2 To nCountBranch
'    sStore = RetZero(nbr, 2)
'    con.Execute "  INSERT INTO NEW1_60H (DOC_NO , BRANCH) VALUES        ( " & addstring(xDoc_No.Text) & " , " & addstring(sStore) & " )  "
'Next nbr
If bModel Then
    myreplaceGrdModel
Else
    myreplaceGrd nRow
End If
con.CommitTrans
myreplace = True
Exit Function
myError:
'prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
On Error GoTo myError
If ActiveControl.Name = grid1.Name Then
    bNew = grid1.Row = grid1.Rows - 1
    Dim nFound As Integer
    nFound = grid1.FindRow(oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    grid1.TextMatrix(grid1.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 8) = "1"
    grid1_AfterEdit grid1.Row, 0
    If Not bNew Then
        CellPos 13, grid1.Row, grid1.Col
        Unload oSearchItem
    Else
        grid1.Row = grid1.Rows - 1
    End If
Else
    xDoc_No.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    myUndo
    Unload oSearchDoc
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
Unload Search
End Sub
Private Sub CMD_ACCEPT_Click()
    On Error GoTo myError
    strConShop = LoadConStringshop
    If MsgBox(" „  „—«Ã⁄… «·„” ‰œ Ê »Ì«‰ «·«’‰«› ’ÕÌÕ", vbYesNo + vbDefaultButton2) = vbYes Then
        If openConShop(conShop) <> "ok" Then
            MsgBox "·« ÌÊÃœ « ’«· »«·”Ì—›—"
            Exit Sub
        End If
        conShop.Execute " UPDATE FR1_60h SET Received = 1 , USERACCEPT = " & addstring(cusername) & " , TIMEACCEPT = GETDATE() WHERE DOC_NO = " & MyParn(xDoc_No.text), nRec
        con.Execute " UPDATE FR1_60h SET Received = 1 , USERACCEPT = " & addstring(cusername) & " , TIMEACCEPT = GETDATE() WHERE DOC_NO = " & MyParn(xDoc_No.text)
        con.Execute " delete from model_dem where item in (select item from FR1_60 where doc_no = " & MyParn(xDoc_No.text) & " ) "
        If nRec = 1 Then
            Inform " „  „—«Ã⁄… «·„” ‰œ Ê »Ì«‰ «·«’‰«› ’ÕÌÕ"
            
            myUndo
        End If
    End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CMD_BAR_Click()
On Error GoTo myError
    With grid1
        con.BeginTrans
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 8)) > 0 Then
                con.Execute "Insert Into ADDPRINT(Item,BRANCH,Quant,isPrint) " & _
                    " Values(" & _
                    addstring(.TextMatrix(i, 0)) & "," & _
                    addstring(cBranch) & "," & _
                    addvalue(.TextMatrix(i, 8)) & "," & _
                    "1" & _
                    ")"
            End If
        Next i
        con.CommitTrans
    End With
    Inform " „ «· —ÕÌ·"
Exit Sub
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub CMD_PrintBar_Click()
    BarCodeFRM.Show 1
End Sub

Private Sub CMD_SEND_Click()
    On Error GoTo myError
    strConShop = LoadConStringshop
    If openConShop(conShop) <> "ok" Then
        MsgBox "·« ÌÊÃœ « ’«· »«·”Ì—›—"
        Exit Sub
    End If
    conShop.Execute " UPDATE FR1_60H  SET Notes_Doc = " & addstring(XNotes_Doc) & " where  DOC_NO = " & MyParn(xDoc_No.text), nRec
    con.Execute " UPDATE FR1_60H      SET Notes_Doc = " & addstring(XNotes_Doc) & " where  DOC_NO = " & MyParn(xDoc_No.text), nRec
    If nRec = 1 Then
        Inform "  „ «—”«· «·»Ì«‰ ··”Ì—›—"
    End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    con.BeginTrans
    On Error GoTo myError
    con.Execute " DELETE FROM NEW1_60H WHERE DOC_NO = " & addstring(xDoc_No.text)
    For nRow = 2 To nCountBranch
        sStore = RetZero(nRow, 2)
        con.Execute "  INSERT INTO NEW1_60H (DOC_NO , BRANCH) VALUES        ( " & addstring(xDoc_No.text) & " , " & addstring(sStore) & " )  "
    Next nRow
    
    con.Execute " Delete  From FR1_60 where Doc_No = " & MyParn(xDoc_No.text)
'   con.Execute " Delete  From FR1_60H where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute " update FR1_60h set ISNOITEM = 0  , [isdate] = getdate() , ISNEW1 = 1 , ISNEW2  = 1 , ISNEW3 = 1 , ISNEW4 = 1 , ISNEW5 = 1 where doc_no = " & MyParn(xDoc_No.text)
    con.CommitTrans
''''''    If XTRANS_DOC.Caption <> "" Then
''''''        con.Execute " update fr1_61h set DOC_TRANS = null , TYPEDOC = null   , ISNEW = 1   where doc_no = " & MyParn(XTRANS_DOC.Caption)
''''''    End If
    myUndo
    Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
    CmdNewInv_Click
End If
Exit Sub
myError:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
     Unload Me
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(3, 4)
Dim GrdArray(6, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT   top 200 DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),BRANCH_FR.DESCA,BRANCH_FR_1.DESCA , Notes_Doc , TRANS_DOC " & _
                  " FROM (FR1_60H INNER JOIN BRANCH_FR ON FR1_60H.Store1 = BRANCH_FR.CODE) INNER JOIN BRANCH_FR AS BRANCH_FR_1 ON FR1_60H.STORE2 = BRANCH_FR_1.CODE "

Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

listarray(1, 0) = "„‰ „Œ“‰"
listarray(1, 1) = "BRANCH_FR.DESCA Like '%cFilter%'  "

listarray(2, 0) = "≈·Ï „Œ“‰"
listarray(2, 1) = "BRANCH_FR_1.DESCA Like '%cFilter%'  "

listarray(3, 0) = "„·«ÕŸ«  «·›—Ê⁄-„” ‰œ"
listarray(3, 1) = "(Notes_Doc Like '%cFilter%' OR TRANS_Doc Like '%cFilter%') "

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "„‰ ›—⁄ "
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "≈·Ì ›—⁄ "
GrdArray(4, 1) = 2000

GrdArray(5, 0) = "„·«ÕŸ«  «·›—Ê⁄"
GrdArray(5, 1) = 2000

GrdArray(6, 0) = " ÕÊÌ· Ê«—œ "
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
Dim lIsNoItem As Boolean
'foundOther
If Check1.Value <> 0 Then
'    con.Execute " DELETE FROM NEW1_60H WHERE DOC_NO = " & addstring(xDoc_No.Text)
'    For nRow = 2 To nCountBranch
'        sStore = RetZero(nRow, 2)
'        con.Execute "  INSERT INTO NEW1_60H (DOC_NO , BRANCH) VALUES        ( " & addstring(xDoc_No.Text) & " , " & addstring(sStore) & " )  "
'    Next nRow
End If

If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub

cStr1 = " UPDATE FR1_60 SET COST = FILE1_10.COSTITEM_FR FROM   FR1_60 INNER JOIN FILE1_10 ON FR1_60.ITEM = FILE1_10.ITEM  WHERE FR1_60.DOC_NO = " & MyParn(xDoc_No.text)
con.Execute cStr1, nRec
    
'con.Execute " delete from model_dem where STORE = " & MyParn(xStore2.BoundText) & " AND item in (select item from FR1_60 where doc_no = " & MyParn(xDoc_No.Text) & " ) "

'lIsNoItem = IIf(GetDesca("SELECT ITEM FROM FILE1_10 WHERE ISNOITEM = 0 AND ITEM IN (SELECT ITEM FROM FR1_60 WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " )") = "", True, False)
'If lIsNoItem Then
'    con.Execute " UPDATE FR1_60H SET ISNOITEM = 1 where doc_no = " & MyParn(xDoc_No.Text)
'Else
'    con.Execute " UPDATE FR1_60H SET ISNOITEM = 0 where doc_no = " & MyParn(xDoc_No.Text)
'End If
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"

myUndo
End Sub
Private Sub CmdUndo_Click()
    myUndo
End Sub
Private Sub Command1_Click()
    Dim cBrc As String
    cBrc = GetDesca("select CODE  from BRANCH_FR where store = " & MyParn(xStore2.BoundText), con)
    If cBrc <> "" Then
        con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBrc) & " FROM  FILE1_10  WHERE ITEM IN (SELECT ITEM FROM FR1_60 WHERE DOC_NO = " & MyParn(xDoc_No.text) & " ) ", nRec
        Inform "  „ «÷«›… " & nRec
    End If
End Sub
Private Sub Command2_Click()
    cStr1 = " UPDATE    FR1_60 SET  cost = (SELECT TOP 1 (FILE7_20.COST_INV ) FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FR1_60.ITEM = FILE7_20.ITEM AND FILE7_20H.DATE <= FR1_60H.DATE ORDER BY DATE DESC ) " & _
            " FROM      FR1_60 INNER JOIN FR1_60H ON FR1_60.DOC_NO = FR1_60H.DOC_NO "
    con.Execute cStr1, nRec
    Inform "  „ Ÿ»ÿ ⁄œœ  " & nRec
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 And KeyCode = 83 Then cmdSave_Click
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Me.Caption = sCaption
Check3.Value = Val(RetSetting("Check31", App.Path & "\other.txt"))

openCon con
CMD_ACCEPT.Visible = (cBranch <> "00")
Command1.Visible = (cBranch = "00")
CMD_SEND.Visible = (cBranch <> "00")

'Set CardTable = New ADODB.Recordset

Set data1.Recordset = myRecordSet("SELECT * from BRANCH_FR WHERE OWNER IS NOT NULL ORDER BY code ", con)
Set xStore1.RowSource = data1
xStore1.ListField = "Desca"
xStore1.BoundColumn = "Code"


Set DATA2.Recordset = myRecordSet("SELECT * from BRANCH_FR WHERE OWNER IS NOT NULL ORDER BY code ", con)
Set xStore2.RowSource = DATA2
xStore2.ListField = "Desca"
xStore2.BoundColumn = "Code"

Set grid1.DataSource = DATA10
'DATA10.ConnectionString = strCon
myUndo

cmdSave.Visible = (cBranch = "00")
CmdDelInv.Visible = (cBranch = "00")
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload oSearchItem
Unload oSearchDoc
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
Set Transfrm = Nothing
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)


Dim bNew As Boolean
If grid1.Col = 0 Then
    GrdDesc grid1.Row
End If

If Not validRow(Row) Then Exit Sub
If Row = grid1.Rows - 1 Then
    myaddItem
    bNew = True
End If
CalcTotals

MYVALID (False)

With grid1
If myreplace(, Row) Then
    If xDoc_No.Tag = DefineMode Then xDoc_No.Tag = LoadMode
Else
    bNew = False
End If
myloadgrd
If bNew Then
    If Check3.Value = 0 Then
        grid1.Row = grid1.Row
        grid1.Col = 8
        grid1.ShowCell grid1.Row, 8
    Else
        grid1.Row = grid1.Rows - 1
        grid1.ShowCell grid1.Rows - 1, 0
        grid1.Col = 0
    End If
End If

End With

End Sub
Private Sub grid1_EnterCell()
With grid1
If (grid1.Col = 0 Or grid1.Col = 8) Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub
Private Sub Grid1_GotFocus()
With grid1
    If grid1.Row <= 1 And grid1.Rows > 1 Then
    .Select 1, 0, 1, 0
    '.ShowCell 1, 0
    End If
End With
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 2 And Trim(grid1.TextMatrix(Row, 3)) = "" Then
'    nBalance = LastBalance(grid1.TextMatrix(Row, 0), xStore1.BoundText, con)
'    grid1.TextMatrix(Row, 9) = nBalance
End If
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    With grid1
    If .EditText <> "" And Len(.EditText) > 2 Then
        If UCase(Mid(.EditText, 1, 1)) = "J" Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(.EditText), con)
        ElseIf UCase(Mid(.EditText, 1, 1)) = "S" Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(.EditText), con)
        ElseIf Len(.EditText) >= 10 Then
            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE13 = " & MyParn(.EditText) & " OR BARCODE = " & MyParn(.EditText) & " OR BARCODE_GS1 = " & MyParn(.EditText), con)
            If cMyItem = "" Then cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE = " & MyParn(.EditText), con)
        End If
        If cMyItem <> "" Then
            .EditText = Val(cMyItem)
        End If
    End If
    End With
    If Not validItem(grid1.EditText, con) Then
        Inform "»«—ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        Cancel = True
    End If
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
If xDoc_No.text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore1.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·«Ê·"
    Exit Function
End If

If xStore2.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·À«‰Ì"
    Exit Function
End If

If bCheckRows Then
    If grid1.Rows < 3 Then
        MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
        Exit Function
    End If

    With grid1
    For i = 1 To .Rows - 2
        If .TextMatrix(i, 0) = "" Then
            .Select i, 0, i, grid1.Cols - 1
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        ElseIf Not validItem(.TextMatrix(i, 0), con) Then
            MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
            Exit Function
        End If
        If Val(.TextMatrix(i, 8)) = 0 Then
            .Select i, 0, i, grid1.Cols - 1
            MsgBox "ﬂ„Ì… «·’‰› €Ì— „”Ã·…"
            Exit Function
        End If
    Next
    End With
End If
MYVALID = True
End Function

Private Function MYVALID2() As Boolean
If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
If xStore1.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·›—⁄ «·«Ê·"
    Exit Function
End If

If xStore2.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·›—⁄ «·À«‰Ì"
    Exit Function
End If
MYVALID2 = True
End Function
Private Sub myload()
xDoc_No.text = CardTable!doc_no
XNotes_Doc.text = CardTable!Notes_Doc & ""
XTRANS_DOC.Caption = CardTable!TRANS_DOC & ""
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
xStore1.BoundText = CardTable!STORE1 & ""
xStore2.BoundText = CardTable!STORE2 & ""

xUserAccept.Caption = CardTable!UserAccept & ""
xUserSend.Caption = CardTable!USERSEND & ""
xTimeAccept.Caption = CardTable!TimeAccept & ""
xTimeSend.Caption = CardTable!TimeSend & ""
xReceived.Value = IIf(CardTable!Received, 1, 0)
xsent.Value = IIf(CardTable!SENT, 1, 0)
CMD_ACCEPT.Enabled = (xReceived.Value = 0)
xDoc_No.Enabled = False
xDoc_No.Tag = LoadMode
myloadgrd
Handlecontrols LoadMode
StatusBar1.Panels(1).text = "”Ã· " & CardTable.AbsolutePosition & " „‰ " & CardTable.RecordCount
StatusBar1.Panels(2).text = "⁄œœ «·”Ã·«  : " & CardTable.RecordCount
grid1.Select grid1.Rows - 1, 0
grid1.ShowCell grid1.Rows - 1, 0
CalcTotals
End Sub
Private Sub myDefine()
XTRANS_DOC.Caption = ""
xDoc_No.Enabled = True
xDoc_No.text = RetZero(Newflag("FR1_60H", "DOC_NO", con))
xDate.text = Format(Date, "dd-mm-yyyy")
xStore1.BoundText = ""
xStore2.BoundText = ""
StatusBar1.Panels(1).text = ""
xCount.Caption = ""
xTq.Caption = ""
StatusBar1.Panels(3).text = ""
StatusBar1.Panels(4).text = ""
XNotes_Doc.text = ""
xUserAccept.Caption = ""
xUserSend.Caption = ""
xTimeAccept.Caption = ""
xTimeSend.Caption = ""
xReceived.Value = 0
If cmdSave.Enabled Then
    StatusBar1.Panels(1).text = "«÷«›… ”Ã· —ﬁ„ " & (CardTable.RecordCount + 1)
    StatusBar1.Panels(2).text = "⁄œœ «·”Ã·«  : " & CardTable.RecordCount
ElseIf CardTable.EOF And CardTable.BOF Then
    StatusBar1.Panels(1).text = ""
    StatusBar1.Panels(2).text = "·«  ÊÃœ ”Ã·« "
End If
grid1.Rows = 1
myaddItem

Handlecontrols DefineMode
Fixgrd
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode
cmdSave.Enabled = (xReceived.Value = 0) Or bopt2
CmdDelInv.Enabled = nMode = LoadMode And ((xReceived.Value = 0) Or bopt2)
TRANS_MODEL.Enabled = nMode = LoadMode And ((xReceived.Value = 0) Or bopt2)
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
If xDoc_No.text = "" Then Exit Sub
xDoc_No.text = RetZero(xDoc_No.text, 6)
myUndo
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If Not cmdSave.Enabled Or Not cmdSave.Visible Then Exit Sub
If KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And nFlag = 0 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myError
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from FR1_60 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
        CalcTotals
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
If grid1.TextMatrix(Row, 0) = "" Then Exit Sub
If grid1.TextMatrix(Row, 1) = "" Then Exit Sub
KeyCode = 0
If Col = 0 Then
    If Check3.Value = 0 Then
        If IsNumeric(grid1.TextMatrix(grid1.Row, 0)) Then grid1.Col = 8
    Else
        If Row <= grid1.Rows - 2 Then
            grid1.ShowCell grid1.Row, 0
            grid1.Row = Row + 1
            grid1.Col = 0
        Else
            grid1.Col = 8
        End If
    End If
ElseIf Col = 8 Then
    If Row <= grid1.Rows - 2 Then
        grid1.ShowCell grid1.Row, 0
        grid1.Row = Row + 1
        grid1.Col = 0
    Else
        grid1.Col = 8
    End If
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
Select Case Col
    Case 0
        If KeyCode = 27 Then Exit Sub
        If KeyCode = 112 Then ItemsLookupAll Me, oSearchItem
End Select
End Sub
Private Sub GrdDesc(Row)
Dim nBalance As Double
If Not IsNumeric(grid1.TextMatrix(Row, 0)) Then Exit Sub
'nBalance = LastBalance(grid1.TextMatrix(Row, 0), xStore1.BoundText, con)
'grid1.TextMatrix(Row, 9) = nBalance

Dim ItemTable As New ADODB.Recordset
Set ItemTable = ItemFind(grid1.TextMatrix(Row, 0), con)
With grid1
    If Not ItemTable.EOF Then
        If ItemTable!ISSTOP Then
            MsgBox "’‰› „ Êﬁ›"
            Exit Sub
        End If
        grid1.TextMatrix(grid1.Row, 1) = ItemTable!MOSM
        grid1.TextMatrix(grid1.Row, 2) = GetField("select desca from fact where code = " & MyParn(ItemTable!Fact), con) & ""
        grid1.TextMatrix(grid1.Row, 3) = ItemTable!CODE & ""
        grid1.TextMatrix(grid1.Row, 4) = ItemTable!modelfact0
        grid1.TextMatrix(grid1.Row, 5) = ItemTable!DESCA
        grid1.TextMatrix(grid1.Row, 6) = ItemTable!scal
        grid1.TextMatrix(grid1.Row, 7) = ItemTable!Color
        If Check3.Value <> 0 Then grid1.TextMatrix(grid1.Row, 8) = 1
    End If
End With
Set ItemTable = Nothing
'If Val(grid1.TextMatrix(Row, 9)) <= 0 Then
'    Inform "·« ÌÊÃœ —’Ìœ ··’‰›"
'End If
End Sub
Private Function CalcTotals()
Dim nTotalQuant As Double
Dim nTotal As Double
With grid1
For i = 1 To grid1.Rows - 2
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 8))
    nTotal = nTotal + (Val(grid1.TextMatrix(i, 8)) * Val(grid1.TextMatrix(i, 9)))
Next
StatusBar1.Panels(3) = ""
StatusBar1.Panels(4) = ""
StatusBar1.Panels(5) = ""
If grid1.Rows > 1 Then StatusBar1.Panels(3) = "⁄œœ «·√’‰«› : " & grid1.Rows - 2
If nTotalQuant > 0 Then StatusBar1.Panels(4) = "≈Ã„«·Ì «·ﬂ„Ì… : " & nTotalQuant
StatusBar1.Panels(5) = "≈Ã„«·Ï «·ﬁÌ„… : " & nTotal
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
Private Sub foundOther()
Dim nRow As Long
For i = 1 To grid1.Rows - 2
    nRow = grid1.FindRow(grid1.TextMatrix(i, 0), i + 1, 0)
    If nRow > -1 Then
        Inform "«·’‰› " & grid1.TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ ====> " & nRow
    End If
Next
End Sub
Private Sub doprint()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim nTot As Double
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
sourcetable.Open "SELECT FR1_60.* ,FACT.DESCA AS FACTNAME ,FILE1_10.SCAL,FILE1_10.COLOR,FILE1_10.MOSM,FILE1_10.DESCA,MODELFACT0,FILE1_10.MOSM , FILE1_10.PRICE AS PRICE_2 , FILE1_10.MODEL FROM FR1_60 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FR1_60.ITEM INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT WHERE DOC_NO = " & MyParn(xDoc_No.text), con, adOpenStatic, adLockReadOnly, adCmdText
nTot = Val(GetDesca("SELECT SUM(QUANT) FROM FR1_60 WHERE DOC_NO = " & MyParn(xDoc_No.text), con) & "")

With sourcetable
    Do While Not .EOF
        temptable.AddNew
        temptable!str6 = xDoc_No.text
        temptable!str8 = TurnValue(XNotes_Doc.text)
        temptable!Date1 = DateValue(xDate.text)
        temptable!str14 = TurnValue(xStore1.text)
        temptable!str15 = TurnValue(!BOX_NO)
    
        temptable!str13 = TurnValue(xStore2.text)
    
        temptable!str4 = TurnValue(!Item)
        temptable!str16 = TurnValue(!MODEL)
        temptable!STR11 = TurnValue(!MOSM)
        temptable!str12 = TurnValue(!modelfact0)
        temptable!str10 = TurnValue(!FACTNAME)
        temptable!str2 = TurnValue(!DESCA)
        temptable!str5 = TurnValue(!scal)
        temptable!STR7 = TurnValue(!Color)
        temptable!VAL1 = Val(!Quant)
        temptable!val2 = Val(!PRICE_2)
        temptable!val18 = nTot
        temptable!str1 = TurnValue(grid1.TextMatrix(i, 2)) & TurnValue(grid1.TextMatrix(i, 1)) & RetZero(TurnValue(grid1.TextMatrix(i, 4)), 10)
        temptable!Val10 = i
        temptable.Update
        
        .MoveNext
    Loop
End With

'''For i = 1 To grid1.Rows - 2
'''    temptable.AddNew
'''    temptable!str6 = xdoc_no.text
'''    temptable!str8 = TurnValue(xNotes.text)
'''    temptable!Date1 = DateValue(xdate.text)
'''    temptable!str14 = TurnValue(xStore1.text)
'''    temptable!str15 = TurnValue(grid1.TextMatrix(i, 11))
'''
'''    temptable!str13 = TurnValue(XSTORE2.text)
'''
'''    temptable!str4 = TurnValue(grid1.TextMatrix(i, 0))
'''    temptable!str16 = TurnValue(GetDesca("SELECT MODEL FROM FILE1_10 WHERE ITEM = " & Val(grid1.TextMatrix(i, 0))))
'''    temptable!str11 = TurnValue(grid1.TextMatrix(i, 1))
'''    temptable!str12 = TurnValue(grid1.TextMatrix(i, 4))
'''    temptable!str10 = TurnValue(grid1.TextMatrix(i, 2))
'''    temptable!str2 = TurnValue(grid1.TextMatrix(i, 5))
'''    temptable!str5 = TurnValue(grid1.TextMatrix(i, 6))
'''    temptable!STR7 = TurnValue(grid1.TextMatrix(i, 7))
'''    temptable!VAL1 = Val(grid1.TextMatrix(i, 8))
'''    temptable!val2 = Val(GetDesca("select price from file1_10 where item = " & grid1.TextMatrix(i, 0)) & "")
'''    temptable!val18 = nTot
'''    temptable!str1 = TurnValue(grid1.TextMatrix(i, 2)) & TurnValue(grid1.TextMatrix(i, 1)) & RetZero(TurnValue(grid1.TextMatrix(i, 4)), 10)
'''    temptable!Val10 = i
'''    temptable.Update
'''Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans

main.REPORT1.ReportFileName = App.Path & "\Reports\TRANS.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1

temptable.Close
Set temptable = Nothing
End Sub
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
With grid1
'                   0            1      2               3        4           5       6           7          8           9
.FormatString = "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„Ê—œ|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "«·ﬂ„Ì…|" & "«·”⁄—|"
.WordWrap = True
.RowHeight(0) = 500
.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 1800
.ColWidth(3) = 1000
.ColWidth(4) = 1500
.ColWidth(5) = 3600
.ColWidth(6) = 900
.ColWidth(7) = 900
.ColWidth(8) = 900
.ColWidth(9) = 900

.ColHidden(9) = (cBranch > "00" And cBranch < "60")
.ColHidden(.Cols - 1) = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub myreplaceGrd(Optional nRow As Long = -1)
Dim aInsert As Variant, i As Long, bSaveRow As Boolean
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert = AddFlag(Empty, "[DOC_NO]", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "[ITEM]", addvalue(grid1.TextMatrix(i, 0)))
        aInsert = AddFlag(aInsert, "[QUANT]", Val(grid1.TextMatrix(i, 8)))
        aInsert = AddFlag(aInsert, "[COST]", Val(ItemField(grid1.TextMatrix(i, 0), "COSTITEM_FR", con) & ""))
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute addInsert(aInsert, "FR1_60")
        Else
            con.Execute addUpdate(aInsert, "FR1_60", " ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub myloadgrd()
Dim cString As String
'                    0               1               2               3           4                       5               6               7               8                   9
cString = "SELECT FR1_60.item , file1_10.mosm , fact.desca , file1_10.CODE , file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  FR1_60.quant , FR1_60.cost ,FR1_60.ID from FR1_60 inner join file1_10 on file1_10.item = FR1_60.item  left join fact on fact.code = file1_10.fact "
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.text)
cString = cString & " ORDER BY ID"
Set DATA10.Recordset = myRecordSet(cString, con)
myaddItem
Fixgrd
End Sub
Sub myProcAddModel()
If Not MYVALID(False) Then Exit Sub
If myreplace(True) Then
    Inform " „  «÷«›… «·„ÊœÌ·«  »‰Ã«Õ"
End If

myUndo

'myloadgrd
grid1.ShowCell grid1.Rows - 1, 2 + 1
grid1.Select grid1.Rows - 1, 2 + 1
End Sub

Private Sub xIgNet_Click()

End Sub

Private Sub xStore2_Validate(Cancel As Boolean)
    If xStore2.BoundText = xStore1.BoundText Then
        xStore2.BoundText = ""
        Cancel = True
    End If
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow Then
    If OldRow <> .Rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
        If Not validRow(OldRow) Then
            .RemoveItem OldRow
            CalcTotals
        End If
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If Not validRow(grid1.Row) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
    CalcTotals
End If
End Sub
Private Function validRow(Row) As Boolean
With grid1
If Trim(.TextMatrix(Row, 0)) = "" Then Exit Function
'If Not validItem(.TextMatrix(Row, 0), con) Then Exit Function
'If grid1.TextMatrix(Row, 8) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub myaddItem()
grid1.Row = grid1.Row
grid1.Col = 8
grid1.Select grid1.Row, 8

grid1.AddItem ""

End Sub
Private Function myreplaceGrdModel() As Boolean
Dim cString As String, aitem As Variant
cString = " DELETE FROM FR1_60 FROM FR1_60 INNER JOIN FILE1_10 ON FR1_60.ITEM = FILE1_10.ITEM"
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.text)
cString = cString & turn(cString) & "FILE1_10.MODEL = " & MyParn(transModelFrm.XMODEL.text)
con.Execute cString, nDel
With transModelFrm.grid1
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                aitem = ItemFields(transModelFrm.grid2.TextMatrix(nRow, nCol), con)
                If Not IsEmpty(aitem) Then
                    cString = "Insert into FR1_60 (doc_no,item,cost,Quant)" & _
                               "Values(" & _
                               addstring(xDoc_No.text) & "," & _
                               addvalue(retFlag(aitem, "item")) & "," & _
                               Val(retFlag(aitem, "cost") & "") & "," & _
                               Val(.TextMatrix(nRow, nCol)) & _
                               ")"
                    con.Execute cString
                End If
            End If
        Next
    Next
End With
End Function
Private Sub Check3_Click()
    addSetting "Check31", Check3.Value, App.Path & "\other.txt"
End Sub
Private Function openCardTable(Optional pDoc_No As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset
cString = "SELECT TOP 1 * FROM FR1_60H"
If sDoc_no <> "" Then
    cString = cString & " WHERE DOC_NO = " & MyParn(sDoc_no)
Else
    If pSign = "=" Then
        If pDoc_No <> "" Then cWhere = "DOC_NO  " & pSign & addstring(pDoc_No)
    Else
        If pDoc_No <> "" And sDoc_no = "" Then cWhere = "DOC_NO  " & pSign & addstring(pDoc_No)
    End If
    'If sDoc_no <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "DOC_NO = " & MyParn(sDoc_no)
    
    ' «·«Œ Ì«—« 
    If cFilter <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & cFilter
    If cWhere <> "" Then cString = cString & " WHERE " & cWhere
    If pSign = "<" Or pSign = "<=" Then
        cString = cString & " order by doc_no desc"
    ElseIf pSign = ">=" Or pSign = ">" Then
        cString = cString & " order by doc_no ASC"
    End If
End If
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function
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
myError:
MsgBox Err.Description
Err.Clear
End Sub

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
