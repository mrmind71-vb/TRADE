VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form salesfrm
   BorderStyle     =	   1  'Fixed Single
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15375
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
   ScaleHeight     =   10095
   ScaleWidth      =   15375
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Caption         =   " «—ÌŒ «·„” ‰œ"
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
      Left            =   2160
      TabIndex        =   91
      Top             =   1935
      Width           =   2580
      Begin VB.TextBox xdate_open 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   675
         MaxLength       =   10
         TabIndex        =   93
         Top             =   225
         Width           =   1770
      End
      Begin VB.CommandButton DAYALL 
         Height          =   450
         Left            =   45
         Picture         =   "Sales.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   180
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "—’Ìœ «·’‰›"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4770
      TabIndex        =   54
      Top             =   1935
      Width           =   1365
      Begin VB.Label xBalance 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Left            =   45
         TabIndex        =   55
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Frame Frame14 
      Height          =   555
      Left            =   2160
      TabIndex        =   187
      Top             =   2565
      Width           =   3975
      Begin VB.CommandButton Command5 
         Caption         =   "Send Receipt"
         Height          =   375
         Left            =   90
         TabIndex        =   188
         Top             =   135
         Width           =   1725
      End
      Begin VB.Label xUUID_RC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   189
         Top             =   135
         Width           =   2070
      End
   End
   Begin VB.CheckBox XISNODEL 
      Caption         =   "„— Ã⁄ ⁄—÷"
      Enabled         =   0   'False
      Height          =   330
      Left            =   15255
      TabIndex        =   179
      Top             =   8370
      Width           =   1230
   End
   Begin VB.Frame Frame13 
      Caption         =   "»ÕÀ ›Ï «·„” ‰œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   15255
      RightToLeft     =   -1  'True
      TabIndex        =   165
      Top             =   8775
      Width           =   6090
      Begin VB.TextBox xfilter4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   169
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   585
         Width           =   2355
      End
      Begin VB.TextBox xfilter3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   168
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   585
         Width           =   1770
      End
      Begin VB.TextBox xfilter2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   167
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   180
         Width           =   1770
      End
      Begin VB.TextBox xfilter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
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
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   166
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   180
         Width           =   1410
      End
      Begin VB.Label Label33 
         Caption         =   "«”„ «·’‰› "
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
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   173
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label32 
         Caption         =   "„’‰⁄"
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
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   172
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label31 
         Caption         =   "—ﬁ„ „ÊœÌ· "
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
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   171
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label8 
         Caption         =   "»«—ﬂÊœ"
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
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   170
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.Frame Frame12 
      Height          =   870
      Left            =   9540
      TabIndex        =   160
      Top             =   9855
      Width           =   5685
      Begin VB.CommandButton CND_PRICE3 
         Caption         =   "SALES PRICE"
         Height          =   690
         Left            =   4635
         TabIndex        =   174
         Top             =   135
         Width           =   1005
      End
      Begin VB.CommandButton Command4 
         Caption         =   "PRICE 2021"
         Height          =   330
         Left            =   3600
         TabIndex        =   164
         Top             =   495
         Width           =   1005
      End
      Begin VB.CommandButton Command3 
         Caption         =   "PRICE"
         Height          =   330
         Left            =   3600
         TabIndex        =   163
         Top             =   135
         Width           =   1005
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·≈Ã„«·Ï »«·„” Â·ﬂ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2730
         RightToLeft     =   -1  'True
         TabIndex        =   162
         Top             =   180
         Width           =   1050
         WordWrap        =   -1  'True
      End
      Begin VB.Label XTOTAL2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000040&
         Height          =   555
         Left            =   90
         TabIndex        =   161
         Top             =   180
         Width           =   2535
      End
   End
   Begin VB.CommandButton CMD_PAST_PURCH 
      BackColor       =   &H0080FF80&
      Caption         =   "‰”Œ „‰ ›« Ê—… „‘ —Ì« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   15300
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   156
      Top             =   765
      Width           =   915
   End
   Begin Threed.SSCommand CMD_ONLINE 
      Height          =   480
      Left            =   45
      TabIndex        =   153
      Top             =   1395
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   847
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      BackColor       =   8454143
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "»Ê‰ «Ê‰ ·«Ì‰  ”·Ì„"
      ButtonStyle     =   2
   End
   Begin VB.CommandButton FIX_VISA 
      Caption         =   "FIX VISA"
      Height          =   510
      Left            =   225
      TabIndex        =   146
      Top             =   6615
      Width           =   555
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   135
      Top             =   1980
   End
   Begin VB.CommandButton sned_purch 
      BackColor       =   &H0080FF80&
      Caption         =   " ÕÊÌ· ·„‘ —Ì«   ÊﬂÌ·« "
      Height          =   375
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   137
      Top             =   2430
      Width           =   1905
   End
   Begin VB.CommandButton cmd_delall 
      BackColor       =   &H000000C0&
      Caption         =   "Õ–› «·ﬂ·"
      Height          =   825
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   132
      Top             =   9090
      Width           =   465
   End
   Begin VB.CommandButton Command2 
      Caption         =   "xxxxxx"
      Height          =   690
      Left            =   2160
      TabIndex        =   131
      Top             =   1350
      Width           =   330
   End
   Begin VB.Frame Frame99 
      Caption         =   "—ﬁ„ «·„” ‰œ ··„— Ã⁄"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   225
      TabIndex        =   119
      Top             =   7245
      Width           =   3885
      Begin VB.CommandButton CMD_ISRET 
         BackColor       =   &H0080FF80&
         Caption         =   "”„«Õ »„— Ã⁄"
         Height          =   600
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   1125
         Width           =   645
      End
      Begin VB.CheckBox XISRET 
         Caption         =   "Check1"
         Height          =   195
         Left            =   1485
         TabIndex        =   148
         Top             =   1440
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.TextBox xdoc_ret 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1845
         MaxLength       =   12
         TabIndex        =   121
         Top             =   270
         Width           =   1950
      End
      Begin VB.CommandButton cmd_showdoc 
         Caption         =   "⁄—÷  ›’Ì·Ï «·»Ê‰"
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
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   120
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label XSALES_RET 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1845
         TabIndex        =   127
         Top             =   675
         Width           =   1950
      End
      Begin VB.Label xusername_RET 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1845
         TabIndex        =   126
         Top             =   1035
         Width           =   1950
      End
      Begin VB.Label XUSERNAME_DISC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   1845
         TabIndex        =   125
         Top             =   1395
         Width           =   1950
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "»Ê‰ «·„»Ì⁄« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   675
         TabIndex        =   124
         Top             =   720
         Width           =   960
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„— Ã⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   123
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·Œ’„"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   122
         Top             =   1395
         Width           =   1680
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Height          =   915
      Left            =   45
      TabIndex        =   111
      Top             =   0
      Width           =   1995
      Begin VB.OptionButton XONLINE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "«Ê‰ ·«Ì‰ web"
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
         Index           =   2
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   152
         Top             =   450
         Width           =   1725
      End
      Begin VB.OptionButton XONLINE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "«Ê‰ ·«Ì‰ ‰ﬁœÏ"
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
         Index           =   1
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   151
         Top             =   -180
         Visible         =   0   'False
         Width           =   1725
      End
      Begin VB.OptionButton XONLINE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "„»Ì⁄«  ‰ﬁœÏ"
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
         Index           =   0
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   150
         Top             =   135
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.Frame FRM_CLOSED 
         Height          =   1680
         Left            =   45
         TabIndex        =   115
         Top             =   3030
         Visible         =   0   'False
         Width           =   2040
         Begin Threed.SSCommand cmdAddItems 
            Height          =   825
            Left            =   90
            TabIndex        =   128
            Top             =   3330
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   1455
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
            Picture         =   "Sales.frx":24F2
            Caption         =   " ”ÃÌ· „ÊœÌ·«  "
            Alignment       =   8
            PictureAlignment=   6
         End
         Begin Threed.SSCommand cmd_CLOSEDDATE 
            CausesValidation=   0   'False
            Height          =   1005
            Left            =   1035
            TabIndex        =   116
            Top             =   1305
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   1773
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
            Picture         =   "Sales.frx":4CFA
            Caption         =   "≈€·«ﬁ › —…"
            Alignment       =   8
            PictureAlignment=   6
         End
         Begin Threed.SSCommand cmd_open 
            CausesValidation=   0   'False
            Height          =   1005
            Left            =   90
            TabIndex        =   117
            Top             =   1305
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   1773
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
            Picture         =   "Sales.frx":77CF
            Caption         =   "› Õ › —…"
            Alignment       =   8
            PictureAlignment=   6
         End
      End
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·„” ‰œ „€·ﬁ"
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
         Height          =   330
         Left            =   405
         TabIndex        =   114
         Top             =   1890
         Width           =   1320
      End
      Begin VB.Frame Frame_okaz 
         Caption         =   "Œ’„ «·√Êﬂ«“ÊÌ‰"
         Height          =   690
         Left            =   45
         TabIndex        =   112
         Top             =   3240
         Width           =   1995
         Begin VB.CheckBox xisokaz 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   " Ã«Â· ”⁄— «Êﬂ«“ÊÌ‰"
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
            Height          =   285
            Left            =   135
            TabIndex        =   113
            Top             =   315
            Width           =   1860
         End
      End
      Begin VB.Label xF_date 
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
         Left            =   135
         TabIndex        =   118
         Top             =   1890
         Width           =   1995
      End
   End
   Begin VB.Frame Frame10 
      Height          =   780
      Left            =   2025
      TabIndex        =   106
      Top             =   0
      Width           =   4695
      Begin Threed.SSCommand LOOKPH 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         TabIndex        =   107
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   1058
         _Version        =   196610
         CaptionStyle    =   1
         ForeColor       =   0
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
         Picture         =   "Sales.frx":A301
         Caption         =   "≈” ⁄·«„ ’Ê—"
         Alignment       =   1
         ButtonStyle     =   1
         PictureAlignment=   10
         BevelWidth      =   10
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdOpen 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   3420
         TabIndex        =   108
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   1058
         _Version        =   196610
         CaptionStyle    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«” ⁄·«„ »Ê‰«  „› ÊÕ…"
         PictureAlignment=   3
      End
      Begin Threed.SSCommand cmdinformbarcode 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   2295
         TabIndex        =   109
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1058
         _Version        =   196610
         CaptionStyle    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«” ⁄·«„ „»Ì⁄«  ’‰›"
         PictureAlignment=   3
      End
      Begin Threed.SSCommand LOOKPH_BARCODE 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   1350
         TabIndex        =   110
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1058
         _Version        =   196610
         CaptionStyle    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈” ⁄·«„ »«—ﬂÊœ"
         PictureAlignment=   3
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3390
      Left            =   45
      TabIndex        =   1
      Top             =   3150
      Width           =   15270
      _cx             =   26935
      _cy             =   5980
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
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   900
      Picture         =   "Sales.frx":C85F
      ScaleHeight     =   2400
      ScaleWidth      =   3165
      TabIndex        =   50
      Top             =   4095
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.CheckBox XGIFT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ÿ»«⁄… »Ê‰ ÂœÌ…"
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
      Height          =   510
      Left            =   6255
      TabIndex        =   98
      Top             =   855
      Width           =   1275
   End
   Begin VB.Frame Frame11 
      Height          =   690
      Left            =   2475
      TabIndex        =   96
      Top             =   720
      Width           =   3660
      Begin Threed.SSCommand xphone 
         Height          =   390
         Left            =   45
         TabIndex        =   99
         Top             =   180
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   688
         _Version        =   196610
         Font3D          =   3
         ForeColor       =   255
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmd_toexel 
         Height          =   435
         Left            =   2205
         TabIndex        =   158
         Top             =   180
         Width           =   1380
         _ExtentX        =   2434
         _ExtentY        =   767
         _Version        =   196610
         Font3D          =   3
         CaptionStyle    =   1
         ForeColor       =   128
         BackColor       =   14737632
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   " ÕÊÌ· ·≈ﬂ”Ì·"
         ButtonStyle     =   2
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   285
      Left            =   2385
      TabIndex        =   95
      Top             =   765
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.CheckBox chkprint 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "«·€«¡ «·ÿ»«⁄…"
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
      Left            =   6165
      TabIndex        =   90
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1365
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   4140
      Picture         =   "Sales.frx":FF36
      ScaleHeight     =   2445
      ScaleWidth      =   3210
      TabIndex        =   74
      Top             =   6615
      Width           =   3210
      Begin VB.CheckBox ISPICT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "⁄—÷ «·’Ê—"
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
         Left            =   45
         TabIndex        =   144
         Top             =   2070
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin ImgXCtrl6.ImgXCtrl imgx1 
         Height          =   2175
         Left            =   45
         TabIndex        =   103
         Top             =   45
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   3836
         BackColor       =   16777215
         BorderStyle     =   0
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   11340
      Picture         =   "Sales.frx":1360D
      ScaleHeight     =   3255
      ScaleWidth      =   3885
      TabIndex        =   64
      Top             =   6615
      Width           =   3885
      Begin VB.CommandButton cmd_disc 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1395
         TabIndex        =   101
         Top             =   1665
         Width           =   375
      End
      Begin VB.TextBox xRate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   12
         Top             =   1665
         Width           =   690
      End
      Begin VB.TextBox xDiscount 
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
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   270
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1665
         Width           =   1095
      End
      Begin VB.TextBox xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   26.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00292952&
         Height          =   540
         Left            =   225
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   13
         Top             =   2385
         Width           =   3435
      End
      Begin VB.Label xDOC_DISC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   2835
         TabIndex        =   175
         Top             =   2070
         Width           =   915
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "„— Ã⁄ :"
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   140
         Top             =   90
         Width           =   585
      End
      Begin VB.Label xTotalQuant2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   139
         Top             =   45
         Width           =   780
      End
      Begin VB.Label xNOTE_DISC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   270
         TabIndex        =   138
         Top             =   2070
         Width           =   2535
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ ﬁÿ⁄"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   90
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï «·ﬁÌ„…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   540
         Width           =   1005
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ √’‰«›"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   945
         Width           =   960
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï ›« Ê—…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   1305
         Width           =   1035
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ ›« Ê—…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   1710
         Width           =   900
      End
      Begin VB.Label xTotalQuant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1800
         TabIndex        =   68
         Top             =   45
         Width           =   690
      End
      Begin VB.Label xTotalItemNoDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   270
         TabIndex        =   67
         Top             =   450
         Width           =   2220
      End
      Begin VB.Label xTotalItemDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   270
         TabIndex        =   66
         Top             =   855
         Width           =   2220
      End
      Begin VB.Label xTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   270
         TabIndex        =   65
         Top             =   1260
         Width           =   2220
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   10845
      TabIndex        =   58
      Top             =   0
      Width           =   4425
      Begin VB.CommandButton cmdExit 
         Height          =   600
         Left            =   45
         Picture         =   "Sales.frx":1730B
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton CmdDel 
         Appearance      =   0  'Flat
         CausesValidation=   0   'False
         Height          =   600
         Left            =   1125
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales.frx":19777
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   2205
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales.frx":1C011
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   3285
         Picture         =   "Sales.frx":1E5BD
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   6165
      TabIndex        =   51
      Top             =   1665
      Width           =   1455
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
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales.frx":20D90
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Õ›Ÿ"
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Sales.frx":230F3
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   630
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   7380
      Picture         =   "Sales.frx":2566C
      ScaleHeight     =   2355
      ScaleWidth      =   3885
      TabIndex        =   40
      Top             =   6615
      Width           =   3885
      Begin VB.CheckBox xisbankahly 
         Alignment       =   1  'Right Justify
         Caption         =   "»‰ﬂ «Â·Ï"
         Height          =   240
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   141
         Top             =   945
         Visible         =   0   'False
         Width           =   1005
      End
      Begin MSDataListLib.DataCombo xBoxVISA 
         Height          =   315
         Left            =   45
         TabIndex        =   105
         TabStop         =   0   'False
         Top             =   2250
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XCODEVISA 
         Height          =   315
         Left            =   45
         TabIndex        =   129
         TabStop         =   0   'False
         Top             =   540
         Visible         =   0   'False
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label XPOINT 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   90
         TabIndex        =   147
         Top             =   135
         Width           =   915
      End
      Begin VB.Label XVISA2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   45
         TabIndex        =   145
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label XTYPEVISA 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   45
         TabIndex        =   143
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label XVISADOC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   1080
         TabIndex        =   142
         Top             =   945
         Width           =   2625
      End
      Begin VB.Label xPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   1080
         TabIndex        =   49
         Top             =   1350
         Width           =   1455
      End
      Begin VB.Label xRest 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   1080
         TabIndex        =   48
         Top             =   1755
         Width           =   1455
      End
      Begin VB.Label xcash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   1080
         TabIndex        =   47
         Top             =   135
         Width           =   1455
      End
      Begin VB.Label xlate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   1080
         TabIndex        =   46
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label xvisa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00000040&
         Height          =   375
         Left            =   1350
         TabIndex        =   45
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ ‰ﬁœÏ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   180
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ ›Ì“«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   585
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„œ›Ê⁄ ‰ﬁœÌ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1395
         Width           =   1035
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»«ﬁÌ «·‰ﬁœÌ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1800
         Width           =   990
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   495
      Picture         =   "Sales.frx":28892
      ScaleHeight     =   780
      ScaleWidth      =   10770
      TabIndex        =   31
      Top             =   9090
      Width           =   10770
      Begin VB.Frame Frame8 
         Height          =   600
         Left            =   5850
         TabIndex        =   133
         Top             =   45
         Width           =   4650
         Begin VB.CommandButton CMD_ADDITEM 
            Caption         =   "≈÷«›… »«—ﬂÊœ „‰ «·”Ì—›—"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   135
            Style           =   1  'Graphical
            TabIndex        =   135
            Top             =   135
            Visible         =   0   'False
            Width           =   2130
         End
         Begin VB.TextBox XBARCODEADD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2610
            MaxLength       =   12
            TabIndex        =   134
            Top             =   180
            Visible         =   0   'False
            Width           =   1950
         End
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         Caption         =   "«·»Ì⁄ »”⁄— «·√Êﬂ«“ÊÌ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   420
         Left            =   1665
         TabIndex        =   104
         Top             =   135
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   225
         TabIndex        =   39
         Top             =   225
         Width           =   315
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ–› ’‰›"
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
         Left            =   630
         TabIndex        =   38
         Top             =   225
         Width           =   825
      End
      Begin VB.Shape Shape9 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   90
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   1500
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   480
         Left            =   1845
         TabIndex        =   37
         Top             =   90
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "1-"
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
         Left            =   2160
         TabIndex        =   36
         Top             =   225
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Shape Shape8 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   1665
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   2835
         TabIndex        =   35
         Top             =   180
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "1+"
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
         Left            =   3150
         TabIndex        =   34
         Top             =   225
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.Shape Shape7 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   2655
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "F5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Left            =   4455
         TabIndex        =   33
         Top             =   225
         Width           =   225
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "”œ«œ ›« Ê—…"
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
         Left            =   4725
         TabIndex        =   32
         Top             =   225
         Width           =   870
      End
      Begin VB.Shape Shape6 
         BackStyle       =   1  'Opaque
         Height          =   510
         Left            =   4410
         Shape           =   4  'Rounded Rectangle
         Top             =   90
         Width           =   1365
      End
   End
   Begin VB.Frame Frame5 
      Height          =   780
      Left            =   6750
      TabIndex        =   29
      Top             =   0
      Width           =   4065
      Begin Threed.SSCommand cmdPrint 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   3150
         TabIndex        =   30
         Top             =   135
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1058
         _Version        =   196610
         CaptionStyle    =   1
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
         Picture         =   "Sales.frx":2BD2A
         Alignment       =   1
      End
      Begin Threed.SSCommand SSCommand2 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   1020
         TabIndex        =   63
         Top             =   135
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1058
         _Version        =   196610
         CaptionStyle    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "ÿ»«⁄… ≈Ã„«·Ï «·ÌÊ„"
         PictureAlignment=   3
      End
      Begin Threed.SSCommand cmdPrintDay 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   2115
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   135
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   1058
         _Version        =   196610
         CaptionStyle    =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "≈€·«ﬁ Ê €ÌÌ— «·ÌÊ„"
         PictureAlignment=   3
      End
      Begin Threed.SSCommand cmd_tsal 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   45
         TabIndex        =   94
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1058
         _Version        =   196610
         CaptionStyle    =   1
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
         Picture         =   "Sales.frx":2E164
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      Left            =   7605
      TabIndex        =   16
      Top             =   720
      Width           =   7710
      Begin VB.TextBox xcard_disc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5040
         MaxLength       =   200
         TabIndex        =   183
         TabStop         =   0   'False
         Top             =   1980
         Width           =   1500
      End
      Begin VB.CheckBox XISONEST 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   7425
         TabIndex        =   159
         Top             =   540
         Width           =   195
      End
      Begin VB.CheckBox XISRETS 
         Caption         =   "Check1"
         Height          =   285
         Left            =   4095
         TabIndex        =   157
         Top             =   180
         Width           =   195
      End
      Begin VB.TextBox xinv_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2970
         MaxLength       =   6
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   135
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdClient2 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   2970
         TabIndex        =   75
         Top             =   900
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox xDoc_no2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5355
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         IMEMode         =   3  'DISABLE
         Left            =   5355
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   2460
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3015
         MaxLength       =   200
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1620
         Width           =   3525
      End
      Begin VB.CheckBox chkCash 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "»Ì⁄ ‰ﬁœÌ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5355
         MaxLength       =   10
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   530
         Width           =   1185
      End
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         MaxLength       =   10
         TabIndex        =   4
         Top             =   180
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   540
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   900
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   90
         TabIndex        =   9
         Top             =   1980
         Width           =   105
         _ExtentX        =   185
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2970
         MaxLength       =   12
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   540
         Visible         =   0   'False
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xMan 
         Height          =   330
         Left            =   2970
         TabIndex        =   0
         Top             =   1260
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
      Begin VB.Label xcard_doc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   2115
         TabIndex        =   186
         Top             =   1980
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label xcard_value 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   135
         TabIndex        =   185
         Top             =   1980
         Width           =   1980
      End
      Begin VB.Label xcard_desca 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   2340
         TabIndex        =   184
         Top             =   1980
         Width           =   2655
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "ﬂ«—  Œ’„"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   182
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label XBRANCH 
         Caption         =   "Label14"
         Height          =   195
         Left            =   2295
         TabIndex        =   180
         Top             =   1350
         Width           =   420
      End
      Begin VB.Label xIpName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   176
         Top             =   1260
         Width           =   2025
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "„” ‰œ  ÕÊÌ· „‰ «·›—⁄ "
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5580
         RightToLeft     =   -1  'True
         TabIndex        =   136
         Top             =   810
         Width           =   1800
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "«·»«∆⁄ :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   1170
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "⁄„·«¡ ﬂ—Ê  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   2430
         Visible         =   0   'False
         Width           =   900
      End
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
         Left            =   3330
         TabIndex        =   56
         Top             =   900
         Width           =   1995
      End
      Begin VB.Label xtime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Left            =   90
         TabIndex        =   14
         Top             =   180
         Width           =   780
      End
      Begin VB.Label xusername 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   90
         TabIndex        =   26
         Top             =   1620
         Width           =   2025
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„” Œœ„ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1665
         Width           =   705
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1665
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰… :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   990
         Width           =   495
      End
      Begin VB.Label xCodeDesca 
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
         Left            =   2970
         TabIndex        =   21
         Top             =   540
         Width           =   2355
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì· :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   585
         Width           =   510
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -3240
      Top             =   6660
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   -855
      Top             =   6885
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
   Begin Crystal.CrystalReport REPORT1 
      Left            =   630
      Top             =   1350
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowProgressCtls=   0   'False
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -4770
      Top             =   8145
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Left            =   -3060
      Top             =   6975
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Left            =   -1755
      Top             =   6885
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin VB.CheckBox xPrinted 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   7380
      TabIndex        =   27
      Top             =   855
      Visible         =   0   'False
      Width           =   195
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   22
      Top             =   9945
      Visible         =   0   'False
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   2115
      Top             =   3465
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
   Begin Crystal.CrystalReport report2 
      Left            =   2115
      Top             =   900
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport Report_T 
      Left            =   3015
      Top             =   3195
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   810
      TabIndex        =   77
      Top             =   6570
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   78
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
         Picture         =   "Sales.frx":30D1B
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Sales.frx":32EEB
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   79
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
         Picture         =   "Sales.frx":35033
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Sales.frx":371FB
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   80
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
         Picture         =   "Sales.frx":3934A
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Sales.frx":3B52A
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   81
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
         Picture         =   "Sales.frx":3D685
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "Sales.frx":3F841
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   7740
      Picture         =   "Sales.frx":41990
      ScaleHeight     =   1140
      ScaleWidth      =   4470
      TabIndex        =   82
      Top             =   810
      Visible         =   0   'False
      Width           =   4470
      Begin VB.Label xCard_Sales2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   2880
         TabIndex        =   89
         Top             =   540
         Width           =   1365
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "√Œ—  «—ÌŒ »Ì⁄"
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
         Left            =   585
         TabIndex        =   88
         Top             =   45
         Width           =   1365
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì Œ’„"
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
         Left            =   1620
         TabIndex        =   87
         Top             =   45
         Width           =   1185
      End
      Begin VB.Label xcard_Sales 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   2925
         TabIndex        =   86
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label xcard_discount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   1665
         TabIndex        =   85
         Top             =   540
         Width           =   1050
      End
      Begin VB.Label xcard_date 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   180
         TabIndex        =   84
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì „»Ì⁄« "
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
         Left            =   2880
         TabIndex        =   83
         Top             =   45
         Width           =   1410
      End
   End
   Begin Threed.SSCommand cmd_charg 
      CausesValidation=   0   'False
      Height          =   525
      Left            =   2475
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   1395
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   926
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   0
      BackColor       =   16761024
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Traditional Arabic"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "⁄—÷ «·⁄„·«¡"
      PictureAlignment=   1
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   1170
      Top             =   3060
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      RTSEnable       =   -1  'True
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc data8 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin Threed.SSCommand cmd_addexel 
      Height          =   435
      Left            =   4725
      TabIndex        =   154
      Top             =   1440
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      CaptionStyle    =   1
      ForeColor       =   128
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«÷«›… „‰ «ﬂ”Ì·"
      ButtonStyle     =   2
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand cmd_closed 
      CausesValidation=   0   'False
      Height          =   465
      Left            =   135
      TabIndex        =   155
      Top             =   1935
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   820
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
      Picture         =   "Sales.frx":4527C
      Alignment       =   4
      PictureAlignment=   9
   End
   Begin VB.CommandButton FIX_COST 
      Caption         =   "FIX COST"
      Height          =   330
      Left            =   135
      TabIndex        =   130
      Top             =   -90
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Left            =   0
      TabIndex        =   181
      Top             =   0
      Width           =   780
   End
   Begin VB.Label XSHIP_NO 
      Caption         =   "Label14"
      Height          =   825
      Left            =   0
      TabIndex        =   178
      Top             =   7470
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label XONLINE_DOC 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Label14"
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
      Left            =   135
      TabIndex        =   177
      Top             =   990
      Width           =   1815
   End
   Begin VB.Label DATE0 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   " «—ÌŒ «·ÃÂ«“ „Œ ·› ⁄‰  «—ÌŒ «·„»Ì⁄« "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   45
      TabIndex        =   97
      Top             =   1935
      Visible         =   0   'False
      Width           =   2085
   End
End
Attribute VB_Name = "salesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temptable As New ADODB.Recordset, cMsgDisplay As String, nDayDiscout As Double, cDescDiscout As String
Dim lManSales As Boolean, nCountPrint As Double
Public sDoc_no As String, lSave As Boolean, sUserDisc As String, lIsPrice_2 As Boolean, lIsdisc2   As Boolean, lDiscModelRet As Boolean
Dim cFilter As String, lIsDocRet As Boolean, cFilterLook As String
Dim tCard As ADODB.Recordset, lDiscOpenDay As Boolean
Dim nLeftMargin, nCardHeight, nPageWidth, nRightMargin, cHeadBoon As String
Dim nRound As Long, clist1 As String, lIsBoxOnline As Boolean
Dim CardTable As ADODB.Recordset, cId_NoDesca1 As String, cId_NoDesca2 As String
Public bRetvalue As Boolean, cIpName As String
Dim cDefBox As String, cDefClient As String, cDefClientDesca As String, cDefCasher As String, cDefStore As String, cdefman As String
Dim dDiscBranch1 As Date, dDiscBranch2 As Date, nDiscBranch As Double
Dim oSearchItem As New Search3, oSearchDoc   As New Search3, oSearchCl1 As New Search3, bMarket As Boolean
Dim oSearchCl2 As New Search3, oSearchCl3 As New Search3
Dim bedit As Boolean, lISNOGRDISC As Boolean
Dim cFile As String, cCodeDesca As String
Dim con As New ADODB.Connection
Dim formMode, dDateLast As String
Public myPublic As Integer, lNotBranch As Boolean
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
    ItemsLookupAll Me, oSearchItem
End Sub
Private Function myreplace(Optional nRow As Long = -1, Optional bModel As Boolean) As Boolean
Dim pBoxDoc As String
If IsPrinted Then Exit Function
Dim aInsert(38, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.text)

aInsert(1, 0) = "code"
aInsert(1, 1) = addstring(xCode.text)

aInsert(2, 0) = "[Date]"
aInsert(2, 1) = addDate(xDate.text)

aInsert(3, 0) = "store"
aInsert(3, 1) = addstring(xStore.BoundText)

aInsert(4, 0) = "[Notes]"
aInsert(4, 1) = addstring(xNotes.text)

aInsert(5, 0) = "Rate"
aInsert(5, 1) = Val(xRate.text)

aInsert(6, 0) = "Discount"
aInsert(6, 1) = Val(xDiscount.text)

aInsert(7, 0) = "Cash"
aInsert(7, 1) = Val(xcash.Caption)

aInsert(8, 0) = "Box"
aInsert(8, 1) = addstring(xBox.BoundText)

aInsert(9, 0) = "Pay"
aInsert(9, 1) = Val(xPay.Caption)

'aInsert(10, 0) = "visa"
'aInsert(10, 1) = Val(xvisa.Caption)

aInsert(11, 0) = "late"
aInsert(11, 1) = Val(xlate.Caption)

aInsert(12, 0) = "rest"
aInsert(12, 1) = Val(xRest.Caption)

aInsert(13, 0) = "userName"
aInsert(13, 1) = addstring(xusername.Caption)

aInsert(14, 0) = "branch"
aInsert(14, 1) = addstring(XBRANCH.Caption)

If cBranch > "60" Then
    If xMan.BoundText = "" Then xMan.BoundText = "0001"
Else
    If xMan.BoundText = "" Then xMan.BoundText = "00"
End If
If xMan.BoundText = "" Then
    If cBranch > "60" Then
        xMan.BoundText = "00"
    Else
        xMan.BoundText = "0001"
    End If
End If
aInsert(15, 0) = "MAN"
aInsert(15, 1) = addstring(xMan.BoundText)

aInsert(16, 0) = "MOSM"
aInsert(16, 1) = addstring(xMosm.BoundText)

aInsert(17, 0) = "[TIME]"
aInsert(17, 1) = addTime(xtime.Caption)

aInsert(18, 0) = "PHONE"
aInsert(18, 1) = addstring(xphone.Caption)

aInsert(19, 0) = "isnew"
aInsert(19, 1) = 1

aInsert(20, 0) = "username_disc"
aInsert(20, 1) = addstring(XUSERNAME_DISC.Caption)

aInsert(21, 0) = "username_ret"
aInsert(21, 1) = addstring(xusername_RET.Caption)

aInsert(22, 0) = "SALES_RET"
aInsert(22, 1) = addstring(XSALES_RET.Caption)

aInsert(23, 0) = "CODEVISA"
aInsert(23, 1) = addvalue(XCODEVISA.BoundText)

aInsert(24, 0) = "NOTE_DISC"
aInsert(24, 1) = addstring(xNOTE_DISC.Caption)

aInsert(32, 0) = "DOC_DISC"
aInsert(32, 1) = addstring(xDOC_DISC.Caption)

aInsert(25, 0) = "isbankahly"
aInsert(25, 1) = IIf(xisbankahly.Value, 1, 0)

aInsert(26, 0) = "VISADOC"
aInsert(26, 1) = addstring(XVISADOC.Caption)

aInsert(27, 0) = "VISA2"
aInsert(27, 1) = Val(XVISA2.Caption)

aInsert(28, 0) = "TYPEVISA"
aInsert(28, 1) = addvalue(XTYPEVISA.Caption)

aInsert(29, 0) = "POINT"
aInsert(29, 1) = Val(XPOINT.Caption)

aInsert(30, 0) = "ONLINE"
If XONLINE(0).Value Then aInsert(30, 1) = 0
If XONLINE(1).Value Then aInsert(30, 1) = 1
If XONLINE(2).Value Then aInsert(30, 1) = 2

aInsert(31, 0) = "inv_no"
aInsert(31, 1) = addstring(xinv_no.text)

aInsert(33, 0) = "USER_IP"
aInsert(33, 1) = addstring(xIpName.Caption)

aInsert(34, 0) = "ISNODEL"
aInsert(34, 1) = IIf(XISNODEL.Value, 1, 0)


aInsert(35, 0) = "card_disc"
aInsert(35, 1) = addstring(xcard_disc.text)

aInsert(36, 0) = "card_desca"
aInsert(36, 1) = addstring(xcard_desca.Caption)

aInsert(37, 0) = "card_VALUE"
aInsert(37, 1) = addstring(xcard_value.Caption)

aInsert(38, 0) = "card_DOC"
aInsert(38, 1) = addstring(xcard_doc.Caption)


'xcard_disc.text = CardTable!card_disc & ""
'xcard_desca.Caption = CardTable!card_desca & ""
'xcard_value.Caption = CardTable!card_VALUE & ""
'xcard_doc.Caption = CardTable!card_DOC & ""

con.BeginTrans
'On Error GoTo myerror
If cBranch > "60" Then
    pBoxDoc = cBranchBox
Else
    pBoxDoc = cBranchBox
End If
If xDoc_No.Tag = DefineMode Then
    If cBranch <> "00" Then
        If lIsBranchStore Then
            dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES WHERE BRANCH = " & MyParn(cBranch), con), "DD-MM-YYYY")
        Else
            dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES", con), "DD-MM-YYYY")
        End If
        If Format(dSalesDate, "DD-MM-YYYY") <> Format(xDate.text, "DD-MM-YYYY") Then
            MsgBox " „  €Ì— «· «—ÌŒ"
        End If
        xDate.text = Format(dSalesDate, "DD-MM-YYYY")
        aInsert(2, 1) = addDate(xDate.text)
    End If
    xDoc_No.text = NewflagDoc(xDate.text, pBoxDoc, con)
    xDoc_no2.text = Mid(xDoc_No.text, 7, 6)
    aInsert(0, 1) = addstring(xDoc_No.text)
    con.Execute CreateInsert(aInsert, "FILE6_20H")
Else
    aInsert(7, 0) = "Cash"
    aInsert(7, 1) = Val(xcash.Caption)
    con.Execute CreateUpdate(aInsert, "FILE6_20H", " where doc_no = " & addstring(xDoc_No.text))
End If

If bModel Then
    myreplaceGrdModel
Else
    myreplaceGrd nRow
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
xDoc_No.Tag = 1
End Function
Sub myProc()
'On Error GoTo myerror
If ActiveControl.Name = grid1.Name Or ActiveControl.Name = LOOKPH.Name Or ActiveControl.Name = LOOKPH_BARCODE.Name Then
    Dim bNew As Boolean
    bNew = grid1.Row = grid1.Rows - 1
    If ActiveControl.Name = grid1.Name Then
        grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    ElseIf ActiveControl.Name = LOOKPH.Name Then
        grid1.TextMatrix(grid1.Row, 1) = ItemLookPh_pict.grid1.TextMatrix(ItemLookPh_pict.grid1.Row, 0)
    ElseIf ActiveControl.Name = LOOKPH_BARCODE.Name Then
        grid1.TextMatrix(grid1.Row, 1) = ItemLookPh.grid1.TextMatrix(ItemLookPh.grid1.Row, 0)
    End If
    If Not bNew Then
        If ActiveControl.Name = grid1.Name Then Unload oSearchItem Else Unload ItemLookPh
        CellPos 13, grid1.Row, grid1.Col
    ElseIf ActiveControl.Name = LOOKPH.Name Then
        grid1.SetFocus
        grid1.TextMatrix(grid1.Row, 1) = ItemLookPh_pict.GRID2.TextMatrix(ItemLookPh_pict.GRID2.Row, 0)
        Unload ItemLookPh_pict
    ElseIf ActiveControl.Name = LOOKPH_BARCODE.Name Then
        grid1.SetFocus
        grid1.Select grid1.Rows - 1, 1
        Unload ItemLookPh
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    grid1_AfterEdit grid1.Row, grid1.Col
'If ActiveControl.Name = grid1.Name Or ActiveControl.Name = LOOKPH.Name Then
'    Dim bNew As Boolean
'    bNew = grid1.Row = grid1.Rows - 1
'    If ActiveControl.Name = grid1.Name Then
'        grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
'    Else
'        grid1.TextMatrix(grid1.Row, 1) = ItemLookPh.grid1.TextMatrix(ItemLookPh.grid1.Row, 0)
'    End If
'    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
'    grid1_AfterEdit grid1.Row, grid1.Col
ElseIf ActiveControl.Name = CmdInform.Name Or ActiveControl.Name = cmdOpen.Name Then
    xDoc_No.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
'    xDate.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 2)
    Unload oSearchDoc
    myUndo
ElseIf ActiveControl.Name = xCode.Name Then
    ActiveControl.text = oSearchCl1.grid1.TextMatrix(oSearchCl1.grid1.Row, 0)
    xCode_LostFocus
    Unload oSearchCl1

End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub chkprint_Click()
    addSetting "print", chkprint.Value, App.Path & "\other.txt"
End Sub
Private Sub CMD_ISRET_Click()
    If MsgBox("«·”„«Õ »„— Ã⁄ ··›« Ê—…", vbYesNo + vbDefaultButton2) = vbYes Then
        con.Execute " UPDATE FILE6_20H SET ISRET = 1 , ISNEW = 1 WHERE DOC_NO = " & MyParn(xDoc_No.text), nRec
        If nRec = 1 Then Inform "  „  ⁄œÌ· «·„” ‰œ "
    End If
End Sub

Private Sub CMD_ONLINE_Click()
On Error GoTo myerror
    lSave = False
    If grid1.Rows = 2 Then Exit Sub

    con.Execute " UPDATE FILE6_20H SET DISCOUNT = " & Val(xDiscount.text) & "  , RATE = " & Val(xRate.text) & " , username_disc  = " & addstring(XUSERNAME_DISC.Caption) & " , DOC_disc  = " & addstring(xDOC_DISC.Caption) & " , ISNEW =  1  WHERE DOC_NO = " & MyParn(xDoc_No.text)
    Set CashOnLine.myForm = Me
    CashOnLine.Show 1
    
    
    
    myUndo

    If XONLINE_DOC.Caption <> "" Then
        Dim con_SERVER As New Connection
        If openCon(con_SERVER, LoadConString_OnLine) = "ok" Then
            con_SERVER.Execute " UPDATE FILE6_90H SET SHIP_NO = " & addstring(XSHIP_NO.Caption) & " WHERE doc_no = " & MyParn(XONLINE_DOC.Caption), nRec
        End If
        con_SERVER.Close
    End If

'    If Not lSave Then Exit Sub
'
'    If doprint(False) Then SavePrint
'    DoEvents
'    CmdNewInv_Click
'    On Error Resume Next
'    XMAN.SetFocus
'    Err.Clear
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CMD_PAST_PURCH_Click()
    If cDocCopySales = "" Then Exit Sub
    If xStore.BoundText = "" Then
        MsgBox " ≈Œ Ì«— «·„Œ“‰ "
        Exit Sub
    End If
    If xCode.text = "" Then
        MsgBox " ≈Œ Ì«— «·⁄„Ì· "
        Exit Sub
    End If
    If Not myreplace Then Exit Sub
    con.Execute " INSERT INTO FILE6_20 (doc_no , MAN,ITEM, QUANT, PRICE, ROW ) SELECT " & addstring(xDoc_No.text) & "  , '000' ,FILE7_20.ITEM, FILE7_20.QUANT, FILE1_10.COST2 , FILE7_20.ROW  FROM  FILE7_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE7_20.ITEM WHERE DOC_NO = " & MyParn(cDocCopySales)
    Inform "  „ «÷«›… «·«’‰«› "
    myUndo
End Sub

Private Sub cmd_toexel_Click()
    grid1.ColHidden(18) = False
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
    grid1.ColHidden(18) = True
End Sub

Private Sub CND_PRICE3_Click()
    If InputBox("", "") = "987654321" Then
        con.Execute " UPDATE FILE6_20 SET  file6_20.price = file1_10.PRICE , file6_20.price_C = file1_10.PRICE  FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM where DOC_NO = " & MyParn(xDoc_No.text)
        myloadgrd
    End If

End Sub

Private Sub Command3_Click()
    If InputBox("", "") = "987654321" Then
        con.Execute " UPDATE FILE6_20 SET  file6_20.price_C = file1_10.PRICE  FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM where DOC_NO = " & MyParn(xDoc_No.text)
        myloadgrd
    End If
End Sub

Private Sub Command4_Click()
    If InputBox("", "") = "987654321" Then
        con.Execute " UPDATE FILE6_20 SET  file6_20.price_C = file1_10.PRICE_2021   FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM where DOC_NO = " & MyParn(xDoc_No.text)
        con.Execute " UPDATE FILE6_20 SET  file6_20.price_C = file1_10.PRICE        FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM where ( file6_20.price_C =0 0 OR file6_20.price_C IS NULL ) AND DOC_NO = " & MyParn(xDoc_No.text)
        myloadgrd
    End If
End Sub

Private Sub Command5_Click()
Dim sError As String
Dim sType As String
sType = ReceiptType
If sType = "" Then
    MsgBox "«·„” ‰œ ·« ÌÕ ÊÌ ⁄·Ì ﬂ„Ì« "
    Exit Sub
ElseIf sType = "v" Then
    MsgBox "«·„” ‰œ ÌÕ ÊÌ ⁄·Ì ﬂ„Ì«  „ÊÃ»… Êﬂ„Ì«  ”«·»…"
    Exit Sub
End If


SubmitReceipt xDoc_No.text, con, sType, sError

If sError <> "" Then
    MsgBox sError
Else
    myUndo
End If
End Sub
Private Function ReceiptType() As String
Dim i As Long, nType As Integer
For i = 1 To grid1.Rows - 2
    If grid1.ValueMatrix(i, 10) = 0 Then
    ElseIf nType = 0 Then
        nType = IIf(grid1.ValueMatrix(i, 10) > 0, 1, 2)
    Else
        If nType <> IIf(grid1.ValueMatrix(i, 10) > 0, 1, 2) Then
           ReceiptType = "v"
          Exit Function
        End If
    End If
Next
If nType <> 0 Then ReceiptType = IIf(nType = 1, "s", "r")
End Function
Private Sub FIX_VISA_Click()
    con.Execute " update file6_20h set typevisa  = 0 from file6_20h where isbankahly = 0"
    con.Execute " update file6_20h set typevisa  = 1 from file6_20h where isbankahly = 1"
End Sub

Private Sub ISPICT_Click()
    addSetting "ISPICT", ISPICT.Value, App.Path & "\other.txt"
End Sub
Private Sub CMD_ADDITEM_Click()
'    On Error GoTo myerror
    Inform " CMD_ADDITEM "
''''    strConShop = LoadConStringshop
''''    nBranchPost = Val(cBranch)
''''    pBranch = cBranch
''''    If openConShop(conShop) = "ok" Then
''''        Inform " „ «·≈ ’«· »‰Ã«Õ »«·”Ì—›—"
''''        sendItems
''''    Else
''''        MsgBox "›‘· «·≈ ’«·"
''''    End If
''''    Exit Sub
''''myerror:
''''MsgBox Err.Description
''''Err.Clear
End Sub

Private Sub cmd_charg_Click()
    VsCustPhone.Show 1
End Sub
Private Sub cmd_charge_Click()
    chargefrm.bedit = True
    chargefrm.myPublic = 1
    chargefrm.Show 1
End Sub
Private Sub cmd_closed_Click()
If (cBranch = "00" And bopt2) Or bSupermode Then
    If CardTable!ISCLOSED Then
        AddLod_Data cusername, 0, " › Õ „” ‰œ ·· ⁄œÌ· ", con, xDoc_No.text, xDate.text, , xCodeDesca.Caption
        
        con.Execute " update file6_20h set isnew = 1  , isclosed = 0 , printed = 0 where doc_no = " & MyParn(xDoc_No.text)
    Else
        con.Execute " update file6_20h set isnew = 1 , isclosed = 1 , printed = 1 where doc_no = " & MyParn(xDoc_No.text)
    End If
    myUndo
End If
End Sub
Private Sub cmd_delall_Click()
''''    If GetDesca("select count(*) from file6_20h ", CON) > 500 Then Exit Sub
''''    If InputBox("Õ–› ﬂ· «·»Ì«‰« ", "") = "654321" Then
''''        If MsgBox("”Ê› Ì „ Õ–› ﬂ· „»Ì⁄«  «·›—⁄ ", vbYesNo + vbDefaultButton2) = vbYes Then
''''            CON.Execute " delete from file6_20  "
''''            CON.Execute " delete from file6_20h "
''''            CON.Execute " delete from SUBCUST   "
''''        End If
''''    End If
End Sub

Private Sub cmd_disc_Click_OLD()
If xPrinted.Value <> 0 Then Exit Sub
'If Not bOpt6 Then
    
    Dim cString As String
    cString = "Select Option6 from USERS"
    Managerfrm.sString = cString
    Managerfrm.sFlag = ActiveControl.Name
    Managerfrm.sFilter = "option6 = 1"
    Managerfrm.Show 1
    If Not IsEmpty(aUser) Then
        If Not aUser(0) Then
            MsgBox "·Ì” ·œÌﬂ ’·«ÕÌ… «·Œ’„ "
            Exit Sub
        Else
            MsgBox "Ì„ﬂ‰  ”ÃÌ· «·Œ’„ "
            xDiscount.Locked = False
            xRate.Locked = False
            XUSERNAME_DISC.Caption = sUserDisc
            xNOTE_DISC.Caption = ""
            xDOC_DISC.Caption = ""
        End If
    Else
        Exit Sub
    End If

'End If
End Sub

Private Sub cmd_move_Click()
    ModelMove.Show
End Sub
Private Sub cmd_newph_Click()
    PhoneCust.lAdd = True
    PhoneCust.Show
End Sub

Private Sub cmd_open_Click()
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "› Õ  „” ‰œ«  „‰  «—ÌŒ", ""), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "› Õ „” ‰œ«  Õ Ï  «—ÌŒ", ""), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ › Õ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update file6_20h set isclosed = 0 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ › Õ „” ‰œ«  «·› —…" & nRec
            myUndo
        End If
    End If
End Sub

Private Sub CMD_PRINT_Click()
    doprint_doc

End Sub
Private Sub cmd_ret_Click()
    'TDaySalRet.Show 1
End Sub
Private Sub cmd_printsal_Click()
'    Print_TimeItemSales
End Sub
Private Sub cmd_showdoc_Click_OLD()
        
    Dim pDocRet As String
    Dim dDateRet As Date
    pDocRet = xdoc_ret.text
    xdoc_ret.text = pDocRet
    If Len(xdoc_ret.text) >= 11 Then
        dDateRet = DateValue(Format(Mid(xdoc_ret.text, 1, 2) & "-" & Mid(xdoc_ret.text, 3, 2) & "-" & Mid(xdoc_ret.text, 5, 2), "DD-MM-YYYY"))
        If DateDiff("D", dDateRet, dSalesDate) > 30 Then
            MsgBox " »Ê‰ „»Ì⁄«   «—ÌŒ… «ﬂ»— „‰ 30 ÌÊ„"
            Exit Sub
        End If
        If Not lRetDoc Then
            Dim cString As String
            cString = "Select Option7 from USERS"
            Managerfrm.sString = cString
            Managerfrm.sFlag = ActiveControl.Name
            Managerfrm.sFilter = "option7 = 1"
            Managerfrm.Show 1
            If Not IsEmpty(aUser) Then
                If Not aUser(0) Then
                    MsgBox "·Ì” ·œÌﬂ ’·«ÕÌ… «·„— Ã⁄ "
                    Exit Sub
                Else
                    lIsDocRet = IIf(GetDesca("SELECT Option8 FROM USERS WHERE DESCA =  " & MyParn(sUserDisc), con) = "True", True, False)
                    lRetDoc = True
'                    MsgBox "Ì„ﬂ‰  ”ÃÌ· «·„— Ã⁄"
                    xusername_RET.Caption = sUserDisc
                End If
            Else
                Exit Sub
            End If
        End If
        SalesDoc.sDoc_no = xdoc_ret.text
        SalesDoc.Show 1
    End If
End Sub

Private Sub CMD_TRANS_Click()
    Transfrm.Show
End Sub
Private Sub cmd_tsal_Click()
'''''If Not bOpt4 Then
'''''    Dim cString As String
'''''    cString = "Select Option6 from USERS"
'''''    Managerfrm.sString = cString
'''''    Managerfrm.sFlag = ActiveControl.Name
'''''    Managerfrm.sFilter = "option4 = 1"
'''''    Managerfrm.Show 1
'''''    If Not IsEmpty(aUser) Then
'''''        If Not aUser(0) Then
'''''            MsgBox "·Ì” ·œÌﬂ ’·«ÕÌ… ⁄—÷ ≈Ã„«·Ï «·„»Ì⁄« "
'''''            Exit Sub
'''''        Else
'''''            TDaySal.Show 1
'''''        End If
'''''    Else
'''''        Exit Sub
'''''    End If
'''''Else
'''''    TDaySal.Show 1
'''''End If

    Dim cPassword As String, aUser As Variant
    ManagPass.pPassWord = ""
    ManagPass.Show 1
    cPassword = ManagPass.pPassWord
    If Trim(cPassword) = "" Then Exit Sub
    aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
    If UBound(aUser) = 0 Then Exit Sub
    TDaySal.Show 1

End Sub

Private Sub cmd_update_Click()
    transDatafrm.Show 1
End Sub
Private Sub cmdClient2_Click()
    Clients_sub.Show
End Sub
Private Sub CmdDel_Click()
Dim cPassword As String, aUser As Variant
    myUndo
    If IsPrinted And (cBranch <> "00") Then
        MsgBox " „  ÿ»«⁄… «·»Ê‰ „‰ ﬁ»·"
        Exit Sub
    End If
    If xPrinted.Value <> 0 And (cBranch <> "00") Then Exit Sub
    
    ManagPass.pPassWord = ""
    ManagPass.Show 1
    cPassword = ManagPass.pPassWord
    If Trim(cPassword) = "" Then Exit Sub
    aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
    If UBound(aUser) = 0 Then Exit Sub


If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    
    AddLod_Data cusername, 2, " Õ–› „»Ì⁄«  ", con, xDoc_No.text, xDate.text, , xCodeDesca.Caption
    
    
    con.BeginTrans
    
    cStr1 = " INSERT INTO COPY6_20 (DOC_NO, ITEM, QUANT, PRICE, TOTAL, DISCOUNT,  cost, PRICE_C, row, MAN, S_OKAZ, ISDISC2_1, DATE, DELUSER, ID2,DEL_IP) " & _
             " SELECT        DOC_NO, ITEM, QUANT, PRICE, TOTAL, DISCOUNT, cost, PRICE_C, row, MAN, S_OKAZ, ISDISC2_1, " & addDate(xDate.text) & " , " & addstring(aUser(2)) & ", ID  , " & addstring(xIpName.Caption) & _
            "   From FILE6_20 WHERE DOC_NO = " & addstring(xDoc_No.text)
    con.Execute cStr1
    
    ' Õ–› «·„” ‰œ
    con.Execute "Delete  From FILE6_20 where Doc_No = " & MyParn(xDoc_No.text)
    con.Execute "Delete  From FILE6_20H where Doc_No = " & MyParn(xDoc_No.text)

    If XONLINE_DOC.Caption <> "" Then
        Dim con_SERVER As New Connection
        If openCon(con_SERVER, LoadConString_OnLine) = "ok" Then
            con_SERVER.Execute " UPDATE FILE6_90H SET SHIP_NO  = NULL , SALES_DOC = NULL , SALES_DATE = NULL WHERE doc_no = " & MyParn(XONLINE_DOC.Caption), nRec
            If nRec = 1 Then Inform "  „ «·€«¡  —ÕÌ· «·ÿ·»  "
        End If
        con_SERVER.Close
    End If


'   con.Execute "update file6_20h set isdate = getdate() , discount = 0 , cash = 0 , visa = 0  where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    
'''''''    If XTRANS_DOC.Caption <> "" Then
'''''''        con.Execute " update fr1_61h set DOC_TRANS = null , TYPEDOC = null   ,ISNEW = 1    where doc_no = " & MyParn(XTRANS_DOC.Caption)
'''''''    End If
    
    myUndo
    CmdNewInv_Click
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
    sDoc_no = ""
    Unload Me
End Sub
Private Sub cmdGo_Click()
    VsITEMBAL.Show 1
End Sub
Private Sub CmdInform_Click()
    CardLookup
End Sub

Private Sub cmdinformbarcode_Click()
    grdsales.Show 1
End Sub

Private Sub CmdNewInv_Click()
xdate_open.text = ""
myDefine
On Error Resume Next
xMan.SetFocus
grid1.Select 1, 1
Err.Clear
End Sub
Private Sub cmdPrintDay_Click()

DoBackUp

If cBranch = "19" Then
    InvoiceTransfrm.bSend = True
    InvoiceTransfrm.Show 1
End If

Dim nBALBox As Double
If CheckOpen(True) > 0 Then Exit Sub
If lIsBranchStore Then
    nBALBox = Round(Val(GetDesca("SELECT  SUM(BOXMOVE.PLUS - BOXMOVE.MINUS) FROM  FILE0_50 INNER JOIN BOXMOVE ON FILE0_50.CODE = BOXMOVE.BOX WHERE FILE0_50.branch = " & MyParn(cBranch) & " and  (FILE0_50.TYPE = 1)", con) & ""), 0)
Else
    nBALBox = Round(Val(GetDesca("SELECT  SUM(BOXMOVE.PLUS - BOXMOVE.MINUS) FROM  FILE0_50 INNER JOIN BOXMOVE ON FILE0_50.CODE = BOXMOVE.BOX WHERE (FILE0_50.TYPE = 1)", con) & ""), 0)
End If
If nBALBox <> 0 Then
    MsgBox " —’Ìœ Œ“‰ «·ﬂ«‘Ì—  " & Format(nBALBox, "#0.00") & " »—Ã«¡  ÕÊÌ· «·‰ﬁœÌ… ﬁ»· «·√€·«ﬁ "
    Exit Sub
End If
If Not bOpt4 Then
    Dim cString As String
    cString = "Select Option6 from USERS"
    Managerfrm.sString = cString
    Managerfrm.sFlag = ActiveControl.Name
    Managerfrm.sFilter = "option4 = 1"
    Managerfrm.Show 1
    If Not IsEmpty(aUser) Then
        If Not aUser(0) Then
            MsgBox "·Ì” ·œÌﬂ ’·«ÕÌ… ⁄—÷ ≈Ã„«·Ï «·„»Ì⁄« "
            Exit Sub
        Else
            If MsgBox(" —ÕÌ· Ê ≈€·«ﬁ «·Õ—ﬂ«  Ê «· ”ÃÌ·« ", vbYesNo + vbDefaultButton2) = vbYes Then
                PostData
                If Not lIsBranchStore Then
                    If UCase(RetSetting("COPY", App.Path & "\COPY.txt")) <> "COPY" Then
                        strConShop = LoadConStringshop
                        If openConShop(conShop) <> "ok" Then
                            MsgBox "·« ÌÊÃœ « ’«· »«·”Ì—›—"
                        Else
                            Inform "  „ «·≈ ’«· »«·”Ì—›— "
                            If cBranch < "60" Then
                                
                                'DoBackUp
                                

                                SendSales
                                sendCharge
                                SendTransBox
                            Else
                                SendSales_fr
                            End If
                        End If
                    End If
                Else
'                    ClearTemp6_20
                End If
                CmdNewInv_Click
            End If
        End If
    Else
        Exit Sub
    End If
Else
    If MsgBox(" —ÕÌ· Ê ≈€·«ﬁ «·Õ—ﬂ«  Ê «· ”ÃÌ·« ", vbYesNo + vbDefaultButton2) = vbYes Then
        PostData
        If Not lIsBranchStore Then
            If UCase(RetSetting("COPY", App.Path & "\COPY.txt")) <> "COPY" Then
                strConShop = LoadConStringshop
                If openConShop(conShop) <> "ok" Then
                    MsgBox "·« ÌÊÃœ « ’«· »«·”Ì—›—"
                Else
                    Inform "  „ «·≈ ’«· »«·”Ì—›— "
                    If cBranch < "60" Then
'                        DoBackUp
                        SendSales
                        sendCharge
                        SendTransBox
                    Else
                        SendSales_fr
                    End If
                End If
            End If
        End If
        CmdNewInv_Click
    End If
End If

'nDayDiscout = IsDayDisc(dSalesDate)
'If nDayDiscout > 0 Then lDiscOpenDay = True
'If lDiscOpenDay Then
'    Label20.Caption = cDescDiscout
'    Label20.Visible = True
'Else
'    Label20.Caption = ""
'    Label20.Visible = False
'End If
End Sub
Private Sub cmdSave_Click()
On Error GoTo myerror
Dim aRet1 As Variant, cDoc1 As String
If grid1.Rows = 2 Then Exit Sub
If Not MYVALID Then Exit Sub
If XONLINE(2).Value = 0 Or (XSALES_RET.Caption <> "") Then

If cBranch = "00" Then
    AddLod_Data cusername, 1, " Õ÷Ÿ „»Ì⁄«  ", con, xDoc_No.text, xDate.text, , xCodeDesca.Caption
End If


If cBranch <> "00" And xcard_disc.text <> "" Then
    xDiscount.text = CheckCardDisc
    xDOC_DISC.Caption = ""
    xNOTE_DISC.Caption = ""
    CalcTotals
End If

If cBranch <> "00" And (XSALES_RET.Caption = "") And Not lISNOGRDISC And xcard_disc.text = "" Then
    If XUSERNAME_DISC.Caption = "" Then
        con.Execute " UPDATE FILE6_20 SET FILE6_20.PRICE = FILE6_20.Price_C2 , FILE6_20.S_OKAZ = FILE6_20.S_OKAZ_2 WHERE DOC_NO =  " & MyParn(xDoc_No.text)
        myloadgrd
        xDiscount.text = ""
        xRate.text = ""
        aRet = CalcDisc3()
        cDoc1 = aRet(2)
        If cDoc1 <> "" Then
            cStr1 = " SELECT COUNT(*) FROM  FILE1_10 INNER JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM LEFT OUTER JOIN FILE0_90 ON FILE1_10.modelno = FILE0_90.MODELNO  WHERE  (FILE0_90.MODELNO IS NULL) AND FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)
            If Val(GetDesca(cStr1, con) & "") = 0 Then
                aRet = CalcDisc3(cDoc1)
                xDiscount.text = Val(aRet(1))
                xNOTE_DISC.Caption = aRet(3)
                xDOC_DISC.Caption = aRet(2)
                Inform xNOTE_DISC.Caption
                xinv_no.text = cDoc1
            Else
                xNOTE_DISC.Caption = aRet(3)
                xDOC_DISC.Caption = aRet(2)
                Inform xNOTE_DISC.Caption
                MsgBox " ÌÊÃœ «’‰«› Œ«—Ã «·⁄—÷ ·« Ì„ﬂ‰ Õ›Ÿ «·⁄„·Ì… "
                For nRow = 1 To grid1.Rows - 2
                    If GetDesca("SELECT FILE1_10.ITEM FROM FILE1_10 INNER JOIN FILE0_90 ON FILE0_90.MODELNO = FILE1_10.MODELNO WHERE FILE0_90.DOC_NO = " & MyParn(cDoc1) & " AND FILE1_10.ITEM =  " & Val(grid1.TextMatrix(nRow, 1)), con) = "" Then
                        grid1.Cell(flexcpBackColor, nRow, 0, nRow, grid1.Cols - 1) = vbRed
                    End If
                Next nRow
                Exit Sub
            End If
        End If
        If Val(xDiscount.text) = 0 Then
            If CalcDiscRateDoc Then
                MsgBox " ÌÊÃœ Œ’„ ⁄—Ê÷  "
                myloadgrd
            End If
        End If
        
        If nDiscBranch > 0 And IsDate(dDiscBranch1) And IsDate(dDiscBranch1) And xDOC_DISC.Caption = "" Then
            If dSalesDate >= dDiscBranch1 And dSalesDate <= dDiscBranch2 Then
                nRate = ((100 - Val(nDiscBranch & "")) / 100)
                con.Execute " UPDATE FILE6_20 SET FILE6_20.S_OKAZ =  " & Val(nDiscBranch) & " , FILE6_20.PRICE = FILE1_10.PRICE * " & nRate & "  FROM  FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE DOC_NO =  " & MyParn(xDoc_No.text) & " AND FILE6_20.S_OKAZ = 0 ", nRec
                If nRec > 0 Then
                    MsgBox " ÌÊÃœ Œ’„ Œ«’ »«·›—⁄ "
                    myloadgrd
                End If
            End If
        End If
        CalcTotals
    End If
End If
End If
If cBranch <> "00" Then
    myDisplay "Total Qty  :" & Format(Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption), "#0"), "Total Price:" & Format(Val(xTotal.text), "#0.00")
End If
lSave = False
If xCode.text = "0000" Then
    If Not XONLINE(2).Value Then
        Set CashReg.myForm = Me
        CashReg.Show 1
    Else
        If IsOnLine(xDoc_No.text) Then
             lSave = True
        End If
    End If
    If Not lSave Then Exit Sub
ElseIf xCode.text <> "0000" Then
    If MsgBox(" ”ÃÌ· „»Ì⁄«  √Ã· ··⁄„Ì·", vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    Else
        If MsgBox("”Ê› Ì „ Õ›Ÿ «·»Ê‰ «Ã· - Â· «‰  „ √ﬂœ", vbOKCancel + vbDefaultButton2) = vbOK Then
            lSave = True
        End If
    End If
End If

mysave

mySendInvoices

'If XSALES_RET.Caption <> "" Then con.Execute " UPDATE FILE6_20H SET ISRET = 0 WHERE DOC_NO = " & MyParn(XSALES_RET.Caption)

'Timer1.Enabled = True
DoEvents
If XONLINE(2).Value Then Exit Sub
CmdNewInv_Click
On Error Resume Next
xMan.SetFocus
'grid1.SetFocus

Err.Clear
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdTransFrom_Click()
    Dim cString As String
    transManfrm.sDate = xDate.text
    transManfrm.sCaption = "”Õ» „‰ Œ“Ì‰… " & xBox.text
    transManfrm.sBox1 = xBox.BoundText
    transManfrm.Sbox2 = GetDesca("Select code from file0_50 where type = 1", con)
    transManfrm.Show 1
End Sub
Private Sub cmdTransTo_Click()
    Dim cString As String
    transManfrm.sDate = xDate.text
    transManfrm.sCaption = "«Ìœ«⁄ ›Ì Œ“Ì‰… " & xBox.text
    transManfrm.Sbox2 = xBox.BoundText
    transManfrm.sBox1 = GetDesca("Select code from file0_50 where type = 2", con)
    transManfrm.Show 1
End Sub
Private Sub CmdUndo_Click()
    Dim cString As String
    If cBranch <> "00" Then
        cString = "Select Option6 from USERS"
        Managerfrm.sString = cString
        Managerfrm.sFlag = ActiveControl.Name
        Managerfrm.sFilter = "option6 = 1"
        Managerfrm.Show 1
        If Not IsEmpty(aUser) Then
            If Not aUser(0) Then
                MsgBox "·Ì” ·œÌﬂ ’·«ÕÌ… "
                Exit Sub
            Else
    '           MsgBox " OK "
                myUndo
                grid1.ShowCell grid1.Rows - 1, 1
                grid1.Select grid1.Rows - 1, 1
            End If
        Else
            Exit Sub
        End If
    Else
        myUndo
        grid1.ShowCell grid1.Rows - 1, 1
        grid1.Select grid1.Rows - 1, 1
    End If
On Error Resume Next
grid1.SetFocus
End Sub
Private Sub cmdopen_Click()
    CardLookup " AND PRINTED = 0"
End Sub
Private Sub Command2_Click()
If cBranch <> "00" Then
    If InputBox(" -- ") = "654321" Then
        If InputBox(" -- ") = "123456" Then
        If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
            con.Execute "Delete  From FILE6_20 WHERE DOC_NO IN (SELECT DOC_NO FROM FILE6_20H WHERE BRANCH = " & MyParn(cBranch) & ")"
            con.Execute "Delete  From FILE6_20H WHERE BRANCH = " & MyParn(cBranch)
            Unload Me
        End If
        End If
    End If
End If
End Sub

Private Sub command31_Click()
End Sub

Private Sub Command1_Click()
Dim SalTable As New ADODB.Recordset
Dim nRate As Double, nCount As Double, i As Double
prog1.Visible = True
cStr1 = "SELECT     FILE6_20H.DOC_NO, FILE6_20H.discount, SUM(FILE6_20.TOTAL) AS TOTAL ,rate AS Expr1 FROM         FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO where FILE6_20H.discount <> 0 GROUP BY FILE6_20H.DOC_NO, FILE6_20H.discount,rate "
nCount = GetDesca("SELECT COUNT(DOC_NO) FROM FILE6_20H WHERE FILE6_20H.discount <> 0 ", con)
SalTable.Open cStr1, con, adOpenKeyset, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Min = 0
prog1.Max = nCount
With SalTable
    .MoveFirst
    Do While Not .EOF
        i = i + 1
        prog1.Value = i
        If !TOTAL <> 0 Then nRate = !discount / !TOTAL * 100
        con.Execute " UPDATE FILE6_20H SET RATE = " & nRate & " WHERE DOC_NO = " & MyParn(!doc_no)
        Me.Caption = !doc_no & ""
        .MoveNext
    Loop
End With

End Sub
Private Sub DAYALL_Click()
'    cFilterLook = " FILE6_20H.date = " & DateSq(Format(xdate_open.Text, "DD-MM-YYYY"))
    myUndo
End Sub
Private Sub FIX_COST_Click()
    On Error GoTo myerror
    con.CommandTimeout = 500
    cStr1 = " UPDATE    FILE6_20 SET  cost = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST WHERE FILE6_20.ITEM = Q_ITEMCOST.ITEM AND Q_ITEMCOST.DATE <= FILE6_20H.DATE ORDER BY DATE DESC ) " & _
            " FROM      FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO where year(file6_20h.date) >= " & Year(Date) - 1
    con.Execute cStr1, nRec
    Inform "  „ Ÿ»ÿ  ﬂ·›… «·„»Ì⁄«  " & nRec
    If cBranch = "00" Then
        cStr1 = " UPDATE    FR6_20 SET  cost = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST_FR WHERE FR6_20.ITEM = Q_ITEMCOST_FR.ITEM  AND FR6_20H.STORE = Q_ITEMCOST_FR.STORE AND Q_ITEMCOST_FR.DATE <= FR6_20H.DATE ORDER BY DATE DESC ) " & _
                " FROM      FR6_20 INNER JOIN FR6_20H ON FR6_20.DOC_NO = FR6_20H.DOC_NO where year(fR6_20h.date) >= " & Year(Date) - 1
        con.Execute cStr1, nRec
        Inform "  „ Ÿ»ÿ  ﬂ·›… «·„»Ì⁄«  «· ÊﬂÌ·« " & nRec
    End If
    
    cStr1 = " UPDATE    FILE1_60 SET  cost = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST WHERE FILE1_60.ITEM = Q_ITEMCOST.ITEM AND Q_ITEMCOST.DATE <= FILE1_60H.DATE ORDER BY DATE DESC ) " & _
            " FROM      FILE1_60 INNER JOIN FILE1_60H ON FILE1_60.DOC_NO = FILE1_60H.DOC_NO where year(file1_60h.date) >= " & Year(Date) - 1
    con.Execute cStr1, nRec
    Inform "  „ Ÿ»ÿ  ﬂ·›… «· ÕÊÌ·« " & nRec
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub Form_Activate()
'    On Error Resume Next
    grid1.SetFocus
    Err.Clear
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If xPrinted.Value = 1 Or bedit = False Then Exit Sub
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   On Error Resume Next
   If KeyCode = 117 Then  '  «·ﬂ„Ì…
        If grid1.Row <= grid1.Rows - 2 Then
            grid1.Select grid1.Row, 10, grid1.Row, 10
        Else
            grid1.Select 1, 10, 1, 10
        End If
    End If
    
    If KeyCode = 118 Then  '  «·”⁄—
        grid1.Select 1, 11, 1, 11
    End If
    
    If KeyCode = 114 Then  '  ‰”Ì… Œ’„
        xDiscount.SetFocus
    End If
    
'    If ActiveControl.Name <> GRID1.Name Then
'        If KeyCode = 116 Then   '  ”œ«œ ›« Ê—…
'            If Not IsPrinted Then cmdSave_Click
'        End If
'    End If
    
'    If KeyAscii = 27 Then xTotal.SetFocus   '  «·≈Ã„«·Ï
    
    Err.Clear
End Sub
Private Sub Form_Load()
'On Error GoTo myerror
Dim ImgXTmp As ImgX
Set ImgXTmp = New ImgX
bedit = True
'xinv_no.Locked = (cBranch <> "00")

cmdOpen.Visible = (cBranch <> "00")
cmd_closed.Visible = (cBranch = "00" And bopt2) Or bSupermode
Command3.Visible = (cBranch = "00" And bSupermode)
cNameIp = GetComputerName
cmd_toexel.Visible = (cBranch = "00")
cmd_addexel.Visible = (cBranch = "00")
openCon con

If cBranch <> "00" Then
'    con.Execute " UPDATE FILE6_20H SET DOC_DISC = '000047' WHERE (DOC_DISC IS NULL OR DOC_DISC = '') AND NOTE_DISC = 'EOSS V1 Receit 1'"
'    con.Execute " UPDATE FILE6_20H SET DOC_DISC = '000048' WHERE (DOC_DISC IS NULL OR DOC_DISC = '') AND NOTE_DISC = 'EOSS V1 Receipt 2'"
'    con.Execute " UPDATE FILE6_20H SET DOC_DISC = '000049' WHERE (DOC_DISC IS NULL OR DOC_DISC = '') AND NOTE_DISC = 'EOSS V1 Receipt 3'"
'    con.Execute " UPDATE FILE6_20H SET DOC_DISC = '000041' WHERE  NOTE_DISC = 'chrimtmas sale'"
'    con.Execute " UPDATE FILE6_20H SET DOC_DISC = '000043' WHERE  NOTE_DISC = 'christmas sale'"
End If
CMD_ISRET.Visible = bOpt9
Frame9.Visible = bOpt9
Command2.Visible = lSupperVisor
FIX_VISA.Visible = lSupperVisor
xisbankahly.Visible = (cBranch <= "60")
FIX_COST.Visible = bSupermode
Set ImgXTmp = Nothing
Me.Caption = dSalesDate
myLoadSetting


cIpName = GetComputerName

Frame13.Visible = (cBranch = "00")

aRet1 = aGetDesca("SELECT MAX(ID_NO1)  , MAX(ID_NO2) , MAX(HEAD1) FROM ADDRESS ", con)

aReT2 = aGetDesca("SELECT DATE1_DISC , DATE2_DISC , DISC FROM ADDRESS  WHERE BRANCH = " & MyParn(cBranch), con)

If UBound(aRet1) > 0 Then
    cId_NoDesca1 = "—ﬁ„  ”ÃÌ· ÷—Ì»Ï : " & aRet1(1)
    cId_NoDesca2 = "”Ã·  Ã«—Ï : " & aRet1(2)
    cHeadBoon = aRet1(3) & ""
End If

If UBound(aReT2) > 0 Then
    If aReT2(1) <> "" Then dDiscBranch1 = DateValue(aReT2(1))
    If aReT2(2) <> "" Then dDiscBranch2 = DateValue(aReT2(2))
    nDiscBranch = Val(aReT2(3) & "")
End If


If lIsBranchStore Then
    aReT2 = aGetDesca("SELECT DISPLAY , PORT , CountPrint FROM COMP WHERE BRANCH = " & MyParn(cBranch), con)
    If UBound(aReT2) > 0 Then
        cMsgDisplay = aReT2(1)
        cComPort = aReT2(2)
        nCountPrint = Val(aReT2(3) & "")
    End If
Else
    aReT2 = aGetDesca("SELECT DISPLAY , PORT , CountPrint FROM COMP ", con)
    If UBound(aReT2) > 0 Then
        cMsgDisplay = aReT2(1)
        cComPort = aReT2(2)
        nCountPrint = Val(aReT2(3) & "")
    End If
End If
If nCountPrint = 0 Then nCountPrint = 1

If cBranch <> "00" Then
    opencom
    Frame12.Visible = (cBranch = "00")
    ClearDisplay
    myDisplay cMsgDisplay, " "
End If

CMD_PAST_PURCH.Visible = (cBranch = "00")

nMaxDisc2 = Val(GetDesca("SELECT MAXDISC FROM DISCOUNT", con) & "")
sned_purch.Visible = (cBranch = "00")


nRound = 2
'nCountPrint = 1
If cBranch = "00" And bopt2 Then
    FRM_CLOSED.Visible = True
Else
    FRM_CLOSED.Visible = False
End If

CmdInform.Visible = Not bopt1
Frame6.Visible = Not bopt1

xBox.Enabled = False
xStore.Enabled = (cBranch = "00")

If cBranch = "00" Then
    cBranchStore = retFlag(aPassword, "store")
    cBranchBox = retFlag(aPassword, "BOX")
    If cBranchStore = "" Then
        cBranchStore = "000"
        cBranchBox = ""
    End If
End If

FixAddress
Dim aRet As Variant

lManSales = False

cDefClient = "0000"
cDefClientDesca = "⁄„Ì· ‰ﬁœÌ"

If cUserBox = "" Then cBranchBox = ""
cMoveName = "„»Ì⁄« "
Me.Caption = " „»Ì⁄« " & DS
Me.cmd_tsal.Caption = "„»Ì⁄« " & vbCrLf & "«·ÌÊ„"

If cBranch > "60" Then
    Set data12.Recordset = myRecordSet("SELECT * FROM FILE6_25 WHERE ISSTOP = 0 ORDER BY DESCA ", con)
Else
    If lNotBranch Then
        Set data12.Recordset = myRecordSet("SELECT * FROM FILE6_25  ORDER BY DESCA ", con)
    Else
        Set data12.Recordset = myRecordSet("SELECT * FROM FILE6_25 WHERE ISSTOP = 0 AND ( CODE = '0001' OR BRANCH = " & MyParn(cBranch) & " ) ORDER BY DESCA ", con)
    End If
End If
Set xMan.RowSource = data12
xMan.ListField = "Desca"
xMan.BoundColumn = "Code"

If lIsBranchStore Then
    Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_40 WHERE BRANCH = " & MyParn(cBranch), con)
Else
    Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_40", con)
End If
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
xStore.BoundText = cBranchStore

Set DATA2.Recordset = myRecordSet("SELECT * FROM MOSM ", con)
Set xMosm.RowSource = DATA2
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"

If lIsBranchStore Then
    Set DATA4.Recordset = myRecordSet("SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch), con)
Else
    Set DATA4.Recordset = myRecordSet("SELECT * FROM FILE0_50", con)
End If
Set xBox.RowSource = DATA4
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"
xBox.BoundText = cBranchBox

If GetBoolean("SELECT ONLINE FROM FILE0_40 WHERE CODE = " & MyParn(cBranchStore), con) Then
    lIsBoxOnline = True
    If cBranchBox <> "" Then
        lIsBoxOnline = GetBoolean("SELECT ISBOXONLINE FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con)
    End If
End If

Frame7.Visible = lIsBoxOnline
CMD_ONLINE.Visible = lIsBoxOnline Or (cBranch = "00") Or lSupperVisor

If cBranch = "00" Or lIsBranchStore Then
    Set data8.Recordset = myRecordSet("SELECT * FROM VISA WHERE BRANCH = " & MyParn("00"), con)
Else
    Set data8.Recordset = myRecordSet("SELECT * FROM VISA WHERE BRANCH = " & MyParn(cBranch), con)
End If
Set XCODEVISA.RowSource = data8
XCODEVISA.ListField = "Desca"
XCODEVISA.BoundColumn = "Code"

Set grid1.DataSource = DATA11

cmd_delall.Visible = bSupermode And cBranch <> "00"

'nDayDiscout = IsDayDisc(dSalesDate)
'If nDayDiscout > 0 Or lIsPrice_2 Then lDiscOpenDay = True
'If lDiscOpenDay Then
'    Label20.Caption = cDescDiscout
'    Label20.Visible = True
'Else
'    Label20.Caption = ""
'    Label20.Visible = False
'End If

If cBranchBox = "" And cBranch <> "00" Then
    MsgBox "«·„” Œœ„ «·Õ«·Ï ·Ì” ·Â… Œ“‰… „”Ã·… ° ·« Ì„ﬂ‰  ”ÃÌ· „»Ì⁄«  ·…"
    bedit = False Or bopt2
Else
    bedit = True
End If

'myUndo


Label19.Visible = (cBranch = "00")
If sDoc_no = "" Then
    xBox.BoundText = cManBox
    CmdNewInv_Click
Else
    xdate_open.text = GetDesca("SELECT DATE FROM FILE6_20H WHERE DOC_NO = " & MyParn(sDoc_no), con)
'   myUndo
    xDoc_No.text = sDoc_no
    myUndo
    Exit Sub
End If
If Not IsDate(dSalesDate) Then dSalesDate = Date
If DateValue(dSalesDate) <> DateValue(Date) Then
    MsgBox " «—ÌŒ «·ÃÂ«“ „Œ ·› ⁄‰  «—ÌŒ «·„»Ì⁄«  ,  »—Ã«¡ „—«Ã⁄…  «—ÌŒ «·»Ì⁄"
    DATE0.Visible = True
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

CardTable.Close
Set CardTable = Nothing

closeCon con

Unload Search3
Unload oSearchDoc
Unload search32
Set salesFrm = Nothing

'If sBranchCode <> "00" And bopt1 And Not lMainShow Then
'    End
'End If

'If CBRAsBranchCode <> "00" Then End
Err.Clear
End Sub
Public Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim cLine1 As String
Dim cLine2 As String
On Error GoTo myerror
Dim bNew As Boolean
If IsPrinted Then
    Exit Sub
End If
With grid1
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
    myaddItem
End If
CalcTotals
If myreplace(Row) Then
    HandleCntEdit
    bNew = grid1.TextMatrix(Row, .Cols - 1) = ""
End If

myloadgrd

If bNew Then
    grid1.Row = grid1.Rows - 1
    grid1.ShowCell grid1.Rows - 1, 1
End If

If Not lIsBranchStore Then
    If Col = 10 And Val(.TextMatrix(.Row, 10)) > 0 Then
        xBalance.Caption = LastBalance(grid1.TextMatrix(Row, 1), xStore.BoundText, con) ' + Val(GetDesca("select quant from file6_20 where doc_no = " & MyParn(xDoc_No.Text)) & "")
        If Val(xBalance.Caption) < Val(.TextMatrix(Row, 10)) Then
            MsgBox "«·—’Ìœ «ﬁ· „‰ «Ê = ’›—"
        End If
    End If
End If
cLine1 = .TextMatrix(.Rows - 2, 5) & " Price : " & Format(.TextMatrix(.Rows - 2, 11), "#0.00")
cLine2 = "Qty:" & Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption) & " Total:" & Format(Val(xTotal.text), "#0.00")

myDisplay cLine1, Mid(cLine2, 1, 20)

End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myEdit(Row As Long, Col As Long)
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        grid1.RemoveItem OldRow
        CalcTotals
    End If
End If
End Sub
Private Sub grid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    GRID1.EditMaxLength = IIf(Col = 10, 7, 0)
End Sub
Private Sub grid1_EnterCell()
If XISNODEL.Value <> 0 Then Exit Sub
If IsPrinted Or (XSALES_RET.Caption <> "" And grid1.ValueMatrix(grid1.Row, 10) < 0) Then
    grid1.Editable = flexEDNone
    Exit Sub
End If
If grid1.TextMatrix(grid1.Row, 17) <> "" Then
    If ISPICT.Value <> 0 Then myLoadPicture grid1.TextMatrix(grid1.Row, 17)
ElseIf grid1.Row = grid1.Rows - 1 And Trim(grid1.TextMatrix(grid1.Row, 1)) = "" And grid1.Row > 1 Then
    If ISPICT.Value <> 0 Then myLoadPicture grid1.TextMatrix(grid1.Row - 1, 17)
End If

If grid1.TextMatrix(grid1.Row, 15) = "" Then
    grid1.TextMatrix(grid1.Row, 15) = xMan.BoundText
    grid1.TextMatrix(grid1.Row, 16) = xMan.text
End If

If ((xPrinted.Value = 1 And xClosed.Value = -1) Or bedit = False) Then
    grid1.Editable = flexEDNone
    Exit Sub
End If


With grid1
    If .Col = 1 Or .Col = 10 Or (.Col = 11 And ((cBranch = "00") Or lIsBoxOnline)) Or .Col = 17 Or .Col = 15 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
    
'    .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = vbWhite
'    .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = vbYellow

End With
End Sub
Private Sub Grid1_GotFocus()
On Error Resume Next
If xMan.BoundText = "" Then xMan.SetFocus
If grid1.Rows < 2 Then Exit Sub
If grid1.Row = 0 Then
    grid1.Row = 1
    grid1.Col = 1
End If
grid1_EnterCell
Err.Clear
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If xPrinted.Value = 1 And Not bopt2 Then Exit Sub
'If KeyCode = 115 Or (KeyCode = 13 And Shift = 2) Then xDiscount.SetFocus
'If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 And validRows(grid1.Row) Then
'    grid1.AddItem "", grid1.Row
'End If
If KeyCode = 112 And xPrinted.Value = 0 And grid1.TextMatrix(grid1.Row, 1) = "" Then
'    Grid1.Row = Grid1.Rows - 1
'    Grid1.Col = 1
    ItemsLookupAll Me, oSearchItem
'    VsModelLook2.Show 1
End If
End Sub
Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If xPrinted.Value = 1 And Not bopt2 Then Exit Sub
End Sub
Private Sub Grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.Row = grid1.Rows - 1 And grid1.Row > 1 And Trim(grid1.TextMatrix(grid1.Row, 1)) = "" Then
        KeyAscii = 0
'        On Error Resume Next
'        xTotal.SetFocus
        If Not IsPrinted Then cmdSave_Click
        Err.Clear
        
    End If
End If
If KeyAscii = 43 Then
    KeyAscii = 0
    If grid1.Row > 1 Then
        grid1.Row = grid1.Row - 1
        grid1.Col = 10
    End If
End If
End Sub

Private Sub grid1_LostFocus()
imgx1.Images.Clear
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
    'CalcTotals
End If
'    xBalance.Caption = ""
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Dim sBarCode As String
Dim nSal As Double
Dim nRet As Double
Dim cPassword As String, aUser As Variant
With grid1
    If Col = 10 And Val(.EditText) < 0 Then
        Dim cString As String
        ManagPass.pPassWord = ""
        ManagPass.Show 1
        cPassword = ManagPass.pPassWord
        If Trim(cPassword) = "" Then
            Cancel = True
            Exit Sub
        End If
        aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( isreturn = 1 and (( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & ")))", con)
        If UBound(aUser) > 0 Then
            lRetDoc = True
            xusername_RET.Caption = aUser(2)
            MsgBox "Ì„ﬂ‰  ”ÃÌ· «·„— Ã⁄"
        Else
            Cancel = True
            Exit Sub
        End If
    End If
    
    If Col = 1 Then
        If .EditText <> "" And Len(.EditText) > 2 Then
            If GrdDesc(.EditText, Row) Then
                .EditText = .TextMatrix(.Row, 1)
            Else
                Cancel = True
            End If
             
            If Len(.EditText) > 1 And Mid(.EditText, 1, 1) = "+" Then
                .TextMatrix(.Row - 1, 10) = Val(Mid(.EditText, 2))
                .TextMatrix(.Row - 1, 13) = Val(.TextMatrix(.Row - 1, 10)) * Val(.TextMatrix(.Row - 1, 11))
                .EditText = ""
                CalcTotals
                grid1_AfterEdit .Row - 1, 10
                Exit Sub
            End If
        End If
    ElseIf Col = 10 Then
        If Val(.EditText) > 0 And Val(.TextMatrix(Row, 10)) < 0 Then
            MsgBox "ÌÃ» Õ–› «·”Ã· ﬁ»·  ⁄œÌ·… „‰ „— Ã⁄ ≈·Ï „»Ì⁄« "
            Cancel = True
            Exit Sub
        End If
    End If
End With
End Sub
Private Sub cmdPrint_Click()
    If cBranch = "00" And Not lNotBranch Then
        doprint_doc
        SavePrint
    Else
        If grid1.Rows = 2 Then Exit Sub
        If Not XONLINE(2).Value Then
            If Val(xTotal.text) <> Val(xvisa.Caption) + Val(xcash.Caption) + Val(XPOINT.Caption) Then
                MsgBox "„—«Ã⁄… ”œ«œ «·»Ê‰ "
                Exit Sub
            End If
            If XONLINE(2).Value = 0 Then
                If doprint(True) Then SavePrint
            Else
                If doprint_online() Then SavePrint
            End If
        Else
            If XONLINE(2).Value = 0 Then
                If doprint(True) Then SavePrint
            Else
                If doprint_online() Then SavePrint
            End If
        End If
    End If
End Sub

Private Sub Text1_LostFocus()
    If xCode.text = "" Then Exit Sub
    xCode.text = RetZero(xCode.text, 3)
    Dim aRet
    aRet = aGetDesca("select desca,cash from FILE3_10 where code = " & MyParn(xCode.text), con)
    If UBound(aRet) > 0 Then
        xCodeDesca.Caption = aRet(1) & ""
        chkCash.Value = IIf(aRet(2), 1, 0)
    Else
        xCode.text = ""
        xCode.SetFocus
    End If
End Sub

Private Sub Label35_Click()

End Sub

Private Sub LOOKPH_BARCODE_Click()
    Set ItemLookPh.myForm = salesFrm
    ItemLookPh.Show
End Sub

Private Sub LOOKPH_Click()
    If xPrinted.Value = 0 Then
        Set ItemLookPh_pict.myForm = Me
        ItemLookPh_pict.Show
    End If
End Sub
Private Sub SSCommand1_Click()
    Set ItemLookPh.myForm = Me
    ItemLookPh.Show
End Sub


Private Sub SSCommand2_Click()
''''If cManBox = "" Then
''''    MsgBox "·« ÌÊÃœ ﬂ«‘Ì—"
''''    Exit Sub
''''End If
''''If Not bOpt4 Then
''''    Dim cString As String
''''    cString = "Select Option6 from USERS"
''''    Managerfrm.sString = cString
''''    Managerfrm.sFlag = ActiveControl.Name
''''    Managerfrm.sFilter = "option4 = 1"
''''    Managerfrm.Show 1
''''    If Not IsEmpty(aUser) Then
''''        If Not aUser(0) Then
''''            MsgBox "·Ì” ·œÌﬂ ’·«ÕÌ… ⁄—÷ ≈Ã„«·Ï «·„»Ì⁄« "
''''            Exit Sub
''''        Else
''''            PrintTDay
''''        End If
''''    Else
''''        Exit Sub
''''    End If
''''Else
''''    PrintTDay
''''End If



Dim cPassword As String, aUser As Variant
    ManagPass.pPassWord = ""
    ManagPass.Show 1
    cPassword = ManagPass.pPassWord
    If Trim(cPassword) = "" Then Exit Sub
    aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
    If UBound(aUser) = 0 Then Exit Sub
    PrintTDay
End Sub
Private Sub XBARCODEADD_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CMD_ADDITEM_Click
End Sub
Private Sub xBox_GotFocus()
myGotFocus xBox
If Not xBox.MatchedWithList Then xBox.BoundText = ""
End Sub
Private Sub xBox_LostFocus()
myLostFocus xBox
xcash.Enabled = (Trim(xBox.BoundText) <> "")
CalcTotals
End Sub
Private Sub xcard_disc_LostFocus()
    If xPrinted.Value = 0 Then CheckCardDisc
End Sub
Private Sub xcash_Click()
    If bopt2 And xDoc_No.text <> "" And XONLINE(2).Value = 0 Then
        Set CashReg.myForm = Me
        CashReg.Show 1
'        myreplace_sub
        myUndo
    End If
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And xPrinted.Value = 0 Then CLIENTLOOKUP
End Sub
Private Sub xCode_LostFocus()
    myLostFocus xCode
    xCodeDesca.Caption = ""
    If xCode.text = "" Then Exit Sub
    xCode.text = RetZero(xCode.text, 3)
    Dim aRet
    aRet = aGetDesca("select desca,cash from FILE3_10 where code = " & MyParn(xCode.text), con)
    If UBound(aRet) > 0 Then
        xCodeDesca.Caption = aRet(1) & ""
        chkCash.Value = IIf(aRet(2), 1, 0)
    Else
        xCode.text = ""
        xCode.SetFocus
    End If
End Sub
Private Sub xcode_Validate(Cancel As Boolean)
'If Trim(xCode.Text) = "" Then Cancel = True
End Sub


Private Sub xcodevisa_LostFocus()
'    con.Execute " UPDATE FILE6_20H SET ISNEW = 1 , CODEVISA = " & addvalue(xcodevisa.BoundText) & " WHERE DOC_NO = " & MyParn(xdoc_no.Text), nRec
'    Inform "  „  ⁄œÌ· ⁄œœ  " & nRec
End Sub

Private Sub xdate_open_LostFocus()
    myValidDate xdate_open
End Sub

Private Sub xDate_Validate(Cancel As Boolean)
If Not IsDate(xDate.text) Then Cancel = True
End Sub

Private Sub xDiscount_LostFocus()
myLostFocus xDiscount
CalcTotals
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
'If xDoc_No.Text = "" Then
'    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
'    Exit Function
'End If

If Trim(xCode.text) = "" Then
    MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì· «Ê ﬂÊœ „Ê—œ"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
If cBranch <> "00" Then
End If
If xStore.BoundText = "" Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If xBox.BoundText = "" And cBranch <> "00" Then
    If Not bIgMsg Then MsgBox "·„ Ì „  ÕœÌœ «·Œ“‰…"
    Exit Function
End If

If xCodeDesca.Caption = "" Then
    If Not bIgMsg Then MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì·"
    Exit Function
End If

'cCodeStore = GetDesca("select CUST CODE from FILE0_40 WHERE CUST = " & MyParn(xCode.Text) & " AND CODE = " & MyParn(XSTORE.BoundText))
'If cCodeStore <> "" Then
'    If cCodeStore <> XSTORE.BoundText Then
'        If Not bIgMsg Then MsgBox "„Œ“‰ €Ì— ’ÕÌÕ "
'        Exit Function
'    End If
'End If

If cBranch = "00" And Not lNotBranch Then
    cCustStore = GetDesca("select CUST from FILE0_40 WHERE CODE = " & MyParn(xStore.BoundText), con)
    If cCustStore <> "" Then
        If cCustStore <> xCode.text Then
            If Not bIgMsg Then MsgBox "⁄„Ì· €Ì— ’ÕÌÕ "
            Exit Function
        End If
    End If

    cCodeStore = GetDesca("select CODE from FILE0_40 WHERE CUST = " & MyParn(xCode.text) & " AND CODE = " & MyParn(xStore.BoundText), con)
    If cCodeStore <> "" Then
        If cCodeStore <> xStore.BoundText Then
            If Not bIgMsg Then MsgBox "„Œ“‰ €Ì— ’ÕÌÕ "
            Exit Function
        End If
    End If
End If
MYVALID = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
xdoc_ret.text = ""

If cBranch <> "00" Then
    If CheckOpen(False, xBox.BoundText) > 0 Then
        cmdOpen.BackColor = vbRed
    Else
        cmdOpen.BackColor = &HE0E0E0
    End If
End If

xUUID_RC.Caption = CardTable!UUID_RC & ""
XBRANCH.Caption = CardTable!branch & ""
xCode.Enabled = False
XISRETS.Value = 0
XISONEST.Value = 0
XISONEST.Value = IIf(CardTable!ISONEST, 1, 0)
XISNODEL.Value = IIf(CardTable!ISNODEL, 1, 0)
xIpName.Caption = CardTable!user_ip & ""
XSHIP_NO.Caption = CardTable!SHIP_NO & ""
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
Select Case CardTable!online
    Case 0
        XONLINE(0).Value = 1
    Case 1
        XONLINE(1).Value = 1
    Case 2
        XONLINE(2).Value = 1
End Select
xDoc_No.text = CardTable!doc_no
xinv_no.text = CardTable!INV_NO & ""
xDoc_no2.text = CardTable!Doc_no2
xphone.Caption = CardTable!phone & ""
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!STORE & ""
xMan.BoundText = CardTable!MAN & ""
xBox.BoundText = CardTable!BOX & ""
XVISADOC.Caption = CardTable!VISADOC & ""
XCODEVISA.BoundText = CardTable!CODEVISA & ""
xisbankahly.Value = IIf(CardTable!isbankahly, 1, 0)
xNotes.text = CardTable!NOTES & ""
XONLINE_DOC.Caption = CardTable!ONLINE_DOC & ""
XTRANS_DOC.Caption = CardTable!TRANS_DOC & ""
XSALES_RET.Caption = CardTable!SALES_RET & ""
xusername_RET.Caption = CardTable!username_ret & ""
xNOTE_DISC.Caption = CardTable!NOTE_DISC & ""
xDOC_DISC.Caption = CardTable!DOC_DISC & ""

xCode.text = CardTable!CODE & ""
xCodeDesca.Caption = CardTable!ClientDesca & ""

xcard_disc.text = CardTable!card_disc & ""
xcard_desca.Caption = CardTable!card_desca & ""
xcard_value.Caption = CardTable!card_VALUE & ""
xcard_doc.Caption = CardTable!card_DOC & ""



xusername.Caption = CardTable!UserName & ""
XUSERNAME_DISC.Caption = CardTable!username_disc & ""
xphone.Caption = CardTable!phone & ""

chkCash.Value = IIf(CardTable!CASH, 1, 0)
xDiscount.text = TurnValue(Val(CardTable!discount & ""), 0, "")
xRate.text = TurnValue(Val(CardTable!Rate & ""), 0, "")
xcash.Caption = Format(CardTable!CASH, "#0.00")
xvisa.Caption = TurnValue(Val(CardTable!Visa & "") + Val(CardTable!VISA3 & ""))
XVISA2.Caption = TurnValue(CardTable!VISA2)
XTYPEVISA.Caption = CardTable!TYPEVISA & ""

XISRET.Value = 0
xPay.Caption = Myvalue(CardTable!PAY)
XPOINT.Caption = Format(CardTable!Point, "#0.00")
xRest.Caption = Myvalue(CardTable!rest)
xlate.Caption = Myvalue(CardTable!late)
xPrinted.Value = IIf(CardTable!PRINTED, 1, 0)
'xTotal.Enabled = IIf(xPrinted.Value = 0, True, False)

xtime.Caption = Format(CardTable!Time, "hh:nn")
myloadgrd
Handlecontrols LoadMode
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myDefine()
XISRET.Value = 0
If cBranch <> "00" Then
    If CheckOpen(False, xBox.BoundText) > 0 Then
        cmdOpen.BackColor = vbRed
    Else
        cmdOpen.BackColor = &HE0E0E0
    End If
End If
xUUID_RC.Caption = ""

XISNODEL.Value = 0
XBRANCH.Caption = cBranch
XSHIP_NO.Caption = ""
XONLINE_DOC.Caption = ""
lDiscModelRet = False
xIpName.Caption = GetComputerName
XONLINE(0).Value = 1
XTRANS_DOC.Caption = ""
XISRETS.Value = 0
'XCODEVISA.Enabled = True
XPOINT.Caption = ""
sUserDisc = ""
lIsDocRet = False
xMan.BoundText = ""
XUSERNAME_DISC.Caption = ""
xRate.Locked = True
xDiscount.Locked = True
xdoc_ret.text = ""
myDisplay cMsgDisplay, " "
cmd_closed.BackColor = &H8080FF
cmd_closed.Caption = "› Õ „” ‰œ"
XVISADOC.Caption = ""
xClosed.Visible = False
xClosed.Value = ssCBUnchecked
XGIFT.Value = 0
xusername_RET = ""
xNOTE_DISC.Caption = ""
xDOC_DISC.Caption = ""

xcard_disc.text = ""
xcard_desca.Caption = ""
xcard_value.Caption = ""
xcard_doc.Caption = ""

XSALES_RET.Caption = ""
xphone.Caption = ""
xDoc_No.text = ""
xDoc_No.Tag = DefineMode
xDoc_no2.text = ""
xinv_no.text = ""
xusername.Caption = cusername
If cBranch = "00" Then
    f = Format(Date, "DD-MM-YYYY")
    xStore.BoundText = cBranchStore
    xBox.BoundText = cBranchBox
    xCode.text = IIf(cBranchBox = "", "", "0000")
    xCodeDesca.Caption = IIf(cBranchBox = "", "", cDefClientDesca)
    xDate.Enabled = True
    xDate.text = Format(Date, "dd-mm-yyyy")
Else
    xStore.BoundText = cBranchStore
    If lIsBranchStore Then
        dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES WHERE BRANCH = " & MyParn(cBranch), con), "DD-MM-YYYY")
    Else
        dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES", con), "DD-MM-YYYY")
    End If
    xBox.BoundText = cBranchBox
    xCode.text = "0000"
    xCodeDesca.Caption = cDefClientDesca
End If
xDate.text = Format(dSalesDate, "DD-MM-YYYY")
xBalance.Caption = ""




xMosm.BoundText = cPMosm
'xBoxVISA.BoundText = ""
XCODEVISA.BoundText = ""

xDiscount.text = ""
chkCash.Value = 1
xisbankahly.Value = 0
xTotalQuant.Caption = ""
xTotalQuant2.Caption = ""
xTotalItemNoDiscount.Caption = ""
xTotalItem.Caption = ""
xTotalItemDiscount.Caption = ""
xRate.text = ""
xTotal.text = ""
xDiscount.text = ""
xlate.Caption = ""
xvisa.Caption = ""

XVISA2.Caption = ""
XTYPEVISA.Caption = ""

xcash.Caption = ""
xPrinted.Value = 0
xNotes.text = ""
xRest.Caption = ""
xPay.Caption = ""
xtime.Caption = Format(Time, "hh:nn")



grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
xDoc_No.Tag = 0
xTotal.text = ""
imgx1.Images.Clear
Fixgrd

'If lManSales Then
'    grid1.TextMatrix(1, 1) = cdefman
'    GrdDesc grid1.TextMatrix(1, 1), 1
'    grid1.TextMatrix(1, 1) = ""
'End If
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
'cmdNewInv.Enabled = nMode = LoadMode
cmdSave.Enabled = (bedit And xPrinted.Value = 0 And grid1.Rows > 2)
xdoc_ret.Enabled = (xPrinted.Value = 0)
cmd_showdoc.Enabled = (xPrinted.Value = 0)
cmd_disc.Enabled = (bedit And xPrinted.Value = 0 And grid1.Rows > 2)
CmdDel.Enabled = (bedit And xPrinted.Value = 0)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
'cmdPrint.Enabled = (nMode = LoadMode)

xBox.Enabled = False
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
xcash.Enabled = (Trim(xBox.BoundText) <> "")
xCode.Enabled = xPrinted.Value = 0 And bedit = True

If xBox.text = "" Or xStore.text = "" Then
    CmdDel.Visible = True
    CmdDel.Enabled = True
End If


End Sub
Private Function retBool(cFieldName) As Boolean
If Not (CardTable.EOF Or CardTable.BOF) Then
    retBool = CardTable(cFieldName)
End If
End Function
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload True
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If IsPrinted Then Exit Sub
If XISNODEL.Value <> 0 Then Exit Sub
If (XSALES_RET.Caption <> "" And grid1.ValueMatrix(grid1.Row, 10) < 0) Then Exit Sub

If KeyCode = 46 And grid1.Row = grid1.Rows - 1 And xPrinted.Value = 0 And bedit And grid1.Rows > 3 Then
    grid1.Select grid1.Rows - 2, 0
End If
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And CmdDel.Enabled And grid1.Rows > 3 Then
'    If IsPrinted Then
'        MsgBox " „  ÿ»«⁄… «·»Ê‰ „‰ ﬁ»·"
'        Exit Sub
'    End If
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem (grid1.Row)
        grid1.Select grid1.Rows - 1, 1
        grid1.ShowCell grid1.Rows - 1, 1
        
        CalcTotals
        'UpdateHeader
        MakeSerial grid1.Row
    End If
End If
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
'If KeyCode = 116 Then
'    cmdSave_Click
'End If
End Sub
Public Function GrdDesc(sitem As String, Row As Long) As Boolean
'On Error GoTo myerror
Dim lMosmDisc As Boolean, cBarCode As String, nItem As Double
With grid1
    If Len(sitem) > 2 Then
        If Len(sitem) > 7 Or Not IsNumeric(sitem) Then
            cBarCode = sitem
            nItem = 0
        Else
            cBarCode = ""
            nItem = sitem
        End If
        Dim loctable As ADODB.Recordset
        If nItem <> 0 Then
            Set loctable = ItemFind(sitem, con)
        Else
            Set loctable = ItemFind_BARCODE(cBarCode, con)
        End If
        If Not (loctable.EOF And loctable.BOF) Then
'            If loctable!ISSTOP Then
'                MsgBox "«·’‰› „ Êﬁ› "
'                Exit Function
'            End If
            
            If Val(loctable!price & "") = 0 Then
                MsgBox "”⁄— «·»Ì⁄ ’›— "
                Exit Function
            End If
              
            grid1.TextMatrix(Row, 2) = loctable!MOSM
            grid1.TextMatrix(Row, 1) = loctable!Item
            grid1.TextMatrix(Row, 3) = loctable!Fact
            grid1.TextMatrix(Row, 4) = loctable!SUPP & ""
            grid1.TextMatrix(Row, 5) = loctable!modelfact0 & ""
            grid1.TextMatrix(Row, 6) = loctable!DESCA
            grid1.TextMatrix(Row, 7) = loctable!scal
            grid1.TextMatrix(Row, 8) = loctable!color
            If Trim(.TextMatrix(Row, 10)) = "" Then grid1.TextMatrix(Row, 10) = "1"
            If cBranch = "00" Then
                grid1.TextMatrix(Row, 9) = Format(loctable!price, "#0.00")
                If cBranchBox = "" Then
                    grid1.TextMatrix(Row, 11) = Val(loctable!cost2 & "")
                    grid1.TextMatrix(Row, 19) = Val(loctable!cost2 & "")
                    grid1.TextMatrix(Row, 14) = Val(loctable!cost2 & "")
                Else
                    grid1.TextMatrix(Row, 11) = Val(loctable!price & "")
                    grid1.TextMatrix(Row, 19) = Val(loctable!price & "")
                    grid1.TextMatrix(Row, 14) = Val(loctable!price & "")
                End If
                grid1.TextMatrix(Row, 12) = 0
            Else
                grid1.TextMatrix(Row, 9) = Format(loctable!price, "#0.00")
                If lIsPrice_2 And loctable!PRICE_2 <> 0 Then
                    If Val(loctable!PRICE_2 & "") <= Val(loctable!price & "") Then
                        grid1.TextMatrix(Row, 11) = IIf(Val(loctable!PRICE_2 & "") > 0, Format(loctable!PRICE_2, "#0.00"), Format(loctable!price, "#0.00"))
                        grid1.TextMatrix(Row, 12) = loctable!okaz_2 & ""
                        grid1.TextMatrix(Row, 20) = loctable!okaz_2 & ""
                    
                        grid1.TextMatrix(Row, 19) = IIf(Val(loctable!PRICE_2 & "") > 0, Format(loctable!PRICE_2, "#0.00"), Format(loctable!price, "#0.00"))
                    
                    Else
                        grid1.TextMatrix(Row, 11) = Format(loctable!price, "#0.00")
                        grid1.TextMatrix(Row, 19) = Format(loctable!price, "#0.00")
                        grid1.TextMatrix(Row, 12) = 0
                    End If
                Else
                    If lokazItem And loctable!ISOKAZITEM Then
                        If Val(loctable!PRICE2 & "") <= Val(loctable!price & "") Then
                            grid1.TextMatrix(Row, 11) = IIf(Val(loctable!PRICE2 & "") > 0, Format(loctable!PRICE2, "#0.00"), Format(loctable!price, "#0.00"))
                            grid1.TextMatrix(Row, 19) = IIf(Val(loctable!PRICE2 & "") > 0, Format(loctable!PRICE2, "#0.00"), Format(loctable!price, "#0.00"))
                            grid1.TextMatrix(Row, 12) = loctable!okaz
                            grid1.TextMatrix(Row, 20) = loctable!okaz
                        Else
                            grid1.TextMatrix(Row, 11) = Format(loctable!price, "#0.00")
                            grid1.TextMatrix(Row, 19) = Format(loctable!price, "#0.00")
                            grid1.TextMatrix(Row, 12) = 0
                        End If
                    Else
                        If lokaz And Val(loctable!PRICE2 & "") <= Val(loctable!price & "") And xisokaz.Value = 0 Then
                            grid1.TextMatrix(Row, 11) = IIf(Val(loctable!PRICE2 & "") > 0, Format(loctable!PRICE2, "#0.00"), Format(loctable!price, "#0.00"))
                            grid1.TextMatrix(Row, 19) = IIf(Val(loctable!PRICE2 & "") > 0, Format(loctable!PRICE2, "#0.00"), Format(loctable!price, "#0.00"))
                            grid1.TextMatrix(Row, 12) = loctable!okaz
                            grid1.TextMatrix(Row, 20) = loctable!okaz
                        Else
                            grid1.TextMatrix(Row, 11) = Format(loctable!price, "#0.00")
                            grid1.TextMatrix(Row, 19) = Format(loctable!price, "#0.00")
                            grid1.TextMatrix(Row, 12) = 0
                        End If
                    End If
                End If
            End If
            grid1.TextMatrix(Row, 13) = Val(grid1.TextMatrix(Row, 10)) * Val(grid1.TextMatrix(Row, 11))
            grid1.TextMatrix(Row, 17) = loctable!MODELNO & ""
            
            If Not lIsBranchStore Then
                xBalance.Caption = LastBalance(grid1.TextMatrix(Row, 1), xStore.BoundText, con) - Val(GetDesca("select quant from file6_20 where ITEM = " & Val(grid1.TextMatrix(Row, 1)) & " AND doc_no = " & MyParn(xDoc_No.text), con) & "")
                If Val(xBalance.Caption) <= 0 And Val(grid1.TextMatrix(grid1.Row, 10)) > 0 Then
                    MsgBox "«·—’Ìœ «ﬁ· „‰ «Ê = ’›—"
                End If
            End If
            GrdDesc = True
            If ISPICT.Value <> 0 Then myLoadPicture loctable!MODELNO & ""
        End If
    Else
'        Dim aRet As Variant
'        aRet = aGetDesca("SELECT CODE,DESCA FROM FILE6_25 WHERE ISSTOP = 0 AND CODE = " & MyParn(sitem))
'        If UBound(aRet) > 0 Then
'            grid1.TextMatrix(Row, 15) = aRet(1) & ""
'            grid1.TextMatrix(Row, 16) = aRet(2) & ""
'            GrdDesc = True
'        Else
'            Inform "ﬂÊœ «·»«∆⁄ €Ì— ’ÕÌÕ"
'        End If
    End If
End With
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotalQuant As Double, nTotalQuant2 As Double, nTotalItem As Double, nTotalItemNoDiscount As Double, nTotalDiscount As Double, nTotalMin As Double
Dim nTotal2 As Double
With grid1
For i = 1 To grid1.Rows - 2
    nTotal2 = nTotal2 + (Val(.TextMatrix(i, 10)) * Val(.TextMatrix(i, 9)))
    .TextMatrix(i, 13) = Val(.TextMatrix(i, 10)) * Val(.TextMatrix(i, 11))
    If cBranch = "00" Then
        If Val(.TextMatrix(i, 9)) <> Val(.TextMatrix(i, 11)) And xStore.BoundText <> "000" Then
            .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = vbRed
            .Cell(flexcpFontStrikethru, i, 0, i, .Cols - 1) = True
        End If
    End If
    
    If Val(grid1.TextMatrix(i, 10)) > 0 Then
        nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 10))
    Else
        nTotalQuant2 = nTotalQuant2 + Abs(Val(grid1.TextMatrix(i, 10)))
    End If
    nTotalItem = nTotalItem + Val(.TextMatrix(i, 10)) * Val(.TextMatrix(i, 11))
    If cBranch = "00" Then
        nTotalItemNoDiscount = nTotalItemNoDiscount + (Val(.TextMatrix(i, 11)) * Val(.TextMatrix(i, 10)))
        nTotalItemDiscount = nTotalItemDiscount + (Val(.TextMatrix(i, 11)) * Val(.TextMatrix(i, 10))) - (Val(.TextMatrix(i, 11)) * Val(.TextMatrix(i, 10)))
    Else
        nTotalItemNoDiscount = nTotalItemNoDiscount + (Val(.TextMatrix(i, 9)) * Val(.TextMatrix(i, 10)))
        nTotalItemDiscount = nTotalItemDiscount + (Val(.TextMatrix(i, 9)) * Val(.TextMatrix(i, 10))) - (Val(.TextMatrix(i, 11)) * Val(.TextMatrix(i, 10)))
    End If
Next
XTOTAL2.Caption = Format(nTotal2, "#0.00")
xTotalQuant.Caption = Myvalue(nTotalQuant, "#0")
xTotalQuant2.Caption = Myvalue(nTotalQuant2, "#0")
xTotalItemNoDiscount.Caption = Format(nTotalItemNoDiscount, "#0.00")
xTotalItemDiscount.Caption = Format(nTotalItemDiscount, "#0.00")
xTotalItem.Caption = Format(nTotalItem, "#0.00")

If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xRate.text = Myvalue(Round((Val(xDiscount.text) / Val(xTotalItem.Caption)) * 100, nRound))
    End If
Else
    xRate.text = ""
End If

xTotal.text = Round(nTotalItem - Val(xDiscount.text), 2)
xlate.Caption = Format(Val(xTotal.text) - (Val(xcash.Caption) + Val(xvisa.Caption) + Val(XPOINT.Caption)), "#0.00")
End With
End Function
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray
Dim GrdArray(9, 1)
Set Generalarray(0) = Me
If cBranch = "00" Then
    Generalarray(1) = "SELECT   FILE6_20H.DOC_NO , FILE6_20H.DOC_NO2, FILE6_20H.TRANS_DOC, CONVERT(VARCHAR(10), FILE6_20H.date, 111) , FILE0_50.DESCA, CONVERT(VARCHAR(10), FILE6_20H.time, 108) AS Expr2, FILE3_10.DESCA  , SUM(FILE6_20.TOTAL) AS total , file0_40.desca  , file6_20h.ONLINE_DOC" & _
                      " FROM    FILE6_20H LEFT JOIN  FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE LEFT OUTER JOIN    FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE  INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE WHERE FILE6_20H.ISINVOICE = 0 "
Else
'                                   0               1                       2                   3                                       4                   5                   6
    Generalarray(1) = "SELECT   FILE6_20H.DOC_NO , FILE6_20H.DOC_NO2, FILE6_20H.INV_NO, CONVERT(VARCHAR(10), FILE6_20H.date, 111) , FILE0_50.DESCA, CONVERT(VARCHAR(10), FILE6_20H.time, 108) AS Expr2, FILE3_10.DESCA  , SUM(FILE6_20.TOTAL) AS total , file0_40.desca , file6_20h.ONLINE_DOC" & _
                      " FROM    FILE6_20H LEFT JOIN  FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE LEFT OUTER JOIN    FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE WHERE FILE6_20H.ISINVOICE = 0 "
End If
If lIsBranchStore Then Generalarray(1) = Generalarray(1) & " AND FILE6_20H.BRANCH = " & MyParn(cBranch)
If cBranch = "00" Then
    If lNotBranch Then
'        Generalarray(1) = Generalarray(1) & " AND FILE0_40.ISSTOREBRANCH = 1 "
        Generalarray(1) = Generalarray(1) & " AND FILE6_20H.branch <> '00' "
    Else
        Generalarray(1) = Generalarray(1) & " AND FILE6_20H.branch = '00' "
    End If
End If

If cBranch <> "00" Or lNotBranch Then
    If IsDate(xdate_open.text) Then
        cFilter = cFilter & " AND  FILE6_20H.date = " & DateSq(Format(xdate_open.text, "DD-MM-YYYY"))
        cFilterLook = cFilterLook & " AND FILE6_20H.date = " & DateSq(Format(xdate_open.text, "DD-MM-YYYY"))
    ElseIf IsDate(dSalesDate) Then
        cFilter = cFilter & " AND FILE6_20H.date = " & DateSq(dSalesDate)
        cFilterLook = cFilterLook & " AND FILE6_20H.date = " & DateSq(dSalesDate)
    End If
    If pFilter <> "" Then Generalarray(1) = Generalarray(1) & pFilter
    If cFilterLook <> "" Then Generalarray(1) = Generalarray(1) & cFilterLook
End If

If cBranch = "00" Then
    Generalarray(2) = "GROUP BY FILE6_20H.DOC_NO, FILE3_10.DESCA , FILE6_20H.DOC_NO2, FILE6_20H.TRANS_DOC, CONVERT(VARCHAR(10), FILE6_20H.date, 111), FILE0_50.DESCA, CONVERT(VARCHAR(10),  FILE6_20H.time, 108) , file0_40.desca , file6_20h.ONLINE_DOC Order by doc_no2 DESC "
Else
    Generalarray(2) = "GROUP BY FILE6_20H.DOC_NO, FILE3_10.DESCA , FILE6_20H.DOC_NO2, FILE6_20H.INV_NO, CONVERT(VARCHAR(10), FILE6_20H.date, 111), FILE0_50.DESCA, CONVERT(VARCHAR(10),  FILE6_20H.time, 108) , file0_40.desca , file6_20h.ONLINE_DOC Order by doc_no2 DESC "
End If
Generalarray(3) = 6000
Generalarray(5) = False

If bopt3 Then
   ReDim listarray(3, 5)
   listarray(0, 0) = "«·—ﬁ„-«·Œ“‰…"
   listarray(0, 1) = "(@@Doc_No2@@5 or %%file0_50.DESCA%% or %%file3_10.DESCA%%) "

   listarray(1, 0) = "«· «—ÌŒ"
   listarray(1, 1) = "(##DATE##)"


    If cBranch = "00" Then
        listarray(2, 0) = "„” ‰œ «· ÕÊÌ·"
        listarray(2, 1) = "(%%TRANS_DOC%%)"
    Else
        listarray(2, 0) = "—ﬁ„ «·«Ì’«·"
        listarray(2, 1) = "(%%INV_NO%%)"
    End If

   listarray(3, 0) = "«·„Œ“‰"
   listarray(3, 1) = "%%file0_40.DESCA%%"

Else
   ReDim listarray(2, 5)

    listarray(0, 0) = "«·—ﬁ„"
    listarray(0, 1) = "(@@Doc_No2@@4 or %%file3_10.DESCA%%)"
   
    If cBranch = "00" Then
        listarray(1, 0) = "„” ‰œ «· ÕÊÌ·"
        listarray(1, 1) = "(%%TRANS_DOC%%)"
    Else
        listarray(1, 0) = "—ﬁ„ «·«Ì’«·"
        listarray(1, 1) = "(%%INV_NO%%)"
    End If
   
   listarray(2, 0) = "«·„Œ“‰"
   listarray(2, 1) = "%%file0_40.DESCA%%"

End If

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(1, 1) = 1200

GrdArray(2, 0) = "„” ‰œ  ÕÊÌ· "
If cBranch = "00" Then
    GrdArray(2, 1) = 1000
Else
    GrdArray(2, 1) = 0
End If
GrdArray(3, 0) = "«· «—ÌŒ"
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "«·Õ“‰…"
GrdArray(4, 1) = 2000

GrdArray(5, 0) = "«·Êﬁ "
GrdArray(5, 1) = 1300

GrdArray(6, 0) = "«·⁄„Ì·"
GrdArray(6, 1) = 2500

GrdArray(7, 0) = "«·ﬁÌ„…"
GrdArray(7, 1) = 1300

GrdArray(8, 0) = "«·„Œ“‰"
GrdArray(8, 1) = 2000

GrdArray(9, 0) = "ONLINE"
GrdArray(9, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.sId = "sales" & IIf(bopt3, "1", "0")
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.text = "" Then Cancel = True
End Sub


Private Sub xdoc_ret_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmd_showdoc_Click
End Sub

Private Sub xMan_KeyPress(KeyAscii As Integer)
Dim sCodEmp As String
If KeyAscii = 13 Then
    If Not xMan.MatchedWithList Then
        sCodEmp = xMan.text
        xMan.BoundText = RetZero(xMan.text, 4)
        If Not xMan.MatchedWithList Then
            xMan.BoundText = GetDesca("SELECT CODE FROM FILE6_25 WHERE ID_EMP = " & MyParn(sCodEmp), con)
        End If
    End If
End If
End Sub
Private Sub XONLINE_DOC_Click()
Dim cAddre As String

If ((xClosed.Value <> 0 Or lMainShow Or lSupperVisor) And cBranch <> "00") Then
    On erorr GoTo myerror
    XONLINE_DOC.Caption = InputBox(" —ﬁ„ «Ê—œ— «Ê‰ ·«Ì‰ ", , XONLINE_DOC.Caption)
    If MsgBox("  ⁄œÌ· —ﬁ„ «Ê—œ— «Ê‰ ·«Ì‰ ··›« Ê—… ", vbYesNo + vbDefaultButton2) = vbYes Then
        Dim con_SERVER As New Connection
        If openCon(con_SERVER, LoadConString_OnLine) = "ok" Then
            con_SERVER.Execute " UPDATE FILE6_90H SET SHIP_NO = " & addstring(XSHIP_NO.Caption) & " , sales_doc = " & addstring(xDoc_No.text) & " , sales_date = " & addDate(xDate.text) & " WHERE doc_no = " & MyParn(XONLINE_DOC.Caption), nRec
        
            If nRec > 0 Then
                aRet = aGetDesca("select name , Shipping_City , STREET , phone from file6_90h where doc_no = " & MyParn(XONLINE_DOC.Caption), con_SERVER)
                If UBound(aRet) > 0 Then
                    cAddre = aRet(2) + "   " + aRet(3)
                    con.Execute " UPDATE  FILE6_20H SET file6_20h.NAME = " & addstring(aRet(1)) & " , file6_20h.ADDRESS = " & addstring(cAddre) & " ,  file6_20h.PHONE = " & addstring(aRet(4)) & "   WHERE file6_20h.doc_no = " & MyParn(xDoc_No.text)
                End If
            
            Else
                MsgBox " —ﬁ„ ÿ·»Ì… €Ì— ’ÕÌÕ "
            End If
        End If
        con_SERVER.Close
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub XPHONE_Click()
    PhoneCust.Show 1
End Sub
Private Sub xRate_LostFocus()
myLostFocus xRate
If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Or xDiscount.Locked Then
        xDiscount.text = Round((Val(xRate.text) * Val(xTotalItem.Caption)) / 100, 0)
    End If
Else
    xDiscount.text = ""
End If
CalcTotals
End Sub
Private Function RetItemBalance(cItem, cStore, dDate) As Double
If cItem = "" Then Exit Function
movetable.Seek Array(cItem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> cItem Or cStore <> movetable!STORE Or DateValue(movetable!Date) > DateValue(Format(dDate, "dd-mm-yyyy")) Then Exit Do
    RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    movetable.MoveNext
Loop
End Function
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub Fixgrd()

With grid1
.Cols = 22
'                0        1           2          3          4           5           6           7       8              9           10          11           12            13          14                    15              16          17
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄— „” Â·ﬂ|" & "⁄œœ|" & "”⁄— «·»Ì⁄|" & "‰”»… Œ’„|" & "«·≈Ã„«·Ï|" & "”⁄— √Êﬂ«“ÌÊ‰|" & "ﬂÊœ «·»«∆⁄|" & "«·»«∆⁄|" & "—ﬁ„ «·„ÊœÌ·|"
.RowHeight(0) = 900
.WordWrap = True
.ColHidden(.Cols - 1) = True
.ColHidden(14) = True
.ColHidden(18) = True
'.ColHidden(16) = True
.ColHidden(17) = True
.ColWidth(17) = 2000
.ColWidth(0) = 600
.ColWidth(1) = 1000
.ColWidth(2) = 0
.ColWidth(3) = 1300
.ColWidth(4) = 0
.ColWidth(5) = 1500
.ColWidth(6) = 2800
.ColWidth(7) = 700
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 800
.ColWidth(11) = 1000
.ColWidth(12) = 500
.ColWidth(13) = 1200
.ColWidth(15) = 1700
.ColWidth(16) = 0

.ColHidden(19) = True
.ColHidden(20) = True
If cBranch = "00" And Not lNotBranch Then
    .ColHidden(12) = True
    .ColHidden(14) = True
'   .ColHidden(9) = True
End If

'.ColFormat(12) = "#.##%"
'.ColComboList(3) = clist1

If cBranch > "60" Then
    .ColComboList(15) = StrList("SELECT * FROM FILE6_25 WHERE ISSTOP = 0 ORDER BY DESCA ", con)
Else
    If lNotBranch Then
        .ColComboList(15) = StrList("SELECT * FROM FILE6_25  ORDER BY DESCA ", con)
    Else
        .ColComboList(15) = StrList("SELECT * FROM FILE6_25 WHERE ISSTOP = 0 AND ( BRANCH = " & MyParn(cBranch) & " OR CODE = '0001' ) ORDER BY DESCA ", con)
    End If
End If
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter

.ShowCell .Rows - 1, 0


'''''''If cBranch = "00" And XISONEST.Value <> 0 Then
'''''''    For nRow = 1 To .Rows - 1
'''''''        If .ValueMatrix(nRow, 9) <> .ValueMatrix(nRow, 11) Then
'''''''            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
'''''''        End If
'''''''    Next nRow
'''''''End If

'If Not bopt3 And DateValue(Date) <> DateValue(xDate.Text) Then
'    grid1.Visible = False
'    Picture3.Visible = False
'    Picture1.Visible = False
'Else
'    grid1.Visible = True
'    SSCommand2.Visible = True
'    Picture3.Visible = True
'    Picture1.Visible = True
'End If
End With
End Sub
Private Sub CLIENTLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From file3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCl1.Caption = "«” ⁄·«„"
oSearchCl1.Show 1
End Sub
Private Sub CLIENTLOOKUP2()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select BARCODE , DescA,phone1,fax,ADDRESS From file3_20 "
Generalarray(2) = " Order by file3_20.Desca"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·«”„-«· ·Ì›Ê‰-«·„Ê»Ì·-«·⁄‰Ê«‰"
listarray(0, 1) = "(%%DESCA%% OR %%PHONE1%% OR %%FAX%% OR %%ADDRESS%%)"

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "«· ·Ì›Ê‰"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·›«ﬂ”"
GrdArray(3, 1) = 1300

GrdArray(4, 0) = "«·⁄‰Ê«‰"
GrdArray(4, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCl2.Caption = "«” ⁄·«„ ⁄„·«¡ «·‰ﬁœÌ"
oSearchCl2.Show 1
End Sub
Private Sub CLIENTLOOKUP3()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From file4_10"
Generalarray(2) = "Order by file4_10.Desca"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·„Ê—œ"
GrdArray(1, 1) = 5000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCl3.Caption = "«” ⁄·«„ «·„Ê—œÌ‰"
oSearchCl3.Show 1
End Sub


Private Sub xRateDis_Lostfocus()
    xDiscount.text = Fix((Val(xTotalItem.Caption) * Val(xRateDis.text) / 100))
End Sub

Private Sub xRet1_Change()

End Sub


Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xStore_Validate(Cancel As Boolean)
    If Trim(xStore.BoundText) = "" Then Cancel = True
End Sub
Private Sub xTax_LostFocus()
    xTax.BackColor = &H80000005
    CalcTotals
End Sub
Private Function RemoveItem(nRow) As Boolean
Dim cPassword As String, aUser As Variant
On Error GoTo myerror
'con.BeginTrans
If grid1.TextMatrix(nRow, grid1.Cols - 1) <> "" Then
    
    ManagPass.pPassWord = ""
    ManagPass.Show 1
    cPassword = ManagPass.pPassWord
    
    If Trim(cPassword) = "" Then Exit Function
    aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
    If UBound(aUser) = 0 Then Exit Function
    
    con.Execute "Delete  From FILE6_20 where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
'    For I = nRow + 1 To grid1.Rows - 2
'        con.Execute "update " & cFile & " set [row] = " & (I - 1) & " where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
'    Next
End If
grid1.RemoveItem nRow
'con.CommitTrans
Exit Function
myerror:
'con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myproc2(nDoc_no)

xDoc_No = nDoc_no
myUndo
'CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
'If Not CardTable.EOF Then
'    myload
'Else
'    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
'    Unload Me
'End If
End Sub
Function itemPrice(cItem) As Single
    itemPrice = GetDesca("select PRICE from file1_10 where item = " & MyParn(cItem), con)
End Function
Private Function myreplaceGrd(nRow) As Boolean
Dim aInsert(12, 1)
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        If Val(.TextMatrix(i, 10)) = 0 Or (Val(.TextMatrix(i, 10)) > 100 And cBranch <> "00") Then
            .TextMatrix(i, 10) = ""
            MsgBox "—«Ã⁄ ⁄œœ «·„»Ì⁄« "
            Exit Function
        End If
        If Len(grid1.TextMatrix(i, 1)) <= 2 Then
            MsgBox " ’‰› ⁄Ì— ’ÕÌÕ"
            Exit Function
        End If
        If Val(.TextMatrix(i, 11)) = 0 Then
            .TextMatrix(i, 11) = Val(.TextMatrix(i, 9))
        End If
        
        If Val(grid1.TextMatrix(i, 11)) = 0 Then
            MsgBox " ”⁄— ’‰› ’›— "
            Exit Function
        End If
        
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 1))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = Val(.TextMatrix(i, 10))

        aInsert(3, 0) = "Price"
        aInsert(3, 1) = Val(.TextMatrix(i, 11))

        aInsert(4, 0) = "Discount"
        aInsert(4, 1) = 0

        aInsert(5, 0) = "PRICE_C"
        aInsert(5, 1) = Val(.TextMatrix(i, 9))

        aInsert(6, 0) = "row"
        aInsert(6, 1) = i
        
        aInsert(7, 0) = "COST"
        aInsert(7, 1) = Item_cost(.TextMatrix(i, 1), con)

        aInsert(8, 0) = "MAN"
        aInsert(8, 1) = addstring(.TextMatrix(i, 15))

        aInsert(9, 0) = "S_OKAZ"
        aInsert(9, 1) = Val(.TextMatrix(i, 12))

        aInsert(10, 0) = "USER_IP"
        aInsert(10, 1) = addstring(cIpName)

        aInsert(11, 0) = "Price_C2"
        aInsert(11, 1) = Val(.TextMatrix(i, 19))

        aInsert(12, 0) = "S_OKAZ_2"
        aInsert(12, 1) = Val(.TextMatrix(i, 20))

        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "FILE6_20")
        Else
            con.Execute CreateUpdate(aInsert, "FILE6_20", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
    Dim cString As String
    Dim cPrice As String
    cField1 = "case when file6_20.Discount = 0 then Null else file6_20.Discount end "
    cPrice = " FILE6_20.PRICE_C"
    
'                           1          2                3                4                   5               6             7               8                9                  10              11               12               13                 14                  15                  16                  17              18                      19                  20          21
    cString = "SELECT FILE6_20.ITEM, FILE1_10.MOSM, FACT.DESCA , FILE1_10.SUPP, FILE1_10.MODELFACT0, FILE1_10.DESCA, FILE1_10.SCAL , FILE1_10.COLOR , " & cPrice & " , file6_20.Quant, file6_20.Price, FILE6_20.S_OKAZ ,FILE6_20.TOTAL  ,   FILE6_20.PRICE_C2 ,   FILE6_20.MAN    ,FILE6_25.DESCA ,   FILE1_10.MODELNO   , FILE1_10.BARCODE13   ,  FILE6_20.PRICE_C2 , FILE6_20.S_OKAZ_2,ID " & _
          " FROM (FILE6_20 LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM) LEFT JOIN FACT ON FACT.CODE = FILE1_10.FACT LEFT JOIN FILE6_25 ON FILE6_20.MAN  = FILE6_25.CODE "
    cString = cString & turn(cString) & " DOC_NO = " & MyParn(xDoc_No.text)
    cString = cString & " ORDER BY FILE6_20.ID"
    Set DATA11.Recordset = myRecordSet(cString, con)
    myaddItem
End With
CalcTotals
Fixgrd
grid1.ShowCell grid1.Rows - 1, 1
If cBranch = "00" Then
    grid1.ShowCell grid1.Row, 1
    grid1.Select grid1.Row, 1
End If
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xDate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xRate_GotFocus()
myGotFocus xRate
End Sub
Private Sub xRateDis_GotFocus()
xRateDis.SelStart = 0
xRateDis.SelLength = Len(xRateDis.text)
End Sub
Private Sub xDiscount_GotFocus()
myGotFocus xDiscount
End Sub
Private Sub xTax_GotFocus()
xTax.SelStart = 0
xTax.SelLength = Len(xTax.text)
End Sub
Private Function mysave(Optional bEnd As Boolean = True, Optional bPrint As Boolean = True) As Boolean
If Not MYVALID Then Exit Function
CalcTotals
If Not myreplace Then Exit Function
con.Execute " update file6_20 set item = item where doc_no = " & MyParn(xDoc_No.text)
If bEnd And lSave Then
    SavePrint
    If XONLINE(2).Value Then
'       If doprint(False) Then SavePrint
        If doprint_online() Then SavePrint
        Exit Function
    Else
        If xCode.text = "0000" Then
            Inform_OK " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ" & Chr(13) & " «·»«ﬁÏ " & Format(xRest.Caption, "#0.00")
        Else
            Inform_OK " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ" & Chr(13) & "›« Ê—… »Ì⁄ √Ã· ··⁄„Ì· " & xCodeDesca.Caption
        End If
        If cBranch <> "00" Then
            If Val(xTotal.text) <> Val(xvisa.Caption) + Val(xcash.Caption) + Val(XPOINT.Caption) Then
                MsgBox " »—Ã«¡ „—«Ã⁄…  ”ÃÌ· «·”œ«œ "
                Exit Function
            End If
            PhoneCust.Show 1
            If chkprint.Value = 0 Then
                If doprint(False) Then SavePrint
            Else
                SavePrint
            End If
        End If
    End If
'        ”Ì· —ﬁ„ «·»«—ﬂÊœ ·ﬂ«—  «·Œ’„
    If lIsBranchStore And xcard_disc.text <> "" Then
        con.Execute " UPDATE CARD_DISC SET SALES_DOC_NO = " & addstring(xDoc_No.text) & " WHERE BARCODE_DISC = " & MyParn(xcard_disc.text)
    Else
        If xcard_disc.text <> "" Then SAveCardDisc_No
    End If
    
    myDefine
Else
    myUndo
End If
End Function
Private Sub SavePrint()
On Error GoTo myerror
'con.BeginTrans
con.Execute "update file6_20h set isnew = 1 , FILE6_20H.PRINTED = 1 , ISCLOSED = 1  WHERE DOC_NO = " & MyParn(xDoc_No.text)
'con.CommitTrans
xPrinted.Value = 1
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Function doprint(lRePrint As Boolean)
Dim loctable As New ADODB.Recordset, cString As String
Dim n As Double, nD As Double
cString = "SELECT FILE6_20.PRICE_C, FILE6_20.ITEM,FILE6_20.QUANT,FILE6_20.PRICE,FILE1_10.DESCA AS ITEM_DESCA,FILE6_20.TOTAL,FILE6_20.MAN,FILE6_20H.TIME,FILE1_10.DESCA , file6_20.man , file1_10.modelfact , FACT.DESCA AS FACTDESCA FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO  INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM LEFT JOIN FACT ON FACT.CODE = FILE1_10.[FACT]  "
cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
contemp.Execute "DELETE * FROM TEMP"
If temptable.State = adStateOpen Then temptable.Close
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Dim sMan As String
Dim cVisaDoc As String, cVisaDoc3 As String, nVisa1 As Double, nVisa3 As Double
If cBranch <> "00" Then AddSales_Def
'                           1   2       3       4           5       6           7                           8
aRet = aGetDesca("SELECT VISA,VISA3,VisaDoc , VisaDoc3 , VISADOC ,VISADOC3, VISA.DESCA AS NAMEVISA1 , VISA_2.DESCA AS NAMEVISA3   FROM FILE6_20H LEFT JOIN VISA ON VISA.CODE = FILE6_20H.CODEVISA LEFT JOIN VISA AS VISA_2 ON VISA_2.CODE = FILE6_20H.CODEVISA3  WHERE DOC_NO = " & MyParn(xDoc_No.text), con)
If UBound(aRet) > 0 Then
    cVisaDoc = "—ﬁ„ ⁄„·Ì… " & aRet(7) & " " & aRet(3)
    cVisaDoc3 = "—ﬁ„ ⁄„·Ì… " & aRet(8) & " " & aRet(4)
    nVisa1 = Val(aRet(1) & "")
    nVisa3 = Val(aRet(2) & "")
End If
With loctable
Do Until loctable.EOF
    temptable.AddNew
    If cBranch <> "00" Then
        temptable!str13 = TurnValue(cComp_Name)
        temptable!str11 = TurnValue(cComp_address)
        temptable!str12 = "Phone : " & TurnValue(cComp_Phone)
    Else
        temptable!str13 = TurnValue(xStore.text)
'        temptable!str11 = TurnValue(cComp_address)
 '        temptable!str12 = "Phone : " & TurnValue(cComp_Phone)
    End If
    If nVisa1 <> 0 Then
        temptable!str17 = TurnValue(cVisaDoc)
        temptable!VAL13 = Val(nVisa1)
    End If
    If nVisa3 <> 0 Then
        temptable!str22 = TurnValue(cVisaDoc3)
        temptable!val18 = Val(nVisa3)
    End If
    
'    temptable!str17 = TurnValue(XVISADOC.Caption)
    temptable!str4 = TurnValue(cId_NoDesca1)
    temptable!STR19 = TurnValue(cId_NoDesca2)

    temptable!str16 = loctable!FACTDESCA & " " & DelZero(loctable!modelfact)
    
    temptable!str8 = TurnValue(xBox.text, "", Null)
    temptable!Date1 = Format(xDate.text, "YYYY-MM-DD")
    temptable!date2 = Format(xtime.Caption, "HH:NN")
    temptable!str3 = xDoc_no2.text
    temptable!STR20 = xDoc_No.text
    temptable!str5 = loctable!Item
    temptable!VAL1 = Val(loctable!Quant & "")
    temptable!STR21 = TurnValue(cHeadBoon)
    If (Val(loctable!price & "") = Val(loctable!PRICE_C & "")) Or XONLINE(2).Value <> 0 Then
        temptable!val2 = Val(loctable!price & "")
        temptable!val3 = Val(loctable!TOTAL & "")
    Else
        temptable!str2 = "Price after Discount"
        temptable!val2 = Val(loctable!PRICE_C & "")
        temptable!VAL4 = Val(loctable!price & "")
        temptable!val5 = Val(loctable!TOTAL & "")
    End If
    If XONLINE(2).Value = 0 Then
        If Val(xDiscount.text) <> 0 Then
            If lIsdisc2 Then
                temptable!str10 = TurnValue(xNOTE_DISC.Caption, "", Null) & " Buy 2 GET 1 free"
                temptable!str14 = "«·≈” »œ«· Ê«·„— Ã⁄ ·⁄—÷ ﬁÿ⁄ Ì‰ Ê ﬁÿ⁄… „Ã«‰«  ”«—Ï Œ·«· › —… «·⁄—÷ ›ﬁÿ "
            Else
                temptable!str10 = TurnValue(xNOTE_DISC.Caption, "", Null)
            End If
        End If
    Else
    End If
    If xcard_disc.text <> "" Then temptable!str10 = TurnValue(xcard_desca.Caption & "")
    
    temptable!Val21 = Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption)
    
    
    If XONLINE(2).Value = 0 Then
        temptable!val12 = Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption)
        temptable!VAL6 = Val(xTotalItemNoDiscount.Caption)
        temptable!VAL7 = Val(xTotalItemDiscount.Caption)
        temptable!Val8 = Val(xDiscount.text)
        temptable!val9 = Val(xTotal.text)
        temptable!Val10 = Val(xPay.Caption)
        temptable!val11 = Val(xRest.Caption)
'        temptable!VAL13 = Val(xvisa.Caption)
        temptable!VAL14 = Val(XPOINT.Caption)
    Else
        temptable!val12 = Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption)
        temptable!VAL6 = Val(xTotalItem.Caption)
        temptable!VAL7 = 0
        temptable!Val8 = Val(xDiscount.text)
        temptable!val9 = Val(xTotal.text)
        temptable!Val10 = Val(xPay.Caption)
        temptable!val11 = Val(xRest.Caption)
        temptable!VAL14 = Val(XPOINT.Caption)
        temptable!VAL13 = Val(xvisa.Caption)
    End If
    If XONLINE(2).Value Then
        aRet = aGetDesca("SELECT NAME , PHONE , ADDRESS , CHARGE1 , CHARGE2 , SHIP_NO , SHIP FROM FILE6_20H WHERE DOC_NO = " & MyParn(xDoc_No.text), con)
        temptable!str6 = TurnValue(ArbString(GetDesca("SELECT DESCA FROM SHIP WHERE CODE = " & MyParn(aRet(7)), con)))
        temptable!str15 = TurnValue((aRet(6)))
        temptable!STR7 = TurnValue(aRet(1))
        temptable!str9 = TurnValue(aRet(2))
        temptable!str18 = TurnValue(ArbString(aRet(3)))
        temptable!val15 = Val(aRet(4)) + Val(aRet(5))
        temptable!val16 = Val(aRet(4)) + Val(aRet(5)) + Val(xTotal.text)
    End If
    If lRePrint Then temptable!str1 = "C  O  P  Y"
    
    temptable.Update
    loctable.MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
On Error Resume Next
If Not lNotBranch Then
    REPORT1.Reset
    FixPrinter REPORT1
End If
If XGIFT.Value = 1 Then
    REPORT1.ReportFileName = App.Path & "\Reports\sales_G.rpt"
Else
    REPORT1.ReportFileName = App.Path & "\Reports\sales.rpt"
End If
REPORT1.DataFiles(0) = tempFile
If lRePrint Then
    REPORT1.CopiesToPrinter = 1
Else
    REPORT1.CopiesToPrinter = 1 ' nCountPrint
End If
If Not lNotBranch Then REPORT1.Destination = crptToPrinter
If lNotBranch Then REPORT1.Destination = crptToWindow
REPORT1.ProgressDialog = False
REPORT1.Action = 1
doprint = True
GoTo closeCon


Exit Function

myerror:
MsgBox Err.Description
Err.Clear
closeCon:
temptable.Close
Set temptable = Nothing
End Function
Private Sub myLoadSetting()
    chkprint.Value = Val(RetSetting("print", App.Path & "\other.txt"))
    ISPICT.Value = Val(RetSetting("ISPICT", App.Path & "\other.txt"))
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not MYVALID(bIgMsg) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 10)) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 11)) Then Exit Function
If Trim(grid1.TextMatrix(Row, 15)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub HandleCntEdit()
xDoc_No.Tag = LoadMode
xDoc_No.Enabled = False
CmdDel.Enabled = True
cmdSave.Enabled = (bedit) And xPrinted.Value = 0 And grid1.Rows > 2
cmd_disc.Enabled = (bedit) And xPrinted.Value = 0 And grid1.Rows > 2
cmdAddItems.Enabled = (bedit) And xPrinted.Value = 0 And grid1.Rows > 2
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col = 1 Then
    If grid1.TextMatrix(Row, 1) = "" Then Exit Sub
    If Row = grid1.Rows - 1 Then
        grid1.Row = grid1.Rows - 1
        grid1.Col = 1
        grid1.ShowCell grid1.Rows - 1, 1
    Else
        If Row < grid1.Rows - 1 Then
            grid1.Row = Row + 1
            grid1.Col = 1
            grid1.ShowCell grid1.Row, 1
        End If
    End If
ElseIf Col = 10 Or Col = 11 Then
    If Row < grid1.Rows - 1 Then
        grid1.Row = Row + 1
        grid1.Col = 1
        grid1.ShowCell grid1.Row, 1
    End If
End If
End Sub
Private Sub myaddItem()
grid1.AddItem ""
If grid1.Rows > 2 Then
    grid1.TextMatrix(grid1.Rows - 1, 15) = grid1.TextMatrix(grid1.Rows - 2, 15)
    grid1.TextMatrix(grid1.Rows - 1, 16) = grid1.TextMatrix(grid1.Rows - 2, 16)
End If
MakeSerial
End Sub
Private Sub xTotal_LostFocus()
myLostFocus xTotal
xDiscount.text = Round(Val(xTotalItem.Caption) - Val(xTotal.text), 2)
CalcTotals
End Sub
Private Sub xDoc_no2_GotFocus()
myGotFocus xDoc_no2
End Sub
Private Sub xDoc_no2_LostFocus()
myLostFocus xDoc_no2
End Sub



Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xMosm
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xMosm
End Sub
Sub PostData()
If DateValue(salesFrm.xDate.text) = DateValue(dSalesDate) Then
    Dim dDate As String
    dDate = DateAdd("d", 1, dSalesDate)
    If DateDiff("d", dSalesDate, Date) < -1 Then
        MsgBox " «—ÌŒ «·ÃÂ«“ ·« Ì”„Õ » —ÕÌ·  «—ÌŒ «·„»Ì⁄« "
        Exit Sub
    End If
    If IsDate(dDate) Then
        CD = "”Ê› Ì „  €Ì— «· «—ÌŒ «·Õ«·Ï ≈·Ï " & Format(dDate, "DD-MM-YYYY")
        If MsgBox(CD, vbOKCancel + vbDefaultButton2) = vbOK Then
            dSalesDate = dDate
            con.Execute " UPDATE FILE8_50H SET isclosed = 1 WHERE ISCLOSED = 0 AND BRANCH = " & MyParn(cBranch)
            con.Execute " UPDATE FILE8_60H SET isclosed = 1 WHERE ISCLOSED = 0 AND BRANCH = " & MyParn(cBranch)
                        
            If lIsBranchStore Then
                con.Execute " UPDATE FILE0_52  SET ISCLOSED = 1 WHERE ISCLOSED = 0 AND BRANCH = " & MyParn(cBranch)
                con.Execute " UPDATE DSALES SET DSALES = " & addDate(dSalesDate) & " WHERE BRANCH = " & MyParn(cBranch)
            Else
                con.Execute " UPDATE FILE0_51  SET ISCLOSED = 1 WHERE ISCLOSED = 0 "
                con.Execute " UPDATE FILE8_20H SET ISCLOSED = 1 WHERE ISCLOSED = 0 "
                con.Execute " UPDATE FILE8_70H SET ISCLOSED = 1 WHERE ISCLOSED = 0 AND BRANCH = " & MyParn(cBranch)
                con.Execute " UPDATE DSALES SET DSALES = " & addDate(dSalesDate)
            End If
            Firsttitle = cComp_Name & Format(dSalesDate, "DD-MM-YYYY")
            Main.Caption = Firsttitle
            salesFrm.Caption = Format(dSalesDate, "DD-MM-YYYY")
            MsgBox " „  €Ì— «· «—ÌŒ"
        End If
    End If
End If
End Sub
Private Function CheckOpen(lMsg As Boolean, Optional ByVal pBox As String = "") As Long
Dim cString As String, nCount As Long
cString = "Select count(*) from file6_20h"
cString = cString & turn(cString) & "File6_20h.printed = 0"
cString = cString & turn(cString) & " [DATE] = " & DateSq(dSalesDate)
If lMsg Then cString = cString & turn(cString) & " online = 0 "
If pBox <> "" Then cString = cString & turn(cString) & " [BOX] = " & MyParn(pBox)
If lIsBranchStore Then
    cString = cString & turn(cString) & " BRANCH  = " & MyParn(cBranch)
End If
nCount = Val(GetDesca(cString, con))
If nCount > 0 Then
    If lMsg Then
        MsgBox "Â‰«ﬂ ⁄œœ " & nCount & " »Ê‰«  »Ì⁄ „› ÊÕ…!!«·—Ã«¡ «·Õ–› «Ê «· ”ÃÌ·", vbCritical
    End If
    CheckOpen = nCount
End If
End Function
Function MyRet5(nVal)
    Dim nRet As Double
    MyRet5 = nVal
    nRet = nVal - Fix(nVal)
    If nRet > 0.5 Then MyRet5 = Fix(nVal) + 1
    If nRet < 0.5 And nRet > 0 Then MyRet5 = Fix(nVal) + 0.5
End Function

Function GetDescafact(pString) As String
Dim loctable As Recordset
Set loctable = Factdb.OpenRecordset(pString)
If loctable.RecordCount >= 0 Then GetDescafact = loctable(0) & ""
loctable.Close
Set loctable = Nothing
End Function
Private Function CalcY()
    CalcY = ((nRow - 1) * nCardHeight) + nUpMargin + temptable!Top
End Function
Private Function Calcx()
    Calcx = nLeftMargin + temptable!Left + ((nCol - 1) * nCardWidth)
End Function


Private Sub Timer1_Timer()
'    myDisplay cMsgDisplay, " "
'    Timer1.Enabled = False
'dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES"), "DD-MM-YYYY")
'If Format(dSalesDate, "DD-MM-YYYY") <> Format(xDate.Text, "DD-MM-YYYY") Then
'    MsgBox " „  €Ì— «· «—ÌŒ"
'End If
'xDate.Text = Format(dSalesDate, "DD-MM-YYYY")
End Sub
Sub opencom()
On Error GoTo myerror
Dim i As Integer
Dim comFlag As Boolean
If cComPort = "" Then Exit Sub
  comFlag = False
    mOpenUSBpd
  IntrCh23$ = Chr$(&H23)
  IntrCh24$ = Chr$(&H24)
  IntrCh40$ = Chr$(&H40)
  IntrCh5B$ = Chr$(&H5B)
  IntrCh5C$ = Chr$(&H5C)
  IntrCh5D$ = Chr$(&H5D)
  IntrCh5E$ = Chr$(&H5E)
  IntrCh60$ = Chr$(&H60)
  IntrCh7B$ = Chr$(&H7B)
  IntrCh7C$ = Chr$(&H7C)
  IntrCh7D$ = Chr$(&H7D)
  IntrCh7E$ = Chr$(&H7E)
   lDisplayCom = True
  Exit Sub
myerror:
lDisplayCom = False
  Err.Clear
End Sub
Sub ClearDisplay()
    On Error GoTo myerror
    Dim i, Pagei, nn As Integer
    Dim ss As String
    Dim TmpTime$
    Dim TmpMM, TmpHH, TmpSS
    If cComPort = "" Then Exit Sub
    If Not lDisplayCom Then Exit Sub
    DoEvents
    Call Epson_Init
    ss = "CLEAR LINE1" + Chr$(&HA)
    ss = ss + "and LINE2"
    DoEvents
'    Sleep (1000)
    ss = Chr$(&HC)
    DoEvents
    Exit Sub
myerror:
    Err.Clear
End Sub
Private Sub Epson_Init()
Dim ss As String
    ss = Chr$(&H1B) + Chr$(&H40)
    DoEvents
    mWritePD ss, Len(ss)
    DoEvents
End Sub
Function mWritePD(mdata As String, ByVal mlength As Long) As Long
    If cComPort = "USB" Then
       mWritePD = WritePD(mdata, mlength)
    ElseIf Left(cComPort, 3) = "COM" Then
       MSComm1.Output = mdata
    End If
End Function
Function mOpenUSBpd() As Long
    MSComm1.CommPort = Mid$(cComPort, 4, 2)
    MSComm1.PortOpen = True
End Function
Sub myDisplay(pLine1, pLine2)
    Dim ss As String
    If cBranch = "00" Then Exit Sub
    If Not lDisplayCom Then Exit Sub
    ClearDisplay
    ss = pLine1
    ss = ss + Chr$(&H1F) + Chr$(&H43) + Chr$(&H0)
    mWritePD ss, Len(ss)
    DoEvents
    ss = vbCrLf + pLine2
    ss = ss + Chr$(&H1F) + Chr$(&H43) + Chr$(&H1)
    mWritePD ss, Len(ss)
    DoEvents
End Sub

Public Sub PrintTDay()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(1)
Dim pDate As Date, pstore As String
Dim nFBalItem As Double
Dim nTInlItem As Double
Dim nSallItem As Double
Dim nBalItem As Double
Dim nOverCash As Double, nVisa2 As Double
Dim VisaTable As New ADODB.Recordset
pDate = xDate.text
pstore = xStore.BoundText
contemp.Execute "DELETE * FROM TEMP"


    temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

    cString = " SELECT [SECTION] ,  SUM(QUANT) AS t_q, SUM(TOTAL * ((100 -rate )/ 100 )) AS t_tot FROM SALES_MODEL "
    cString = cString & " where date = " & DateSq(xDate.text)
    cString = cString & " GROUP BY [SECTION] "
    aHeader(0) = "[" & BetweenString(Format(xDate.text, "d-m-yyyy"), Format(xDate.text, "d-m-yyyy")) & "]"
    
    cString = "SELECT  SUM(T_QUANT) AS T_Q, SUM(T_QUANT1) AS T_Q1, SUM(T_QUANT2) AS T_Q2, SUM(t_total) AS T_Item, SUM(t_total1) AS T_Item1, SUM(t_total2) AS T_Item2, SUM(discount) AS T_Disc, SUM(CASH) AS T_Cash, SUM(VISA+VISA3) AS T_Visa FROM   T_SALESDOC "
    cString = cString & " where date = " & DateSq(xDate.text)
    aHeader(0) = "[" & BetweenString(Format(xDate.text, "d-m-yyyy"), Format(xDate.text, "d-m-yyyy")) & "]"
    
    
    If cManBox <> "" Then
        cString = cString & " AND BOX = " & MyParn(cManBox)
        aHeader(1) = "[" & salesFrm.xBox.text & "]"
    End If
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

cStr1 = " SELECT SUM(VALUE) FROM FILE8_60 INNER JOIN FILE8_60H ON FILE8_60.DOC_NO = FILE8_60H.DOC_NO WHERE DATE = " & DateSq(xDate.text)
If cManBox <> "" Then
    cStr1 = cStr1 & " AND BOX = " & MyParn(cManBox)
End If
nOverCash = Val(GetDesca(cStr1, con) & "")
If Not (sourcetable.EOF And sourcetable.BOF) Then
With sourcetable
    Do While Not sourcetable.EOF
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "⁄œœ „»Ì⁄« "
    
    temptable!val3 = !t_q1
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "⁄œœ „— Ã⁄« "
    
    temptable!val3 = !t_q2
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "’«›Ï ⁄œœ „»Ì⁄« "
    temptable!val3 = !t_q
    temptable.Update

    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "ﬁÌ„… „»Ì⁄« "
    temptable!val3 = !T_Item1
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "ﬁÌ„… „— Ã⁄« "
    temptable!val3 = !T_Item2
    temptable.Update
    
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "’«›Ï ﬁÌ„… „ÊœÌ·« "
    temptable!val3 = !T_Item
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "ﬁÌ„… «·Œ’„"
    temptable!val3 = !T_Disc
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "’«›Ï „»Ì⁄« "
    
    temptable!val3 = Val(!T_Item & "") - Val(!T_Disc & "")
    temptable.Update

    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "”œ«œ ‰ﬁœÏ"
    temptable!val3 = Val(!T_CASH & "")
    temptable.Update
    
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "”œ«œ ›Ì“« "
    temptable!val3 = Val(!T_VISA & "")
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "√Ã· "
    temptable!val3 = Val(!T_Item & "") - Val(!T_Disc & "") - Val(!T_CASH & "") - Val(!T_VISA & "")
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "OVER CASH"
    temptable!val3 = nOverCash
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "≈Ã„«·Ï ‰ﬁœÌ…"
    temptable!val3 = nOverCash + Val(!T_CASH & "")
    temptable.Update
    
''  «Ã„«·Ï ›Ì“«
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "⁄„Ê·…  ﬁ”Ìÿ ›Ì“«"
    cString = "select sum(visa2+visa3_2) from file6_20h where date = " & DateSq(xDate.text)
    If cManBox <> "" Then cString = cString & " AND BOX = " & MyParn(cManBox)
    nVisa2 = Val(GetDesca(cString, con) & "")
    temptable!val3 = nVisa2
    temptable.Update
    
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    If cManBox <> "" Then
        temptable!str1 = TurnValue(cManBox)
        temptable!str2 = TurnValue(GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(cBranchBox), con))
    End If
    temptable!str11 = 3
    temptable!str12 = "≈Ã„«·Ï «·ÌÊ„ "
    temptable!str5 = "’«›Ï ›Ì“« "
    temptable!val3 = nVisa2 + Val(!T_VISA & "")
    temptable.Update
    
''  «Ã„«·Ï ›Ì“«
    
    .MoveNext
    
    Loop
End With
End If
'
cString = "SELECT   FILE0_50.DESCA AS BOXDESCA , SUM(BOXMOVE.PLUS ) AS T_IN, SUM(BOXMOVE.MINUS ) AS T_OUT, BOXMOVE.BOX, BOXMOVE.DESCA as desca, BOXMOVE.FLAG FROM         BOXMOVE INNER JOIN FILE0_50 ON FILE0_50.CODE = BOXMOVE.BOX WHERE  DATE = " & DateSq(salesFrm.xDate.text)
If cManBox <> "" Then cString = cString & " AND BOX = " & MyParn(cManBox)
cString = cString & " GROUP BY FILE0_50.DESCA ,BOXMOVE.BOX, BOXMOVE.DESCA , BOXMOVE.FLAG ORDER BY BOXMOVE.FLAG desc "
sourcetable.Close
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (sourcetable.EOF And sourcetable.BOF) Then
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str11 = 6
    temptable!str12 = "≈Ã„«·Ï «·Õ“‰…"
    temptable!str3 = " ÌÊ„Ì… " & xDate.text
    temptable!str1 = !BOX
    temptable!str2 = TurnValue(!BOXDESCA)
    temptable!str5 = TurnValue(!DESCA)
    
    temptable!val3 = !T_IN + !T_OUT
    temptable!Val10 = 2
    temptable.Update
    sourcetable.MoveNext
Loop
End With
End If
    
'   ›’Ì·Ï «·›Ì“«

cString = "SELECT  SUM(FILE6_20H.VISA) AS TVISA ,SUM(FILE6_20H.VISA3) AS TVISA3 , FILE6_20H.CODEVISA, FILE6_20H.CODEVISA3, VISA.DESCA , VISA3.DESCA AS DESCA3 FROM    FILE6_20H LEFT JOIN VISA ON FILE6_20H.CODEVISA = VISA.CODE  LEFT JOIN VISA AS VISA3 ON FILE6_20H.CODEVISA3 = VISA3.CODE " & _
            " WHERE FILE6_20H.VISA <> 0  AND FILE6_20H.date = " & DateSq(xDate.text)
If cBranchStore <> "" Then cString = cString & " AND STORE = " & MyParn(cBranchStore)
If cManBox <> "" Then cString = cString & " AND BOX = " & MyParn(cManBox)
cString = cString & " GROUP BY FILE6_20H.CODEVISA, VISA.DESCA ,FILE6_20H.CODEVISA3, VISA3.DESCA"
VisaTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (VisaTable.EOF And VisaTable.BOF) Then
    With VisaTable
    Do Until .EOF
        temptable.AddNew
        temptable!str11 = 7
        temptable!str12 = "≈Ã„«·Ï ›Ì“«"
        temptable!str3 = " ÌÊ„Ì… " & xDate.text
        temptable!str1 = cManBox
'       temptable!str2 = TurnValue(!BOXDESCA)
        temptable!str5 = TurnValue(!DESCA)
        temptable!val3 = !TVISA
        temptable!Val10 = 2
        temptable.Update
        
        If !TVISA3 <> 0 Then
        temptable.AddNew
        temptable!str11 = 7
        temptable!str12 = "≈Ã„«·Ï ›Ì“«"
        temptable!str3 = " ÌÊ„Ì… " & xDate.text
        temptable!str1 = cManBox
'       temptable!str2 = TurnValue(!BOXDESCA)
        temptable!str5 = TurnValue(!DESCA3)
        temptable!val3 = !TVISA3
        temptable!Val10 = 2
        temptable.Update
        End If
        
        .MoveNext
    Loop
    End With
End If

    
    
cString = "SELECT   FILE0_50.DESCA AS BOXDESCA , BOXMOVE.BOX , SUM(BOXMOVE.PLUS - BOXMOVE.MINUS ) AS BALBOX FROM         BOXMOVE INNER JOIN FILE0_50 ON FILE0_50.CODE = BOXMOVE.BOX WHERE  DATE = " & DateSq(salesFrm.xDate.text)
If cManBox <> "" Then cString = cString & " AND BOX = " & MyParn(cManBox)
cString = cString & " GROUP BY FILE0_50.DESCA ,BOXMOVE.BOX ORDER BY BOXMOVE.BOX "
sourcetable.Close
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (sourcetable.EOF And sourcetable.BOF) Then
    With sourcetable
        Do Until sourcetable.EOF
            temptable.AddNew
            temptable!str11 = 6
            temptable!str12 = "≈Ã„«·Ï «·Õ“‰…"
            temptable!str3 = " ÌÊ„Ì… " & xDate.text
            temptable!str1 = !BOX
            temptable!str2 = !BOXDESCA
            temptable!str5 = "—’Ìœ «·Œ“‰…"
            
            temptable!val3 = !BalBox
            temptable!Val10 = 2
            temptable.Update
            sourcetable.MoveNext
        Loop
    End With
End If
    
contemp.BeginTrans
contemp.CommitTrans

REPORT1.Reset
FixPrinter REPORT1

REPORT1.ReportFileName = App.Path & "\Reports\TDAY.RPT"
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Destination = crptToWindow
REPORT1.WindowState = crptMaximized
REPORT1.Action = 1
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function IsPrinted() As Boolean
     IsPrinted = TurnValue(GetField("SELECT PRINTED FROM FILE6_20H WHERE DOC_NO = " & MyParn(xDoc_No.text), con), Null, False)
'     IsPrinted = ISPRINTED_sales(xDoc_No.text, con)
End Function
Sub FixAddress()
Dim loctable As New ADODB.Recordset
If lIsBranchStore Then
    loctable.Open "select * From Address where branch = " & MyParn(cBranch), con, adOpenStatic, adLockReadOnly
Else
    loctable.Open "select * From Address", con, adOpenStatic, adLockReadOnly
End If
If Not (loctable.EOF And loctable.BOF) Then
    cComp_Name = loctable!DESCA & ""
    cComp_address = loctable!Address & ""
    cComp_Phone = loctable!phone & ""
    cComp_Head1 = loctable!HEAD1 & ""
    cComp_Head2 = loctable!HEAD2 & ""
    nCountPrint = Val(loctable!COUNTPRINT & "")
    nLang_Boon = Val(loctable!lang & "")
End If
loctable.Close
Set loctable = Nothing
End Sub
Function TestRet() As Boolean
    TestRet = True
    With grid1
        For nRow = 1 To .Rows - 2
            If Val(.TextMatrix(nRow, 10)) < 0 Then
                nSal = Val(GetDesca("SELECT SUM(QUANT) FROM FILE6_20 WHERE DOC_NO = " & MyParn(XSALES_RET.Caption) & " AND ITEM = " & Val(.TextMatrix(nRow, 1)), con) & "")
                nRet = Val(GetDesca("SELECT SUM(QUANT) FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO WHERE SALES_RET = " & MyParn(XSALES_RET.Caption) & " AND ITEM = " & Val(.TextMatrix(nRow, 1)), con) & "")
                If nSal + nRet < 0 Then
                    MsgBox "„— Ã⁄ «ﬂ»— „‰ „»Ì⁄«  »Ê‰ «·»Ì⁄ - »Ê‰ «·»Ì⁄ €Ì— „”Ã· "
                    .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
                    TestRet = False
                End If
            End If
        Next nRow
    End With
End Function
Public Sub Print_TimeItemSales()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(1)
Dim pDate As Date, pstore As String
Dim nH1 As Double
Dim nH2 As Double

nH1 = Val(InputBox("", "„‰ ”«⁄… ") & "")
nH2 = Val(InputBox("", "„‰ ”«⁄… ", Hour(Time)) & "")

contemp.Execute "DELETE * FROM TEMP"


cF1 = " ( SELECT SUM([IN]-[OUT]) FROM  FILE1_11 WHERE SALES_MODEL.ITEM = FILE1_11.ITEM ) AS BALITEM  "
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
cString = " SELECT item ,FACTDESCA , MODELFACT , SCAL , COLOR ,  SUM(QUANT) AS t_q , " & cF1 & " FROM SALES_MODEL "
cString = cString & " where date = " & DateSq(xDate.text)
If nH1 <> 0 Then cString = cString & " and { fn HOUR(TIME) }  >=  " & nH1
If nH2 <> 0 Then cString = cString & " and { fn HOUR(TIME) }  <=  " & nH2
cString = cString & " GROUP BY item ,FACTDESCA , MODELFACT , SCAL , COLOR  "
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (sourcetable.EOF And sourcetable.BOF) Then
With sourcetable
    Do While Not sourcetable.EOF
        temptable.AddNew
        temptable!Date1 = dSalesDate
        temptable!date2 = Time
        temptable!str13 = TurnValue(cComp_Name)
        temptable!str1 = " „»Ì⁄«  „‰ : " & nH1 & " Õ Ï " & nH2
        temptable!str5 = TurnValue(!FACTDESCA)
        temptable!str6 = TurnValue(DelZero(!modelfact))
        temptable!str3 = !scal
        temptable!str4 = IIf(!color = "---" Or !color = "„ÊÕœ", Null, !color)
        temptable!val3 = !t_q
        temptable!VAL1 = !BALITEM
        temptable.Update
        .MoveNext
    Loop
End With
End If
    
contemp.BeginTrans
contemp.CommitTrans

REPORT1.Reset
FixPrinter REPORT1

REPORT1.ReportFileName = App.Path & "\Reports\item_sales.RPT"
REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
REPORT1.Destination = crptToWindow
REPORT1.WindowState = crptMaximized
REPORT1.Action = 1
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Function IsDayDisc(pDate) As Double
    Dim aRetDay(2)
    Dim DiscTable As New ADODB.Recordset
    DiscTable.Open "DISCOUNT", con, adOpenStatic, adLockReadOnly, adCmdTable
    With DiscTable
        Do While Not .EOF
            If DateValue(pDate) >= DateValue(!Date1) And DateValue(pDate) <= DateValue(!date2) Then
                IsDayDisc = !disc
                lIsPrice_2 = IIf(!ISPRICE_2, True, False)
                lIsdisc2 = IIf(!ISdisc2, True, False)
                cDescDiscout = TurnValue(!DESCA, Null, "")
                xNOTE_DISC.Caption = TurnValue(!DESCA, Null, "")
                xDOC_DISC.Caption = TurnValue(!doc_no, Null, "")
            End If
            .MoveNext
        Loop
    End With
End Function
Private Sub doprint_doc()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double

Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT FILE6_20.PRICE_C, FILE6_20.ITEM,FILE6_20.QUANT,FILE6_20.PRICE,FILE1_10.DESCA AS ITEM_DESCA,FILE6_20.TOTAL,FILE6_20.MAN,FILE6_20H.TIME,file6_20.man , file1_10.modelfact , FACT.DESCA AS FACTDESCA , FILE1_10.COLOR , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.MODEL FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO  INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM LEFT JOIN FACT ON FACT.CODE = FILE1_10.[FACT]  "
cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Do While Not loctable.EOF
    temptable.AddNew
    temptable!str6 = RetZero(DelZero(xDoc_No.text), 5)
    temptable!Date1 = xDate.text
    
    temptable!STR7 = xCodeDesca.Caption
    temptable!STR19 = TurnValue(cComp_Name)
    
    temptable!str2 = loctable!FACTDESCA
    temptable!str1 = loctable!Item
    temptable!str13 = TurnValue(DelZero(loctable!modelfact))
    temptable!str3 = loctable!ITEM_DESCA
    temptable!str4 = loctable!color
    temptable!str5 = loctable!scal
    temptable!str16 = loctable!MODEL
    
    temptable!val3 = loctable!Quant
    temptable!VAL1 = loctable!price
    temptable!VAL4 = loctable!TOTAL
    temptable!val5 = Val(xDiscount.text)
    temptable!VAL7 = Val(xTotal.text)
    temptable!val11 = Val(xcash.Caption)
    temptable!val12 = nBal
    temptable!VAL14 = i
    temptable.Update
    loctable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_SALES.rpt"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub cmdAddItems_Click()
    If grid1.Row = 0 Then grid1.Row = grid1.Rows - 1
    Set salesModel.myForm = Me
    salesModel.nColItem = 0
    salesModel.nColQuant = 10
    salesModel.strItem = grid1.TextMatrix(grid1.Row, 0)
    salesModel.strStore = xStore.BoundText
    salesModel.bshowBal = True
    salesModel.Show 1
End Sub

Sub myProcAddModel()
If Not MYVALID(False) Then Exit Sub
If myreplace(, True) Then
    Inform " „  «÷«›… «·„ÊœÌ·«  »‰Ã«Õ"
End If
myloadgrd

myUndo
If cBranch <> "00" Then
    grid1.ShowCell grid1.Rows - 1, 2 + 1
    grid1.Select grid1.Rows - 1, 2 + 1
End If
End Sub
Private Function myreplaceGrdModel() As Boolean
Dim cString As String, aitem As Variant
Dim cFPrice As String
cFPrice = "COST"
cString = " DELETE FROM FILE6_20 FROM FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM"
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.text)
cString = cString & turn(cString) & "FILE1_10.MODEL = " & MyParn(salesModel.XMODEL.text)
con.Execute cString, nDel
With salesModel.grid1
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                aitem = ItemFields(salesModel.GRID2.TextMatrix(nRow, nCol), con)
                If Not IsEmpty(aitem) Then
                    cString = "Insert into FILE6_20 (doc_no,item,price,price_c,cost,man,Quant)" & _
                               "Values(" & _
                                addstring(xDoc_No.text) & "," & _
                                addvalue(retFlag(aitem, "item")) & "," & _
                                Val(retFlag(aitem, cFPrice) & "") & "," & _
                                Val(retFlag(aitem, cFPrice) & "") & "," & _
                                Val(retFlag(aitem, "COST") & "") & "," & _
                                addstring(salesFrm.xMan.BoundText) & "," & _
                                Val(.TextMatrix(nRow, nCol)) & _
                               ")"
                    con.Execute cString
                End If
            End If
        Next
    Next
End With
End Function
Private Function sendItems() As Long
Dim cFile As String
cFile = "FILE1_10"

cString = "SELECT * FROM FILE1_10 WHERE  ITEM  = " & Val(XBARCODEADD.text) & " OR BARCODE =  " & MyParn(XBARCODEADD.text) & " OR BARCODE2 =  " & MyParn(XBARCODEADD.text) & " OR BARCODE13 =  " & MyParn(XBARCODEADD.text) & " OR BARCODE_GS1 =  " & MyParn(XBARCODEADD.text)
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    con.Execute " DELETE FROM FILE1_10 where item = " & loctable!Item
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!MODEL))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable!MODELNO))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "FACT", addstring(loctable!Fact))
    aInsert = AddFlag(aInsert, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "MODELFACT", addstring(loctable!modelfact))
    aInsert = AddFlag(aInsert, "MODELFACT0", addstring(loctable!modelfact0))
    aInsert = AddFlag(aInsert, "SUPP", addstring(loctable!SUPP))
    aInsert = AddFlag(aInsert, "OKAZ", Val(loctable!okaz & ""))
    aInsert = AddFlag(aInsert, "OKAZ_2", Val(loctable!okaz_2 & ""))
    aInsert = AddFlag(aInsert, "RATE", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "[SECTION]", addvalue(loctable!Section))
    aInsert = AddFlag(aInsert, "[ITEM]", addvalue(loctable!Item))
    aInsert = AddFlag(aInsert, "[COST]", Val(loctable!cost & ""))
     
    aInsert = AddFlag(aInsert, "[COSTITEM]", Val(loctable!costITEM & ""))
    aInsert = AddFlag(aInsert, "[COST2]", Val(loctable!cost2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE]", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "[PRICE2]", Val(loctable!PRICE2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE_2]", Val(loctable!PRICE_2 & ""))
    aInsert = AddFlag(aInsert, "[SCAL]", addstring(loctable!scal))
    aInsert = AddFlag(aInsert, "[C_SCAL]", addvalue(loctable!C_SCAL))
    aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable!color))
    aInsert = AddFlag(aInsert, "[C_COLOR]", addvalue(loctable!c_Color))
    aInsert = AddFlag(aInsert, "[BARCODE]", addstring(loctable!BARCODE))
    aInsert = AddFlag(aInsert, "[BARCODE2]", addstring(loctable!BARCODE2))
    aInsert = AddFlag(aInsert, "[BARCODE13]", addstring(loctable!BARCODE13))
    aInsert = AddFlag(aInsert, "[BARCODE_GS1]", addstring(loctable!BARCODE_GS1))
    aInsert = AddFlag(aInsert, "[ISOKAZITEM]", IIf(loctable!ISOKAZITEM, 1, 0))
    aInsert = AddFlag(aInsert, "[MOSM2]", addstring(loctable!MOSM2))
    aInsert = AddFlag(aInsert, "[ISNOITEM]", IIf(loctable!ISNOITEM, 1, 0))
    aInsert = AddFlag(aInsert, "[ISNODEM]", IIf(loctable!ISNODEM, 1, 0))
    
    
    
    
    
    con.Execute addInsert(aInsert, "FILE1_10"), nAffect
    Inform "  „ ≈÷«›… »«—ﬂÊœ " & loctable!Item & "  " & loctable!DESCA
    loctable.MoveNext
Loop
End Function
Private Function SendSales() As Long
On Error GoTo myerror
cFile = "FILE6_20"
Set loctable = New ADODB.Recordset
Dim cStr1 As String, cStr2 As String, cDelStr1 As String, cDelStr2 As String, cStrNew As String, cStrPhone As String
cString = "Select * from " & cFile & "H WHERE (PRINTED = 1 and isnew = 1) "
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long, lSkip As Boolean
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
'    cDelStr1 = cDelStr1 & " delete from file6_20 where doc_no = " & MyParn(loctable!doc_no) & " ; "
'    cDelStr2 = cDelStr2 & " delete from file6_20H where doc_no = " & MyParn(loctable!doc_no) & " ; "
    
    conShop.Execute "delete from file6_20 where doc_no = " & MyParn(loctable!doc_no)
    conShop.Execute "delete from file6_20h where doc_no = " & MyParn(loctable!doc_no)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[TIME]", addTime(Format(loctable!Time, "SHORT TIME")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "[NOTES]", addstring(loctable!NOTES))
    aInsert = AddFlag(aInsert, "[RATE]", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[CASH]", Val(loctable!CASH & ""))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
'    aInsert = AddFlag(aInsert, "[BOXVISA]", addstring(loctable!BOXVISA))
    aInsert = AddFlag(aInsert, "[PAY]", Val(loctable!PAY & ""))
    aInsert = AddFlag(aInsert, "[VISA]", Val(loctable!Visa & ""))
    aInsert = AddFlag(aInsert, "[LATE]", Val(loctable!late & ""))
    aInsert = AddFlag(aInsert, "[REST]", Val(loctable!rest & ""))
    aInsert = AddFlag(aInsert, "[MAN]", addstring(loctable!MAN))
    aInsert = AddFlag(aInsert, "[PRINTED]", IIf(loctable!PRINTED, "1", "0"))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "[phone]", addstring(loctable!phone))
                                                                                
    aInsert = AddFlag(aInsert, "[username_disc]", addstring(loctable!username_disc))
    aInsert = AddFlag(aInsert, "[username_ret]", addstring(loctable!username_ret))
    aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(loctable!SALES_RET))
    aInsert = AddFlag(aInsert, "CODEVISA", addvalue(loctable!CODEVISA))
    aInsert = AddFlag(aInsert, "[NOTE_DISC]", addstring(loctable!NOTE_DISC))
    aInsert = AddFlag(aInsert, "[DOC_DISC]", addstring(loctable!DOC_DISC))
    aInsert = AddFlag(aInsert, "[isbankahly]", IIf(loctable!isbankahly, 1, 0))
    aInsert = AddFlag(aInsert, "[VISADOC]", addstring(loctable!VISADOC))
    
    aInsert = AddFlag(aInsert, "VISA2", addvalue(loctable!VISA2))
    aInsert = AddFlag(aInsert, "TYPEVISA", addvalue(loctable!TYPEVISA))
    aInsert = AddFlag(aInsert, "point", Val(loctable!Point & ""))
    
    aInsert = AddFlag(aInsert, "ONLINE", addvalue(loctable!online))
    aInsert = AddFlag(aInsert, "CHARGE1", Val(loctable!CHARGE1 & ""))
    aInsert = AddFlag(aInsert, "CHARGE2", Val(loctable!charge2 & ""))
    aInsert = AddFlag(aInsert, "[NAME]", addstring(loctable!Name))
    aInsert = AddFlag(aInsert, "[SHIP]", addstring(loctable!SHIP))
    aInsert = AddFlag(aInsert, "[ADDRESS]", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "[SHIP_NO]", addstring(loctable!SHIP_NO))
    aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
    aInsert = AddFlag(aInsert, "[DATE_PAY]", addDate(loctable!DATE_PAY))
    aInsert = AddFlag(aInsert, "[BANK]", addstring(loctable!BANK))
    aInsert = AddFlag(aInsert, "[note_ret]", addstring(loctable!note_ret))
    aInsert = AddFlag(aInsert, "[ONLINE_DOC]", addstring(loctable!ONLINE_DOC))
    aInsert = AddFlag(aInsert, "[NORET_CHARGE1]", IIf(loctable!NORET_CHARGE1, 1, 0))
    
    
    aInsert = AddFlag(aInsert, "[card_disc]", addstring(loctable!card_disc))
    aInsert = AddFlag(aInsert, "[card_desca]", addstring(loctable!card_desca))
    aInsert = AddFlag(aInsert, "[card_VALUE]", addvalue(loctable!card_VALUE))
    aInsert = AddFlag(aInsert, "[card_DOC]", addstring(loctable!card_DOC))
    
    aInsert = AddFlag(aInsert, "[VISA3]", Val(loctable!VISA3 & ""))
    aInsert = AddFlag(aInsert, "[CODEVISA3]", addstring(loctable!CODEVISA3))
    aInsert = AddFlag(aInsert, "[VISADOC3]", addstring(loctable!VISADOC3))
    aInsert = AddFlag(aInsert, "[VISA3_2]", Val(loctable!VISA3_2))
    aInsert = AddFlag(aInsert, "[TYPEVISA3 ]", addstring(loctable!TYPEVISA3))
    
    If Not IsNull(loctable!card_DOC) Then
        conShop.Execute " UPDATE CARD_DISC SET SALES_DOC_NO = " & addstring(loctable!doc_no) & " WHERE BARCODE_DISC = " & MyParn(loctable!card_DOC)
    End If

'    cStr1 = cStr1 & addInsert(aInsert, cFile & "H") & " ; "
    conShop.Execute addInsert(aInsert, cFile & "H")
    lSkip = False
    
    Set loctable2 = New ADODB.Recordset
    cString = "Select FILE6_20.* from FILE6_20 WHERE FILE6_20.doc_no = " & MyParn(loctable!doc_no)
    loctable2.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    Dim sitem As String
        
    cStr2 = ""
    Dim aitem As Variant
    Do Until loctable2.EOF
        sitem = loctable2!Item
        If GetDesca("select item from file1_10 where  item = " & sitem, conShop) <> "" Then
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "MAN", addstring(loctable2!MAN))
            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
            aInsert2 = AddFlag(aInsert2, "DISCOUNT", Val(loctable2!discount & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE_C", Val(loctable2!PRICE_C & ""))
            aInsert2 = AddFlag(aInsert2, "ROW", Val(loctable2!Row & ""))
            aInsert2 = AddFlag(aInsert2, "COST", Val(loctable2!cost & ""))
            aInsert2 = AddFlag(aInsert2, "S_OKAZ", Val(loctable2!S_OKAZ & ""))
            aInsert2 = AddFlag(aInsert2, "ISDISC2_1", IIf(loctable2!ISDISC2_1, 1, 0))
            cStr2 = cStr2 & addInsert(aInsert2, "file6_20") & " ; "
        Else
'            MsgBox " „—«Ã⁄… „»Ì⁄«  »«—ﬂÊœ " & sitem
            lSkip = True
        End If
        loctable2.MoveNext
    Loop
    If cStr2 <> "" Then conShop.Execute cStr2
    
    If Not lSkip Then
'        cStrNew = cStrNew & " update file6_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no) & " ; "
'        cStrNew = cStrNew & " update file6_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no) & " ; "
        con.Execute " update file6_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no)
    End If
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
'If cDelStr1 <> "" Then conShop.Execute cDelStr1
'If cDelStr2 <> "" Then conShop.Execute cDelStr2

'If cStr1 <> "" Then conShop.Execute cStr1
'If cStr2 <> "" Then conShop.Execute cStr2
  

'If cStrNew <> "" Then
'    con.Execute cStrNew
'End If

Inform " „»Ì⁄«  " & nRecord
SendSales = nAffectTotal

Inform "”Õ» «· Ì·›Ê‰« "
Set loctable = New ADODB.Recordset
cString = "Select * from SUBCUST WHERE isnew = 1 "
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "PHONE", addstring(loctable!phone))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "E_MAIL", addstring(loctable!E_MAIL))
    aInsert = AddFlag(aInsert, "branch", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "F_DATE", addDate(loctable!F_DATE))
    If cStrPhone = "" Then
        cStrPhone = addInsert(aInsert, "SUBCUST")
    Else
        cStrPhone = cStrPhone & " ; " & addInsert(aInsert, "SUBCUST")
    End If
    loctable.MoveNext
Loop
If cStrPhone <> "" Then
    conShop.Execute cStrPhone
    con.Execute " UPDATE SUBCUST SET isnew = 0 "
End If
Inform " phone : " & nRecord

'ClearTemp6_20

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendSales = -1
GoTo lastsub
End Function


Private Function sendCharge() As Long
cFile = "FILE8_60"
Dim loctable   As New ADODB.Recordset
Dim loctable2  As New ADODB.Recordset
cString = "Select * from FILE8_60H WHERE ISNEW = 1 "

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    
    conShop.Execute " delete from file8_60 where DOC_NO = " & MyParn(loctable!doc_no) & " AND DOC_NO IN ( SELECT DOC_NO FROM FILE8_60H WHERE BRANCH = " & MyParn(sBranch) & " ) "
    conShop.Execute " delete from file8_60h where DOC_NO = " & MyParn(loctable!doc_no) & " AND BRANCH = " & MyParn(sBranch)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranch))
    aInsert = AddFlag(aInsert, "ISCLOSED", 1)
    conShop.Execute addInsert(aInsert, cFile & "H"), nAffect

'''''''''''''
    Dim aInsert2 As Variant
    cString = "Select * FROM FILE8_60 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    Do Until loctable2.EOF
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
        aInsert2 = AddFlag(aInsert2, "[BOX]", addstring(loctable2!BOX))
        aInsert2 = AddFlag(aInsert2, "[CHARGE]", addstring(loctable2!CHARGE))
        aInsert2 = AddFlag(aInsert2, "[DESCA]", addstring(loctable2!DESCA))
        aInsert2 = AddFlag(aInsert2, "[BRANCH]", addstring(loctable2!branch))
        aInsert2 = AddFlag(aInsert2, "[VALUE]", Val(loctable2!Value & ""))
        aInsert2 = AddFlag(aInsert2, "[ROW]", Val(loctable2!Row & ""))
        conShop.Execute addInsert(aInsert2, "FILE8_60"), nAffect
        loctable2.MoveNext
    Loop
'''''''''''''
    con.Execute " UPDATE FILE8_60H SET ISNEW = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendCharge = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendCharge = -1
GoTo lastsub
End Function

Private Function SendTransBox() As Long
Set loctable = New ADODB.Recordset
Dim cDoc As String

cString = "Select * from FILE0_51 WHERE isnew = 1"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Do Until loctable.EOF
    cDoc = cBranch & loctable!CODE
    conShop.Execute " delete from file0_52  where CODE = " & MyParn(cDoc)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(cDoc))
    aInsert = AddFlag(aInsert, "NO1", addstring(loctable!no1))
    aInsert = AddFlag(aInsert, "NO2", addstring(loctable!no2))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "BRANCH", addstring(cBranch))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(loctable!Value))
                                                                                
    conShop.Execute addInsert(aInsert, "FILE0_52"), nAffect
    con.Execute " update FILE0_51 set isnew = 0 where CODE = " & MyParn(loctable!CODE)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
SendTransBox = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendTransBox = -1
GoTo lastsub
End Function
Private Function SendSales_fr() As Long
On Error GoTo myerror
cFile = "FILE6_20"
Set loctable = New ADODB.Recordset
Dim cStr1 As String, cStr2 As String, cDelStr1 As String, cDelStr2 As String, cStrNew As String, cStrPhone As String
cString = "Select * from " & cFile & "H WHERE ( PRINTED = 1 and isnew = 1 ) or ONLINE > 0 "
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long, lSkip As Boolean
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
    cDelStr1 = cDelStr1 & " delete from FR6_20      where doc_no = " & MyParn(loctable!doc_no) & " ; "
    cDelStr2 = cDelStr2 & " delete from FR6_20H     where doc_no = " & MyParn(loctable!doc_no) & " ; "
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[TIME]", addTime(Format(loctable!Time, "SHORT TIME")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "[NOTES]", addstring(loctable!NOTES))
    aInsert = AddFlag(aInsert, "[RATE]", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[CASH]", Val(loctable!CASH & ""))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
'    aInsert = AddFlag(aInsert, "[BOXVISA]", addstring(loctable!BOXVISA))
    aInsert = AddFlag(aInsert, "[PAY]", Val(loctable!PAY & ""))
    aInsert = AddFlag(aInsert, "[VISA]", Val(loctable!Visa & ""))
    aInsert = AddFlag(aInsert, "[LATE]", Val(loctable!late & ""))
    aInsert = AddFlag(aInsert, "[REST]", Val(loctable!rest & ""))
    aInsert = AddFlag(aInsert, "[MAN]", addstring(loctable!MAN))
    aInsert = AddFlag(aInsert, "[PRINTED]", IIf(loctable!PRINTED, "1", "0"))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "[phone]", addstring(loctable!phone))
                                                                                
    aInsert = AddFlag(aInsert, "[username_disc]", addstring(loctable!username_disc))
    aInsert = AddFlag(aInsert, "[username_ret]", addstring(loctable!username_ret))
    aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(loctable!SALES_RET))
'    aInsert = AddFlag(aInsert, "CODEVISA", addvalue(loctable!CODEVISA))
                                                                                
    aInsert = AddFlag(aInsert, "CHARGE1", Val(loctable!CHARGE1 & ""))
    aInsert = AddFlag(aInsert, "CHARGE2", Val(loctable!charge2 & ""))
    aInsert = AddFlag(aInsert, "[NAME]", addstring(loctable!Name))
    aInsert = AddFlag(aInsert, "[SHIP]", addstring(loctable!SHIP))
    aInsert = AddFlag(aInsert, "[ADDRESS]", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "[SHIP_NO]", addstring(loctable!SHIP_NO))
    aInsert = AddFlag(aInsert, "[DATE_PAY]", addDate(loctable!DATE_PAY))
    aInsert = AddFlag(aInsert, "[BANK]", addstring(loctable!BANK))
    aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
    aInsert = AddFlag(aInsert, "[NOTE_DISC]", addstring(loctable!NOTE_DISC))
    aInsert = AddFlag(aInsert, "[DOC_DISC]", addstring(loctable!DOC_DISC))
    aInsert = AddFlag(aInsert, "ONLINE", addvalue(loctable!online))
    aInsert = AddFlag(aInsert, "[note_ret]", addstring(loctable!note_ret))
    aInsert = AddFlag(aInsert, "[card_disc]", addstring(loctable!card_disc))
    aInsert = AddFlag(aInsert, "[card_desca]", addstring(loctable!card_desca))
    aInsert = AddFlag(aInsert, "[card_VALUE]", addvalue(loctable!card_VALUE))
    aInsert = AddFlag(aInsert, "[card_DOC]", addstring(loctable!card_DOC))
    
    aInsert = AddFlag(aInsert, "[VISA3]", Val(loctable!VISA3 & ""))
    aInsert = AddFlag(aInsert, "[CODEVISA3]", addstring(loctable!CODEVISA3))
    aInsert = AddFlag(aInsert, "[VISADOC3]", addstring(loctable!VISADOC3))
    aInsert = AddFlag(aInsert, "[VISA3_2]", Val(loctable!VISA3_2))
    aInsert = AddFlag(aInsert, "[TYPEVISA3 ]", addstring(loctable!TYPEVISA3))
                                                                                
    If Not IsNull(loctable!card_DOC) Then
        conShop.Execute " UPDATE CARD_DISC SET SALES_DOC_NO = " & addstring(loctable!doc_no) & " WHERE BARCODE_DISC = " & MyParn(loctable!card_DOC)
    End If
                                                                                
    cStr1 = cStr1 & addInsert(aInsert, "FR6_20H") & " ; "
    
    lSkip = False
    
    Set loctable2 = New ADODB.Recordset
    cString = "Select FILE6_20.* from FILE6_20 WHERE FILE6_20.doc_no = " & MyParn(loctable!doc_no)
    loctable2.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    Dim sitem As String
    Dim aitem As Variant
    Do Until loctable2.EOF
        sitem = loctable2!Item
        If GetDesca("select item from file1_10 where  item = " & sitem, conShop) <> "" Then
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "MAN", addstring(loctable2!MAN))
            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
            aInsert2 = AddFlag(aInsert2, "DISCOUNT", Val(loctable2!discount & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE_C", Val(loctable2!PRICE_C & ""))
            aInsert2 = AddFlag(aInsert2, "ROW", Val(loctable2!Row & ""))
            aInsert2 = AddFlag(aInsert2, "COST", Val(loctable2!cost & ""))
            aInsert2 = AddFlag(aInsert2, "S_OKAZ", Val(loctable2!S_OKAZ & ""))
            cStr2 = cStr2 & addInsert(aInsert2, "FR6_20") & " ; "
        Else
'            MsgBox " „—«Ã⁄… „»Ì⁄«  »«—ﬂÊœ " & sitem
            lSkip = True
        End If
        loctable2.MoveNext
    Loop
    If Not lSkip Then
        cStrNew = cStrNew & " update file6_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no) & " ; "
    End If
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
If cDelStr1 <> "" Then
    conShop.Execute cDelStr1
    conShop.Execute cDelStr2
End If
If cStr1 <> "" Then
    conShop.Execute cStr1
    conShop.Execute cStr2
End If
If cStrNew <> "" Then
    con.Execute cStrNew
End If
Inform " „»Ì⁄«  " & nRecord
SendSales_fr = nAffectTotal


Inform "”Õ» «· Ì·›Ê‰« "
Set loctable = New ADODB.Recordset
cString = "Select * from SUBCUST WHERE isnew = 1 "
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "PHONE", addstring(loctable!phone))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "E_MAIL", addstring(loctable!E_MAIL))
    aInsert = AddFlag(aInsert, "branch", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "F_DATE", addDate(loctable!F_DATE))
    cStrPhone = cStrPhone & addInsert(aInsert, "SUBCUST") & " ; "
    loctable.MoveNext
Loop
If cStrPhone <> "" Then
    conShop.Execute cStrPhone
    con.Execute " UPDATE SUBCUST SET isnew = 0 "
End If
Inform " phone : " & nRecord
'ClearTemp6_20
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendSales_fr = -1
GoTo lastsub
End Function


Private Sub sned_purch_Click()
    Dim cDoc As String, cStore As String, cCode As String
    cCode = "000"
    cStore = GetDesca("SELECT BRANCH FROM FILE4_10 WHERE CUST = " & MyParn(xCode.text), con)
    If cStore = "" Then
        MsgBox " «·⁄„Ì· ·Ì”  ÊﬂÌ· "
        Exit Sub
    End If
    Dim aInsert As Variant
    If MsgBox("≈·Ï „” ‰œ „‘ —Ì«   ÊﬂÌ·«  ", vbYesNo + vbDefaultButton2) = vbYes Then
        aInsert = AddFlag(Empty, "[DATE]", DateSq(Date))
        aInsert = AddFlag(aInsert, "[mosm]", addstring(cPMosm))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(cStore))
        aInsert = AddFlag(aInsert, "[code]", addstring(cCode))
        aInsert = AddFlag(aInsert, "[isdate]", "getdate()")
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(cusername))
        aInsert = AddFlag(aInsert, "[Notes_Doc]", addstring(" „»Ì⁄«  Junior "))
        aInsert = AddFlag(aInsert, "[TRANS_DOC]", addstring(xDoc_No.text))
        aInsert = AddFlag(aInsert, "[ISNEW1]", 1)
        cDoc = Newflag_PurchBr("FR7_20H", "DOC_NO", cStore, con)
        aInsert = AddFlag(aInsert, "DOC_NO", addstring(cDoc))
        con.Execute addInsert(aInsert, "FR7_20h")
        
        con.Execute "  INSERT INTO FR7_20 (DOC_NO , quant, ITEM, price )  SELECT  " & addstring(cDoc) & " , quant, FILE6_20.ITEM, file6_20.price From FILE6_20 inner join file1_10 on FILE6_20.item = file1_10.item WHERE DOC_NO = " & MyParn(xDoc_No.text)
        Inform " „  ÕÊÌ· ≈·Ï „” ‰œ „‘ —Ì«   ÊﬂÌ·«  —ﬁ„ " & cDoc
        myUndo
    End If
End Sub
Function CalcDisc2(pQ) As Double
    Dim DiscTable As New ADODB.Recordset, nCountQ As Double
    Dim SubItemTable As New ADODB.Recordset
    Dim lIsTestQty As Boolean
    cStr1 = "SELECT FILE6_20.ITEM , FILE6_20.QUANT , FILE1_10.PRICE  , DOC_NO FROM FILE6_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_20.ITEM INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE WHERE FILE1_10SC.ISDISC = 1 AND  ( FILE1_10.OKAZ_2 < " & nMaxDisc2 & " and FILE1_10.okaz < " & nMaxDisc2 & " )  And doc_no = " & MyParn(xDoc_No.text)
    SubItemTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    lIsTestQty = True
    nQDisc = Int(pQ / 3)
    With SubItemTable
    Do While Not SubItemTable.EOF
        For nQ = 1 To !Quant
            nCountQ = nCountQ + 1
            cString = "Insert into [PRICE]  (doc_no,PRICE) " & _
                               "Values(" & _
                               addstring(!doc_no) & ", " & _
                               Val(!price) & _
                               ")"
             
            con.Execute cString
        Next nQ
        SubItemTable.MoveNext
    Loop
    End With
    If nCountQ >= 3 And Val(xTotalQuant2.Caption) > 0 Then
        MsgBox "  ·‰ Ì „ Õ”«» Œ’„ «·⁄—÷ - ÌÊÃœ „— Ã⁄ ›Ï «·›« Ê—… "
       CalcDisc2 = -9999
        Exit Function
    End If
    If nCountQ >= 3 And pQ <> nCountQ Then
        MsgBox "  ·‰ Ì „ Õ”«» Œ’„ «·⁄—÷ - ÌÊÃœ «’‰«› Œ«—Ã «·⁄—÷ "
       CalcDisc2 = -9999
        Exit Function
    End If
    
    If Int(nCountQ / 3) <> (nCountQ / 3) Then
        MsgBox "·‰ Ì „ Õ”«» Œ’„ «·⁄—÷ - ⁄œœ «·ﬁÿ⁄ ›Ï «·⁄—÷ " & nCountQ & Chr(13) & " »—Ã«¡ “Ì«œ… ⁄œœ ﬁÿ⁄ ··≈” ›«œ… „‰ «·⁄—÷ 2+1"
'       CalcDisc2 = -9999
        Exit Function
    End If
    
    If nCountQ >= 3 Then
        If Int(nCountQ / 3) = (nCountQ / 3) Then
            Inform "⁄œœ «·ﬁÿ⁄ ›Ï «·⁄—÷ " & nCountQ
        Else
            Inform_OK "⁄œœ «·ﬁÿ⁄ ›Ï «·⁄—÷ " & nCountQ & Chr(13) & " »—Ã«¡ “Ì«œ… ⁄œœ ﬁÿ⁄ ··≈” ›«œ… „‰ «·⁄—÷ 2+1"
        End If
        con.Execute " UPDATE  FILE6_20 SET  file6_20.price = file6_20.PRICE_C , isdisc2_1 = 1  , S_OKAZ = 0 FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM INNER JOIN FILE1_10SC ON FILE1_10.SECTION = FILE1_10SC.CODE where FILE1_10SC.ISDISC = 1 AND DOC_NO = " & MyParn(xDoc_No.text)
    Else
        con.Execute " UPDATE  FILE6_20 SET  isdisc2_1 = 0  WHERE DOC_NO = " & MyParn(xDoc_No.text)
    End If
    myloadgrd
    CalcTotals
    DiscTable.Open "SELECT * FROM PRICE WHERE DOC_NO = " & MyParn(xDoc_No.text) & " ORDER BY PRICE  DESC ", con, adOpenStatic, adLockReadOnly, adCmdText
    If Not (DiscTable.EOF Or DiscTable.BOF) Then
        DiscTable.MoveFirst
        For i = 1 To 1000
            If Int(i / 3) = (i / 3) Then
                nTDiscInv = nTDiscInv + Val(DiscTable!price)
            End If
            DiscTable.MoveNext
            If DiscTable.EOF Then Exit For
        Next i
    End If
    CalcDisc2 = nTDiscInv
'   xNOTE_DISC.Caption = "Œ’„ ⁄—÷ 2+1"
    con.Execute " DELETE FROM PRICE WHERE DOC_NO = " & MyParn(xDoc_No.text)
End Function
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

Private Sub XVISADOC_Click()
'    Dim cVisaDoc As String
'    cVisaDoc = InputBox("", "—ﬁ„ «·⁄„Ì·… ", XVISADOC.Caption)
'    XVISADOC.Caption = cVisaDoc
'    con.Execute " UPDATE FILE6_20H SET ISNEW = 1 , VISADOC = " & addvalue(cVisaDoc) & " WHERE DOC_NO = " & MyParn(xDoc_No.Text), nRec
'    Inform "  „  ⁄œÌ· ⁄œœ  " & nRec
End Sub
Private Function myreplace_sub()
Dim aInsert(13, 1)

aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.text)

aInsert(1, 0) = "Cash"
aInsert(1, 1) = Val(xcash.Caption)

aInsert(2, 0) = "Pay"
aInsert(2, 1) = Val(xPay.Caption)

aInsert(3, 0) = "visa"
aInsert(3, 1) = Val(xvisa.Caption)

aInsert(4, 0) = "late"
aInsert(4, 1) = Val(xlate.Caption)

aInsert(5, 0) = "rest"
aInsert(5, 1) = Val(xRest.Caption)

aInsert(6, 0) = "CODEVISA"
aInsert(6, 1) = addvalue(XCODEVISA.BoundText)

aInsert(7, 0) = "isbankahly"
aInsert(7, 1) = IIf(Val(xisbankahly.Value) = "0", 0, 1)

aInsert(8, 0) = "VISADOC"
aInsert(8, 1) = addstring(XVISADOC.Caption)

aInsert(9, 0) = "VISA2"
aInsert(9, 1) = Val(XVISA2.Caption)

aInsert(10, 0) = "TYPEVISA"
aInsert(10, 1) = addvalue(XTYPEVISA.Caption)

aInsert(11, 0) = "isnew"
aInsert(11, 1) = 1

aInsert(12, 0) = "phone"
aInsert(12, 1) = addstring(xphone.Caption)

aInsert(13, 0) = "POINT"
aInsert(13, 1) = Val(XPOINT.Caption)

con.Execute CreateUpdate(aInsert, "FILE6_20H", " where doc_no = " & addstring(xDoc_No.text))
End Function
Private Sub cmd_showdoc_Click_AMR()
    Dim pDocRet As String
    Dim dDateRet As Date
    Dim cPassword As String, aUser As Variant
    
    pDocRet = xdoc_ret.text
    xdoc_ret.text = pDocRet
    If Len(xdoc_ret.text) >= 11 Then
        XISRET.Value = IIf(GetBoolean("SELECT ISRET FROM FILE6_20H WHERE DOC_NO = " & MyParn(xdoc_ret.text), con), 1, 0)
        dDateRet = DateValue(Format(Mid(xdoc_ret.text, 1, 2) & "-" & Mid(xdoc_ret.text, 3, 2) & "-" & Mid(xdoc_ret.text, 5, 2), "DD-MM-YYYY"))
        If DateDiff("D", dDateRet, dSalesDate) > 30 And XISRET.Value = 0 Then
            MsgBox " »Ê‰ „»Ì⁄«   «—ÌŒ… «ﬂ»— „‰ 30 ÌÊ„"
            Exit Sub
        End If

         
        ManagPass.pPassWord = ""
        ManagPass.Show 1
        cPassword = ManagPass.pPassWord
        
        If Trim(cPassword) = "" Then Exit Sub
        aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
        If UBound(aUser) > 0 Then
            lRetDoc = True
            XISRET.Value = 1
            xusername_RET.Caption = aUser(2)
        Else
            Exit Sub
        End If
        SalesDoc.sDoc_no = xdoc_ret.text
        SalesDoc.Show 1
        If xinv_no.text <> "" Then
            cmdSave_Click
        End If
    End If
End Sub
Private Sub cmd_disc_Click()
Dim cPassword As String, aUser As Variant
If xPrinted.Value <> 0 Then Exit Sub

    ManagPass.pPassWord = ""
    ManagPass.Show 1
    cPassword = Trim(ManagPass.pPassWord)
    If cPassword = "" Then Exit Sub
    aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
    If UBound(aUser) > 0 Then
        MsgBox "Ì„ﬂ‰  ”ÃÌ· «·Œ’„ "
        xDiscount.Locked = False
        xRate.Locked = False
        XUSERNAME_DISC.Caption = aUser(2)
        xNOTE_DISC.Caption = ""
        xDOC_DISC.Caption = ""
    Else
        Exit Sub
    End If
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
Private Function openCardTable(Optional pDoc_No As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset
cFilter = ""
cFilterLook = ""
cString = "SELECT TOP 1 FILE6_20H.* , FILE3_10.DESCA AS CLIENTDESCA , FILE3_10.ISONEST from FILE6_20H inner join file3_10 on file3_10.code = FILE6_20H.code INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE WHERE FILE6_20H.ISINVOICE = 0"
If cBranch = "00" And sDoc_no = "" Then
    If Not lNotBranch Then
        cFilter = cFilter & " AND FILE6_20H.BRANCH = '00' "
        cFilterLook = cFilterLook & " AND FILE6_20H.BRANCH = '00' "
    Else
'        cFilter = cFilter & " AND FILE0_40.ISSTOREBRANCH = 1"
'        cFilterLook = cFilterLook & " AND FILE0_40.ISSTOREBRANCH = 1 "
    
        cFilter = cFilter & " AND FILE6_20H.BRANCH <> '00' "
        cFilterLook = cFilterLook & " AND FILE6_20H.BRANCH <> '00' "
    
    End If
    If lNotBranch Then
        If IsDate(xdate_open.text) Then
            cFilter = cFilter & " AND  FILE6_20H.date = " & DateSq(Format(xdate_open.text, "DD-MM-YYYY"))
            cFilterLook = cFilterLook & " AND FILE6_20H.date = " & DateSq(Format(xdate_open.text, "DD-MM-YYYY"))
        ElseIf IsDate(dSalesDate) Then
            cFilter = cFilter & " AND FILE6_20H.date = " & DateSq(dSalesDate)
            cFilterLook = cFilterLook & " AND FILE6_20H.date = " & DateSq(dSalesDate)
        End If
    End If
Else
    If sDoc_no <> "" Then
        cString = cString & " AND DOC_NO = " & MyParn(sDoc_no)
    Else
        If IsDate(xdate_open.text) Then
            cFilter = cFilter & " AND  FILE6_20H.BRANCH = " & MyParn(cBranch) & " AND FILE6_20H.date = " & DateSq(Format(xdate_open.text, "DD-MM-YYYY"))
            cFilterLook = cFilterLook & " AND FILE6_20H.date = " & DateSq(Format(xdate_open.text, "DD-MM-YYYY"))
        ElseIf IsDate(dSalesDate) Then
            cFilter = cFilter & " AND  FILE6_20H.BRANCH = " & MyParn(cBranch) & " AND FILE6_20H.date = " & DateSq(dSalesDate)
            cFilterLook = cFilterLook & " AND FILE6_20H.date = " & DateSq(dSalesDate)
        End If
        If cBranchBox <> "" Then
            cFilter = cFilter & " AND  FILE6_20H.BRANCH = " & MyParn(cBranch) & " AND  FILE6_20H.BOX = " & MyParn(cBranchBox)
            cFilterLook = cFilterLook & " AND FILE6_20H.BOX = " & MyParn(cBranchBox)
        End If
    End If
End If
If pSign = "=" Then
    If pDoc_No <> "" Then cWhere = " AND DOC_NO  " & pSign & addstring(pDoc_No)
Else
    If pDoc_No <> "" Then cWhere = " AND DOC_NO  " & pSign & addstring(pDoc_No)
End If
' «·«Œ Ì«—« 
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
Private Sub OLDDDopenCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cFilter = ""
cString = "SELECT FILE6_20H.*,FILE3_10.DESCA AS CLIENTDESCA FROM FILE6_20H INNER JOIN FILE3_10 ON FILE6_20H.Code = FILE3_10.CODE"
If cBranch = "00" And sDoc_no = "" Then
    cFilter = cFilter & turn(cFilter, " AND ") & " FILE6_20H.BRANCH = '00' "
    If sDoc_no <> "" Then cFilter = cFilter & turn(cFilter, " AND ") & " DOC_NO = " & MyParn(sDoc_no)
Else
    If sDoc_no = "" Then
        If IsDate(xdate_open.text) Then
            cFilter = cFilter & turn(cFilter, " AND ") & " FILE6_20H.date = " & DateSq(Format(xdate_open.text, "DD-MM-YYYY"))
        ElseIf IsDate(dSalesDate) Then
            cFilter = cFilter & turn(cFilter, " AND ") & " FILE6_20H.date = " & DateSq(dSalesDate)
        End If
    Else
        If sDoc_no <> "" Then cFilter = cFilter & turn(cFilter, " AND ") & " DOC_NO = " & MyParn(sDoc_no)
    End If
        
    If cBranchBox <> "" Then
        cFilter = cFilter & turn(cFilter, " AND ") & " FILE6_20H.BOX = " & MyParn(cBranchBox)
    End If
End If
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY DATE , DOC_NO2"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Function IsOnLine(pDoc_No) As Boolean
    aRet = aGetDesca("SELECT NAME , PHONE , ADDRESS , CHARGE1 , CHARGE2 , SHIP_NO , SHIP FROM FILE6_20H WHERE DOC_NO = " & MyParn(pDoc_No), con)
    If UBound(aRet) > 0 Then
        If aRet(1) = "" Then
            MsgBox "—«Ã⁄  ”ÃÌ· «”„ «·⁄„Ì· "
            Exit Function
        End If
        If aRet(2) = "" Then
            MsgBox "—«Ã⁄  ”ÃÌ·  ·Ì›Ê‰ "
            Exit Function
        End If
        If aRet(3) = "" Then
            MsgBox "—«Ã⁄  ”ÃÌ· «·⁄‰Ê«‰ "
            Exit Function
        End If
        If Val(aRet(4) & "") + Val(aRet(5) & "") = 0 Then
            MsgBox "—«Ã⁄  ”ÃÌ· „’«—Ì› «·‘Õ‰ "
            Exit Function
        End If
        If aRet(6) = "" Then
            MsgBox "—«Ã⁄  ”ÃÌ· —ﬁ„ «·»Ê·Ì’…"
            Exit Function
        End If
        If aRet(7) = "" Then
            MsgBox "—«Ã⁄  ”ÃÌ· ‘—ﬂ… «·‘Õ‰ "
            Exit Function
        End If
    End If
    IsOnLine = True
End Function
Private Sub cmd_addexel_Click()
    If xStore.BoundText = "" Then
        MsgBox " ≈Œ Ì«— «·„Œ“‰ "
        Exit Sub
    End If
    If xCode.text = "" Then
        MsgBox " ≈Œ Ì«— «·⁄„Ì· "
        Exit Sub
    End If
    If Not myreplace Then Exit Sub
    AddFromExel
    Inform "  „ «÷«›… «·«’‰«› "
    myUndo
End Sub
Sub AddFromExel()
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet, nItem As String
    Dim xlwbook As Excel.Workbook, cItem As String
    Dim cFileName As String, cBarCode As String
    Dim nCount As Double, cModel As String, nquant As Double, nPrice As Double, nPriceC As Double
    If MsgBox("≈÷«›… »Ì«‰«  „‰ „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
            With grid1
            i = 1
            Do While True
                i = i + 1
                Me.Caption = i
                If (xlsheet.Cells(i, 1)) <> "" Then
                    nCount = i
                Else
                    Exit Do
                End If
            Loop
            prog1.Visible = True
            prog1.Value = 0
            prog1.Max = nCount
            prog1.Min = 0
            .Rows = 1
            For nRow = 2 To nCount
                If xlsheet.Cells(nRow, 2) <> "" Then
                    cBarCode = (xlsheet.Cells(nRow, 1))
                    cItem = ""
                    If IsNumeric(cBarCode) And Len(cBarCode) <= 7 Then
                        cItem = cBarCode
                    End If
                    Dim loctable As ADODB.Recordset
                    If cItem <> "" Then
                        Set loctable = ItemFind(cItem, con)
                    Else
                        Set loctable = ItemFind_BARCODE(cBarCode, con)
                    End If
                    cItem = ""
                    If Not (loctable.EOF And loctable.BOF) Then
    '                    cItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE  ITEM = " & Val(cBarCode) & " OR BARCODE2 = " & MyParn(cBarCode) & " OR BARCODE13 = " & MyParn(cBarCode) & " OR BARCODE = " & MyParn(cBarCode), con)
    '                    nPriceC = Val(GetDesca("SELECT PRICE FROM FILE1_10 WHERE  ITEM = " & Val(cBarCode) & " OR BARCODE2 = " & MyParn(cBarCode) & " OR BARCODE13 = " & MyParn(cBarCode) & " OR BARCODE = " & MyParn(cBarCode), con) & "")
                        cItem = loctable!Item
                        nPriceC = loctable!price
                        If cItem <> "" Then
                            nquant = Val(xlsheet.Cells(nRow, 2))
                            nPrice = Abs(Val(xlsheet.Cells(nRow, 3)))
                            
                            Dim aInsert(9, 1)
                            aInsert(0, 0) = "doc_no"
                            aInsert(0, 1) = addstring(xDoc_No.text)
                            aInsert(1, 0) = "item"
                            aInsert(1, 1) = addstring(cItem)
                            aInsert(2, 0) = "quant"
                            aInsert(2, 1) = nquant
                            aInsert(3, 0) = "Price"
                            aInsert(3, 1) = nPrice
                            aInsert(4, 0) = "Discount"
                            aInsert(4, 1) = 0
                            aInsert(5, 0) = "PRICE_C"
                            aInsert(5, 1) = nPriceC
                            aInsert(6, 0) = "row"
                            aInsert(6, 1) = nRow
                            aInsert(7, 0) = "COST"
                            aInsert(7, 1) = Item_cost(cItem, con)
                            aInsert(8, 0) = "MAN"
                            aInsert(8, 1) = addstring(xMan.BoundText)
                            aInsert(9, 0) = "S_OKAZ"
                            aInsert(9, 1) = 0
                            con.Execute CreateInsert(aInsert, "FILE6_20")
                        End If
                    Else
                        MsgBox " »«—ﬂÊœ €Ì— ’ŒÌŒ " & cBarCode & "  ”ÿ—  " & nRow
                    End If
                End If
            Next nRow
            MsgBox "«„ «÷«›… «·«’‰«›  "
            End With
        End If
    End If
End Sub
Function CalcDisc3(Optional pDocDisc As String) As Variant
    Dim DiscHTable As New ADODB.Recordset, cDescDisc As String, nQty1 As Double, nQTY2 As Double, nDiscModelRate    As Double
    Dim SubItemTable As New ADODB.Recordset, nQDisc As Double, nTQty As Double, nNewCount As Double, nTDiscInv As Double
    Dim cMaxDoc As String, nMaxDiscout, cMaxDescDisc As String, DiscTable As New ADODB.Recordset, nTTotQty As Double
    ReDim aRet(3)
    con.Execute " DELETE FROM PRICE WHERE DOC_NO = " & MyParn(xDoc_No.text)
    If pDocDisc = "" Then
        cStr1 = " SELECT  FILE0_90H.DOC_NO, FILE0_90H.DESCA, FILE0_90H.QTY1, FILE0_90H.QTY2, FILE0_90H.DISC, FILE0_90H.ISSTOP , SUM(FILE6_20.QUANT) AS TQUANT FROM  FILE0_90H INNER JOIN FILE0_90 ON FILE0_90H.DOC_NO = FILE0_90.DOC_NO INNER JOIN FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM ON FILE0_90.MODELNO = FILE1_10.modelno WHERE  DATE1 <= " & DateSq(xDate.text) & " AND DATE2 >=" & DateSq(xDate.text) & " AND FILE0_90H.ISSTOP = 0  AND QTY2 > 0 AND FILE6_20.DOC_NO = " & MyParn(xDoc_No.text) & _
                " GROUP BY FILE0_90H.DOC_NO, FILE0_90H.DESCA, FILE0_90H.QTY1, FILE0_90H.QTY2, FILE0_90H.DISC , FILE0_90H.ISSTOP   HAVING SUM(FILE6_20.QUANT) >=  (FILE0_90H.QTY1 + FILE0_90H.QTY2 ) "
        DiscHTable.Open cStr1, con, adOpenStatic, adCmdText
    Else
        DiscHTable.Open "SELECT * FROM   FILE0_90H WHERE  DATE1 <= " & DateSq(xDate.text) & " AND DATE2 >=" & DateSq(xDate.text) & " AND QTY2 > 0 AND ISSTOP = 0 AND DOC_NO = " & MyParn(pDocDisc), con, adOpenStatic, adCmdText
    End If
    Do While Not DiscHTable.EOF
        cDescDisc = DiscHTable!DESCA & ""
        nQty1 = DiscHTable!QTY1
        nQTY2 = DiscHTable!QTY2
        nDiscModelRate = DiscHTable!disc
        If SubItemTable.State = adStateOpen Then SubItemTable.Close
        cStr1 = " SELECT FILE0_90.DOC_NO AS DOC_DISC, FILE6_20.DOC_NO , FILE6_20.ITEM, FILE6_20.QUANT, FILE1_10.PRICE FROM  FILE0_90 INNER JOIN FILE1_10 ON FILE0_90.MODELNO = FILE1_10.modelno INNER JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM WHERE FILE6_20.DOC_NO =   " & MyParn(xDoc_No.text) & " AND FILE0_90.DOC_NO = " & MyParn(DiscHTable!doc_no) & " ORDER BY PRICE DESC "
        SubItemTable.Open cStr1, con, adOpenStatic, adLockReadOnly
        nTQty = (nQty1 + nQTY2)
        If Not (SubItemTable.EOF And SubItemTable.BOF) Then
            With SubItemTable
            nTTotQty = 0
            Do While Not SubItemTable.EOF
                nTTotQty = nTTotQty + !Quant
                For nQ = 1 To !Quant
                    nCountQ = nCountQ + 1
                    cString = "Insert into [PRICE]  (doc_no,DOC_DISC,PRICE) " & _
                                       "Values(" & _
                                       addstring(!doc_no) & ", " & _
                                       addstring(!DOC_DISC) & ", " & _
                                       Val(!price) & _
                                       ")"
        
                    con.Execute cString
                     
                Next nQ
                SubItemTable.MoveNext
            Loop
            End With
            
            cStr1 = "SELECT * FROM PRICE WHERE DOC_NO = " & MyParn(xDoc_No.text) & " AND DOC_DISC = " & MyParn(DiscHTable!doc_no) & " ORDER BY PRICE  DESC "
            If DiscTable.State = adStateOpen Then DiscTable.Close
            DiscTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
            If Not (DiscTable.EOF Or DiscTable.BOF) Then
                DiscTable.MoveFirst
                nNewCount = 0
                nTDiscInv = 0
                For i = 1 To 1000
                    nNewCount = nNewCount + 1
                    If nNewCount > nQty1 And nNewCount <= nTQty Then
                        nTDiscInv = nTDiscInv + Val(DiscTable!price * (nDiscModelRate / 100))
                    End If
                    DiscTable.MoveNext
                    If DiscTable.EOF Then Exit For
                    If (i / nTQty) = Int(i / nTQty) Then nNewCount = 0
                Next i
            End If
            If nMaxDiscout < nTDiscInv Then
                nMaxDiscout = nTDiscInv
                cMaxDoc = DiscHTable!doc_no
                cMaxDescDisc = DiscHTable!DESCA
            End If
        End If
        If pDocDisc <> "" Then
            If Int(nTTotQty / nTQty) <> (nTTotQty / nTQty) Then
                MsgBox " »—Ã«¡ “Ì«œ… «·ﬁÿ⁄ œ«Œ· «·⁄—÷ ··≈” ›«œ…"
            End If
        End If
        DiscHTable.MoveNext
    Loop
    
    If nMaxDiscout > 0 Then
        con.Execute " UPDATE FILE6_20 SET  file6_20.price = file1_10.PRICE , file6_20.price_C = file1_10.PRICE , isdisc2_1 = 1  , S_OKAZ = 0 FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM where DOC_NO = " & MyParn(xDoc_No.text)
        myloadgrd
        CalcTotals
    End If
    
    aRet(1) = nMaxDiscout
    aRet(2) = cMaxDoc
    aRet(3) = cMaxDescDisc
    CalcDisc3 = aRet
End Function
'''''''''''''''''''''

Function CalcDisc_Item(Optional pDocDisc As String) As Variant
    Dim DiscHTable As New ADODB.Recordset, cDescDisc As String, nQty1 As Double, nQTY2 As Double, nDiscModelRate    As Double
    Dim SubItemTable As New ADODB.Recordset, nQDisc As Double, nTQty As Double, nNewCount As Double, nTDiscInv As Double
    Dim cMaxDoc As String, nMaxDiscout, cMaxDescDisc As String, DiscTable As New ADODB.Recordset, nTTotQty As Double
    ReDim aRet(3)
    con.Execute " DELETE FROM PRICE WHERE DOC_NO = " & MyParn(xDoc_No.text)
    If pDocDisc = "" Then
        cStr1 = " SELECT  FILE0_90H.DOC_NO, FILE0_90H.DESCA, FILE0_90H.QTY1, FILE0_90H.QTY2, FILE0_90H.DISC, FILE0_90H.ISSTOP , SUM(FILE6_20.QUANT) AS TQUANT FROM  FILE0_90H INNER JOIN FILE0_90 ON FILE0_90H.DOC_NO = FILE0_90.DOC_NO INNER JOIN FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM ON FILE0_90.MODELNO = FILE1_10.modelno WHERE  DATE1 <= " & DateSq(xDate.text) & " AND DATE2 >=" & DateSq(xDate.text) & " AND FILE0_90H.ISSTOP = 0  AND QTY2 =  0 AND FILE6_20.DOC_NO = " & MyParn(xDoc_No.text) & _
                " GROUP BY FILE0_90H.DOC_NO, FILE0_90H.DESCA, FILE0_90H.QTY1, FILE0_90H.QTY2, FILE0_90H.DISC , FILE0_90H.ISSTOP   HAVING SUM(FILE6_20.QUANT) >=  (FILE0_90H.QTY1 + FILE0_90H.QTY2 ) "
        DiscHTable.Open cStr1, con, adOpenStatic, adCmdText
    Else
        DiscHTable.Open "SELECT * FROM   FILE0_90H WHERE  DATE1 <= " & DateSq(xDate.text) & " AND DATE2 >=" & DateSq(xDate.text) & " AND QTY2 =  0 AND ISSTOP = 0 AND DOC_NO = " & MyParn(pDocDisc), con, adOpenStatic, adCmdText
    End If
    Do While Not DiscHTable.EOF
        cDescDisc = DiscHTable!DESCA & ""
        nQty1 = DiscHTable!QTY1
        nDiscModelRate = DiscHTable!disc
        If SubItemTable.State = adStateOpen Then SubItemTable.Close
        cStr1 = " SELECT FILE0_90.DOC_NO AS DOC_DISC, FILE6_20.DOC_NO , SUM(FILE6_20.QUANT) AS QUANT , SUM(FILE6_20.QUANT * FILE1_10.PRICE) AS TOTAL FROM  FILE0_90 INNER JOIN FILE1_10 ON FILE0_90.MODELNO = FILE1_10.modelno INNER JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM WHERE FILE6_20.DOC_NO =   " & MyParn(xDoc_No.text) & " AND FILE0_90.DOC_NO = " & MyParn(DiscHTable!doc_no)
        SubItemTable.Open cStr1, con, adOpenStatic, adLockReadOnly
        nTQty = nQty1
        If Not (SubItemTable.EOF And SubItemTable.BOF) Then
            With SubItemTable
            nTTotQty = 0
            Do While Not SubItemTable.EOF
                nTTotQty = nTTotQty + !Quant
                For nQ = 1 To !Quant
                    nCountQ = nCountQ + 1
                    cString = "Insert into [PRICE]  (doc_no,DOC_DISC,PRICE) " & _
                                       "Values(" & _
                                       addstring(!doc_no) & ", " & _
                                       addstring(!DOC_DISC) & ", " & _
                                       Val(!price) & _
                                       ")"
        
                    con.Execute cString
                     
                Next nQ
                SubItemTable.MoveNext
            Loop
            End With
            
            cStr1 = "SELECT * FROM PRICE WHERE DOC_NO = " & MyParn(xDoc_No.text) & " AND DOC_DISC = " & MyParn(DiscHTable!doc_no) & " ORDER BY PRICE  DESC "
            If DiscTable.State = adStateOpen Then DiscTable.Close
            DiscTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
            If Not (DiscTable.EOF Or DiscTable.BOF) Then
                DiscTable.MoveFirst
                nNewCount = 0
                nTDiscInv = 0
                For i = 1 To 1000
                    nNewCount = nNewCount + 1
                    If nNewCount > nQty1 And nNewCount <= nTQty Then
                        nTDiscInv = nTDiscInv + Val(DiscTable!price * (nDiscModelRate / 100))
                    End If
                    DiscTable.MoveNext
                    If DiscTable.EOF Then Exit For
                    If (i / nTQty) = Int(i / nTQty) Then nNewCount = 0
                Next i
            End If
            If nMaxDiscout < nTDiscInv Then
                nMaxDiscout = nTDiscInv
                cMaxDoc = DiscHTable!doc_no
                cMaxDescDisc = DiscHTable!DESCA
            End If
        End If
        If pDocDisc <> "" Then
            If Int(nTTotQty / nTQty) <> (nTTotQty / nTQty) Then
                MsgBox " »—Ã«¡ “Ì«œ… «·ﬁÿ⁄ œ«Œ· «·⁄—÷ ··≈” ›«œ…"
            End If
        End If
        DiscHTable.MoveNext
    Loop
    If nMaxDiscout > 0 Then
        con.Execute " UPDATE FILE6_20 SET  file6_20.price = file1_10.PRICE , file6_20.price_C = file1_10.PRICE , isdisc2_1 = 1  , S_OKAZ = 0 FROM   FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM where DOC_NO = " & MyParn(xDoc_No.text)
        myloadgrd
        CalcTotals
    End If
    aRet(1) = nMaxDiscout
    aRet(2) = cMaxDoc
    aRet(3) = cMaxDescDisc
    CalcDisc3 = aRet
End Function
Function CalcDiscRateDoc() As Boolean
    Dim DiscDocTable As New ADODB.Recordset, nRate As Double, lDiscOk As Boolean, DiscDocHTable As New ADODB.Recordset, cDoc_Disc As String
    cStr1 = " SELECT   SUM(FILE6_20.QUANT) AS Qty_Disc , FILE0_90H.DOC_NO, FILE0_90H.desca , FILE0_90H.QTY1, FILE0_90H.DISC, FILE0_90.DISC2, FILE6_20.DOC_NO AS DOC_SAL " & _
            " FROM     FILE0_90H INNER JOIN FILE0_90 ON FILE0_90H.DOC_NO = FILE0_90.DOC_NO INNER JOIN FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM ON FILE0_90.MODELNO = FILE1_10.modelno WHERE        (FILE6_20.DOC_NO = " & MyParn(xDoc_No.text) & ") AND (FILE0_90H.QTY2 = 0) AND (FILE0_90H.ISSTOP = 0) AND FILE0_90H.DATE1 <= " & DateSq(xDate.text) & " AND FILE0_90H.DATE2 >= " & DateSq(xDate.text) & _
            " AND (SELECT SUM(QUANT) FROM Q_DOC_SALES_DISC WHERE DOC_SALES = FILE6_20.DOC_NO AND DOC_NO = FILE0_90H.DOC_NO ) >= QTY1 " & _
            " GROUP BY FILE0_90H.DOC_NO, FILE0_90H.desca , FILE0_90H.QTY1, FILE0_90H.DISC, FILE0_90.DISC2, FILE6_20.DOC_NO " & _
            " HAVING SUM(FILE6_20.QUANT) >= FILE0_90H.QTY1 ORDER BY QTY1 DESC"
    DiscDocHTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    If DiscDocHTable.RecordCount > 0 Then cDoc_Disc = DiscDocHTable!doc_no
    If cDoc_Disc = "" Then Exit Function
    
    cStr1 = " SELECT   FILE6_20.ITEM, FILE6_20.QUANT , FILE6_20.PRICE, FILE0_90H.DOC_NO, FILE0_90H.desca , FILE0_90H.QTY1, FILE0_90H.DISC, FILE0_90.DISC2, FILE6_20.DOC_NO AS DOC_SAL " & _
            " FROM     FILE0_90H INNER JOIN FILE0_90 ON FILE0_90H.DOC_NO = FILE0_90.DOC_NO INNER JOIN FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM ON FILE0_90.MODELNO = FILE1_10.modelno " & _
            " WHERE  FILE0_90H.DOC_NO = " & MyParn(cDoc_Disc) & " AND FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)
    DiscDocTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    con.Execute " UPDATE FILE6_20 SET FILE6_20.PRICE = FILE6_20.PRICE_C2 , FILE6_20.S_OKAZ = FILE6_20.S_OKAZ_2 WHERE DOC_NO =  " & MyParn(xDoc_No.text)
    myloadgrd
    
    With DiscDocTable
        xDOC_DISC.Caption = DiscDocTable!doc_no
        xNOTE_DISC.Caption = DiscDocTable!DESCA
        con.Execute " update file6_20h set doc_disc = " & addstring(xDOC_DISC.Caption) & " , note_disc = " & addstring(xNOTE_DISC.Caption) & " where doc_no = " & MyParn(xDoc_No.text)
        
        Do While Not .EOF
            CalcDiscRateDoc = True
'            nRate = ((100 - Val(!disc & "")) / 100)
            nRate = ((100 - Val(!disc2 & "")) / 100)
            con.Execute " UPDATE FILE6_20 SET FILE6_20.S_OKAZ =  " & Val(!disc2) & " , FILE6_20.PRICE = FILE1_10.PRICE * " & nRate & "  FROM  FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE DOC_NO =  " & MyParn(xDoc_No.text) & " AND FILE6_20.ITEM = " & MyParn(!Item)
            .MoveNext
        Loop
    End With

    If nDiscBranch > 0 And IsDate(dDiscBranch1) And IsDate(dDiscBranch1) Then
        If dSalesDate >= dDiscBranch1 And dSalesDate <= dDiscBranch2 Then
            nRate = ((100 - Val(nDiscBranch & "")) / 100)
             
            con.Execute " UPDATE FILE6_20 SET FILE6_20.S_OKAZ =  " & Val(nDiscBranch) & " , FILE6_20.PRICE = FILE1_10.PRICE * " & nRate & "  FROM  FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE DOC_NO =  " & MyParn(xDoc_No.text) & " AND FILE6_20.S_OKAZ = 0 ", nRec
              
            If nRec > 0 Then
                MsgBox " ÌÊÃœ Œ’„ Œ«’ »«·›—⁄ "
                xNOTE_DISC.Caption = "Œ’„ Œ«’ »«·›—⁄ "
                myloadgrd
            End If
        End If
    End If



End Function
Private Sub xfilter2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter2.text, 5
    End If
End Sub
Private Sub xfilter3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter3.text, 3
    End If
End Sub
Private Sub xfilter4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter4.text, 6
    End If
End Sub
Private Sub xfilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter.text, 1
    End If
End Sub
Sub AddSales_Def()
On Error GoTo myerror
    cString = "Insert into FILE6_20H (doc_no,CODE,[DATE],STORE,BOX,BRANCH,MAN,PHONE,BANK,USERID,CASH,VISA)" & _
            "Values(" & _
             addstring(xDoc_No.text) & "," & _
             addstring(xCode.text) & "," & _
             addDate(xDate.text) & "," & _
             addstring(xStore.text) & "," & _
             addstring(xBox.text) & "," & _
             addstring(cBranch) & "," & _
             addstring(xMan.text) & "," & _
             addstring(xphone.Caption) & "," & _
             addstring(XCODEVISA.text) & "," & _
             addstring(xIpName.Caption) & "," & _
             Val(xcash.Caption) & "," & _
             Val(xvisa.Caption) & _
            ")"
    condef.Execute cString
    
    condef.Execute " DELETE FROM FILE6_20 WHERE DOC_NO = " & MyParn(xDoc_No.text)
    With grid1
    For nRow = 1 To .Rows - 2
        cString = "Insert into FILE6_20 (doc_no,ITEM,MAN,QUANT,PRICE,PRICE_C,PRICE_C2,TOTAL)" & _
            "Values(" & _
             addstring(xDoc_No.text) & "," & _
             addstring(.TextMatrix(.Row, 1)) & "," & _
             addstring(.TextMatrix(.Row, 16)) & "," & _
             Val(.TextMatrix(.Row, 10)) & "," & _
             Val(.TextMatrix(.Row, 11)) & "," & _
             Val(.TextMatrix(.Row, 9)) & "," & _
             Val(.TextMatrix(.Row, 14)) & "," & _
             Val(.TextMatrix(.Row, 13)) & _
            ")"
        condef.Execute cString
    Next nRow
    End With
    Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub ClearTemp6_20()
On Error GoTo myerror
Dim fs As New FileSystemObject
condef.Close
fs.CopyFile App.Path & "\DEF.mdb", "c:\TempMrshd\def.MDB"
condef.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & "c:\TempMrshd\def.MDB"
Exit Sub
myerror:
Err.Clear
End Sub
Private Sub cmd_showdoc_Click()
    Dim pDocRet As String, cRetBranch As String
    Dim dDateRet As Date
    Dim cPassword As String, aUser As Variant
        If grid1.Rows > 2 Then
            MsgBox " ÌÊÃœ «À‰«› „”Ã·… ›Ï «·»Ê‰"
            Exit Sub
        End If
        pDocRet = xdoc_ret.text
        xdoc_ret.text = pDocRet
        ManagPass.pPassWord = ""
        ManagPass.Show 1
        cPassword = ManagPass.pPassWord
        If Trim(cPassword) = "" Then Exit Sub
        aUser = aGetDesca("SELECT CODE , DESCA FROM SUB_USER WHERE ( ( PASSWORD = " & MyParn(cPassword) & " AND STOP1 = 0 ) OR ( STOP2 = 0 AND PASSWORD2 = " & MyParn(cPassword) & "))", con)
        
        AddLod_Data cusername, 0, " › Õ »Ê‰ ··„— Ã⁄  ", con, xDoc_No.text, xDate.text, , xCodeDesca.Caption
        
        
        If UBound(aUser) > 0 Then
            lRetDoc = True
            XISRET.Value = 1
            xusername_RET.Caption = aUser(2)
        Else
            Exit Sub
        End If
        SalesDoc.sDoc_no = xdoc_ret.text
        SalesDoc.Show 1
        If xinv_no.text <> "" Then
            cmdSave_Click
        End If
End Sub
Private Function LoadConString_OnLine()
Dim cServerName As String, cUserId As String, cPassword As String, sCatalog_Online As String
    cServerName = "junior-sql.database.windows.net"
    cUserId = decrypt("062F8C407C77", "dr")
    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
    sCatalog_Online = "JUNIOR_SHOP"
    LoadConString_OnLine = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
        & "catalog=" & sCatalog_Online & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Private Function doprint_online() As Boolean
Dim aHeader(2)
If Not MYVALID Then Exit Function
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT FILE6_20.ITEM,FILE6_20.QUANT,FILE6_20.SKU  ,FILE6_20.PRICE,FILE1_10.DESCA AS ITEM_DESCA,FILE6_20.DESCA2 , FILE6_20.TOTAL,FILE6_20H.TIME, file1_10.modelfact0 , FACT.DESCA AS FACTDESCA , FILE1_10.COLOR , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.MODEL   , file6_20h.[name] , file6_20h.phone , file6_20h.address , file6_20h.street , file6_20h.Payment_Method , CITY , Shipping_City , ONLINE_DOC  , CHARGE1 , CHARGE2 , ship_no FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM LEFT JOIN FACT ON FACT.CODE = FILE1_10.[FACT]  "
cString = cString & turn(cString) & "FILE6_20.DOC_NO = " & MyParn(xDoc_No.text)
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Do While Not loctable.EOF
    temptable.AddNew
    temptable!str16 = xDoc_no2.text
    temptable!str6 = loctable!ONLINE_DOC
    temptable!str12 = xStore.text
    temptable!Date1 = xDate.text
    
    temptable!STR7 = loctable!Name
    temptable!str8 = loctable!phone
    temptable!str9 = loctable!Shipping_City
    temptable!str9 = loctable!Shipping_City
    temptable!str14 = loctable!city
    temptable!str15 = loctable!Street
    
    temptable!str2 = loctable!modelfact0
    temptable!str3 = IIf(loctable!Desca2 <> "", loctable!Desca2, loctable!ITEM_DESCA & " " & loctable!color & " " & loctable!scal)
    
    temptable!val3 = loctable!Quant
    temptable!VAL1 = loctable!price
    temptable!val2 = loctable!TOTAL
    
    temptable!val5 = Val(xDiscount.text)
    temptable!VAL7 = Val(xTotal.text)
    temptable!val11 = Val(loctable!CHARGE1 & "") + Val(loctable!charge2 & "")
    temptable!val12 = Val(xTotal.text) + Val(loctable!CHARGE1 & "") + Val(loctable!charge2 & "")
    temptable!VAL6 = Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption)
    temptable!str10 = loctable!SHIP_NO
    temptable!STR20 = xDoc_No.text
    
    If Val(xTotal.text) + Val(loctable!CHARGE1 & "") + Val(loctable!charge2 & "") > 0 Then
        temptable!str11 = MyOnly(Val(xTotal.text) + Val(loctable!CHARGE1 & "") + Val(loctable!charge2 & ""))
    End If
    
    temptable!str4 = TurnValue(cId_NoDesca1)
    temptable!str5 = TurnValue(cId_NoDesca2)
    temptable!str18 = TurnValue(cHeadBoon)
    temptable!str17 = loctable!SKU
    
    temptable!VAL14 = i
    temptable.Update
    loctable.MoveNext
Loop

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

contemp.BeginTrans
contemp.CommitTrans

Main.REPORT1.ReportFileName = App.Path & "\Reports\Print_SALES_ONLINE.rpt"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1

temptable.Close
Set temptable = Nothing
doprint_online = True
End Function
Private Sub mySendInvoices()
Dim sError As String
Dim nCount As Long
Dim bDone As Long
If cBranch = "051" Then
    bDone = SendInvoicesArab(con, "051", sError, nCount)
    If sError <> "" Then
        MsgBox sError
    End If
    If nCount > 0 Then
        Inform " „ «—”«· " & nCount & " ›« Ê—… „»Ì⁄«  ·«œ«—… «·„Ê·"
    ElseIf nCount = 0 And bDone Then
        Inform "·« ÌÊÃœ «Ì ›Ê« Ì— ·«—”«·Â«"
    End If
End If
End Sub
Function CheckCardDisc() As Double
    Dim lConServerDisc As Boolean
    If xcard_disc.text = "" Then Exit Function
    If Not lIsBranchStore Then
        strConShop = LoadConStringshop
        If openConShop(conShop) <> "ok" Then
            MsgBox "·« ÌÊÃœ « ’«· »«·”Ì—›—"
            lConServerDisc = False
            Exit Function
        Else
            lConServerDisc = True
            Inform "  „ «·≈ ’«· »«·”Ì—›— "
        End If
    Else
        strConShop = LoadConString
        If openConShop(conShop) = "ok" Then lConServerDisc = True
    End If
    If lConServerDisc Then
        aRet = aGetDesca("SELECT DOC_NO ,  VALUE , SALES , DATE2 , DESCA ,  SALES_DOC_NO FROM QCARD_DISC WHERE BARCODE_DISC = " & MyParn(xcard_disc.text), conShop)
        If UBound(aRet) > 0 Then
            If aRet(6) <> "" Then
                MsgBox " ﬂ«—  „” Œœ„ „‰ ﬁ»·  "
                xcard_desca.Caption = ""
                xcard_disc.text = ""
                xcard_doc.Caption = ""
                xcard_value.Caption = ""
                CheckCardDisc = 0
                Exit Function
            End If
            If Val(xTotalItemNoDiscount.Caption & "") < Val(aRet(3) & "") Then
                MsgBox " ﬁÌ„… «·„»Ì⁄«  ÌÃ» «‰  ﬂÊ‰ «ﬂ»— „‰  " & Val(aRet(3))
                xcard_desca.Caption = ""
                xcard_disc.text = ""
                xcard_doc.Caption = ""
                xcard_value.Caption = ""
                CheckCardDisc = 0
            Else
                xcard_desca.Caption = aRet(5) & ""
                xcard_doc.Caption = aRet(1) & ""
                xcard_value.Caption = aRet(2) & ""
                CheckCardDisc = aRet(2) & ""
            
                con.Execute " UPDATE FILE6_20 SET FILE6_20.PRICE = FILE6_20.Price_C2 , FILE6_20.S_OKAZ = FILE6_20.S_OKAZ_2 WHERE DOC_NO =  " & MyParn(xDoc_No.text)
                myloadgrd
            End If
        Else
            MsgBox "—ﬁ„ ﬂ«—  €Ì— ’«·Õ"
            xcard_desca.Caption = ""
            xcard_disc.text = ""
            xcard_doc.Caption = ""
            xcard_value.Caption = ""
            CheckCardDisc = 0
        End If
    End If
End Function
Sub SAveCardDisc_No()
    Dim lConServerDisc As Boolean
    If Not lIsBranchStore Then
        strConShop = LoadConStringshop
        If openConShop(conShop) <> "ok" Then
            lConServerDisc = False
            Exit Sub
        Else
            lConServerDisc = True
            conShop.Execute " UPDATE CARD_DISC SET SALES_DOC_NO = " & addstring(xDoc_No.text) & " WHERE BARCODE_DISC = " & MyParn(xcard_disc.text)
        End If
    End If
End Sub
Private Function DoBackUp() As Boolean
On Error GoTo myerror
Dim fs As FileSystemObject, cDir As String, cFileName As String, aDrive As Variant, cBackDrive  As String
Set fs = CreateObject("Scripting.FileSystemObject")
cBackDrive = GetDesca("SELECT BACK FROM ADDRESS", con)
If cBackDrive = "" Then cBackDrive = "D"
'aDrive = aLastDrive(False)
'cDir = retFlag(aDrive, "LETTER") & ":\DataBackup"
cFileName = cBackDrive & ":\" & sCatalog & "_BACK.bak"
createBackUp (cFileName)
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function createBackUp(pFileName) As Boolean
Dim cFile As String
Dim con As New ADODB.Connection
openCon con

Dim cmd
Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
cString = "BACKUP DATABASE " & sCatalog & " TO  DISK = N'" & pFileName & "' WITH  RETAINDAYS = 1, NOFORMAT, INIT,  NAME = N'over-Full Database Backup', SKIP,  NOREWIND, NOUNLOAD, STATS = 10"
FS1.CommandText = cString
FS1.CommandTimeout = 6000
FS1.Execute
Set FS1 = Nothing
closeCon con
createBackUp = True
End Function

