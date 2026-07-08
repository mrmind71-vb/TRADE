VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form purchasefrm 
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
   ScaleHeight     =   11100
   ScaleWidth      =   20310
   Visible         =   0   'False
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
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
      Left            =   9315
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   116
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   8415
      UseMaskColor    =   -1  'True
      Width           =   1050
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
      Left            =   8280
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   115
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   8415
      UseMaskColor    =   -1  'True
      Width           =   1050
   End
   Begin VB.CommandButton CMD_NOTDEM 
      BackColor       =   &H0080FFFF&
      Caption         =   "„ÊœÌ·«  ·Ì” ·Â« ÿ·»Ì…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8325
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   114
      Top             =   7650
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton reprice_inv 
      BackColor       =   &H00FFFFC0&
      Caption         =   " ⁄œÌ· «”⁄«— «·›« Ê—…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   8325
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   113
      Top             =   6975
      Width           =   1905
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "»œÊ‰  ÕœÌÀ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   112
      Top             =   1350
      Width           =   735
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
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   110
      Top             =   0
      Width           =   4110
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   2355
         Left            =   90
         TabIndex        =   111
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
         LicenseUserName =   "AMR01"
         LicenseRegCode  =   "íß“≠Ω≤•™µΩÆ“Æ®∫´Ω•ÆΩNZOV-HRVEOHJJ-EPOP6gI"
      End
   End
   Begin VB.TextBox XMODELNOT 
      Alignment       =   2  'Center
      Height          =   1365
      Left            =   4905
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   109
      Text            =   "purchase3.frx":0000
      Top             =   6930
      Width           =   1455
   End
   Begin VB.CommandButton cmd_exRET 
      BackColor       =   &H00FFFFC0&
      Caption         =   "„— Ã⁄ „‰ «·≈ﬂ”Ì·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   15255
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   108
      Top             =   8415
      Width           =   1590
   End
   Begin VB.CommandButton FIX_PRICE 
      BackColor       =   &H00FF8080&
      Caption         =   " ⁄œÌ· «”⁄«— „ÊœÌ·«  «·›« Ê—… ·«”⁄«—«·„’‰⁄"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   10395
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   104
      Top             =   8415
      Width           =   1500
   End
   Begin VB.Frame Frame14 
      Caption         =   "„” ‰œ  ÕÊÌ· „‰ «·›—⁄ "
      Height          =   600
      Left            =   2295
      RightToLeft     =   -1  'True
      TabIndex        =   102
      Top             =   135
      Visible         =   0   'False
      Width           =   1860
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
         Left            =   90
         TabIndex        =   103
         Top             =   225
         Width           =   1545
      End
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   90
      Top             =   765
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CMD_ADDINV 
      BackColor       =   &H0080FF80&
      Caption         =   "«” Ì—«œ „‰ «·„’‰⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11925
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   99
      Top             =   8415
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "send items"
      Height          =   510
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   98
      Top             =   2070
      Visible         =   0   'False
      Width           =   780
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2625
      Left            =   90
      TabIndex        =   86
      Top             =   6840
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   4630
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "SEND"
      TabPicture(0)   =   "purchase3.frx":0008
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame11 
         Height          =   2175
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   375
         Width           =   4560
         Begin VB.TextBox XTIMESEND 
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
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   93
            Top             =   180
            Width           =   2400
         End
         Begin VB.TextBox XUSERSEND 
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
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   92
            Top             =   540
            Width           =   2400
         End
         Begin VB.TextBox XUSERACCEPT 
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
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   91
            Top             =   1260
            Width           =   2400
         End
         Begin VB.TextBox XTIMEACCEPT 
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
            Left            =   1755
            Locked          =   -1  'True
            MaxLength       =   10
            RightToLeft     =   -1  'True
            TabIndex        =   90
            Top             =   900
            Width           =   2400
         End
         Begin VB.CheckBox XReceived 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Received"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3375
            RightToLeft     =   -1  'True
            TabIndex        =   88
            Top             =   1710
            Width           =   1050
         End
         Begin Threed.SSCommand CMD_RECEVED 
            Height          =   465
            Left            =   45
            TabIndex        =   89
            Top             =   1620
            Width           =   3300
            _ExtentX        =   5821
            _ExtentY        =   820
            _Version        =   196610
            BackColor       =   33023
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arabic Transparent"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   " „  „—«Ã⁄… «·„” ‰œ Ê »Ì«‰ «·„” ‰œ ’ÕÌÕ"
            Alignment       =   8
            PictureAlignment=   6
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Send Time"
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
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   97
            Top             =   225
            Width           =   1140
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Send User"
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
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   96
            Top             =   540
            Width           =   1125
         End
         Begin VB.Label Label19 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receive Time"
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
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   95
            Top             =   900
            Width           =   1470
         End
         Begin VB.Label Label20 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Receive User"
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
            Left            =   105
            RightToLeft     =   -1  'True
            TabIndex        =   94
            Top             =   1260
            Width           =   1455
         End
      End
   End
   Begin VB.CommandButton cmd_ex 
      BackColor       =   &H00FFFFC0&
      Caption         =   " «” Ì—«œ „‰ «·«ﬂ”Ì· IFS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   13455
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   85
      Top             =   8415
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.Frame Frame13 
      Height          =   555
      Left            =   13185
      RightToLeft     =   -1  'True
      TabIndex        =   80
      Top             =   2070
      Width           =   6990
      Begin VB.CommandButton CMD_FIXSEC 
         Caption         =   " ⁄œÌ· «·„ÊœÌ·«  ··ﬁ”„"
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
         Left            =   90
         TabIndex        =   83
         Top             =   135
         Width           =   2220
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   2385
         TabIndex        =   81
         Top             =   180
         Width           =   3750
         _ExtentX        =   6615
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
         Left            =   6255
         RightToLeft     =   -1  'True
         TabIndex        =   82
         Top             =   225
         Width           =   570
      End
   End
   Begin VB.Frame Frame12 
      Height          =   645
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   8370
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
         TabIndex        =   79
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
         TabIndex        =   78
         Top             =   180
         Width           =   705
      End
   End
   Begin VB.CommandButton cmd_barcode_fix 
      Caption         =   "Ÿ»ÿ »«—ﬂÊœ «·„’‰⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   9000
      TabIndex        =   76
      Top             =   2160
      Width           =   1770
   End
   Begin VB.Frame Frame10 
      Caption         =   "‰”»… Œ’„ ·ﬂ· „ÊœÌ·«  «·›« Ê—…"
      Height          =   645
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   73
      Top             =   1980
      Width           =   2580
      Begin VB.TextBox xDiscAll 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1440
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   225
         Width           =   1080
      End
      Begin VB.CommandButton CM_OK 
         Caption         =   "Œ’„ ·ﬂ· «·›« Ê—…"
         Height          =   420
         Left            =   45
         TabIndex        =   74
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1860
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   45
      Width           =   2085
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         TabIndex        =   69
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
         Picture         =   "purchase3.frx":0024
         Alignment       =   4
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   1035
         TabIndex        =   70
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
         Picture         =   "purchase3.frx":25F0
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   90
         TabIndex        =   71
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
         Picture         =   "purchase3.frx":4C68
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   16875
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   8370
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   60
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
         Picture         =   "purchase3.frx":72A9
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase3.frx":9479
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   61
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
         Picture         =   "purchase3.frx":B5C1
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase3.frx":D789
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   62
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
         Picture         =   "purchase3.frx":F8D8
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase3.frx":11AB8
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   63
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
         Picture         =   "purchase3.frx":13C13
         Caption         =   " «Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase3.frx":15DCF
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   58
      Top             =   10725
      Width           =   20310
      _ExtentX        =   35825
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
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   0
      Width           =   5415
      Begin Threed.SSCommand CMD_PRINT 
         Height          =   510
         Left            =   3780
         TabIndex        =   29
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
         Picture         =   "purchase3.frx":17F1E
         Caption         =   "ÿ»«⁄… «·›« Ê—…"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand CMD_PrintBar 
         Height          =   510
         Left            =   90
         TabIndex        =   30
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
         Picture         =   "purchase3.frx":1A32D
         Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand cmdBarCode 
         Height          =   510
         Left            =   1905
         TabIndex        =   31
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
         Picture         =   "purchase3.frx":1CCAF
         Caption         =   " ÕÊÌ· ··»«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1230
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   720
      Width           =   1365
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase3.frx":1F648
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   49
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
         Picture         =   "purchase3.frx":21BC1
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   48
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
      TabIndex        =   42
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton CmdDelInv 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1275
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase3.frx":23F24
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   46
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
         Picture         =   "purchase3.frx":267BE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   3735
         Picture         =   "purchase3.frx":28D6A
         Style           =   1  'Graphical
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "purchase3.frx":2B53D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   43
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
      TabIndex        =   41
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
         TabIndex        =   50
         Top             =   100
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   720
      Width           =   11175
      Begin VB.CommandButton CMD_CHANGE 
         Caption         =   "..."
         Height          =   285
         Left            =   10710
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   540
         Width           =   420
      End
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
         Left            =   5400
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   855
         Visible         =   0   'False
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
         Left            =   5985
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   100
         Top             =   855
         Width           =   3975
      End
      Begin VB.CheckBox XONEST 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   "»÷«⁄… «„«‰« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   135
         Width           =   1680
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
         TabIndex        =   4
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
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   -225
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   1485
         TabIndex        =   5
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
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   1485
         TabIndex        =   2
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
         Picture         =   "purchase3.frx":2D9A9
         Caption         =   " ”ÃÌ· „ÊœÌ·«  "
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_ret 
         Height          =   375
         Left            =   -990
         TabIndex        =   67
         Top             =   1125
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   661
         _Version        =   196610
         BackColor       =   8421631
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "»Ì«‰«  «·‘Õ‰"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Œ“‰ :"
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
         TabIndex        =   72
         Top             =   900
         Width           =   660
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
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   855
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
         TabIndex        =   54
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
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   855
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   225
         Width           =   555
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   375
      Left            =   4890
      TabIndex        =   17
      Top             =   9045
      Visible         =   0   'False
      Width           =   15330
      _ExtentX        =   27040
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
      TabIndex        =   10
      Top             =   6930
      Width           =   9915
      Begin VB.CommandButton sned_purch 
         BackColor       =   &H0080FF80&
         Caption         =   " ÕÊÌ· ·„‘ —Ì«  JUNIOR"
         Height          =   465
         Left            =   8415
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   107
         Top             =   900
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.CommandButton cmd_fixcost 
         Caption         =   "FIX COST"
         Height          =   420
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   945
         Visible         =   0   'False
         Width           =   1275
      End
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
         TabIndex        =   24
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   57
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
         TabIndex        =   56
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
      TabIndex        =   9
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
         TabIndex        =   53
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
         TabIndex        =   52
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
   Begin Threed.SSCommand CMD_ALL 
      Height          =   600
      Left            =   4905
      TabIndex        =   64
      Top             =   8415
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1058
      _Version        =   196610
      CaptionStyle    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ⁄œÌ· „ÊœÌ·«  «·›« Ê—…"
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4125
      Left            =   135
      TabIndex        =   7
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
   Begin VB.CommandButton FIX_FACTITEM 
      Caption         =   " ÕœÌÀ «·»«—ﬂÊœ ··„’‰⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   11115
      TabIndex        =   84
      Top             =   2160
      Width           =   2085
   End
   Begin VB.Label lblBarCode 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      Caption         =   "·„ Ì „ «· ÕÊÌ· ··»«—ﬂÊœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2250
      RightToLeft     =   -1  'True
      TabIndex        =   66
      Top             =   90
      Visible         =   0   'False
      Width           =   1995
   End
End
Attribute VB_Name = "Purchasefrm"
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

Dim bEdit As Boolean, pGrCust As String
Dim cFileClient, cFieldClient, cCodeDesca As String
Dim formMode, dDateLast As String
Dim bUpdated As Boolean
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Private Function myreplace(Optional bAddModel As Boolean = False) As Boolean
Dim aInsert(17, 1), nR_Tax As Double, nR_DISC As Double, nRCost As Double
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xdoc_no.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(xDate.Text)

aInsert(2, 0) = "Code"
aInsert(2, 1) = addstring(xCode.Text)

aInsert(3, 0) = "Discount"
aInsert(3, 1) = Val(xdiscount.Text)

aInsert(4, 0) = "Tax"
aInsert(4, 1) = Val(xTax.Text)

aInsert(5, 0) = "Store"
aInsert(5, 1) = addstring(XSTORE.BoundText)

aInsert(6, 0) = "MOSM"
aInsert(6, 1) = addstring(xMosm.BoundText)

aInsert(7, 0) = "INV_NO"
aInsert(7, 1) = addstring(xinv_no.Text)

aInsert(8, 0) = "ONEST"
aInsert(8, 1) = addstring(XONEST.Value)

aInsert(9, 0) = "ISDATE"
aInsert(9, 1) = "GETDATE()"

aInsert(10, 0) = "ISNEW1"
aInsert(10, 1) = 1

aInsert(12, 0) = "Notes_Doc"
aInsert(12, 1) = addstring(XNotes_Doc.Text)

aInsert(13, 0) = "TIMESEND"
aInsert(13, 1) = "GETDATE()"

aInsert(14, 0) = "USERSEND"
aInsert(14, 1) = addstring(cusername)


If Val(xTotalNoTax.Caption) <> 0 Then nR_Tax = Round(Val(xTax.Text) / Val(xTotalNoTax.Caption) * 100, 2)
aInsert(16, 0) = "R_TAX"
aInsert(16, 1) = nR_Tax

If Val(xTotalItem.Caption) <> 0 Then nR_DISC = Round(Val(xdiscount.Text) / Val(xTotalItem.Caption) * 100, 2)
aInsert(17, 0) = "R_DISC"
aInsert(17, 1) = nR_DISC

con.BeginTrans
If xdoc_no.Tag = DefineMode Then
    If myPublic = 0 Or myPublic = 1 Then
        xdoc_no.Text = RetZero(Newflag(cFileHeader, "doc_no"))
    Else
        xdoc_no.Text = RetZero(Newflag_PurchBr(cFileHeader, "doc_no", XSTORE.BoundText))
    End If
    aInsert(0, 1) = addstring(xdoc_no.Text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
   con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & MyParn(xdoc_no.Text))
End If
If Check1.Value = 0 Then
    If myPublic = 0 Then
        con.Execute " DELETE FROM NEW7_20H WHERE DOC_NO = " & addstring(xdoc_no.Text)
        For nRow = 2 To nCountBranch
            sStore = RetZero(nRow, 2)
            con.Execute "  INSERT INTO NEW7_20H (DOC_NO , BRANCH) VALUES        ( " & addstring(xdoc_no.Text) & " , " & addstring(sStore) & " )  "
        Next nRow
    ElseIf myPublic = 1 Then
        con.Execute " DELETE FROM NEW7_10H WHERE DOC_NO = " & addstring(xdoc_no.Text)
        For nRow = 2 To nCountBranch
            sStore = RetZero(nRow, 2)
            con.Execute "  INSERT INTO NEW7_10H (DOC_NO , BRANCH) VALUES        ( " & addstring(xdoc_no.Text) & " , " & addstring(sStore) & " )  "
        Next nRow
    End If
End If

'If XONEST.Value = 1 Then
'    cString = "UPDATE " & cFile & " SET TRUST_MODEL = 1"
'    cString = cString & turn(cString) & "DOC_NO = " & MyParn(xdoc_no.Text)
'    con.Execute cString
'End If

If bAddModel Then
    myreplaceGrd
    myReplaceOrder
End If

If myPublic = 0 Then
    cStr1 = " UPDATE FILE1_10 SET FILE1_10.COSTITEM = [FILE7_20].[PRICE]*(1-([FILE7_20].[DISCOUNT ]/100)) FROM   FILE1_10 LEFT JOIN FILE7_20 ON FILE1_10.ITEM = FILE7_20.ITEM  WHERE DOC_NO = " & MyParn(xdoc_no.Text)
    con.Execute cStr1
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
    xdoc_no.Text = Search1.grid1.TextMatrix(Search1.grid1.Row, 0)
    Search1.Hide
    myUndo
ElseIf ActiveControl.Name = CMD_ADDINV.Name Then
    If xCode.Text = "001" And XSTORE.BoundText <> "" Then
        If xdoc_no.Enabled Then
            If Not mysave Then Exit Sub
        End If
        AddInvFact Search_FACT.grid1.TextMatrix(Search_FACT.grid1.Row, 0)
        Unload Search_FACT
    Else
        MsgBox " «·„Ê—œ €Ì— ’ÕÌÕ - «Œ Ì«— «·›—⁄ "
    End If
ElseIf ActiveControl.Name = CMD_CHANGE.Name Then
    xCode.Text = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
ElseIf TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    Unload oSearchSup
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CMD_FIX_Click()
    FIXINV.Show 1
End Sub
Private Sub CM_OK_Click()
If xDiscAll.Text <> "" Then
    If MsgBox(" ⁄œÌ· Œ’„ «·„ÊœÌ·«  ·ﬂ· «·›« Ê—…", vbOKCancel) = vbOK Then
        If myPublic = 0 Then
            con.Execute " UPDATE FILE7_20 SET DISCOUNT = " & Val(xDiscAll.Text) & " WHERE DOC_NO = " & MyParn(xdoc_no.Text)
            cStr1 = " UPDATE FILE1_10 SET FILE1_10.COSTITEM = [FILE7_20].[PRICE]*(1-([FILE7_20].[DISCOUNT ]/100)) FROM   FILE1_10 LEFT JOIN FILE7_20 ON FILE1_10.ITEM = FILE7_20.ITEM  WHERE DOC_NO = " & MyParn(xdoc_no.Text)
            con.Execute cStr1
        Else
            con.Execute " UPDATE " & cFile & " SET DISCOUNT = " & Val(xDiscAll.Text) & " WHERE DOC_NO = " & MyParn(xdoc_no.Text)
        End If
    End If
End If
myUndo
'If bMsg Then Inform " „  ⁄œÌ· «·Œ’„ »‰Ã«Õ"
myload
End Sub
Private Sub CMD_ALL_Click()
    If myPublic = 0 Or myPublic = 2 Then
        FIXINV.Show 1
        myloadgrd
    End If
End Sub
Private Sub CMD_CHANGE_Click()
    SuppLookupAll Me, oSearchSup
End Sub
Private Sub cmd_fixcost_Click()
Dim SalTable As New ADODB.Recordset
Dim nRate As Double, nCount As Double, i As Double

If InputBox("", "") <> "123456" Then Exit Sub
prog1.Visible = True
Select Case myPublic
    Case 0
        cStr1 = "SELECT FILE7_20H.DOC_NO, FILE7_20H.TAX  ,SUM(FILE7_20.TOTAL) AS TOTAL , FILE7_20H.DISCOUNT FROM         FILE7_20H INNER JOIN FILE7_20 ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO where ( FILE7_20H.TAX + FILE7_20H.DISCOUNT ) <> 0 GROUP BY FILE7_20H.DOC_NO, FILE7_20H.discount,TAX "
        nCount = GetDesca("SELECT COUNT(DOC_NO) FROM FILE7_20H WHERE ( FILE7_20H.TAX + FILE7_20H.DISCOUNT ) <> 0 ")
    Case 1
        cStr1 = "SELECT FILE7_10H.DOC_NO, FILE7_10H.TAX  ,SUM(FILE7_10.TOTAL) AS TOTAL , FILE7_10H.DISCOUNT FROM         FILE7_10H INNER JOIN FILE7_10 ON FILE7_10H.DOC_NO = FILE7_10.DOC_NO where ( FILE7_10H.TAX + FILE7_10H.DISCOUNT ) <> 0 GROUP BY FILE7_10H.DOC_NO, FILE7_10H.discount,TAX "
        nCount = GetDesca("SELECT COUNT(DOC_NO) FROM FILE7_10H WHERE ( FILE7_10H.TAX + FILE7_10H.DISCOUNT ) <> 0 ")
    Case 2
        cStr1 = "SELECT FR7_20H.DOC_NO, FR7_20H.TAX  ,SUM(FR7_20.TOTAL) AS TOTAL , FR7_20H.DISCOUNT FROM         FR7_20H INNER JOIN FR7_20 ON FR7_20H.DOC_NO = FR7_20.DOC_NO where ( FR7_20H.TAX + FR7_20H.DISCOUNT ) <> 0 GROUP BY FR7_20H.DOC_NO, FR7_20H.discount,TAX "
        nCount = GetDesca("SELECT COUNT(DOC_NO) FROM FR7_20H WHERE ( FR7_20H.TAX + FR7_20H.DISCOUNT ) <> 0 ")
    Case 3
        cStr1 = "SELECT FR7_10H.DOC_NO, FR7_10H.TAX  ,SUM(FR7_10.TOTAL) AS TOTAL , FR7_10H.DISCOUNT FROM         FR7_10H INNER JOIN FR7_10 ON FR7_10H.DOC_NO = FR7_10.DOC_NO where ( FR7_10H.TAX + FR7_10H.DISCOUNT ) <> 0 GROUP BY FR7_10H.DOC_NO, FR7_10H.discount,TAX "
        nCount = GetDesca("SELECT COUNT(DOC_NO) FROM FR7_10H WHERE ( FR7_10H.TAX + FR7_10H.DISCOUNT ) <> 0 ")
End Select
SalTable.Open cStr1, con, adOpenKeyset, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Min = 0
prog1.Max = nCount
prog1.Visible = True

'UPDATE       FILE7_20H SET                R_DISC = QDOC7_20H.R_DISC FROM            FILE7_20H INNER JOIN                          QDOC7_20H ON FILE7_20H.DOC_NO = QDOC7_20H.DOC_NO
'UPDATE       FILE7_10H SET                R_DISC = QDOC7_10H.R_DISC FROM            FILE7_10H INNER JOIN                          QDOC7_10H ON FILE7_10H.DOC_NO = QDOC7_10H.DOC_NO
'UPDATE       FR7_20H SET                R_DISC = QDOCFR7_20H.R_DISC FROM            FR7_20H INNER JOIN                          QDOCFR7_20H ON FR7_20H.DOC_NO = QDOCFR7_20H.DOC_NO
'UPDATE       FR7_10H SET                R_DISC = QDOCFR7_10H.R_DISC FROM            FR7_10H INNER JOIN                          QDOCFR7_10H ON FR7_10H.DOC_NO = QDOCFR7_10H.DOC_NO


'UPDATE FILE7_20 SET  R_TAX = QDOC7_20H.RATE_INV FROM  FILE7_20 INNER JOIN QDOC7_20H ON FILE7_20.DOC_NO = QDOC7_20H.DOC_NO
'UPDATE FILE7_10 SET  R_TAX = QDOC7_10H.RATE_INV FROM  FILE7_10 INNER JOIN QDOC7_10H ON FILE7_10.DOC_NO = QDOC7_10H.DOC_NO
'UPDATE FR7_10 SET  R_TAX = QDOCFR7_10H.RATE_INV FROM  FR7_10 INNER JOIN QDOCFR7_10H ON FR7_10.DOC_NO = QDOCFR7_10H.DOC_NO
'UPDATE FR7_20 SET  R_TAX = QDOCFR7_20H.RATE_INV FROM  FR7_20 INNER JOIN QDOCFR7_20H ON FR7_20.DOC_NO = QDOCFR7_20H.DOC_NO


con.Execute " UPDATE " & cFileHeader & " SET R_TAX = 0 WHERE (TAX) = 0 "
con.CommandTimeout = 1000
con.Execute " UPDATE " & cFile & " SET R_TAX = 1 "
With SalTable
    .MoveFirst
    Do While Not .EOF
        i = i + 1
        prog1.Value = i
        If (!TOTAL) <> 0 Then nRate = (!TOTAL - !discount + !tax) / (!TOTAL)
        If (!TOTAL - !discount) <> 0 And (!tax) <> 0 Then nRatet = (!tax) / (!TOTAL - !discount) * 100
        con.Execute " UPDATE " & cFile & " SET R_TAX = " & nRate & " WHERE DOC_NO = " & MyParn(!doc_no)
        con.Execute " UPDATE " & cFileHeader & " SET R_TAX = " & Val(nRatet & "") & " WHERE DOC_NO = " & MyParn(!doc_no)
        Me.Caption = !doc_no & ""
        .MoveNext
    Loop
End With

If myPublic = 0 Then
    cStr1 = " UPDATE    FILE6_20 SET  cost = (SELECT TOP 1 ( FILE7_20.cost_inv ) FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FILE6_20.ITEM = FILE7_20.ITEM AND FILE7_20H.DATE <= FILE6_20H.DATE ORDER BY DATE DESC ) " & _
            " FROM      FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO"
    con.CommandTimeout = 1000
    con.Execute cStr1, nRec
    Inform "  „ Ÿ»ÿ „»Ì⁄«  " & nRec

    If cBranch < "60" Then
        cStr1 = " UPDATE    FILE1_60 SET  cost = (SELECT TOP 1 ( FILE7_20.cost_inv ) FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FILE1_60.ITEM = FILE7_20.ITEM AND FILE7_20H.DATE <= FILE1_60H.DATE ORDER BY DATE DESC ) " & _
                " FROM      FILE1_60 INNER JOIN FILE1_60H ON FILE1_60.DOC_NO = FILE1_60H.DOC_NO"
        con.CommandTimeout = 1000
        con.Execute cStr1, nRec
        Inform "  „ Ÿ»ÿ  ÕÊÌ·« " & nRec
    End If
    
    cStr1 = " UPDATE    FILE0_10 SET  cost = Coalesce((SELECT TOP 1 ( FILE7_20.cost_inv ) FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FILE0_10.ITEM = FILE7_20.ITEM AND FILE7_20H.DATE <= FILE0_10H.DATE ORDER BY DATE DESC ),0) " & _
            " FROM      FILE0_10 INNER JOIN FILE0_10H ON FILE0_10.DOC_NO = FILE0_10H.DOC_NO"
    con.CommandTimeout = 1000
    con.Execute cStr1, nRec
    Inform "  „ Ÿ»ÿ  ”ÊÌ… «·Ã—œ " & nRec
    
    
    cStr1 = " UPDATE    FILE1_10 SET  COSTITEM = Coalesce((SELECT TOP 1 ( FILE7_20.cost_inv ) FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FILE1_10.ITEM = FILE7_20.ITEM ORDER BY DATE DESC ),0) " & _
            " FROM      FILE1_10 "
    con.CommandTimeout = 1000
    con.Execute cStr1
    Inform "  „ Ÿ»ÿ  ﬂ·›… «·„ÊœÌ·«  " & nRec
End If

If myPublic = 2 Then
    cStr1 = " UPDATE    FR6_20 SET  cost = (SELECT TOP 1 ( FILE7_20.cost_inv ) FROM FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FR6_20.ITEM = FILE7_20.ITEM AND FILE7_20H.DATE <= FR6_20H.DATE ORDER BY DATE DESC ) " & _
            " FROM      FR6_20 INNER JOIN FR6_20H ON FR6_20.DOC_NO = FR6_20H.DOC_NO"
    con.CommandTimeout = 600
    con.Execute cStr1, nRec
    Inform "  „ Ÿ»ÿ „»Ì⁄«  " & nRec
End If
MsgBox ok

End Sub

Private Sub CMD_FIXSEC_Click()
If xSection.BoundText = "" Then Exit Sub
If MsgBox(" ⁄œÌ· „ÊœÌ·«  «·›« Ê—… ··ﬁ”„", vbYesNo) = vbYes Then
    con.Execute " UPDATE FILE1_10 SET [SECTION] = " & Val(xSection.BoundText) & " WHERE ITEM IN (SELECT ITEM FROM FILE7_20 WHERE DOC_NO = " & MyParn(xdoc_no.Text) & " )"
    Inform " „  «· ⁄œÌ· »‰Ã«Õ"
End If
End Sub

Private Sub CMD_NOTDEM_Click()
    ModelNotDem.pDoc_No = xdoc_no.Text
    ModelNotDem.Show
End Sub
Private Sub CMD_PrintBar_Click()
    BarCodeFRM.Show 1
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
Private Sub CMD_RECEVED_Click()
    On Error GoTo myerror
    strConShop = LoadConStringshop
    If openConShop(conShop) <> "ok" Then
        MsgBox "·« ÌÊÃœ « ’«· »«·”Ì—›—"
        Exit Sub
    End If
    If cBranch < "60" Then
        conShop.Execute " UPDATE " & cFileHeader & " SET Received = 1 , USERACCEPT = " & addstring(cusername) & " , TIMEACCEPT = GETDATE() WHERE DOC_NO = " & MyParn(xdoc_no.Text), nRec
    Else
        conShop.Execute " UPDATE " & cFileHeader_FR & " SET Received = 1 , USERACCEPT = " & addstring(cusername) & " , TIMEACCEPT = GETDATE() WHERE DOC_NO = " & MyParn(xdoc_no.Text), nRec
    End If
    con.Execute " UPDATE " & cFileHeader & " SET Received = 1 , USERACCEPT = " & addstring(cusername) & " , TIMEACCEPT = GETDATE() WHERE DOC_NO = " & MyParn(xdoc_no.Text)
    If myPublic = 0 Then
        con.Execute " delete from model_dem where item in (select item from file7_20 where doc_no = " & MyParn(xdoc_no.Text) & " ) "
    End If
    If nRec = 1 Then
        Inform " „  „—«Ã⁄… «·„” ‰œ Ê »Ì«‰ «·«’‰«› ’ÕÌÕ"
        myUndo
    End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmd_ret_Click()
    SubRet.Show 1
End Sub
Private Sub CMD_SEND_Click()
    On Error GoTo myerror
    strConShop = LoadConStringshop
    If openConShop(conShop) <> "ok" Then
        MsgBox "·« ÌÊÃœ « ’«· »«·”Ì—›—"
        Exit Sub
    End If
    conShop.Execute " update " & cFileHeader & " SET Notes_Doc = " & addstring(XNotes_Doc) & " where  DOC_NO = " & MyParn(xdoc_no.Text), nRec
    con.Execute " update " & cFileHeader & " SET Notes_Doc = " & addstring(XNotes_Doc) & " where  DOC_NO = " & MyParn(xdoc_no.Text), nRec
    If nRec = 1 Then
        Inform "  „ «—”«· «·»Ì«‰ ··”Ì—›—"
    End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdAddItems_Click()
    Dim lIsNoItem  As Boolean
    myAddModels
'    cmd_barcode_fix_Click
    If cBranch = "00" And myPublic = 0 Then
        lIsNoItem = IIf(GetDesca("SELECT ITEM FROM FILE1_10 WHERE ISNOITEM = 0 AND ITEM IN (SELECT ITEM FROM " & cFile & " WHERE DOC_NO = " & MyParn(xdoc_no.Text) & " )") = "", True, False)
        If lIsNoItem Then
            con.Execute " UPDATE " & cFileHeader & " SET ISNOITEM = 1 where doc_no = " & MyParn(xdoc_no.Text)
        Else
            con.Execute " UPDATE " & cFileHeader & " SET ISNOITEM = 0 where doc_no = " & MyParn(xdoc_no.Text)
        End If
    End If
End Sub
Private Sub cmdAddItems_GotFocus()
    cmdAddItems.BackColor = vbYellow
End Sub
Private Sub cmdAddItems_LostFocus()
    cmdAddItems.BackColor = &H8000000F
End Sub
Private Sub cmdBarCode_Click()
Select Case myPublic
Case 0
    addBarCode
    BarCodeFRM.Show 1
Case 1
    DelBarCode
End Select
End Sub
Private Sub cmdcopy_Click()
    cDocCopy = xdoc_no.Text
    cCodeCopy = xCode.Text
    Inform " Copy " & cDocCopy
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    
    con.BeginTrans
    If myPublic = 0 Then
        con.Execute " DELETE FROM NEW7_20H WHERE DOC_NO = " & addstring(xdoc_no.Text)
        For nRow = 2 To nCountBranch
            sStore = RetZero(nRow, 2)
            con.Execute "  INSERT INTO NEW7_20H (DOC_NO , BRANCH) VALUES        ( " & addstring(xdoc_no.Text) & " , " & addstring(sStore) & " )  "
        Next nRow
    ElseIf myPublic = 1 Then
        con.Execute " DELETE FROM NEW7_10H WHERE DOC_NO = " & addstring(xdoc_no.Text)
        For nRow = 2 To nCountBranch
            sStore = RetZero(nRow, 2)
            con.Execute "  INSERT INTO NEW7_10H (DOC_NO , BRANCH) VALUES        ( " & addstring(xdoc_no.Text) & " , " & addstring(sStore) & " )  "
        Next nRow
    End If
    
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xdoc_no.Text)
    con.Execute "update " & cFileHeader & " set isdate = getdate() , ISNOITEM = 0, isnew1 = 1 , discount =0 , tax = 0 where Doc_No = " & MyParn(xdoc_no.Text)
    con.CommitTrans
''''''    If XTRANS_DOC.Caption <> "" Then
''''''        If myPublic = 1 Then
''''''            con.Execute " update file1_61h set DOC_TRANS = null , TYPEDOC = null  , ISNEW = 1   where doc_no = " & MyParn(XTRANS_DOC.Caption)
''''''        End If
''''''        If myPublic = 3 Then
''''''            con.Execute " update fr1_61h set DOC_TRANS = null , TYPEDOC = null  ,ISNEW = 1     where doc_no = " & MyParn(XTRANS_DOC.Caption)
''''''        End If
''''''    End If
    myUndo
    CmdNewInv_Click
End If
Exit Sub
myerror:
con.RollbackTrans
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
    mydefine
    xinv_no.SetFocus
End Sub
Private Sub cmdpast_Click()
    If cDocCopy <> "" And xCode.Text = cCodeCopy Then
        If xdoc_no.Enabled Then
            If MsgBox("  ﬂ—«— ›« Ê—… —ﬁ„  " & cDocCopy, vbYesNo) = vbYes Then
                If Not mysave Then Exit Sub
                con.Execute " INSERT INTO FILE7_20 (doc_no , ITEM, QUANT, PRICE, DISCOUNT, ROW, MODEL, REMARK, r_tax) SELECT " & addstring(xdoc_no.Text) & "  , ITEM, QUANT, PRICE, DISCOUNT, ROW, MODEL, REMARK, r_tax FROM  FILE7_20 AS FILE7_20_1 WHERE DOC_NO = " & MyParn(cDocCopy)
                myUndo
            End If
        End If
    End If
End Sub
Private Sub cmdSave_Click()
    mysave
    If Val(xTotalItem.Caption) <> 0 Then
        nRCost = Val(xTotal.Caption) / (Val(xTotalItem.Caption))
        con.Execute " UPDATE " & cFile & " SET R_TAX = " & nRCost & " WHERE DOC_NO = " & MyParn(xdoc_no.Text)
    End If
    
    If myPublic = 0 Then
        cStr1 = " UPDATE   FILE1_10 SET  costITEM  = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST WHERE FILE1_10.ITEM = Q_ITEMCOST.ITEM ORDER BY DATE DESC ) From FILE1_10 WHERE ITEM IN (SELECT ITEM FROM FILE7_20 WHERE DOC_NO = " & MyParn(xdoc_no.Text) & ")"
        con.Execute cStr1, nRec
        Inform "  „ Ÿ»ÿ  ﬂ·›… «·„ÊœÌ·«  " & nRec
        
        If cBranch < "60" Then
            con.Execute " delete from model_dem where STORE = " & MyParn(XSTORE.BoundText) & " AND item in (select item from file7_20 where doc_no = " & MyParn(xdoc_no.Text) & " ) "
        Else
            con.Execute " delete from model_dem where STORE = " & MyParn(XSTORE.BoundText) & " AND item in (select item from fR7_20 where doc_no = " & MyParn(xdoc_no.Text) & " ) "
        End If
    End If
End Sub
Private Sub CmdUndo_Click()
    myUndo
End Sub
Private Sub CMD_PRINT_Click()
    myUndo
    doprint
End Sub
Private Sub Command1_Click()
    Dim cBrc As String
    If myPublic = 0 Or myPublic = 1 Then
        cBrc = GetDesca("SELECT BRANCH FROM FILE0_40 WHERE CODE = " & MyParn(XSTORE.BoundText))
        If cBrc <> "" Then
            con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBrc) & " FROM  FILE1_10  WHERE ITEM IN (SELECT ITEM FROM " & cFile & " WHERE DOC_NO = " & MyParn(xdoc_no.Text) & " ) ", nRec
            Inform "  „ «÷«›… " & nRec
        End If
    Else
        cBrc = XSTORE.BoundText
        If cBrc <> "" Then
            con.Execute "  INSERT INTO NEW_ITEM_FR (ITEM ,  BRANCH) SELECT ITEM , " & addstring(cBrc) & " FROM  FILE1_10  WHERE ITEM IN (SELECT ITEM FROM " & cFile & " WHERE DOC_NO = " & MyParn(xdoc_no.Text) & " ) ", nRec
            Inform "  „ «÷«›… " & nRec
        End If
    End If
End Sub
Private Sub cmd_exRET_Click()
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet
    Dim xlwbook As Excel.Workbook
    Dim cModelFact As String
    Dim cBarCode As String
    Dim cModel As String
    Dim cItem As String
    Dim nQTY As Double
    Dim nPrice As Double
    Dim nDisc As Double
    Dim nNewItem As String
    Dim lAddTrue As Boolean
    prog1.Min = 0
    prog1.Value = 0
    prog1.Visible = True
    If MsgBox("≈÷«›… »Ì«‰«  „‰ „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            If xdoc_no.Enabled Then
                If Not mysave Then Exit Sub
            End If
            
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
            For nRow = 1 To 10000
                If Len(xlsheet.Cells(nRow, 1)) > 1 Then
                    nRec = nRec + 1
                End If
                If Len(xlsheet.Cells(nRow, 1)) = 0 Then Exit For
            Next nRow
            prog1.Max = nRec
            i = 0
            For nRow = 1 To nRec
                i = i + 1
                prog1.Value = i
                cItem = xlsheet.Cells(nRow, 1)
                nQTY = xlsheet.Cells(nRow, 2)
                aRet = aGetDesca("select model , cost from  FILE1_10 where item = " & Val(cItem))
                If cItem = "" Then Exit For
                If UBound(aRet) > 0 Then
                    If GetDesca("select item from " & cFile & " where item = " & Val(cItem) & " and doc_no = " & MyParn(xdoc_no.Text)) = "" Then
                        cString = "Insert into " & cFile & " (doc_no,item,MODEL,price,Quant)" & _
                                   "Values(" & _
                                   addstring(xdoc_no.Text) & "," & _
                                   addvalue(cItem) & "," & _
                                   addstring(aRet(1)) & "," & _
                                   Val(aRet(2)) & "," & _
                                   Val(nQTY) & _
                                   ")"
                        con.Execute cString
                    End If
                    lAddTrue = True
                End If
            Next nRow
            If lAddTrue Then
                Inform " „   ”ÃÌ· «·„— Ã⁄ "
            Else
                Inform "·„ Ì „ «· ”ÃÌ· "
            End If
            myUndo
        End If
    End If
End Sub
Private Sub FIX_DOC_FR_Click()
    If cBranch > "60" Then
        con.Execute " UPDATE FILE7_20H SET STORE = " & addstring(cBranch)
        con.Execute " UPDATE FILE7_10H SET STORE = " & addstring(cBranch)
        con.Execute " UPDATE FILE6_20H SET STORE = " & addstring(cBranch)
        
        con.Execute " UPDATE FILE7_20H SET DOC_NO = FILE7_20H.STORE + FILE7_20H.doc_no WHERE LEN(DOC_NO )=  6 "
        con.Execute " UPDATE FILE7_10H SET DOC_NO = FILE7_10H.STORE + FILE7_10H.doc_no WHERE LEN(DOC_NO )=  6 "
        
        con.Execute " UPDATE FILE7_20  SET DOC_NO = " & addstring(cBranch) & " + FILE7_20.doc_no WHERE LEN(DOC_NO )=  6 "
        con.Execute " UPDATE FILE7_10  SET DOC_NO = " & addstring(cBranch) & " + FILE7_10.doc_no WHERE LEN(DOC_NO )=  6 "
        
        con.CommandTimeout = 1000
        con.Execute " UPDATE FILE6_20H SET DOC_NO = SUBSTRING(FILE6_20H.DOC_NO,1,6) + FILE6_20H.STORE + SUBSTRING(DOC_NO,7,5) WHERE LEN(DOC_NO) = 10 "
        
        
        con.Execute " UPDATE FILE0_10H SET BRANCH = " & addstring(cBranch)
        con.Execute " UPDATE FILE0_10H SET ISNEW = 1 , DOC_NO = FILE0_10H.BRANCH + FILE0_10H.DOC_NO WHERE LEN(DOC_NO) = 6 "
        Inform " END "
    End If
End Sub

Private Sub FIX_FACTITEM_Click()
Dim ItemFactTable As New ADODB.Recordset, i As Double
prog1.Visible = True
prog1.Min = 0
prog1.Value = 0
ItemFactTable.Open "SELECT * FROM FILE1_10 WHERE CODE = '001' AND BARCODE IS NOT NULL ", con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Max = ItemFactTable.RecordCount
ItemFactTable.MoveFirst
With ItemFactTable
    Do While Not .EOF
        i = i + 1
        prog1.Value = i
        confact.Execute " UPDATE FILE1_10 SET BARCODE_SHOP = " & !Item & " WHERE  ITEM = " & MyParn(!BARCODE)
        .MoveNext
    Loop
End With
End Sub


Private Sub FIX_PRICE_Click()
If MsgBox("”Ê› Ì „  €Ì— ﬂ· «”⁄«— «·„Õ· ·„’‰⁄ ÃÊ‰ÌÊ— ”⁄— »Ì⁄ Ê „’‰⁄ ··„ÊœÌ·«  ", vbYesNo + vbDefaultButton2) = vbYes Then
    If MsgBox("«” ﬂ„«· «· €Ì— ", vbYesNo + vbDefaultButton2) = vbYes Then
'        EditAllItem
        EditAllItemPrice
    End If
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
    Frame4.Visible = bopt2
    nRound = 2
    CMD_RECEVED.Visible = (cBranch <> "00")
    reprice_inv.Visible = bopt2
    CMD_NOTDEM.Visible = (cBranch = "00")
    If cBranch = "00" Then
        If myPublic = 0 Or myPublic = 1 Then
            cmd_fixcost.Visible = True
        End If
        Frame14.Visible = True
        openConFACT confact
        openConFACT2 confact2
        cmd_ex.Visible = True
        CMD_ADDINV.Visible = True
    Else
        cmdCopy.Visible = False
        cmdpast.Visible = False
        If cBranch < "60" Then
            cmd_fixcost.Visible = False
        Else
            cmd_fixcost.Visible = True
        End If
    End If
    'SetKbLayout Lang_EN
    con.CursorLocation = adUseClient
    con.Open strCon
    CMD_SEND.Visible = (cBranch <> "00")
    'Frame8.Visible = lSupperVisor
    
    If cBranch <> "00" Then
        cmd_exRET.Visible = False
        CMD_ALL.Visible = False
        CmdDelInv.Visible = False
        cmdAddItems.Visible = False
        CmdSave.Visible = False
        cmd_barcode_fix.Visible = False
        Frame4.Visible = False
        If cBranch < "60" Then
            CMD_SEND.Visible = True
            Frame7.Visible = False
            Frame10.Visible = False
        End If
                
    End If
    
    Frame7.Visible = bOpt10
    CmdDelInv.Visible = bOpt10
    CmdSave.Visible = bOpt10
    cmdAddItems.Visible = bOpt10
    cmdCopy.Visible = bOpt10
    cmdpast.Visible = bOpt10
    reprice_inv.Visible = bOpt10
    FIX_PRICE.Visible = bOpt10
    cmd_exRET.Visible = bOpt10
    cmd_ex.Visible = bOpt10
    CMD_ADDINV.Visible = bOpt10
    bEdit = True
    Select Case myPublic
    Case 0
        cmd_exRET.Visible = False

        pGrCust = "04"
        cmd_ret.Visible = False
        cFile = "File7_20"
        cFileHeader = "File7_20H"
        
        cFile_FR = "FR7_20"
        cFileHeader_FR = "FR7_20H"
        
        cFileClient = "File4_10"
        Me.Caption = "›« Ê—… „‘ —Ì« "
        If cBranch = "00" Then
            Command1.Visible = True
        End If
    Case 1
        cmd_exRET.Visible = True
        pGrCust = "04"
        Frame14.Visible = True
        If cBranch = "00" Then
            Command1.Visible = True
        End If
        CMD_ALL.Visible = False
        cmdBarCode.Caption = "«·€«¡ »«—ﬂÊœ"
        cFile = "FILE7_10"
        cFileHeader = "FILE7_10H"
        
        cFile_FR = "FR7_10"
        cFileHeader_FR = "FR7_10H"
        
        cFileClient = "File4_10"
        Me.Caption = "›« Ê—… „—œÊœ „‘ —Ì« "
        Me.BackColor = &H8080FF
        Me.Frame1.BackColor = &H8080FF
        Me.Frame2.BackColor = &H8080FF
        Me.Frame3.BackColor = &H8080FF
        Me.Frame4.BackColor = &H8080FF
        Me.Frame5.BackColor = &H8080FF
        Me.Frame6.BackColor = &H8080FF
        Me.Frame7.BackColor = &H8080FF
        Me.Frame8.BackColor = &H8080FF
        Me.Frame9.BackColor = &H8080FF
    Case 2
        cmd_exRET.Visible = False
'        FIX_DOC_FR.Visible = (cBranch > "60")
        Me.BackColor = &HFFFFC0
        Me.Frame1.BackColor = &HFFFFC0
        Me.Frame2.BackColor = &HFFFFC0
        Me.Frame3.BackColor = &HFFFFC0
        Me.Frame4.BackColor = &HFFFFC0
        Me.Frame5.BackColor = &HFFFFC0
        Me.Frame6.BackColor = &HFFFFC0
        Me.Frame7.BackColor = &HFFFFC0
        Me.Frame8.BackColor = &HFFFFC0
        Me.Frame9.BackColor = &HFFFFC0
        
        pGrCust = "16"
        cmd_ret.Visible = False
        cFile = "FR7_20"
        cFileHeader = "FR7_20H"
        cFileClient = "File4_10"
        Me.Caption = "›« Ê—… „‘ —Ì«   ÊﬂÌ·« "
        If cBranch = "00" Then
            Command1.Visible = True
        End If
'        CMD_ALL.Visible = True
    Case 3
        cmd_exRET.Visible = True
        If cBranch = "00" Then sned_purch.Visible = True
        Me.BackColor = &H80FF&
        Me.Frame1.BackColor = &H80FF&
        Me.Frame2.BackColor = &H80FF&
        Me.Frame3.BackColor = &H80FF&
        Me.Frame4.BackColor = &H80FF&
        Me.Frame5.BackColor = &H80FF&
        Me.Frame6.BackColor = &H80FF&
        Me.Frame7.BackColor = &H80FF&
        Me.Frame8.BackColor = &H80FF&
        Me.Frame9.BackColor = &H80FF&
        pGrCust = "16"
'        CMD_ALL.Visible = False
        cmd_ret.Visible = False
        cFile = "FR7_10"
        cFileHeader = "FR7_10H"
        cFileClient = "File4_10"
        Me.Caption = "›« Ê—… „—œÊœ „‘ —Ì«   ÊﬂÌ·« "
        Command1.Visible = False
        If cBranch = "00" Then
            Command1.Visible = True
        End If
    End Select
        
''''''Set data4.Recordset = myRecordSet("SELECT * FROM fact ORDER BY code ", con)
''''''Set xFact.RowSource = data4
''''''xFact.ListField = "Desca"
''''''xFact.BoundColumn = "Code"
        
    If myPublic = 0 Or myPublic = 1 Then
        Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_40", con)
    Else
        Set data1.Recordset = myRecordSet("SELECT * FROM BRANCH_FR", con)
    End If
    Set XSTORE.RowSource = data1
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    
    
    Set data4.Recordset = myRecordSet("SELECT * FROM MOSM ORDER BY DATE DESC ", con)
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    Set DATA2.Recordset = myRecordSet("SELECT * FROM FILE1_10SC ", con)
    Set xSection.RowSource = DATA2
    xSection.ListField = "Desca"
    xSection.BoundColumn = "CODE"
    
    With grid1
        .Cols = 10
        .Rows = 1
    End With
    
    
    Set grid1.DataSource = data3
        
    xdoc_no.Text = sDoc_no
    myUndo
    If sDoc_no = "" Then mydefine
        
'    Fixgrd
'    openCardTable
'    If sDoc_no <> "" Then
'        xDoc_No.Text = sDoc_no
'        myUndo
'    End If
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
        con.Execute " update " & cFile & " set discount = " & Val(.TextMatrix(.Row, 9)) & " FROM   " & cFile & " JOIN FILE1_10 ON " & cFile & ".ITEM = FILE1_10.ITEM  where FILE1_10.model = " & MyParn(.TextMatrix(.Row, 13)) & " and doc_no = " & MyParn(xdoc_no.Text)
    End If
    If .Col = 14 Then
        con.Execute " update file1_10 set barcode = " & addstring(.TextMatrix(.Row, 14)) & " ,ISDATE = GETDATE() where item = " & MyParn(.TextMatrix(.Row, 0))
    End If
    myloadgrd
End With
End Sub
Private Sub grid1_DblClick()
    If xClosed.Value <> 0 Then Exit Sub
    If cBranch <> "00" Then Exit Sub
    If grid1.Col = 7 Then
        addBarCodeModel GetDesca("SELECT MODEL FROM FILE1_10 WHERE ITEM = " & grid1.TextMatrix(grid1.Row, 0))
    ElseIf grid1.Col = 8 Then
        myAddModels True
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
'Me.Caption = grid1.TextMatrix(grid1.Row, 13)
If xClosed.Value = 1 Then Exit Sub

If grid1.TextMatrix(grid1.Row, 13) <> "" Then LoadPhoto grid1.TextMatrix(grid1.Row, 13) Else ImgX1.Images.Clear
'If grid1.Col = 0 Then grid1.Editable = flexEDKbdMouse Else grid1.Editable = flexEDNone
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
    oMaxPhoto.cdesca = grid1.TextMatrix(grid1.Row, 4)
    oMaxPhoto.Show 1
End If
End Sub
Private Sub reprice_inv_Click()
    If MsgBox(" ⁄œÌ· «”⁄«— «·„” ‰œ »«”⁄«— «·„’‰⁄ «·Õ«·Ì…", vbYesNo + vbDefaultButton2) = vbYes Then
        Select Case myPublic
        Case 0
            con.Execute " UPDATE       FILE7_20  SET  PRICE = FILE1_10.COST FROM  FILE1_10 INNER JOIN FILE7_20 ON FILE1_10.ITEM = FILE7_20.ITEM where doc_no = " & MyParn(xdoc_no.Text), nRec
        Case 1
            con.Execute " UPDATE       FILE7_10  SET  PRICE = FILE1_10.COST FROM  FILE1_10 INNER JOIN FILE7_10 ON FILE1_10.ITEM = FILE7_10.ITEM where doc_no = " & MyParn(xdoc_no.Text), nRec
        Case 2
            con.Execute " UPDATE       FR7_20  SET    PRICE = FILE1_10.COST FROM  FILE1_10 INNER JOIN FR7_20 ON FILE1_10.ITEM = FR7_20.ITEM where doc_no = " & MyParn(xdoc_no.Text), nRec
        Case 3
            con.Execute " UPDATE       FR7_10  SET    PRICE = FILE1_10.COST FROM  FILE1_10 INNER JOIN FR7_10 ON FILE1_10.ITEM = FR7_10.ITEM where doc_no = " & MyParn(xdoc_no.Text), nRec
                
        End Select
        myUndo
        Inform "  „  ⁄œÌ·  " & nRec
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
xr_Discount.Text = ""
xSubCode.Caption = ""
xSupp.Value = 0
xsup_desca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 3)
If xCode.Text = "" Then Exit Sub

Dim cString As String, cWhere As String
cString = "select code,desca, SUPP , DISC ,SUBCODE  FROM FILE4_10"
cString = cString & " WHERE CODE = " & MyParn(xCode.Text)
aRet = aGetDesca(cString)
If UBound(aRet) > 0 Then
    xCode.Text = aRet(1)
    xCodeDesca.Caption = aRet(2)
    xSupp.Value = IIf(aRet(3), 1, 0)
    xsup_desca.Caption = IIf(aRet(3), "„ﬂ » Ã„·…", "„’‰⁄")
    xr_Discount.Text = Val(aRet(4) & "")
    xRateSup.Caption = Val(aRet(4) & "")
    xSubCode.Caption = aRet(5) & ""
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
    If Not IsDate(xDate.Text) Then
        MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
        Exit Function
    End If
    If XSTORE.BoundText = "" Then
        MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
        Exit Function
    End If
    
    If xMosm.BoundText = "" Then
        MsgBox "·„ Ì „  ÕœÌœ «·„Ê”„ "
        Exit Function
    End If
    
    If cBranch = "00" Then
        If GetBoolean("select closed from mosm where mosm = " & MyParn(xMosm.BoundText)) Then
            MsgBox " „Ê”„ „€·ﬁ  "
            Exit Function
        End If
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
xDiscAll.Text = ""
XTRANS_DOC.Caption = CardTable!TRANS_DOC & ""
xUserAccept.Text = CardTable!UserAccept & ""
xUserSend.Text = CardTable!USERSEND & ""
xTimeAccept.Text = CardTable!TimeAccept & ""
xTimeSend.Text = CardTable!TimeSend & ""
XNotes_Doc.Text = CardTable!Notes_Doc & ""
xReceived.Value = IIf(CardTable!Received, 1, 0)
CMD_RECEVED.Enabled = (xReceived.Value = 0)

xdoc_no.Text = CardTable!doc_no
xinv_no.Text = CardTable!INV_NO & ""
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")

XSTORE.BoundText = CardTable!STORE & ""
If myPublic = 2 Or myPublic = 3 Then
    XSTORE.Locked = True
End If
xMosm.BoundText = Trim(CardTable!MOSM & "")
xCode.Text = CardTable!code & ""


xCodeDesca.Caption = CardTable!CodeDesca & ""
xSupp.Value = IIf(CardTable!SUPP, 1, 0)
xsup_desca.Caption = IIf(CardTable!SUPP, "„ﬂ » Ã„·…", "„’‰⁄")
xSubCode.Caption = CardTable!SUBCODE & ""

xdiscount.Text = CardTable!discount & ""
xTax.Text = TurnValue(Val(CardTable!tax & ""), 0, "")
xRateSup.Caption = Val(CardTable!disc & "")
XONEST.Value = IIf(CardTable!ONEST, 1, 0)
If GetDesca("Select doc_no from addprint where doc_no = " & MyParn(xdoc_no.Text)) <> "" Then
    Me.StatusBar1.Panels(1).Text = "„”Ã· ·Â« »«—ﬂÊœ"
Else
    Me.StatusBar1.Panels(1).Text = ""
End If
'xDiscount.Text = TurnValue(Val(CardTable!DISCOUNT & ""), 0, "")
myloadgrd
grid1.Row = 1
grid1.Col = 0
grid1_EnterCell
End Sub
Public Sub myloadgrd()
Dim cString As String
With grid1
'                                   0        1           2                   3                   4           5           6           7                   8       9       10              11          12          13                 14              15
    If cBranch = "00" Then
        cString = "SELECT " & cFile & ".ITEM,file1_10.mosm,FACT.DESCA ,file1_10.modelfact0,file1_10.desca ,file1_10.scal,file1_10.color,Quant," & cFile & ".Price, DISCOUNT, total   , FILE1_10.PRICE, COST_INV ,FILE1_10.MODELNO,FILE1_10.barcode ,FILE1_10.barcode2 " & _
              " FROM (" & cFile & " INNER JOIN FILE1_10 ON " & cFile & ".ITEM = FILE1_10.ITEM) LEFT JOIN FACT ON FILE1_10.FACT = FACT.CODE"
    Else
        cString = "SELECT " & cFile & ".ITEM,file1_10.mosm,FACT.DESCA ,file1_10.modelfact0,file1_10.desca ,file1_10.scal,file1_10.color,Quant," & cFile & ".Price, DISCOUNT, total   , FILE1_10.PRICE, 0 ,FILE1_10.MODELNO,FILE1_10.barcode ,FILE1_10.barcode2 " & _
              " FROM (" & cFile & " INNER JOIN FILE1_10 ON " & cFile & ".ITEM = FILE1_10.ITEM) LEFT JOIN FACT ON FILE1_10.FACT = FACT.CODE"
    End If
    cString = cString & turn(cString) & "DOC_NO = " & MyParn(xdoc_no.Text)
    
    Set data3.Recordset = myRecordSet(cString, con)
    grid1.AddItem ""
    xCode.Enabled = .Rows <= 2
    xdoc_no.Enabled = .Rows <= 2
End With
Handlecontrols LoadMode
CalcTotals
Fixgrd
End Sub
Private Sub mydefine()
    xRateSup.Caption = ""
    xClosed.Value = 0
    XMODELNOT.Text = ""
    If cmd_closed.Caption <> "≈€·«ﬁ „” ‰œ" Then
        cmd_closed.BackColor = &H8000000F
        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
        xClosed.Visible = False
    End If
    
    XTRANS_DOC.Caption = ""
    xDiscAll.Text = ""
    xUserAccept.Text = ""
    xUserSend.Text = ""
    xTimeAccept.Text = ""
    xTimeSend.Text = ""
    xReceived.Value = 0
    CMD_RECEVED.Enabled = False
    XNotes_Doc.Text = ""
    cmdAddItems.Enabled = True
    xSubCode.Caption = ""
    xClosed.Visible = False
    xClosed.Value = ssCBUnchecked
    XSTORE.BoundText = ""
    xCode.Enabled = True
    xinv_no.Text = ""
    xDate.Text = Format(Date, "dd-mm-yyyy")
    XSTORE.Locked = False
    If myPublic = 0 Or myPublic = 1 Then
        xdoc_no.Text = RetZero(Newflag(cFileHeader, "doc_no"))
    Else
        xdoc_no.Text = ""
'        xdoc_no.Text = RetZero(Newflag(cFileHeader, "doc_no"))
    End If
    XONEST.Value = 0
    xCodeDesca.Caption = ""
    xBalance.Caption = ""
    xCode.Text = ""
    xdiscount.Text = ""
'   lblBarCode.Visible = True
    
    xtotalOrg.Caption = ""
    xDiscountItem.Caption = ""
    xTotalItem.Caption = ""
    xr_Discount.Text = ""
    xTotalNoTax.Caption = ""
    xr_tax.Text = ""
    xTax.Text = ""
    xTotal.Caption = ""
    xtotalQuant.Caption = ""
    xsup_desca.Caption = ""

'    xTotalDis.Caption = ""
    xUserName.Caption = ""
'    xRate.Text = ""
    xMosm.BoundText = cPMosm
    grid1.Rows = 1
    grid1.AddItem ""
    ImgX1.Images.Clear
    Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
    
    CmdSave.Enabled = (xClosed.Value = 0) And (xReceived.Value = 0)
    CmdDelInv.Enabled = nMode = LoadMode And (xClosed.Value = 0) And (xReceived.Value = 0)
    cmdAddItems.Enabled = (xClosed.Value = 0) And (xClosed.Value = 0) And (xReceived.Value = 0)
    CMD_ALL.Enabled = (xClosed.Value = 0)
    cmdNewInv.Enabled = nMode = LoadMode
    cmd_closed.Enabled = nMode = LoadMode
    
    cmdFirst.Enabled = (nMode = LoadMode)
    cmdLast.Enabled = (nMode = LoadMode)
    cmdNext.Enabled = (nMode = LoadMode)
    cmdPrevious.Enabled = (nMode = LoadMode)
    xClosed.Enabled = (nMode = LoadMode)
    xdoc_no.Enabled = (nMode = DefineMode)
    cmd_closed.Enabled = (nMode = LoadMode)
    xdoc_no.Tag = nMode
'    If xClosed.Value = 1 Then
'        cmd_closed.BackColor = &H8000000F
'        cmd_closed.Caption = "› Õ „” ‰œ"
'        xClosed.Value = ssCBChecked
'        xClosed.Visible = True
'    Else
'        cmd_closed.BackColor = &H8080FF
'        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
'        xClosed.Value = ssCBUnchecked
'        xClosed.Visible = False
'    End If
    LoadPict xClosed.Value
End Sub
Private Sub xDiscount_Validate(Cancel As Boolean)
CalcTotals
End Sub

Private Sub xDoc_No_GotFocus()
myGotFocus xdoc_no
End Sub
Private Sub xDoc_No_LostFocus()
    myLostFocus xdoc_no
    If xdoc_no.Text = "" Then Exit Sub
    xdoc_no.Text = RetZero(xdoc_no.Text)
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
    If Round(Val(xr_Discount.Text), nRound) <> Round(Val(xdiscount.Text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xr_Discount.Text = Myvalue(Round((Val(xdiscount.Text) / Val(xTotalItem.Caption)) * 100, nRound))
    End If
Else
    xr_Discount.Text = ""
End If

xTotalNoTax.Caption = Format(Val(xTotalItem.Caption) - Val(xdiscount.Text), "#0.00")

If Val(xTotalNoTax.Caption) <> 0 Then
    If Round(Val(xr_tax.Text), nRound) <> Round(Val(xTax.Text) / Val(xTotalNoTax.Caption) * 100, nRound) Then
        xr_tax.Text = Myvalue(Round((Val(xTax.Text) / Val(xTotalNoTax.Caption)) * 100, nRound))
    End If
Else
    xr_tax.Text = ""
End If
xTotal.Caption = Format(Val(xTotalNoTax.Caption) + Val(xTax.Text))
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(2, 4)
Dim GrdArray(8, 1)

Set Generalarray(0) = Me
If myPublic = 0 Or myPublic = 1 Then
    Generalarray(1) = "SELECT     top 200 DOC_NO,[DATE] AS DATE1,DATE, FILE4_10.Desca , inv_no , mosm , file0_40.desca , Notes_Doc , TRANS_DOC " & _
                      " FROM  (" & cFileHeader & " left JOIN FILE4_10 ON " & cFileHeader & ".CODE " & " = FILE4_10.CODE inner join file0_40 on file0_40.code = " & cFileHeader & ".store )"
Else
    Generalarray(1) = "SELECT     top 200 DOC_NO,[DATE] AS DATE1,DATE, FILE4_10.Desca , inv_no , mosm , BRANCH_FR.desca , Notes_Doc , TRANS_DOC " & _
                      " FROM  (" & cFileHeader & " left JOIN FILE4_10 ON " & cFileHeader & ".CODE " & " = FILE4_10.CODE inner join BRANCH_FR on BRANCH_FR.code = " & cFileHeader & ".store )"
End If
If cBranch = "00" And Not bOpt5 Then Generalarray(1) = Generalarray(1) & " where store in (select store from USERSHOP where code = " & nusercode & " ) "
Generalarray(2) = "Order by DATE DESC , doc_no desc "
Generalarray(3) = 6000
Generalarray(5) = True


listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
listarray(0, 1) = "(inv_no Like '%cFilter%' or Doc_No Like '%cFilter%' or  FILE4_10.DESCA LIKE '%cFilter%' OR " & _
                  "##date##)"

listarray(1, 0) = "«·›—⁄"
If myPublic = 0 Or myPublic = 1 Then
    listarray(1, 1) = "file0_40.desca Like '%cFilter%' "
Else
    listarray(1, 1) = "BRANCH_FR.desca Like '%cFilter%' "
End If

listarray(2, 0) = "„·«ÕŸ«  «·›—⁄„” ‰œ"
listarray(2, 1) = "(Notes_Doc Like '%cFilter%' OR TRANS_Doc Like '%cFilter%' ) "

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

GrdArray(6, 0) = "«·„Œ“‰"
GrdArray(6, 1) = 1500

GrdArray(7, 0) = "„·«ÕŸ«  «·›—⁄"
GrdArray(7, 1) = 1500

GrdArray(8, 0) = "„” ‰œ «· ÕÊÌ·"
GrdArray(8, 1) = 1200

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search1
Search1.Caption = "«” ⁄·«„"
Search1.Show 1
End Sub
Private Sub xInv_No_GotFocus()
myGotFocus xinv_no
End Sub
Private Sub xInv_No_LostFocus()
myLostFocus xinv_no
End Sub

Private Sub xIsClosed_Click()

End Sub

Private Sub xMosm_GotFocus()
myGotFocus xinv_no
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xinv_no
End Sub
Private Sub xRate_LostFocus()
If Val(XRATE.Text) <> 0 Then
    xTax.Text = Format(Val(xTotalDis.Caption) * (Val(XRATE.Text) / 100), "Fixed")
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

.ColHidden(13) = True

.ColHidden(8) = (cBranch <> "00" And cBranch < "60") Or Not bOpt10
.ColHidden(9) = (cBranch <> "00" And cBranch < "60") Or Not bOpt10
.ColHidden(10) = (cBranch <> "00" And cBranch < "60") Or Not bOpt10
.ColHidden(12) = (cBranch <> "00" And cBranch < "60") Or Not bOpt10

'.ColComboList(0) = "..."
'.ColComboList(2) = StrList("SELECT CODE , DESCA FROM FACT ")
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub xr_Discount_Validate(Cancel As Boolean)

If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xr_Discount.Text), nRound) <> Round(Val(xdiscount.Text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xdiscount.Text = Round((Val(xr_Discount.Text) * Val(xTotalItem.Caption)) / 100, nRound)
    End If
Else
    xdiscount.Text = ""
End If
CalcTotals
End Sub

Private Sub xr_tax_Validate(Cancel As Boolean)
If Val(xTotalNoTax.Caption) <> 0 Then
    If Round(Val(xr_tax.Text), nRound) <> Round(Val(xTax.Text) / Val(xTotalNoTax.Caption) * 100, nRound) Then
        xTax.Text = Round((Val(xr_tax.Text) * Val(xTotalNoTax.Caption)) / 100, nRound)
    End If
Else
    xTax.Text = ""
End If
CalcTotals
End Sub

Private Sub xRateSup_Click()
xr_Discount.Text = Val(xRateSup.Caption)
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
    temptable.AddNew
    temptable!str10 = Me.Caption & "  " & XSTORE.Text
    temptable!str6 = xdoc_no.Text
    temptable!STR7 = xCodeDesca.Caption
    temptable!str5 = TurnValue(xinv_no.Text)
    temptable!Date1 = xDate.Text
    temptable!str1 = TurnValue(grid1.TextMatrix(i, 0))
    temptable!str13 = TurnValue(grid1.TextMatrix(i, 3))
    temptable!str2 = TurnValue(grid1.TextMatrix(i, 2), "", Null)
    temptable!str3 = TurnValue(grid1.TextMatrix(i, 4))
    temptable!str4 = TurnValue(grid1.TextMatrix(i, 5))
    temptable!str8 = TurnValue(grid1.TextMatrix(i, 6))
    temptable!val3 = Val(grid1.TextMatrix(i, 7))
    temptable!val1 = Val(grid1.TextMatrix(i, 8))
    temptable!VAL4 = Val(grid1.TextMatrix(i, 11))
    temptable!val15 = Val(grid1.TextMatrix(i, 10))
    temptable!Val10 = Val(grid1.TextMatrix(i, 9))
    temptable!val2 = Val(grid1.TextMatrix(i, 12))
    temptable!str15 = TurnValue(grid1.TextMatrix(i, 13))
    
    If myPublic = 0 Then
        temptable!str16 = " ÕÊÌ· ≈·Ï : " & XSTORE.Text
    ElseIf myPublic = 1 Then
        temptable!str16 = " ÕÊÌ· „‰ : " & XSTORE.Text
    End If
    temptable!val5 = Val(xdiscount.Text)
    temptable!Val6 = Val(xTax.Text)
    temptable!val17 = Val(xTotal.Caption)
    temptable!str19 = Firsttitle
    temptable!VAL14 = i
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If (cBranch = "00" And bOpt10) Or cBranch > "60" Then
    If MsgBox("≈ŸÂ«— «·«”⁄«—", vbYesNo) = vbYes Then
        If myPublic = 0 Or myPublic = 2 Then
            main.Report1.ReportFileName = App.Path & "\Reports\Print_Purch.rpt"
        Else
            main.Report1.ReportFileName = App.Path & "\Reports\Print_PurchR.rpt"
        End If
    Else
        If myPublic = 0 Or myPublic = 2 Then
            main.Report1.ReportFileName = App.Path & "\Reports\Print_Purch0.rpt"
        Else
            main.Report1.ReportFileName = App.Path & "\Reports\Print_PurchR0.rpt"
        End If
    End If
Else
    If myPublic = 0 Or myPublic = 2 Then
        main.Report1.ReportFileName = App.Path & "\Reports\Print_Purch0.rpt"
    Else
        main.Report1.ReportFileName = App.Path & "\Reports\Print_PurchR0.rpt"
    End If
End If
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
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
If bopt2 Then
    cString = " update " & cFileHeader & " set closed = " & IIf(xClosed.Value = 1, "0", "1") & " , Received = " & IIf(xClosed.Value = 1, "0", "1")
Else
    cString = " update " & cFileHeader & " set closed = " & IIf(xClosed.Value = 1, "0", "1")
End If
cString = cString & turn(cString) & "doc_no = " & MyParn(xdoc_no.Text)
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
oClosefrm.sCaption = IIf(myPublic = 0 Or myPublic = 2, "„‘ —Ì« ", "„—œÊœ „‘ —Ì« ")
oClosefrm.nMode = 0
oClosefrm.Show 1
myUndo
End Sub
Private Sub cmd_open_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sCaption = IIf(myPublic = 0 Or myPublic = 2, "„‘ —Ì« ", "„—œÊœ „‘ —Ì« ")
oClosefrm.nMode = 1
oClosefrm.Show 1
myUndo
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
Search1.Caption = "«” ⁄·«„"
Search1.Show 1
End Sub
Private Function myreplaceGrd() As Boolean
Dim loctable As ADODB.Recordset, cString As String, nItem As Double, cQDem As String
cString = " DELETE FROM " & cFile & " FROM   FILE1_10 INNER JOIN " & cFile & " ON FILE1_10.ITEM = " & cFile & ".ITEM"
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xdoc_no.Text)
cString = cString & turn(cString) & "FILE1_10.MODEL = " & MyParn(addModelFrm.XMODEL.Caption)
con.Execute cString, nDel
With addModelFrm.grid1
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                nItem = addModelFrm.grid2.TextMatrix(nRow, nCol)
                cQDem = GetDesca("SELECT QUANT FROM Model_Dem WHERE ITEM = " & nItem & " AND STORE = " & MyParn(XSTORE.BoundText))
                If Val(cQDem) < Val(.TextMatrix(nRow, nCol)) Then
                    cString = "Insert into SUB7_20  (doc_no,STORE,item,QUANT,Q_DEM)" & _
                               "Values(" & _
                               addstring(xdoc_no.Text) & "," & _
                               addstring(XSTORE.BoundText) & "," & _
                               addvalue(nItem) & "," & _
                               Val(.TextMatrix(nRow, nCol)) & "," & _
                               Val(cQDem) & _
                               ")"
                      
                    con.Execute cString
                End If
                Set loctable = ItemFind(addModelFrm.grid2.TextMatrix(nRow, nCol), con)
                If Not (loctable.EOF And loctable.BOF) Then
                    If myPublic = 0 Or myPublic = 1 Then
                    cString = "Insert into " & cFile & " (doc_no,item,MODEL,REMARK,discount,price,Quant,trust_model)" & _
                               "Values(" & _
                               addstring(xdoc_no.Text) & "," & _
                               addvalue(loctable!Item) & "," & _
                               addstring(loctable!MODEL) & "," & _
                               addstring(addModelFrm.xRemark.Text) & "," & _
                               Val(addModelFrm.xdiscount.Text) & "," & _
                               Val(loctable!cost & "") & "," & _
                               Val(.TextMatrix(nRow, nCol)) & "," & _
                               addModelFrm.xTrust_model.Value & _
                               ")"
                    Else
                    cString = "Insert into " & cFile & " (doc_no,item,MODEL,REMARK,discount,price,Quant,trust_model)" & _
                               "Values(" & _
                               addstring(xdoc_no.Text) & "," & _
                               addvalue(loctable!Item) & "," & _
                               addstring(loctable!MODEL) & "," & _
                               addstring(addModelFrm.xRemark.Text) & "," & _
                               Val(addModelFrm.xdiscount.Text) & "," & _
                               Val(loctable!cost2 & "") & "," & _
                               Val(.TextMatrix(nRow, nCol)) & "," & _
                               addModelFrm.xTrust_model.Value & _
                               ")"
                    End If
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
    If xdoc_no.Tag = DefineMode Then
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
Private Function addBarCode() As String
con.BeginTrans
On Error GoTo myerror
con.Execute "DELETE FROM ADDPRINT WHERE DOC_NO = " & MyParn(xdoc_no.Text)
cString = "INSERT INTO ADDPRINT(DOC_NO,ITEM,QUANT,ISPRINT)"

cString = cString & _
          " SELECT DOC_NO,ITEM,QUANT,1 FROM FILE7_20"
cString = cString & turn(cString) & " DOC_NO = " & MyParn(xdoc_no.Text)
con.Execute cString


cString = " update FILE7_20H set isbar = 1"
cString = cString & turn(cString) & " doc_no = " & MyParn(xdoc_no.Text)
con.Execute cString
 
con.CommitTrans
Inform " „ «· ÕÊÌ· »‰Ã«Õ"
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub DelBarCode()
Dim cString As String, nQuant As Long, nQuantDel As Long
Dim loctable As ADODB.Recordset
con.BeginTrans
For i = 1 To grid1.Rows - 2
    Set loctable = New ADODB.Recordset
    cString = "SELECT * FROM ADDPRINT"
    cString = cString & turn(cString) & "ITEM = " & MyParn(grid1.TextMatrix(i, 0))
    cString = cString & " ORDER BY DOC_NO"
    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    nQuant = Val(grid1.TextMatrix(i, 7))
    Do Until loctable.EOF Or nQuant < 0
        nQuantDel = IIf(loctable!Quant >= nQuant, nQuant, loctable!Quant)
        cString = "UPDATE ADDPRINT SET ADDPRINT.QUANT = ADDPRINT.QUANT - " & nQuantDel
        cString = cString & turn(cString) & " ADDPRINT.ID = " & loctable!ID
        con.Execute cString, nRecord
        nQuant = nQuant - nQuantDel
        loctable.MoveNext
    Loop
Next

cString = " update FILE7_10H set isbar = 1"
cString = cString & turn(cString) & " doc_no = " & MyParn(xdoc_no.Text)
con.Execute cString

con.CommitTrans
Inform " „ «·€«¡ »‰Ã«Õ"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myAddModels(Optional bAddModel As Boolean = False)
If Not MYVALID Then Exit Sub
Set addModelFrm.myForm = Me
addModelFrm.nColItem = 0
addModelFrm.nColQuant = 7
addModelFrm.sStore = XSTORE.BoundText
addModelFrm.bshowBal = True
addModelFrm.SMosm = xMosm.BoundText
addModelFrm.Sdiscount = Val(grid1.TextMatrix(grid1.Row, 9))
If xSupp.Value = 1 Then addModelFrm.sSupp = xCode.Text Else addModelFrm.SFact = xSubCode.Caption
If IsNumeric(grid1.TextMatrix(grid1.Row, 0)) And bAddModel Then
    Dim loctable As ADODB.Recordset
    Set loctable = ItemFind(grid1.TextMatrix(grid1.Row, 0), con)
    If Not (loctable.EOF And loctable.BOF) Then
        addModelFrm.SModelFact = loctable!modelfact0 & ""
        addModelFrm.SFact = loctable!Fact & ""
        addModelFrm.sSupp = loctable!SUPP & ""
        addModelFrm.SMosm = loctable!MOSM & ""
        addModelFrm.sModel = loctable!MODEL & ""
        addModelFrm.sSection = loctable!Section & ""
    End If
    loctable.Close
    Set loctable = Nothing
End If
bUpdated = False
addModelFrm.Show 1
If bUpdated Then
    If Not MYVALID Then Exit Sub
    xr_Discount.Text = Val(xRateSup.Caption)
    xr_Discount_Validate False
    myreplace
End If

End Sub



Private Function addBarCodeModel(cModel) As String
con.BeginTrans
On Error GoTo myerror
'con.Execute "DELETE FROM ADDPRINT WHERE DOC_NO = " & MyParn(xDoc_No.Text)
cString = "INSERT INTO ADDPRINT(DOC_NO,ITEM,QUANT,ISPRINT)"

cString = cString & _
          " SELECT DOC_NO,ITEM,QUANT,1 FROM FILE7_20"
cString = cString & turn(cString) & " DOC_NO = " & MyParn(xdoc_no.Text) & " AND MODEL = " & MyParn(cModel)
con.Execute cString


cString = " update FILE7_20H set isbar = 1"
cString = cString & turn(cString) & " doc_no = " & MyParn(xdoc_no.Text)
con.Execute cString
 
con.CommitTrans
Inform " „ «· ÕÊÌ· »‰Ã«Õ"
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub myReplaceOrder()
If myPublic = 0 Then
    con.Execute "DELETE FROM ORDERS1 WHERE MODELNO IN (SELECT FILE1_10.MODELNO FROM FILE7_20 INNER JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xdoc_no.Text) & ")"
ElseIf myPublic = 1 Then
    con.Execute "DELETE FROM ORDERS2 WHERE MODELNO IN (SELECT FILE1_10.MODELNO FROM FILE7_10 INNER JOIN FILE1_10 ON FILE7_10.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xdoc_no.Text) & ")"
End If
End Sub
Private Sub cmd_ex_Click()
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet
    Dim xlwbook As Excel.Workbook
    Dim cModelFact As String
    Dim cBarCode As String
    Dim cModel As String
    Dim cItem As String
    Dim nQTY As Double
    Dim nPrice As Double
    Dim nDisc As Double
    Dim nNewItem As String
    Dim lAddTrue As Boolean
    prog1.Min = 0
    prog1.Value = 0
    prog1.Visible = True
    If MsgBox("≈÷«›… »Ì«‰«  „‰ „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            If xdoc_no.Enabled Then
                If Not mysave Then Exit Sub
            End If
            
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
            For nRow = 2 To 10000
                If Len(xlsheet.Cells(nRow, 1)) > 1 Then
                    nRec = nRec + 1
                End If
                If Len(xlsheet.Cells(nRow, 1)) = 0 Then Exit For
            Next nRow
            prog1.Max = nRec
            i = 0
            For nRow = 2 To 10000
                If Len(xlsheet.Cells(nRow, 2)) = 0 Then Exit For
                If Len(xlsheet.Cells(nRow, 2)) > 1 Then
                    i = i + 1
'                   prog1.Value = i
                    cBarCode = xlsheet.Cells(nRow, 2)
                    nQTY = xlsheet.Cells(nRow, 4)
                    nPrice = xlsheet.Cells(nRow, 5)
                    nPrice2 = xlsheet.Cells(nRow, 9)
                    nDisc = xlsheet.Cells(nRow, 7)
                    cdesca = xlsheet.Cells(nRow, 3)
                    cItem = GetDesca("select ITEM from file1_10 where BARCODE = " & MyParn(cBarCode))
                    If Len(cBarCode) <= 10 Then
                        cModelFact = cBarCode
                    Else
                        cModelFact = Replace(cBarCode, "IFS", "")
                        cModelFact = Replace(cModelFact, "-", "")
                    End If
                    If cItem = "" Then
                        cItem = GetDesca("select ITEM from file1_10 where code = '066' and MODELFACT0 = " & MyParn(cModelFact))
                    End If
                    If cItem = "" Then
                        If Len(cBarCode) <= 10 Then
                            cModelFact = cBarCode
                        Else
                            cModelFact = Replace(cBarCode, "IFS", "")
                            cModelFact = Replace(cModelFact, "-", "")
                        End If
                        If Len(cModelFact) > 10 Then
                            MsgBox cBarCode
                            MsgBox cModelFact
                        Else
                            cItem = AddNewItem(cModelFact, cdesca, cBarCode, nPrice, nPrice2)
                            cModel = xCode.Text & "000" & RetZero(cModelFact)
                        End If
                    Else
                        con.Execute " UPDATE FILE1_10 SET PRICE = " & nPrice2 & " , COST = " & nPrice & " , DESCA = " & addstring(cdesca) & " WHERE ITEM = " & Val(cItem)
                    End If
                    If cItem <> "" Then
                        If GetDesca("select item from " & cFile & " where item = " & Val(cItem) & " and doc_no = " & MyParn(xdoc_no.Text)) = "" Then
                            cString = "Insert into " & cFile & " (doc_no,item,MODEL,discount,price,Quant)" & _
                                       "Values(" & _
                                       addstring(xdoc_no.Text) & "," & _
                                       addvalue(cItem) & "," & _
                                       addstring(cModel) & "," & _
                                       addstring(nDisc) & "," & _
                                       Val(nPrice) & "," & _
                                       Val(nQTY) & _
                                       ")"
                            con.Execute cString
                        End If
                        lAddTrue = True
                    End If
                End If
            Next nRow
            If lAddTrue Then
                Inform " „   ”ÃÌ· ›« Ê—… «·„‘ —Ì«  "
            Else
                Inform "·„ Ì „ «· ”ÃÌ· "
            End If
            myUndo
        End If
    End If
End Sub
Private Function AddNewItem(pModelFact, pDesca, pBarCode, pPRICE, pPRICE2)
Dim aInsert(19, 1)
Dim sitem As String, aRet As Variant
Dim cBarCode As String

cModel = xCode.Text & "000" & RetZero(pModelFact, 10)

aInsert(0, 0) = "Model"
aInsert(0, 1) = addstring(cModel)

aInsert(1, 0) = "ModelNO"
aInsert(1, 1) = addstring(cModel)

aInsert(2, 0) = "Desca"
aInsert(2, 1) = addstring(pDesca)

aInsert(3, 0) = "FACT"
aInsert(3, 1) = addstring(xCode.Text)

aInsert(4, 0) = "MOSM"
aInsert(4, 1) = addstring("000")

aInsert(5, 0) = "MODELFACT"
aInsert(5, 1) = addstring(RetZero(pModelFact, 10))

aInsert(6, 0) = "MODELFACT0"
aInsert(6, 1) = addstring(DelZero(pModelFact))

aInsert(7, 0) = "[section]"
aInsert(7, 1) = addstring(xSection.BoundText)

aInsert(8, 0) = "Code"
aInsert(8, 1) = addstring(xCode.Text)

aInsert(9, 0) = "[group]"
aInsert(9, 1) = addstring("00")

sitem = Val(GetDesca("SELECT MAX(ITEM) FROM FILE1_10", con)) + 1
aInsert(10, 0) = "item"
aInsert(10, 1) = addvalue(sitem)
                
aInsert(11, 0) = "Cost"
aInsert(11, 1) = pPRICE
                
aInsert(12, 0) = "PRICE"
aInsert(12, 1) = pPRICE2

aInsert(13, 0) = "Scal"
aInsert(13, 1) = addstring("---")
            
aInsert(14, 0) = "C_Scal"
aInsert(14, 1) = 1
                
aInsert(15, 0) = "Color"
aInsert(15, 1) = addstring("---")
        
aInsert(16, 0) = "C_Color"
aInsert(16, 1) = 1
        
aInsert(17, 0) = "Code_Color"
aInsert(17, 1) = 1
        
aInsert(18, 0) = "Code_Scal"
aInsert(18, 1) = 1
      
aInsert(19, 0) = "BARCODE"
aInsert(19, 1) = addstring(pBarCode)

con.Execute CreateInsert(aInsert, "File1_10")
AddNewItem = sitem
End Function
Private Sub Sales_FactLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
If myPublic = 0 Or myPublic = 2 Then
    Generalarray(1) = "SELECT   top 200 DOC_NO,[DATE] AS DATE1,DATE, FILE3_10.DESCA , FILE6_20H.DESCA " & _
                      " FROM  FILE6_20H INNER JOIN FILE3_10 ON FILE3_10.CODE = FILE6_20H.CODE WHERE FILE3_10.[GROUP] =  " & MyParn(pGrCust)
Else
    Generalarray(1) = "SELECT   top 200 DOC_NO,[DATE] AS DATE1,DATE, FILE3_10.DESCA , FILE6_10H.DESCA " & _
                      " FROM  FILE6_10H INNER JOIN FILE3_10 ON FILE3_10.CODE = FILE6_10H.CODE WHERE FILE3_10.[GROUP] =  " & MyParn(pGrCust)
End If
Generalarray(2) = "Order by Date DESC , DOC_NO "
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·—ﬁ„-≈”„  - «· «—ÌŒ "
If myPublic = 0 Or myPublic = 2 Then
    listarray(0, 1) = "(Doc_No Like '%cFilter%' or  FILE3_10.DESCA LIKE '%cFilter%' or  FILE6_20H.DESCA LIKE '%cFilter%' OR " & _
                      "##date##)"
Else
    listarray(0, 1) = "(Doc_No Like '%cFilter%' or  FILE3_10.DESCA LIKE '%cFilter%' or  FILE6_10H.DESCA LIKE '%cFilter%' OR " & _
                      "##date##)"
End If
GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·›—⁄ "
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "«·»Ì«‰ "
GrdArray(4, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search_FACT
Search_FACT.Caption = "«” ⁄·«„"
Search_FACT.Show 1
End Sub
Private Sub CMD_ADDINV_Click()
    XMODELNOT.Text = ""
    Sales_FactLookup
End Sub
Sub AddInvFact(pDoc)
    Dim InvFactTable As New ADODB.Recordset, nQ As Double, nPrice1 As Double, nPrice2 As Double
    Dim lIsEditPrice As Boolean
    Dim nQuant As Double
    If MsgBox(" ⁄œÌ· «”⁄«— «·»Ì⁄ Ê «·„’‰⁄  ", vbYesNo + vbDefaultButton2) = vbYes Then
        lIsEditPrice = True
    Else
        lIsEditPrice = False
    End If
    If myPublic = 0 Or myPublic = 2 Then
        InvFactTable.Open "SELECT FILE6_20.* , FILE1_10.BARCODE2 , FILE1_10.GRITEM ,file1_10.isdoz    FROM FILE1_10 INNER JOIN FILE6_20 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(pDoc), confact, adOpenStatic, adLockReadOnly
    Else
        InvFactTable.Open "SELECT FILE6_10.* , FILE1_10.BARCODE2 , FILE1_10.GRITEM ,file1_10.isdoz FROM FILE1_10 INNER JOIN FILE6_10 ON FILE6_10.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(pDoc), confact, adOpenStatic, adLockReadOnly
    End If
    With InvFactTable
        Do While Not .EOF
            If Not IsNull(!BARCODE2) Then
                If !Item = "18139942108001" Then
                    A = A
                End If
                
                nItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(!BARCODE2))
                If !GrItem = "01" And !ISDOZ = 1 Then
                    nQ = 12
                Else
                    nQ = 1
                End If
                If nItem <> "" Then
                    If lIsEditPrice Then
                        If !GrItem <> "01" Or !ISDOZ = 0 Then
                            nQ = 1
                            aRet1 = aGetDesca("SELECT PRICE , PRICE2 FROM FILE1_10 WHERE BARCODE2 = " & MyParn(!BARCODE2), confact)
                            aRet2 = aGetDesca("SELECT COST  , PRICE FROM FILE1_10 WHERE ITEM = " & nItem, con)
                            If UBound(aRet1) > 1 Then
                                nPrice1 = Val(aRet1(1) & "")
                                nPrice2 = Val(aRet1(2) & "")
                                If nPrice1 <> aRet2(1) Or nPrice2 <> aRet2(2) Then
                                    con.Execute " UPDATE FILE1_10 SET ISNEWPRICE=1 , COST = " & nPrice1 & " , cost2 = " & nPrice1 & " , PRICE = " & nPrice2 & " WHERE ITEM = " & nItem
                                    For sbr = 3 To nCountBranch
                                        cBr = RetZero(sbr, 2)
                                        con.Execute "   DELETE FROM NEW_ITEM WHERE ITEM = " & nItem
                                        con.Execute "   INSERT INTO NEW_ITEM (ITEM ,  BRANCH) Values ( " & nItem & " , " & addstring(cBr) & " ) "
                                    Next sbr
                                    
                                    For sbr = 61 To nCountBranch_fr
                                        cBr = RetZero(sbr, 2)
                                        con.Execute "   DELETE FROM NEW_ITEM_FR WHERE ITEM = " & nItem
                                        con.Execute "   INSERT INTO NEW_ITEM_FR (ITEM ,  BRANCH) Values ( " & nItem & " , " & addstring(cBr) & " ) "
                                    Next sbr
                                    Inform "  ⁄œÌ· ”⁄— " & !Item
                                End If
                            End If
                        Else
                            nQ = 12
                            aRet1 = aGetDesca("SELECT PRICE , PRICE2 FROM FILE1_10 WHERE BARCODE2 = " & MyParn(!BARCODE2), confact)
                            aRet2 = aGetDesca("SELECT COST  , PRICE FROM FILE1_10 WHERE ITEM = " & nItem, con)
                            If UBound(aRet1) > 1 Then
                                nPrice1 = Round(Val(aRet1(1) & "") / 12, 2)
                                nPrice2 = Val(aRet1(2) & "")
                                If nPrice1 <> Round(Val(aRet2(1) & "") / 12, 2) Or nPrice2 <> aRet2(2) Then
                                    con.Execute " UPDATE FILE1_10 SET ISNEWPRICE=1 ,COST = " & nPrice1 & " , cost2 = " & nPrice1 & " , PRICE = " & nPrice2 & " WHERE ITEM = " & nItem
                                    Inform "  ⁄œÌ· ”⁄— " & !Item
                                    con.Execute "   DELETE FROM NEW_ITEM WHERE ITEM = " & nItem
                                    For sbr = 3 To nCountBranch
                                        cBr = RetZero(sbr, 2)
                                        con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) Values ( " & nItem & " , " & addstring(cBr) & " ) "
                                    Next sbr
                                
                                    con.Execute "DELETE FROM NEW_ITEM_FR WHERE ITEM = " & nItem
                                    For sbr = 61 To nCountBranch_fr
                                        cBr = RetZero(sbr, 2)
                                        con.Execute "  INSERT INTO NEW_ITEM_FR (ITEM ,  BRANCH) Values ( " & nItem & " , " & addstring(cBr) & " ) "
                                    Next sbr
                                End If
                            End If
                        End If
                    End If
''''''''''''''''''''''''''''
'''''''''''''''''''«’‰«› »œÊ‰ ÿ·»Ì…
                    cQDem = GetDesca("SELECT QUANT FROM Model_Dem WHERE ITEM = " & nItem & " AND STORE = " & MyParn(XSTORE.BoundText))
                    If Val(cQDem) < (!Quant * nQ) Then
                        cString = "Insert into SUB7_20  (doc_no,STORE,item,QUANT,Q_DEM)" & _
                                   "Values(" & _
                                   addstring(xdoc_no.Text) & "," & _
                                   addstring(XSTORE.BoundText) & "," & _
                                   addvalue(nItem) & "," & _
                                   Val(!Quant * nQ) & "," & _
                                   Val(cQDem) & _
                                   ")"
                        con.Execute cString
                    End If

''''''''''''''''''''''''''''
                    nQuant = Round((!Quant * nQ), 0)
                    cString = "Insert into " & cFile & " (doc_no,item,Quant)" & _
                               "Values(" & _
                               addstring(xdoc_no.Text) & "," & _
                               addvalue(nItem) & "," & _
                               nQuant & _
                               ")"
                    con.Execute cString, nRec
                    Me.Caption = !Item
                Else
                    XMODELNOT.Text = XMODELNOT.Text & Chr(13) & !Item
                    Inform " »«—ﬂÊœ €Ì— ’ÕÌÕ " & !Item
                End If
            Else
                XMODELNOT.Text = XMODELNOT.Text & Chr(13) & !Item
                Inform " ’‰› €Ì— „”Ã· ›Ï «·„Õ· " & !Item
            End If
            .MoveNext
        Loop
            
        con.Execute " delete from NEWPRICEITEM where MODEL in ( SELECT     MODEL FROM        FILE1_10 WHERE     FILE1_10.ISNEWPRICE = 1 )  "
        con.Execute " INSERT INTO NEWPRICEITEM (MODEL) SELECT     MODEL FROM        FILE1_10 WHERE     (FILE1_10.ISNEWPRICE = 1) GROUP BY MODEL ", nRec
        
        Inform "  „  ⁄œÌ· «”⁄«— ⁄œœ „ÊœÌ·«   " & nRec
        
        Select Case myPublic
            Case 0
                con.Execute " UPDATE       FILE7_20  SET                MODEL = FILE1_10.MODEL, PRICE = FILE1_10.COST FROM            FILE1_10 INNER JOIN FILE7_20 ON FILE1_10.ITEM = FILE7_20.ITEM WHERE DOC_NO = " & MyParn(xdoc_no.Text)
                cDocSaleS = "Junior SHOP " & xdoc_no.Text
                confact.Execute " UPDATE   FILE6_20H SET  DESCA = " & addstring(cDocSaleS) & " WHERE DOC_NO = " & MyParn(pDoc)
                con.Execute " UPDATE   FILE6_20H SET  INV_NO = " & addstring(pDoc) & " WHERE DOC_NO = " & MyParn(xdoc_no.Text)
            Case 1
                con.Execute " UPDATE       FILE7_10  SET                MODEL = FILE1_10.MODEL, PRICE = FILE1_10.COST FROM            FILE1_10 INNER JOIN FILE7_10 ON FILE1_10.ITEM = FILE7_10.ITEM WHERE DOC_NO = " & MyParn(xdoc_no.Text)
                cDocSaleS = "Junior SHOP " & xdoc_no.Text
                confact.Execute " UPDATE   FILE6_10H SET  DESCA = " & addstring(cDocSaleS) & " WHERE DOC_NO = " & MyParn(pDoc)
                con.Execute " UPDATE   FILE6_10H SET  INV_NO = " & addstring(pDoc) & " WHERE DOC_NO = " & MyParn(xdoc_no.Text)
            Case 2
                con.Execute " UPDATE       FR7_20   SET MODEL = FILE1_10.MODEL, PRICE = FILE1_10.COST FROM  FILE1_10 INNER JOIN FR7_20 ON FILE1_10.ITEM = FR7_20.ITEM WHERE DOC_NO = " & MyParn(xdoc_no.Text)
                cDocSaleS = "Junior_Fr SHOP " & xdoc_no.Text
                confact.Execute " UPDATE   FILE6_20H SET  DESCA = " & addstring(cDocSaleS) & " WHERE DOC_NO = " & MyParn(pDoc)
                con.Execute " UPDATE   FILE6_20H SET  INV_NO = " & addstring(pDoc) & " WHERE DOC_NO = " & MyParn(xdoc_no.Text)
            Case 3
                con.Execute " UPDATE       FR7_10  SET                MODEL = FILE1_10.MODEL, PRICE = FILE1_10.COST FROM            FILE1_10 INNER JOIN FR7_10 ON FILE1_10.ITEM = FR7_10.ITEM WHERE DOC_NO = " & MyParn(xdoc_no.Text)
                cDocSaleS = "Junior SHOP " & xdoc_no.Text
                confact.Execute " UPDATE   FILE6_10H SET  DESCA = " & addstring(cDocSaleS) & " WHERE DOC_NO = " & MyParn(pDoc)
                con.Execute " UPDATE        FILE6_10H SET  INV_NO = " & addstring(pDoc) & " WHERE DOC_NO = " & MyParn(xdoc_no.Text)
        End Select
        myUndo
    End With
    MsgBox " OK "
End Sub
Sub EditAllItem()
    Dim InvFactTable As New ADODB.Recordset, nQ As Double, nPrice1 As Double, nPrice2 As Double
    Dim lIsEditPrice As Boolean
    If MsgBox(" ⁄œÌ· «”⁄«— «·»Ì⁄ Ê «·„’‰⁄  ", vbYesNo + vbDefaultButton2) = vbYes Then
        lIsEditPrice = True
    Else
        lIsEditPrice = False
        Exit Sub
    End If
    
    InvFactTable.Open "SELECT file1_10.item , file1_10.desca , FILE1_10.BARCODE2 , FILE1_10.GRITEM ,file1_10.isdoz FROM FILE1_10 where file1_10.model is not null = ", confact, adOpenStatic, adLockReadOnly
    prog1.Max = InvFactTable.RecordCount
    prog1.Min = 0
    prog1.Value = 0
    prog1.Visible = True
    With InvFactTable
        Do While Not .EOF
            i = i + 1
            prog1.Value = i
            If Not IsNull(!BARCODE2) Then
                nItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(!BARCODE2))
                If !GrItem = "01" And !ISDOZ = 1 Then
                    nQ = 12
                Else
                    nQ = 1
                End If
                If nItem <> "" Then
                    If lIsEditPrice Then
                        If !GrItem <> "01" Or !ISDOZ = 0 Then
                            nQ = 1
                            aRet1 = aGetDesca("SELECT PRICE , PRICE2 FROM FILE1_10 WHERE BARCODE2 = " & MyParn(!BARCODE2), confact)
                            aRet2 = aGetDesca("SELECT COST  , PRICE FROM FILE1_10 WHERE ITEM = " & nItem, con)
                            If UBound(aRet1) > 1 Then
                                nPrice1 = Val(aRet1(1) & "")
                                nPrice2 = Val(aRet1(2) & "")
                                If nPrice1 <> aRet2(1) Or nPrice2 <> aRet2(2) Then
                                    con.Execute " UPDATE FILE1_10 SET ISNEWPRICE=1 , ISNEWPRICE2=1 , COST = " & nPrice1 & " , PRICE = " & nPrice2 & " WHERE ITEM = " & nItem
                                    For sbr = 3 To nCountBranch
                                        cBr = RetZero(sbr, 2)
                                        con.Execute " DELETE FROM NEW_ITEM WHERE ITEM = " & nItem
                                        con.Execute " INSERT INTO NEW_ITEM (ITEM ,  BRANCH) Values ( " & nItem & " , " & addstring(cBr) & " ) "
                                    Next sbr

                                    For sbr = 61 To nCountBranch_fr
                                        cBr = RetZero(sbr, 2)
                                        con.Execute "   DELETE FROM NEW_ITEM_fr WHERE ITEM = " & nItem
                                        con.Execute "   INSERT INTO NEW_ITEM_fr (ITEM ,  BRANCH) Values ( " & nItem & " , " & addstring(cBr) & " ) "
                                    Next sbr

                                    Inform "  ⁄œÌ· ”⁄— " & !Item
                                End If
                            End If
                        Else
                            nQ = 12
                            aRet1 = aGetDesca("SELECT PRICE , PRICE2 FROM FILE1_10 WHERE BARCODE2 = " & MyParn(!BARCODE2), confact)
                            aRet2 = aGetDesca("SELECT COST  , PRICE FROM FILE1_10 WHERE ITEM = " & nItem, con)
                            If UBound(aRet1) > 1 Then
                                nPrice1 = Round(Val(aRet1(1) & "") / 12, 2)
                                nPrice2 = Val(aRet1(2) & "")
                                If nPrice1 <> Round(Val(aRet2(1) & "") / 12, 2) Or nPrice2 <> aRet2(2) Then
                                    con.Execute " UPDATE FILE1_10 SET ISNEWPRICE=1 , ISNEWPRICE2=1 , COST = " & nPrice1 & " , PRICE = " & nPrice2 & " WHERE ITEM = " & nItem
'                                    Inform "  ⁄œÌ· ”⁄— " & !Item
                                    con.Execute "   DELETE FROM NEW_ITEM WHERE ITEM = " & nItem
                                    For sbr = 2 To nCountBranch
                                        cBr = RetZero(sbr, 2)
                                        con.Execute "  INSERT INTO NEW_ITEM (ITEM ,  BRANCH) Values ( " & nItem & " , " & addstring(cBr) & " ) "
                                    Next sbr
                                
                                    For sbr = 61 To nCountBranch_fr
                                        cBr = RetZero(sbr, 2)
                                        con.Execute "  INSERT INTO NEW_ITEM_fr (ITEM ,  BRANCH) Values ( " & nItem & " , " & addstring(cBr) & " ) "
                                    Next sbr
                                End If
                            End If
                        End If
                    End If
                Else
'                   Inform " »«—ﬂÊœ €Ì— ’ÕÌÕ " & !Item
                End If
            Else
                Inform " ’‰› €Ì— „”Ã· ›Ï «·„Õ· " & !Item
            End If
            .MoveNext
        Loop
        con.Execute " delete from NEWPRICEITEM where MODEL in ( SELECT     MODEL FROM        FILE1_10 WHERE     FILE1_10.ISNEWPRICE = 1 )  "
        con.Execute " INSERT INTO NEWPRICEITEM (MODEL) SELECT     MODEL FROM        FILE1_10 WHERE     (FILE1_10.ISNEWPRICE = 1) GROUP BY MODEL ", nRec
        con.Execute " UPDATE FILE1_10 SET ISNEWPRICE = 0 "
        Inform "  „  ⁄œÌ· «”⁄«— ⁄œœ „ÊœÌ·«   " & nRec
    End With
    MsgBox " OK "
End Sub
Private Sub sned_purch_Click()
    Dim cDoc As String, cCode As String
    Dim aInsert As Variant
    If MsgBox("≈·Ï „” ‰œ „‘ —Ì«  ··›—Ê⁄", vbYesNo + vbDefaultButton2) = vbYes Then
        cCode = GetDesca("SELECT CODE FROM FILE4_10 WHERE BRANCH = " & MyParn(XSTORE.BoundText))
        If cCode = "" Then
            MsgBox " ·« ÌÊÃœ „Ê—œ ·Â–« «·›—⁄  "
            Exit Sub
        End If
        aInsert = AddFlag(Empty, "[DATE]", DateSq(Date))
        aInsert = AddFlag(aInsert, "[mosm]", addstring(cPMosm))
        aInsert = AddFlag(aInsert, "[STORE]", addstring("000"))
        aInsert = AddFlag(aInsert, "[code]", addstring(cCode))
        aInsert = AddFlag(aInsert, "[isdate]", "getdate()")
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(cusername))
        aInsert = AddFlag(aInsert, "[REMARK]", addstring(XNotes_Doc.Text & " " & XTRANS_DOC.Caption))
        aInsert = AddFlag(aInsert, "[ISNEW1]", 1)
            
        cDoc = RetZero(Newflag("FILE7_20H", "DOC_NO"))
        aInsert = AddFlag(aInsert, "DOC_NO", addstring(cDoc))
        con.Execute addInsert(aInsert, "FILE7_20h")
        
        con.Execute " INSERT INTO FILE7_20 (DOC_NO , quant, ITEM, price )  SELECT  " & addstring(cDoc) & " , quant, FR7_10.ITEM, file1_10.COST  From FR7_10 inner join file1_10 on FR7_10.item = file1_10.item WHERE FR7_10.item <> 0 and DOC_NO = " & MyParn(xdoc_no.Text)
        con.Execute " DELETE FROM NEW7_20H WHERE DOC_NO = " & addstring(cDoc)
        For nRow = 2 To nCountBranch
            sStore = RetZero(nRow, 2)
            con.Execute "  INSERT INTO NEW7_20H (DOC_NO , BRANCH) VALUES        ( " & addstring(cDoc) & " , " & addstring(sStore) & " )  "
        Next nRow
        
        Inform " „  ÕÊÌ· ≈·Ï „” ‰œ „‘ —Ì«  ··›—Ê⁄ —ﬁ„ " & cDoc
        openCardTable
        myUndo
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
Private Function SendPurchaseHeader(cFile, cFile2) As Boolean
'''''''On Error GoTo myerror
''''''Dim loctable As New ADODB.Recordset
''''''Dim loctable2 As New ADODB.Recordset, cStr1 As String, cStr2 As String, cStrDel1 As String, cStrDel2 As String, cNew1 As String, cNew2 As String
''''''loctable.Open "select * from " & cFile2 & " where doc_no = " & MyParn(xDoc_No.Text), con, adOpenStatic, adLockReadOnly, adCmdText
''''''Dim aInsert As Variant
''''''prog1.Value = 0
''''''prog1.Visible = True
''''''Dim nRecordCount As Long, nRecord As Long, nAffect As Long
''''''nRecordCount = loctable.RecordCount
''''''Do Until loctable.EOF
''''''    cStrDel1 = " delete from " & cFile & " where doc_no = " & MyParn(loctable!doc_no)
''''''    cStrDel2 = " delete from " & cFile2 & " where doc_no = " & MyParn(loctable!doc_no)
''''''    conShop2.Execute cStrDel1
''''''    conShop2.Execute cStrDel2
''''''    nRecord = nRecord + 1
''''''    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
''''''        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
''''''        aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
''''''        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
''''''        aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
''''''        aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
''''''        aInsert = AddFlag(aInsert, "[TAX]", Val(loctable!tax & ""))
''''''        aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM))
''''''        aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
''''''        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
''''''        conShop2.Execute addInsert(aInsert, cFile), nAffect
''''''
''''''        cStr2 = ""
''''''        Dim aInsert2 As Variant
''''''        If loctable2.State = adStateOpen Then loctable2.Close
''''''        loctable2.Open "SELECT " & cFile2 & ".* FROM " & cFile2 & " WHERE DOC_NO = " & MyParn(loctable!doc_no), con, adOpenStatic, adLockReadOnly, adCmdText
''''''        Do While Not loctable2.EOF
''''''            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
''''''            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
''''''            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
''''''            aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
''''''            cStr2 = cStr2 & addInsert(aInsert2, cFile2) & " ; "
''''''            loctable2.MoveNext
''''''        Loop
''''''        If cStr2 <> "" Then conShop2.Execute cStr2
''''''    End If
''''''    con.Execute cNew1
''''''Loop
''''''lastsub:
''''''prog1.Visible = False
''''''Exit Function
''''''myerror:
''''''MsgBox Err.Description
''''''Err.Clear
''''''SendPurchaseHeader = -1
''''''GoTo lastsub
End Function
Private Sub myUndo()
If IsNumeric(xdoc_no.Text) Then
    openCardTable xdoc_no.Text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then mydefine Else myload
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
openCardTable xdoc_no.Text, ">"
If CardTable.EOF Then openCardTable xdoc_no.Text
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xdoc_no.Text, "<"
If CardTable.EOF Then openCardTable xdoc_no.Text
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    mydefine
End If
End Sub
Sub EditAllItemPrice()
    Dim InvFactTable As New ADODB.Recordset, nQ As Double, nPrice1 As Double, nPrice2 As Double
    Dim SubModel As New ADODB.Recordset
    Dim lIsEditPrice As Boolean
    If MsgBox(" ⁄œÌ· «”⁄«— «·»Ì⁄ Ê «·„’‰⁄  ", vbYesNo + vbDefaultButton2) = vbYes Then
        lIsEditPrice = True
    Else
        lIsEditPrice = False
        Exit Sub
    End If
    If myPublic = 0 Then
        InvFactTable.Open "SELECT file1_10.MODELFACT0 FROM FILE1_10 INNER JOIN FILE7_20 ON FILE7_20.ITEM = FILE1_10.ITEM  where DOC_NO = " & MyParn(xdoc_no.Text) & " GROUP BY MODELFACT0 ", con, adOpenStatic, adLockReadOnly
    End If
    
    If myPublic = 2 Then
        InvFactTable.Open "SELECT file1_10.MODELFACT0 FROM FILE1_10 INNER JOIN FR7_20 ON FR7_20.ITEM = FILE1_10.ITEM  where DOC_NO = " & MyParn(xdoc_no.Text) & " GROUP BY MODELFACT0 ", con, adOpenStatic, adLockReadOnly
    End If
    prog1.Max = InvFactTable.RecordCount
    prog1.Min = 0
    prog1.Value = 0
    prog1.Visible = True
    With InvFactTable
        Do While Not .EOF
            i = i + 1
            prog1.Value = i
            If SubModel.State = adStateOpen Then SubModel.Close
            SubModel.Open " SELECT * FROM FILE1_10 WHERE MODELFACT0 = " & MyParn(!modelfact0), confact, adOpenStatic, adLockReadOnly, adCmdText
            Do While Not SubModel.EOF
                If Not IsNull(SubModel!BARCODE2) Then
                    nItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(SubModel!BARCODE2))
                    If SubModel!GrItem = "01" And SubModel!ISDOZ = 1 Then
                        nQ = 12
                    Else
                        nQ = 1
                    End If
                    If nItem <> "" Then
                        If lIsEditPrice Then
                            If SubModel!GrItem <> "01" Or SubModel!ISDOZ = 0 Then
                                nQ = 1
                                aRet1 = aGetDesca("SELECT PRICE , PRICE2 FROM FILE1_10 WHERE BARCODE2 = " & MyParn(SubModel!BARCODE2), confact)
                                aRet2 = aGetDesca("SELECT COST  , PRICE FROM FILE1_10 WHERE ITEM = " & nItem, con)
                                If UBound(aRet1) > 1 Then
                                    nPrice1 = Val(aRet1(1) & "")
                                    nPrice2 = Val(aRet1(2) & "")
                                    If nPrice1 <> aRet2(1) Or nPrice2 <> aRet2(2) Then
                                        con.Execute " UPDATE FILE1_10 SET COST = " & nPrice1 & " , PRICE = " & nPrice2 & " WHERE ITEM = " & nItem
                                    End If
                                End If
                            Else
                                nQ = 12
                                aRet1 = aGetDesca("SELECT PRICE , PRICE2 FROM FILE1_10 WHERE BARCODE2 = " & MyParn(SubModel!BARCODE2), confact)
                                aRet2 = aGetDesca("SELECT COST  , PRICE FROM FILE1_10 WHERE ITEM = " & nItem, con)
                                If UBound(aRet1) > 1 Then
                                    nPrice1 = Round(Val(aRet1(1) & "") / 12, 2)
                                    nPrice2 = Val(aRet1(2) & "")
                                    If nPrice1 <> Round(Val(aRet2(1) & "") / 12, 2) Or nPrice2 <> aRet2(2) Then
                                        con.Execute " UPDATE FILE1_10 SET ISNEWPRICE=1 , ISNEWPRICE2=1 , COST = " & nPrice1 & " , PRICE = " & nPrice2 & " WHERE ITEM = " & nItem
                                    End If
                                End If
                            End If
                         End If
                    End If
                Else
                    Inform " ’‰› €Ì— „”Ã· ›Ï «·„Õ· " & !Item
                End If
                SubModel.MoveNext
            Loop
            InvFactTable.MoveNext
        Loop
'        con.Execute " delete from NEWPRICEITEM where MODEL in ( SELECT     MODEL FROM        FILE1_10 WHERE     FILE1_10.ISNEWPRICE = 1 )  "
'        con.Execute " INSERT INTO NEWPRICEITEM (MODEL) SELECT     MODEL FROM        FILE1_10 WHERE     (FILE1_10.ISNEWPRICE = 1) GROUP BY MODEL ", nRec
'        con.Execute " UPDATE FILE1_10 SET ISNEWPRICE = 0 "
'        Inform "  „  ⁄œÌ· «”⁄«— ⁄œœ „ÊœÌ·«   " & nRec
    End With
    MsgBox " OK "
End Sub


