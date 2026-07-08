VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form CashReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9465
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   9465
   Begin Threed.SSCommand CMD_PAY 
      Height          =   1095
      Left            =   90
      TabIndex        =   60
      Top             =   4455
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   1931
      _Version        =   196610
      Font3D          =   2
      BackColor       =   16241358
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
      Picture         =   "Cashreg.frx":0000
      Caption         =   "”œ«œ «·»Ê‰"
      Alignment       =   1
      PictureAlignment=   3
   End
   Begin VB.Frame v 
      Caption         =   "”œ«œ ›Ì“« 2"
      Height          =   4335
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   45
      Width           =   4200
      Begin VB.OptionButton XTYPEVISA3 
         Alignment       =   1  'Right Justify
         Caption         =   "›Ì“« «Œ—Ï"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   2475
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.OptionButton XTYPEVISA3 
         Alignment       =   1  'Right Justify
         Caption         =   "PayMob"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   2430
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton XTYPEVISA3 
         Alignment       =   1  'Right Justify
         Caption         =   "›Ì“« »‰ﬂ €Ì— CIB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   1890
         Width           =   1860
      End
      Begin VB.OptionButton XTYPEVISA3 
         Alignment       =   1  'Right Justify
         Caption         =   "›Ì“« »‰ﬂ CIB"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   945
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1890
         Width           =   1410
      End
      Begin VB.OptionButton XTYPEVISA3 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ì“…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1890
         Width           =   690
      End
      Begin VB.OptionButton XTYPEVISA3 
         Alignment       =   1  'Right Justify
         Caption         =   "›Ì“« WIFI"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   1440
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.TextBox XVISA3_2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   270
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   3600
         Width           =   2175
      End
      Begin VB.TextBox XVISADOC3 
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
         Height          =   420
         Left            =   270
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   3105
         Width           =   2175
      End
      Begin VB.TextBox XVISA3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   270
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   315
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo XCODEVISA3 
         Height          =   390
         Left            =   270
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   810
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
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
      Begin VB.Line Line5 
         X1              =   4335
         X2              =   0
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line4 
         X1              =   4335
         X2              =   0
         Y1              =   2970
         Y2              =   2970
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         Caption         =   "⁄„Ê·…  ﬁ”Ìÿ : "
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
         Height          =   330
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   3690
         Width           =   1275
      End
      Begin VB.Label Label13 
         Appearance      =   0  'Flat
         Caption         =   "—ﬁ„ «·⁄„·Ì… :"
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
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   3240
         Width           =   1275
      End
      Begin VB.Label Label12 
         Appearance      =   0  'Flat
         Caption         =   "›Ì“« »‰ﬂ 2 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   900
         Width           =   1275
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         Caption         =   "”œ«œ ›Ì“« 2 : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   360
         Width           =   1365
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Height          =   645
      Left            =   4860
      TabIndex        =   28
      Top             =   8280
      Width           =   4605
      Begin VB.Frame Frame_okaz 
         Caption         =   "Œ’„ «·√Êﬂ«“ÊÌ‰"
         Height          =   690
         Left            =   45
         TabIndex        =   38
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
            TabIndex        =   39
            Top             =   315
            Width           =   1860
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
         TabIndex        =   37
         Top             =   1890
         Width           =   1320
      End
      Begin VB.Frame FRM_CLOSED 
         Height          =   1680
         Left            =   45
         TabIndex        =   32
         Top             =   3030
         Visible         =   0   'False
         Width           =   2040
         Begin Threed.SSCommand cmdAddItems 
            Height          =   825
            Left            =   90
            TabIndex        =   33
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
            Picture         =   "Cashreg.frx":0B29
            Caption         =   " ”ÃÌ· „ÊœÌ·«  "
            Alignment       =   8
            PictureAlignment=   6
         End
         Begin Threed.SSCommand cmd_closed 
            CausesValidation=   0   'False
            Height          =   555
            Left            =   90
            TabIndex        =   34
            Top             =   150
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   979
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "≈€·«ﬁ «·„” ‰œ"
         End
         Begin Threed.SSCommand cmd_CLOSEDDATE 
            CausesValidation=   0   'False
            Height          =   1005
            Left            =   1035
            TabIndex        =   35
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
            Picture         =   "Cashreg.frx":3331
            Caption         =   "≈€·«ﬁ › —…"
            Alignment       =   8
            PictureAlignment=   6
         End
         Begin Threed.SSCommand cmd_open 
            CausesValidation=   0   'False
            Height          =   1005
            Left            =   90
            TabIndex        =   36
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
            Picture         =   "Cashreg.frx":5E06
            Caption         =   "› Õ › —…"
            Alignment       =   8
            PictureAlignment=   6
         End
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   135
         Value           =   -1  'True
         Width           =   1230
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
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   135
         Width           =   1455
      End
      Begin VB.OptionButton XONLINE 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Caption         =   "«Ê‰ ·«Ì‰  Ê’Ì·"
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
         Height          =   375
         Index           =   2
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   135
         Width           =   1545
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
         TabIndex        =   40
         Top             =   1890
         Width           =   1995
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   225
      Top             =   -135
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
   Begin VB.Frame Frame1 
      Caption         =   "”œ«œ ›Ì“« 1"
      Height          =   8205
      Left            =   4860
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   4605
      Begin VB.TextBox XVISADOC 
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
         Height          =   420
         Left            =   810
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   7200
         Width           =   1995
      End
      Begin VB.TextBox XPOINT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3195
         Width           =   2175
      End
      Begin VB.Frame FRMVISA 
         Caption         =   "ÿ—Ìﬁ… ”œ«œ «·›Ì“« "
         Height          =   3390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   3690
         Width           =   4380
         Begin VB.OptionButton XTYPEVISA 
            Alignment       =   1  'Right Justify
            Caption         =   "›Ì“« «Œ—Ï "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   9
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   62
            Top             =   2925
            Width           =   1815
         End
         Begin VB.OptionButton XTYPEVISA 
            Alignment       =   1  'Right Justify
            Caption         =   "„Ì“…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   8
            Left            =   90
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2295
            Width           =   690
         End
         Begin VB.OptionButton XTYPEVISA 
            Alignment       =   1  'Right Justify
            Caption         =   "PayMob"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   7
            Left            =   2340
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   2835
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.OptionButton XTYPEVISA 
            Alignment       =   1  'Right Justify
            Caption         =   "›Ì“« »‰ﬂ CIB"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   6
            Left            =   945
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   2295
            Width           =   1410
         End
         Begin VB.OptionButton XTYPEVISA 
            Alignment       =   1  'Right Justify
            Caption         =   "›Ì“« »‰ﬂ €Ì— CIB"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   5
            Left            =   2295
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   2295
            Width           =   1860
         End
         Begin VB.OptionButton XTYPEVISA 
            Alignment       =   1  'Right Justify
            Caption         =   "›Ì“« WIFI"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   4
            Left            =   2070
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   1800
            Visible         =   0   'False
            Width           =   2085
         End
         Begin VB.TextBox XVISA2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   90
            MaxLength       =   15
            RightToLeft     =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1080
            Width           =   2040
         End
         Begin VB.OptionButton XTYPEVISA 
            Alignment       =   1  'Right Justify
            Caption         =   "›Ì“« »‰ﬂ «Â·Ï"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   2340
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   690
            Width           =   1815
         End
         Begin VB.OptionButton XTYPEVISA 
            Alignment       =   1  'Right Justify
            Caption         =   " ﬁ”Ìÿ »⁄„Ê·…"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   3
            Left            =   135
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   720
            Width           =   2085
         End
         Begin VB.OptionButton XTYPEVISA 
            Alignment       =   1  'Right Justify
            Caption         =   " ﬁ”Ìÿ «Â·Ï 6 ‘ÂÊ—"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   180
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   270
            Width           =   1995
         End
         Begin VB.OptionButton XTYPEVISA 
            Alignment       =   1  'Right Justify
            Caption         =   "›Ì“« »‰ﬂ €Ì— «·«Â·Ï"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   2115
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   315
            Width           =   2040
         End
         Begin VB.CheckBox xisbankahly 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬂ«—  «·⁄„Ì· »‰ﬂ «Â·Ï"
            Height          =   645
            Left            =   3330
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1035
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Line Line1 
            X1              =   4335
            X2              =   0
            Y1              =   2745
            Y2              =   2745
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            Caption         =   "⁄„Ê·…  ﬁ”Ìÿ :"
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
            Height          =   330
            Left            =   2250
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   1170
            Width           =   1140
         End
         Begin VB.Line Line3 
            X1              =   4335
            X2              =   0
            Y1              =   2205
            Y2              =   2205
         End
         Begin VB.Line Line2 
            X1              =   4335
            X2              =   0
            Y1              =   1755
            Y2              =   1755
         End
      End
      Begin VB.TextBox xPay 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   1215
         Width           =   2175
      End
      Begin VB.TextBox xVisa 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1710
         Width           =   2175
      End
      Begin MSDataListLib.DataCombo xBoxVISA 
         Height          =   390
         Left            =   3780
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2700
         Visible         =   0   'False
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
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
      Begin MSDataListLib.DataCombo xcodevisa 
         Height          =   390
         Left            =   135
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2700
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
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
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         Caption         =   "—ﬁ„ «·⁄„·Ì… :"
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
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   7290
         Width           =   1275
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         Caption         =   "»œ· ‰ﬁ«ÿ »‰ﬂ «Â·Ï "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   3240
         Width           =   1770
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         Caption         =   "—ﬁ„  ·Ì›Ê‰ «·⁄„Ì· :"
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
         Height          =   330
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   7785
         Width           =   1545
      End
      Begin VB.Label xphone 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   7740
         Width           =   1995
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "›Ì“« »‰ﬂ 1 :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2790
         Width           =   1275
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "«·»«ﬁÏ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2250
         Width           =   1275
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "«·„œ›Ê⁄ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1260
         Width           =   1275
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "≈Ã„«·Ï ﬁÌ„… «·»Ê‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   765
         Width           =   1995
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "⁄œœ «·ﬁÿ⁄ «·„»«⁄… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   315
         Width           =   1905
      End
      Begin VB.Label xtotalQuant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   2175
      End
      Begin VB.Label xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label xRest 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   2205
         Width           =   2175
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "”œ«œ ›Ì“« 1 : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1755
         Width           =   1410
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   2381
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
Attribute VB_Name = "CashReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim lExit As Boolean
Dim bSave As Boolean, cTypeVisa As String, cTypeVisa3 As String
Public myForm As Form
Private Sub CMD_PAY_Click()
    MakePayment

End Sub

Private Sub Form_Load()
Dim LocVisaTable As New ADODB.Recordset
con.CursorLocation = adUseClient
If con.State = adStateOpen Then con.Close
con.Open strCon

LocVisaTable.Open "SELECT * FROM FILE6_20H WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text), con, adOpenStatic, adLockReadOnly, adCmdText
With LocVisaTable
'xisbankahly.Visible = (cBranch <= "60")

data2.ConnectionString = strCon
If lIsBranchStore Then
    data2.RecordSource = "SELECT * FROM VISA WHERE branch = '00' ORDER BY CODE "
Else
    data2.RecordSource = "SELECT * FROM visa WHERE branch = " & MyParn(cBranch) & " ORDER BY CODE "
End If
Set xcodevisa.RowSource = data2
xcodevisa.ListField = "Desca"
xcodevisa.BoundColumn = "Code"

DATA3.ConnectionString = strCon
If lIsBranchStore Then
    DATA3.RecordSource = "SELECT * FROM VISA WHERE branch = '00' and code <> '14' ORDER BY CODE "
Else
    DATA3.RecordSource = "SELECT * FROM visa WHERE branch = " & MyParn(cBranch) & " and code <> '14' ORDER BY CODE "
End If
Set XCODEVISA3.RowSource = DATA3
XCODEVISA3.ListField = "Desca"
XCODEVISA3.BoundColumn = "Code"

xTotal.Caption = Myvalue(myForm.xTotal.text, "Fixed")
xPay.text = Myvalue(!PAY, "Fixed")
XPOINT.text = Myvalue(!Point, "Fixed")
XVISADOC.text = !VISADOC & ""
XVISADOC3.text = !VISADOC3 & ""

xVisa.text = Myvalue(!Visa, "Fixed")
XVISA3.text = Myvalue(!VISA3, "Fixed")
XVISA3_2.text = Myvalue(!VISA3_2, "Fixed")
XVISA2.text = Myvalue(!VISA2, "Fixed")
xphone.Caption = myForm.xphone.Caption
xcodevisa.BoundText = !CODEVISA & ""
XCODEVISA3.BoundText = !CODEVISA3 & ""

If !TYPEVISA = 0 Then XTYPEVISA(0).Value = 1
If !TYPEVISA = 2 Then XTYPEVISA(2).Value = 1
If !TYPEVISA = 3 Then XTYPEVISA(3).Value = 1
If !TYPEVISA = 4 Then XTYPEVISA(4).Value = 1
If !TYPEVISA = 5 Then XTYPEVISA(5).Value = 1
If !TYPEVISA = 6 Then XTYPEVISA(6).Value = 1
If !TYPEVISA = 7 Then XTYPEVISA(7).Value = 1
If !TYPEVISA = 8 Then XTYPEVISA(8).Value = 1
If !TYPEVISA = 9 Then XTYPEVISA(9).Value = 1




If !TYPEVISA3 = 4 Then XTYPEVISA3(4).Value = 1
If !TYPEVISA3 = 5 Then XTYPEVISA3(5).Value = 1
If !TYPEVISA3 = 6 Then XTYPEVISA3(6).Value = 1
If !TYPEVISA3 = 7 Then XTYPEVISA3(7).Value = 1
If !TYPEVISA3 = 8 Then XTYPEVISA3(8).Value = 1
If !TYPEVISA3 = 9 Then XTYPEVISA3(9).Value = 1

XONLINE(0).Value = myForm.XONLINE(0).Value
XONLINE(1).Value = myForm.XONLINE(1).Value
XONLINE(2).Value = myForm.XONLINE(2).Value

'If myForm.XTYPEVISA.Caption = "" And cBranch > "60" Then XTYPEVISA(0).Value = True

'If Val(xvisa.text) <> 0 Then FRMVISA.Visible = True

xtotalQuant.Caption = Myvalue(myForm.xtotalQuant.Caption, "Fixed")
xisbankahly.Value = IIf(myForm.xisbankahly.Value, 1, 0)
End With

If lOneVisa Then xcodevisa.BoundText = sCodeVisaBranch
CalcTotals
End Sub
Private Sub Form_Unload(Cancel As Integer)
    myForm.lSave = bSave
    Set CashReg = Nothing
End Sub
Private Sub xcodevisa_LostFocus()
If xcodevisa.BoundText = "14" Then XTYPEVISA(0).Value = 1
If xcodevisa.BoundText = "16" Then XTYPEVISA(5).Value = 1
If xcodevisa.BoundText = "17" Then XTYPEVISA(7).Value = 1
If xcodevisa.BoundText <> "16" And xcodevisa.BoundText <> "17" Then XTYPEVISA(9) = True

End Sub

Private Sub XCODEVISA3_LostFocus()
If XCODEVISA3.BoundText <> "16" And XCODEVISA3.BoundText <> "17" Then XTYPEVISA3(9) = True
End Sub
Private Sub xPay_Change()
    CalcTotals
End Sub
Private Sub xPOINT_Change()
    CalcTotals
End Sub
Private Sub xPay_GotFocus()
xPay.SelStart = 0
xPay.SelLength = Len(xPay.text)
End Sub
Private Sub xPOINT_GotFocus()

XPOINT.SelStart = 0
XPOINT.SelLength = Len(XPOINT.text)
End Sub
Private Sub xRet_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Exit Sub
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub
Private Sub CalcTotals()
    xRest.Caption = Format(Val(xPay.text) + Val(xVisa.text) + Val(XVISA3.text) + Val(XPOINT.text) - Val(xTotal.Caption), "#0.00")
End Sub
Private Sub xPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then MakePayment
End Sub

Private Sub xvisa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then MakePayment
End Sub
Private Sub xPOINT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then MakePayment
End Sub
Private Sub xvisaDOC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then MakePayment
End Sub
Private Sub xvisa2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then MakePayment
End Sub
Private Sub xVisa_Change()
    CalcTotals
End Sub
Private Sub xVisa3_Change()
    CalcTotals
End Sub

Private Sub xVisa_GotFocus()
    FRMVISA.Visible = True
End Sub
Private Function MakePayment() As Boolean
On Error Resume Next
Dim cPhone As String, nCash As Double

nCash = Val(xTotal.Caption) - Val(xVisa.text) - Val(XPOINT.text) - Val(XVISA3.text)
xRest.Caption = Format(Val(xPay.text) + Val(xVisa.text) + Val(XVISA3.text) + Val(XPOINT.text) - Val(xTotal.Caption), "#0.00")


If xcodevisa.BoundText = "" And Val(XPOINT.text) > 0 Then
    MsgBox " ÌÃ»  ”ÃÌ· ‰Ê⁄ «·»‰ﬂ "
    Exit Function
End If
If SalesFrm.xCode.text = "0000" Then
    If Abs(Round(Val(xPay.text) + Val(xVisa.text) + Val(XVISA3.text) + Val(XPOINT.text), 2)) < (Round(Val(xTotal.Caption), 2)) Then
        MsgBox " ﬁÌ„… «ﬁ· „‰ ﬁÌ„… «·»Ê‰ "
        Exit Function
    End If
End If
''''  «Ê· ›Ì“«
If Val(xVisa.text) <> 0 Then
    If xcodevisa.BoundText <> "" And XVISADOC.text = "" Then
        MsgBox " ÌÃ»  ”ÃÌ· —ﬁ„ «·⁄„Ì·… "
        XVISADOC.SetFocus
        Exit Function
    End If

    If xcodevisa.BoundText = "14" Then
        If XTYPEVISA(0).Value = 0 And XTYPEVISA(1).Value = 0 And XTYPEVISA(2).Value = 0 And XTYPEVISA(3).Value = 0 And XTYPEVISA(4).Value = 0 Then
            MsgBox " ÌÃ»  ÕœÌœ ‰Ê⁄ ⁄„·Ì… ›Ì“« «·«Â·Ï"
            Exit Function
        End If
        If XTYPEVISA(5).Value = 1 Or XTYPEVISA(6).Value = 1 Then
            MsgBox " ÌÃ»  ÕœÌœ ‰Ê⁄ ⁄„·Ì… ›Ì“« «·«Â·Ï"
            Exit Function
        End If
    End If

    If xcodevisa.BoundText = "16" Then
'        If XTYPEVISA(5).Value = 1 Or XTYPEVISA(6).Value = 1 Or XTYPEVISA(8).Value = 1 Or XTYPEVISA(9).Value = 1 Then
'            MsgBox " ÌÃ»  ÕœÌœ ‰Ê⁄ ⁄„·Ì… ›Ì“« CIB"
'            Exit Function
'        End If
    
        If XTYPEVISA(5).Value = False And XTYPEVISA(6).Value = False And XTYPEVISA(8).Value = FALCE And XTYPEVISA(9).Value Then
            MsgBox " ÌÃ»  ÕœÌœ ‰Ê⁄ ⁄„·Ì… ›Ì“« CIB"
            Exit Function
        End If
    
        If XTYPEVISA(5).Value = False And XTYPEVISA(6).Value = False And XTYPEVISA(8).Value = FALCE And XTYPEVISA(9).Value Then
            MsgBox " ÌÃ»  ÕœÌœ ‰Ê⁄ ⁄„·Ì… ›Ì“« CIB"
            Exit Function
        End If
    
    End If


    If XTYPEVISA(3).Value <> 0 And Val(XVISA2.text) = 0 Then
        MsgBox " ”ÃÌ· ﬁÌ„… ⁄„Ê·… «· ﬁ”Ìÿ ··»‰ﬂ"
        XVISA2.SetFocus
        Exit Function
    End If

    If xcodevisa.BoundText = "17" And XTYPEVISA(3).Value = 0 And XTYPEVISA(7).Value = 0 Then
        MsgBox " ÕœÌœ ‰Ê⁄ ›Ì“« PayMob €Ì— ’ÕÌÕ "
        Exit Function
    End If

    If XTYPEVISA(1).Value <> 0 Or XTYPEVISA(2).Value <> 0 Or XTYPEVISA(3).Value <> 0 Then xisbankahly.Value = True

    If Abs(Val(xVisa.text) + Val(XVISA3.text)) > Abs(Val(xTotal.Caption)) + 100 And (Val(xVisa.text) + Val(XVISA3.text)) > 0 Then
        MsgBox "«·„”œœ ›Ì“« «ﬂ»— „‰ ﬁÌ„… «·»Ê‰"
        Exit Function
    End If

    If (Val(xVisa.text) + Val(XPOINT.text)) <> 0 And xcodevisa.BoundText = "" Then
        MsgBox " ÕœÌœ ‰Ê⁄ «·›Ì“« «·„” Œœ„… "
        xcodevisa.SetFocus
        Exit Function
    End If
End If
'''''''''''''''  ›Ì“« 2
If Val(XVISA3.text) <> 0 Then
    If XCODEVISA3.BoundText <> "" And XVISADOC3.text = "" Then
        MsgBox " ÌÃ»  ”ÃÌ· —ﬁ„ «·⁄„Ì·… ··›Ì“« 2"
        XVISADOC3.SetFocus
        Exit Function
    End If

    If XCODEVISA3.BoundText = "16" Then
        If XTYPEVISA3(5).Value = 0 And XTYPEVISA3(6).Value = 0 And XTYPEVISA3(8).Value = 0 And XTYPEVISA3(9).Value = 0 Then
            MsgBox " ÌÃ»  ÕœÌœ ‰Ê⁄ ⁄„·Ì… ›Ì“«2 CIB"
            Exit Function
        End If
    End If
    If XCODEVISA3.BoundText <> "16" Then XTYPEVISA3(9).Value = 1


'    If XTYPEVISA3(3).Value <> 0 And Val(XVISA3_2.text) = 0 Then
'        MsgBox " ”ÃÌ· ﬁÌ„… ⁄„Ê·… «· ﬁ”Ìÿ ··»‰ﬂ 2"
'        XVISA3_2.SetFocus
'        Exit Function
'    End If

'    If XCODEVISA3.BoundText = "17" And XTYPEVISA3(7).Value = 0 Then
'        MsgBox " ÕœÌœ ‰Ê⁄ ›Ì“«2  PayMob €Ì— ’ÕÌÕ "
'        Exit Function
'    End If


    If Abs(Val(xVisa.text) + Val(XVISA3.text)) > Abs(Val(xTotal.Caption)) + 100 And (Val(xVisa.text) + Val(XVISA3.text)) > 0 Then
        MsgBox "«·„”œœ ›Ì“« «ﬂ»— „‰ ﬁÌ„… «·»Ê‰"
        Exit Function
    End If

    If (Val(xVisa.text) + Val(XVISA3.text) + Val(XPOINT.text)) <> 0 And XCODEVISA3.BoundText = "" Then
        MsgBox " ÕœÌœ ‰Ê⁄ «·›Ì“«2 «·„” Œœ„… "
        XCODEVISA3.SetFocus
        Exit Function
    End If
End If



If Trim(xPay.text) = "" And Trim(xVisa.text) = "" And Trim(XVISA3.text) = "" And Trim(XPOINT.text) = "" Then Exit Function
'nTot = Val(GetDesca("SELECT SUM(TOTAL) FROM FILE6_20 WHERE DOC_NO = " & MyParn(SalesFrm.xDoc_No.text), con) & "") - Val(SalesFrm.xDiscount.text)
'If Int(Val(xTotal.Caption)) <> Int(nTot) Then
'    MsgBox "—«Ã⁄ ﬁÌ„… «·ÌÊ‰ "
'    Unload Me
'End If
If Val(xRest.Caption) >= 0 Then
    If Val(xVisa.text) + Val(XVISA3.text) > 0 Then
        If xphone.Caption = "" Then
            cPhone = InputBox("", " —ﬁ„  ·Ì›Ê‰ «·⁄„Ì· ")
            If cPhone <> "" Then xphone.Caption = cPhone
        End If
    End If
    myForm.xPay.Caption = Val(xPay.text)
    myForm.XPOINT.Caption = Val(XPOINT.text)
    myForm.xVisa.Caption = Val(xVisa.text) + Val(XVISA3.text)
    myForm.XVISA2.Caption = XVISA2.text
    myForm.xphone.Caption = xphone.Caption
    
    myForm.XONLINE(0).Value = XONLINE(0).Value
    myForm.XONLINE(1).Value = XONLINE(1).Value
    myForm.XONLINE(2).Value = XONLINE(2).Value
    
    If XTYPEVISA(0).Value Then myForm.XTYPEVISA.Caption = 0
    If XTYPEVISA(1).Value Then myForm.XTYPEVISA.Caption = 1
    If XTYPEVISA(2).Value Then myForm.XTYPEVISA.Caption = 2
    If XTYPEVISA(3).Value Then myForm.XTYPEVISA.Caption = 3
    If XTYPEVISA(4).Value Then myForm.XTYPEVISA.Caption = 4
    If XTYPEVISA(5).Value Then myForm.XTYPEVISA.Caption = 5
    If XTYPEVISA(6).Value Then myForm.XTYPEVISA.Caption = 6
    If XTYPEVISA(7).Value Then myForm.XTYPEVISA.Caption = 7
    If XTYPEVISA(8).Value Then myForm.XTYPEVISA.Caption = 8
    If XTYPEVISA(9).Value Then myForm.XTYPEVISA.Caption = 9
    
    If XTYPEVISA(1).Value Then cTypeVisa = 1
    If XTYPEVISA(2).Value Then cTypeVisa = 2
    If XTYPEVISA(3).Value Then cTypeVisa = 3
    If XTYPEVISA(4).Value Then cTypeVisa = 4
    If XTYPEVISA(5).Value Then cTypeVisa = 5
    If XTYPEVISA(6).Value Then cTypeVisa = 6
    If XTYPEVISA(7).Value Then cTypeVisa = 7
    If XTYPEVISA(8).Value Then cTypeVisa = 8
    If XTYPEVISA(9).Value Then cTypeVisa = 9

    If XTYPEVISA3(4).Value Then cTypeVisa3 = 4
    If XTYPEVISA3(5).Value Then cTypeVisa3 = 5
    If XTYPEVISA3(6).Value Then cTypeVisa3 = 6
    If XTYPEVISA3(7).Value Then cTypeVisa3 = 7
    If XTYPEVISA3(8).Value Then cTypeVisa3 = 8
    If XTYPEVISA3(9).Value Then cTypeVisa3 = 9

    
    
    myForm.xRest.Caption = xRest.Caption
    myForm.XVISADOC.Caption = XVISADOC.text
    myForm.xcash.Caption = Val(xTotal.Caption) - Val(xVisa.text) - Val(XVISA3.text) - Val(XPOINT.text)
    
    If Val(XPOINT.text) > 0 And Val(myForm.xcash.Caption) < 0 Then myForm.xcash.Caption = 0
    myForm.xisbankahly.Value = IIf(xisbankahly.Value, 1, 0)
    myForm.xcodevisa.BoundText = xcodevisa.BoundText
        
    bSave = True
'    Unload Me
ElseIf ActiveControl.Name = xPay.Name Then
    xVisa.SetFocus
ElseIf myForm.xCode.text <> "0000" Or myForm.xcodeDesca3.Caption <> "" Then

ElseIf myForm.xCode.text <> "0000" Or myForm.xcodeDesca3.Caption <> "" Then
    If MsgBox("”œ«œ " & xPay.text & " „‰ ﬁÌ„… «·›« Ê—… Ê «·»«ﬁÏ  " & Abs(Val(xRest.Caption)) & " √Ã·", vbYesNo + vbDefaultButton2) = vbYes Then
        myForm.xPay.Caption = xPay.text
        myForm.XPOINT.Caption = XPOINT.text
        myForm.xVisa.Caption = Val(xVisa.text) + Val(XVISA3.text)
        
        myForm.xRest.Caption = xRest.Caption
        myForm.XVISADOC.Caption = XVISADOC.text
        myForm.xcash.Caption = Val(xTotal.Caption) - Val(xVisa.text) - Val(XVISA3.text) - Val(XPOINT.text)
        
        myForm.XVISA2.Caption = XVISA2.text
        If XTYPEVISA(0).Value = 1 Then myForm.XTYPEVISA.Caption = 0
        If XTYPEVISA(1).Value = 1 Then myForm.XTYPEVISA.Caption = 1
        If XTYPEVISA(2).Value = 1 Then myForm.XTYPEVISA.Caption = 2
        If XTYPEVISA(3).Value = 1 Then myForm.XTYPEVISA.Caption = 3
        If XTYPEVISA(4).Value = 1 Then myForm.XTYPEVISA.Caption = 4
        If XTYPEVISA(5).Value = 1 Then myForm.XTYPEVISA.Caption = 5
        If XTYPEVISA(6).Value = 1 Then myForm.XTYPEVISA.Caption = 6
        If XTYPEVISA(7).Value = 1 Then myForm.XTYPEVISA.Caption = 7
        If XTYPEVISA(8).Value = 1 Then myForm.XTYPEVISA.Caption = 8
        If XTYPEVISA(9).Value = 1 Then myForm.XTYPEVISA.Caption = 9
        
        myForm.XVISADOC.Caption = XVISADOC.text
        myForm.xcash.Caption = Val(xTotal.Caption) - Val(xVisa.text) - Val(XVISA3.text) - Val(XPOINT.text)
        myForm.xisbankahly.Value = IIf(xisbankahly.Value, 1, 0)
        myForm.xcodevisa.BoundText = xcodevisa.BoundText
        myForm.xRest.Caption = ""
        
        bSave = True
 '       Unload Me
    End If
End If
nCash = Val(xTotal.Caption) - Val(xVisa.text) - Val(XPOINT.text) - Val(XVISA3.text)
xRest.Caption = Format(Val(xPay.text) + Val(xVisa.text) + Val(XVISA3.text) + Val(XPOINT.text) - Val(xTotal.Caption), "#0.00")


If bSave Then
'''''''''''''  SAVE VISA
        
        
        
    con.Execute " UPDATE FILE6_20H SET " & _
                "  cash = " & Val(nCash) & " , REST = " & Val(xRest.Caption) & _
                ", VISA = " & Val(xVisa.text) & " , CODEVISA = " & addstring(xcodevisa.BoundText) & " , VISADOC = " & addstring(XVISADOC.text) & " , VISA2 = " & Val(XVISA2.text) & " , TYPEVISA = " & MyParn(cTypeVisa) & " , POINT = " & Val(XPOINT.text) & _
                " ,VISA3 = " & Val(XVISA3.text) & " , CODEVISA3= " & addstring(XCODEVISA3.BoundText) & " , VISADOC3 = " & addstring(XVISADOC3.text) & " , VISA3_2 = " & Val(XVISA3_2.text) & " , TYPEVISA3 = " & MyParn(cTypeVisa3) & _
                " ,ISNEW = 1 , PHONE = " & addstring(xphone.Caption) & _
                " WHERE DOC_NO = " & MyParn(SalesFrm.xdoc_no.text)
    
    Unload Me
 
End If
Err.Clear
End Function
