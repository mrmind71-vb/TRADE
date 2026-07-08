VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form banktrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕÊÌ·«  »Ì‰ Õ”«» «·»‰Êﬂ"
   ClientHeight    =   4485
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   9300
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   4485
   ScaleWidth      =   9300
   Begin VB.Frame Frame9 
      Height          =   870
      Left            =   5085
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   3510
      Width           =   4155
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
         TabIndex        =   34
         Top             =   -270
         Visible         =   0   'False
         Width           =   105
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
         TabIndex        =   33
         Top             =   135
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
         TabIndex        =   32
         Top             =   135
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
         TabIndex        =   31
         Top             =   495
         Width           =   1815
      End
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
         TabIndex        =   30
         Top             =   495
         Width           =   2220
      End
   End
   Begin VB.CheckBox xClosed 
      Alignment       =   1  'Right Justify
      Caption         =   "„” ‰œ „€·ﬁ"
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
      Height          =   315
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   855
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame Frame1 
      Height          =   1725
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   720
      Width           =   1995
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   555
         Left            =   45
         TabIndex        =   25
         Top             =   1080
         Width           =   1905
         _ExtentX        =   3360
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
         Picture         =   "banktrans.frx":0000
         Alignment       =   4
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   915
         Left            =   990
         TabIndex        =   26
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1614
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
         Picture         =   "banktrans.frx":25CC
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   915
         Left            =   45
         TabIndex        =   27
         Top             =   135
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1614
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
         Picture         =   "banktrans.frx":4C44
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   3735
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   20
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
         Picture         =   "banktrans.frx":7285
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "banktrans.frx":9455
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   21
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
         Picture         =   "banktrans.frx":B59D
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "banktrans.frx":D765
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   22
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
         Picture         =   "banktrans.frx":F8B4
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "banktrans.frx":11A94
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   23
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
         Picture         =   "banktrans.frx":13BEF
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "banktrans.frx":15DAB
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   1170
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   7260
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   6030
         Picture         =   "banktrans.frx":17EFA
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdAdd 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   4830
         MaskColor       =   &H00FFFFFF&
         Picture         =   "banktrans.frx":1A6CD
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1275
         MaskColor       =   &H00FFFFFF&
         Picture         =   "banktrans.frx":1CC79
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "banktrans.frx":1F513
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         Picture         =   "banktrans.frx":2197F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
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
         Left            =   3645
         MaskColor       =   &H00FFFFFF&
         Picture         =   "banktrans.frx":23EF8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1170
      Top             =   2880
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
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
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5895
      Locked          =   -1  'True
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   810
      Width           =   1365
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5895
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1245
      Width           =   1365
   End
   Begin VB.TextBox xValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   5715
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2925
      Width           =   1545
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1395
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2490
      Width           =   5865
   End
   Begin MSDataListLib.DataCombo xNo1 
      Height          =   360
      Left            =   2025
      TabIndex        =   2
      Top             =   1665
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo xNo2 
      Height          =   360
      Left            =   2025
      TabIndex        =   3
      Top             =   2085
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1125
      Top             =   0
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
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
   Begin Threed.SSCommand CMD_PRINT 
      Height          =   465
      Left            =   1170
      TabIndex        =   35
      Top             =   3195
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   820
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "banktrans.frx":2625B
      Caption         =   "ÿ»«⁄… «·„” ‰œ  "
      Alignment       =   1
      PictureAlignment=   3
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "„”·”· :"
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
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   900
      Width           =   675
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï Õ”«» »‰ﬂ :"
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
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2205
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁÌ„… :"
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
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3015
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ :"
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
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1350
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰ :"
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
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2610
      Width           =   450
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ Õ”«» »‰ﬂ :"
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
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1755
      Width           =   1155
   End
End
Attribute VB_Name = "banktrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public bedit As Boolean, bClosed As Boolean
Dim formMode As Byte, cFileHeader As String
Dim con As New ADODB.Connection
Dim oSearch As New Search3
Dim CardTable As ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode And bedit)
CmdDel.Enabled = (nMode = LoadMode And xClosed.Value = 0)
cmdSave.Enabled = xClosed.Value = 0
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1
cmdNext.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount
cmdLast.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition < CardTable.RecordCount And CardTable.RecordCount > 2
cmdFirst.Enabled = (nMode = LoadMode) And CardTable.AbsolutePosition > 1 And CardTable.RecordCount > 2
xCode.Enabled = Not (nMode = LoadMode)
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
Dim cString As String
cString = "SELECT FILE5_51.CODE,FILE5_10.DESCA,FILE5_10_1.DESCA,FILE5_51.DESCA,CONVERT(VARCHAR(10),[DATE],111) , FILE5_51.[VALUE] " & _
           " FROM      FILE5_51 LEFT OUTER JOIN FILE5_10 ON FILE5_51.no1 = FILE5_10.CODE LEFT OUTER JOIN FILE5_10 AS FILE5_10_1 ON FILE5_51.no2 = FILE5_10_1.CODE "
Generalarray(1) = cString
Generalarray(2) = "ORDER BY FILE5_51.DATE"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰ √Ê «· «—ÌŒ"
listarray(0, 1) = "(%%FILE5_51.DESCA%% or ##date##)"

listarray(1, 0) = "„‰ Õ”«» »‰ﬂ"
listarray(1, 1) = "(%%FILE5_10.DESCA%%)"

listarray(2, 0) = "≈·Ì Õ”«» »‰ﬂ"
listarray(2, 1) = "(%%FILE5_10_1.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "„‰ Õ”«» »‰ﬂ "
GrdArray(1, 1) = 2500

GrdArray(2, 0) = "≈·Ì Õ”«» »‰ﬂ "
GrdArray(2, 1) = 2500

GrdArray(3, 0) = "«·»Ì«‰"
GrdArray(3, 1) = 4500

GrdArray(4, 0) = "«· «—ÌŒ"
GrdArray(4, 1) = 1400

GrdArray(5, 0) = "«·ﬁÌ„…"
GrdArray(5, 1) = 1400

searchArray = Array(Generalarray, listarray, GrdArray)
oSearch.Caption = "≈” ⁄·«„  ÕÊÌ·«  »Ì‰ «·»‰Êﬂ "
oSearch.Show 1
End Sub
Sub mydefine()


cmdSave.Enabled = True
CmdDel.Enabled = True

cmd_closed.BackColor = &H8080FF
cmd_closed.Caption = "› Õ „” ‰œ"
xClosed.Value = ssCBUnchecked

xUserName.Caption = ""
xUserName2.Caption = ""
XTIME1.Caption = ""
XTIME2.Caption = ""


xCode.text = RetZero(Val(Newflag("FILE5_51", "CODE", con)), 6)
xDescA.text = ""
xDate.text = ""
xNo1.BoundText = ""
xNo2.BoundText = ""
xValue.text = ""
Handlecontrols DefineMode
End Sub
Sub myProc()
CardTable.Find "CODE = " & MyParn(oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
oSearch.Hide
End Sub
Sub myload()
If CardTable!ISCLOSED Then
    cmd_closed.BackColor = &H8080FF
    cmd_closed.Caption = "› Õ „” ‰œ"
    xClosed.Value = 1
Else
    cmd_closed.BackColor = vbGreen
    cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
    xClosed.Value = 0
End If


xUserName.Caption = CardTable!UserName & ""
xUserName2.Caption = CardTable!UserName2 & ""
XTIME1.Caption = CardTable!Time & ""
XTIME2.Caption = CardTable!TIME2 & ""

xCode.text = CardTable!CODE
xDescA.text = TurnValue(CardTable!DESCA, Null, "")
xNo1.BoundText = TurnValue(CardTable!no1, Null, "")
xNo2.BoundText = TurnValue(CardTable!no2, Null, "")
xDate.text = Format(CardTable!Date, "DD-MM-YYYY")
xValue.text = Format(CardTable!Value, "Fixed")
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert(8, 1)
aInsert(0, 0) = "CODE"
aInsert(0, 1) = addstring(xCode.text)

aInsert(1, 0) = "DESCA"
aInsert(1, 1) = addstring(xDescA.text)

aInsert(2, 0) = "Date"
aInsert(2, 1) = addDate(xDate.text)

aInsert(3, 0) = "NO1"
aInsert(3, 1) = addstring(xNo1.BoundText)

aInsert(4, 0) = "NO2"
aInsert(4, 1) = addstring(xNo2.BoundText)

aInsert(5, 0) = "[VALUE]"
aInsert(5, 1) = Val(xValue.text)

aInsert(8, 0) = "[ISCLOSED]"
aInsert(8, 1) = (xClosed.Value)


On Error GoTo myerror
con.BeginTrans
If xCode.Enabled Then
    aInsert(6, 0) = "username"
    aInsert(6, 1) = addstring(cusername)
    aInsert(7, 0) = "time"
    aInsert(7, 1) = "getdate()"
    
    xCode.text = RetZero(Val(Newflag("FILE5_51", "CODE", con)), 6)
    aInsert(0, 1) = addstring(xCode.text)
    con.Execute CreateInsert(aInsert, "FILE5_51")
Else
    aInsert(6, 0) = "username2"
    aInsert(6, 1) = addstring(cusername)
    aInsert(7, 0) = "time2"
    aInsert(7, 1) = "getdate()"
    con.Execute CreateUpdate(aInsert, "FILE5_51", " where CODE = " & addstring(xCode.text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox " ”ÃÌ· „”·”· "
    Exit Function
End If
If xDate.text = "" Then
    MsgBox " ”ÃÌ·  «—ÌŒ "
    Exit Function
End If
If Trim(xNo1.BoundText) = "" Then
    MsgBox " ”ÃÌ· «·»‰ﬂ «·«Ê· ÷—Ê—Ì"
    Exit Function
End If
If Trim(xNo2.BoundText) = "" Then
    MsgBox " ”ÃÌ· «·»‰ﬂ «·À«‰Ì ÷—Ê—Ì"
    Exit Function
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
mydefine
On Error Resume Next
xCode.SetFocus
Err.Clear
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE5_51 Where Code = " & MyParn(xCode.text)
    con.CommitTrans
    openCardTable
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "Code < " & MyParn(xCode.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        mydefine
    End If
End If
Exit Sub
myerror:
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
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
If xCode.Enabled Then
    CmdAdd_Click
Else
    openCardTable
    myUndo
End If
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
ElseIf KeyAscii = 19 And cmdSave.Enabled Then
    cmdSave_Click
End If
End Sub
Private Sub Form_Load()
bedit = True
openCon con
Frame1.Visible = bopt2
cFileHeader = "file5_51"
data1.ConnectionString = strCon
data1.RecordSource = "Select * From file5_10"
DATA2.ConnectionString = strCon

DATA2.RecordSource = "Select * From file5_10"

Set xNo1.RowSource = data1
xNo1.ListField = "Desca"
xNo1.BoundColumn = "Code"

Set xNo2.RowSource = DATA2
xNo2.ListField = "Desca"
xNo2.BoundColumn = "Code"

openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub xCode_LostFocus()
xCode.text = RetZero(xCode.text, 6)
CardTable.Find "Code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
myLostFocus xCode
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
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

Private Sub xDesca_LostFocus()
myLostFocus xDescA
End Sub

Private Sub xValue_GotFocus()
myGotFocus xValue
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDescA
End Sub
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    mydefine
Else
    If Trim(xCode.text) <> "" Then
        CardTable.Find "CODE = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub openCardTable()
Dim cString As String
cString = "SELECT * FROM FILE5_51"
cString = cString & " order by code"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub xValue_LostFocus()
myLostFocus xValue
End Sub

Private Sub cmd_CLOSEDDATE_Click()
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "≈€·«ﬁ  „” ‰œ«  „‰  «—ÌŒ", ""), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  Õ Ï  «—ÌŒ", ""), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ ≈€·«ﬁ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update FILE5_51 set isclosed  = 1 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ ≈€·«ﬁ „” ‰œ«  «·› —…" & nRec
            CardTable.Requery
            CardTable.Find "code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
            If CardTable.EOF Then CardTable.MoveLast
            myload
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
            con.Execute " update FILE5_51 set isclosed = 0 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ › Õ „” ‰œ«  «·› —…" & nRec
            CardTable.Requery
            CardTable.Find "code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
            If CardTable.EOF Then CardTable.MoveLast
            myload
        End If
    End If
End Sub
Private Sub cmd_closed_Click()
If CardTable!ISCLOSED Then
    con.Execute " update FILE5_51 set isclosed = 0 where code = " & MyParn(xCode.text)
Else
    con.Execute " update FILE5_51 set isclosed  = 1 where code = " & MyParn(xCode.text)
End If
CardTable.Requery
CardTable.Find "code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub





Private Sub CMD_PRINT_Click()
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
temptable.AddNew
temptable!STR20 = Firsttitle
temptable!STR21 = "≈Ì’«· ≈” ·«„ ‰ﬁœÌ… :  "
temptable!STR7 = xCode.text
temptable!STR11 = " „‰ : " & xNo1.text
temptable!str12 = " ≈·Ï : " & xNo2.text

temptable!str2 = xDate.text
temptable!str6 = TurnValue(xDescA.text, "", Null)
temptable!VAL1 = Val(xValue.text)
temptable!str5 = MyOnly(Val(xValue.text))
temptable.Update
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\P_box.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing

End Sub

