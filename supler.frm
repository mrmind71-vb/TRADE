VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form suplerfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„Ê—œÌ‰"
   ClientHeight    =   7965
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   8385
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   7965
   ScaleWidth      =   8385
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   7245
      Width           =   8205
      Begin VB.TextBox XF_DATE 
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
         Left            =   180
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   180
         Width           =   1500
      End
      Begin VB.TextBox XF_BALANCE 
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
         Left            =   5085
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   180
         Width           =   1500
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ «Ê· : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   180
         Width           =   840
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ  «Ê· : "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   225
         Width           =   840
      End
   End
   Begin VB.Frame Frame7 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   6660
      Width           =   8205
      Begin MSDataListLib.DataCombo xcust 
         Height          =   315
         Left            =   3960
         TabIndex        =   49
         Top             =   180
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xbranch 
         Height          =   315
         Left            =   90
         TabIndex        =   51
         Top             =   180
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· ÊﬂÌ· : "
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
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   225
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄„Ì· :"
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
         Left            =   7245
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   225
         Width           =   690
      End
   End
   Begin VB.Frame Frame13 
      Height          =   600
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   5985
      Width           =   6135
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
         TabIndex        =   44
         Top             =   135
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   1980
         TabIndex        =   45
         Top             =   180
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
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
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   180
         Width           =   570
      End
   End
   Begin VB.CommandButton Cmd_gom 
      BackColor       =   &H00FFC0C0&
      Caption         =   " €Ì— «·„Ê—œ ≈·Ï „ﬂ »"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   675
      UseMaskColor    =   -1  'True
      Width           =   2085
   End
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   990
      Width           =   2085
      Begin VB.OptionButton xFact 
         Alignment       =   1  'Right Justify
         Caption         =   "„ﬂ » Ã„·…"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   675
         Width           =   1680
      End
      Begin VB.OptionButton xFact 
         Alignment       =   1  'Right Justify
         Caption         =   "„’‰⁄"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.Frame Frame5 
      Height          =   690
      Left            =   990
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   0
      Width           =   7260
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   6030
         Picture         =   "supler.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
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
         Picture         =   "supler.frx":27D3
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
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
         Picture         =   "supler.frx":4D7F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
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
         Picture         =   "supler.frx":7619
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
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
         Picture         =   "supler.frx":9A85
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
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
         Picture         =   "supler.frx":BFFE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.TextBox xRemark 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   90
      MaxLength       =   200
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   5175
      Width           =   6540
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   675
      Width           =   6090
      Begin VB.CheckBox xisonest 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ê—œ «„«‰«  "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   180
         Width           =   1365
      End
      Begin VB.TextBox xSubCode 
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
         Left            =   90
         MaxLength       =   3
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.TextBox xDescA 
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
         Left            =   90
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   4695
      End
      Begin VB.TextBox xCode 
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
         Left            =   3465
         MaxLength       =   3
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox xManager 
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
         Left            =   90
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   4695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„œÌ— „”ƒ· :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   990
         Width           =   870
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·«”„ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   450
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   495
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
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   4545
      Width           =   8160
      Begin VB.TextBox xdisc2 
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
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   180
         Width           =   690
      End
      Begin VB.TextBox xRate 
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
         Left            =   5580
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox xDisc 
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
         Left            =   2610
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ «„«‰«  ⁄·Ï «·»Ì⁄"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   225
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·„” Â·ﬂ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   225
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·Œ’„ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   180
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2400
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2115
      Width           =   8160
      Begin VB.CheckBox xisopen 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì— „ Êﬁ›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         Height          =   330
         Left            =   1845
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   180
         Width           =   330
      End
      Begin VB.TextBox xEMAIL 
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
         Left            =   135
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1935
         Width           =   6450
      End
      Begin VB.TextBox xFAx 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   135
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1535
         Width           =   6450
      End
      Begin VB.TextBox xPhone1 
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
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   135
         MaxLength       =   200
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1135
         Width           =   6450
      End
      Begin VB.TextBox xAddress 
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
         Height          =   555
         Left            =   90
         MaxLength       =   100
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   540
         Width           =   6495
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   315
         Left            =   2205
         TabIndex        =   4
         Top             =   180
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·≈Ì„Ì· :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1980
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê»Ì· :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1575
         Width           =   495
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«· ·Ì›Ê‰ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1170
         Width           =   615
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ã„Ê⁄… :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   750
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·⁄‰Ê«‰ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   540
         Width           =   570
      End
   End
   Begin VB.Frame Frame6 
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   5985
      Width           =   2040
      Begin VB.CommandButton cmdLast 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1530
         Picture         =   "supler.frx":E361
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "«·«ŒÌ—"
         Top             =   180
         Width           =   420
      End
      Begin VB.CommandButton cmdPrevious 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   540
         Picture         =   "supler.frx":105F8
         Style           =   1  'Graphical
         TabIndex        =   37
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdFirst 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         Picture         =   "supler.frx":12825
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "«·«Ê·"
         Top             =   180
         Width           =   465
      End
      Begin VB.CommandButton cmdNext 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1035
         Picture         =   "supler.frx":14B02
         Style           =   1  'Graphical
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   180
         Width           =   465
      End
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
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "‰”»… «·Œ’„ :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   0
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "„·«ÕŸ«  :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   5625
      Width           =   735
   End
End
Attribute VB_Name = "suplerfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public myFlag As Integer
Dim formMode As Byte, cTableName As String, cGroupname As String
Dim oSearch As New Search3
Dim CardTable As New ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub cmdPerviious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub

Private Sub CMD_FIXSEC_Click()
If xSection.BoundText = "" Then Exit Sub
If MsgBox(" ⁄œÌ· „ÊœÌ·«  «·›« Ê—… ··ﬁ”„", vbYesNo) = vbYes Then
    con.Execute " UPDATE FILE1_10 SET [SECTION] = " & Val(xSection.BoundText) & " WHERE CODE = " & MyParn(xCode.text)
    Inform " „  «· ⁄œÌ· »‰Ã«Õ"
End If

End Sub

Private Sub Cmd_gom_Click()
    If MsgBox(" €Ì— «·„Ê—œ ≈·Ï „ﬂ » Ã„·…", vbOKCancel) = vbOK Then
        con.Execute " update file4_10 set fact = 0 , supp = 1 where code = " & MyParn(xCode.text)
        con.Execute " update file1_10 set supp = " & addstring(xCode.text) & " , model = file1_10.modelno + file1_10.code where code = " & MyParn(xCode.text)
    End If
    MsgBox " „  €Ì— ﬂ· „ÊœÌ·«  «·„Ê—œ ≈·Ï „ÊœÌ·«  ·‰›” «·„’‰⁄ „‰ „ﬂ » Ã„·… - ÌÃ» „—«⁄«…  ⁄œÌ· «·„ÊœÌ·«  Õ”» „’«‰⁄Â«"
    openCardTable
    myUndo
End Sub
Private Sub Command3_Click()
Dim myPublic(6)
myPublic(0) = "FILE4_50"
myPublic(1) = "Code"
myPublic(2) = "Desca"
myPublic(3) = "ﬂÊœ «·„Ã„Ê⁄…"
myPublic(4) = "≈”„ «·„Ã„Ê⁄…"
myPublic(5) = "„Ã„Ê⁄«  «·„Ê—œÌ‰"
FlagFrm2.bedit = True
FlagFrm2.myPublic = myPublic
FlagFrm2.Show 1
data1.Refresh
End Sub
Private Sub Form_Load()
    openCon con

    Set data1.Recordset = myRecordSet("SELECT * FROM file4_50", con)
    Set xGroup.RowSource = data1
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"

    Set DATA2.Recordset = myRecordSet("SELECT * FROM file1_10sc order by desca ", con)
    Set xSection.RowSource = DATA2
    xSection.ListField = "DESCA"
    xSection.BoundColumn = "code"

    Set data3.Recordset = myRecordSet("SELECT * FROM file3_10 order by desca ", con)
    Set xcust.RowSource = data3
    xcust.ListField = "Desca"
    xcust.BoundColumn = "Code"

    Set data4.Recordset = myRecordSet("SELECT * FROM branch_fr order by desca ", con)
    Set XBRANCH.RowSource = data4
    XBRANCH.ListField = "Desca"
    XBRANCH.BoundColumn = "Code"

    openCardTable
    myUndo

cmdSave.Visible = (cBranch = "00")
CmdDel.Visible = (cBranch = "00")
CmdAdd.Visible = (cBranch = "00")
Cmd_gom.Visible = (cBranch = "00")
End Sub
Private Sub CmdAdd_Click()
myDefine
xCode.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.BeginTrans
    con.Execute "Delete  From file4_22  Where f_bal <> 0  and code = " & MyParn(xCode.text)
    con.Execute "Delete  From file4_10  Where code = " & MyParn(xCode.text)
    con.CommitTrans
    
    AddLod_Data cusername, 2, " Õ–›  „Ê—œÌ‰ ", con, xCode.text, , , xDesca.text
    
    openCardTable
    CmdAdd_Click
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
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub

Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
AddLod_Data cusername, 1, " Õ÷Ÿ „Ê—œÌ‰ ", con, xCode.text, , , xDesca.text

openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
SuppLookupAll Me, oSearch
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
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xCode.Enabled = Not (nMode = LoadMode)
xsubcode.Enabled = (nMode = DefineMode)
xFact(0).Enabled = (nMode = DefineMode)
xFact(1).Enabled = (nMode = DefineMode)
End Sub
Sub myDefine()


xCode.text = RetZero(Newflag("file4_10", "code", con), 3)
xDesca.text = ""
xisopen.Value = 1
XISONEST.Value = 1


xF_date.text = ""
XF_BALANCE.text = ""

xManager.text = ""
XEMAIL.text = ""
xGroup.BoundText = ""
xcust.BoundText = ""
XBRANCH.BoundText = ""
xAddress.text = ""
xPhone1.text = ""
xFAx.text = ""
XDISC.text = ""
xdisc2.text = ""
xRate.text = ""
xFact(0).Value = 1
xRemark.text = ""
xsubcode.text = ""
Handlecontrols DefineMode
End Sub
Sub myload()
Cmd_gom.Visible = xFact(0).Value = True
xisopen.Value = IIf(CardTable!IsOpen, 1, 0)
XISONEST.Value = IIf(CardTable!ISONEST, 1, 0)

xF_date.text = Format(CardTable!F_DATE, "DD-MM-YYYY")
XF_BALANCE.text = Format(CardTable!F_Balance, "#0.00")

xCode.text = CardTable!CODE & ""
xDesca.text = CardTable!Desca
xManager.text = CardTable!Manager & ""
xAddress.text = CardTable!Address & ""
XEMAIL.text = CardTable!email & ""
xPhone1.text = CardTable!PHONE1 & ""
xFAx.text = CardTable!Fax & ""
XDISC.text = Format(CardTable!disc, "##0.00")
xdisc2.text = Format(CardTable!disc2, "##0.00")
xRate.text = Format(CardTable!Rate, "##0.00")
xFact(0).Value = IIf(CardTable!Fact, 1, 0)
xFact(1).Value = IIf(CardTable!SUPP, 1, 0)
xGroup.BoundText = CardTable!Group & ""
xcust.BoundText = CardTable!CUST & ""
XBRANCH.BoundText = CardTable!branch & ""
xRemark.text = CardTable!remark & ""
xsubcode.text = CardTable!SUBCODE & ""
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert(20, 1)
aInsert(0, 0) = "Code"
aInsert(0, 1) = addstring(xCode.text)

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(xDesca.text)

aInsert(2, 0) = "Manager"
aInsert(2, 1) = addstring(xManager.text)

aInsert(3, 0) = "Address"
aInsert(3, 1) = addstring(xAddress.text)

aInsert(4, 0) = "phone1"
aInsert(4, 1) = addstring(xPhone1.text)

aInsert(5, 0) = "fax"
aInsert(5, 1) = addstring(xFAx.text)

aInsert(6, 0) = "[Group]"
aInsert(6, 1) = addstring(xGroup.BoundText)

aInsert(7, 0) = "email"
aInsert(7, 1) = addstring(XEMAIL.text)

aInsert(8, 0) = "disc"
aInsert(8, 1) = Val(XDISC.text)

aInsert(9, 0) = "RATE"
aInsert(9, 1) = Val(xRate.text)

aInsert(10, 0) = "FACT"
aInsert(10, 1) = IIf(xFact(0).Value, 1, 0)

aInsert(11, 0) = "SUPP"
aInsert(11, 1) = IIf(xFact(1).Value, 1, 0)

aInsert(12, 0) = "remark"
aInsert(12, 1) = addstring(xRemark.text)

If xFact(0).Value Then
    aInsert(13, 0) = "SUBCODE"
    aInsert(13, 1) = addstring(xCode.text)
Else
    aInsert(13, 0) = "SUBCODE"
    aInsert(13, 1) = "NULL"
End If

aInsert(14, 0) = "ISOPEN"
aInsert(14, 1) = xisopen.Value

aInsert(15, 0) = "[cust]"
aInsert(15, 1) = addstring(xcust.BoundText)

aInsert(16, 0) = "branch"
aInsert(16, 1) = addstring(XBRANCH.BoundText)

aInsert(17, 0) = "disc2"
aInsert(17, 1) = Val(xdisc2.text)

aInsert(18, 0) = "isonest"
aInsert(18, 1) = XISONEST.Value

aInsert(19, 0) = "f_Balance"
aInsert(19, 1) = Val(XF_BALANCE.text)

aInsert(20, 0) = "F_DATE"
aInsert(20, 1) = addDate(xF_date.text)


con.BeginTrans
If xCode.Enabled Then
    xCode.text = RetZero(Newflag("FILE4_10", "CODE", con), 3)
    aInsert(0, 1) = addstring(xCode.text)
    cString = CreateInsert(aInsert, "file4_10")
    con.Execute cString
Else
    cString = CreateUpdate(aInsert, "file4_10", " WHERE CODE = " & MyParn(xCode.text))
    con.Execute cString
End If
If xFact(0).Value Then
    If GetDesca("SELECT CODE FROM FACT WHERE CODE = " & MyParn(xCode.text), con) = "" Then
        con.Execute "INSERT INTO FACT (CODE , DESCA , [RATE] )" & _
                    "VALUES( " & _
                    addstring(xCode.text) & "," & _
                    addstring(xDesca.text) & "," & _
                    Val(xRate.text) & _
                    ")"
    Else
        con.Execute "UPDATE FACT " & _
                    " SET DESCA = " & addstring(xDesca.text) & _
                    ",[RATE] = " & Val(xRate.text) & _
                    " WHERE code = " & MyParn(xCode.text)
    End If
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
If Err.Number <> 0 Then MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
xCode.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
oSearch.Hide
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Err.Clear
End Sub

Private Sub xCode_Change()
xsubcode.text = xCode.text
End Sub

Private Sub xCode_LostFocus()
If xCode.text = "" Then Exit Sub
xCode.text = xCode.text
CardTable.Find "CODE = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

'If Len(Trim(xsubcode.Text)) <> 3 Then
'    MsgBox "—«Ã⁄ «·ﬂÊœ «·„Œ ’—  "
'    Exit Function
'End If

If xDesca.text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

MYVALID = True
End Function
Private Sub openCardTable()
Dim cString As String
cString = "SELECT FILE4_10.* From file4_10"
cString = cString & " ORDER BY CODE"
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myUndo()
If (CardTable.BOF And CardTable.EOF) Then
    myDefine
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
