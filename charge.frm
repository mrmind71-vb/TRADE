VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form chargefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁœÌ…"
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15225
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
   ScaleHeight     =   9225
   ScaleWidth      =   15225
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   870
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   135
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   30
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
         TabIndex        =   29
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
         TabIndex        =   28
         Top             =   495
         Width           =   2220
      End
   End
   Begin VB.Frame FRM_CLOSED 
      Height          =   1815
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   135
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
         TabIndex        =   23
         Top             =   135
         Visible         =   0   'False
         Width           =   1410
      End
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         TabIndex        =   24
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
         Picture         =   "charge.frx":0000
         Alignment       =   4
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   1035
         TabIndex        =   25
         Top             =   135
         Visible         =   0   'False
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
         Picture         =   "charge.frx":25CC
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   90
         TabIndex        =   26
         Top             =   135
         Visible         =   0   'False
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
         Picture         =   "charge.frx":4C44
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   6750
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   945
      Width           =   1275
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge.frx":7285
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   630
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
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge.frx":97FE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Õ›Ÿ"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   8145
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   15
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
         Picture         =   "charge.frx":BB61
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge.frx":DD31
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   16
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
         Picture         =   "charge.frx":FE79
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge.frx":12041
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   17
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
         Picture         =   "charge.frx":14190
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge.frx":16370
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   18
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
         Picture         =   "charge.frx":184CB
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge.frx":1A687
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   9675
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   5460
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4095
         Picture         =   "charge.frx":1C7D6
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   2790
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge.frx":1EFA9
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge.frx":21555
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge.frx":23DEF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   8925
      Width           =   15225
      _ExtentX        =   26855
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
            TextSave        =   "04:32 „"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   8055
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   675
      Width           =   7125
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
         Left            =   4500
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1320
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
         Left            =   270
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo XSUPP 
         Height          =   315
         Left            =   2340
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   585
         Visible         =   0   'False
         Width           =   3480
         _ExtentX        =   6138
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
      Begin MSDataListLib.DataCombo XMOSM 
         Height          =   315
         Left            =   2340
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   3480
         _ExtentX        =   6138
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê”„ : "
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
         Left            =   5940
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   990
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Left            =   5970
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   630
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   1650
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   285
         Width           =   600
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
         Left            =   5970
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   255
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   6480
      Top             =   1035
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Left            =   6480
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Height          =   6000
      Left            =   45
      TabIndex        =   0
      Top             =   2070
      Width           =   15090
      _cx             =   26617
      _cy             =   10583
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
   Begin Threed.SSCommand CMD_PRINT 
      Height          =   465
      Left            =   7605
      TabIndex        =   34
      Top             =   180
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
      Picture         =   "charge.frx":23F39
      Caption         =   "ÿ»«⁄… «·„” ‰œ  "
      Alignment       =   1
      PictureAlignment=   3
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
   Begin VB.Label xtotal 
      Alignment       =   2  'Center
      Caption         =   "0.00"
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
      Left            =   3420
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   8235
      Width           =   2220
   End
   Begin VB.Label xbranch 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   330
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   1260
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "chargefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte, bedit As Boolean
Dim cFilter2 As String, oSearchSup As New Search3
Dim cStrBox As String
Dim CardTable As ADODB.Recordset, GrdTable As New ADODB.Recordset
Dim cFile As String, cFileHeader As String, sName As String
Dim docMoveType As String
Dim DocTitle As String
Dim DocClient As String, cGroup As String
Dim dLastdate As String, defBox As String
Dim DocField As String, dDateLast As String
Dim formMode
Dim con As New ADODB.Connection
Dim lCellButton As Boolean
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(8, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xdoc_no.text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.text)

aInsert(2, 0) = "branch"
aInsert(2, 1) = addstring(xbranch.Caption)

aInsert(3, 0) = "isnew"
aInsert(3, 1) = 1

aInsert(4, 0) = "isclosed"
aInsert(4, 1) = 0

aInsert(5, 0) = "SUPP"
aInsert(5, 1) = addstring(xSupp.BoundText)

aInsert(8, 0) = "MOSM"
aInsert(8, 1) = addstring(xMosm.BoundText)


On Error GoTo myerror
con.BeginTrans
If xdoc_no.Enabled Then
    aInsert(6, 0) = "username"
    aInsert(6, 1) = addstring(cusername)
    aInsert(7, 0) = "time"
    aInsert(7, 1) = "getdate()"
    
    If cBranch > "60" Then
        xdoc_no.text = NewflagBranch(cFileHeader, "doc_no", "00", con)
    Else
        xdoc_no.text = NewflagBranch(cFileHeader, "doc_no", cBranch, con)
    End If
    aInsert(0, 1) = addstring(xdoc_no.text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
    aInsert(6, 0) = "username2"
    aInsert(6, 1) = addstring(cusername)
    aInsert(7, 0) = "time2"
    aInsert(7, 1) = "getdate()"
    con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xdoc_no.text))
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
Dim aInsert(7, 1)
With grid1
    For i = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xdoc_no.text)
        
        aInsert(1, 0) = "Box"
        aInsert(1, 1) = addstring(.TextMatrix(i, 0))
        
        aInsert(2, 0) = "Charge"
        aInsert(2, 1) = addstring(grid1.TextMatrix(i, 1))
        
        aInsert(3, 0) = "Desca"
        aInsert(3, 1) = addstring(grid1.TextMatrix(i, 3))
        
        aInsert(4, 0) = "[value]"
        aInsert(4, 1) = Val(grid1.TextMatrix(i, 4))

        aInsert(5, 0) = "[ser_no]"
        aInsert(5, 1) = addvalue(grid1.TextMatrix(i, 5))


        aInsert(6, 0) = "BRANCH"
        If cBranch = "00" Then
            aInsert(6, 1) = addstring(grid1.TextMatrix(i, 6))
        Else
            aInsert(6, 1) = addstring(cBranch)
        End If

        aInsert(7, 0) = "[Row]"
        aInsert(7, 1) = i

        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, cFile)
        Else
            con.Execute CreateUpdate(aInsert, cFile, " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 1 Then
        grid1.TextMatrix(grid1.Row, 1) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
        GrdDesc grid1.Row
        If grid1.Row = grid1.Rows - 1 And validRow(grid1.Row) Then
            myaddItem
        End If
        Unload Search3
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xdoc_no.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
    myUndo
ElseIf ActiveControl.Name = xSupp.Name Then
    xSupp.BoundText = oSearchSup.grid1.TextMatrix(oSearchSup.grid1.Row, 0)
    Unload oSearchSup
End If
End Sub

Private Sub CMD_PRINT_Click()
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
'On Error GoTo myerror
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
For i = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!STR21 = "≈Ì’«· ’—› ‰ﬁœÌ… ··„’«—Ì› : " & xdoc_no.text
    temptable!STR7 = xdoc_no.text
    temptable!str1 = cComp_Name
    temptable!val2 = grid1.TextMatrix(i, 4)
    temptable!str3 = TurnValue(grid1.TextMatrix(i, 3), "", Null)
    temptable!str6 = TurnValue(grid1.TextMatrix(i, 2), "", Null)
    temptable!str8 = TurnValue(GetDesca("select desca from file0_50 where code = " & MyParn(grid1.TextMatrix(1, 0)), con), "", Null)
    temptable!str2 = xDate.text
    temptable!VAL1 = Val(xTotal.Caption)
    temptable!str5 = MyOnly(xTotal.Caption)
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans

Main.REPORT1.ReportFileName = App.Path & "\Reports\P_CHARGE.rpt"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\Temp.MDB"
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear

End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xdoc_no.text)
    If cBranch = ("00" Or cBranch > "60") Or (myPublic = 1) Then
        con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xdoc_no.text)
    Else
        con.Execute " UPDATE " & cFileHeader & " SET ISNEW = 1 where Doc_No = " & MyParn(xdoc_no.text)
    End If
    con.CommitTrans
    openCardTable
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
Dim listarray(1, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
cString = "SELECT " & cFileHeader & ".Doc_No, Convert(Varchar," & cFileHeader & ".Date,111),Min(" & DocClient & ".Desca)  ,  BRANCH.DESCA " & _
          " FROM (" & cFileHeader & " inner join " & cFile & " on " & cFileHeader & ".doc_no = " & cFile & ".Doc_NO) Inner Join " & DocClient & " on " & cFile & ".Charge = " & DocClient & ".Code  inner join branch on BRANCH.CODE = " & cFileHeader & ".BRANCH WHERE " & cFileHeader & ".DOC_NO IS NOT NULL "

If lIsBranchStore Then
    cString = cString & " AND " & cFileHeader & ".BRANCH = " & MyParn(cBranch)
End If
Generalarray(1) = cString
Generalarray(2) = " group by " & cFileHeader & ".Doc_No," & cFileHeader & ".Date ,  BRANCH.DESCA order by " & cFileHeader & ".Doc_No," & cFileHeader & ".Date "
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·«”„- «—ÌŒ «·„” ‰œ"
listarray(0, 1) = "(%%" & DocClient & ".Desca%% or " & _
                  " ##" & cFileHeader & ".Date##)"

listarray(1, 0) = "«·›—⁄"
listarray(1, 1) = "%%BRANCH.DESCA%%"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·≈”„"
GrdArray(2, 1) = 3000

GrdArray(3, 0) = "«·›—⁄"
GrdArray(3, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
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
        openCardTable
        myUndo
End Sub
Private Sub CmdUndo_Click()
    openCardTable
    myUndo
End Sub

Private Sub Command1_Click()
Dim loctable As New ADODB.Recordset
loctable.Open "SELECT * FROM " & cFileHeader & " ORDER BY DATE", con, adOpenStatic, adLockReadOnly
Dim cString As String, i As Long
con.BeginTrans
Do
    i = i + 1
    Me.Caption = i
    cString = "UPDATE " & cFileHeader & " SET " & cFileHeader & ".DOC_NO = " & MyParn(RetZero(i, 6)) & " FROM " & cFileHeader
    cString = cString & turn(cString) & cFileHeader & ".DOC_NO = " & MyParn(loctable!doc_no & "")
    con.Execute cString
    loctable.MoveNext
Loop Until loctable.EOF
con.CommitTrans
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()

FRM_CLOSED.Visible = lSupperVisor Or (bopt2 And (cBranch = "00" Or cBranch > "60")) Or bOpt10 Or bOpt11
CmdDelInv.Visible = (cBranch = "00" Or cBranch > "60")
openCon con

Select Case myPublic
    Case 1 '„’«—Ì›
'        con.Execute " UPDATE FILE8_50H SET BRANCH = " & addstring(cBranch) & " WHERE BRANCH IS NULL "
        sName = "«·„’—Ê›"
        cFile = "File8_50"
        cFileHeader = "FILE8_50H"
        DocClient = "FILE8_51"
        If cBranch = "00" Or cBranch > "60" Then
            xSupp.Visible = True
            Label4.Visible = True
            Label2.Visible = True
            xMosm.Visible = True
        End If
    Case 4 '„’«—Ì›
        sName = "„’—Ê›  ÊﬂÌ·«  »«·‰”»…"
        cFile = "File8_55"
        cFileHeader = "FILE8_55H"
        DocClient = "FILE8_51"
        xSupp.Visible = True
        Label4.Visible = True
        Label2.Visible = True
        xMosm.Visible = True
    Case 2 '«·«Ì—«œ
        sName = "«·«Ì—«œ"
        cFile = "File8_60"
        cFileHeader = "FILE8_60H"
        DocClient = "FILE8_61"
        xDate.Enabled = lSupperVisor Or (cBranch > "60" Or cBranch = "00")
    Case 3 '«·«Ì—«œ
        sName = "‰ﬁœÌ… «·›—Ê⁄"
        cFile = "File8_80"
        cFileHeader = "FILE8_80H"
        DocClient = "FILE8_81"
        CmdDelInv.Visible = False
        cmdSave.Visible = False
End Select
Me.Caption = DocTitle

xSupp.Visible = Not lIsBranchStore
xMosm.Visible = Not lIsBranchStore
Label4.Visible = Not lIsBranchStore
Label2.Visible = Not lIsBranchStore

Set data1.Recordset = myRecordSet("SELECT * FROM FILE4_10 ORDER BY DESCA ", con)
Set xSupp.RowSource = data1
xSupp.ListField = "Desca"
xSupp.BoundColumn = "Code"

Set data3.Recordset = myRecordSet("SELECT * FROM MOSM ORDER BY DATE DESC ", con)
Set xMosm.RowSource = data3
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"


If cBranch = "00" Or cBranch > "60" Then
    cStrBox = StrList("select code , desca from file0_50 order by code ", con)
Else
    If lIsBranchStore Then
        cStrBox = StrList("select code , desca from file0_50 where BRANCH = " & MyParn(cBranch) & " AND type <> '1' order by code ", con)
    Else
        cStrBox = StrList("select code , desca from file0_50 where type <> '1' order by code ", con)
    End If
End If
Set grid1.DataSource = DATA11

'defBox = RetDefBox
Fixgrd

openCardTable
Dim cString As String, sBegin As String
cString = "SELECT TOP 1 DOC_NO FROM " & cFileHeader
If lIsBranchStore Then cString = cString & " WHERE BRANCH = " & MyParn(cBranch)
cString = cString & turn(cString) & "DATE = " & DateSq(Format(dSalesDate, "DD-MM-YYYY"))

cString = cString & " ORDER BY DATE DESC,DOC_NO DESC"
sBegin = GetDesca(cString, con)

If sBegin = "" Then
    myDefine
Else
    xdoc_no.text = sBegin
    myUndo
End If
grid1.Row = grid1.Rows - 1
grid1.ShowCell grid1.Rows - 1, 1
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
GrdTable.Close
Set CardTable = Nothing
Set GrdTable = Nothing
closeCon con
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 2 Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.SetFocus
    grid1.Select 1, 0
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If Not cmdSave.Enabled Then Exit Sub
If KeyCode = 112 And grid1.Col = 1 Then grdLookup
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 And bedit Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from " & cFile & " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        grid1.RemoveItem grid1.Row
        grid1_EnterCell
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If Not cmdSave.Enabled Then Exit Sub
If KeyCode = 112 And Col = 2 Then grdLookup
End Sub

Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub

Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 0 Then CellPos KeyCode, Row, Col
End Sub
Private Function MYVALID() As Boolean
If Trim(xdoc_no.text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If cBranch <> "00" And cBranch < "60" And Not lShowBranch Then
    If DateValue(xDate.text) < dMaxDate And Not lSupperVisor Then
        MsgBox " «—ÌŒ €Ì— ’«·Õ "
        Exit Function
    End If
End If


If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If

If cBranch = "00" Or cBranch > "60" Then
With grid1
    For i = 1 To .Rows - 2
        If .TextMatrix(i, 0) = "" And xSupp.BoundText = "" Then
            .Select i, 0, i, grid1.Cols - 1
            MsgBox "«·Œ“‰… €Ì— „Õœœ…"
            Exit Function
        End If
        
        If .TextMatrix(i, 0) <> "" And xSupp.BoundText <> "" Then
            .Select i, 0, i, grid1.Cols - 1
            MsgBox " „ «Œ Ì«— Œ“‰… Ê „Ê—œ "
            Exit Function
        End If
        
        If .TextMatrix(i, 1) = "" Then
            .Select i, 0, i, grid1.Cols - 1
            MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
            Exit Function
        End If
    
        If .TextMatrix(i, 6) = "" And myPublic = 4 Then
            .Select i, 0, i, grid1.Cols - 1
            MsgBox "«·›—⁄ €Ì— „ÊÃÊœ"
            Exit Function
        End If
    
    Next
End With
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
    xbranch.Caption = CardTable!branch & ""
    xusername.Caption = CardTable!UserName & ""
    xUserName2.Caption = CardTable!UserName2 & ""
    XTIME1.Caption = CardTable!Time & ""
    XTIME2.Caption = CardTable!TIME2 & ""
    xbranch.Caption = CardTable!branch & ""
    xSupp.BoundText = CardTable!SUPP & ""
    xMosm.BoundText = CardTable!MOSM & ""
xdoc_no.text = CardTable!doc_no
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")

Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub myloadgrd()
With grid1
    Dim cString As String
    
    cString = "SELECT " & cFile & ".[BOX], " & cFile & ".Charge," & DocClient & ".DESCA," & cFile & ".desca,[VALUE],ser_no , BRANCH , [ID] " & _
               " FROM " & cFile & " LEFT JOIN " & DocClient & " ON " & cFile & ".Charge = " & DocClient & ".CODE "
    cString = cString & turn(cString) & "Doc_no = " & MyParn(xdoc_no.text)
    Set DATA11.Recordset = myRecordSet(cString, con)
    myaddItem
End With
CalcTotals
Fixgrd
End Sub
Private Sub myDefine()
    xusername.Caption = ""
    xUserName2.Caption = ""
    XTIME1.Caption = ""
    XTIME2.Caption = ""
    xMosm.BoundText = cPMosm
    cmd_closed.BackColor = &H8080FF
    cmd_closed.Caption = "› Õ „” ‰œ"
    xClosed.Visible = False
    xClosed.Value = ssCBUnchecked
    xbranch.Caption = cBranch
    xdoc_no.text = NewflagBranch(cFileHeader, "doc_no", cBranch, con)

    xSupp.BoundText = ""
If (cBranch = "00") Then
    xDate.text = ""
Else
    xDate.text = Format(dSalesDate, "DD-MM-YYYY")
End If

grid1.Rows = 1
myaddItem
Handlecontrols DefineMode
CalcTotals
CmdDelInv.Visible = True
cmdSave.Visible = True
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode And bedit) Or bopt2
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
CmdDelInv.Enabled = (nMode = LoadMode) And (xClosed.Value = 0)
cmdPrevious.Enabled = (nMode = LoadMode)
xdoc_no.Enabled = (nMode = DefineMode) ' And (DateValue(xdate.Text) = DateValue(dSalesDate)) Or bopt2

'CmdSave.Enabled = (nMode = LoadMode And bEdit And DateValue(xdate.Text) = DateValue(dSalesDate)) Or bopt2
cmdSave.Enabled = (xClosed.Value = 0)

'CmdDelInv.Visible = (cBranch = xbranch.Caption)
'cmdSave.Visible = (cBranch = xbranch.Caption)

End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
    If Col = 1 And Trim(.EditText) <> "" And IsNumeric(.EditText) Then .EditText = RetZero(.EditText, 3)
End With
End Sub

Private Sub xDoc_No_LostFocus()
If Trim(xdoc_no.text) = "" Then Exit Sub
xdoc_no.text = RetZero(xdoc_no.text)
CardTable.Find "Doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function StrBox()
Dim boxtable As ADODB.Recordset
Set boxtable = New ADODB.Recordset

boxtable.Open "SELECT * FROM file0_50 WHERE ISSTOP = 0  ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
If Not (boxtable.EOF And boxtable.BOF) Then
    boxtable.MoveFirst
    StrBox = "#  " & ";       "
    Do Until boxtable.EOF
        StrBox = StrBox & "|#" & boxtable!CODE & ";" & boxtable!DESCA
        boxtable.MoveNext
    Loop
End If
End Function
Private Sub grdLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

If myPublic <> 2 Then
    Generalarray(1) = "Select code ,DescA From " & DocClient
Else
    If cBranch = "00" Or cBranch >= "60" Then
        Generalarray(1) = "Select code ,DescA From " & DocClient
    Else
'       Generalarray(1) = "Select code , DescA From FILE8_61 WHERE MAINGROUP = '00' "
        Generalarray(1) = "Select code , DescA From FILE8_61  "
    End If
End If
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Private Function CalcTotals()
Dim nTotal As Double
With grid1
For i = 1 To grid1.Rows - 2
    nTotal = nTotal + Round(Val(grid1.TextMatrix(i, 4)), 2)
Next
xTotal.Caption = Format(nTotal, "#0.00")
StatusBar1.Panels(1).text = "«·«Ã„«·Ì : " & Format(nTotal, "Fixed")
End With
End Function
Private Sub GrdDesc(nRow)
    grid1.TextMatrix(nRow, 2) = GetDesca("Select Desca From " & DocClient & " WHERE code = " & MyParn(grid1.TextMatrix(nRow, 1)), con) & ""
End Sub
Private Function RetDefBox() As String
Dim loctable As New ADODB.Recordset
loctable.Open "file0_50", con, adOpenStatic, adLockReadOnly, adCmdTable
If loctable.EOF And loctable.BOF Then Exit Function
loctable.MoveLast
If loctable.RecordCount = 1 Then
    loctable.MoveFirst
    RetDefBox = Trim(loctable!CODE & "")
End If
End Function

Private Sub xDoc_No_Validate(Cancel As Boolean)
If xdoc_no.text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With grid1
    .Cols = 8
    If myPublic <> 2 Then
'                           0           1           2                   3       4           5           6               7
        .FormatString = "Œ“‰…|" & "«·„’—Ê›|" & "Ê’› «·„’—Ê›|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "„” ‰œ|" & "«·›—⁄ "
    Else
        .FormatString = "Œ“‰…|" & "«·«Ì—«œ|" & "Ê’› «·«Ì—«œ|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "„” ‰œ|" & "«·›—⁄"
    End If
    .ColWidth(0) = 2000
    
    .ColWidth(1) = 800
    .ColWidth(2) = 2800
    .ColWidth(3) = 4500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1200
    .ColWidth(6) = 2000
    
'    .ColHidden(1) = (cBranch <> "00")
'    .ColHidden(2) = (cBranch <> "00")
'    .ColHidden(0) = (cBranch <> "00")
'    If cBranch <> "00" Then .ColWidth(3) = 8000

    .ColHidden(6) = (cBranch <> "00" And cBranch < "60")
    .ColDataType(4) = flexDTDouble
    If myPublic = 4 Then
        .ColComboList(6) = StrList("SELECT CODE , DESCA FROM BRANCH WHERE ISBRANCH2 =  1", con)
    Else
        If Not lIsBranchStore Then
            .ColComboList(6) = StrList("SELECT CODE , DESCA FROM BRANCH ORDER BY CODE ", con)
        Else
            .ColComboList(6) = StrList("SELECT CODE , DESCA FROM BRANCH WHERE CODE = " & MyParn(cBranch) & " ORDER BY CODE ", con)
        End If
    End If
    
    .ColFormat(4) = "#0.00"
    .ColHidden(.Cols - 1) = True
    '.ColHidden(0) = True
    For i = 1 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColComboList(0) = cStrBox
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
        CalcTotals
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 And grid1.TextMatrix(.Row, .Cols - 1) = "" Then
    .RemoveItem .Row
    CalcTotals
End If
End With
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Not (Val(.TextMatrix(nRow, 4)) <> 0) Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Col = 1 Then GrdDesc Row
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
    myaddItem
    CalcTotals
End If
End With
End Sub
Private Sub myUndo()
'On Error GoTo myerror
If CardTable.BOF And CardTable.EOF Then
    myDefine
Else
    If xdoc_no.text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT * FROM " & cFileHeader
If lIsBranchStore Then cString = cString & " WHERE BRANCH = " & MyParn(cBranch)
cString = cString & " ORDER BY DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myaddItem()
With grid1
.AddItem ""
If cBranch <> "00" And cBranch < "60" Then

End If
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    If Col = 1 Then
        grid1.Select Row, Col + 2
    Else
        grid1.Select Row, Col + 1
    End If
ElseIf Row < grid1.Rows - 1 Then
    grid1.Row = Row + 1
    grid1.Select Row + 1, 1
    If Not grid1.RowIsVisible(Row + 1) Then grid1.ShowCell Row + 1, 1
End If
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
Else
    con.Execute " update " & cFileHeader & " set isclosed = 1 where doc_no = " & MyParn(xdoc_no.text)
End If
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload

End Sub
Private Sub xSupp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    If KeyCode = 112 Then SuppLookupAll Me, oSearchSup
End If
End Sub
