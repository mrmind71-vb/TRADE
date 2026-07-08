VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Cashfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁœÌ…"
   ClientHeight    =   8235
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   15375
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame9 
      Height          =   870
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   45
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
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   45
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
         Picture         =   "cash.frx":0000
         Alignment       =   4
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   1035
         TabIndex        =   25
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
         Picture         =   "cash.frx":25CC
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
         Picture         =   "cash.frx":4C44
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   9135
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
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
         Picture         =   "cash.frx":7285
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "cash.frx":9455
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
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
         Picture         =   "cash.frx":B59D
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "cash.frx":D765
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
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
         Picture         =   "cash.frx":F8B4
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "cash.frx":11A94
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   19
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
         Picture         =   "cash.frx":13BEF
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "cash.frx":15DAB
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   9765
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   5415
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cash.frx":17EFA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1410
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cash.frx":18044
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   2700
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cash.frx":1A8DE
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4050
         Picture         =   "cash.frx":1CE8A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   150
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   6255
      Top             =   855
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   5
      Top             =   7800
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "02:07 „"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial (Arabic)"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   1365
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   585
      Width           =   4740
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
         Left            =   2295
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
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
         Left            =   2295
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xBox2 
         Height          =   315
         Left            =   135
         TabIndex        =   20
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
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰… :"
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
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   990
         Visible         =   0   'False
         Width           =   585
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   600
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   255
         Width           =   930
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7080
      Left            =   180
      TabIndex        =   6
      Top             =   1980
      Width           =   15000
      _cx             =   26458
      _cy             =   12488
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   0
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
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   8955
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   900
      Width           =   1455
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cash.frx":1F65D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cash.frx":1FA9F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   6255
      Top             =   990
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
      Height          =   555
      Left            =   7380
      TabIndex        =   33
      Top             =   90
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   979
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
      Picture         =   "cash.frx":22018
      Caption         =   "ÿ»«⁄… «·„” ‰œ  "
      Alignment       =   1
      PictureAlignment=   3
   End
End
Attribute VB_Name = "Cashfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte
Dim cFilter As String
Public bedit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim cFile As String, cFileHeader As String, sName As String
Dim cStrBox As String
Dim cStrMosm As String
Dim cFileMove As String
Dim DocTitle As String
Dim DocClient As String, CGROUP As String
Dim dLastdate As String, defBox As String, defMosm As String
Dim dDateLast As String
Dim formMode
Dim lCellButton As Boolean
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(3, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.text)

On Error GoTo myError
con.BeginTrans
If xDoc_No.Enabled Then
    
    aInsert(2, 0) = "username"
    aInsert(2, 1) = addstring(cusername)
    aInsert(3, 0) = "time"
    aInsert(3, 1) = "getdate()"
    
    xDoc_No.text = RetZero(Val(Newflag(cFileHeader, "doc_no", con)))
    aInsert(0, 1) = addstring(xDoc_No.text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
    
    aInsert(2, 0) = "username2"
    aInsert(2, 1) = addstring(cusername)
    aInsert(3, 0) = "time2"
    aInsert(3, 1) = "getdate()"
    con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xDoc_No.text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd()
Dim aInsert(8, 1)
With GRID1
    For i = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.text)
                
        aInsert(1, 0) = "Box"
        aInsert(1, 1) = addstring(.TextMatrix(i, 0))
        
        aInsert(2, 0) = "code"
        aInsert(2, 1) = addstring(GRID1.TextMatrix(i, 1))
                
        aInsert(3, 0) = "Desca"
        aInsert(3, 1) = addstring(GRID1.TextMatrix(i, 3))
        
        aInsert(4, 0) = "[value]"
        aInsert(4, 1) = Val(GRID1.TextMatrix(i, 4))

        aInsert(5, 0) = "MOSM"
        aInsert(5, 1) = addstring(GRID1.TextMatrix(i, 5))
        
        aInsert(6, 0) = "STORE"
        aInsert(6, 1) = addstring(GRID1.TextMatrix(i, 6))
        
        
        aInsert(7, 0) = "SER_NO"
        aInsert(7, 1) = addstring(GRID1.TextMatrix(i, 8))
        
        aInsert(8, 0) = "row"
        aInsert(8, 1) = i
        
        If GRID1.TextMatrix(i, GRID1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, cFile)
        Else
            con.Execute CreateUpdate(aInsert, cFile, " where ID = " & GRID1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = GRID1.Name Then
    If GRID1.Col = 1 Then
        GRID1.TextMatrix(GRID1.Row, 1) = Search3.GRID1.TextMatrix(Search3.GRID1.Row, 0)
        GrdDesc GRID1.Row
        If GRID1.Row = GRID1.Rows - 1 Then
            myaddItem
        End If
        Unload Search3
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.text = Search3.GRID1.TextMatrix(Search3.GRID1.Row, 0)
    Unload Search3
    myUndo
End If
End Sub

Private Sub CMD_PRINT_Click()
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
For i = 1 To GRID1.Rows - 2
    temptable.AddNew
    Select Case myPublic
        Case 1
            temptable!STR21 = ArbString("„” ‰œ „ﬁ»Ê÷«  „‰ ⁄„Ì·  —ﬁ„ : " & xDoc_No.text)
            temptable!str4 = " «·⁄„Ì· / " & " ( " & GRID1.TextMatrix(i, 2) & " ) "
        Case 2
            temptable!STR21 = ArbString("„” ‰œ ”œ«œ ‰ﬁœÌ… ·„Ê—œ —ﬁ„ : " & xDoc_No.text)
            temptable!str4 = " «·„Ê—œ / " & " ( " & GRID1.TextMatrix(i, 2) & " ) "
        Case 3
            temptable!STR21 = ArbString("„” ‰œ ”œ«œ ‰ﬁœÌ… ·⁄„Ì· —ﬁ„ : " & xDoc_No.text)
            temptable!str4 = " ⁄„Ì·   / " & " ( " & GRID1.TextMatrix(i, 2) & " ) "
        Case 4
            temptable!STR21 = ArbString("„” ‰œ ≈” ·«„ ‰ﬁœÌ… „‰ „Ê—œ —ﬁ„ : " & xDoc_No.text)
            temptable!str4 = " «·„Ê—œ / " & " ( " & GRID1.TextMatrix(i, 2) & " ) "
        Case 0
            temptable!STR21 = ArbString("„” ‰œ  ”ÊÌ… —ﬁ„ : " & xDoc_No.text)
            temptable!str4 = " «·„Ê—œ / " & " ( " & GRID1.TextMatrix(i, 2) & " ) "
        Case 11
            temptable!STR21 = ArbString("„” ‰œ  ”ÊÌ… —ﬁ„ : " & xDoc_No.text)
            temptable!str4 = " «·⁄„Ì· / " & " ( " & GRID1.TextMatrix(i, 2) & " ) "
    End Select
    
    temptable!str1 = cComp_Name
    temptable!val2 = GRID1.TextMatrix(i, 4)
    temptable!str3 = TurnValue(GRID1.TextMatrix(i, 3), "", Null)
    temptable!str6 = TurnValue(GRID1.TextMatrix(i, 2), "", Null)
    temptable!str8 = TurnValue(GetDesca("select desca from file0_50 where code = " & MyParn(GRID1.TextMatrix(1, 0)), con), "", Null)
    temptable!STR7 = TurnValue(GRID1.TextMatrix(i, 5), "", Null)
    temptable!str2 = xDate.text
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\P_CASH.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\Temp.MDB"
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing
Exit Sub
myError:
MsgBox Err.Description
Err.Clear

End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myError
    con.BeginTrans
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.text)
    con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.text)
    con.CommitTrans
    openCardTable
    If CardTable.EOF And CardTable.EOF Then
        myDefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub cmdExit_Click()
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
cString = "SELECT " & cFileHeader & ".Doc_No, Convert(Varchar(10)," & cFileHeader & ".Date,111),Min(" & DocClient & ".Desca)" & _
          " FROM (" & cFileHeader & " inner join " & cFile & " on " & cFileHeader & ".doc_no = " & cFile & ".Doc_NO) Inner Join " & DocClient & " on " & cFile & ".Code = " & DocClient & ".Code"
'If Not bopt2 Then cString = cString & " and date = " & DateSq(dSalesDate)
Generalarray(1) = cString
Generalarray(2) = " group by " & cFileHeader & ".Doc_No," & cFileHeader & ".Date order by " & cFileHeader & ".DATE ," & cFileHeader & ".DOC_NO "
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·«”„- «—ÌŒ «·„” ‰œ"
listarray(0, 1) = "(%%" & DocClient & ".Desca%% or " & _
                  " ##" & cFileHeader & ".Date##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·≈”„"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "Customers Query"
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
xDoc_No.SetFocus
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


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
'bEdit = Not bOpt5
bedit = True
openCon con
cStrBox = StrBox

FRM_CLOSED.Visible = lSupperVisor Or (bopt2 And ((cBranch = "00" Or cBranch >= "60") Or lShowBranch))
CmdDelInv.Visible = (cBranch = "00" Or lShowBranch Or cBranch >= "60")

Select Case myPublic
    Case 1 ' „ﬁ»Ê÷«  „‰ ⁄„·«¡
        sName = "«·⁄„Ì·"
        cFile = "File8_10"
        cFileHeader = "FILE8_10H"
        docMoveType = "4"
        DocTitle = "„ﬁ»Ê÷«  „‰ ⁄„·«¡"
        DocClient = "File3_10"
        CGROUP = "File3_20"
        cFileMove = "File3_11"
        
    Case 2 '„œ›Ê⁄«  ≈·Ï „Ê—œÌ‰
        sName = "«·„Ê—œ"
        cFile = "File8_20"
        cFileHeader = "FILE8_20H"
        docMoveType = "4"
        DocTitle = "„œ›Ê⁄«  «·Ì „Ê—œÌ‰"
        DocClient = "File4_10"
        CGROUP = "FILE4_50"
        cFileMove = "File4_11"
        
    Case 3 ' „œ›Ê⁄«  ≈·Ï ⁄„·«¡
        sName = "«·⁄„Ì·"
        cFile = "File8_30"
        cFileHeader = "FILE8_30H"
        docMoveType = "5"
        DocTitle = "„œ›Ê⁄«  ≈·Ï ⁄„·«¡"
        DocClient = "File3_10"
        CGROUP = "File3_20"
        cFileMove = "File3_11"
    Case 4 ' „ﬁ»Ê÷«  „‰ „Ê—œÌ‰
        sName = "«·„Ê—œ"
        cFile = "File8_40"
        cFileHeader = "FILE8_40H"
        docMoveType = "5"
        DocTitle = "„ﬁ»Ê÷«  „‰ „Ê—œÌ‰"
        DocClient = "File4_10"
        CGROUP = "FILE4_50"
        cFileMove = "File4_11"
'       CmdSave.Visible = False
'       CmdDelInv.Visible = False
    Case 0 '  ”ÊÌ«  „Ê—œÌ‰
        sName = "«·„Ê—œ"
        cFile = "File8_00"
        cFileHeader = "FILE8_00H"
        docMoveType = "5"
        DocTitle = " ”ÊÌ«  „Ê—œÌ‰ "
        DocClient = "File4_10"
        CGROUP = "FILE4_50"
        cFileMove = "File4_11"
    Case 11 '  ”ÊÌ«  „Ê—œÌ‰
        sName = "«·„Ê—œ"
        cFile = "File8_11"
        cFileHeader = "FILE8_11H"
        docMoveType = "5"
        DocTitle = " ”ÊÌ«  ⁄„·«¡ "
        DocClient = "File3_10"
        CGROUP = "FILE4_30"
        cFileMove = "File3_11"
End Select
'data1.ConnectionString = strCon
'data1.RecordSource = "SELECT * FROM FILE0_50"
'Set xBox2.RowSource = data1
'xBox2.ListField = "Desca"
'xBox2.BoundColumn = "Code"

Set GRID1.DataSource = DATA11
DATA11.ConnectionString = strCon

Me.Caption = DocTitle


defBox = RetDefBox()
defMosm = cPMosm
Fixgrd
openCardTable
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub

Private Sub grid1_EnterCell()
    If GRID1.Col = 2 Then
        GRID1.Editable = flexEDNone
    Else
        GRID1.Editable = flexEDKbdMouse
    End If
End Sub
Private Sub Grid1_GotFocus()
If GRID1.Row = 0 Then
    GRID1.SetFocus
    GRID1.Select 1, 0
End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And GRID1.Col = 1 Then CLIENTLOOKUP
If KeyCode = 46 And GRID1.Row <> GRID1.Rows - 1 And GRID1.Rows > 3 And bedit Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myError
        con.BeginTrans
        If GRID1.TextMatrix(GRID1.Row, GRID1.Cols - 1) <> "" Then
            con.Execute "Delete from " & cFile & " where ID = " & GRID1.TextMatrix(GRID1.Row, GRID1.Cols - 1)
        End If
        con.CommitTrans
        GRID1.RemoveItem GRID1.Row
        CalcTotals
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Col = 2 Then CLIENTLOOKUP
If KeyCode = 46 And Row <> GRID1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        GRID1.RemoveItem Row
    End If
End If
End Sub

Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case GRID1.Col
    Case 0
        If KeyCode = 27 Then
            Exit Sub
        End If
        If KeyCode = 112 Then
            CLIENTLOOKUP
        End If
End Select
End Sub
Private Function MYVALID() As Boolean
If Trim(xDoc_No.text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

With GRID1
    For nRow = 1 To .Rows - 1
        If cBranch = "00" Then
            If GetBoolean("select closed from mosm where mosm = " & MyParn(.TextMatrix(nRow, 5)), con) Then
                MsgBox " „Ê”„ „€·ﬁ  "
                Exit Function
            End If
        End If
    Next nRow
End With
'If IsDate(dLastdate) Then
'    If DateValue(xDate.Text) <= DateValue(dLastdate) Then
'        MsgBox "«· «—ÌŒ «ﬁ· „‰ «Œ—  «—ÌŒ «€·«ﬁ"
'        Exit Function
'    End If
'End If
If GRID1.Rows < 3 Then
    MsgBox "·«  ÊÃœ »‰Êœ  „  ”ÃÌ·Â«"
    Exit Function
End If

If cBranch > "60" Then
    If DateValue(xDate.text) < dMaxDate And Not lSupperVisor Then
        MsgBox " «—ÌŒ €Ì— ’«·Õ "
        Exit Function
    End If
End If
With GRID1
For i = 1 To .Rows - 2
    If .TextMatrix(i, 1) = "" Then
        .Select i, 0, i, GRID1.Cols - 1
        MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
        Exit Function
    End If
'    If Val(.TextMatrix(I, 4)) = 0 Then
'        MsgBox "ﬁÌ„… «·»‰œ €Ì— „”Ã·…"
'        Exit Function
'    End If
    If myPublic = 0 Or myPublic = 2 Or myPublic = 4 Then
        If GRID1.TextMatrix(i, 5) = "" Then
            .Select i, 0, i, .Cols - 1
            .ShowCell i, 5
            MsgBox "«·„Ê”„ €Ì— „Õœœ "
            Exit Function
        End If
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
    Dim GrdTable As New ADODB.Recordset
    If CardTable!ISCLOSED Then
        cmd_closed.BackColor = &H8080FF
        cmd_closed.Caption = "› Õ „” ‰œ"
        xclosed.Value = ssCBChecked
        xclosed.Visible = True
    Else
        cmd_closed.BackColor = vbGreen
        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
        xclosed.Value = ssCBUnchecked
        xclosed.Visible = False
    End If
    
    xUserName.Caption = CardTable!UserName & ""
    xUserName2.Caption = CardTable!USERNAME2 & ""
    XTIME1.Caption = CardTable!Time & ""
    XTIME2.Caption = CardTable!TIME2 & ""

xDoc_No.text = CardTable!doc_no
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub myloadgrd()
With GRID1
'                                   0                       1                   2                   3       4       5       6       7           8
    cString = "SELECT " & cFile & ".[BOX], " & cFile & ".CODE," & DocClient & ".DESCA," & cFile & ".desca, VALUE, MOSM , STORE ,' ' AS N7 ,SER_NO , ID " & _
               " FROM " & cFile & " LEFT JOIN " & DocClient & " ON " & cFile & ".CODE = " & DocClient & ".CODE " & _
               " where " & cFile & ".Doc_no = " & MyParn(xDoc_No.text)

    DATA11.RecordSource = cString
    DATA11.Refresh
    myaddItem
End With
CalcTotals
Fixgrd
End Sub
Private Sub myDefine()
    
    xUserName.Caption = ""
    xUserName2.Caption = ""
    XTIME1.Caption = ""
    XTIME2.Caption = ""
    
    cmd_closed.BackColor = &H8080FF
    cmd_closed.Caption = "› Õ „” ‰œ"
    xclosed.Visible = False
    xclosed.Value = ssCBUnchecked
    

xDoc_No.text = RetZero(Val(Newflag(cFileHeader, "doc_no", con)))
If cBranch = "00" Then
    xDate.text = ""
Else
    xDate.text = Format(dSalesDate, "DD-MM-YYYY")
End If
GRID1.Rows = 1
myaddItem
Handlecontrols DefineMode
CalcTotals
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode And bedit)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
CmdDelInv.Enabled = (nMode = LoadMode) And (xclosed.Value = 0)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
'xBox2.Enabled = cManBox = ""
CmdSave.Enabled = (xclosed.Value = 0)
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With GRID1
    If Col = 1 And Trim(.EditText) <> "" And IsNumeric(.EditText) Then .EditText = RetZero(.EditText, 3)
    If Col = 1 And Trim(.EditText) <> "" And Not IsNumeric(.EditText) Then
        If GetDesca("SELECT CODE FROM FILE4_10 WHERE code = " & MyParn(.EditText), con) = "" Then
            
            cCode = GetDesca("SELECT CODE FROM FILE4_10 WHERE SUBCODE = " & MyParn(.EditText), con)
            .EditText = cCode
        End If
    End If
End With
End Sub

Private Sub xDoc_No_LostFocus()
If Trim(xDoc_No.text) = "" Then Exit Sub
xDoc_No.text = RetZero(xDoc_No.text)
CardTable.Find "Doc_no = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function StrBox()
Dim boxtable As New ADODB.Recordset
If cBranch = "00" Then
    boxtable.Open "SELECT * FROM file0_50 WHERE ISSTOP = 0 ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
Else
    boxtable.Open "SELECT * FROM file0_50 WHERE ISSTOP = 0 AND TYPE <> 1  ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
End If
StrBox = "#  " & ";       "
Do Until boxtable.EOF
    StrBox = StrBox & "|#" & boxtable!CODE & ";" & boxtable!DESCA
    boxtable.MoveNext
Loop
End Function
Private Sub CLIENTLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From " & DocClient
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Private Function CalcTotals()
Dim nTotal As Double
With GRID1
For i = 1 To GRID1.Rows - 2
    nTotal = nTotal + Round(Val(GRID1.TextMatrix(i, 4)), 2)
Next
'xTotal.Text = nTotal
StatusBar1.Panels(1).text = turn(Myvalue(nTotal), "«·«Ã„«·Ì : ") & Myvalue(nTotal, "Fixed")
End With
End Function
Private Sub GrdDesc(nRow)
    GRID1.TextMatrix(nRow, 2) = GetDesca("Select Desca From " & DocClient & " Where code = " & MyParn(GRID1.TextMatrix(nRow, 1)), con) & ""
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
If xDoc_No.text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With GRID1
    .Editable = flexEDKbdMouse
    .Cols = 10
'                       0       1         2           3           4           5           6           7                 8
    .FormatString = "Œ“‰…|" & "ﬂÊœ|" & "«·≈”„|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "«·„Ê”„|" & "«·›—⁄|" & "—ﬁ„ «·‘Ìﬂ|" & "—ﬁ„ „” ‰œ|"
    .ColWidth(0) = 2000
    .ColWidth(1) = 1000
    .ColWidth(2) = 2500
    .ColWidth(3) = 4000
    .ColWidth(4) = 1500
    .ColWidth(5) = 1000
    .ColWidth(6) = 2000
    .ColWidth(7) = 0
    .ColWidth(8) = 1500
    
    If myPublic = 0 Or myPublic = 11 Then
        .ColHidden(0) = True
    End If
    
    If DocClient = "File3_10" Then
        .ColHidden(5) = True
        .ColHidden(6) = True
        .ColHidden(7) = True
    End If
    
    .ColHidden(.Cols - 1) = True
    .ColDataType(4) = flexDTDouble
    .ColFormat(4) = "#0.00"
    For i = 1 To GRID1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColComboList(0) = cStrBox
    .ColComboList(6) = StrList("SELECT CODE , DESCA FROM FILE0_40 ORDER BY CODE ", con)
    .ColComboList(5) = StrList("SELECT MOSM , MOSM FROM MOSM ORDER BY DATE DESC ", con)
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With GRID1
'If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And .TextMatrix(OldRow, .Cols - 1) = "" Then
'    If Not validRow(OldRow) Then
'        .RemoveItem OldRow
'        CalcTotals
'    End If
'End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
With GRID1
'If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 And .TextMatrix(.Row, .Cols - 1) = "" Then
'    .RemoveItem .Row
'    CalcTotals
'End If
End With
End Sub
Private Function validRow(nRow) As Boolean
With GRID1
If Not (IsNumeric(.TextMatrix(nRow, 1)) And Val(.TextMatrix(nRow, 4)) <> 0) Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim nBal As Double
With GRID1

If Col = 1 Then GrdDesc Row

If Not validRow(Row) Then Exit Sub
'If myPublic = 1 Or myPublic = 3 Then
'    nBal = Val(GetDesca("SELECT SUM(sal - SAL ) FROM FILE3_11 WHERE CODE = " & MyParn(.TextMatrix(Row, 1))) & "")
'Else
'    nBal = Val(GetDesca("SELECT SUM(SAL - PAY ) FROM FILE4_11 WHERE CODE = " & MyParn(.TextMatrix(Row, 1)) & " AND MOSM = " & MyParn(.TextMatrix(Row, 5))) & "")
'End If
'.TextMatrix(Row, 6) = Format(nBal, "#0.00")

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
    If xDoc_No.text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
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
cString = "SELECT * FROM " & cFileHeader
'If cManBox <> "" Then
'    cFilter = cFilter & turn(cFilter, " and ") & cFileHeader & ".DATE = " & DateSq(Format(dSalesDate, "DD-MM-YYYY"))
'    cFilter = cFilter & turn(cFilter, " and ") & cFileHeader & ".[BOX2] = " & MyParn(cManBox)
'End If
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myaddItem()
With GRID1
.AddItem ""
.TextMatrix(.Rows - 1, 5) = defMosm
'If xBox2.MatchedWithList Then
'    .TextMatrix(.Rows - 1, 0) = xBox2.BoundText
'Else
    If GRID1.Rows > 3 And .TextMatrix(.Rows - 2, 0) <> "" Then
        .TextMatrix(.Rows - 1, 0) = .TextMatrix(.Rows - 2, 0)
    End If
'End If
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
CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload

End Sub
Private Sub CMD_OPEN_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sFieldClose = "isclosed"
oClosefrm.sCaption = Me.Caption
oClosefrm.nMode = 1
oClosefrm.Show 1

CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload

End Sub



Private Sub cmd_closed_Click()
If CardTable!ISCLOSED Then
    con.Execute " update " & cFileHeader & " set isclosed = 0 where doc_no = " & MyParn(xDoc_No.text)
Else
    con.Execute " update " & cFileHeader & " set isclosed = 1 where doc_no = " & MyParn(xDoc_No.text)
End If
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload

End Sub


