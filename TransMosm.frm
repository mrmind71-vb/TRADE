VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form TransMosm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " —ÕÌ· —’Ìœ »Ì‰ «·„Ê«”„"
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
   Begin VB.Frame Frame5 
      Height          =   1860
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   -45
      Width           =   2085
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
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1800
         Visible         =   0   'False
         Width           =   1410
      End
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         TabIndex        =   23
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
         Picture         =   "TransMosm.frx":0000
         Alignment       =   4
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   1035
         TabIndex        =   24
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
         Picture         =   "TransMosm.frx":25CC
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   90
         TabIndex        =   25
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
         Picture         =   "TransMosm.frx":4C44
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin Threed.SSCommand cmd_supl 
      Height          =   555
      Left            =   4095
      TabIndex        =   20
      Top             =   1260
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   979
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "-"
      ButtonStyle     =   3
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
         Picture         =   "TransMosm.frx":7285
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TransMosm.frx":9455
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1650
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
         Picture         =   "TransMosm.frx":B59D
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TransMosm.frx":D765
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
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
         Picture         =   "TransMosm.frx":F8B4
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TransMosm.frx":11A94
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
         Picture         =   "TransMosm.frx":13BEF
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "TransMosm.frx":15DAB
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
         Picture         =   "TransMosm.frx":17EFA
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
         Picture         =   "TransMosm.frx":18044
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
         Left            =   2730
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TransMosm.frx":1A8DE
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
         Picture         =   "TransMosm.frx":1CE8A
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   150
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   2565
      Top             =   810
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
            TextSave        =   "11:30 ’"
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
      Height          =   1230
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   585
      Width           =   5820
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
         Left            =   3330
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
         Left            =   3330
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xmosm1 
         Height          =   315
         Left            =   90
         TabIndex        =   28
         Top             =   270
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xmosm2 
         Height          =   315
         Left            =   90
         TabIndex        =   31
         Top             =   675
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï „Ê”„ :"
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
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   675
         Width           =   930
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„‰ „Ê”„ :"
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
         Left            =   2310
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   270
         Width           =   870
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
         Left            =   4710
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
         Left            =   4710
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   255
         Width           =   930
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7215
      Left            =   180
      TabIndex        =   6
      Top             =   1845
      Width           =   15000
      _cx             =   26458
      _cy             =   12726
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
      Height          =   1230
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   585
      Width           =   1455
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TransMosm.frx":1F65D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   465
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "TransMosm.frx":1FA9F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   675
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   1620
      Top             =   315
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
   Begin Threed.SSCommand cmd_bal 
      Height          =   555
      Left            =   2340
      TabIndex        =   21
      Top             =   1260
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   979
      _Version        =   196610
      Font3D          =   3
      ForeColor       =   128
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "-"
      ButtonStyle     =   3
   End
   Begin MSAdodcLib.Adodc data2 
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
   Begin Threed.SSCommand cmd_balall 
      Height          =   510
      Left            =   5535
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   675
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   900
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " —ÕÌ· ﬂ· «·«—’œ…"
      PictureAlignment=   4
   End
End
Attribute VB_Name = "TransMosm"
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

aInsert(2, 0) = "MOSM1"
aInsert(2, 1) = addstring(xmosm1.BoundText)

aInsert(3, 0) = "MOSM2"
aInsert(3, 1) = addstring(xmosm2.BoundText)

On Error GoTo myerror
con.BeginTrans
If xDoc_No.Enabled Then
    xDoc_No.text = RetZero(Val(Newflag(cFileHeader, "doc_no", con)))
    aInsert(0, 1) = addstring(xDoc_No.text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
    con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xDoc_No.text))
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
Dim aInsert(9, 1)
With grid1
    For i = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.text)
                
        aInsert(1, 0) = "MOSM1"
        aInsert(1, 1) = addstring(.TextMatrix(i, 0))
        
        aInsert(2, 0) = "code"
        aInsert(2, 1) = addstring(grid1.TextMatrix(i, 1))
                
        aInsert(3, 0) = "Desca"
        aInsert(3, 1) = addstring(grid1.TextMatrix(i, 3))
        
        aInsert(4, 0) = "[value]"
        aInsert(4, 1) = Val(grid1.TextMatrix(i, 4))

        aInsert(5, 0) = "MOSM2"
        aInsert(5, 1) = addstring(grid1.TextMatrix(i, 5))
        
        
        aInsert(6, 0) = "row"
        aInsert(6, 1) = i
        
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
        If grid1.Row = grid1.Rows - 1 Then
            myaddItem
        End If
        Unload Search3
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
    myUndo
End If
End Sub

Private Sub cmd_balall_Click()
    Dim BalTable As New ADODB.Recordset
    Dim aInsert(5, 1), i As Double

    If MsgBox("”Ê› Ì „  —ÕÌ· «·«—’œ… „‰ „Ê”„ ≈·Ï „Ê”„ ·ﬂ· «·»Ì«‰« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    con.Execute " delete from " & cFile & " where doc_no = " & MyParn(xDoc_No.text)
    grid1.Rows = 1
    cStr1 = " SELECT code , SUM(SAL - PAY ) as BalCode FROM " & cFileMove & " WHERE MOSM = " & MyParn(xmosm1.BoundText) & " GROUP BY CODE "
    BalTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    With BalTable
        Do While Not .EOF
        If Val(!BalCode & "") <> 0 Then
            i = i + 1
            aInsert(0, 0) = "doc_no"
            aInsert(0, 1) = addstring(xDoc_No.text)
                        
            aInsert(1, 0) = "MOSM1"
            aInsert(1, 1) = addstring(xmosm1.BoundText)
                
            aInsert(2, 0) = "code"
            aInsert(2, 1) = addstring(!CODE)
            
            aInsert(3, 0) = "[value]"
            aInsert(3, 1) = Val(!BalCode & "")
        
            aInsert(4, 0) = "MOSM2"
            aInsert(4, 1) = addstring(xmosm2.BoundText)
                
            aInsert(5, 0) = "row"
            aInsert(5, 1) = i
                
            con.Execute CreateInsert(aInsert, cFile)
        End If
        .MoveNext
        Loop
    End With
    openCardTable
    myUndo
End Sub
Private Sub cmd_supl_Click()
'    supMovefrm.sCode = grid1.TextMatrix(grid1.Row, 1)
'    supMovefrm.SMosm = grid1.TextMatrix(grid1.Row, 5)
'    supMovefrm.Show
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.text)
    con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.text)
    con.CommitTrans
    openCardTable
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
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
mydefine
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
'bEdit = Not bOpt5
bedit = True
openCon con

DATA1.ConnectionString = strCon
DATA1.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC"
Set xmosm1.RowSource = DATA1
xmosm1.ListField = "Desca"
xmosm1.BoundColumn = "MOSM"
xmosm1.BoundText = cPMosm

DATA1.ConnectionString = strCon
DATA1.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC"
Set xmosm2.RowSource = DATA1
xmosm2.ListField = "Desca"
xmosm2.BoundColumn = "MOSM"

Select Case myPublic
    Case 1 ' ⁄„·«¡
        sName = "«·⁄„Ì·"
        cFile = "File3_30"
        cFileHeader = "FILE3_30H"
        docMoveType = "0"
        DocTitle = " —ÕÌ· —’Ìœ ⁄„·«¡ »Ì‰ „Ê«”„"
        DocClient = "File3_10"
        CGROUP = "File3_20"
        cFileMove = "File3_11"
    Case 2 ' „Ê—œÌ‰
        sName = "«·„Ê—œÌ‰"
        cFile = "File4_30"
        cFileHeader = "FILE4_30H"
        docMoveType = "0"
        DocTitle = " —ÕÌ· —’Ìœ „Ê—œÌ‰ »Ì‰ „Ê«”„"
        DocClient = "File4_10"
        CGROUP = "File4_20"
        cFileMove = "File4_11"
End Select

Set grid1.DataSource = DATA11
DATA11.ConnectionString = strCon

Me.Caption = DocTitle


Fixgrd
openCardTable
myUndo
mydefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub grid1_EnterCell()
    If grid1.Col = 2 Then
        grid1.Editable = flexEDNone
    Else
        grid1.Editable = flexEDKbdMouse
    End If
    With grid1
        cmd_supl.Caption = .TextMatrix(.Row, 2)
        If myPublic = 1 Then
            cmd_bal.Caption = Format(Val(GetDesca("select sum(sal-pay) from file3_11 where code = " & MyParn(.TextMatrix(.Row, 1)) & " and mosm = " & MyParn(.TextMatrix(.Row, 0)), con) & ""), "#0.00")
        Else
            cmd_bal.Caption = Format(Val(GetDesca("select sum(sal-pay) from file4_11 where code = " & MyParn(.TextMatrix(.Row, 1)) & " and mosm = " & MyParn(.TextMatrix(.Row, 0)), con) & ""), "#0.00")
        End If
    End With
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.SetFocus
    grid1.Select 1, 0
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 1 Then CLIENTLOOKUP
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 And bedit Then
    If MsgBox("Õ–› «·”Ã·  „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from " & cFile & " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        grid1.RemoveItem grid1.Row
        CalcTotals
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Col = 2 Then CLIENTLOOKUP
If KeyCode = 46 And Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·”Ã·  „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        grid1.RemoveItem Row
    End If
End If
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

'If IsDate(dLastdate) Then
'    If DateValue(xDate.Text) <= DateValue(dLastdate) Then
'        MsgBox "«· «—ÌŒ «ﬁ· „‰ «Œ—  «—ÌŒ «€·«ﬁ"
'        Exit Function
'    End If
'End If
'If grid1.Rows < 3 Then
'    MsgBox "·«  ÊÃœ »‰Êœ  „  ”ÃÌ·Â«"
'    Exit Function
'End If
'
'With grid1
'For I = 1 To .Rows - 2
'    If grid1.TextMatrix(I, 5) = "" And grid1.TextMatrix(I, 0) = "" Then
'        .Select I, 0, I, .Cols - 1
'        .ShowCell I, 5
'        MsgBox "«·„Ê”„ €Ì— „Õœœ "
'        Exit Function
'    End If
'Next
'End With
MYVALID = True
End Function
Private Sub myload()
Dim GrdTable As New ADODB.Recordset

xClosed.Value = IIf(CardTable!ISCLOSED, 1, 0)
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


xDoc_No.text = CardTable!doc_no
xmosm1.BoundText = CardTable!MOSM1 & ""
xmosm2.BoundText = CardTable!MOSM2 & ""
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
'xusername.Text = TurnValue(CardTable!UserName, Null, "")
Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub myloadgrd()
With grid1
'                                   0                       1                   2                   3       4       5       6
    cString = "SELECT " & cFile & ".MOSM1, " & cFile & ".CODE," & DocClient & ".DESCA," & cFile & ".desca, VALUE, MOSM2 , ID " & _
               " FROM " & cFile & " LEFT JOIN " & DocClient & " ON " & cFile & ".CODE = " & DocClient & ".CODE " & _
               " where " & cFile & ".Doc_no = " & MyParn(xDoc_No.text) & " ORDER BY CODE "

    DATA11.RecordSource = cString
    DATA11.Refresh
    myaddItem
End With
CalcTotals
Fixgrd
End Sub
Private Sub mydefine()
    If cmd_closed.Caption <> "≈€·«ﬁ „” ‰œ" Then
        cmd_closed.BackColor = &H8000000F
        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
        xClosed.Visible = False
    End If

xDoc_No.text = RetZero(Val(Newflag(cFileHeader, "doc_no", con)))
xDate.text = Format(dSalesDate, "DD-MM-YYYY")
xmosm1.BoundText = ""
xmosm2.BoundText = ""
xClosed.Value = 0
'xusername.Text = ""
grid1.Rows = 1
myaddItem
Handlecontrols DefineMode
CalcTotals
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
CmdDelInv.Enabled = (nMode = LoadMode) And (xClosed.Value = 0)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)

CmdSave.Enabled = (xClosed.Value = 0)

End Sub
Private Sub xDoc_No_LostFocus()
If Trim(xDoc_No.text) = "" Then Exit Sub
xDoc_No.text = RetZero(xDoc_No.text)
CardTable.Find "Doc_no = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
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
With grid1
For i = 1 To grid1.Rows - 2
    nTotal = nTotal + Round(Val(grid1.TextMatrix(i, 4)), 2)
Next
StatusBar1.Panels(1).text = turn(Myvalue(nTotal), "«·«Ã„«·Ì : ") & Myvalue(nTotal, "Fixed")
End With
End Function
Private Sub GrdDesc(nRow)
    grid1.TextMatrix(nRow, 2) = GetDesca("Select Desca From " & DocClient & " Where code = " & MyParn(grid1.TextMatrix(nRow, 1)), con) & ""
    grid1.TextMatrix(nRow, 4) = GetDesca("Select SUM(SAL-PAY) From FILE4_11 Where code = " & MyParn(grid1.TextMatrix(nRow, 1)) & " AND MOSM = " & MyParn(grid1.TextMatrix(grid1.Row, 0)), con) & ""
End Sub

Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With grid1
    .Editable = flexEDKbdMouse
    .Cols = 7
    .FormatString = "„‰ „Ê”„|" & "ﬂÊœ|" & sName & "|" & "«·»Ì«‰|" & "«·ﬁÌ„…|" & "≈·Ï „Ê”„|"
    .ColWidth(0) = 1500
    .ColWidth(1) = 800
    .ColWidth(2) = 3500
    .ColWidth(3) = 4000
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
        
    .ColHidden(.Cols - 1) = True
    .ColDataType(4) = flexDTDouble
    .ColFormat(4) = "#0.00"
    For i = 1 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColComboList(0) = StrList("SELECT MOSM , MOSM FROM MOSM ORDER BY DATE DESC ", con)
    .ColComboList(5) = StrList("SELECT MOSM , MOSM FROM MOSM ORDER BY DATE DESC ", con)
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
End With
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Not (IsNumeric(.TextMatrix(nRow, 1)) And Val(.TextMatrix(nRow, 4)) <> 0) Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim nBal As Double
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
If CardTable.BOF And CardTable.EOF Then
    mydefine
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
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT * FROM " & cFileHeader
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Sub myaddItem()
With grid1
.AddItem ""
End With
End Sub
Private Sub cmd_closed_Click()
Dim cString As String
con.BeginTrans
On Error GoTo myerror
cString = " update " & cFileHeader & " set ISclosed = " & IIf(xClosed.Value = 1, "0", "1")
cString = cString & turn(cString) & "doc_no = " & MyParn(xDoc_No.text)
con.Execute cString
con.CommitTrans
openCardTable
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
oClosefrm.sCaption = DocTitle
oClosefrm.sFieldClose = "isclosed"
oClosefrm.nMode = 0
oClosefrm.Show 1
openCardTable
myUndo
End Sub
Private Sub cmd_open_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sCaption = DocTitle
oClosefrm.sFieldClose = "isclosed"
oClosefrm.nMode = 1
oClosefrm.Show 1
openCardTable
myUndo
End Sub
