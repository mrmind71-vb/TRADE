VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Ship_Bills_PAY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”œ«œ »Ê«·’ «·‘Õ‰ «Ê‰ ·«Ì‰"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14955
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
   ScaleHeight     =   8085
   ScaleWidth      =   14955
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand cmd_addexel 
      Height          =   435
      Left            =   6390
      TabIndex        =   38
      Top             =   90
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
   Begin VB.Frame FRM_CLOSED 
      Height          =   2085
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   0
      Visible         =   0   'False
      Width           =   2040
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   31
         Top             =   1215
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   741
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
         TabIndex        =   32
         Top             =   180
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
         Picture         =   "Ship_Bills_PAY.frx":0000
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   1005
         Left            =   90
         TabIndex        =   33
         Top             =   180
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
         Picture         =   "Ship_Bills_PAY.frx":2AD5
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCheck xclosed 
         Height          =   330
         Left            =   90
         TabIndex        =   34
         Top             =   1665
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   8421631
         PictureMaskColor=   14737632
         Enabled         =   0   'False
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«·„” ‰œ „€·ﬁ"
         Alignment       =   1
         MaskColor       =   8421631
      End
   End
   Begin VB.Frame Frame5 
      Height          =   870
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   1215
      Width           =   4200
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
         TabIndex        =   29
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
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   135
         Width           =   1950
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
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   135
         Width           =   2175
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
         Left            =   2205
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   495
         Width           =   1950
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
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   495
         Width           =   2175
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   0
      Width           =   1500
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Ship_Bills_PAY.frx":5607
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Ship_Bills_PAY.frx":5A49
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   510
      ScaleWidth      =   14955
      TabIndex        =   9
      Top             =   7275
      Width           =   14955
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   -45
         Width           =   3975
         Begin VB.CommandButton cmdFirst 
            Height          =   375
            Left            =   45
            Picture         =   "Ship_Bills_PAY.frx":7FC2
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "√Ê·"
            Top             =   135
            Width           =   870
         End
         Begin VB.CommandButton cmdLast 
            Height          =   375
            Left            =   915
            Picture         =   "Ship_Bills_PAY.frx":A6A1
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "√ŒÌ—"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdPrevious 
            Height          =   375
            Left            =   2070
            Picture         =   "Ship_Bills_PAY.frx":CD7B
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "«·”«»ﬁ"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdNext 
            Height          =   375
            Left            =   2985
            Picture         =   "Ship_Bills_PAY.frx":F34E
            Style           =   1  'Graphical
            TabIndex        =   15
            TabStop         =   0   'False
            ToolTipText     =   "«· «·Ì"
            Top             =   135
            Width           =   915
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   8
      Top             =   7785
      Width           =   14955
      _ExtentX        =   26379
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
            TextSave        =   "03:02 „"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   5505
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4095
         Picture         =   "Ship_Bills_PAY.frx":1190E
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   2790
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Ship_Bills_PAY.frx":140E1
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Ship_Bills_PAY.frx":1668D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Ship_Bills_PAY.frx":18F27
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1500
      Left            =   6435
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   585
      Width           =   8385
      Begin VB.TextBox XNOTES 
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
         Height          =   360
         Left            =   270
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1035
         Width           =   6375
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
         Height          =   360
         Left            =   5220
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1425
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
         Height          =   360
         Left            =   315
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1470
      End
      Begin MSDataListLib.DataCombo xbank 
         Height          =   360
         Left            =   2085
         TabIndex        =   36
         Top             =   585
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12648447
         ListField       =   "54"
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·»‰ﬂ : "
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
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   675
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ« "
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
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1080
         Width           =   690
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
         Left            =   1830
         RightToLeft     =   -1  'True
         TabIndex        =   7
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
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   930
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4995
      Left            =   45
      TabIndex        =   3
      Top             =   2115
      Width           =   14775
      _cx             =   26061
      _cy             =   8811
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   1485
      Top             =   225
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
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
      Left            =   3690
      TabIndex        =   22
      Top             =   90
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
      Picture         =   "Ship_Bills_PAY.frx":19071
      Caption         =   "ÿ»«⁄… «·„” ‰œ  "
      Alignment       =   1
      PictureAlignment=   3
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
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
      Width           =   1245
      _ExtentX        =   2196
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   5805
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   39
      Top             =   7125
      Visible         =   0   'False
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label xbranch 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   240
      Left            =   3690
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "Ship_Bills_PAY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte, bedit As Boolean
Dim con As New ADODB.Connection, cDelFile As String
Dim CardTable As ADODB.Recordset, cStrBox As String
Dim cFile As String, cFileHeader As String, sName As String
Dim DocTitle As String, oSearchDoc As New Search3
Dim DocClient As String
Dim dLastdate As String, defBox As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(5, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.text)

aInsert(2, 0) = "NOTES"
aInsert(2, 1) = addstring(XNOTES.text)

aInsert(3, 0) = "bank"
aInsert(3, 1) = addstring(xbank.BoundText)

On Error GoTo myerror
'con.BeginTrans
If xDoc_No.Enabled Then
    
    aInsert(4, 0) = "username"
    aInsert(4, 1) = addstring(cusername)
    
    aInsert(5, 0) = "time"
    aInsert(5, 1) = "getdate()"
    
    xDoc_No.text = RetZero(Val(Newflag(cFileHeader, "doc_no", con)))
    aInsert(0, 1) = addstring(xDoc_No.text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
    aInsert(4, 0) = "username2"
    aInsert(4, 1) = addstring(cusername)
    
    aInsert(5, 0) = "time2"
    aInsert(5, 1) = "getdate()"
    
    con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xDoc_No.text))
End If
myreplaceGrd
'con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
'con.RollbackTrans
Err.Clear
End Function
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    myload
    Unload Search3
End If
If ActiveControl.Name = grid1.Name Then
    grid1.TextMatrix(grid1.Row, 0) = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    GrdDesc grid1.Row, grid1.TextMatrix(grid1.Row, 0)
    grid1.AddItem ""
    grid1.Select grid1.Rows - 1, 0
    grid1.ShowCell grid1.Rows - 1, 0
End If
End Sub

Private Sub cmd_addexel_Click()
    If Not myreplace Then Exit Sub
    AddFromExel
    cmdSave_Click

    CardTable.Requery
    If CardTable.EOF And CardTable.BOF Then
        myDefine
    Else
        If xDoc_No.Enabled Then CardTable.MoveLast Else CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
        myload
    End If

End Sub
Private Sub cmdDelinv_Click()
If GetDesca("SELECT DOC_NO FROM FILE6_91 WHERE DOC_NO = " & MyParn(xDoc_No.text), con) = "" Then
    If MsgBox("Õ–› «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
    '   con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.Text)
        con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.text)
        con.CommitTrans
        CardTable.Requery
        CmdNewInv_Click
    End If
Else
    MsgBox " ÌÊÃœ ”œ«œ ‘Õ‰ ›Ï «·„” ‰œ "
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
cString = "SELECT FILE6_91H.DOC_NO , Convert(VarChar(10),FILE6_91H.Date,111),file5_10.DESCA , FILE6_91H.NOTES " & _
          " FROM FILE6_91H INNER JOIN file5_10 ON FILE6_91H.bank = file5_10.CODE "
Generalarray(1) = cString
Generalarray(2) = " ORDER BY DATE DESC , DOC_NO "
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "»Ì«‰- «—ÌŒ «·„” ‰œ"
listarray(0, 1) = "(%%FILE6_91H.NOTES%% or " & _
                  " ##Date##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·»‰ﬂ "
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«·»Ì«‰"
GrdArray(3, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
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
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Dim cShip_NO As String
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    myDefine
Else
    If xDoc_No.Enabled Then CardTable.MoveLast Else CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
    myload
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
FRM_CLOSED.Visible = bedit
cFileHeader = "FILE6_91H"

    Set DATA3.Recordset = myRecordSet("select * from FILE5_10 where ISSHIP = 1 ", con)
    Set xbank.RowSource = DATA3
    xbank.ListField = "Desca"
    xbank.BoundColumn = "Code"

Set grid1.DataSource = DATA1
DATA1.ConnectionString = strCon
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM FILE6_91H ORDER BY DOC_NO", con, adOpenStatic, adLockOptimistic, adCmdText
myDefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub grid1_EnterCell()
    If xclosed.Value = 0 Then
        If grid1.Col = 0 Or grid1.Col = 7 Then grid1.Editable = flexEDKbdMouse Else grid1.Editable = flexEDNone
    Else
        grid1.Editable = flexEDNone
    End If
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.SetFocus
    grid1.Select 1, 0
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If xclosed.Value <> 0 Then Exit Sub
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 2 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "update file6_20h SET Date_Pay = NULL , bank = NULL WHERE DOC_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
            con.Execute "Delete from FILE6_91 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
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
Private Function MYVALID() As Boolean
If Trim(xDoc_No.text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If


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
    xbank.BoundText = CardTable!BANK & ""
    xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
    XNOTES.text = CardTable!NOTES & ""

    Handlecontrols LoadMode
    myloadgrd
End Sub
Private Sub myDefine()

    xclosed.Value = 0
    cmd_closed.BackColor = vbGreen
    cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
    xclosed.Value = ssCBUnchecked
    xclosed.Visible = False

    xUserName.Caption = ""
    xUserName2.Caption = ""
    
    XTIME1.Caption = ""
    XTIME2.Caption = ""

    xDoc_No.text = RetZero(Val(Newflag(cFileHeader, "doc_no", con)))
    xDate.text = Format(Date, "dd-mm-yyyy")
    xbank.BoundText = ""
    XNOTES.text = ""
grid1.Rows = 1
grid1.AddItem ""
Fixgrd
grid1.TextMatrix(grid1.Rows - 1, 0) = defBox
Handlecontrols DefineMode
CalcTotals
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
CmdDelInv.Enabled = (nMode = LoadMode) And (xclosed.Value = 0)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
CmdSave.Enabled = (xclosed.Value = 0)
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
With grid1
    If xclosed.Value = 0 And .Col = 0 And KeyCode = 112 Then LookDocSalesOnline
End With
End Sub

Private Sub xDoc_No_LostFocus()
    If Trim(xDoc_No.text) = "" Then Exit Sub
    xDoc_No.text = RetZero(xDoc_No.text)
    CardTable.Find "Doc_no = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
    If Not CardTable.EOF Then myload
End Sub
Private Function CalcTotals()
Dim nTotal As Double
With grid1
For i = 1 To grid1.Rows - 2
    nTotal = nTotal + Round(Val(grid1.TextMatrix(i, 3)), 2)
Next
StatusBar1.Panels(1).text = "«·«Ã„«·Ì : " & Format(nTotal, "Fixed")
End With
End Function
Private Function GrdDesc(nRow, pDoc_No) As Boolean
    With grid1
    If Len(pDoc_No) >= 10 Then
        aRet = aGetDesca("SELECT SHIP_NO , DATE , [TOTAL], doc_no , NAME , FILE0_40.DESCA  , DATE_PAY  , ONLINE_DOC FROM Q_SHIP2 INNER JOIN FILE0_40 ON FILE0_40.CODE = Q_SHIP2.STORE WHERE  DOC_NO = " & MyParn(pDoc_No), con)
    Else
        aRet = aGetDesca("SELECT SHIP_NO , DATE , [TOTAL], doc_no , NAME , FILE0_40.DESCA  , DATE_PAY  , ONLINE_DOC FROM Q_SHIP2 INNER JOIN FILE0_40 ON FILE0_40.CODE = Q_SHIP2.STORE WHERE  ONLINE_DOC = " & MyParn(pDoc_No), con)
    End If
    If UBound(aRet) > 0 Then
        If IsDate(aRet(7)) Then
            Inform "  „ ”œ«œ  «·»Ê·Ì’… " & grid1.TextMatrix(nRow, 0) & " » «—ÌŒ " & Format(aRet(7), "DD-MM-YYYY")
        Else
            .TextMatrix(nRow, 0) = aRet(4) & ""
            .TextMatrix(nRow, 1) = aRet(1) & ""
            .TextMatrix(nRow, 2) = Format(aRet(2), "DD-MM-YYYY")
            .TextMatrix(nRow, 3) = Format(aRet(3), "#0.00")
            .TextMatrix(nRow, 4) = aRet(8) & ""
            .TextMatrix(nRow, 5) = aRet(5)
            .TextMatrix(nRow, 6) = aRet(6)
            GrdDesc = True
        End If
    Else
        Inform " —ﬁ„ €Ì— ’ÕÌÕ " & grid1.TextMatrix(nRow, 0)
    End If
    End With
End Function
Private Sub xDoc_No_Validate(Cancel As Boolean)
    If xDoc_No.text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With grid1
    .Editable = flexEDKbdMouse
    .Cols = 9
  '                     0                  1                2               3                 4               5         6            7
    .FormatString = "—ﬁ„ ›« Ê—…|" & "—ﬁ„ »Ê·Ì’… «·‘Œ‰|" & " «—ÌŒ |" & "ﬁÌ„… «·»Ê·Ì’…|" & "—ﬁ„ «·ÿ·»|" & "«·⁄„Ì·|" & "«·›—⁄|" & "„·«ÕŸ« |"
    .ColWidth(0) = 1500
    .ColWidth(1) = 2000
    .ColWidth(2) = 1300
    .ColWidth(3) = 1200
    .ColWidth(4) = 1000
    .ColWidth(5) = 3000
    .ColWidth(6) = 2000
    .ColWidth(7) = 2000

    .ColHidden(grid1.Cols - 1) = True
    
    For i = 1 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub myreplaceGrd()
Dim aInsert(5, 1), cStr1 As String, cStr2 As String
With grid1
    For i = 1 To .Rows - 2
        If .TextMatrix(nRow, 4) <> "" Then
            aInsert(0, 0) = "doc_no"
            aInsert(0, 1) = addstring(xDoc_No.text)
                    
            aInsert(1, 0) = "DOC_SALES"
            aInsert(1, 1) = addstring(.TextMatrix(i, 0))
            
            aInsert(2, 0) = "PAYNOTES"
            aInsert(2, 1) = addstring(grid1.TextMatrix(i, 7))
            
            aInsert(3, 0) = "VALUE_PAY"
            aInsert(3, 1) = Val(grid1.TextMatrix(i, 3))
            
            aInsert(4, 0) = "row"
            aInsert(4, 1) = i
            
            aInsert(5, 0) = "ONLINE_DOC"
            aInsert(5, 1) = addstring(.TextMatrix(i, 4))
            
            If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
                cStr1 = cStr1 & CreateInsert(aInsert, "FILE6_91") & " ; "
'                con.Execute CreateInsert(aInsert, "FILE6_91")
            Else
                
                cStr1 = cStr1 & CreateUpdate(aInsert, "FILE6_91", " where ID = " & grid1.TextMatrix(i, .Cols - 1)) & " ; "
            End If
            cStr2 = cStr2 & " update file6_20h SET Date_Pay = " & addDate(xDate.text) & " , bank = " & addstring(xbank.BoundText) & "  WHERE DOC_NO = " & MyParn(.TextMatrix(i, 0)) & " ; "
'           con.Execute " update file6_20h SET Date_Pay = " & addDate(xDate.text) & " , bank = " & addstring(xbank.BoundText) & "  WHERE DOC_NO = " & MyParn(.TextMatrix(i, 0))
        End If
    Next
    If cStr1 <> "" Then con.Execute cStr1
    If cStr2 <> "" Then con.Execute cStr2
End With
End Sub
Private Sub myloadgrd()
Dim cString As String
With grid1
'                       0                       1               2           3               4                       5               6               7               8
    cString = " SELECT FILE6_91.DOC_SALES, Q_SHIP2.SHIP_NO , Q_SHIP2.DATE ,Q_SHIP2.TOTAL , FILE6_91.ONLINE_DOC  , Q_SHIP2.NAME , FILE0_40.DESCA ,  FILE6_91.PAYNOTES , FILE6_91.ID " & _
              " FROM FILE6_91 INNER JOIN Q_SHIP2 ON Q_SHIP2.DOC_NO = FILE6_91.DOC_SALES INNER JOIN FILE0_40 ON FILE0_40.CODE = Q_SHIP2.STORE  " & _
              " where FILE6_91.doc_no = " & MyParn(xDoc_No.text) & " Order By FILE6_91.ID "
    Set DATA1.Recordset = myRecordSet(cString, con)
    grid1.AddItem ""
End With
CalcTotals
Fixgrd
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With grid1
    If Col = 0 Then GrdDesc Row, .TextMatrix(Row, 0)
    If Row = .Rows - 1 Then
        .AddItem ""
        CalcTotals
    End If
    End With
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
Private Sub cmd_CLOSEDDATE_Click()
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  „‰  «—ÌŒ", "1-1"), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  Õ Ï  «—ÌŒ", Format(Date, "DD-MM-YYYY")), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ ≈€·«ﬁ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update " & cFileHeader & " set isclosed = 1 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ ≈€·«ﬁ „” ‰œ«  «·› —…" & nRec
            CardTable.Requery
            CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
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
            con.Execute " update " & cFileHeader & " set isclosed = 0 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ › Õ „” ‰œ«  «·› —…" & nRec
            CardTable.Requery
            CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
            If CardTable.EOF Then CardTable.MoveLast
            myload
        End If
    End If
End Sub
Private Sub CMD_PRINT_Click()
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable


    
With grid1
For i = 1 To .Rows - 2
    temptable.AddNew
    temptable!str1 = xDoc_No.text
    temptable!str2 = TurnValue(xbank.text)
    temptable!Date1 = Format(xDate.text, "DD-MM-YYYY")
    temptable!str4 = TurnValue(XNOTES.text)
    
    temptable!str5 = TurnValue(.TextMatrix(i, 0))
    temptable!str6 = TurnValue(.TextMatrix(i, 5))
    temptable!STR7 = TurnValue(.TextMatrix(i, 3))
    temptable!date2 = Format(.TextMatrix(i, 2), "DD-MM-YYYY")
    temptable!str3 = TurnValue(.TextMatrix(i, 6))
    
    temptable!VAL1 = Val(.TextMatrix(i, 3))
    temptable!str9 = TurnValue(.TextMatrix(i, 7))
    temptable.Update
Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans

Main.REPORT1.ReportFileName = App.Path & "\Reports\REP_SHIP_PAY.rpt"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\Temp.MDB"
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear

End Sub
Sub AddFromExel()
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet, nItem As String
    Dim xlwbook As Excel.Workbook, cItem As String
    Dim cFileName As String, cBarCode As String, cDoc_No As String
    Dim nCount As Double, cModel As String, nquant As Double, nPrice As Double, nPriceC As Double
    If MsgBox("≈÷«›… ”œ«œ «·»Ê«·’ „‰ „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
            With grid1
            i = 8
            Do While True
                i = i + 1
                Me.Caption = i
                If (xlsheet.Cells(i, 2)) <> "" Then
                    nCount = i
                Else
                    Exit Do
                End If
            Loop
            prog1.Visible = True
            prog1.Value = 0
            prog1.Max = nCount
            prog1.Min = 0
            For nRow = 9 To nCount
                prog1.Value = nRow
                  If xlsheet.Cells(nRow, 2) <> "" Then
                    cDoc_No = Replace(xlsheet.Cells(nRow, 2), "#", "")
                    grid1.TextMatrix(grid1.Rows - 1, 4) = cDoc_No
                    If Not GrdDesc(grid1.Rows - 1, cDoc_No) Then .RemoveItem .Rows - 1
                    grid1.AddItem ""
                End If
            Next nRow
            MsgBox " „ «÷«›… ”œ«œ «·»Ê«·’"
            End With
        End If
    End If
End Sub
Sub LookDocSalesOnline()
Dim Generalarray(5)
Dim listarray(2, 5)
Dim GrdArray(7, 1)
Set Generalarray(0) = Me


'                               0                   1                       2                                       3               4                       5                   6               7
Generalarray(1) = "SELECT   FILE6_20H.DOC_NO , FILE6_20H.DOC_NO2, CONVERT(VARCHAR(10), FILE6_20H.date, 111) , file0_40.desca  , FILE6_20H.ONLINE_DOC , FILE6_20H.SHIP_NO , FILE6_20H.NAME , FILE6_20H.PHONE " & _
                  " FROM    FILE6_20H INNER JOIN    FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE  WHERE FILE6_20H.ONLINE = 2 AND DATE_PAY IS NULL     "
Generalarray(2) = "ORDER BY FILE6_20H.DATE , FILE6_20H.DOC_NO2"

Generalarray(3) = 6000
Generalarray(5) = False

listarray(0, 0) = "—ﬁ„ «·»Ê‰ - —ﬁ„ «·»Ê·Ì’… - «·⁄„Ì· - —ﬁ„ «· ·Ì›Ê‰"
listarray(0, 1) = "(@@Doc_No2@@5 or %%file6_20H.NAME%% or %%file6_20H.PHONE%% or %%file6_20H.SHIP_NO%%) "

listarray(1, 0) = "«· «—ÌŒ"
listarray(1, 1) = "(##DATE##)"

listarray(2, 0) = "—ﬁ„ ÿ·»"
listarray(2, 1) = "(%%file6_20H.ONLINE_DOC%%)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(1, 1) = 1200

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·›—⁄ "
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "—ﬁ„ ÿ·»"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "—ﬁ„ »Ê·Ì’…"
GrdArray(5, 1) = 2000

GrdArray(6, 0) = "«·⁄„Ì·"
GrdArray(6, 1) = 2500

GrdArray(7, 0) = " ·Ì›Ê‰"
GrdArray(7, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
