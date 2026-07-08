VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form AccDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· ﬁÌÊœ  ”ÊÌ…"
   ClientHeight    =   10035
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   17850
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
   ScaleHeight     =   10035
   ScaleWidth      =   17850
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame FramClosed 
      Height          =   1815
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   540
      Width           =   1545
      Begin Threed.SSCommand cmd_closed 
         Height          =   375
         Left            =   45
         TabIndex        =   36
         Top             =   900
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         _Version        =   196610
         BackColor       =   8421631
         Caption         =   "≈€·«ﬁ «·„” ‰œ"
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         Height          =   375
         Left            =   45
         TabIndex        =   37
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         _Version        =   196610
         Caption         =   "≈€·«ﬁ › —…"
      End
      Begin Threed.SSCommand cmd_open 
         Height          =   375
         Left            =   45
         TabIndex        =   38
         Top             =   540
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   661
         _Version        =   196610
         BackColor       =   16777152
         Caption         =   "› Õ › —…"
      End
      Begin Threed.SSCheck xclosed 
         Height          =   330
         Left            =   45
         TabIndex        =   39
         Top             =   1350
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   8421631
         PictureMaskColor=   14737632
         Enabled         =   0   'False
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "„” ‰œ „€·ﬁ"
         Alignment       =   1
         MaskColor       =   8421631
      End
   End
   Begin VB.Frame Frame6 
      Height          =   870
      Left            =   1890
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1710
      Width           =   5145
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
         TabIndex        =   42
         Top             =   495
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
         TabIndex        =   41
         Top             =   495
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
         TabIndex        =   40
         Top             =   135
         Width           =   2175
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
         TabIndex        =   34
         Top             =   135
         Width           =   1950
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "«·„” Œœ„ :"
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
         Left            =   4185
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   180
         Width           =   885
      End
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
   End
   Begin VB.Frame Frame4 
      Caption         =   "ÿ»«⁄… ﬁÌÊœ «· ”ÊÌ…"
      Height          =   1005
      Left            =   1890
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   540
      Width           =   3660
      Begin VB.TextBox xDate2 
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
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   570
         Width           =   1515
      End
      Begin VB.TextBox xDate1 
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
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   135
         Width           =   1515
      End
      Begin VB.CommandButton Cmd_Print 
         Caption         =   "ÿ»«⁄… «·„” ‰œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1665
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   270
         Width           =   1860
      End
   End
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   9090
      Width           =   5280
      Begin VB.Label XT1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label4 
         Caption         =   "≈Ã„«·Ï „œÌ‰"
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
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   225
         Width           =   1080
      End
      Begin VB.Label XT2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label Label15 
         Caption         =   "≈Ã„«·Ï œ«∆‰"
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
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   225
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   12105
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   5460
      Begin VB.CommandButton CmdDelInv 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–› «·„” ‰œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1410
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2730
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4050
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1590
      Left            =   9945
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   765
      Width           =   7620
      Begin VB.TextBox xNote 
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
         Height          =   600
         Left            =   90
         MaxLength       =   1000
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   570
         Width           =   6420
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
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   1635
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
         Left            =   5175
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1320
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·»Ì«‰"
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
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   675
         Width           =   480
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   1830
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   225
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
         Left            =   6555
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   255
         Width           =   930
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   5580
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   540
      Width           =   1500
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   540
         Width           =   1320
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   135
         Width           =   1320
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   6705
      Top             =   90
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
      Height          =   300
      Left            =   0
      TabIndex        =   11
      Top             =   9735
      Width           =   17850
      _ExtentX        =   31485
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
            TextSave        =   "02:50 „"
         EndProperty
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6135
      Left            =   135
      TabIndex        =   17
      Top             =   2565
      Width           =   17430
      _cx             =   30745
      _cy             =   10821
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      Cols            =   4
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
   Begin VB.Frame Frame8 
      Height          =   570
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   9000
      Width           =   1920
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   525
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1395
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Move Last"
         Top             =   135
         Width           =   435
      End
   End
   Begin MSAdodcLib.Adodc DATA5 
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
End
Attribute VB_Name = "AccDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection, cstramsn As String
Dim bedit As Boolean
Dim CardTable As New ADODB.Recordset, codeTable As New ADODB.Recordset, CLIST As String
Dim formMode, cListBox As String
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim cId As String
Dim aInsert(4, 1)

aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xdoc_no.text)

aInsert(1, 0) = "DATE"
aInsert(1, 1) = addDate(xDate.text)

If xNote.text = "" Then xNote.text = "."
aInsert(2, 0) = "NOTE"
aInsert(2, 1) = addstring(xNote.text)

con.BeginTrans
If xdoc_no.Enabled Then
    
    aInsert(3, 0) = "username"
    aInsert(3, 1) = addstring(cusername)
    
    aInsert(4, 0) = "time"
    aInsert(4, 1) = "getdate()"
    
    
    xdoc_no.text = RetZero(Val(Newflag("ACC1_20h", "doc_no", con)))
    aInsert(0, 1) = addstring(xdoc_no.text)
    con.Execute CreateInsert(aInsert, "ACC1_20H")
Else
    aInsert(3, 0) = "username2"
    aInsert(3, 1) = addstring(cusername)
    
    aInsert(4, 0) = "time2"
    aInsert(4, 1) = "getdate()"
    con.Execute CreateUpdate(aInsert, "ACC1_20H", " where doc_no = " & MyParn(xdoc_no.text))
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
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 0 Then
        If grid1.Row = grid1.Rows - 1 Then grid1.AddItem ""
        grid1.TextMatrix(grid1.Row, 0) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 1) = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
        If grid1.Row = grid1.Rows - 2 Then
            grid1.AddItem ""
            grid1.Select grid1.Rows - 2, 0
        End If
    End If

ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    myload
    Unload Search3
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CMD_PRINT_Click()
Dim sourcetable As New ADODB.Recordset
Dim TargetTable As New ADODB.Recordset
Dim nDoc1 As Double
Dim nDoc2 As Double
contemp.Execute "DELETE * FROM TEMP"

TargetTable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
cString = " SELECT  acc1_20h.date , acc1_20.DOC_NO ,  acc1_20h.note , acc1_20.DESCA ,   acc1_20.Val1, acc1_20.Val2, acc1_20.ACC,  acc1_10.DESCA AS CodeDesc   " & _
          " FROM (acc1_20  inner join acc1_20h on acc1_20h.doc_no = acc1_20.doc_no )  LEFT OUTER JOIN ACC1_10 ON acc1_20.ACC = ACC1_10.CODE " & _
          "  where acc1_20.acc is not null  "
If IsDate(xDate1.text) Then
    If IsDate(xDate1.text) Then cString = cString & " and acc1_20h.date >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cString = cString & " and acc1_20h.date <= " & DateSq(xdate2.text)
Else
    cString = cString & " and acc1_20h.doc_no = " & MyParn(xdoc_no.text)
End If
cString = cString & " order by acc1_20h.date , acc1_20h.doc_no  "
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·« ÌÊÃœ ﬁÌÊœ „”Ã·… ··› —…"
    Exit Sub
End If
sourcetable.MoveFirst
Do
    TargetTable.AddNew
    TargetTable!str1 = "—ﬁ„ «·„”·”· " & sourcetable!doc_no
    TargetTable!str10 = RetZero(sourcetable!doc_no, 10)
    TargetTable!str2 = sourcetable!note
    TargetTable!str5 = sourcetable!codedesc
    TargetTable!str6 = sourcetable!DESCA
'    TargetTable!str9 = sourcetable!note
    TargetTable!str4 = sourcetable!ACC
    TargetTable!Date1 = sourcetable!Date
    
    TargetTable!VAL1 = sourcetable!VAL1
    TargetTable!val2 = sourcetable!val2
    
    TargetTable!STR19 = Firsttitle
    TargetTable!STR20 = " ›’Ì·Ï ﬁÌÊœ ÌÊ„Ì… "
    TargetTable.Update
    sourcetable.MoveNext
Loop Until sourcetable.EOF
    contemp.BeginTrans
    contemp.CommitTrans
    Main.REPORT1.ReportFileName = App.Path & "\Reports\Doc_Day_doc.rpt"
    Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    Main.REPORT1.Action = 1
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From ACC1_20 where Doc_No = " & MyParn(xdoc_no.text)
    con.Execute "Delete  From ACC1_20H where Doc_No = " & MyParn(xdoc_no.text)
    con.CommitTrans
    CardTable.Requery
    
    myDefine
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT ACC1_20H.Doc_No, ACC1_20H.ID_NO , ACC1_20H.ID_NO2 , CONVERT(VARCHAR(10), ACC1_20H.[DATE],111),ACC1_20H.NOTE " & _
                  " FROM ACC1_20H "
Generalarray(2) = " order by ACC1_20H.Doc_No"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = " «—ÌŒ-»Ì«‰"
listarray(0, 1) = "(##ACC1_20H.DATE##) " & _
                  " or ( %%ACC1_20H.NOTE%% ) "
                  
listarray(1, 0) = "—ﬁ„ ﬁÌœ - —ﬁ„ ≈–‰ «·’—›"
listarray(1, 1) = "(%%ACC1_20H.ID_NO%%) " & _
                  " or ( %%ACC1_20H.ID_NO2%% ) "
                  
GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "—ﬁ„ «·ﬁÌœ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "—ﬁ„ ≈–‰ «·’—›"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = " «—ÌŒ"
GrdArray(3, 1) = 1500


GrdArray(4, 0) = "»Ì«‰"
GrdArray(4, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„ ﬁÌÊœ ÌÊ„Ì… "
Search3.Show 1
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
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
CardTable.Requery
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If xdoc_no.Enabled Then
    CmdNewInv_Click
Else
    CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
    Handlecontrols LoadMode
    myload
End If
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.BOF And CardTable.EOF Then
    myDefine
    Exit Sub
End If
CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}" Else KeyAscii = 13
End If
End Sub
Private Sub Form_Load()
bedit = True
FramClosed.Visible = bIsClosed
con.CursorLocation = adUseClient
con.Open strCon

CardTable.CursorLocation = adUseClient
CardTable.Open "SELECT * FROM ACC1_20H ORDER BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText

codeTable.CursorLocation = adUseClient
codeTable.Open "ACC1_10", con, adOpenStatic, adLockReadOnly, adCmdTable


'CmdDelInv.Visible = bOpt5
Set grid1.DataSource = data1
data1.ConnectionString = strCon
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
    Fixgrd
    xdoc_no.text = "000001"
End If
End Sub
Sub dispProc()
formMode = dispMode
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set codeTable = Nothing
Set CardTable = Nothing
closeCon con
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then
    grid1.TextMatrix(Row, 1) = ""
    If grid1.TextMatrix(Row, 0) <> "" Then
        codeTable.Find "code = " & grid1.TextMatrix(Row, 0), , adSearchForward, adBookmarkFirst
        If Not codeTable.EOF Then grid1.TextMatrix(Row, 1) = codeTable!DESCA & ""
    End If
End If
CalcTotals
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 1 Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
With grid1
    .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = &H80000005
    .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = vbYellow
    If .Row = .Rows - 1 And .TextMatrix(.Row, 0) <> "" Then .AddItem ""
End With
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then AcntLookup_end Me, Search3
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If (Not bedit) Then Exit Sub
    If MsgBox("Õ–› «·ﬁÌœ „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "delete from ACC1_20 where SYS_KEY = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myloadgrd
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = 112 And grid1.Col = 0 Then AcntLookup_end Me, Search3
    
End Sub
Private Function MYVALID() As Boolean
CardTable.Find "Doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF And xdoc_no.Enabled Then
    MsgBox "„” ‰œ »‰›” «·—ﬁ„ „‰ ﬁ»·"
    Exit Function
End If

If xdoc_no.text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— „”Ã·"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "»‰Êœ €Ì— „”Ã·…"
    Exit Function
End If
With grid1
For i = 1 To .Rows - 2
    If .TextMatrix(i, 6) <> "" Then
        If GetBoolean("select closed from mosm where mosm = " & MyParn(.TextMatrix(i, 6)), con) Then
            MsgBox " „Ê”„ „€·ﬁ  "
            Exit Function
        End If
    End If
    If .TextMatrix(i, 0) <> "" Then
        codeTable.Find "code = " & MyParn(.TextMatrix(i, 0)), , adSearchForward, adBookmarkFirst
        If codeTable.EOF Then
            .Select i, 0, i, 2
            MsgBox "ﬂÊœ «·»‰œ €Ì— ”·Ì„"
            Exit Function
        End If
    End If
    If .TextMatrix(i, 8) <> "" Then
    If Not (IsNumeric(.TextMatrix(i, 8))) Then
        MsgBox " —«Ã⁄  ”ÃÌ· «·„’—Ê› "
        .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbRed
    End If
    codeTable.Find "code = " & MyParn(.TextMatrix(i, 0)), , adSearchForward, adBookmarkFirst
    If codeTable.EOF Then
        .Select i, 0, i, 2
        MsgBox "ﬂÊœ «·»‰œ €Ì— ”·Ì„"
        Exit Function
    End If

End If
If Val(.TextMatrix(i, 2)) = 0 And Val(.TextMatrix(i, 3)) = 0 And .TextMatrix(i, 0) <> "" Then
    MsgBox "ﬁÌ„… «·ﬁÌœ €Ì— „”Ã·…"
    Exit Function
End If
Next
CalcTotals
If Val(XT1.Caption) <> Val(XT2.Caption) Then
    MsgBox "ÿ—›Ï «·ﬁÌœ €Ì— „ ”«ÊÏ"
    Exit Function
End If
End With
MYVALID = True
End Function
Private Sub myload()
Dim GrdTable As ADODB.Recordset
Set GrdTable = New ADODB.Recordset
If CardTable!ISCLOSED2 Then
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
xdoc_no.text = CardTable!doc_no
xNote.text = CardTable!note & ""

xDate.text = Format(CardTable!Date, "DD-MM-YYYY")
xusercode.Caption = CardTable!userCode & ""

xusername.Caption = CardTable!UserName & ""
xUserName2.Caption = CardTable!USERNAME2 & ""

XTIME1.Caption = CardTable!Time & ""
XTIME2.Caption = CardTable!TIME2 & ""

Handlecontrols LoadMode
myloadgrd
CalcTotals
End Sub
Private Sub myDefine()



cmd_closed.BackColor = &H8080FF
cmd_closed.Caption = "› Õ „” ‰œ"
xClosed.Visible = False
xClosed.Value = ssCBUnchecked



xdoc_no.text = RetZero(Newflag("ACC1_20H", "DOC_NO", con))
xDate.text = ""
xNote.text = ""
xusercode.Caption = ""

xusername.Caption = ""
xUserName2.Caption = ""
XTIME1.Caption = ""
XTIME2.Caption = ""

grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode And bedit)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xdoc_no.Enabled = (nMode = DefineMode)
CmdSave.Enabled = bedit And (xClosed.Value = 0)
CmdDelInv.Enabled = bedit And (xClosed.Value = 0)
End Sub
Private Sub xDoc_No_LostFocus()
xdoc_no.text = RetZero(xdoc_no.text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub myreplaceGrd()
Dim aInsert(9, 1)
With grid1
    For i = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xdoc_no.text)
        
        aInsert(1, 0) = "ACC"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 0))
        
        aInsert(2, 0) = "[VAL1]"
        aInsert(2, 1) = Val(grid1.TextMatrix(i, 2))
        
        aInsert(3, 0) = "[VAL2]"
        aInsert(3, 1) = Val(grid1.TextMatrix(i, 3))
        
        If grid1.TextMatrix(i, 4) = "" Then grid1.TextMatrix(i, 4) = "."
        aInsert(4, 0) = "DESCA"
        aInsert(4, 1) = addstring(grid1.TextMatrix(i, 4))

        aInsert(5, 0) = "BOXNO"
        aInsert(5, 1) = addstring(grid1.TextMatrix(i, 5))
        
        aInsert(6, 0) = "MOSM"
        aInsert(6, 1) = addstring(grid1.TextMatrix(i, 6))
        
        aInsert(7, 0) = "BRANCH"
        aInsert(7, 1) = addstring(grid1.TextMatrix(i, 7))
        
        aInsert(8, 0) = "CHARGE2"
        aInsert(8, 1) = addstring(grid1.TextMatrix(i, 8))
        
        aInsert(9, 0) = "row"
        aInsert(9, 1) = i
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "ACC1_20")
        Else
            con.Execute CreateUpdate(aInsert, "ACC1_20", " where SYS_KEY = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub myloadgrd()
cString = "SELECT ACC1_20.ACC, ACC1_10.DESCA, ACC1_20.Val1, ACC1_20.Val2, ACC1_20.DESCA , BOXNO , ACC1_20.MOSM , ACC1_20.BRANCH  , ACC1_20.CHARGE2 , ACC1_20.SYS_KEY FROM ACC1_20 LEFT OUTER JOIN  ACC1_10 ON ACC1_20.ACC = ACC1_10.CODE  " & _
          " where doc_no = " & MyParn(xdoc_no.text) & " ORDER BY SYS_KEY "
data1.RecordSource = cString
data1.Refresh
grid1.AddItem ""
Fixgrd
End Sub
Private Sub Fixgrd()
With grid1
    .Cols = 10
'                       0               1           2       3       4               5           6               7           8
    .FormatString = "ﬂÊœ Õ”«»|" & "«·Õ”«»|" & "„œÌ‰|" & "œ«∆‰|" & "«·»Ì«‰|" & "„” ‰œ —ﬁ„|" & "«·„Ê”„|" & "«·›—⁄|" & "„’—Ê›  ÊﬂÌ·|"
    .ColWidth(0) = 1000
    .ColWidth(1) = 2200
    .ColWidth(2) = 1100
    .ColWidth(3) = 1100
    .ColWidth(4) = 4000
    .ColWidth(5) = 1500
    .ColWidth(6) = 1000
    .ColWidth(7) = 1500
    .ColWidth(8) = 2000
    
    .ColDataType(0) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTString
    .ColComboList(6) = StrList("SELECT MOSM , MOSM FROM MOSM ORDER BY DATE DESC ", con)
    .ColComboList(7) = StrListA("SELECT CODE , DESCA FROM BRANCH ORDER BY CODE ", con)
    .ColComboList(8) = StrListA("SELECT CODE , DESCA FROM FILE8_51 ORDER BY DESCA ", con)
    For i = 1 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    grid1.ColHidden(grid1.Cols - 1) = True
End With
End Sub
Private Function CalcTotals()
Dim nT1 As Double, nT2 As Double

With grid1
For i = 1 To grid1.Rows - 1
    If Val(.TextMatrix(i, 2)) <> 0 Then nT1 = nT1 + Val(.TextMatrix(i, 2))
    If Val(.TextMatrix(i, 3)) <> 0 Then nT2 = nT2 + Val(.TextMatrix(i, 3))
Next
XT1.Caption = Format(nT1, "Fixed")
XT2.Caption = Format(nT2, "Fixed")
End With
End Function
Private Sub cmd_CLOSEDDATE_Click()
    Dim cSecDoc As String
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  „‰  «—ÌŒ", ""), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  Õ Ï  «—ÌŒ", ""), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ ≈€·«ﬁ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update ACC1_20H set isclosed2 = 1 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ ≈€·«ﬁ „” ‰œ«  «·› —…" & nRec
            CardTable.Requery
            CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
            If CardTable.EOF Then CardTable.MoveLast
            myload
        End If
    End If
End Sub

Private Sub cmd_open_Click()
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "› Õ „” ‰œ«  „‰  «—ÌŒ", ""), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "› Õ  „” ‰œ«  Õ Ï  «—ÌŒ", ""), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ › Õ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update ACC1_20H set isclosed2 = 0 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ › Õ „” ‰œ«  «·› —…" & nRec
            CardTable.Requery
            CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
            If CardTable.EOF Then CardTable.MoveLast
            myload
        End If
    End If
End Sub
Private Sub cmd_closed_Click()
If CardTable!ISCLOSED2 Then
    con.Execute " update ACC1_20H set isclosed2 = 0 where doc_no = " & MyParn(xdoc_no.text)
Else
    con.Execute " update ACC1_20H set isclosed2 = 1 where doc_no = " & MyParn(xdoc_no.text)
End If
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Sub Fix_Data()
On Error Resume Next
Dim FS1 As New ADODB.Command
cString = ReadFile(App.Path & "\sql_1.TXT")
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
Err.Clear
End Sub
