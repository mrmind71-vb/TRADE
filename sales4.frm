VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form Salesfrm 
   Caption         =   "›Ê« Ì— «·„»Ì⁄« "
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14235
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   14235
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame8 
      Height          =   555
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   6885
      Width           =   1905
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
         TabIndex        =   47
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
         TabIndex        =   46
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
         TabIndex        =   45
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
         TabIndex        =   44
         ToolTipText     =   "Move Last"
         Top             =   135
         Width           =   435
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   2250
      Top             =   225
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
         Name            =   "MS Sans Serif"
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
      Height          =   600
      Left            =   7650
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   6885
      Width           =   4560
      Begin VB.Label xTotal2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   135
         TabIndex        =   41
         Top             =   180
         Width           =   2445
      End
      Begin VB.Label Label13 
         Caption         =   "’«›Ì ﬁÌ„… «·„»Ì⁄«  :"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   225
         Width           =   1770
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Height          =   4110
      Left            =   135
      TabIndex        =   12
      Top             =   2745
      Width           =   13965
      _cx             =   24633
      _cy             =   7250
      _ConvInfo       =   1
      Appearance      =   2
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
      AutoResize      =   -1  'True
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
      TabBehavior     =   0
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
   Begin VB.Frame Frame5 
      Height          =   1680
      Left            =   90
      TabIndex        =   34
      Top             =   360
      Width           =   2715
      Begin VB.TextBox xCredit 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   90
         MaxLength       =   10
         TabIndex        =   11
         Top             =   1260
         Width           =   1725
      End
      Begin VB.TextBox xBoon 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   90
         MaxLength       =   10
         TabIndex        =   10
         Top             =   900
         Width           =   1725
      End
      Begin VB.TextBox xVisa 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   90
         MaxLength       =   10
         TabIndex        =   9
         Top             =   540
         Width           =   1725
      End
      Begin VB.TextBox xcash 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   90
         MaxLength       =   10
         TabIndex        =   8
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "¬Ã· :"
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
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»Ê‰ :"
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
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "›Ì“« :"
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
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   630
         Width           =   390
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ«‘ :"
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
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   270
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   2970
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   8385
      Begin VB.CommandButton cmdUndo 
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
         Height          =   420
         Left            =   2925
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Õ›Ÿ"
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
         Left            =   4275
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
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
         Left            =   1440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   180
         Width           =   1455
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
         Left            =   6975
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
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
         Left            =   5625
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
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
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   11385
      TabIndex        =   13
      Top             =   0
      Width           =   2760
      Begin VB.CommandButton cmdPrint 
         Caption         =   "ÿ»«⁄…"
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
         Left            =   1440
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton CMD_MOVE 
         Caption         =   "«” ⁄·«„ «’‰«›"
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
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   5625
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame3 
      Height          =   1365
      Left            =   2835
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   675
      Width           =   11265
      Begin VB.TextBox xInv_No 
         Alignment       =   2  'Center
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
         Left            =   135
         MaxLength       =   10
         TabIndex        =   2
         Top             =   225
         Width           =   1725
      End
      Begin VB.TextBox xTotal 
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
         Left            =   135
         MaxLength       =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   945
         Width           =   1725
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
         Left            =   3285
         MaxLength       =   10
         TabIndex        =   1
         Top             =   180
         Width           =   1860
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
         Left            =   8730
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   3
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox xStore 
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
         Left            =   8730
         MaxLength       =   10
         TabIndex        =   0
         Top             =   180
         Width           =   1275
      End
      Begin MSDataListLib.DataCombo xPayment 
         Height          =   315
         Left            =   8010
         TabIndex        =   6
         Top             =   900
         Width           =   1995
         _ExtentX        =   3519
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
      Begin MSDataListLib.DataCombo xCode 
         Height          =   315
         Left            =   3285
         TabIndex        =   4
         Top             =   540
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xMan 
         Bindings        =   "sales4.frx":0000
         DataSource      =   "DATA1"
         Height          =   315
         Left            =   135
         TabIndex        =   5
         Top             =   585
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·»«∆⁄"
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
         Left            =   1980
         TabIndex        =   32
         Top             =   630
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„” ‰œ"
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
         Left            =   1935
         TabIndex        =   31
         Top             =   270
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï ﬁÌ„… «·›« Ê—…"
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
         Left            =   1935
         TabIndex        =   30
         Top             =   1035
         Width           =   1620
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ "
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
         Left            =   5265
         TabIndex        =   29
         Top             =   270
         Width           =   585
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·⁄„Ì·"
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
         Left            =   5220
         TabIndex        =   28
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "„”·”·"
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
         Left            =   10170
         TabIndex        =   27
         Top             =   585
         Width           =   690
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ—Ìﬁ… «·œ›⁄"
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
         Left            =   10155
         TabIndex        =   26
         Top             =   945
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Õ·"
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
         Left            =   10155
         TabIndex        =   25
         Top             =   225
         Width           =   480
      End
      Begin VB.Label xStoreDesca 
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   6255
         TabIndex        =   24
         Top             =   180
         Width           =   2445
      End
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   9360
      TabIndex        =   33
      Top             =   2025
      Width           =   4740
      Begin VB.CommandButton Command2 
         Caption         =   " ”ÃÌ· «’‰«› «·„— Ã⁄« "
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
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   180
         Width           =   2265
      End
      Begin VB.CommandButton cmdAddItems 
         Caption         =   " ”ÃÌ· «’‰«› «·„»Ì⁄« "
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
         Left            =   2385
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   180
         Width           =   2265
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
         Name            =   "MS Sans Serif"
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA4 
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data10 
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
         Name            =   "MS Sans Serif"
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
Attribute VB_Name = "Salesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CardTable As ADODB.Recordset

Dim searchitem As New Search
Dim Search1 As New Search, Search2 As New Search, bMarket As Boolean
Dim formMode, dDateLast As String
Public myPublic As Integer
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca,file1_50.desca,file1_10.cost From file1_10 left join file1_50 on file1_10.group = file1_50.code"
Generalarray(2) = "Order by file1_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = True

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter%' or  %%FILE1_10.DESCA%%) "

listarray(1, 0) = "«·„Ã„Ê⁄…"
listarray(1, 1) = "(%%FILE1_50.DESCA%%)"

GrdArray(0, 0) = "ﬂÊœ «·’‰›"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·’‰›"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«· ﬂ·›…"
GrdArray(3, 1) = 0

searchArray = Array(Generalarray, listarray, GrdArray)
Load searchitem
searchitem.Caption = "«” ⁄·«„ «·«’‰«›"
searchitem.Show 1
End Sub
Private Function myreplace() As Boolean
Dim aInsert(11, 1)

aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_no.Text)

aInsert(1, 0) = "inv_no"
aInsert(1, 1) = addstring(xInv_No.Text)

aInsert(2, 0) = "[Date]"
aInsert(2, 1) = datesq(xDate.Text)

aInsert(3, 0) = "Code"
aInsert(3, 1) = addstring(xCode.BoundText)

aInsert(4, 0) = "Store"
aInsert(4, 1) = addstring(xStore.Text)

aInsert(5, 0) = "man"
aInsert(5, 1) = addstring(xMan.BoundText)

aInsert(6, 0) = "payment"
aInsert(6, 1) = addstring(xPayment.BoundText)

aInsert(7, 0) = "Cash"
aInsert(7, 1) = Val(xcash.Text)

aInsert(8, 0) = "Visa"
aInsert(8, 1) = Val(xVisa.Text)

aInsert(9, 0) = "Boon"
aInsert(9, 1) = Val(xBoon.Text)

aInsert(10, 0) = "Credit"
aInsert(10, 1) = Val(xCredit.Text)

aInsert(11, 0) = "Total"
aInsert(11, 1) = Val(xTotal.Text)

con.BeginTrans
If xDoc_no.Tag = DefineMode Then
    xDoc_no.Text = RetZero(Newflag("File6_20h", "doc_no"))
    aInsert(0, 1) = addstring(xDoc_no.Text)
    con.Execute CreateInsert(aInsert, "File6_20H")
Else
    con.Execute CreateUpdate(aInsert, "File6_20H", " where doc_no = " & MyParn(xDoc_no.Text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myproc()
On Error GoTo myerror
If ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "DOC_NO = " & MyParn(Search1.grid1.TextMatrix(Search1.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    Search1.Hide
    myload
ElseIf TypeOf ActiveControl Is TextBox Then
    ActiveControl.Text = Search2.grid1.TextMatrix(Search2.grid1.Row, 0)
    Unload Search2
End If
Exit Sub
myerror:
End Sub
Private Sub cmdClient_Click()
publicFlag = 2
Clients.Show 1
End Sub

Private Sub cmdAdditem_Click()

End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    ' Õ–› «·„” ‰œ
   con.Execute "Delete  From FILE6_20H where Doc_No = " & MyParn(xDoc_no.Text)
   con.Execute "Delete  From FILE6_20 where Doc_No = " & MyParn(xDoc_no.Text)
              
    con.CommitTrans
    CardTable.Requery
    If CardTable.BOF And CardTable.EOF Then
        myDefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_no.Text), , adSearchBackward, adBookmarkLast
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
Private Sub cmdExit_Click()
If MsgBox("Œ—ÊÃ !! ” ›ﬁœ ﬂ· «·»Ì«‰«  «·€Ì— „Õ›ÊŸ… ! „Ê«›ﬁ ø", vbYesNo + vbDefaultButton2) = vbYes Then Unload Me
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
Private Sub CmdNewInv_Click()
myDefine
xDoc_no.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID2 Then Exit Sub
mysave
End Sub
Private Sub CmdUndo_Click()
If MsgBox(" —«Ã⁄ ⁄‰  ”ÃÌ· «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Sub
If CardTable.BOF And CardTable.EOF Then
    myDefine
    Exit Sub
End If
If xDoc_no.Enabled Then
    CardTable.Find "Doc_No = " & MyParn(xDoc_no.Text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
End If
myload
End Sub
Private Sub Command3_Click()
mysave False
doprint
End Sub
Private Sub cmdAddItems_Click()
If Not MYVALID Then Exit Sub
AddSalefrm.myFlag = 0
AddSalefrm.sdoc_no = xDoc_no.Text
AddSalefrm.Show 1
End Sub

Private Sub Command5_Click()
AddItems
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
bEdit = True
Set CardTable = New ADODB.Recordset
CardTable.CursorLocation = adUseClient
CardTable.Open "SELECT *,FILE0_40.DESCA AS StoreDesca FROM FILE6_20H INNER JOIN FILE0_40 ON FILE6_20H.STORE = FILE0_40.CODE ORDER BY DOC_NO", con, adOpenKeyset, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "FILE3_10"
Set xCode.RowSource = data1
xCode.ListField = "Desca"
xCode.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "EMP"
Set xMan.RowSource = DATA2
xMan.ListField = "Desca"
xMan.BoundColumn = "Code"

data3.ConnectionString = strCon
data3.RecordSource = "FILE0_90"
Set xPayment.RowSource = data3
xPayment.ListField = "Desca"
xPayment.BoundColumn = "Code"

Set grid1.DataSource = data10
data10.ConnectionString = strCon

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
    Fixgrd
    xDoc_no.Text = RetZero("1", 6)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
ItemTable.Close
Set ItemTable = Nothing
Unload Search
Unload Search1
Unload Search2
Err.Clear
End Sub
Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.Select 1, 1
End If
End Sub
Private Sub XBOX_Click(Area As Integer)
If Not XBOX.MatchedWithList Then XBOX.BoundText = "'"
End Sub


Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ClientLookup
End Sub
Private Sub xCode_LostFocus()
'xBalance.Caption = Format(GetDesca("Select sum(Format(val(SAL & '') - val (pay & ''),'Fixed')) FROM " & cFileMove & " WHERE CODE = " & MyParn(xCode.Text)), "fixed")
End Sub
Private Sub xDiscount_LostFocus()
CalcTotals
End Sub
Private Function MYVALID2() As Boolean
If Val(xcash.Text) + Val(xVisa.Text) + Val(xBoon.Text) + Val(xCredit.Text) <> Val(xTotal2.Caption) Then
    MsgBox "”œ«œ «·»Ê‰ €Ì— ’ÕÌÕ"
    Exit Function
End If
MYVALID2 = True
End Function
Private Sub myload()
xDoc_no.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xInv_No.Text = CardTable!INV_NO & ""
xStore.Text = CardTable!STORE & ""
xStoreDesca.Caption = CardTable!StoreDesca & ""
xMan.Text = CardTable!MAN & ""
xCode.BoundText = CardTable!CODE & ""
xPayment.BoundText = CardTable!PAYMENT & ""
xcash.Text = TurnValue(Val(CardTable!CASH & ""), 0, "")
xVisa.Text = TurnValue(Val(CardTable!VISA & ""), 0, "")
xBoon.Text = TurnValue(Val(CardTable!BOON & ""), 0, "")
xCredit.Text = TurnValue(Val(CardTable!CREDIT & ""), 0, "")
myloadgrd
End Sub
Sub myloadgrd()
With grid1
    cString = "SELECT FILE6_20.ITEM,FILE1_10.BARCODE,FILE1_10.MODEL,FILE1_10.DESCA,FILE1_10.C_COLOR,FILE1_10.SCAL,Quant,FILE6_20.Price, FILE6_20.DISCOUNT, FILE6_20.COST, FILE6_20.TOTAL,FILE6_20.SYS_KEY " & _
          " FROM FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_no.Text)
    data10.RecordSource = cString
    data10.Refresh
    grid1.AddItem ""
End With
Handlecontrols LoadMode
CalcTotals
Fixgrd
End Sub
Private Sub myDefine()
xDoc_no.Text = RetZero(Newflag("FILE6_20H", "doc_no"))
xDate.Text = Format(Date, "dd-mm-yyyy")
'xStore.Text = ""
xCode.BoundText = ""
xInv_No.Text = ""

xMan.BoundText = ""
xPayment.BoundText = ""

xcash.Text = ""
xVisa.Text = ""
xBoon.Text = ""
xCredit.Text = ""

grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewinv.Enabled = nMode = LoadMode And bEdit
cmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = nMode = LoadMode And bEdit
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_no.Enabled = (nMode = DefineMode)
xDoc_no.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
xDoc_no.Text = RetZero(xDoc_no.Text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_no.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
With grid1
On Error GoTo myerror
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
     If MsgBox("Õ–› «·’‰› !‰⁄„ «„ ·« ø", vbOKCancel) = vbOK Then
        If .TextMatrix(.Row, .Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "DELETE FROM FILE6_20 WHERE SYS_KEY = " & .TextMatrix(.Row, .Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
        CalcTotals
    End If
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
myloadgrd
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case grid1.Col
    Case 1
        If KeyCode = 27 Then
            Exit Sub
        End If
End Select
End Sub
Private Function CalcTotals()
With grid1
For i = 1 To .Rows - 2
    nTotalItemRow = Val(.TextMatrix(i, 6)) * Val(.TextMatrix(i, 7))
    nTotalitem = nTotalitem + (nTotalItemRow)
    nDiscount = (Val(.TextMatrix(i, 8)) / 100)
    nTotalDiscount = nTotalDiscount + (nDiscount * nTotalItemRow)
    grid1.TextMatrix(i, 10) = nTotalItemRow - (nDiscount * nTotalItemRow)
Next
xTotal2.Caption = nTotalitem - nTotalDiscount
End With
End Function
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,[DATE] AS DATE1,DATE, " & cFileClient & ".Desca " & _
                  " FROM  (FILE6_20H left JOIN FILE3_10 ON FILE6_20H.CODE  = FILE3_10.CODE )"

Generalarray(2) = "Order by Date"
Generalarray(3) = 6000
Generalarray(5) = True


listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
listarray(0, 1) = "(Doc_No Like '%cFilter%' or  " & cFileClient & ".DESCA LIKE '%cFilter%' OR " & _
                  "##date##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "≈”„ " & cCodeDesca
GrdArray(3, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search1
Search1.Caption = "«” ⁄·«„"
Search1.Show 1
End Sub
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow And Trim(grid1.TextMatrix(i, nCol)) <> "" Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function
Private Function nofoundOther() As Boolean
For i = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(i, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 2) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow
        Exit Function
    End If
Next
nofoundOther = True
End Function

Private Sub xRate_LostFocus()
If Val(xRate.Text) <> 0 Then
    xTax.Text = Format(Val(xTotalDis.Caption) * (Val(xRate.Text) / 100), "Fixed")
    CalcTotals
End If
End Sub
Private Function validRow(nRow) As Boolean
If nRow > 0 Then
    If Trim(grid1.TextMatrix(nRow, 1)) = "" Then Exit Function
    If Trim(grid1.TextMatrix(nRow, 2)) = "" Then Exit Function
   ' If Val(grid1.TextMatrix(nRow, 5)) = 0 Then Exit Function
End If
validRow = True
End Function
Sub additemProc()
grid1.RemoveItem grid1.Rows - 1
With additemfrm.grid1
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 4)) <> 0 Then
            grid1.AddItem ""
            grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
            grid1.TextMatrix(grid1.Rows - 1, 1) = .TextMatrix(i, 0)
            grid1.TextMatrix(grid1.Rows - 1, 2) = retitem(.TextMatrix(i, 0), "desca")
            grid1.TextMatrix(grid1.Rows - 1, 3) = .TextMatrix(i, 4)
            grid1.TextMatrix(grid1.Rows - 1, 4) = .TextMatrix(i, 5)
        End If
    Next
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
    CalcTotals
End With
End Sub
Private Function RetItemBalance(cItem, cStore, DDate) As Double
If cItem = "" Then Exit Function
movetable.Seek Array(cItem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> cItem Or cStore <> movetable!STORE Or DateValue(movetable!Date) > DateValue(Format(DDate, "dd-mm-yyyy")) Then Exit Do
    'If Not (movetable!Type = cItemmove And movetable!Doc_Id = xDoc_No.Text) Then
        RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    'End If
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
.FormatString = "ﬂÊœ|" & "»«—ﬂÊœ|" & "«·„ÊœÌ·|" & "«·»Ì«‰|" & "—ﬁ„ «··Ê‰|" & "«·„ﬁ«”|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "«·Œ’„|" & "«· ﬂ·›…|" & "«·≈Ã„«·Ì|"
.ColWidth(0) = 0
.ColWidth(1) = 1200
.ColWidth(2) = 2000
.ColWidth(3) = 4000
.ColWidth(4) = 900
.ColWidth(5) = 900
.ColWidth(6) = 900
.ColWidth(7) = 900
.ColWidth(8) = 900
.ColWidth(9) = 900
.ColHidden(9) = True
.ColHidden(11) = True
.ColHidden(.Cols - 1) = True
'.ColComboList(1) = "..."
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub ClientLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From " & cFileClient
Generalarray(2) = "Order by file4_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = True

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·„Ê—œ"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search2
Search2.Caption = "«” ⁄·«„"
Search2.Show 1
End Sub

Private Sub xTax_LostFocus()
CalcTotals
End Sub
Private Sub doprint()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For i = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!STR21 = "›« Ê—… „‘ —Ì«  —ﬁ„ : " & Format(xDoc_no.Text)
    temptable!str1 = TurnValue(xCodeDesca.Caption)
    temptable!STR2 = TurnValue(xStore.Text)
    temptable!Str11 = TurnValue(xDate.Text)
    temptable!str3 = TurnValue(grid1.TextMatrix(i, 1))
    temptable!str4 = TurnValue(grid1.TextMatrix(i, 2))
    temptable!val2 = TurnValue(Val(grid1.TextMatrix(i, 3)))
    temptable!VAL1 = TurnValue(Val(grid1.TextMatrix(i, 4)))
    temptable!VAL3 = TurnValue(Val(grid1.TextMatrix(i, 6)))
    temptable!Val11 = Val(retitem(grid1.TextMatrix(i, 1), "Price") & "")
    temptable!Val10 = i
    temptable!VAL4 = Val(xTotalItem.Caption)
    temptable!val5 = Val(xDiscount.Text)
    temptable!Val6 = Val(xTotalDis.Caption)
    temptable!val7 = Val(xTax.Text)
    temptable!val8 = Val(xTotal.Caption)
    temptable!Val10 = Val(grid1.TextMatrix(i, 8))
    temptable!str10 = TurnValue(grid1.TextMatrix(i, 9))
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
REPORT1.ReportFileName = App.Path & "\Reports\purchasewide.rpt"
REPORT1.DataFiles(0) = tempPath
REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Function mysave(Optional bMsg As Boolean = True) As Boolean
If Not MYVALID Then Exit Function
CalcTotals
If Not myreplace Then Exit Function
CardTable.Requery
If bMsg Then Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Find "Doc_No = " & MyParn(xDoc_no.Text), , adSearchForward, adBookmarkFirst
Handlecontrols LoadMode
'xBalance.Caption = Format(GetDesca("Select sum(SAL - pay) as balance FROM " & cFileMove & " WHERE CODE = " & MyParn(xCode.Text)), "fixed")
myload
mysave = True
End Function
Private Sub AddItems(Optional pCode As String = "")
Dim aFields(1, 3)
aFields(0, 0) = "«·ﬂ„Ì…"
aFields(0, 1) = 3
aFields(0, 2) = True
aFields(0, 3) = True

aFields(1, 0) = "«·”⁄—"
aFields(1, 1) = 4
aFields(1, 2) = False
aFields(1, 3) = False

Set itemgrdfrm.oForm = Me
itemgrdfrm.aFields = aFields
itemgrdfrm.xCode.Text = pCode
Set itemgrdfrm.pGrid = grid1
itemgrdfrm.nColItem = 1
itemgrdfrm.xCode_LostFocus
itemgrdfrm.Show 1
End Sub
Sub itemGrdsubAdd(Row)
Dim nRow As Integer
If grid1.TextMatrix(Row, 1) = "" Then Exit Sub
ItemTable.Find "item = " & MyParn(grid1.TextMatrix(Row, 1)), , adSearchForward, adBookmarkFirst
If Not ItemTable.EOF Then
End If
nRow = Row
MakeSerial nRow
End Sub
Sub itemGrdsubDel(Row)
CalcTotals
MakeSerial 1
End Sub
Sub itemGrdsubEdit(Row)
CalcTotals
End Sub
Private Sub xStore_LostFocus()
xStoreDesca.Caption = ""
If Trim(xStore.Text) = "" Then Exit Sub
xStore.Text = GetDesca("Select * from file0_40 where code = " & MyParn(xStore.Text))
End Sub
Private Function MYVALID() As Boolean
If xDoc_no.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If


If xStore.Text = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If

If Not xCode.MatchedWithList Then
    MsgBox "·„ Ì „ «œŒ«· ﬂÊœ «·⁄„Ì·"
    Exit Function
End If

MYVALID = True
End Function

