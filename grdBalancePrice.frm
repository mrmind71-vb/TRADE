VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdBalancePrice 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ì «—»«Õ ‘Â—Ì… ··„Œ«“‰"
   ClientHeight    =   10290
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame fmBranch 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   34
      Top             =   405
      Width           =   3885
      Begin VB.OptionButton optStop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›—Ê⁄ „ Êﬁ›…"
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
         Height          =   270
         Index           =   2
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   225
         Width           =   1365
      End
      Begin VB.OptionButton optStop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›—⁄ Õ«·Ì…"
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
         Height          =   270
         Index           =   1
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.OptionButton optStop 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
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
         Height          =   270
         Index           =   3
         Left            =   3105
         RightToLeft     =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   225
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   600
      Left            =   6345
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   405
      Width           =   3795
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
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
         Height          =   270
         Index           =   0
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   225
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›—Ê⁄ ›ﬁÿ"
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
         Height          =   270
         Index           =   1
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   225
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " ÊﬂÌ·«  ›ﬁÿ"
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
         Height          =   270
         Index           =   2
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   990
      Width           =   7710
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdBalancePrice.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1305
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
         BackColor       =   16777215
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
         Picture         =   "grdBalancePrice.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdBalancePrice.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   6420
         TabIndex        =   10
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "grdBalancePrice.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2565
         TabIndex        =   12
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "grdBalancePrice.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPdf 
         Height          =   555
         Left            =   3870
         TabIndex        =   11
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "grdBalancePrice.frx":BBFA
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSql 
         Height          =   555
         Left            =   5145
         TabIndex        =   33
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
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
         Picture         =   "grdBalancePrice.frx":E0A7
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1680
      Left            =   10215
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   45
      Width           =   10095
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   4860
         TabIndex        =   3
         Top             =   1260
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   330
         Left            =   4860
         TabIndex        =   2
         Top             =   900
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   330
         Left            =   4860
         TabIndex        =   1
         Top             =   555
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   330
         Left            =   135
         TabIndex        =   4
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   900
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   330
         Left            =   135
         TabIndex        =   6
         Top             =   1260
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï"
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
         Height          =   270
         Left            =   8685
         TabIndex        =   25
         Top             =   225
         Width           =   300
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "„Ê”„ «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   585
         Width           =   1230
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "«·›—⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   945
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "«·ﬁ”„ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Ê—œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8625
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   585
         Width           =   645
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„Ã„Ê⁄… „ÊœÌ·« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   8625
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1305
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„’‰⁄"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8625
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   930
         Width           =   690
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   4770
      Top             =   -135
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   6495
      Index           =   1
      Left            =   45
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2340
      Width           =   20220
      _cx             =   35666
      _cy             =   11456
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   945
      Top             =   1530
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   180
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   90
      Top             =   1620
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   720
      Top             =   1125
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA6 
      Height          =   330
      Left            =   675
      Top             =   675
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA7 
      Height          =   330
      Left            =   1125
      Top             =   765
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc data8 
      Height          =   330
      Left            =   135
      Top             =   495
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc data9 
      Height          =   330
      Left            =   90
      Top             =   45
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   26
      Top             =   9825
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   27
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   4095
         TabIndex        =   28
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   8100
         TabIndex        =   29
         Top             =   45
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   3
         Left            =   12150
         TabIndex        =   30
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   4
         Left            =   16155
         TabIndex        =   31
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   32
      Top             =   9630
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   2
      Left            =   14085
      TabIndex        =   38
      Top             =   1800
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   900
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ﬁÌÌ„ —’Ìœ «·„Ã„Ê⁄« "
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   3
      Left            =   10980
      TabIndex        =   39
      Top             =   1800
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ﬁÌÌ„ —’Ìœ «·„Ê—œÌ‰ "
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   5
      Left            =   4770
      TabIndex        =   40
      Top             =   1800
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ﬁÌÌ„ —’Ìœ «·„ÊœÌ·« "
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   4
      Left            =   7875
      TabIndex        =   41
      Top             =   1800
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ﬁÌÌ„ —’Ìœ «·›—Ê⁄"
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6495
      Index           =   2
      Left            =   45
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2340
      Width           =   20220
      _cx             =   35666
      _cy             =   11456
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6495
      Index           =   3
      Left            =   45
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2340
      Width           =   20220
      _cx             =   35666
      _cy             =   11456
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6495
      Index           =   4
      Left            =   45
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2340
      Width           =   20220
      _cx             =   35666
      _cy             =   11456
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6495
      Index           =   5
      Left            =   45
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2340
      Width           =   20220
      _cx             =   35666
      _cy             =   11456
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   12
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Index           =   1
      Left            =   1305
      Top             =   90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Index           =   2
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Index           =   3
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Index           =   4
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Index           =   5
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin Threed.SSRibbon cmdRep 
      Height          =   510
      Index           =   1
      Left            =   17190
      TabIndex        =   46
      Top             =   1800
      Width           =   3060
      _ExtentX        =   5398
      _ExtentY        =   900
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   " ﬁÌÌ„ —’Ìœ «·«’‰«›"
      PictureAlignment=   1
      BevelWidth      =   3
      GroupAllowAllUp =   0   'False
   End
End
Attribute VB_Name = "grdBalancePrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim FieldTable As ADODB.Recordset
Dim cString As String, aHeader()
Private Sub CMD_EXEL_Click()
End Sub
Private Sub CMD_PRINT_Click()
End Sub
Private Sub cmdExcel_Click()
Me.MousePointer = 11
Dim I As Long
For I = 1 To cmdRep.UBound
    Dim aRow As Variant
    If cmdRep(I).Value = -1 Then
        If grid1(I).Rows > 1 Then
            aSub = AddFlag(Empty, "row", 1)
            aSub = AddFlag(aSub, "bold", True)
            aSub = AddFlag(aSub, "word_wrap", False)
            aSub = AddFlag(aSub, "back_color", 19)
            aRow = AddFlag(aRow, aSub)
        End If
        ToFileExelNew grid1(I), , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption & "-" & cmdRep(I).Caption, retHeader(aHeader, 0, 6), retHeader(aHeader, 4, 4), retHeader(aHeader, 8, 4))
    End If
Next
Me.MousePointer = 0
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
If Not MYVALID Then Exit Sub
Me.MousePointer = vbHourglass
myloadgrd
Me.MousePointer = vbNormal
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate1.text) Then
    MsgBox " «—ÌŒ «·„œ… «·«Ê· €Ì— „Õœœ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub cmdPdf_Click()
myPrint "1", True
End Sub
Private Sub cmdPrint_Click()
myPrint
End Sub

Private Sub cmdRep_Click(Index As Integer, Value As Integer)
Handlecontrols
End Sub
Private Sub Handlecontrols()
Dim I As Long
For I = 1 To cmdRep.UBound
    cmdRep(I).BackColor = IIf(cmdRep(I).Value = -1, &HC0C0C0, &HFFFFFF)
    
    grid1(I).Visible = cmdRep(I).Value = -1
Next
End Sub
Private Sub cmdSql_Click()
myloadgrd True
End Sub

Private Sub Form_Load()
openCon con

Set DATA7.Recordset = myCmd("Select Code,DescA From STORE_ALL WHERE " & cUserStore, con)
Set xStore.RowSource = DATA7
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
If cBranch <> "00" Then
    xStore.BoundText = cBranchStore
    xStore.Enabled = False
End If

Set data1.Recordset = myCmd("Select Code,DescA From File1_10SC ORDER BY code ", con)
Set XSECTION.RowSource = data1
XSECTION.ListField = "Desca"
XSECTION.BoundColumn = "Code"

Set data1.Recordset = myCmd("SELECT CODE,DESCA FROM FILE0_40 UNION ALL SELECT CODE,DESCA FROM BRANCH WHERE STORE IS NULL  " & cWhere, con)
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

Set DATA2.Recordset = myCmd("Select Code,DescA From File4_10 order by Desca", con)
Set xSupp.RowSource = DATA2
xSupp.ListField = "Desca"
xSupp.BoundColumn = "Code"

Set DATA3.Recordset = myCmd("Select Code,DescA From File1_50 ORDER BY DESCA", con)
Set xGroup.RowSource = DATA3
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set DATA4.Recordset = myCmd("Select mosm ,descA From mosm ORDER BY date DESC ", con)
Set xMosm.RowSource = DATA4
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"
xMosm.BoundText = cPMosm


Set DATA5.Recordset = myCmd("Select code ,desca From fact ORDER BY DESCA ", con)
Set xFact.RowSource = DATA5
xFact.ListField = "Desca"
xFact.BoundColumn = "Code"


Set DATA7.Recordset = myCmd("Select Code,DescA From STORE_ALL WHERE " & cUserStore, con)
Set xStore.RowSource = DATA7
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
If cBranch <> "00" Then
    xStore.BoundText = cBranchStore
    xStore.Enabled = False
End If

Dim I As Long
For I = 1 To grid1.UBound
    Set grid1(I).DataSource = DATA11(I)
    Fixgrd (I)
Next

grid1(1).Tag = "rp.sp_balance_price_section"
grid1(2).Tag = "rp.sp_balance_price_group"
grid1(3).Tag = "rp.sp_balance_price_sup"
grid1(4).Tag = "rp.sp_balance_price_branch"
grid1(5).Tag = "rp.sp_balance_price_model"


LoadText Me

fmBranch.Visible = lSupperVisor
optStop(1).Value = True

fixControls
cmdRep(1).Value = -1

End Sub
Private Sub myloadgrd(Optional bString As Boolean = False)
ReDim aHeader(13)

Dim aPrm As Variant
If IsDate(xDate1.text) Then
    aPrm = AddFlag(aPrm2, "DATE", addDate(xDate1.text))
    aHeader(0) = " «—ÌŒ Õ Ì : " & xDate1.text
End If


If XSECTION.MatchedWithList Then
    aPrm = AddFlag(aPrm, "SECTION", addvalue(XSECTION.BoundText))
    aHeader(1) = "«·ﬁ”„ : " & XSECTION.text
End If

If xSupp.MatchedWithList Then
    aPrm = AddFlag(aPrm, "SUP", addstring(xSupp.BoundText))
    aHeader(2) = "«·„Ê—œ : " & xSupp.text
End If

If xGroup.MatchedWithList Then
    aPrm = AddFlag(aPrm, "GROUP", addstring(xGroup.BoundText))
    aHeader(3) = "„Ã„Ê⁄… «·’‰› : " & xGroup.text
End If

If xFact.MatchedWithList Then
    aPrm = AddFlag(aPrm, "FACT", addstring(xFact.BoundText))
    aHeader(4) = "«·„’‰⁄ : " & xFact.text
End If

If xMosm.MatchedWithList Then
    aPrm = AddFlag(aPrm, "MOSM", addstring(xMosm.BoundText))
    aHeader(5) = "«·„Ê”„ : " & xMosm.text
End If

If xStore.MatchedWithList Then
    aPrm = AddFlag(aPrm, "STORE", addstring(xStore.BoundText))
    aHeader(6) = "„Œ“‰ : " & xStore.text
End If


If Option1(1).Value Or Option1(2).Value Then
    aPrm = AddFlag(aPrm, "TYPE", IIf(Option1(1).Value, "1", "0"))
    aHeader(7) = IIf(Option1(1).Value, Option1(1).Caption, Option1(2).Caption)
End If

If optStop(1).Value Or optStop(2).Value Then
    aPrm = AddFlag(aPrm, "ISSTOP", IIf(optStop(1).Value, "0", "1"))
    aHeader(8) = IIf(optStop(1).Value, optStop(1).Caption, optStop(2).Caption)
End If

If Not bOpt5 Then
    aPrm = AddFlag(aPrm, "usercode", TurnValue(nusercode))
End If

Dim cString As String
Dim I As Long
If Not bString Then
    For I = 1 To cmdRep.UBound
        cString = myPrcString(grid1(I).Tag, aPrm)
        If cmdRep(I).Value = -1 Then
            Set DATA11(I).Recordset = myCmd(cString, con, , , 600)
            Fixgrd I
            panel1(0).Caption = "⁄œœ «·”Ã·«  : " & grid1(I).Rows - 1
        Else
            If DATA11(I).RecordSource <> cString Then grid1(I).Rows = 1
        End If
    Next
Else
    Clipboard.Clear
    For I = 1 To cmdRep.UBound
        If cmdRep(I).Value = -1 Then
            Clipboard.SetText myPrcString(grid1(I).Tag, aPrm)
        End If
    Next
End If
End Sub
Sub Fixgrd(Index As Long)
With grid1(Index)
If Index = 5 Then
    .RowHeight(0) = 700
    .TextMatrix(0, 0) = "„ÊœÌ·"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "„ÊœÌ· "
    .TextMatrix(0, 4) = "«·’‰›"
    
    .TextMatrix(0, 5) = "«·—’Ìœ"
    .TextMatrix(0, 6) = "”⁄— „’‰⁄"
    .TextMatrix(0, 7) = "”⁄— «· ﬂ·›…"
    .TextMatrix(0, 8) = "”⁄— «·»Ì⁄"
    .TextMatrix(0, 9) = " ﬁÌ„ »”⁄— „’‰⁄"
    .TextMatrix(0, 10) = " ﬁÌ„ »”⁄— «· ﬂ·›…"
    .TextMatrix(0, 11) = " ﬁÌ„ »”⁄— «·»Ì⁄"
    
    .ColHidden(0) = True
    .ColHidden(6) = Not bOpt10
    .ColHidden(7) = Not bOpt10
    .ColHidden(10) = Not bOpt10
    .ColHidden(11) = Not bOpt10
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1500
    .ColWidth(2) = 1000
    .ColWidth(3) = 1200
    .ColWidth(4) = 3200
    
    .ColWidth(5) = 1200
    .ColWidth(6) = 900
    .ColWidth(7) = 900
    .ColWidth(8) = 900
    .ColWidth(9) = 1400
    .ColWidth(10) = 1400
    .ColWidth(11) = 1400
        
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTAbove
    
    .Subtotal flexSTSum, -1, 5, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    For Col = 9 To .Cols - 1
        .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    Next
Else
    .RowHeight(0) = 700
    .TextMatrix(0, 0) = "ﬂÊœ"
    If Index = 1 Then
        .TextMatrix(0, 1) = "«·ﬁ”„"
    ElseIf Index = 2 Then
        .TextMatrix(0, 1) = "«·„Ã„Ê⁄…"
    ElseIf Index = 3 Then
        .TextMatrix(0, 1) = "«·„Ê—œ"
    ElseIf Index = 4 Then
        .TextMatrix(0, 1) = "«·›—⁄"
    End If
    .TextMatrix(0, 2) = "«·—’Ìœ "
    .TextMatrix(0, 3) = " ﬁÌ„ »”⁄— „’‰⁄"
    .TextMatrix(0, 4) = " ﬁÌ„ »”⁄— «· ﬂ·›…"
    .TextMatrix(0, 5) = " ﬁÌ„ »”⁄— «·»Ì⁄"
    
    .ColHidden(3) = Not bOpt10
    .ColHidden(4) = Not bOpt10
    
    
    .ColWidth(0) = 600
    .ColWidth(1) = 3000
    .ColWidth(2) = 1500
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
    .ColWidth(5) = 2000
    
    For Col = 2 To .Cols - 1
        .ColFormat(Col) = "##,##"
        .ColDataType(Col) = flexDTDouble
    Next
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTAbove
    
    For Col = 2 To .Cols - 1
        .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    Next
End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdBalancePrice = Nothing
End Sub

Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
Dim aRow As Variant, aSub As Variant

Dim I As Long
For I = 1 To cmdRep.UBound
    If cmdRep(I).Value = -1 Then
        aSub = AddFlag(Empty, "row", 1)
        aSub = AddFlag(aSub, "col", 0)
        aSub = AddFlag(aSub, "cols", 2)
        aSub = AddFlag(aSub, "text", "≈Ã„«·Ì")
        aRow = AddFlag(aRow, aSub)

        If grid1(I).Rows < 2 Then Exit Sub
        
        Set PrintGrdNew.myform = Me
        PrintGrdNew.DOPRINT grid1(I), 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption & " - " & cmdRep(I).Caption, ArbString(retHeader(aHeader, 0, 5)), , False, I = 5, 10, , aRow, Array(1)
    End If
Next

If Not bIgPreview Then
    PrintGrdNew.Show 1
Else
    Unload PrintGrdNew
End If
End Sub


Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub XMOSM2_GotFocus()
myGotFocus XMOSM2
End Sub
Private Sub XMOSM2_LostFocus()
myLostFocus XMOSM2
End Sub
Private Sub xModelFACT_GotFocus()
myGotFocus xModelFact
End Sub
Private Sub XModelFACT_LostFocus()
myLostFocus xModelFact
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub xFact_GotFocus()
myGotFocus xFact
End Sub
Private Sub xFACT_LostFocus()
myLostFocus xFact
If Not xFact.MatchedWithList Then xFact.BoundText = ""
End Sub
Private Sub xSupp_GotFocus()
myGotFocus xSupp
End Sub
Private Sub xSupp_LostFocus()
myLostFocus xSupp
If Not xSupp.MatchedWithList Then xSupp.BoundText = ""
End Sub
Private Sub xage_GotFocus()
myGotFocus xage
End Sub
Private Sub xage_LostFocus()
myLostFocus xage
If Not xage.MatchedWithList Then xage.BoundText = ""
End Sub
Private Sub xsex_GotFocus()
myGotFocus xsex
End Sub
Private Sub xsex_LostFocus()
myLostFocus xsex
If Not xsex.MatchedWithList Then xsex.BoundText = ""
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xMosm
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xMosm
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xSection_GotFocus()
myGotFocus XSECTION
End Sub
Private Sub xSection_LostFocus()
myLostFocus XSECTION
If Not XSECTION.MatchedWithList Then XSECTION.BoundText = ""
End Sub
Private Sub xmosm20_GotFocus()
myGotFocus xmosm20
End Sub
Private Sub xmosm20_LostFocus()
myLostFocus xmosm20
If Not xmosm20.MatchedWithList Then xmosm20.BoundText = ""
End Sub
Private Sub fixControls()
If Dir(App.Path & "\sys_img\selected.jpg") = "" Then Exit Sub
Dim I As Long
For I = 1 To cmdRep.UBound
    cmdRep(I).PictureDn = LoadPicture(App.Path & "\sys_img\selected.jpg")
Next
End Sub
