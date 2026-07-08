VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form BalBox 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6900
   ScaleWidth      =   13650
   Begin VB.Frame Frame8 
      Height          =   780
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   4815
      Width           =   6180
      Begin VB.CommandButton CMDPRINT3 
         Height          =   465
         Left            =   4860
         Picture         =   "BalBox.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   180
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   90
         TabIndex        =   54
         Top             =   135
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   820
         _Version        =   196610
         BackColor       =   13822956
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "BalBox.frx":242A
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   240
         Left            =   2475
         TabIndex        =   55
         Top             =   405
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   423
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   9495
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   3480
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   1365
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo XBOX 
         Height          =   315
         Left            =   135
         TabIndex        =   2
         Top             =   990
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
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
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Œ“‰…"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1050
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„‰ ‰«—ÌŒ"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "’«œ—"
      Height          =   2805
      Left            =   6210
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1485
      Width           =   3390
      Begin VB.TextBox xPart 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xChq_Out 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   2385
         Width           =   1365
      End
      Begin VB.TextBox xTrans_out 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1665
         Width           =   1365
      End
      Begin VB.TextBox xBank_out 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1305
         Width           =   1365
      End
      Begin VB.TextBox xCharges 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2025
         Width           =   1365
      End
      Begin VB.TextBox xPurchase 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   585
         Width           =   1365
      End
      Begin VB.TextBox xCash_out 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   180
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   945
         Width           =   1365
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "’«›Ì Ã«—Ï «·‘—ﬂ«¡ :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   270
         Width           =   1725
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "√Ê—«ﬁ œ›⁄ :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   2475
         Width           =   870
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì  ÕÊÌ·«  „‰ :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   1800
         Width           =   1620
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ìœ«⁄«  »‰ﬂÌ… :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„’«—Ì› :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2115
         Width           =   735
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„‘ —Ì«  ‰ﬁœÌ… :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   1260
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„œ›Ê⁄«  ‰ﬁœÌ… :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1035
         Width           =   1230
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3750
      Top             =   1275
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -585
      Top             =   1935
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin VSFlex7LCtl.VSFlexGrid vsBox 
      Height          =   4695
      Left            =   45
      TabIndex        =   22
      Top             =   90
      Width           =   6150
      _cx             =   10848
      _cy             =   8281
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
      FixedCols       =   1
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ê«—œ"
      Height          =   2805
      Left            =   9630
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1485
      Width           =   3345
      Begin VB.TextBox xBank_in 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   225
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1305
         Width           =   1365
      End
      Begin VB.TextBox xChq_in 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   225
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   2385
         Width           =   1365
      End
      Begin VB.TextBox xCash_in 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   225
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   945
         Width           =   1365
      End
      Begin VB.TextBox xSales 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   225
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   585
         Width           =   1365
      End
      Begin VB.TextBox xIncome 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   225
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2025
         Width           =   1365
      End
      Begin VB.TextBox xFirst_Balance 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   225
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xTrans_In 
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
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   225
         Locked          =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1665
         Width           =   1365
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„”ÕÊ»«  »‰ﬂÌ… :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1395
         Width           =   1305
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ê—«ﬁ ﬁ»÷ :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   2430
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ﬁ»Ê÷«  ‰ﬁœÌ… :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   990
         Width           =   1245
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï  ÕÊÌ· «·Ì :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1755
         Width           =   1545
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«Ì—«œ«  :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   2115
         Width           =   660
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ ”«»ﬁ :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„»Ì⁄«  ‰ﬁœÌ… :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   675
         Width           =   1065
      End
   End
   Begin VB.Frame Frame5 
      Height          =   600
      Left            =   6210
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   4230
      Width           =   3390
      Begin VB.TextBox xOut 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì ’«œ— :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   225
         Width           =   1110
      End
   End
   Begin VB.Frame Frame7 
      Height          =   600
      Left            =   9630
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   4230
      Width           =   3345
      Begin VB.TextBox xIn 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ì Ê«—œ :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.Frame Frame6 
      Height          =   600
      Left            =   8235
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   4860
      Width           =   3345
      Begin VB.TextBox xLast_Balance 
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   180
         Width           =   1365
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—’Ìœ «Œ— «·ÌÊ„ :"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   225
         Width           =   1200
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   6255
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   5490
      Width           =   6810
      Begin VB.CommandButton Cmd_Ok 
         Caption         =   "⁄—÷"
         Height          =   465
         Left            =   5535
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   180
         Width           =   1185
      End
      Begin VB.CommandButton Cmd_Exit 
         Caption         =   "Œ‹‹‹—ÊÃ"
         Height          =   465
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdPrint2 
         Caption         =   " ›’Ì·Ì «·Õ—ﬂ…"
         Height          =   465
         Left            =   990
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   180
         Width           =   1500
      End
      Begin VB.CommandButton Cmd_Print 
         Caption         =   "ÿ»«⁄… «·—’Ìœ"
         Height          =   465
         Left            =   4230
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton Cmd_Print2 
         Caption         =   "≈Ã„«·Ì Õ—ﬂ… ÌÊ„Ì…"
         Height          =   465
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   1680
      End
   End
End
Attribute VB_Name = "BalBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub cmd_excel_Click()
    ToFileExel2 vsBox, , , , , 1.1, , , , , , Me
End Sub

Private Sub CMD_EXIT_Click()
    Unload Me
End Sub
Private Sub CMD_OK_Click()
Dim sourcetable As New ADODB.Recordset, nBalance As Double
If xBox.BoundText = "" Then
    MsgBox "«œŒ· «·Œ“«‰…"
    Exit Sub
End If

If IsDate(xDate1.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " DATE < " & DateSq(xDate1.Text)
Else
    cWhere = cWhere & turn(cWhere, " and ") & " DATE IS NULL "
End If

If xBox.BoundText <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & " box = " & MyParn(xBox.BoundText)
End If

'--------------  Ê«—œ
cField1 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove " & _
           turn(cWhere) & cWhere & _
           ") as First_Balance"

cField2 = myiif( _
        "( FLAG = 0)", "PLUS - MINUS") & _
        " As First_Bal"
                                  
cField3 = myiif( _
        " (FLAG = 9 or Flag = 10 )", "PLUS - MINUS") & _
        " As Sales"

cField4 = myiif( _
        " (FLAG = 1 or Flag = 4 )", "PLUS - MINUS") & _
        " As Cash_In"

cField5 = myiif( _
        " (FLAG = 14)", "PLUS") & _
        " As Bank_In"

cField6 = myiif( _
        " (FLAG = 8)", "PLUS") & _
        " As Trans_In"

cField7 = myiif( _
        " (FLAG = 6)", "PLUS ") & _
        " As Income"

cField8 = myiif( _
        " (FLAG = 13)", "PLUS") & _
        " As Chq_In"

' ----------- ’«œ—
cField9 = myiif( _
        " (FLAG = 11 or flag = 12)", "MINUS - PLUS") & _
        " As Purchase"

cField10 = myiif( _
        " (FLAG = 2 or flag = 3)", "MINUS - PLUS") & _
        " As Cash_out"

cField11 = myiif( _
        " (FLAG = 15)", "MINUS") & _
        " As BANK_OUT"

cField12 = myiif( _
        " (FLAG = 7)", "MINUS") & _
        " As TRANS_OUT"

cField13 = myiif( _
        " (FLAG = 5)", "MINUS") & _
        " As CHARGES"
        
cField14 = myiif( _
        " (FLAG = 16)", "MINUS") & _
        " As CHQ_OUT"
        
cField15 = myiif( _
        " (FLAG = 17)", "MINUS") & _
        " As PART"
        
cWhere = ""
If IsDate(xdate2.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " DATE <= " & DateSq(xdate2.Text)
End If

If xBox.BoundText <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & " box = " & MyParn(xBox.BoundText)
End If

cField16 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove " & _
           turn(cWhere) & cWhere & _
           ") as Last_Balance"

cString = "Select " & cField1 & "," & cField2 & "," & cField3 & "," & cField4 & "," & cField5 & "," & _
           cField6 & "," & cField7 & "," & cField8 & "," & cField9 & "," & cField10 & "," & cField11 & "," & cField12 & "," & cField13 & "," & cField14 & "," & cField15 & "," & cField16 & _
           " From BOXMOVE "

cWhere = ""
If IsDate(xDate1.Text) Then cWhere = cWhere & " DATE >= " & DateSq(xDate1.Text)

If IsDate(xdate2.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " DATE <= " & DateSq(xdate2.Text)
End If

If xBox.BoundText <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & " BOX = " & MyParn(xBox.BoundText)
End If

cString = cString & turn(cWhere) & cWhere
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (sourcetable.EOF And sourcetable.BOF) Then
    xFirst_Balance.Text = Format(Val(sourcetable!First_Balance & ""), "FIXED")
    xSales.Text = Format(sourcetable!sales, "Fixed")
    xCash_in.Text = Format(sourcetable!Cash_In, "Fixed")
    xChq_in.Text = Format(sourcetable!Chq_in, "Fixed")
    xBank_in.Text = Format(sourcetable!Bank_in, "Fixed")
    xTrans_In.Text = Format(sourcetable!Trans_In, "Fixed")
    xIncome.Text = Format(sourcetable!Income, "Fixed")
    
    xPurchase.Text = Format(sourcetable!purchase, "Fixed")
    xCash_out.Text = Format(sourcetable!Cash_out, "Fixed")
    xBank_out.Text = Format(sourcetable!Bank_Out, "Fixed")
    xTrans_out.Text = Format(sourcetable!TRANS_OUT, "Fixed")
    xCharges.Text = Format(sourcetable!CHARGES, "Fixed")
    xChq_Out.Text = Format(sourcetable!CHQ_OUT, "Fixed")
    XPART.Text = Format(sourcetable!PART, "Fixed")
    
    xIn.Text = Format(Val(xFirst_Balance.Text) + Val(xSales.Text) + Val(xCash_in.Text) + Val(xBank_in.Text) + Val(xTrans_In.Text) + Val(xIncome.Text), "FIXED")
    xOut.Text = Format(Val(xPurchase.Text) + Val(xCash_out.Text) + Val(xTrans_out.Text) + Val(xCharges.Text) + Val(xBank_out.Text) + Val(xChq_Out.Text) + Val(XPART.Text), "FIXED")
    xLast_Balance.Text = Format(Val(sourcetable!Last_Balance & ""), "FIXED")
End If
sourcetable.Close
Set sourcetable = Nothing
myloadgrd
End Sub
Private Sub CMD_PRINT_Click()
Dim temptable As New ADODB.Recordset, aHeader(1)
contemp.Execute "Delete * From Temp"
temptable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
contemp.BeginTrans
If IsDate(xDate1.Text) Or IsDate(xdate2.Text) Then
    aHeader(1) = BetweenString(xDate1.Text, xdate2.Text)
End If
If xBox.BoundText <> "" Then
    aHeader(0) = "«·Œ“‰… : " & xBox.Text
End If

temptable.AddNew
temptable!STR21 = TurnValue(aHeader(0))

temptable!Date1 = DateFix(xDate1.Text)
temptable!date2 = DateFix(xdate2.Text)


temptable!VAL1 = Val(xFirst_Balance.Text)
temptable!val2 = Val(xSales.Text)
temptable!val3 = Val(xCash_in.Text)
temptable!VAL4 = Val(xBank_in.Text)
temptable!val5 = Val(xTrans_In.Text)
temptable!Val6 = Val(xIncome.Text)
temptable!VAL7 = Val(xChq_in.Text)

temptable!Val8 = Val(xPurchase.Text)
temptable!val9 = Val(xCash_out.Text)
temptable!Val10 = Val(xBank_out.Text)
temptable!val11 = Val(xTrans_out.Text)
temptable!val12 = Val(xCharges.Text)
temptable!val13 = Val(xIn.Text)
temptable!val14 = Val(xOut.Text)
temptable!val15 = Val(xLast_Balance.Text)
temptable!STR11 = TurnValue(retHeader(aHeader, 0, 1))
temptable!str12 = TurnValue(retHeader(aHeader, 1, 1))
temptable.Update
contemp.CommitTrans

main.REPORT1.ReportFileName = App.Path & "\Reports\BALBOX.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1

temptable.Close
Set temptable = Nothing
End Sub
Private Sub Cmd_Print2_Click()
    doprint2

End Sub

Private Sub cmdGo_Click()

End Sub

Private Sub CmdPrint2_Click()
Dim sourcetable As New ADODB.Recordset, nBalance As Double
Dim temptable As New ADODB.Recordset
Dim aHeader(2)
contemp.Execute "Delete * From Temp"
temptable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

If xBox.BoundText = "" Then
    MsgBox "«œŒ· «·Œ“«‰…"
    Exit Sub
End If

If IsDate(xDate1.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " DATE < " & DateSq(xDate1.Text)
Else
    cWhere = cWhere & turn(cWhere, " and ") & " DATE IS NULL "
End If

If xBox.BoundText <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & " box = " & MyParn(xBox.BoundText)
End If

cField1 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove " & _
           turn(cWhere) & cWhere & _
           ") as FirstBalance"

cWhere = ""
If IsDate(xdate2.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " DATE <= " & DateSq(xdate2.Text)
End If

If xBox.BoundText <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & " box = " & MyParn(xBox.BoundText)
End If

cString = "Select  BOXMOVE.*," & cField1 & _
           " From boxmove"

cWhere = ""
If IsDate(xDate1.Text) Then
    cWhere = cWhere & " DATE >= " & DateSq(xDate1.Text)
    aHeader(1) = BetweenString(xDate1.Text, xdate2.Text)
End If

If IsDate(xdate2.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " DATE <= " & DateSq(xdate2.Text)
    aHeader(1) = BetweenString(xDate1.Text, xdate2.Text)
End If

If xBox.BoundText <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & " box = " & MyParn(xBox.BoundText)
    aHeader(0) = "—’Ìœ «·Œ“‰… : " & xBox.Text
End If

cString = cString & turn(cWhere) & cWhere & "  ORDER BY DATE,Flag"
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (sourcetable.EOF And sourcetable.EOF) Then
    nBalance = Val(sourcetable!FirstBalance & "")
    If nBalance <> 0 Then
        temptable.AddNew
        temptable!str10 = " ›’Ì·Ì Õ—ﬂ… " & xBox.Text
        temptable!str2 = "—’Ìœ ”«»ﬁ"
        temptable!VAL1 = nBalance
        temptable!val2 = 0
        temptable!val3 = nBalance
        temptable!STR21 = " „‰  «—ÌŒ  " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ  " & Format(xdate2.Text, "DD-MM-YYYY")
        temptable.Update
    End If
End If

With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!Date1 = !Date
    temptable!str1 = !doc_no
    temptable!str2 = TurnValue(!DESCA)
    temptable!str3 = !CodeDesca
    temptable!VAL1 = !plus
    temptable!val2 = !Minus
    nBalance = nBalance + Val(!plus & "") - Val(!Minus & "")
    temptable!val3 = nBalance
    temptable!str10 = " ›’Ì·Ì Õ—ﬂ… " & xBox.Text
    temptable!STR21 = " „‰  «—ÌŒ  " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ  " & Format(xdate2.Text, "DD-MM-YYYY")
    temptable.Update
    sourcetable.MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.REPORT1.ReportFileName = App.Path & "\Reports\Box2.rpt"
    main.REPORT1.DataFiles(0) = tempFile
    main.REPORT1.Action = 1
End If
sourcetable.Close: Set sourcetable = Nothing
temptable.Close: Set temptable = Nothing
End Sub

Private Sub CMDPRINT3_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = xBox.Text
    cHead2 = " „‰  «—ÌŒ  " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ  " & Format(xdate2.Text, "DD-MM-YYYY")
    
    Load PrintGrd
    PrintGrd.doprint vsBox, 1.3, , cHead1, cHead2, , , False, 10
    PrintGrd.Show 1

End Sub

Private Sub Form_Load()
openCon con
Set boxtable = New ADODB.Recordset
If lIsBranchStore Then
    boxtable.Open "SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch), con, adOpenKeyset, adLockOptimistic, adCmdText
Else
    boxtable.Open "FILE0_50", con, adOpenKeyset, adLockOptimistic, adCmdTable
End If
data1.ConnectionString = strCon
If lIsBranchStore Then
    data1.RecordSource = "SELECT CODE , DESCA FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch) & " ORDER BY CODE "
Else
    data1.RecordSource = "SELECT CODE , DESCA FROM FILE0_50 WHERE ISSTOP = 0  ORDER BY CODE "
End If
Set xBox.RowSource = data1
xBox.ListField = "Desca"
xBox.BoundColumn = "Code"

xDate1.Text = Format(Date, "DD-MM-YYYY")
With vsBox
    .Rows = 1
    .Cols = 3
    .TextMatrix(0, 0) = "«·»Ì«‰"
    .TextMatrix(0, 1) = "≈Ã„«·Ì Õ—ﬂ…"
    .TextMatrix(0, 2) = "—’Ìœ «·ÌÊ„"
    .ColWidth(0) = 2500
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub VsBox_Click()
'If vsBox.row <> 0 Then balBoxDtlfrm.Show 1
End Sub
Private Function myValid3() As Boolean
If Not IsDate(xDate1.Text) And Trim(xDate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If Not IsDate(xdate2.Text) And Trim(xdate2.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
If Trim(xBox.BoundText) = "" Then
    MsgBox "«œŒ· «·Œ“«‰…"
    Exit Function
End If
myValid3 = True
End Function
Private Sub myloadgrd()
Dim sourcetable As New ADODB.Recordset, nBalance As Double
cStr1 = "     SELECT        SUM(PLUS) AS V1 , SUM(MINUS) AS V2 , DESCA FROM            BOXMOVE WHERE BOX = " & MyParn(xBox.BoundText)
If IsDate(xDate1.Text) Then cStr1 = cStr1 & " AND DATE >= " & DateSq(xDate1.Text)
If IsDate(xdate2.Text) Then cStr1 = cStr1 & " AND DATE <= " & DateSq(xdate2.Text)
cStr1 = cStr1 & " GROUP BY FLAG, DESCA ORDER BY FLAG"
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText

vsBox.Rows = 1
Do Until sourcetable.EOF
        vsBox.AddItem ""
        vsBox.TextMatrix(vsBox.Rows - 1, 0) = sourcetable!DESCA & ""
        vsBox.TextMatrix(vsBox.Rows - 1, 1) = Format(sourcetable!V1, "fixed")
        vsBox.TextMatrix(vsBox.Rows - 1, 2) = Format(sourcetable!V2, "fixed")
    sourcetable.MoveNext
Loop
vsBox.SubtotalPosition = flexSTBelow
vsBox.Subtotal flexSTSum, -1, 1, "#,00", vbYellow, , True
vsBox.Subtotal flexSTSum, -1, 2, "#,00", vbYellow, , True
sourcetable.Close
Set sourcetable = Nothing
End Sub
Private Sub doprint2()
Dim sourcetable As New ADODB.Recordset, nBalance As Double
Dim temptable As New ADODB.Recordset
Dim aHeader(2)
contemp.Execute "Delete * From Temp"
temptable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

If xBox.BoundText = "" Then
    MsgBox "«œŒ· «·Œ“«‰…"
    Exit Sub
End If

If IsDate(xDate1.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " DATE < " & DateSq(xDate1.Text)
End If

If xBox.BoundText <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & " box = " & MyParn(xBox.BoundText)
End If

'--------------  Ê«—œ
cField1 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove " & _
           turn(cWhere) & cWhere & _
           ") as First_Balance"
cField2 = myiif( _
        "( FLAG = 0)", "PLUS - MINUS") & _
        " As First_Bal"
                                  
cField3 = myiif( _
        " (FLAG = 9 or Flag = 10 )", "PLUS - MINUS") & _
        " As Sales"

cField4 = myiif( _
        " (FLAG = 1 or Flag = 4 )", "PLUS - MINUS") & _
        " As Cash_In"

cField5 = myiif( _
        " (FLAG = 14)", "PLUS") & _
        " As Bank_In"

cField6 = myiif( _
        " (FLAG = 8)", "PLUS") & _
        " As Trans_In"

cField7 = myiif( _
        " (FLAG = 6)", "PLUS") & _
        " As Income"

cField8 = myiif( _
        " (FLAG = 13)", "PLUS") & _
        " As Chq_In"


' ----------- ’«œ—
cField9 = myiif( _
        " (FLAG = 11 or flag = 12)", "MINUS - PLUS") & _
        " As Purchase"

cField10 = myiif( _
        " (FLAG = 2 or flag = 3)", "MINUS - PLUS") & _
        " As Cash_out"

cField11 = myiif( _
        " (FLAG = 15 OR FLAG = 17 )", "MINUS") & _
        " As BANK_OUT"

cField12 = myiif( _
        " (FLAG = 9)", "MINUS") & _
        " As TRANS_OUT"

cField13 = myiif( _
        " (FLAG = 5)", "MINUS") & _
        " As CHARGES"
        
cField14 = myiif( _
        "", "PLUS - MINUS") & _
        " As BalanceLastDay"


' ·÷»ÿ «Œ— Õﬁ·
    cWhere = ""
    If IsDate(xdate2.Text) Then
        cWhere = cWhere & turn(cWhere, " and ") & " DATE <= " & DateSq(xdate2.Text)
    End If
    
    If xBox.BoundText <> "" Then
        cWhere = cWhere & turn(cWhere, " and ") & " box = " & MyParn(xBox.BoundText)
    End If
            
    cField15 = "(" & _
               "Select Sum(PLUS - MINUS) From BoxMove " & _
               turn(cWhere) & cWhere & _
               ") as Last_Balance"
' «‰ Â«¡ «Œ— Õﬁ·

cString = "Select Date," & cField1 & "," & cField2 & "," & cField3 & "," & cField4 & "," & cField5 & "," & _
           cField6 & "," & cField7 & "," & cField8 & "," & cField9 & "," & cField10 & "," & cField11 & "," & cField12 & "," & cField13 & "," & cField14 & "," & cField15 & _
           " From boxmove "
cWhere = ""
If IsDate(xDate1.Text) Then
    cWhere = cWhere & " DATE >= " & DateSq(xDate1.Text)
    aHeader(1) = BetweenString(xDate1.Text, xdate2.Text)
End If

If IsDate(xdate2.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " DATE <= " & DateSq(xdate2.Text)
    aHeader(1) = BetweenString(xDate1.Text, xdate2.Text)
End If

If xBox.BoundText <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & " box = " & MyParn(xBox.BoundText)
    aHeader(0) = "—’Ìœ «·Œ“‰… : " & xBox.Text
End If


cString = cString & turn(cWhere) & cWhere & " Group by Date ORDER BY DATE"

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

If Not (sourcetable.EOF And sourcetable.BOF) Then
    nBalance = Val(sourcetable!First_Balance & "") + Val(sourcetable!First_Bal & "")
    nFirst_Balance = nBalance
End If

Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = "≈Ã„«·Ì Õ—ﬂ… ÌÊ„Ì… " & xBox.Text
    temptable!VAL1 = nBalance
    temptable!val14 = nFirst_Balance
    
    temptable!val2 = sourcetable!sales
    temptable!val3 = sourcetable!Trans_In
    
    temptable!VAL4 = sourcetable!Income + sourcetable!Chq_in + sourcetable!Cash_In
    
    
    temptable!VAL7 = sourcetable!TRANS_OUT
    temptable!Val8 = sourcetable!CHARGES
    
    temptable!Val6 = sourcetable!Cash_out + sourcetable!Bank_Out
    temptable!val9 = sourcetable!Bank_Out
    temptable!val12 = sourcetable!BalanceLastDay + nBalance
    
    temptable!val13 = Val(sourcetable!Last_Balance & "")
    nBalance = Val(temptable!val12 & "")
    
    temptable!Date1 = sourcetable!Date
    temptable!Val21 = nRecord
    temptable!STR21 = retHeader(aHeader, 0, 2)
    nRecord = nRecord + 1
    temptable.Update
    sourcetable.MoveNext
Loop

contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\BALBOX2.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
sourcetable.Close
Set sourcetable = Nothing
temptable.Close
Set temptable = Nothing
End Sub
