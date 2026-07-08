VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form VsOrder_Rec 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ—Ì— ÿ·»Ì«  - «” ·«„ - „»Ì⁄«  - —’Ìœ - ‰”»…  Ê—Ìœ  «·√’‰«›  ›’Ì·Ï"
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15120
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1410
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   1125
      Width           =   2715
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   1260
         Picture         =   "VsOrder_Rec.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "VsOrder_Rec.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   32
         Top             =   630
         Width           =   2505
         _ExtentX        =   4419
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
         Picture         =   "VsOrder_Rec.frx":495E
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   240
         Left            =   90
         TabIndex        =   33
         Top             =   1125
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   423
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame11 
      Height          =   2535
      Left            =   2790
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   0
      Width           =   4785
      Begin ImgXCtrl6.ImgXCtrl imgx1 
         Height          =   2175
         Left            =   90
         TabIndex        =   22
         Top             =   225
         Width           =   4605
         _ExtentX        =   8123
         _ExtentY        =   3836
         BackColor       =   16777215
         BorderStyle     =   0
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "AMR01"
         LicenseRegCode  =   "íß“≠Ω≤•™µΩÆ“Æ®∫´Ω•ÆΩNZOV-HRVEOHJJ-EPOP6gI"
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8025
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2565
      Width           =   20175
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   7665
         Left            =   135
         TabIndex        =   14
         Top             =   225
         Width           =   23460
         _cx             =   41381
         _cy             =   13520
         _ConvInfo       =   1
         Appearance      =   0
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
         RightToLeft     =   0   'False
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
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   9315
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   -45
      Width           =   10905
      Begin VB.TextBox XPO_NO 
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
         Left            =   1800
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1710
         Width           =   1770
      End
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
         Height          =   330
         Left            =   4860
         MaxLength       =   15
         TabIndex        =   24
         Top             =   1665
         Width           =   1545
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
         Height          =   330
         Left            =   7785
         MaxLength       =   15
         TabIndex        =   23
         Top             =   1665
         Width           =   1545
      End
      Begin VB.TextBox xColor 
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
         Left            =   6435
         MaxLength       =   15
         TabIndex        =   17
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox xbarcode 
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
         Left            =   1935
         MaxLength       =   15
         TabIndex        =   15
         Top             =   540
         Width           =   2850
      End
      Begin VB.TextBox xScal 
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
         Left            =   8010
         MaxLength       =   15
         TabIndex        =   11
         Top             =   900
         Width           =   1320
      End
      Begin VB.TextBox xModelFact 
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
         Left            =   180
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1305
         Width           =   3390
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   925
         Width           =   4695
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   4860
         TabIndex        =   2
         Top             =   540
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   6435
         TabIndex        =   3
         Top             =   180
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xcode 
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   180
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   315
         Left            =   4860
         TabIndex        =   18
         Top             =   1305
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No P.O"
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1710
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "To "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   6525
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1710
         Width           =   345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "From Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   9450
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1710
         Width           =   1095
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Supplier"
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   9450
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1350
         Width           =   525
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "BARCODE & ASIN"
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   585
         Width           =   1770
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "SIZE - COLOR"
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
         Left            =   9345
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   945
         Width           =   1515
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "SKU"
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
         Left            =   3630
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1350
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   4
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   945
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   9465
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   960
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Brand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   9465
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   255
         Width           =   645
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10710
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -90
      Top             =   270
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
      Left            =   1188
      Top             =   45
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   891
      Top             =   45
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   297
      Top             =   45
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   0
      Top             =   45
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
   Begin MSAdodcLib.Adodc data6 
      Height          =   330
      Left            =   594
      Top             =   45
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
   Begin MSAdodcLib.Adodc Adodc1 
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
   Begin MSAdodcLib.Adodc DATA7 
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
   Begin MSAdodcLib.Adodc DATA8 
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
   Begin MSAdodcLib.Adodc Adodc2 
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
Attribute VB_Name = "VsOrder_Rec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Variant, sM_Fact As String
Dim oSearch As New Search3, oSearchDoc As New Search3
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    DoDropTable
    DoTable
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    DATA7.ConnectionString = strCon
    DATA7.RecordSource = "Select Code,DescA From File1_10SC ORDER BY DESCA"
    Set xsection.RowSource = DATA7
    xsection.ListField = "Desca"
    xsection.BoundColumn = "Code"
    
    data4.ConnectionString = strCon
    data4.RecordSource = "Select code ,descA From file4_10 ORDER BY DESCA "
    Set xCode.RowSource = data4
    xCode.ListField = "Desca"
    xCode.BoundColumn = "CODE"
    
    data5.ConnectionString = strCon
    data5.RecordSource = "Select code ,desca From fact ORDER BY code "
    Set xFact.RowSource = data5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    xFact.BoundText = sM_Fact
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    
    grid1.Rows = 1
    grid1.Cols = 12
    Fixgrd
    If xpo_no.Text <> "" Then cmdGo_Click
End Sub
Private Sub myload()
    Dim cWhere As String, cFiled1 As String
    xmodelfact.Text = DelZero(xmodelfact.Text)
   '                            0               1               2                   3               4           5               6               7                   8               9               10              11              12              13                                          14                                      15                                      16                                                  17                                                  18                                                      19
    cString = " SELECT   FILE1_10.ITEM, FACT.DESCA , FILE1_10.MODELFACT0, FILE1_10.desca , FILE1_10.COLOR, FILE1_10.SCAL, FILE1_10.COST, FILE1_10.costitem, FILE1_10.PRICE, FILE1_10.PRICE2, FILE1_10.BARCODE, FILE1_10.ENA, FILE4_10.DESCA , ' ',  SUM(T_ORDER_REC_SALES.QUANT) AS QUANT, SUM(T_ORDER_REC_SALES.RECEVD) AS RECEVD, SUM(T_ORDER_REC_SALES.SALES) AS SALES,  COUNT(DISTINCT T_ORDER_REC_SALES.DOC_NO) AS COUNT1, COUNT(DISTINCT T_ORDER_REC_SALES.DOC_REC) AS COUNT2,  COUNT(DISTINCT T_ORDER_REC_SALES.DOC_SALES) AS COUNT3 " & _
                " FROM     FILE1_10 INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE INNER JOIN FILE4_10 ON FILE4_10.CODE = FILE1_10.code INNER JOIN T_ORDER_REC_SALES ON FILE1_10.ITEM = T_ORDER_REC_SALES.ITEM   where file1_10.item is not null "
    If xCode.BoundText <> "" Then cString = cString & " AND FILE1_10.CODE = " & MyParn(xCode.BoundText)
    If xmodelfact.Text <> "" Then cString = cString & " AND FILE1_10.MODELFACT0 = " & MyParn(xmodelfact.Text)
    If xGroup.BoundText <> "" Then cString = cString & " AND  FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xsection.BoundText <> "" Then cString = cString & " AND  FILE1_10.[SECTION] = " & Val(xsection.BoundText)
    If xFact.BoundText <> "" Then cString = cString & " AND  FILE1_10.FACT = " & MyParn(xFact.BoundText)
    
    If IsDate(xDate1.Text) Then cString = cString & " AND  DATE >= " & DateSq(xDate1.Text)
    If IsDate(xdate2.Text) Then cString = cString & " AND  DATE <= " & DateSq(xdate2.Text)
    If xpo_no.Text <> "" Then cString = cString & " AND  DOC_NO = " & MyParn(xpo_no.Text)
    
    If xDesca.Text <> "" Then cString = cString & " AND " & MyParnAnd(xDesca.Text, "file1_10.desca")
    If xbarcode.Text <> "" Then cString = cString & " AND " & MyParnAnd(xbarcode.Text, "file1_10.BARCODE+FILE1_10.ENA")
    If Trim(xScal.Text) <> "" Then cString = cString & " AND  FILE1_10.SCAL = " & MyParn(xScal.Text)
    If Trim(xColor.Text) <> "" Then cString = cString & " AND  " & MyParnAnd(xColor, "FILE1_10.COLOR")
    cString = cString & " GROUP BY FILE1_10.model, FACT.DESCA, FILE1_10.MODELFACT0, FILE1_10.desca, FILE1_10.COLOR, FILE1_10.SCAL, FILE1_10.COST, FILE1_10.costitem, FILE1_10.PRICE,  FILE1_10.PRICE2, FILE1_10.BARCODE, FILE1_10.ENA, FILE4_10.DESCA, FILE1_10.ITEM, FILE1_10.MOSM, FILE1_10.C_SCAL "
    cString = cString & " ORDER BY FILE1_10.model, FACT.DESCA , FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.COLOR, FILE1_10.C_SCAL "
    DATA6.RecordSource = cString
    DATA6.Refresh
    Fixgrd
'    myload2
End Sub
Sub Fixgrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .Cols = 20
    .RowHeight(0) = 1000
    
    .TextMatrix(0, 1) = "Brand"
    .TextMatrix(0, 2) = "SKU"
    .TextMatrix(0, 3) = "Description"
    .TextMatrix(0, 4) = "Color"
    .TextMatrix(0, 5) = "Size"
    .TextMatrix(0, 6) = "PRICE "
''    .TextMatrix(0, 7) = "Net Purchese PRICE"
''    .TextMatrix(0, 8) = "Amazon cost"
''    .TextMatrix(0, 9) = "List Price"
    .TextMatrix(0, 10) = "BARCODE"
    .TextMatrix(0, 11) = "ASIN"
    .TextMatrix(0, 12) = "Supplier"
'    .TextMatrix(0, 13) = "PICT"
    
    .TextMatrix(0, 14) = "Order"
    .TextMatrix(0, 15) = "Recevd"
    .TextMatrix(0, 16) = "Sales"
    
    .TextMatrix(0, 17) = "Count Order"
    .TextMatrix(0, 18) = "Count Recevd"
    .TextMatrix(0, 19) = "Count Sales"
    
    
    .ColHidden(7) = True
    .ColHidden(8) = True
    .ColHidden(9) = True
    .ColHidden(13) = True
    
    .ColHidden(17) = True
    .ColHidden(18) = True
    .ColHidden(19) = True
    
    .ColHidden(0) = True
    
    .ColWidth(1) = 2000
    .ColWidth(2) = 1800
    .ColWidth(3) = 4000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    
    .ColWidth(6) = 800
    .ColWidth(7) = 800
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    
    .ColWidth(10) = 1800
    .ColWidth(11) = 1800
    
    .ColWidth(12) = 3000
    .ColWidth(13) = 0
    
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .ColWidth(16) = 1000
    .ColWidth(17) = 1000
    .ColWidth(18) = 1000
    .ColWidth(19) = 1000
    
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString
    .ColDataType(5) = flexDTString
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 1
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 14, "#0", vbYellow, , True, " "
    .Subtotal flexSTSum, -1, 15, "#0", vbYellow, , True, " "
    .Subtotal flexSTSum, -1, 16, "#0", vbYellow, , True, " "
    If .Rows > 1 Then .Row = 1
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Sub myProc()
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
If ActiveControl.Name = xCode.Name Then
    xCode.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
If ActiveControl.Name = xpo_no.Name Then
    xpo_no.Text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
End If
End Sub
Private Sub grid1_DblClick()
    If IsDate(xDate1.Text) Then ViewMoveOrderItem.pDate1 = DateValue(xDate1.Text)
    If IsDate(xdate2.Text) Then ViewMoveOrderItem.pDate2 = DateValue(xdate2.Text)
    ViewMoveOrderItem.pDoc_No = xpo_no.Text
    ViewMoveOrderItem.pItem = grid1.TextMatrix(grid1.Row, 0)
    ViewMoveOrderItem.Show 1
End Sub
Private Sub grid1_EnterCell()
With grid1
    If grid1.TextMatrix(grid1.Row, 13) <> "" Then LoadPhoto grid1.TextMatrix(grid1.Row, 13) Else ImgX1.Images.Clear
End With
End Sub
Private Sub ImgX1_Click()
If validPhoto(retFile(grid1.TextMatrix(grid1.Row, 13))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(grid1.TextMatrix(grid1.Row, 13))
    oMaxPhoto.cDesca = grid1.TextMatrix(grid1.Row, 3)
    oMaxPhoto.Show 1
End If
End Sub

Private Sub xFACT_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Set Generalarray(0) = Me
Generalarray(1) = "Select code ,DescA From FACT"
Generalarray(2) = "Order by DESCA"
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
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdGo_Click
    End If
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Set Generalarray(0) = Me
Generalarray(1) = "Select code ,DescA From FILE4_10 "
Generalarray(2) = "Order by DESCA"
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
End If
End Sub

Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
ImgX1.Images.Clear
ImgX1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub
Sub DoDropTable()
Dim cString As String
On Error Resume Next
    con.Execute " DROP TABLE T_ORDER_REC_SALES"
    Err.Clear
End Sub
Sub DoTable()
Dim cString As String
On Error Resume Next
    con.Execute "     SELECT        DOC_NO, DATE, ITEM, PRICE, QUANT, RECEVD, SALES, DOC_REC, DATE_REC, DOC_SALES, DATE_SALES Into T_ORDER_REC_SALES From Q_ORDER_REC_SALES"
    cString = "CREATE NONCLUSTERED INDEX [IX_T_ORDER_REC_SALES] ON [dbo].[T_ORDER_REC_SALES] (     [ITEM] ASC,     [DOC_NO] ASC  ) WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]"
    createCommand cString, con
    Err.Clear
End Sub
Private Sub xPO_NO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then OrderNo_LooK
End Sub
Sub OrderNo_LooK()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),NOTES " & _
                  " FROM FILE6_51H  "
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If

Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·»Ì«‰"
GrdArray(3, 1) = 2000


searchArray = Array(Generalarray, listarray, GrdArray)
Load oSearchDoc
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub


