VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AddModelGrDisc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ›’Ì·Ï »Ì«‰«  «·„ÊœÌ·« "
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
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   13065
   ScaleWidth      =   21510
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   9855
      Width           =   14910
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   285
         Left            =   90
         TabIndex        =   16
         Top             =   180
         Width           =   14685
         _cx             =   25903
         _cy             =   503
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
   End
   Begin VB.Frame Frame3 
      Height          =   7980
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1845
      Width           =   18240
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   7620
         Left            =   135
         TabIndex        =   14
         Top             =   225
         Width           =   17925
         _cx             =   31618
         _cy             =   13441
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
         AllowBigSelection=   -1  'True
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
   End
   Begin VB.Frame Frame2 
      Height          =   1275
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   495
      Width           =   6540
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   1620
         Picture         =   "AddModelGrDisc.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1860
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   90
         Picture         =   "AddModelGrDisc.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   1410
      End
      Begin Threed.SSCommand cmdAddItem2 
         Height          =   510
         Left            =   3780
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   720
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "AddModelGrDisc.frx":495E
         Caption         =   "√÷«›… «·«’‰«› «·„Œ «—…"
         Alignment       =   4
         ButtonStyle     =   1
         PictureAlignment=   1
         BevelWidth      =   10
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdAddItem1 
         Height          =   555
         Left            =   3780
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   135
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "AddModelGrDisc.frx":7C39
         Caption         =   " ”ÃÌ· ﬂ· «·√’‰«› "
         Alignment       =   4
         ButtonStyle     =   1
         PictureAlignment=   1
         BevelWidth      =   10
         ShapeSize       =   1
      End
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   26
         Top             =   810
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11625
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÊœÌ·«  ·Ì” ·Â« Œ’„ Õ«·Ï"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   990
         Width           =   3210
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "„« ⁄œ«"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   180
         Width           =   1095
      End
      Begin VB.ComboBox xmosm2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   900
         Width           =   1140
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   532
         Width           =   5505
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   7065
         TabIndex        =   2
         Top             =   540
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   7065
         TabIndex        =   3
         Top             =   180
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   3330
         TabIndex        =   4
         Top             =   180
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   315
         Left            =   7065
         TabIndex        =   18
         Top             =   900
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xmodelage 
         Height          =   315
         Left            =   7065
         TabIndex        =   20
         Top             =   1260
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
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
      Begin VB.Label Label2 
         Caption         =   "„Ê”„ œ«Œ·Ï"
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
         Index           =   6
         Left            =   5850
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   945
         Width           =   1140
      End
      Begin VB.Label Label2 
         Caption         =   "«·›∆… «·⁄„—Ì… :"
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
         Index           =   11
         Left            =   10455
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1305
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ”„ : "
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
         Left            =   10455
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   900
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰ ’‰› :"
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
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   585
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… :"
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
         Left            =   10500
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„’‰⁄ :"
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
         Left            =   10500
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   255
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ê”„ :"
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
         Index           =   3
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   780
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   12735
      Width           =   21510
      _ExtentX        =   37941
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            Key             =   ""
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
End
Attribute VB_Name = "AddModelGrDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Variant, sM_Fact As String
Dim oSearch As New Search3
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Private Sub cmdAddItem1_Click()
    If MsgBox(" ≈÷«›… ﬂ· «·„ÊœÌ·«   ", vbYesNo + vbDefaultButton2) = vbYes Then AddModelToDocOkaz (True)
    
End Sub
Private Sub cmdAddItem2_Click()
    If MsgBox(" ≈÷«›… «·„ÊœÌ·«   «·„Õœœ… ›ﬁÿ", vbYesNo + vbDefaultButton2) = vbYes Then AddModelToDocOkaz (False)
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    data4.ConnectionString = strCon
    data4.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC "
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    
    data5.ConnectionString = strCon
    data5.RecordSource = "Select code ,desca From fact ORDER BY code "
    Set xFact.RowSource = data5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    xFact.BoundText = sM_Fact
    Set grid1.DataSource = data6
    data6.ConnectionString = strCon
    
    DATA8.ConnectionString = strCon
    DATA8.RecordSource = "Select Code,DescA From FILE1_10SC ORDER BY DESCA "
    Set XSECTION.RowSource = DATA8
    XSECTION.ListField = "Desca"
    XSECTION.BoundColumn = "Code"
    
    Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_82 ORDER BY DESCA", con)
    Set xmodelage.RowSource = data1
    xmodelage.ListField = "DESCA"
    xmodelage.BoundColumn = "CODE"
    
    Set grid2.DataSource = DATA2
    DATA2.ConnectionString = strCon
    
    xmosm2.AddItem " "
    xmosm2.AddItem "S"
    xmosm2.AddItem "W"
    xmosm2.AddItem "M"
    
    
    grid1.Rows = 1
    grid1.Cols = 8
    
    grid2.Rows = 1
    
    Fixgrd
    If xFact.BoundText <> "" Then cmdGo_Click
End Sub
Private Sub myload()
    Dim cWhere As String, cFiled1 As String, cField2 As String
    cField2 = "DATEDIFF(dd , " & myiif2("type = '2' ", "[date]", "MIN") & " , GETDATE()) AS DATEDIFF"
   '                        0               1                           2               3                   4                       5                   6               7           8                       9
    cString = " SELECT  FILE1_10.MODELNO, FACT.DESCA AS FACTDESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca AS DESCA,  MIN(FILE1_10.PRICE ) , MIN(FILE1_10.PRICE2 ), MIN(FILE1_10.PRICE_2 ) " & _
                " , SUM(CASE WHEN (TYPE = '2' OR TYPE = '7' ) THEN [IN]-[OUT] ELSE 0 END )  , SUM(CASE WHEN (TYPE = '6' ) THEN [OUT]-[IN] ELSE 0 END )  , SUM( [IN]-[OUT] ) , 0 , " & cField2 & _
                " FROM  FILE1_10 INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE   INNER JOIN FILE1_11_ALL ON FILE1_11_ALL.ITEM = FILE1_10.ITEM WHERE FILE1_10.ITEM IS NOT NULL "
    If Check1.Value = 0 Then
        If xMosm.BoundText <> "" Then cString = cString & " AND FILE1_10.MOSM =  " & MyParn(xMosm.BoundText)
    Else
        If xMosm.BoundText <> "" Then cString = cString & " AND FILE1_10.MOSM <> " & MyParn(xMosm.BoundText)
    End If

    If xGroup.BoundText <> "" Then cString = cString & " AND  FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cString = cString & " AND  FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If XSECTION.BoundText <> "" Then cString = cString & " AND  FILE1_10.SECTION = " & Val(XSECTION.BoundText)
    If xDesca.text <> "" Then cString = cString & " AND " & MyParnAnd(xDesca.text, "file1_10.desca")
    If Check2.Value <> 0 Then
         cString = cString & " AND FILE1_10.MODELNO NOT IN (SELECT MODELNO FROM Q_DISC_MODEL ) "
    End If
    cString = cString & " GROUP BY FILE1_10.MODELNO, FACT.DESCA , FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.desca   ORDER BY FACT.DESCA , FILE1_10.MOSM, FILE1_10.MODELFACT0"
    data6.RecordSource = cString
    data6.Refresh
    Fixgrd
End Sub
Sub Fixgrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .Cols = 13
    .RowHeight(0) = 1000
    
    .TextMatrix(0, 0) = "„ÊœÌ·"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "„ÊœÌ·"
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = " ”⁄— «·»Ì⁄"
    .TextMatrix(0, 6) = " ”⁄— √Êﬂ«“ÊÌ‰"
    .TextMatrix(0, 7) = " ”⁄— √Êﬂ«“ÊÌ‰2"
    
    
    .TextMatrix(0, 8) = " „‘ —Ì« "
    .TextMatrix(0, 9) = " „»Ì⁄« "
    .TextMatrix(0, 10) = " —’Ìœ"
    .TextMatrix(0, 11) = " ‰”»… «·»Ì⁄"
    .TextMatrix(0, 12) = " ⁄„— «·„ÊœÌ·"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 2000
    .ColWidth(2) = 800
    .ColWidth(3) = 1500
    .ColWidth(4) = 5000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    .ColWidth(10) = 800
    .ColWidth(11) = 800
    .ColWidth(12) = 800
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString
    
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    
    .ColDataType(8) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    .ColDataType(10) = flexDTDouble
    .ColDataType(11) = flexDTDouble
    .ColDataType(12) = flexDTDouble
    
    For nRow = 1 To .Rows - 1
        nRate = 0
        If Val(.TextMatrix(nRow, 8)) > 0 Then nRate = Round(((Val(.TextMatrix(nRow, 9)) / Val(.TextMatrix(nRow, 8)))) * 100, 2)
        .TextMatrix(nRow, 11) = nRate
    Next nRow
        
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
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
Sub AddModelToDocOkaz(lSelect)
Dim cStr1 As String, nCount As Double
Dim aInsert(1, 1), nId As Double

prog1.Min = 0
prog1.Max = grid1.Rows
prog1.Value = 0

With grid1
    For i = 1 To grid1.Rows - 1
        If lSelect Or .IsSelected(i) Then
            prog1.Value = i
            aInsert(0, 0) = "doc_no"
            aInsert(0, 1) = addstring(ModelOkazin.xdoc_no.text)
            
            aInsert(1, 0) = "MODELNO"
            aInsert(1, 1) = addstring(grid1.TextMatrix(i, 0))
            
             
            nId = Val(GetDesca("SELECT ID FROM FILE0_90 WHERE doc_no = " & MyParn(ModelOkazin.xdoc_no.text) & " AND MODELNO = " & MyParn(.TextMatrix(i, 0)), con) & "")
            If nId = 0 Then
                nCount = nCount + 1
                cStr1 = cStr1 & CreateInsert(aInsert, "FILE0_90") & " ; "
'               con.Execute CreateInsert(aInsert, "FILE0_90")
            Else
                nCount = nCount + 1
                cStr1 = cStr1 & CreateUpdate(aInsert, "FILE0_90", " where ID = " & nId) & " ; "
                con.Execute CreateUpdate(aInsert, "FILE0_90", " where ID = " & nId)
            End If
        End If
        If nCount = 200 Then
            If cStr1 <> "" Then con.Execute cStr1
            cStr1 = ""
            nCount = 0
        End If
    Next i
    If cStr1 <> "" Then con.Execute cStr1
End With
End Sub
