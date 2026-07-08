VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form salesModel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " ”ÃÌ· „»Ì⁄«  „ÊœÌ·« "
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14850
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4860
      TabIndex        =   13
      Top             =   1260
      Width           =   9870
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·—’Ìœ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8685
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   630
         Width           =   675
      End
      Begin VB.Label xSupp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   8055
         TabIndex        =   19
         Top             =   180
         Width           =   510
      End
      Begin VB.Label xFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2160
         TabIndex        =   18
         Top             =   180
         Width           =   2940
      End
      Begin VB.Label xModelFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5130
         TabIndex        =   17
         Top             =   180
         Width           =   2895
      End
      Begin VB.Label xMosm 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         TabIndex        =   16
         Top             =   180
         Width           =   2040
      End
      Begin VB.Label xBalance 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5130
         TabIndex        =   15
         Top             =   585
         Width           =   3435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„ÊœÌ· :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ﬂ„Ì… «·’‰›"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   540
      Width           =   1995
      Begin VB.Label xTotal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   315
         Width           =   1815
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1725
      Left            =   2070
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   630
      Width           =   2760
      Begin Threed.SSCommand cmdSave 
         Height          =   690
         Left            =   45
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   225
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   1217
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
         Picture         =   "salesModel.frx":0000
         Caption         =   " ”ÃÌ· ﬂ„Ì«  «·„ÊœÌ· "
         Alignment       =   4
         ButtonStyle     =   1
         PictureAlignment=   1
         BevelWidth      =   10
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   690
         Left            =   90
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   945
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   1217
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
         Picture         =   "salesModel.frx":2536
         Caption         =   "Œ—ÊÃ "
         Alignment       =   4
         ButtonStyle     =   1
         PictureAlignment=   1
         BevelWidth      =   10
         ShapeSize       =   1
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4440
      Left            =   45
      TabIndex        =   2
      Top             =   2385
      Width           =   14685
      _cx             =   25903
      _cy             =   7832
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   3
      FixedCols       =   3
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
      Editable        =   2
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   300
      Top             =   -540
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
      Caption         =   "Adodc2"
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
      Left            =   -3015
      Top             =   585
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Left            =   -2745
      Top             =   495
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Left            =   -2295
      Top             =   -135
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Index           =   0
      Left            =   -2880
      Top             =   300
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   4860
      TabIndex        =   3
      Top             =   90
      Width           =   9870
      Begin VB.TextBox xitem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6165
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   2445
      End
      Begin VB.TextBox xModel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   6165
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   675
         Width           =   2445
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·»«—ﬂÊœ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8730
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   315
         Width           =   750
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   135
         TabIndex        =   21
         Top             =   675
         Width           =   6000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ÊœÌ· :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8730
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   705
      End
   End
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   -2475
      Top             =   -45
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
      Left            =   -2925
      Top             =   -450
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin VSFlex7Ctl.VSFlexGrid Grid2 
      Height          =   765
      Left            =   -315
      TabIndex        =   5
      Top             =   315
      Visible         =   0   'False
      Width           =   1320
      _cx             =   2328
      _cy             =   1349
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      Rows            =   1
      Cols            =   1
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
      Index           =   1
      Left            =   -2925
      Top             =   405
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Index           =   2
      Left            =   -2925
      Top             =   -450
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Index           =   3
      Left            =   -2925
      Top             =   -450
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin VB.Frame Frame3 
      Caption         =   "—’Ìœ «·’‰›"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1440
      Width           =   1995
      Begin VB.Label xBalance_Item 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   1725
      End
   End
End
Attribute VB_Name = "salesModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Form, nColItem As Integer, nColQuant As Long, strItem As String, STRCODE As String, strStore As String
Public bshowBal As Boolean, bshowcost As Boolean, bAddItem As Boolean
Public bCost1 As Boolean
Dim con As New ADODB.Connection
Public strModel As String
Dim oSearch As New Search3
Dim bUpDate As Boolean
Const LoadMode = 0, DefineMode = 1
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
myForm.myProcAddModel
xitem.SetFocus
mydefine
xitem.text = ""
XMODEL.text = ""
End Sub
Private Sub Form_Activate()
On Error Resume Next
If strItem <> "" Then grid1.SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
Else
    If KeyCode = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then
        KeyCode = 0
        SendKeys "{tAB}"
    End If
End If
End Sub
Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
openCon con
myDefineGrd
Fixgrd
xitem.text = ""
myvalidItem
myload
End Sub
Private Sub VsModel_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim nTot As Double
    With VsModel
        For r = 4 To .Rows - 1
            For c = 2 To .Cols - 1
                nTot = nTot + Val(.TextMatrix(r, c))
            Next c
        Next r
    End With
End Sub
Private Sub VsModel_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
KeyAscii = RetNumber(KeyAscii, False)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 19 Then
    cmdSave_Click
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set salesModel = Nothing
End Sub
Private Function myloadgrd() As Boolean
Dim aRet As Variant, cFieldas As String, cField As String

myDefineGrd

aRet = retFields
If aRet(0) = "" Then Exit Function

cField = aRet(0)
cFieldas = aRet(1)

FillItem cFieldas, cField
FixCost cFieldas, cField
Fixgrd
CalcTotals
myloadgrd = True
End Function
Private Sub Fixgrd()
With grid1
.ColHidden(0) = True
'.ColWidth(0) = 500
.ColWidth(1) = 1300
'.RowHidden(1) = (Not bopt1) And (Not bshowcost)
nColWidth = (.Width - 200 - .ColWidth(0) - .ColWidth(1)) / .Cols
If nColWidth < 500 Then nColWidth = 500
If nColWidth > 1200 Then nColWidth = 1200
For nCol = 2 To .Cols - 1
    .ColWidth(nCol) = nColWidth
    .ColAlignment(nCol) = flexAlignCenterCenter
Next

End With
End Sub
Private Sub myDefineGrd()
grid2.Rows = 3
grid2.Cols = 2

grid1.Rows = 3
grid1.Cols = 2

grid1.MergeCells = flexMergeRestrictRows
grid1.TextMatrix(0, 1) = "«·„ﬁ«”"
grid1.TextMatrix(1, 1) = "”⁄— „’‰⁄"
grid1.TextMatrix(2, 1) = "”⁄— „” Â·ﬂ"

'grid1.RowHidden(1) = Not bCost1
'grid1.RowHidden(2) = Not bCost1
End Sub
Private Sub myloadModel()
'If Trim(xX1.Text) <> "" And Trim(xX2.Text) <> "" And Trim(xX3.Text) <> "" And Trim(xX4.Text) <> "" Then
'    xModel2.caption = UCase(Trim(xX1.Text)) & "/" & UCase(Trim(xX2.Text)) & "/" & UCase(Trim(xX3.Text)) & "/" & UCase(Trim(xX4.Text))
'    xdesca.Caption = lblX1.Caption & Space(1) & LBLX2.Caption & Space(1) & lblX3.Caption & Space(1) & xX4.Text
'    Cmd_Item.Enabled = True
'    myDefineGrd
'    myloadgrd
'Else
'    xModel2.caption = ""
'    xdesca.Caption = ""
'    myDefineGrd
'    Cmd_Item.Enabled = False
'End If
End Sub
Private Sub FixCost(cFieldas, cField)
' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,cost2 from file1_10 WHERE MODEL = " & MyParn(XMODEL.text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(cost2)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grid1.Cols - 1
        grid1.TextMatrix(1, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If

' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,cost2 from file1_10 WHERE MODEL = " & MyParn(XMODEL.text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(cost2)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

loctable.Close
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grid1.Cols - 1
        grid1.TextMatrix(2, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If
loctable.Close
Set loctable = Nothing
End Sub
Private Function retFields()
Dim aRet(1) As String
Dim FieldTable As New ADODB.Recordset
'  ⁄—Ì› «·«⁄„œ…
Dim cString As String, cWhere As String
FieldTable.Open "Select SCAL from file1_10 where model = " & MyParn(XMODEL.text) & " GROUP BY SCAL,C_SCAL order by c_scal", con, adOpenStatic, adLockReadOnly
Do Until FieldTable.EOF
    If Not IsNull(FieldTable!scal) Then
        cFieldas = cFieldas & turn(cField, ",") & "[" & FieldTable!scal & "]" & " as " & "[" & FieldTable!scal & "]"
        cField = cField & turn(cField, ",") & "[" & FieldTable!scal & "]"
    End If
    FieldTable.MoveNext
Loop

aRet(0) = cField
aRet(1) = cFieldas
retFields = aRet
' ⁄œ„ ÊÃÊœ «⁄„œ…
FieldTable.Close
Set FieldTable = Nothing
End Function
Private Sub FillItem(cFieldas, cField)
Dim GrdTable As New ADODB.Recordset, cString As String
' „·∆ «·ÃœÊ·
cString = "Select c_color as [—ﬁ„ «··Ê‰] ,color as [«··Ê‰] " & turn(cFieldas, ",") & cFieldas & _
          " From " & _
          " (Select c_color,Color,scal,item from file1_10 WHERE MODEL = " & MyParn(XMODEL.text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(item)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  " & _
          " order by pvt.C_color"

GrdTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grid1.Cols = GrdTable.Fields.Count: grid2.Cols = GrdTable.Fields.Count

For nCol = 2 To GrdTable.Fields.Count - 1
    grid1.TextMatrix(0, nCol) = GrdTable.Fields(nCol).Name
Next

Do Until GrdTable.EOF
    grid2.AddItem ""
    grid1.AddItem ""
    For nCol = 0 To GrdTable.Fields.Count - 1
        If nCol <= 1 Then
            grid1.TextMatrix(grid2.Rows - 1, nCol) = GrdTable.Fields(nCol).Value & ""
        Else
            grid2.TextMatrix(grid2.Rows - 1, nCol) = GrdTable.Fields(nCol).Value & ""
            nFoundRow = myForm.grid1.FindRow(GrdTable.Fields(nCol).Value & "", , 1)
            If nFoundRow <> -1 Then
                 grid1.TextMatrix(grid1.Rows - 1, nCol) = Myvalue(myForm.grid1.TextMatrix(nFoundRow, nColQuant))
             Else
                 grid1.TextMatrix(grid1.Rows - 1, nCol) = ""
             End If
        End If
    Next
    GrdTable.MoveNext
Loop
GrdTable.Close
Set GrdTable = Nothing
End Sub
Private Sub mydefine()
xModelFact.Caption = ""
xDesca.Caption = ""
xSupp.Caption = ""
xFact.Caption = ""
xMosm.Caption = ""
xBalance.Caption = ""
xDesca.Caption = ""
myDefineGrd
End Sub
Private Sub Fixgrdx(pGrid)
pGrid.ColHidden(0) = True
pGrid.ColWidth(1) = 1500
End Sub

Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
CalcTotals
End Sub
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    If grid1.Col < grid1.Cols - 1 Then grid1.Col = grid1.Col + 1
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
   KeyCode = 0
   If Col < grid1.Cols - 1 Then grid1.Col = Col + 1
End If
End Sub

Private Sub grid1_EnterCell()
If bshowBal And grid1.Cols > 2 And grid1.Rows > 3 And strStore <> "" Then
    xBalance_Item.Caption = LastBalance(grid2.TextMatrix(grid1.Row, grid1.Col), strStore, con)
End If
With grid1
    If .Col > 2 Then
        If .TextMatrix(.Row, .Col) = "" Then .TextMatrix(.Row, .Col) = .TextMatrix(.Row, .Col - 1)
    End If
End With
End Sub
Private Sub Grid1_GotFocus()
With grid1
    If .Cols > 2 And .Rows > 3 Then
        .Select 3, 2
    End If
End With
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Not IsNumeric(grid1.EditText) Then
    Cancel = True
    Exit Sub
End If
End Sub
Private Sub xitem_GotFocus()
    myGotFocus xitem
End Sub
Private Sub xitem_LostFocus()
    myLostFocus xitem
    myvalidItem
End Sub
Private Sub xModel_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then ModelLookupAll Me, oSearch
End Sub
Private Sub xModel_GotFocus()
    myGotFocus XMODEL
End Sub
Private Sub XModel_LostFocus()
    myLostFocus XMODEL
End Sub
Sub myProc()
On Error GoTo myerror
ActiveControl.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
Unload oSearch
myload
grid1.SetFocus
Exit Sub
myerror:
End Sub
Private Sub myload()
mydefine
If Trim(XMODEL.text) = "" Then Exit Sub
Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT FILE1_10.*,FACT.DESCA AS FACT_DESCA FROM FILE1_10 INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE "
cString = cString & turn(cString) & "FILE1_10.MODEL = " & MyParn(XMODEL.text)
loctable.Open cString, con, adOpenStatic, adLockReadOnly, acmdtext
    If Not loctable.EOF Then
    XMODEL.text = loctable!MODEL
    xSupp.Caption = loctable!SUPP & ""
    xDesca.Caption = loctable!DESCA & ""
    xMosm.Caption = loctable!MOSM & ""
    xSupp.Caption = loctable!SUPP & ""
    xModelFact.Caption = loctable!modelfact0 & ""
    cString = "Select sum([IN] - [OUT]) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM"
    If strStore <> "" Then cString = cString & turn(cString) & " STORE = " & MyParn(strStore)
    If XMODEL.text <> "" Then cString = cString & turn(cString) & " FILE1_10.MODEL = " & MyParn(XMODEL.text)
    xBalance.Caption = GetDesca(cString, con)
    xFact.Caption = loctable!fact_Desca & ""
End If
myloadgrd
End Sub
Private Sub CalcTotals()
Dim nTotal As Double
For nRow = 3 To grid1.Rows - 1
    For nCol = 2 To grid1.Cols - 1
        nTotal = nTotal + Val(grid1.TextMatrix(nRow, nCol))
    Next
Next
xTotal.Caption = Myvalue(nTotal)
End Sub
Private Sub xModel_Validate(Cancel As Boolean)
myload
End Sub
Private Sub myvalidItem()
Dim cString As String
If Trim(xitem.text) <> "" Then
    cString = "Select Model from file1_10"
    cString = cString & turn(cString) & " FILE1_10.item =" & xitem.text
    XMODEL.text = GetDesca(cString, con)
    If XMODEL.text <> "" Then
        xitem.text = ""
        myload
        On Error Resume Next
        grid1.SetFocus
        Err.Clear
    End If
End If
End Sub
