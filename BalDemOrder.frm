VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form RepDemDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·»Ì«  «·›—Ê⁄"
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
   Begin VB.Frame Frame3 
      Height          =   9150
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   990
      Width           =   23010
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   8880
         Left            =   10665
         TabIndex        =   11
         Top             =   225
         Width           =   12210
         _cx             =   21537
         _cy             =   15663
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   8880
         Left            =   90
         TabIndex        =   15
         Top             =   225
         Width           =   10455
         _cx             =   18441
         _cy             =   15663
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
         BackColorSel    =   8454016
         ForeColorSel    =   -2147483630
         BackColorBkg    =   12648447
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
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   5010
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   1725
         Picture         =   "BalDemOrder.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "⁄—÷"
         Top             =   180
         Width           =   1860
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   135
         Picture         =   "BalDemOrder.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   5580
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   9240
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "«ŸÂ«— «·ﬁ”„"
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
         Height          =   510
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   450
         Width           =   1365
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "«ŸÂ«— «·„Ê”„"
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
         Height          =   510
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   90
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo XCODE 
         Height          =   315
         Left            =   4815
         TabIndex        =   2
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   180
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   4815
         TabIndex        =   12
         Top             =   585
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XSECTION 
         Height          =   315
         Left            =   1440
         TabIndex        =   13
         Top             =   540
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬁ”„ :"
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
         Index           =   0
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   540
         Width           =   960
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄ :"
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
         Left            =   8340
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Left            =   8340
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   210
         Width           =   600
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
         Left            =   3825
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1050
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
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   495
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   3885
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
      Left            =   2685
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
      Left            =   1590
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
End
Attribute VB_Name = "RepDemDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Variant, sM_Fact As String
Dim oSearch As New Search3
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    DATA1.ConnectionString = strCon
    DATA1.RecordSource = "Select code,desca From STORE_BR ORDER BY code "
    Set xStore.RowSource = DATA1
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    If cBranch <> "00" Then
        xStore.Enabled = False
        xStore.BoundText = cBranchStore
    End If
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File1_10SC ORDER BY DESCA"
    Set xSection.RowSource = DATA2
    xSection.ListField = "Desca"
    xSection.BoundColumn = "Code"
    
    DATA3.ConnectionString = strCon
    DATA3.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC "
    Set xMosm.RowSource = DATA3
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    
    data4.ConnectionString = strCon
    data4.RecordSource = "Select code ,desca From FILE4_10 WHERE CODE IN (SELECT CODE FROM Q_DOC_DEM ) ORDER BY code "
    Set xCode.RowSource = data4
    xCode.ListField = "Desca"
    xCode.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    
    Set Grid2.DataSource = DATA7
    DATA7.ConnectionString = strCon
    
    grid1.Rows = 1
    grid1.Cols = 10
    
    Fixgrd
    Fixgrd2
End Sub
Private Sub myload()
    Dim cWhere As String, cFiled1 As String
    If Check1.Value = 0 And Check2.Value = 0 Then
        cString = " SELECT  STORE, STORE_NAME, code, DESCA,' ',' ' , ' ', sum(COUNT_MODEL) , sum(QUANT), min(F_DATE), max(E_DATE ) From Q_DOC_DEM WHERE STORE IS NOT NULL "
    End If
    If Check1.Value = 1 And Check2.Value = 1 Then
        cString = " SELECT  STORE, STORE_NAME, code, DESCA,mosm , [SECTION] , SEC_NAME , sum(COUNT_MODEL) , sum(QUANT), min(F_DATE), max(E_DATE ) From Q_DOC_DEM WHERE STORE IS NOT NULL "
    End If
    If Check1.Value = 1 And Check2.Value = 0 Then
        cString = " SELECT  STORE, STORE_NAME, code, DESCA,mosm , ' ' , ' ', sum(COUNT_MODEL) , sum(QUANT), min(F_DATE), max(E_DATE ) From Q_DOC_DEM WHERE STORE IS NOT NULL "
    End If
    If Check1.Value = 0 And Check2.Value = 1 Then
        cString = " SELECT  STORE, STORE_NAME, code, DESCA,' ', [SECTION] , SEC_NAME , sum(COUNT_MODEL) , sum(QUANT), min(F_DATE), max(E_DATE ) From Q_DOC_DEM WHERE STORE IS NOT NULL "
    End If
    If xMosm.BoundText <> "" Then cString = cString & " AND [MOSM] = " & MyParn(xMosm.BoundText)
    If xSection.BoundText <> "" Then cString = cString & " AND [SECTION] = " & Val(xSection.BoundText)
    If xStore.BoundText <> "" Then cString = cString & " AND [STORE] = " & MyParn(xStore.BoundText)
    If xCode.BoundText <> "" Then cString = cString & " AND [CODE] = " & MyParn(xCode.BoundText)
    If Check1.Value = 0 And Check2.Value = 0 Then
        cString = cString & " group by STORE, STORE_NAME, code, DESCA ORDER BY STORE, DESCA"
    End If
    If Check1.Value = 1 And Check2.Value = 1 Then
        cString = cString & " group by STORE, STORE_NAME, code, DESCA ,mosm , [SECTION] , SEC_NAME ORDER BY STORE, DESCA"
    End If
    If Check1.Value = 1 And Check2.Value = 0 Then
        cString = cString & " group by STORE, STORE_NAME, code, DESCA ,mosm  ORDER BY STORE, DESCA"
    End If
    If Check1.Value = 0 And Check2.Value = 1 Then
        cString = cString & " group by STORE, STORE_NAME, code, DESCA ,[SECTION] , SEC_NAME ORDER BY STORE, DESCA"
    End If
    DATA6.RecordSource = cString
    DATA6.Refresh
    Fixgrd
End Sub
Sub Fixgrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .RowHeight(0) = 1000
    .Cols = 13
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·›—⁄"
    .TextMatrix(0, 2) = "ﬂÊœ"
    .TextMatrix(0, 3) = "«·„Ê—œ"
    
    .TextMatrix(0, 4) = "«·„Ê”„"
    .TextMatrix(0, 5) = " ﬂÊœ"
    .TextMatrix(0, 6) = "«·ﬁ”„"
    
    .ColHidden(4) = (Check1.Value = 0)
    
    .ColHidden(5) = (Check2.Value = 0)
    .ColHidden(6) = (Check2.Value = 0)
    
    .TextMatrix(0, 7) = " ⁄œœ „ÊœÌ·« "
    .TextMatrix(0, 8) = " ⁄œœ «·ﬁÿ⁄"
    .TextMatrix(0, 9) = "«Ê· ÿ·»Ì…"
    .TextMatrix(0, 10) = "«Œ— ÿ·»Ì…"
    
    .TextMatrix(0, 11) = " —ÕÌ·"
    .TextMatrix(0, 12) = "Õ–›"
    
    .ColHidden(0) = True
    .ColHidden(2) = True
    .ColHidden(5) = True
    
    .ColHidden(11) = (cBranch <> "00")
    .ColHidden(12) = (cBranch <> "00")
    
    
    .ColWidth(0) = 0
    .ColWidth(1) = 1600
    .ColWidth(2) = 0
    .ColWidth(3) = 2000
    .ColWidth(4) = 700
    .ColWidth(5) = 0
    .ColWidth(6) = 1500
    .ColWidth(7) = 800
    .ColWidth(8) = 1000
    .ColWidth(9) = 1300
    .ColWidth(10) = 1300
    .ColWidth(11) = 700
    .ColWidth(12) = 700
    .ColComboList(12) = "..."
    .ColComboList(11) = "..."
    .ColDataType(9) = flexDTDate
    .ColDataType(10) = flexDTDate
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
Dim cDoc_No As String
On Error GoTo myerror
con.BeginTrans
With grid1
If Col = 12 Then
    If MsgBox(" Õ–› «·ÿ·»Ì… »«·ﬂ«„· ", vbYesNo + vbDefaultButton2) = vbYes Then
        cStr1 = " DELETE FROM Model_Dem FROM  Model_Dem INNER JOIN FILE1_10 ON Model_Dem.ITEM = FILE1_10.ITEM  WHERE STORE = " & MyParn(.TextMatrix(.Row, 0)) & " AND FILE1_10.CODE = " & MyParn(.TextMatrix(.Row, 2))
        If Check1.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(.TextMatrix(.Row, 4))
        If Check2.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(.TextMatrix(.Row, 5))
        If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
        If xSection.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(xSection.BoundText)
        con.Execute cStr1, nRec
        Inform "  „ Õ–›  " & nRec
        .RemoveItem .Row
    End If
End If

If Col = 11 Then
    If MsgBox("  —ÕÌ· «·ÿ·»Ì… ·„” ‰œ «„— ‘—«¡ ", vbYesNo + vbDefaultButton2) = vbYes Then
        
        cDoc_No = RetZero(Newflag("FILE6_50H", "doc_no"))
        cStr1 = "  INSERT INTO FILE6_50H (STORE, CODE, MOSM, date , doc_no  ) SELECT Model_Dem.STORE, FILE1_10.code ,MAX(FILE1_10.MOSM) , " & addDate(Date) & " , " & addstring(cDoc_No) & "  FROM Model_Dem INNER JOIN FILE1_10 ON Model_Dem.ITEM = FILE1_10.ITEM WHERE STORE = " & MyParn(.TextMatrix(.Row, 0)) & " AND FILE1_10.CODE = " & MyParn(.TextMatrix(.Row, 2))
        If Check1.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(.TextMatrix(.Row, 4))
        If Check2.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(.TextMatrix(.Row, 5))
        If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
        If xSection.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(xSection.BoundText)
        cStr1 = cStr1 & " group by file1_10.code , Model_Dem.STORE "
        con.Execute cStr1
        
        cStr1 = "  INSERT INTO FILE6_50 (doc_no , item , quant, price , discount ,model ) SELECT " & addstring(cDoc_No) & " , Model_Dem.item  , quant , file1_10.cost ,0, file1_10.model   FROM Model_Dem INNER JOIN FILE1_10 ON Model_Dem.ITEM = FILE1_10.ITEM  WHERE quant <> 0 AND STORE = " & MyParn(.TextMatrix(.Row, 0)) & " AND FILE1_10.CODE = " & MyParn(.TextMatrix(.Row, 2))
        If Check1.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(.TextMatrix(.Row, 4))
        If Check2.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(.TextMatrix(.Row, 5))
        If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
        If xSection.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(xSection.BoundText)
        con.Execute cStr1
        
        Inform "  „ ⁄„· «·„” ‰œ " & nRec

        cStr1 = " DELETE FROM Model_Dem FROM  Model_Dem INNER JOIN FILE1_10 ON Model_Dem.ITEM = FILE1_10.ITEM  WHERE STORE = " & MyParn(.TextMatrix(.Row, 0)) & " AND FILE1_10.CODE = " & MyParn(.TextMatrix(.Row, 2))
        If Check1.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(.TextMatrix(.Row, 4))
        If Check2.Value <> 0 Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(.TextMatrix(.Row, 5))
        If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
        If xSection.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[SECTION] = " & Val(xSection.BoundText)
        con.Execute cStr1, nRec
        Inform "  „ Õ–›  " & nRec
        .RemoveItem .Row
    End If
End If
End With
con.CommitTrans
Exit Sub
con.RollbackTrans
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_DblClick()
    MYLOAD2
End Sub
Private Sub grid1_EnterCell()
With grid1
    If .Col >= 11 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Sub Fixgrd2()
    With Grid2
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .RowHeight(0) = 1000
    .Cols = 11
    .ColHidden(8) = True
    .TextMatrix(0, 0) = "—ﬁ„ „ÊœÌ·"
    .TextMatrix(0, 1) = "«·„ÊœÌ·"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "«··Ê‰"
    .TextMatrix(0, 4) = "«·„ﬁ«”"
    .TextMatrix(0, 5) = "«·”⁄—"
    .TextMatrix(0, 6) = " ﬂ„Ì…"
    .TextMatrix(0, 7) = " «—ÌŒ"
    .TextMatrix(0, 8) = " «·›—⁄"
    .TextMatrix(0, 10) = " Õ–›"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 1700
    .ColWidth(2) = 700
    .ColWidth(3) = 1000
    .ColWidth(4) = 700
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    .ColWidth(7) = 1300
    .ColWidth(8) = 600
    .ColWidth(9) = 600
    .ColWidth(10) = 600
    .ColComboList(10) = "..."
    .ColDataType(7) = flexDTDate
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    .ColHidden(9) = True
End With
End Sub
Sub MYLOAD2()
    With grid1
'                       0                       1               2               3               4               5               6                   7           8
    cStr1 = " SELECT FILE1_10.MODELFACT0, FILE1_10.desca, FILE1_10.mosm, FILE1_10.COLOR, FILE1_10.SCAL, FILE1_10.price , Model_Dem.QUANT, Model_Dem.DATE, Model_Dem.Branch , MODEL_DEM.ITEM  FROM    Model_Dem INNER JOIN FILE1_10 ON Model_Dem.ITEM = FILE1_10.ITEM where quant > 0 " & _
        " and STORE = " & MyParn(.TextMatrix(.Row, 0)) & _
        " AND FILE1_10.CODE = " & MyParn(.TextMatrix(.Row, 2))
        If xMosm.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
        If xSection.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[SECTION]= " & Val(xSection.BoundText)
        If Check1.Value = 1 Then cStr1 = cStr1 & " AND FILE1_10.[mosm]= " & MyParn(.TextMatrix(.Row, 4))
        If Check1.Value = 2 Then cStr1 = cStr1 & " AND FILE1_10.[SECTION]= " & Val(.TextMatrix(.Row, 5))
        cStr1 = cStr1 & " ORDER BY FILE1_10.MODELFACT0, FILE1_10.desca, FILE1_10.COLOR , FILE1_10.SCAL "
    End With
    DATA7.RecordSource = cStr1
    DATA7.Refresh
    Fixgrd2
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With Grid2
    If .Col = 6 Then
        con.Execute " UPDATE  Model_Dem SET QUANT = " & Val(.TextMatrix(.Row, 5)) & " WHERE ITEM = " & Val(.TextMatrix(.Row, 8)) & " AND STORE = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
    End If
End With
End Sub
Private Sub GRID2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
If Col = 10 Then
    cStr1 = " DELETE FROM Model_Dem WHERE STORE = " & MyParn(grid1.TextMatrix(grid1.Row, 0)) & " AND Model_Dem.item = " & Val(Grid2.TextMatrix(Grid2.Row, 9))
    con.Execute cStr1, nRec
    Inform "  „ Õ–›  " & nRec
    Grid2.RemoveItem Row
End If
End Sub
Private Sub grid2_EnterCell()
With Grid2
    If .Col = 6 Or .Col = 10 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
