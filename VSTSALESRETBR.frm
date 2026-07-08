VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VSTSALESRETBR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÅÌãÇáì ãÈíÚÇÊ - ãÑÊÌÚ - ãÑÊÌÚ ãä ÝÑæÚ ÇÎÑì"
   ClientHeight    =   10365
   ClientLeft      =   0
   ClientTop       =   375
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1365
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   3795
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "VSTSALESRETBR.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "VSTSALESRETBR.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2475
         Picture         =   "VSTSALESRETBR.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "ÚÑÖ"
         Top             =   135
         Width           =   1275
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   15
         Top             =   630
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
         Picture         =   "VSTSALESRETBR.frx":6D88
         Caption         =   "ÊÍæíá ÇßÓíá "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   45
         TabIndex        =   16
         Top             =   1125
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   900
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   10695
      Begin VB.Frame Frame2 
         Height          =   465
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   135
         Width           =   3885
         Begin VB.OptionButton xall 
            Alignment       =   1  'Right Justify
            Caption         =   "ÝÑæÚ ÝÞØ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   2610
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   180
            Width           =   960
         End
         Begin VB.OptionButton xall 
            Alignment       =   1  'Right Justify
            Caption         =   "ÊæßíáÇÊ ÝÞØ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   1035
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton xall 
            Alignment       =   1  'Right Justify
            Caption         =   "Çáßá"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   225
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   180
            Value           =   -1  'True
            Width           =   555
         End
      End
      Begin VB.TextBox xdate2 
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
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox xDate1 
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
         Left            =   7620
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   6000
         TabIndex        =   6
         Top             =   1575
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
         Left            =   1305
         TabIndex        =   7
         Top             =   1470
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ÍÊì ÊÇÑíÎ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   6375
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ãä ÊÇÑíÎ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "ÇáãæÓã :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1530
         Width           =   1230
      End
      Begin VB.Label Label2 
         Caption         =   "ãÌãæÚÉ ãæÏíáÇÊ :"
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
         Index           =   5
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1620
         Width           =   1410
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   10035
      Width           =   15120
      _ExtentX        =   26670
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VSTSALESRETBR.frx":9583
      Height          =   8100
      Left            =   90
      TabIndex        =   4
      Top             =   1395
      Width           =   14910
      _cx             =   26300
      _cy             =   14287
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorSel    =   16776960
      ForeColorSel    =   64
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   2
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
   Begin MSAdodcLib.Adodc data3 
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
   Begin MSAdodcLib.Adodc data5 
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
   Begin MSAdodcLib.Adodc DATA6 
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
      Caption         =   "data7"
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
Attribute VB_Name = "VSTSALESRETBR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = "ÅÌãÇáì ãÈíÚÇÊ ÇáãÚÇÑÖ æ ÇáÊæßíáÇÊ"
    cHead2 = " ãä ÊÇÑíÎ " & Format(xDate1.text, "DD-MM-YYYY") & " ÍÊì ÊÇÑíÎ " & Format(xDate2.text, "DD-MM-YYYY")
    
    Load PrintGrd
        PrintGrd.DOPRINT grid1, 1, , cHead1, cHead2, cHead3, , True, 10
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    If IsDate(xDate1.text) And IsDate(xDate2.text) Then
        myload
    Else
        MsgBox " ÊÍÏíÏ ÇáÊÇÑíÎ "
    End If
End Sub
Private Sub Form_Load()
    openCon con
    
    
    
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    
    grid1.Rows = 1
    grid1.Cols = 9
    FIXGRID
End Sub
Private Sub myload()
Dim I As Double
Dim cString  As String, cStr2 As String
With grid1
'                           0               1                 2                3               4                5
    
    
    
CW = ""
If IsDate(xDate1.text) Then CW = " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xDate2.text) Then CW = CW & " AND DATE >= " & DateSq(xDate1.text)
    
cString = " SELECT STORE_ALL.CODE , STORE_ALL.DESCA , COUNT(DISTINCT DOC_NO ) , SUM(QFILE6_20_RET.TOTAL) " & _
        " , SUM(CASE WHEN T_BOON > 0 THEN QFILE6_20_RET.TOTAL ELSE 0 END ) AS TSALES1 " & _
        " , SUM(CASE WHEN T_BOON < 0 THEN QFILE6_20_RET.TOTAL * -1 ELSE 0 END ) AS TRETURN " & _
        " ,SUM(CASE WHEN STORE = STORE_RET AND  T_BOON < 0 THEN QFILE6_20_RET.TOTAL * -1 ELSE 0 END )   AS        T_RETURN1 " & _
        " ,SUM(CASE WHEN STORE <> STORE_RET AND STORE_RET IS NOT NULL  AND  T_BOON < 0 THEN QFILE6_20_RET.TOTAL * -1 ELSE 0 END )   AS    T_RETURN2 " & _
        " ,SUM(CASE WHEN STORE_RET IS NULL  AND  T_BOON < 0 THEN QFILE6_20_RET.TOTAL * -1 ELSE 0 END )   AS   T_RETURN3 " & _
        " , (SELECT SUM(QFILE6_20_RET_2.TOTAL * -1 ) FROM QFILE6_20_RET AS QFILE6_20_RET_2 WHERE QFILE6_20_RET_2.T_BOON < 0 AND STORE_ALL.CODE = QFILE6_20_RET_2.STORE_RET AND QFILE6_20_RET_2.STORE <>  STORE_ALL.CODE " & CW & " ) " & _
        " FROM QFILE6_20_RET INNER JOIN STORE_ALL ON STORE_ALL.CODE = QFILE6_20_RET.STORE WHERE STORE_ALL.CODE IS NOT NULL "
    
    If IsDate(xDate1.text) Then cString = cString & " AND QFILE6_20_RET.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xDate2.text) Then cString = cString & " AND QFILE6_20_RET.[DATE] <= " & DateSq(xDate2.text)
'    If Not bOpt5 Then cStr2 = cStr2 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    If xall(0).Value <> 0 Then cString = cString & " AND QFILE6_20_RET.ISBRANCH = 1 "
    If xall(1).Value <> 0 Then cString = cString & " AND QFILE6_20_RET.ISBRANCH = 2 "
    
    cString = cString & " GROUP BY STORE_ALL.CODE , STORE_ALL.DESCA  ORDER BY STORE_ALL.CODE "
    DATA6.RecordSource = cString
    DATA6.Refresh
End With
FIXGRID
'If grid1.Rows > 1 Then grid1.TextMatrix(1, 1) = "ÇáÅÌãÇáì"
End Sub
Sub FIXGRID()
With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
'    .ColHidden(0) = True
    .Cols = 10
    .TextMatrix(0, 0) = "ßæÏ"
    .TextMatrix(0, 1) = "ÇáÝÑÚ "
    
    
    .TextMatrix(0, 2) = "ÚÏÏ ÈæäÇÊ"
    .TextMatrix(0, 3) = "ÕÇÝì ÇáãÈíÚÇÊ"
    .TextMatrix(0, 4) = "ÞíãÉ ãÈíÚÇÊ"
    .TextMatrix(0, 5) = "ÞíãÉ ãÑÊÌÚÇÊ"
    .TextMatrix(0, 6) = "ãÑÊÌÚÇÊ ãä ÇáÝÑÚ áãÈíÚÇÊ ÇáÝÑÚ "
    .TextMatrix(0, 7) = "ãÑÊÌÚÇÊ ãä ÇáÝÑÚ áãÈíÚÇÊ ÝÑæÚ ÇÎÑì"
    .TextMatrix(0, 8) = "ãÑÊÌÚÇÊ ãä ÇáÝÑÚ ÈÏæä Èæä ãÈíÚÇÊ"
    .TextMatrix(0, 9) = "ãÑÊÌÚÇÊ ãä ÝÑæÚ ÇÎÑì ãä ãÈíÚÇÊ ÇáÝÑÚ "
    
    .ColWidth(0) = 600
    .ColWidth(1) = 2000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1500
    .ColWidth(9) = 1500
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    For nCol = 3 To .Cols - 1
        .ColFormat(nCol) = "#,###,##0"
        .ColDataType(nCol) = flexDTDouble
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, nCol) = Format(Round(Val(.TextMatrix(nRow, nCol)), 2), "#,###,##0")
        Next nRow
        .Subtotal flexSTSum, -1, nCol, "#,###,##0", vbRed, vbYellow, True, "  "
        
    Next nCol
    If .Rows > 1 Then .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
End Sub

Private Sub grid1_DblClick()
If grid1.Col = 7 Then
    ViewRetSalesBr.pDateRep1 = xDate1.text
    ViewRetSalesBr.pDateRep2 = xDate2.text
    ViewRetSalesBr.pBranch_Store = grid1.TextMatrix(grid1.Row, 0)
    ViewRetSalesBr.Show 1
End If

If grid1.Col = 9 Then
    ViewRetSalesBr2.pDateRep1 = xDate1.text
    ViewRetSalesBr2.pDateRep2 = xDate2.text
    ViewRetSalesBr2.pBranch_Store = grid1.TextMatrix(grid1.Row, 0)
    ViewRetSalesBr2.Show 1
End If
End Sub
