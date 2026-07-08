VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsProftMonthBr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈Ã„«·Ï √—»«Õ ‘Â—Ì… "
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
   Begin VB.Frame Frame2 
      Height          =   510
      Left            =   4005
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   990
      Width           =   3885
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Caption         =   "›—Ê⁄ ›ﬁÿ"
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
         TabIndex        =   18
         Top             =   180
         Width           =   960
      End
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Caption         =   " ÊﬂÌ·«  ›ﬁÿ"
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
         TabIndex        =   17
         Top             =   180
         Width           =   1095
      End
      Begin VB.OptionButton xall 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬂ·"
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
         TabIndex        =   16
         Top             =   180
         Value           =   -1  'True
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1500
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   3840
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "VsProftMonthBr.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "VsProftMonthBr.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2475
         Picture         =   "VsProftMonthBr.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   13
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
         Picture         =   "VsProftMonthBr.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   285
         Left            =   45
         TabIndex        =   14
         Top             =   1170
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   990
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   11235
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
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   5
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
         TabIndex        =   3
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   1305
         TabIndex        =   7
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄ :"
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
         Index           =   1
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ :"
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
         Left            =   4845
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
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
         TabIndex        =   4
         Top             =   270
         Width           =   765
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
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
      Bindings        =   "VsProftMonthBr.frx":9583
      Height          =   7470
      Left            =   90
      TabIndex        =   1
      Top             =   1620
      Width           =   15090
      _cx             =   26617
      _cy             =   13176
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
Attribute VB_Name = "VsProftMonthBr"
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
    cHead1 = "≈Ã„«·Ï «—»«Õ „»Ì⁄«  ‘Â—Ì… " & xStore.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY") & " Õ Ï  «—ÌŒ " & Format(xDate2.text, "DD-MM-YYYY")
    cHead3 = ""

    
    Load PrintGrd
    PrintGrd.DOPRINT grid1, 1, , cHead1, cHead2, cHead3, , True, 9
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
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    DATA7.ConnectionString = strCon
    DATA7.RecordSource = "Select Code,DescA From STORE_BR WHERE " & cUserStore
    Set xStore.RowSource = DATA7
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    If cBranch <> "00" Then
        xStore.BoundText = cBranchStore
        xStore.Enabled = False
    End If
    Set grid1.DataSource = DATA6
    grid1.Rows = 1
    grid1.Cols = 9
    Fixgrd
End Sub
Private Sub myload()
Dim I As Double
Dim cString  As String, cStr2 As String

With grid1
cWhere = " AND YEAR(acc_charge.DATE ) =  Year(QFILE6_20.Date)  AND MONTH (acc_charge.DATE ) =  Month(QFILE6_20.Date)"
If xStore.BoundText <> "" Then cWhere = cWhere & " AND FILE0_40.CODE = " & MyParn(xStore.BoundText)
If IsDate(xDate1.text) Then cWhere = cWhere & " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xDate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xDate2.text) '
If xall(0).Value <> 0 Then cWhere = cWhere & " AND LEN(FILE0_40.CODE) =  3"
If xall(1).Value <> 0 Then cWhere = cWhere & " AND LEN(FILE0_40.CODE) =  2"
If Not bOpt5 Then cWhere = cWhere & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "

cField5 = " SELECT SUM(acc_charge.VALUE) FROM acc_charge INNER JOIN FILE0_40 ON FILE0_40.BRANCH = acc_charge.BRANCH INNER JOIN File8_51 ON File8_51.CODE = ACC_CHARGE.CHARGE WHERE File8_51.ISasset = 0 "
cField5 = " , (" & cField5 & cWhere & ") AS TOTAL_CHARGES"


cWhere = " AND YEAR(FILE8_00H.DATE ) =  Year(QFILE6_20.Date)  AND MONTH (FILE8_00H.DATE ) =  Month(QFILE6_20.Date)"
If xStore.BoundText <> "" Then cWhere = cWhere & " AND FILE8_00.STORE = " & MyParn(xStore.BoundText)
If IsDate(xDate1.text) Then cWhere = cWhere & " AND FILE8_00H.DATE >= " & DateSq(xDate1.text)
If IsDate(xDate2.text) Then cWhere = cWhere & " AND FILE8_00H.DATE <= " & DateSq(xDate2.text)
If xall(0).Value <> 0 Then cWhere = cWhere & " AND LEN(FILE8_00.STORE) =  3 "
If xall(1).Value <> 0 Then cWhere = cWhere & " AND LEN(FILE8_00.STORE) =  2 "
If Not bOpt5 Then cWhere = cWhere & " AND FILE8_00.STORE.[store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
cField6 = "SELECT SUM(FILE8_00.VALUE) FROM FILE8_00 INNER JOIN FILE8_00H ON FILE8_00.DOC_NO = FILE8_00H.DOC_NO  "
cField6 = ",(" & cField6 & cWhere & ") AS TOTAL_DISC "

'                           0               1                 2                3               4                5
    cString = " SELECT YEAR(QFILE6_20.date) , MONTH(QFILE6_20.date) as n2 " & _
                " , SUM(QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100)) AS net_sales " & _
                " , SUM(QFILE6_20.QUANT * QFILE6_20.cost) AS t_cost " & _
                " , SUM((QFILE6_20.TOTAL * ((100-QFILE6_20.RATE )/100))  - (QFILE6_20.QUANT * QFILE6_20.cost)) AS t_proft " & _
                cField5 & _
                cField6 & _
            " FROM  QFILE6_20 " & _
            " WHERE (QFILE6_20.date IS NOT NULL)"
            
    If xStore.BoundText <> "" Then cStr2 = cStr2 & " AND QFILE6_20.STORE = " & MyParn(xStore.BoundText)
    If Not bOpt5 Then cStr2 = cStr2 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    
    If IsDate(xDate1.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] >= " & DateSq(xDate1.text)
    If IsDate(xDate2.text) Then cStr2 = cStr2 & " AND QFILE6_20.[DATE] <= " & DateSq(xDate2.text)
    
    If xall(0).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 1"
    If xall(1).Value <> 0 Then cStr2 = cStr2 & " AND QFILE6_20.ISBRANCH = 0"
    
    cString = cString & cStr2 & " GROUP BY  YEAR(QFILE6_20.date) , MONTH(QFILE6_20.date) ORDER BY YEAR(QFILE6_20.date) , MONTH(QFILE6_20.date)"
    
    Set DATA6.Recordset = myRecordSet(cString, con)
End With
Fixgrd
End Sub
Sub Fixgrd()
Dim nTotal_Sales As Double, nTotal_in As Double
    With grid1
    .Cols = 12
    .RowHeight(0) = 1000
    .WordWrap = True
    .TextMatrix(0, 0) = "«·”‰…"
    .TextMatrix(0, 1) = "«·‘Â—"
    .TextMatrix(0, 2) = "„»Ì⁄« "
    .TextMatrix(0, 3) = " ﬂ·›…"
    .TextMatrix(0, 4) = "«—»«Õ „»Ì⁄« "
    
    .TextMatrix(0, 5) = "„’«—Ì›"
    .TextMatrix(0, 6) = " ”ÊÌ«  „Ê—œÌ‰"
    
    .TextMatrix(0, 7) = "’«›Ì «—»«Õ"
    
    .TextMatrix(0, 8) = "‰”»… —»Õ ··„»Ì⁄« "
    
    .TextMatrix(0, 9) = "«·‰”»… „»Ì⁄«  ··≈Ã„«·Ï"
    
    .TextMatrix(0, 10) = "‰”»… „’«—Ì› ··„»Ì⁄« "
    .TextMatrix(0, 11) = "‰”»… ’«›Ï «·—»Õ ··„»Ì⁄« "
    
    .ColWidth(0) = 800
    .ColWidth(1) = 1000
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    .ColWidth(5) = 1400
    .ColWidth(6) = 1400
    .ColWidth(7) = 1400
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1000
    
    For I = 2 To .Cols - 1
        .ColDataType(I) = flexDTDouble
    Next
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    For I = 1 To .Rows - 1
        .TextMatrix(I, 2) = Round(Val(.TextMatrix(I, 2)), 2)
        .TextMatrix(I, 3) = Round(Val(.TextMatrix(I, 3)), 2)
        .TextMatrix(I, 4) = Round(Val(.TextMatrix(I, 4)), 2)
        .TextMatrix(I, 5) = Round(Val(.TextMatrix(I, 5)), 2)
        .TextMatrix(I, 6) = Round(Val(.TextMatrix(I, 6)), 2)
        
        .TextMatrix(I, 4) = Round(Val(.TextMatrix(I, 2)) - Val(.TextMatrix(I, 3)), 2)
        .TextMatrix(I, 7) = Round(Val(.TextMatrix(I, 4)) - Val(.TextMatrix(I, 5)) + Val(.TextMatrix(I, 6)), 2)
        If Val(.TextMatrix(I, 2)) <> 0 Then .TextMatrix(I, 8) = Round(Val(.TextMatrix(I, 4)) / Val(.TextMatrix(I, 2)) * 100, 2)
        If Val(.TextMatrix(I, 2)) <> 0 Then .TextMatrix(I, 10) = Round(Val(.TextMatrix(I, 5)) / Val(.TextMatrix(I, 2)) * 100, 2)
        If Val(.TextMatrix(I, 2)) <> 0 Then .TextMatrix(I, 11) = Round(Val(.TextMatrix(I, 7)) / Val(.TextMatrix(I, 2)) * 100, 2)
    Next I
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 3, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0.00", vbRed, vbYellow, True, "  "
    If .Rows > 1 Then
        .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
                
        If Val(.TextMatrix(1, 2)) <> 0 Then .TextMatrix(1, 8) = Round(Val(.TextMatrix(1, 4)) / Val(.TextMatrix(1, 2)) * 100, 2)
        If Val(.TextMatrix(1, 2)) <> 0 Then .TextMatrix(1, 10) = Round(Val(.TextMatrix(1, 5)) / Val(.TextMatrix(1, 2)) * 100, 2)
        If Val(.TextMatrix(1, 2)) <> 0 Then .TextMatrix(1, 11) = Round(Val(.TextMatrix(1, 7)) / Val(.TextMatrix(1, 2)) * 100, 2)
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
    For I = 2 To .Rows - 1
        If Val(.TextMatrix(1, 4)) <> 0 Then .TextMatrix(I, 9) = Round(Val(.TextMatrix(I, 4)) / Val(.TextMatrix(1, 4)) * 100, 2)
    Next I
    .Subtotal flexSTSum, -1, 9, "#0.00", vbRed, vbYellow, True, "  "
'    xTotal.Caption = IIf(grid1.Rows > 1, grid1.Rows - 2, "")
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
End Sub
Private Sub xSupp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then SuppLookupAll Me, oSearch
End Sub
Sub myProc()
If ActiveControl.Name = xSupp.Name Then
    xSupp.BoundText = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
End If
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


