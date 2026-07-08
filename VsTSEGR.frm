VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VsTSEGR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬁÌ„… «·—’Ìœ „Ê“⁄ ⁄·Ï «·›—Ê⁄ Ê «·„Ã„Ê⁄« "
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1365
      Left            =   -45
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   45
      Width           =   3795
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "VsTSEGR.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "VsTSEGR.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2475
         Picture         =   "VsTSEGR.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   11
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
         Picture         =   "VsTSEGR.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   45
         TabIndex        =   13
         Top             =   1125
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1350
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   45
      Width           =   11235
      Begin VB.Frame Frame2 
         Height          =   510
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   4470
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
            Left            =   3420
            RightToLeft     =   -1  'True
            TabIndex        =   17
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
            Height          =   195
            Index           =   1
            Left            =   1845
            RightToLeft     =   -1  'True
            TabIndex        =   16
            Top             =   180
            Width           =   1140
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
            Height          =   195
            Index           =   2
            Left            =   225
            RightToLeft     =   -1  'True
            TabIndex        =   15
            Top             =   180
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   6000
         TabIndex        =   4
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
      Begin MSComCtl2.DTPicker xdate1 
         Height          =   375
         Left            =   7560
         TabIndex        =   12
         Top             =   180
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   441647105
         CurrentDate     =   43098.0816782407
      End
      Begin VB.Label Label1 
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   885
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
         Index           =   5
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   675
         Width           =   1230
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
      Left            =   1485
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTSEGR.frx":9583
      Height          =   8490
      Left            =   0
      TabIndex        =   1
      Top             =   1440
      Width           =   14910
      _cx             =   26300
      _cy             =   14975
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      BackColorSel    =   12648447
      ForeColorSel    =   4194368
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
   Begin VB.Label xmodelno 
      Alignment       =   1  'Right Justify
      Caption         =   "Label7"
      Height          =   240
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   270
      Visible         =   0   'False
      Width           =   555
   End
End
Attribute VB_Name = "VsTSEGR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3
Public cMyModel As String
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï —’Ìœ «·„Ã„Ê⁄«  „Ê“⁄ ⁄·Ï «·›—Ê⁄ " & xgmosm
    
    Load PrintGrd
    If grid1.Cols > 8 Then
        PrintGrd.DOPRINT Me.grid1, 1, -2, cHead1, xStore.text, , False, True, 10, , Array(1)
    Else
        PrintGrd.DOPRINT Me.grid1, 1, -2, cHead1, xStore.text, , False, False, 10, , Array(1)
    End If
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    xDate1.Value = Date
    
    DATA7.ConnectionString = strCon
    DATA7.RecordSource = "Select Code,DescA From STORE_BR ORDER BY CODE "
    Set xStore.RowSource = DATA7
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    
    grid1.Rows = 2
    grid1.FixedRows = 2
    grid1.Cols = 1
'   FixGrid
End Sub
Private Sub myload()
    Dim cWhere As String
    Dim cStr2 As String
    Dim DataMosmTable As New ADODB.Recordset
    On Error GoTo myerror
    cStr1 = " SELECT CODE , DESCA FROM STORE_BR   ORDER BY CODE "
    DataMosmTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    DataMosmTable.MoveFirst
    cField2 = ""
    Do Until DataMosmTable.EOF
        If xall(0).Value <> 0 Then
            cWhere = " FILE1_11.[STORE] = " & MyParn(DataMosmTable!CODE)
        End If
        If xall(1).Value <> 0 Then
            cWhere = " FR1_11.[STORE] = " & MyParn(DataMosmTable!CODE)
        End If
        If xall(2).Value <> 0 Then
            cWhere = " FILE1_11_ALL.[STORE] = " & MyParn(DataMosmTable!CODE)
        End If
        cField2 = cField2 & turnFound(cField2, ",") & myiif(cWhere, "([IN] - [OUT]) ")
        If xall(1).Value Then
            cField2 = cField2 & turnFound(cField2, ",") & myiif(cWhere, "([IN] - [OUT]) * FILE1_10.COSTITEM_FR ")
        Else
            cField2 = cField2 & turnFound(cField2, ",") & myiif(cWhere, "([IN] - [OUT]) * FILE1_10.COSTITEM ")
        End If
        DataMosmTable.MoveNext
    Loop
    cString = " SELECT FILE1_10.[GROUP] , FILE1_50.DESCA , " & cField2
    
    If xall(0).Value <> 0 Then
        cString = cString & " FROM    FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM INNER JOIN MOSM ON MOSM.MOSM = FILE1_10.MOSM LEFT OUTER JOIN                       FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]  WHERE FILE1_10.ITEM IS NOT NULL "
        If IsDate(xDate1.Value) Then cString = cString & " AND DATE <= " & DateSq(xDate1.Value)
    End If
    If xall(1).Value <> 0 Then
        cString = cString & " FROM    FILE1_10 INNER JOIN FR1_11 ON FILE1_10.ITEM = FR1_11.ITEM INNER JOIN MOSM ON MOSM.MOSM = FILE1_10.MOSM LEFT OUTER JOIN                       FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]  WHERE FILE1_10.ITEM IS NOT NULL "
        If IsDate(xDate1.Value) Then cString = cString & " AND FR1_11.DATE <= " & DateSq(xDate1.Value)
    End If
    If xall(2).Value <> 0 Then
        cString = cString & " FROM    FILE1_10 INNER JOIN FILE1_11_ALL ON FILE1_10.ITEM = FILE1_11_ALL.ITEM INNER JOIN MOSM ON MOSM.MOSM = FILE1_10.MOSM LEFT OUTER JOIN                       FILE1_50 ON FILE1_50.CODE = FILE1_10.[GROUP]  WHERE FILE1_10.ITEM IS NOT NULL "
        If IsDate(xDate1.Value) Then cString = cString & " AND FILE1_11_ALL.DATE <= " & DateSq(xDate1.Value)
    End If
    If xStore.BoundText <> "" Then cString = cString & " AND store = " & MyParn(xStore.BoundText)
    cString = cString & cStr2 & " group BY FILE1_10.[GROUP] , FILE1_50.DESCA ORDER BY FILE1_10.[GROUP] "
    DATA6.RecordSource = cString
    DATA6.Refresh
    FIXGRID
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Sub FIXGRID()
    Dim nTRow As Double
    Dim DataMosmTable As New ADODB.Recordset
    cStr1 = " SELECT CODE , DESCA FROM STORE_BR   ORDER BY CODE "
    DataMosmTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
    DataMosmTable.MoveFirst
    
    With grid1
    .MergeCells = flexMergeFree
    .MergeRow(0) = True
    If .Cols = 1 Then Exit Sub
    .FixedRows = 2
    .FrozenCols = 2
    .ExplorerBar = flexExSortShow
    
    .RowHeight(0) = 1000
    .ColWidth(0) = 500
    .ColWidth(1) = 2500
    
    .WordWrap = True
    .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
    .TextMatrix(0, 0) = "«·„Ã„Ê⁄…"
    .ColWidth(0) = 1000
    .ColDataType(0) = flexDTString
    nCol = 1
    Do Until DataMosmTable.EOF
       nCol = 2 + nCol
        .TextMatrix(0, nCol - 1) = DataMosmTable!DESCA
        .TextMatrix(0, nCol) = DataMosmTable!DESCA
        
        .TextMatrix(1, nCol - 1) = "⁄œœ"
        .TextMatrix(1, nCol) = "ﬁÌ„…"
        
        DataMosmTable.MoveNext
        .ColWidth(nCol - 1) = 700
        .ColWidth(nCol) = 900
        .ColDataType(nCol) = flexDTDouble
        .ColFormat(nCol - 1) = "#0"
        .ColFormat(nCol) = "#0"
    Loop
    .ColWidth(0) = 500
    .ColWidth(1) = 2500
    
    .Cols = .Cols + 2
    .TextMatrix(0, .Cols - 1) = "«·≈Ã„«·Ï"
    .TextMatrix(0, .Cols - 2) = "«·≈Ã„«·Ï"
    .TextMatrix(1, .Cols - 1) = "ﬁÌ„…"
    .TextMatrix(1, .Cols - 2) = "⁄œœ"
    
    .ColWidth(.Cols - 1) = 1200
    .ColWidth(.Cols - 2) = 800
    .ColDataType(.Cols - 1) = flexDTDouble
    .ColFormat(.Cols - 1) = "#0"
    .ColFormat(.Cols - 2) = "#0"
    
    For nRow = 2 To .Rows - 1
        nTRow1 = 0
        nTRow2 = 0
        For nCol = 3 To .Cols - 3 Step 2
            .TextMatrix(nRow, nCol) = Round(Val(.TextMatrix(nRow, nCol)))
            .TextMatrix(nRow, nCol - 1) = Round(Val(.TextMatrix(nRow, nCol - 1)))
            nTRow1 = nTRow1 + Val(.TextMatrix(nRow, nCol - 1))
            nTRow2 = nTRow2 + Val(.TextMatrix(nRow, nCol))
        Next nCol
        .TextMatrix(nRow, .Cols - 2) = Format(nTRow1, "#0")
        .TextMatrix(nRow, .Cols - 1) = Format(nTRow2, "#0")
'        If nTRow1 + nTRow2 = 0 Then
'            .RemoveItem nRow
'            nRow = nRow - 1
'        End If
    Next nRow
    .SubtotalPosition = flexSTAbove
    For nCol = 2 To .Cols - 1
        .Subtotal flexSTSum, -1, nCol, "#0", vbYellow, , True, " "
    Next nCol
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = 4
    .Cell(flexcpAlignment, 2, 0, .Rows - 1, .Cols - 1) = 7
    If .Rows > 3 Then .TextMatrix(2, 1) = "«·≈Ã„«·Ï"
End With
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
