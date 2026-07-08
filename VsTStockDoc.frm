VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTStockDoc 
   Caption         =   "≈Ã„«·Ï „” ‰œ«   ”ÊÌ«  «·Ã—œ"
   ClientHeight    =   9180
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   13260
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
   ScaleHeight     =   9180
   ScaleWidth      =   13260
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   810
      Width           =   4920
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2430
         Picture         =   "VsTStockDoc.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "VsTStockDoc.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "VsTStockDoc.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1230
         Picture         =   "VsTStockDoc.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   8670
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
         Left            =   5460
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   135
         Width           =   1815
      End
      Begin VB.TextBox xDate2 
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
         Left            =   1005
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   135
         Width           =   1575
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   3840
         TabIndex        =   7
         Top             =   510
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Œ“‰ :"
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
         Index           =   3
         Left            =   7395
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   525
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   7395
         TabIndex        =   5
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
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
         Index           =   0
         Left            =   2640
         TabIndex        =   4
         Top             =   195
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   8850
      Width           =   13260
      _ExtentX        =   23389
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
      Left            =   2520
      Top             =   405
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
      Bindings        =   "VsTStockDoc.frx":9573
      Height          =   7095
      Left            =   150
      TabIndex        =   6
      Top             =   1530
      Width           =   14865
      _cx             =   26220
      _cy             =   12515
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
      Left            =   60
      Top             =   120
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
      Left            =   300
      Top             =   120
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
      Left            =   60
      Top             =   120
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
   Begin ComctlLib.ProgressBar prog1 
      Height          =   240
      Left            =   180
      TabIndex        =   14
      Top             =   495
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "VsTStockDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub cmdExel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "≈Ã„«·Ï  ”ÊÌ«  «·Ã—œ " & xStore.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xDate2.text, "DD-MM-YYYY")
    
    Load PrintGrd
    PrintGrd.DOPRINT Me.grid1, 1, -2, cHead1, cHead2, , False, , 9
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
    xDate1.text = "1-1-" & Year(Date)
    xDate2.text = Format(Date, "dd-mm-yyyy")

    openCon con
    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File0_40 order by Desca"
    Set xStore.RowSource = data1
    xStore.ListField = "Desca"
    xStore.BoundColumn = "Code"
    If cBranch <> "" Or lIsBranchStore Then
        xStore.BoundText = cBranchStore
        xStore.Enabled = False
    End If
    
    Set grid1.DataSource = DATA4
    DATA4.ConnectionString = strCon
    
    FIXGRID
    grid1.Rows = 1
End Sub
Private Sub myload()
Dim cWhere As String
With grid1
'                           0               1
    cString = "SELECT  FILE0_10H.DOC_NO, FILE0_10H.DATE, FILE0_40.DESCA, FILE0_10H.DESCA AS notes, SUM(RealBal) AS T_QUANT1 , SUM(ComputerBal) AS T_QUANT2  , SUM(Differ) AS T_QUANT3 , SUM(Differ * file0_10.COST ) AS T_QUANT3 " & _
                " FROM FILE0_10 INNER JOIN FILE0_10H ON FILE0_10.doc_no = FILE0_10H.DOC_NO INNER JOIN FILE0_40 ON FILE0_10H.STORE = FILE0_40.CODE  inner join file1_10 on file1_10.item = FILE0_10.item where FILE0_10H.DOC_NO is not null "
    cWhere = ""
    If xStore.BoundText <> "" Then cString = cString & " and file0_10H.[STORE]  = " & MyParn(xStore.BoundText)
    If IsDate(xDate1.text) Then cString = cString & " and date >= " & DateSq(xDate1.text)
    If IsDate(xDate2.text) Then cString = cString & " and date <= " & DateSq(xDate2.text)
    cString = cString & " GROUP BY FILE0_10H.DOC_NO, FILE0_10H.DATE, FILE0_40.DESCA, FILE0_10H.DESCA ORDER BY FILE0_10H.DOC_NO"
    DATA4.RecordSource = cString
    DATA4.Refresh
End With
FIXGRID
End Sub
Sub FIXGRID()
    With grid1
    .Cols = 8
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "„” ‰œ"
    .TextMatrix(0, 1) = " «—ÌŒ"
    .TextMatrix(0, 2) = "«·„Œ“‰"
    .TextMatrix(0, 3) = "«·»Ì«‰"
    .TextMatrix(0, 4) = "Ã—œ ›⁄·Ï"
    .TextMatrix(0, 5) = "Ã—œ œ› —Ï"
    .TextMatrix(0, 6) = "⁄œœ ÕÃ“ "
    .TextMatrix(0, 7) = "ﬁÌ„… «·⁄Ã—"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1300
    .ColWidth(2) = 2000
    .ColWidth(3) = 4000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1400
    
    .ColDataType(4) = flexDTDouble
    .ColDataType(1) = flexDTDate
    
    .ColFormat(4) = "#0"
    .ColFormat(5) = "#0"
    .ColFormat(6) = "#0"
    .ColFormat(7) = "#0.00"

    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTBelow
    
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
        .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
        .Subtotal flexSTSum, -1, 5, "#0", vbRed, vbYellow, True, "  "
        .Subtotal flexSTSum, -1, 6, "#0", vbRed, vbYellow, True, "  "
        .Subtotal flexSTSum, -1, 7, "#0.00", vbRed, vbYellow, True, "  "
        StatusBar1.Panels(1).text = "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 2
        .TextMatrix(.Rows - 1, 1) = "«·≈Ã„«·Ï"
    End If
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
