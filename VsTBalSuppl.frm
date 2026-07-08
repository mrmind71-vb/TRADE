VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTBalCust 
   Caption         =   "√—’œ… Ê  ⁄«„·«  «·⁄„·«¡"
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15240
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
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdExit 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1020
      Width           =   1200
   End
   Begin VB.CommandButton Cmd_Print 
      Caption         =   "ÿ»«⁄…"
      Height          =   420
      Left            =   1387
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1020
      Width           =   1200
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "⁄—÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2625
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1020
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
      Begin VB.TextBox Xcode 
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
         TabIndex        =   12
         Top             =   622
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
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGrCust 
         Height          =   315
         Left            =   6000
         TabIndex        =   10
         Top             =   1020
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label xCustName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   5010
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   622
         Width           =   2505
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ ⁄„Ì·"
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
         Index           =   1
         Left            =   9555
         TabIndex        =   13
         Top             =   690
         Width           =   720
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… ⁄„·«¡"
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
         Index           =   0
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1050
         Width           =   1230
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
         Left            =   9555
         TabIndex        =   5
         Top             =   270
         Width           =   675
      End
      Begin VB.Label LLL 
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
         Left            =   4680
         TabIndex        =   4
         Top             =   285
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   15240
      _ExtentX        =   26882
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTBalSuppl.frx":0000
      Height          =   8010
      Left            =   150
      TabIndex        =   9
      Top             =   1575
      Width           =   14865
      _cx             =   26220
      _cy             =   14129
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
      BackColorSel    =   16777088
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
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2475
      Top             =   75
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
      Left            =   150
      Top             =   75
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
Attribute VB_Name = "VsTBalCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ √—’œ… Ê  ⁄«„·«  «·⁄„·«¡ "
    If xGrCust.Text <> "" Then cHead1 = cHead1 & xGrCust.Text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(XDATE2.Text, "DD-MM-YYYY")
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 0.9, -2, cHead1, cHead2, , False, True, 8
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
    xDate1.Text = "1-1-" & Year(Date)
    
    XDATE2.Text = Format(Date, "dd-mm-yyyy")
    
    data1.ConnectionString = strCon
    data1.RecordSource = "SELECT * FROM FILE3_50 "
    
    Set xGrCust.RowSource = data1
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA2
    DATA2.ConnectionString = strCon
    
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub myload()
Dim cString As String

If IsDate(xDate1.Text) Then cWhere = " date < " & DateSq(xDate1.Text)
cField2 = myiif(cWhere, "[SAL] - [PAY]") & " AS F_BAL"

cWhere = ""

If IsDate(xDate1.Text) Then cWhere = " date >= " & DateSq(xDate1.Text)
If IsDate(XDATE2.Text) Then cWhere = cWhere & turn(cWhere, " and ") & " DATE <= " & DateSq(XDATE2.Text)

cField3 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '4' and sal > 0 )", "[SAL]") & " AS T_SALES"
cField4 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '4' and sal < 0 )", "([sal]* -1)") & " AS T_RET"
cField5 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '4' )", "[sal]") & " AS net_sal"

cField8 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '7' or TYPE = '8'  )", "[PAY] - [SAL]") & " AS T_CASH"
cField9 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = 'A' OR TYPE = 'C' )", "[PAY]") & " AS T_CHQ"
cField10 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '9')", "[PAY]") & " AS T_CASH0 "

cWhere = ""
If IsDate(XDATE2.Text) Then cWhere = "date <= " & DateSq(XDATE2.Text)
cField11 = myiif(cWhere, "[sal]- [pay] ") & " AS C_BAL"

With grid1
'                           0                           1
    cString = "  select FILE3_10.CODE AS CODE , FILE3_10.DESCA AS DESCA , " & _
                cField2 & " , " & cField3 & " , " & cField4 & " , " & cField5 & " ,' 'AS N6 , ' ' AS N7  , " & _
                cField8 & " , " & cField9 & " , " & cField10 & " , " & cField11 & " , ' ' AS N12 " & _
                " FROM (FILE3_11 LEFT JOIN FILE3_10 ON FILE3_11.CODE = FILE3_10.CODE) LEFT JOIN file3_50 ON FILE3_10.[group] = file3_50.CODE WHERE FILE3_10.CODE IS NOT NULL "
    If xGrCust.BoundText <> "" Then cString = cString & turn(cString) & " file3_10.[group]  = " & MyParn(xGrCust.BoundText)
    If XCODE.Text <> "" Then cString = cString & turn(cString) & " FILE3_10.[CODE] = " & MyParn(XCODE.Text)
    cString = cString & " GROUP BY FILE3_10.DESCA , FILE3_10.CODE "
    DATA2.RecordSource = cString
    DATA2.Refresh
End With
FixGrid
Fillgrd
End Sub
Sub FixGrid()
    With grid1
    .Cols = 14
    .RowHeight(0) = 1000
    .WordWrap = True
    .FrozenCols = 2
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·⁄„Ì·"
    
    .TextMatrix(0, 2) = "«·—’Ìœ " & xDate1.Text
    .TextMatrix(0, 3) = "Ã. „»Ì⁄« "
    .TextMatrix(0, 4) = "Ã. „— Ã⁄«  "
    .TextMatrix(0, 5) = "’«›Ï ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 6) = " ﬂ·›… «·„»Ì⁄«  "
    .TextMatrix(0, 7) = "—»Õ «·„»Ì⁄« "
    .TextMatrix(0, 8) = " ”œ«œ ‰ﬁœÏ"
    .TextMatrix(0, 9) = "”œ«œ √.ﬁ "
    .TextMatrix(0, 10) = " ”ÊÌ« "
    .TextMatrix(0, 11) = "—’Ìœ " & XDATE2.Text
    .TextMatrix(0, 12) = "‘Ìﬂ«  €Ì— „Õ’·…"
    .TextMatrix(0, 13) = " «—ÌŒ √Œ— „»Ì⁄« "
    
    
    .ColWidth(0) = 800
    .ColWidth(1) = 2000
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    .ColWidth(12) = 1200
    .ColWidth(13) = 1200
    
    For nCol = 2 To .Cols - 2
        .ColDataType(nCol) = flexDTDouble
        .ColFormat(nCol) = "#0.00"
    Next nCol
    .ColDataType(13) = flexDTDate
    .ColFormat(13) = "DD-MM-YYYY"
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
End Sub
Sub Fillgrd()
Dim SalesTable As New ADODB.Recordset
Dim ChqTable As New ADODB.Recordset

cStrSal = "SELECT     code, SUM(T_COST) AS T_COST , MAX(date) AS LASTSALES FROM T_SALES WHERE CODE <> '000' "
If IsDate(xDate1.Text) Then cStrSal = cStrSal & turn(cStrSal) & "DATE >= " & DateSq(xDate1.Text)
If IsDate(XDATE2.Text) Then cStrSal = cStrSal & turn(cStrSal) & "DATE <= " & DateSq(XDATE2.Text)
cStrSal = cStrSal & " GROUP BY code "
SalesTable.Open cStrSal, con, adOpenStatic, adLockReadOnly, adCmdText


cStrChq = "SELECT code1 , Sum(VALUE) AS T_VALUE FROM FILE5_20 WHERE CLOSED = '0' "
cStrChq = cStrChq & " GROUP BY code1"
If ChqTable.State = adStateOpen Then ChqTable.Close
ChqTable.Open cStrChq, con, adOpenStatic, adLockReadOnly, adCmdText

    With grid1
    For I = 1 To .Rows - 1
        SalesTable.Find " CODE = " & MyParn(.TextMatrix(I, 0)), , adSearchForward, adBookmarkFirst
        If Not SalesTable.EOF Then
            .TextMatrix(I, 6) = Format(SalesTable!T_COST, "#0.00")
            .TextMatrix(I, 7) = Format(Val(.TextMatrix(I, 5)) - Val(.TextMatrix(I, 6)), "#0.00")
            .TextMatrix(I, 13) = Format(SalesTable!LASTSALES, "DD-MM-YYYY")
        End If
        ChqTable.Find " CODE1 = " & MyParn(.TextMatrix(I, 0)), , adSearchForward, adBookmarkFirst
        If Not ChqTable.EOF Then
            .TextMatrix(I, 12) = Format(ChqTable!T_VALUE, "FIXED")
        End If
        For nCol = 2 To .Cols - 2
            .TextMatrix(I, nCol) = Format(.TextMatrix(I, nCol), "#0.00")
        Next nCol
    Next I
    .SubtotalPosition = flexSTAbove
    For I = 2 To .Cols - 2
        .Subtotal flexSTSum, -1, I, "#0", vbRed, vbYellow, True, "  "
    Next I
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
    SalesTable.Close
    ChqTable.Close
    Set SalesTable = Nothing
    Set ChqTable = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
Private Sub grid1_DBLClick()
    If grid1.Row > 1 Then
        Load ClientMoveFrm
        ClientMoveFrm.XCODE.Text = grid1.TextMatrix(grid1.Row, 0)
        ClientMoveFrm.xDesca.Caption = grid1.TextMatrix(grid1.Row, 1)
        ClientMoveFrm.xDate1.Text = xDate1.Text
        ClientMoveFrm.XDATE2.Text = XDATE2.Text
        ClientMoveFrm.Fillgrd
        ClientMoveFrm.Show
    End If
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
xCustName.Caption = ""
If XCODE.Text = "" Then Exit Sub
XCODE.Text = RetZero(XCODE.Text, 6)
xCustName.Caption = GetDesca("select desca from FILE3_10 where code = " & MyParn(XCODE.Text)) & ""
End Sub
Sub myProc()
ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Unload Search3
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From FILE3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub
