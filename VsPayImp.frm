VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form VsTBALIMP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈Ã„«·Ï «—’œ… «·—”«∆· «·≈” Ì—«œÌ…"
   ClientHeight    =   10290
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   4920
      Begin VB.CommandButton cmdPrint 
         Height          =   555
         Left            =   2430
         Picture         =   "VsPayImp.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   555
         Left            =   45
         Picture         =   "VsPayImp.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   555
         Left            =   3600
         Picture         =   "VsPayImp.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExel 
         Height          =   555
         Left            =   1260
         Picture         =   "VsPayImp.frx":6D88
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   7500
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   765
         Width           =   4875
         Begin VB.OptionButton xpay 
            Alignment       =   1  'Right Justify
            Caption         =   "«·€Ì— „”œœ ›ﬁÿ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   2
            Left            =   45
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   225
            Width           =   1500
         End
         Begin VB.OptionButton xpay 
            Alignment       =   1  'Right Justify
            Caption         =   "«·„”œœ ›ﬁÿ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   2025
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   225
            Width           =   1230
         End
         Begin VB.OptionButton xpay 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄—÷ «·ﬂ·"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   225
            Value           =   -1  'True
            Width           =   1185
         End
      End
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
         TabIndex        =   9
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
         TabIndex        =   7
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
         TabIndex        =   11
         Top             =   622
         Width           =   2505
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·„Ê—œ :"
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
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   690
         Width           =   885
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… „Ê—œÌ‰ "
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
         TabIndex        =   8
         Top             =   1050
         Width           =   1365
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
         RightToLeft     =   -1  'True
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
      Top             =   9960
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
      Bindings        =   "VsPayImp.frx":9573
      Height          =   8010
      Left            =   240
      TabIndex        =   6
      Top             =   1575
      Width           =   18465
      _cx             =   32570
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
      BackColorSel    =   14220542
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
      Left            =   6075
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
      Left            =   3750
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
Attribute VB_Name = "VsTBALIMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastSalTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim CON As New ADODB.Connection
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï —”«∆· ≈” Ì—«œÌ…"
    If xGrCust.Text <> "" Then cHead1 = cHead1 & xGrCust.Text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xDate2.Text, "DD-MM-YYYY")
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 1, -1, cHead1, cHead2, , False, True, 9
    PrintGrd.Show 1
End Sub

Private Sub cmdExel_Click()
    ToFileExel grid1

End Sub

Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub cmdundo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    MyLoad
End Sub

Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï —”«∆· ≈” Ì—«œÌ…"
    If xGrCust.Text <> "" Then cHead1 = cHead1 & xGrCust.Text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xDate2.Text, "DD-MM-YYYY")
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 1, -1, cHead1, cHead2, , False, True, 9
    PrintGrd.Show 1

End Sub

Private Sub Form_Load()
    openCon CON
    
    xDate1.Text = "1-1-" & Year(Date)
    xDate2.Text = Format(Date, "dd-mm-yyyy")
    
    data1.ConnectionString = strCon
    data1.RecordSource = "SELECT * FROM FILE4_50 "
    Set xGrCust.RowSource = data1
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA2
    DATA2.ConnectionString = strCon
    
    FixGrid
    grid1.Rows = 1
End Sub
Private Sub MyLoad()
Dim cString As String
With grid1
'                           0                           1
    cString = "  SELECT     CODE, DESCA, CUURNAME, date, DOC_NO, date_pay, TOTALFRGN, T_PAY FROM         Q_IMP "
    If xGrCust.BoundText <> "" Then cString = cString & turn(cString) & " [group]  = " & MyParn(xGrCust.BoundText)
    If Xcode.Text <> "" Then cString = cString & turn(cString) & " [CODE] = " & MyParn(Xcode.Text)
    If IsDate(xDate1.Text) Then cString = cString & turn(cString) & " DATE >= " & DateSq(xDate1.Text)
    If IsDate(xDate2.Text) Then cString = cString & turn(cString) & " DATE <= " & DateSq(xDate2.Text)
    If xpay(1).Value <> 0 Then cString = cString & turn(cString) & " TOTALFRGN = T_PAY "
    If xpay(2).Value <> 0 Then cString = cString & turn(cString) & " TOTALFRGN > T_PAY "
    cString = cString & " ORDER BY CODE , DATE "
    DATA2.RecordSource = cString
    DATA2.Refresh
End With
FixGrid
End Sub
Sub FixGrid()
    With grid1
    .Cols = 9
    .RowHeight(0) = 1000
    .WordWrap = True
    .FrozenCols = 2
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    .TextMatrix(0, 2) = "«·⁄„·…"
    .TextMatrix(0, 3) = "«· «—ÌŒ"
    .TextMatrix(0, 4) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 5) = " «—ÌŒ «·≈” Õﬁ«ﬁ"
    .TextMatrix(0, 6) = "ﬁÌ„… «·—”«·…"
    .TextMatrix(0, 7) = "«·„”œœ"
    .TextMatrix(0, 8) = "«·»«ﬁÏ"
    
    .ColWidth(0) = 900
    .ColWidth(1) = 4000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1300
    .ColWidth(4) = 1000
    .ColWidth(5) = 1300
    .ColWidth(6) = 1300
    .ColWidth(7) = 1300
    .ColWidth(8) = 1300
    .SubtotalPosition = flexSTBelow
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 8) = Round(Val(.TextMatrix(nRow, 6)) - Val(.TextMatrix(nRow, 7)), 2)
    Next nRow
    For nCol = 6 To .Cols - 1
        .ColDataType(nCol) = flexDTDouble
        .ColFormat(nCol) = "#0.00"
        .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, vbRed, True, " "
        .Subtotal flexSTSum, 0, nCol, "#0.00", vbYellow, vbRed, True, " "
    Next nCol
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon CON
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
    xCustName.Caption = ""
    If Xcode.Text = "" Then Exit Sub
    Xcode.Text = RetZero(Xcode.Text, 6)
    xCustName.Caption = GetDesca("select desca from FILE4_10 where code = " & MyParn(Xcode.Text)) & ""
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
Generalarray(1) = "Select Code, DescA From FILE4_10"
Generalarray(2) = "Order by file4_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·„Ê—œ"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub
