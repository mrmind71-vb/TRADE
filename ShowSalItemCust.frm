VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ShowSalItemCust 
   Caption         =   " ›’Ì·Ï"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   5595
   ScaleWidth      =   11310
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   4200
      Top             =   5100
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin VB.CommandButton cmd_exit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4950
      Width           =   2715
   End
   Begin VB.CommandButton CMD_PRINT 
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8400
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4950
      Width           =   2715
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "ShowSalItemCust.frx":0000
      Height          =   4710
      Left            =   150
      TabIndex        =   2
      Top             =   75
      Width           =   11040
      _cx             =   19473
      _cy             =   8308
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
End
Attribute VB_Name = "ShowSalItemCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cHead As String
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = " ›’Ì·Ï „»Ì⁄«  Ê „— Ã⁄«  ’‰› " & VsTCustSales.grid1.TextMatrix(VsTCustSales.grid1.Row, 0) & "  " & VsTCustSales.grid1.TextMatrix(VsTCustSales.grid1.Row, 1)
    cHead2 = " „‰  «—ÌŒ " & Format(VsTCustSales.xDate1.Text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(VsTCustSales.xDate2.Text, "DD-MM-YYYY")
    If VsTCustSales.xCustName.Caption <> "" Then cHead2 = cHead2 & " " & VsTCustSales.xCustName.Caption
    If VsTCustSales.xGrCust.BoundText <> "" Then cHead2 = cHead2 & " " & VsTCustSales.xGrCust.Text
    
    Load PrintGrd
    PrintGrd.doprint grid1, 1, -1, cHead1, cHead2, , False, , 10
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()
Dim cItem As String
Dim DDate1 As Date
Dim DDate2 As Date
Dim cCode As String
Dim cGrCust As String
cItem = VsTCustSales.grid1.TextMatrix(VsTCustSales.grid1.Row, 1)
DDate1 = VsTCustSales.xDate1.Text
DDate2 = VsTCustSales.xDate2.Text
cCode = VsTCustSales.Xcode.Text
cGrCust = VsTCustSales.xGrCust.BoundText
'               1                       2               3               4                       5                       6     7                                  8                               8                       9              10
cStrall = " SELECT FILE1_11.date, FILE3_10.DESCA, FILE1_11.DOC_ID, FILE1_11.OUT AS sal, FILE1_11.[IN] AS ret, FILE1_11.PRICE , FILE1_11.DISCOUNT AS discount, FILE1_11.TOTAL AS total, case when [type]='6' then 1 else -1 end AS rATE FROM FILE1_11 LEFT JOIN FILE3_10 ON FILE1_11.CODECUST = FILE3_10.CODE  WHERE (FILE1_11.TYPE = '3' Or FILE1_11.TYPE = '6' ) AND ITEM = " & MyParn(cItem)
If IsDate(DDate1) Then
    cStrall = cStrall & turn(cStrall) & " DATE >= " & DateSq(DDate1)
End If
If IsDate(DDate2) Then
    cStrall = cStrall & turn(cStrall) & " DATE <= " & DateSq(DDate2)
End If

If cCode <> "" Then cStrall = cStrall & " AND FILE3_10.CODE = " & MyParn(cCode)
If cGrCust <> "" Then cStrall = cStrall & " AND FILE3_10.[GROUP] = " & MyParn(cGrCust)
cStrall = cStrall & " ORDER BY DATE "
Set grid1.DataSource = data1
data1.ConnectionString = strCon
data1.RecordSource = cStrall
data1.Refresh
FixGrid


End Sub
Sub FixGrid()
With grid1
    .Cols = 9
    
     
    .RowHeight(0) = 800
    .WordWrap = True
    
    .TextMatrix(0, 0) = " «—ÌŒ"
    .TextMatrix(0, 1) = "«·⁄„Ì·"
    .TextMatrix(0, 2) = "—ﬁ„ „” ‰œ"
    .TextMatrix(0, 3) = "„»Ì⁄« "
    .TextMatrix(0, 4) = "„—œÊœ „»Ì⁄« "
    .TextMatrix(0, 5) = "”⁄—"
    .TextMatrix(0, 6) = "‰”»… Œ’„"
    .TextMatrix(0, 7) = "≈Ã„«·Ï ﬁÌ„…"
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 2000
    .ColWidth(2) = 1000
    .ColWidth(3) = 900
    .ColWidth(4) = 900
    .ColWidth(5) = 900
    .ColWidth(6) = 900
    .ColWidth(7) = 900
    .ColWidth(8) = 0
    For I = 1 To .Rows - 1
        .TextMatrix(I, 7) = Format(Val(.TextMatrix(I, 7)) * Val(.TextMatrix(I, 8)), "#0.00")
    Next I
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 3, "#0", "#0", vbRed, vbYellow, " "
    .Subtotal flexSTSum, -1, 4, "#0", "#0", vbRed, vbYellow, " "
    .Subtotal flexSTSum, -1, 7, "#0", "#0", vbRed, vbYellow, " "

    .ColDataType(3) = flexDTDouble
    .ColDataType(4) = flexDTDouble
    .ColDataType(5) = flexDTDouble
    .ColDataType(6) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ExplorerBar = flexExSortShow
End With
End Sub




