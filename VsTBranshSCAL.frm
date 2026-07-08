VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Begin VB.Form VsTBranshSCAL 
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTBranshSCAL.frx":0000
      Height          =   3960
      Left            =   45
      TabIndex        =   0
      Top             =   630
      Width           =   20085
      _cx             =   35428
      _cy             =   6985
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
End
Attribute VB_Name = "VsTBranshSCAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StoreTable As New adodb.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New adodb.Connection
Private Sub Form_Load()
    openCon con
    
    StoreTable.Open "select FILE0_40.CODE , FILE0_40.DESCA FROM FILE0_40 ", con, adOpenStatic, adLockReadOnly, adCmdText
    myload
End Sub
Private Sub myload()
Dim i As Double
Dim cStrall  As String

grid1.Rows = 0
grid1.Rows = 2
grid1.FixedRows = 2
grid1.Cols = 7

Dim lOk As Boolean
StoreTable.MoveFirst
Do Until StoreTable.EOF
    cwhere = " STORE = " & MyParn(StoreTable!CODE)
    cField2 = cField2 & turnFound(cField2, ",") & myiif(cwhere, "QUANT")
    cField2 = cField2 & turnFound(cField2, ",") & myiif(cwhere, "(TOTAL * ((100-SALES_MODEL.RATE)/100))")
    StoreTable.MoveNext
Loop

With grid1
'                           0               1                 2                3               4                        5
    cStrall = " SELECT  SALES_MODEL.CODE, SALES_MODEL.SUPPDESCA, SUM(QUANT) , SUM(TOTAL * ((100-SALES_MODEL.RATE)/100))  , " & _
                cField2 & _
            " FROM  SALES_MODEL INNER JOIN FILE1_10 ON FILE1_10.ITEM = SALES_MODEL.ITEM WHERE SALES_MODEL.CODE IS NOT NULL  "
    If xMosm.BoundText <> "" Then cStrall = cStrall & " AND SALES_MODEL.MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cStrall = cStrall & " AND SALES_MODEL.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStrall = cStrall & " AND SALES_MODEL.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStrall = cStrall & " AND SALES_MODEL.FACT = " & MyParn(xFact.BoundText)
    If xSection.BoundText <> "" Then cStrall = cStrall & " AND FILE1_10.[Section] = " & xSection.BoundText
    If IsDate(xDate1.Text) Then cStrall = cStrall & " AND SALES_MODEL.[DATE] >= " & DateSq(xDate1.Text)
    If IsDate(xDate1.Text) Then cStrall = cStrall & " AND SALES_MODEL.[DATE] <= " & DateSq(xDate2.Text)
    cStrall = cStrall & " GROUP BY  SALES_MODEL.CODE, SALES_MODEL.SUPPDESCA ORDER BY SALES_MODEL.CODE "
    DATA6.RecordSource = cStrall
    DATA6.Refresh
End With
FixGrid
If grid1.Rows > 0 Then grid1.TextMatrix(1, 1) = "«·≈Ã„«·Ï"
End Sub
Sub FixGrid()
Dim nColS1 As Double
Dim nColS2 As Double

With grid1
    .RowHeight(0) = 700
    .RowHeight(0) = 700
    .WordWrap = True
    .MergeCells = flexMergeFree
    .MergeRow(0) = True
    .MergeRow(1) = True
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    
    .TextMatrix(1, 0) = "ﬂÊœ"
    .TextMatrix(1, 1) = "«·„Ê—œ"
    
    .TextMatrix(0, 2) = "≈Ã„«·Ï"
    .TextMatrix(0, 3) = "≈Ã„«·Ï"
    
    .TextMatrix(1, 2) = "⁄œœ"
    .TextMatrix(1, 3) = "≈Ã„«·Ï"
    
    .ColWidth(0) = 500
    .ColWidth(1) = 2500
    .ColWidth(2) = 1000
    .ColWidth(3) = 1500
    
    .ColDataType(2) = flexDTDouble
    .ColDataType(3) = flexDTDouble
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    nCol = 4
    StoreTable.MoveFirst
    Do Until StoreTable.EOF
        .TextMatrix(0, nCol) = StoreTable!desca
        .TextMatrix(0, nCol + 1) = StoreTable!desca
        .ColWidth(nCol) = 1000
        .ColWidth(nCol + 1) = 1300
        
        .TextMatrix(1, nCol) = "⁄œœ"
        .TextMatrix(1, nCol + 1) = "≈Ã„«·Ï"
        nCol = nCol + 2
        StoreTable.MoveNext
    Loop
    
    For nRow = 2 To .Rows - 1
        For nCol = 2 To .Cols - 2 Step 2
            .TextMatrix(nRow, nCol) = Format(Val(.TextMatrix(nRow, nCol)), "#0")
            .TextMatrix(nRow, nCol + 1) = Format(Val(.TextMatrix(nRow, nCol + 1)), "#0.00")
        Next nCol
    Next nRow
    .SubtotalPosition = flexSTAbove
    For i = 2 To .Cols - 1
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
    End With
End Sub


