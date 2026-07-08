VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form VsTfrBranshSalModel_SCAL 
   ClientHeight    =   9405
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   18960
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   18960
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "⁄—÷ ﬂ· «·„ÊœÌ·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   21150
      TabIndex        =   3
      Top             =   5895
      Width           =   2535
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTfrBranshSalModel_SCAL.frx":0000
      Height          =   5760
      Left            =   45
      TabIndex        =   0
      Top             =   135
      Width           =   23640
      _cx             =   41698
      _cy             =   10160
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   12648384
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
   Begin VSFlex7Ctl.VSFlexGrid GRID2 
      Bindings        =   "VsTfrBranshSalModel_SCAL.frx":0014
      Height          =   5985
      Left            =   12060
      TabIndex        =   1
      Top             =   6255
      Width           =   11625
      _cx             =   20505
      _cy             =   10557
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   8454016
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
   Begin VSFlex7Ctl.VSFlexGrid GRID3 
      Bindings        =   "VsTfrBranshSalModel_SCAL.frx":0028
      Height          =   5985
      Left            =   90
      TabIndex        =   2
      Top             =   6255
      Width           =   11850
      _cx             =   20902
      _cy             =   10557
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   16777152
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
   Begin MSAdodcLib.Adodc DATA3 
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
Attribute VB_Name = "VsTfrBranshSalModel_SCAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sModel As String, sDate1 As Date, sDate2 As Date, cRStore As String
Dim StoreTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection
Private Sub Command1_Click()
    MYLOAD2
End Sub
Private Sub Form_Load()
    openCon con
    
    Set grid1.DataSource = data6
    data6.ConnectionString = strCon
    
    Set grid2.DataSource = DATA2
    DATA2.ConnectionString = strCon
    
    Set grid3.DataSource = data3
    data3.ConnectionString = strCon
    
    grid1.Rows = 2
    grid1.FixedRows = 2
    grid1.Cols = 7
    
'    StoreTable.Open "select STORE_BR.CODE , STORE_BR.DESCA FROM STORE_BR ", con, adOpenStatic, adLockReadOnly, adCmdText
    
    myload

    MYLOAD2
    FixGrid2
End Sub
Sub myload()
Dim i As Double
Dim cStrall  As String
Dim ScalTable As New ADODB.Recordset
'On Error GoTo myerror
ScalTable.Open "select FILE1_10.COLOR , FILE1_10.SCAL  , ITEM FROM FILE1_10 WHERE MODELNO = " & MyParn(sModel) & " ORDER BY COLOR , C_SCAL ", con, adOpenStatic, adLockReadOnly, adCmdText

grid1.Rows = 0
grid1.Rows = 4
grid1.FixedRows = 4

Dim lOk As Boolean
ScalTable.MoveFirst

Do Until ScalTable.EOF
    cWhere = " FILE1_10.ITEM = " & MyParn(ScalTable!Item)
    cWhere = cWhere & " AND DATE <= " & DateSq(sDate2)
    cField4 = cField4 & " , " & myiif(cWhere, "([IN] - [OUT])") & " AS " & addstring(ScalTable!Item)
    ScalTable.MoveNext
Loop
With grid1
    
    cField1 = " , " & myiif("(type = '2'  OR type = 'F'  ) ", "[IN]") & "   AS t1 "
    cField2 = " , " & myiif("(type = '7'  OR type = 'T'  ) ", "[OUT]") & "   AS t2 "
    cField3 = " , " & myiif("(type = '6' ) ", "[OUT]") & " AS t3 "
    
    cStrall = " SELECT STORE_BR.CODE , STORE_BR.desca " & cField1 & cField2 & cField3 & cField4 & _
                " , sum([in] - [OUT]) as bal " & _
            " FROM  STORE_BR INNER JOIN FILE1_11_ALL ON STORE_BR.CODE = FILE1_11_ALL.STORE INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11_ALL.ITEM WHERE FILE1_10.MODELNO = " & MyParn(sModel)
    cStrall = cStrall & " GROUP BY STORE_BR.CODE , STORE_BR.desca ORDER BY STORE_BR.CODE "
    data6.RecordSource = cStrall
    data6.Refresh
End With
FIXGRID
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub FIXGRID()
Dim nColS1 As Double
Dim nColS2 As Double
Dim ScalTable As New ADODB.Recordset
ScalTable.Open "select FILE1_10.COLOR , FILE1_10.SCAL  , ITEM FROM FILE1_10 WHERE MODELno = " & MyParn(sModel) & " ORDER BY COLOR , C_SCAL ", con, adOpenStatic, adLockReadOnly, adCmdText

With grid1
'    .FixedRows = 2
    
    .ExplorerBar = flexExSortShow
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·›—⁄"
    .TextMatrix(0, 2) = "Ê—«œ"
    .TextMatrix(0, 3) = "’«œ—"
    .TextMatrix(0, 4) = "„»Ì⁄« "
    .TextMatrix(0, .Cols - 1) = "«·—’Ìœ"
    
    .TextMatrix(1, 0) = "ﬂÊœ"
    .TextMatrix(1, 1) = "«·›—⁄"
    .TextMatrix(1, 2) = "Ê—«œ"
    .TextMatrix(1, 3) = "’«œ—"
    .TextMatrix(1, 4) = "„»Ì⁄« "
    .TextMatrix(1, .Cols - 1) = "«·—’Ìœ"
    
    .TextMatrix(2, 0) = "ﬂÊœ"
    .TextMatrix(2, 1) = "«·›—⁄"
    .TextMatrix(2, 2) = "Ê—«œ"
    .TextMatrix(2, 3) = "’«œ—"
    .TextMatrix(2, 4) = "„»Ì⁄« "
    .TextMatrix(2, .Cols - 1) = "«·—’Ìœ"
    
    .RowHidden(3) = True
    .ColWidth(0) = 700
    .ColWidth(1) = 2000
    
    .ColWidth(2) = 800
    .ColWidth(3) = 800
    .ColWidth(4) = 800
    
    .WordWrap = True
    .MergeCells = flexMergeFixedOnly
    nCol = 4
    ScalTable.MoveFirst
    Do Until ScalTable.EOF
        nCol = nCol + 1
        .ColWidth(nCol) = 800
        .TextMatrix(0, nCol) = "«·—’Ìœ"
        .TextMatrix(1, nCol) = ScalTable!Color
        .TextMatrix(2, nCol) = ScalTable!scal
        .TextMatrix(3, nCol) = ScalTable!Item
        ScalTable.MoveNext
    Loop
    .ColWidth(.Cols - 1) = 800
        
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(.Cols - 1) = True
    
    .MergeRow(0) = True
    .MergeRow(1) = True
    .MergeRow(2) = True
    
    .SubtotalPosition = flexSTBelow
    For nCol = 2 To .Cols - 2
        .Subtotal flexSTSum, -1, nCol, "#", vbRed, , True
    Next nCol
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End With
End Sub
Sub FixGrid2()
With grid2
    .Cols = 7
    .RowHeight(0) = 500
    .WordWrap = True
    .FixedRows = 1
    .FixedCols = 0
    .MergeCells = flexMergeFixedOnly
    .ExplorerBar = flexExSortShow
    
    .TextMatrix(0, 0) = "»«—ﬂÊœ"
    .TextMatrix(0, 1) = "·Ê‰"
    .TextMatrix(0, 2) = "„ﬁ«”"
    
    .TextMatrix(0, 3) = "„‰ ›—⁄"
    .TextMatrix(0, 4) = "≈·Ï ›—⁄"
    .TextMatrix(0, 5) = "«·ﬂ„Ì…"
    
    .ColWidth(0) = 1500
    .ColWidth(1) = 1200
    .ColWidth(2) = 1200
    .ColWidth(3) = 2000
    .ColWidth(4) = 2000
    .ColWidth(5) = 1000
    .ColHidden(6) = True
    .ColComboList(3) = StrList("SELECT CODE , DESCA FROM STORE_BR ORDER BY S_BRANCH ", con)
    .ColComboList(4) = StrList("SELECT CODE , DESCA FROM STORE_BR ORDER BY S_BRANCH ", con)
    End With
End Sub
Sub MYLOAD2(Optional pItem As Double)
With grid1
'                           0               1                 2                3               4                        5           6
    cStrall = " SELECT SUBTRANS.ITEM , FILE1_10.COLOR , FILE1_10.SCAL , SUBTRANS.STORE1 , SUBTRANS.STORE2 , SUBTRANS.QUANT , SUBTRANS.ID FROM FILE1_10 INNER JOIN SUBTRANS ON SUBTRANS.ITEM = FILE1_10.ITEM WHERE FILE1_10.MODELno = " & MyParn(sModel)
    If pItem <> 0 Then cStrall = cStrall & " AND SUBTRANS.ITEM = " & pItem
    cStrall = cStrall & " ORDER BY SUBTRANS.ID "
    DATA2.RecordSource = cStrall
    DATA2.Refresh
End With
FixGrid2
End Sub
Private Sub grid1_DblClick()
Dim nQunat As Double, cStore As String, nItem As Double
cRStore = grid1.TextMatrix(grid1.Row, 0)
With grid2
    If grid1.Col >= 5 And grid1.Col < grid1.Cols - 1 Then
        nItem = Val(grid1.TextMatrix(3, grid1.Col))
        MYLOAD2 nItem
        nQunat = Val(GetDesca("SELECT SUM(QUANT) FROM SUBTRANS WHERE ITEM = " & nItem & " AND STORE1 = " & MyParn(cRStore), con) & "")
        If Val(grid1.TextMatrix(grid1.Row, grid1.Col)) - nQunat > 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = grid1.TextMatrix(3, grid1.Col)
            .TextMatrix(.Rows - 1, 1) = grid1.TextMatrix(1, grid1.Col)
            .TextMatrix(.Rows - 1, 2) = grid1.TextMatrix(2, grid1.Col)
            .TextMatrix(.Rows - 1, 3) = cRStore
            MYLOAD3 grid1.TextMatrix(3, grid1.Col)
        Else
            MYLOAD3 grid1.TextMatrix(3, grid1.Col)
            Inform " ·« ÌÊÃœ —’Ìœ "
        End If
    End If
End With
End Sub
Private Sub grid2_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
Dim nItem As Double
Dim cStore1 As String
Dim cStore2 As String, nQ As Double
Dim aInsert As Variant
With grid2
    
    nItem = .TextMatrix(.Row, 0)
    cStore1 = .TextMatrix(.Row, 3)
    cStore2 = .TextMatrix(.Row, 4)
    nQ = Val(.TextMatrix(.Row, 5))
        
    If nQ > 0 And cStore1 <> "" And cStore2 <> "" And cStore1 <> cStore2 Then
        aInsert = AddFlag(Empty, "ITEM", addvalue(nItem))
        aInsert = AddFlag(aInsert, "MODEL", addstring(sModel))
        aInsert = AddFlag(aInsert, "STORE1", addstring(cStore1))
        aInsert = AddFlag(aInsert, "STORE2", addstring(cStore2))
        aInsert = AddFlag(aInsert, "QUANT", (nQ))
        If .TextMatrix(.Row, .Cols - 1) = "" Then
            NROW3 = grid3.FindRow(cStore1, 1, 1)
            If NROW3 > 0 Then
                If nQ <= (Val(grid3.TextMatrix(NROW3, 4)) - Val(grid3.TextMatrix(NROW3, 5))) Then
                    con.Execute addInsert(aInsert, "SUBTRANS")
                Else
                    Inform "«·—’Ìœ €Ì— ﬂ«›Ï "
                End If
            End If
        Else
            NROW3 = grid3.FindRow(cStore1, 1, 1)
            If NROW3 > 0 Then
                If nQ >= (Val(grid3.TextMatrix(NROW3, 4)) - Val(grid3.TextMatrix(NROW3, 5))) Then
                    con.Execute addUpdate(aInsert, "SUBTRANS", " ID = " & Val(.TextMatrix(.Row, .Cols - 1)))
                Else
                    Inform "«·—’Ìœ €Ì— ﬂ«›Ï "
                End If
            End If
        End If
        MYLOAD2 nItem
'        .AddItem ""
        MYLOAD3 nItem
    Else
    If nQ > 0 Then Inform "”Ã· €Ì— ’ÕÌÕ "
    End If
End With
Exit Sub
myerror:
MsgBox Err.Description
grid2.RemoveItem grid2.Row
Err.Clear
End Sub
Private Sub grid2_EnterCell()
With grid2
    If .Col >= 4 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub Grid2_KeyUp(KeyCode As Integer, Shift As Integer)
With grid2
    If KeyCode = 46 Then
        If MsgBox(" Õ–› «·”Ã·  ", vbYesNo + vbDefaultButton2) = vbYes Then
            If Val(.TextMatrix(.Row, .Cols - 1)) > 0 Then
                con.Execute " DELETE FROM SUBTRANS WHERE ID = " & Val(.TextMatrix(.Row, .Cols - 1))
                .RemoveItem .Row
            End If
        End If
    End If
End With
End Sub
Sub FixGrid3()
With grid3
    .Cols = 7
    .ExplorerBar = flexExSortShow
    .RowHeight(0) = 500
    .WordWrap = True
    .FixedRows = 1
    .FixedCols = 0
    .MergeCells = flexMergeFixedOnly
    .TextMatrix(0, 0) = "»«—ﬂÊœ"
    .TextMatrix(0, 1) = "«·›—⁄"
    .TextMatrix(0, 2) = "Ê«—œ"
    .TextMatrix(0, 3) = "„»Ì⁄« "
    .TextMatrix(0, 4) = "—’Ìœ"
    .TextMatrix(0, 5) = "’«œ—  Ê“Ì⁄"
    .TextMatrix(0, 6) = "Ê«—œ  Ê“Ì⁄"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 2000
    .ColWidth(2) = 800
    .ColWidth(3) = 800
    .ColWidth(4) = 800
    .ColWidth(5) = 800
    .ColWidth(6) = 800
    .ColComboList(1) = StrList("SELECT CODE , DESCA FROM STORE_BR", con)
    For nRow = 1 To .Rows - 1
        If .TextMatrix(nRow, 1) = cRStore Then
            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
        End If
    Next nRow
    End With
End Sub
Sub MYLOAD3(pItem As Double)
    cWhere = " ( FILE1_11_ALL.TYPE = '2' OR FILE1_11_ALL.TYPE = '7'  OR FILE1_11_ALL.TYPE = 'F'  OR FILE1_11_ALL.TYPE = 'T' ) "
    cWhere = cWhere & " AND ( DATE >= " & DateSq(sDate1) & " AND DATE <= " & DateSq(sDate2) & " ) "
    cField2 = myiif(cWhere, "[IN] - [OUT]")
    
    cWhere = " FILE1_11_ALL.TYPE = '6' "
    cWhere = cWhere & " AND ( DATE >= " & DateSq(sDate1) & " AND DATE <= " & DateSq(sDate2) & " ) "
    cField2 = cField2 & " , " & myiif(cWhere, "[OUT] ")
    
    cWhere = " ( DATE <= " & DateSq(sDate2) & " ) "
    cField2 = cField2 & " , " & myiif(cWhere, "[IN] - [OUT]")
    
    
    cStrall = " SELECT FILE1_11_ALL.ITEM , FILE1_11_ALL.STORE, " & cField2 & ", (SELECT SUM(QUANT) FROM SUBTRANS WHERE ITEM = FILE1_11_ALL.ITEM AND STORE1 = FILE1_11_ALL.STORE ) " & ", (SELECT SUM(QUANT) FROM SUBTRANS WHERE ITEM = FILE1_11_ALL.ITEM AND STORE2 = FILE1_11_ALL.STORE ) " & _
            " FROM  FILE1_11_ALL WHERE ITEM = " & pItem
    
    cStrall = cStrall & " GROUP BY ITEM , STORE ORDER BY STORE "
    data3.RecordSource = cStrall
    data3.Refresh
    FixGrid3
End Sub
Private Sub grid3_DblClick()
Exit Sub
Dim nQunat As Double, cStore As String
cStore = grid3.TextMatrix(grid3.Row, 1)
With grid2
    MYLOAD2 grid3.TextMatrix(grid3.Row, 0)
'    nQunat = Val(GetDesca("SELECT SUM(QUANT) FROM SUBTRANS WHERE ITEM = " & Val(GRID3.TextMatrix(GRID3.Row, 0))) & " AND STORE1 = " & MyParn(GRID3.TextMatrix(GRID3.Row, 1)) & "")
    If Val(grid3.TextMatrix(grid3.Row, 4)) - Val(grid3.TextMatrix(grid3.Row, 5)) + Val(grid3.TextMatrix(grid3.Row, 6)) > 0 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = grid3.TextMatrix(grid3.Row, 0)
        .TextMatrix(.Rows - 1, 1) = grid1.TextMatrix(1, grid1.Col)
        .TextMatrix(.Rows - 1, 2) = grid1.TextMatrix(2, grid1.Col)
        .TextMatrix(.Rows - 1, 3) = .TextMatrix(.Rows - 2, 3)
        .TextMatrix(.Rows - 1, 4) = cStore
        MYLOAD3 grid3.TextMatrix(grid3.Row, 0)
    Else
        MYLOAD3 grid3.TextMatrix(grid3.Row, 0)
        Inform " ·« ÌÊÃœ —’Ìœ "
    End If
End With

End Sub
