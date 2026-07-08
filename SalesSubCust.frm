VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form SalesSubCust 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   15270
   Begin VB.CommandButton CMR_PRINT 
      Caption         =   "ØÈÇÚÉ "
      Height          =   555
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   6660
      Width           =   1860
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   90
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
      Bindings        =   "SalesSubCust.frx":0000
      Height          =   6345
      Left            =   90
      TabIndex        =   1
      Top             =   225
      Width           =   15045
      _cx             =   26538
      _cy             =   11192
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
End
Attribute VB_Name = "SalesSubCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CMR_PRINT_Click()
    Load PrintGrd
    PrintGrd.doprint grid1, 1, , Me.Caption, , , , True, 8
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()
    openCon con
    With VsTBalCustSub.grid1
        Me.Caption = "ÊÝÕíáì ÝæÇÊíÑ ãÈíÚÇÊ ááÚãíá " & .TextMatrix(.Row, 1)
        Me.Caption = Me.Caption & " ãä ÊÇÑíÎ " & VsTBalCustSub.xDate1.Text
        Me.Caption = Me.Caption & " Åáì ÊÇÑíÎ " & VsTBalCustSub.xdate2.Text
'                               0               1                       2                   3                          4                       5            6                       7                                               8                                                                               9               10                  11                                                                              12
        cString = " SELECT  T_PRICE_CUST2.DATE , T_PRICE_CUST2.DOC_ID , T_PRICE_CUST2.DOC_NO2 , T_PRICE_CUST2.QUANT, T_PRICE_CUST2.T_PRICE, T_PRICE_CUST2.TOTAL, T_PRICE_CUST2.DISCOUNT , T_PRICE_CUST2.TOTAL - T_PRICE_CUST2.discount , T_PRICE_CUST2.T_PRICE - T_PRICE_CUST2.TOTAL - T_PRICE_CUST2.discount AS T_DISCOUNT, ' ' AS N8 ,T_PRICE_CUST2.T_COST, T_PRICE_CUST2.TOTAL - T_PRICE_CUST2.discount - T_PRICE_CUST2.T_COST AS T_PROFT , ' ' AS N11 " & _
        " FROM T_PRICE_CUST2 LEFT OUTER JOIN FILE0_40 ON T_PRICE_CUST2.store = FILE0_40.CODE " & _
        " where CODE2 = " & MyParn(.TextMatrix(.Row, 0))
'        If IsDate(VsTBalCustSub.xDate1.Text) Then cString = cString & " AND DATE >= " & DateSq(VsTBalCustSub.xDate1.Text)
'        If IsDate(VsTBalCustSub.xDate2.Text) Then cString = cString & " AND DATE <= " & DateSq(VsTBalCustSub.xDate2.Text)
        cString = cString & " ORDER BY T_PRICE_CUST2.DATE , T_PRICE_CUST2.DOC"
    End With
    
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    data1.RecordSource = cString
    data1.Refresh
    
    With grid1
        .MergeCells = flexMergeFree
        .MergeCol(0) = True
        .Editable = flexEDNone
        .ColWidth(0) = 1500
        .ColWidth(1) = 0
        .ColWidth(2) = 800
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        
        .RowHeight(0) = 800
        .WordWrap = True
        .TextMatrix(0, 0) = "ÊÇÑíÎ"
        .TextMatrix(0, 1) = "ÑÞã ãÓÊäÏ"
        .TextMatrix(0, 2) = "ãÓÊäÏ"
        .TextMatrix(0, 3) = "ÇáÚÏÏ"
        .TextMatrix(0, 4) = "ÅÌãÇáì ãÓÊåáß"
        .TextMatrix(0, 5) = "ÅÌãÇáì ÇÕäÇÝ"
        .TextMatrix(0, 6) = "ÎÕã Èæä"
        .TextMatrix(0, 7) = "ÕÇÝì ÇáÈæä"
        .TextMatrix(0, 8) = "ÅÌãÇáì ÇáÎÕã"
        .TextMatrix(0, 9) = "äÓÈÉ ÎÕã"
        .TextMatrix(0, 10) = "ÊßáÝÉ ÃÕäÇÝ"
        .TextMatrix(0, 11) = "ÑÈÍ ÇáÈæä"
        .TextMatrix(0, 12) = "ÑÕíÏ ãÈíÚÇÊ"
        nTot = 0
        For nRow = 1 To .Rows - 1
            .TextMatrix(nRow, 9) = Format(Val(.TextMatrix(nRow, 8)) / Val(.TextMatrix(nRow, 4)) * 100, "#0.00")
            nTot = nTot + Val(.TextMatrix(nRow, 7))
            .TextMatrix(nRow, 12) = nTot
        Next nRow
        
        .SubtotalPosition = flexSTAbove
        .ColDataType(0) = flexDTDate
        For nCol = 3 To .Cols - 2
            .ColDataType(nCol) = flexDTDouble
            .Subtotal flexSTSum, -1, nCol, "#0", vbGreen, , True, " "
        Next nCol
        nRow = 1
        If Val(.TextMatrix(nRow, 4)) <> 0 Then .TextMatrix(nRow, 9) = Format(Val(.TextMatrix(nRow, 8)) / Val(.TextMatrix(nRow, 4)) * 100, "#0.00")
        If Val(.TextMatrix(nRow, 10)) <> 0 Then .TextMatrix(nRow, 12) = Format(Val(.TextMatrix(nRow, 11)) / Val(.TextMatrix(nRow, 10)) * 100, "#0.00")
    End With
End Sub
