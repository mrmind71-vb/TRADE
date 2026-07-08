VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ViewRetSalesBr 
   ClientHeight    =   6945
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   10065
   StartUpPosition =   3  'Windows Default
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6390
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   9825
      _cx             =   17330
      _cy             =   11271
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
      Rows            =   1
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
End
Attribute VB_Name = "ViewRetSalesBr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con2 As New ADODB.Connection
Public pBranch_Store As String, pDateRep1 As Date, pDateRep2 As Date
Private Sub Form_Load()
Dim cString As String
con2.CursorLocation = adUseClient
con2.Open strCon
Me.Caption = " ãÑÊÌÚÇÊ ÇáÝÑÚ ÊãÊ ãä ÝÑæÚ ÇÎÑì " & VSTSALESRETBR.grid1.TextMatrix(VSTSALESRETBR.grid1.Row, 1) & "   "
 
Set grid1.DataSource = data1
cString = " SELECT STORE_ALL.CODE , STORE_ALL.DESCA " & _
            " , COUNT(DISTINCT DOC_NO ) " & _
            " , SUM( QFILE6_20_RET.TOTAL * -1 )   AS    T_RETURN2 " & _
            " FROM QFILE6_20_RET INNER JOIN STORE_ALL ON STORE_ALL.CODE = QFILE6_20_RET.STORE_RET " & _
            " WHERE QFILE6_20_RET.STORE = " & MyParn(pBranch_Store) & "  AND QFILE6_20_RET.[DATE] >= " & DateSq(pDateRep1) & " AND QFILE6_20_RET.[DATE] <= " & DateSq(pDateRep2) & _
            " AND QFILE6_20_RET.STORE  <> QFILE6_20_RET.STORE_RET AND T_BOON < 0 " & _
            " GROUP BY STORE_ALL.CODE , STORE_ALL.DESCA  ORDER BY STORE_ALL.CODE "
Set data1.Recordset = myRecordSet(cString, con2)
With grid1
    .Cols = 4
    .ColWidth(0) = 1000
    .ColWidth(1) = 3000
    .ColWidth(2) = 1000
    .ColWidth(3) = 2000
    .TextMatrix(0, 0) = "ßæÏ"
    .TextMatrix(0, 1) = "ÇáÝÑÚ "
    .TextMatrix(0, 2) = "ÚÏÏ ÈæäÇÊ"
    .TextMatrix(0, 3) = "ÞíãÉ ÇáãÑÊÌÚ"
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 2, "#.00", vbYellow, , True, " "
    .Subtotal flexSTSum, -1, 3, "#.00", vbYellow, , True, " "
    .ExplorerBar = flexExSortShow
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con2
End Sub
