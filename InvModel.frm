VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form InvModel 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8250
   ScaleWidth      =   10530
   Begin VB.Frame Frame4 
      Caption         =   "„ﬁ«”«  Ê«·Ê«‰"
      Height          =   3930
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4005
      Width           =   10185
      Begin VSFlex7Ctl.VSFlexGrid grdModel 
         Height          =   3540
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   10005
         _cx             =   17648
         _cy             =   6244
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   4
         FixedCols       =   2
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
   Begin VSFlex7Ctl.VSFlexGrid grdModelCodes 
      Height          =   945
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
      _cx             =   2249
      _cy             =   1667
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      Rows            =   1
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "InvModel.frx":0000
      Height          =   3840
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   10185
      _cx             =   17965
      _cy             =   6773
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   12648384
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   300
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
      TabBehavior     =   0
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
Attribute VB_Name = "InvModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pModelNo As String
Dim con As New ADODB.Connection
Private Sub CMR_PRINT_Click()
    Load PrintGrd
    PrintGrd.doprint grid1, 1, , Me.Caption
    PrintGrd.Show 1
End Sub
Private Sub Form_Load()
    Dim cString As String
    openCon con
'    With VsTItem.grid1
'        Me.Caption = " ›’Ì·Ï ›Ê« Ì— „‘ —Ì«  Ê „— Ã⁄«  ··„ÊœÌ· —ﬁ„ " & .TextMatrix(.Row, 3) & " „’‰⁄ " & .TextMatrix(.Row, 1)
        cString = " select FILE1_11.date , FILE1_11.doc_id , FILE4_10.DESCA , SUM([IN]) , SUM([OUT]) FROM (FILE1_11 INNER JOIN FILE4_10 ON FILE4_10.CODE = FILE1_11.CODESUP ) INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM"
        cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(pModelNo)
        cString = cString & " GROUP BY date , doc_id , FILE4_10.DESCA ORDER BY date "
        
'    End With
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    Set data1.Recordset = myRecordSet(cString, con)
    With grid1
        .Editable = flexEDNone
        .ColWidth(0) = 1500
        .ColWidth(1) = 1000
        .ColWidth(2) = 3000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1000
        .TextMatrix(0, 0) = " «—ÌŒ"
        .TextMatrix(0, 1) = "—ﬁ„ „” ‰œ"
        .TextMatrix(0, 2) = "«·„Ê—œ"
        .TextMatrix(0, 3) = "„‘ —Ì« "
        .TextMatrix(0, 4) = "„— Ã⁄« "
        .SubtotalPosition = flexSTBelow
        For nRow = 1 To .Rows - 1
            If Val(.TextMatrix(nRow, 4)) > 0 Then
                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
            End If
        Next nRow
'        .Subtotal flexSTCount, -1, 1, "#0", vbGreen, , True, " "
        .Subtotal flexSTSum, -1, 3, "#0", vbYellow, , True, " "
        .Subtotal flexSTSum, -1, 4, "#0", vbYellow, , True, " "
    End With
    myloadGrdModel " FILE1_10.modelno = " & MyParn(pModelNo)
    If grid1.Rows > 1 Then
        grid1.Row = 1
        grid1.Col = 0
        GRID1_Click
    End If
End Sub
Private Function myloadGrdModel(pFilter As String) As Boolean
Dim aRet As Variant, cFieldas As String, cField As String, cFilter As String

mydefineGrdModel
aRet = retFields(pFilter)
If aRet(0) = "" Then Exit Function

cField = aRet(0)
cFieldas = aRet(1)

FillItem cFieldas, cField, pFilter
FixCost cFieldas, cField, pFilter
fixGrdModel
myloadgrd = True
End Function
Private Sub fixGrdModel()
With grdModel
.ColHidden(0) = True
'.ColWidth(0) = 500
.ColWidth(1) = 1300
nColWidth = (.Width - 200 - .ColWidth(0) - .ColWidth(1)) / .Cols
If nColWidth < 500 Then nColWidth = 500
If nColWidth > 1200 Then nColWidth = 1200
For nCol = 2 To .Cols - 1
    .ColWidth(nCol) = nColWidth
    .ColAlignment(nCol) = flexAlignCenterCenter
Next
.RowHidden(1) = Not bopt2
End With
End Sub
Private Sub mydefineGrdModel()
grdModel.Rows = 4
grdModel.Cols = 2

grdModel.MergeCells = flexMergeRestrictRows
grdModel.TextMatrix(0, 1) = "«·„ﬁ«”"
grdModel.TextMatrix(1, 1) = "”⁄— „’‰⁄"
grdModel.TextMatrix(2, 1) = "”⁄— „” Â·ﬂ"
grdModel.TextMatrix(3, 1) = "”⁄— «Êﬂ«“ÌÊ‰"
End Sub
Private Sub FillItem(cFieldas, cField, pFilter As String)
Dim GRDTABLE As New ADODB.Recordset, cString As String
' „·∆ «·ÃœÊ·
cString = "Select c_color as [—ﬁ„ «··Ê‰] ,color as [«··Ê‰] " & turn(cFieldas, ",") & cFieldas & _
          " From " & _
          " (Select c_color,Color,scal,item from file1_10 " & turn(pFilter) & pFilter & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(item)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  " & _
          " order by pvt.C_color"

GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grdModel.Cols = GRDTABLE.Fields.Count

For nCol = 2 To GRDTABLE.Fields.Count - 1
    grdModel.TextMatrix(0, nCol) = GRDTABLE.Fields(nCol).Name & ""
Next

Do Until GRDTABLE.EOF
    grdModel.AddItem ""
    For nCol = 0 To 1
        grdModel.TextMatrix(grdModel.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
    Next
    GRDTABLE.MoveNext
Loop
With grdModel
.AddItem ""
.TextMatrix(.Rows - 1, 1) = "≈Ã„«·Ì „ﬁ«”"
.Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HE0E0E0
.ColWidth(.Cols - 1) = 1000

.Cols = .Cols + 1
.TextMatrix(0, .Cols - 1) = "≈Ã„«·Ì ·Ê‰"
.TextMatrix(1, .Cols - 1) = "≈Ã„«·Ì ·Ê‰"
.TextMatrix(2, .Cols - 1) = "≈Ã„«·Ì ·Ê‰"
.TextMatrix(3, .Cols - 1) = "≈Ã„«·Ì ·Ê‰"
.MergeCells = flexMergeFixedOnly
.MergeCol(.Cols - 1) = True
.Cell(flexcpBackColor, 0, .Cols - 1, .Rows - 1, .Cols - 1) = &HE0E0E0
End With
GRDTABLE.Close
Set GRDTABLE = Nothing
End Sub
Private Sub FixCost(cFieldas, cField, pFilter As String)
Dim loctable As ADODB.Recordset
' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,cost from file1_10 " & turn(pFilter) & pFilter & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(cost)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grdModel.Cols - 2
        grdModel.TextMatrix(1, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If

' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,price from file1_10 " & turn(pFilter) & pFilter & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(price)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grdModel.Cols - 2
        grdModel.TextMatrix(2, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If
Set loctable = Nothing

' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,price2 from file1_10 " & turn(pFilter) & pFilter & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(price2)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grdModel.Cols - 2
        grdModel.TextMatrix(3, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If
Set loctable = Nothing
End Sub
Private Function retFields(pFilter As String)
Dim aRet(1) As String
Dim FieldTable As New ADODB.Recordset
'  ⁄—Ì› «·«⁄„œ…
Dim cString As String, cwhere As String
FieldTable.Open "Select SCAL from file1_10 " & turn(pFilter) & pFilter & " GROUP BY SCAL,C_SCAL order by c_scal", con, adOpenStatic, adLockReadOnly
Do Until FieldTable.EOF
    If Not IsNull(FieldTable!SCAL) Then
        cFieldas = cFieldas & turn(cField, ",") & "[" & FieldTable!SCAL & "]" & " as " & "[" & FieldTable!SCAL & "]"
        cField = cField & turn(cField, ",") & "[" & FieldTable!SCAL & "]"
    End If
    FieldTable.MoveNext
Loop

aRet(0) = cField
aRet(1) = cFieldas
retFields = aRet
' ⁄œ„ ÊÃÊœ «⁄„œ…
FieldTable.Close
Set FieldTable = Nothing
End Function

Private Sub GRID1_Click()
For nRow = 4 To grdModelCodes.Rows - 1
    For nCol = 2 To grdModelCodes.Cols - 1
        grdModel.TextMatrix(nRow, nCol) = ""
    Next
Next

If Val(grid1.TextMatrix(grid1.Row, 3)) = 0 And Val(grid1.TextMatrix(grid1.Row, 4)) = 0 Then Exit Sub
Dim loctable As New ADODB.Recordset, cString As String
If Val(grid1.TextMatrix(grid1.Row, 3)) <> 0 Then
    cString = "SELECT SUM(QUANT) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL FROM FILE7_20 INNER JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM"
    cString = cString & turn(cString) & "FILE7_20.DOC_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 1))
    cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(pModelNo)
Else
    cString = "SELECT SUM(QUANT) as sumofQuant,FILE1_10.COLOR,FILE1_10.SCAL FROM FILE7_10 INNER JOIN FILE1_10 ON FILE7_10.ITEM = FILE1_10.ITEM"
    cString = cString & turn(cString) & "FILE7_10.DOC_NO = " & MyParn(grid1.TextMatrix(grid1.Row, 1))
    cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(pModelNo)
End If
cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL "
loctable.Open cString, con, adOpenStatic, adLockReadOnly
Do Until loctable.EOF
    For nRow = 4 To grdModel.Rows - 1
        If grdModel.TextMatrix(nRow, 1) = loctable!Color Then
            For nCol = 2 To grdModel.Cols - 1
                If grdModel.TextMatrix(0, nCol) = loctable!SCAL Then
                    grdModel.TextMatrix(nRow, nCol) = loctable!sumofQuant
                    Exit For
                End If
            Next
        End If
    Next
    loctable.MoveNext
Loop
CalcTotals
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CalcTotals()
With grdModel
If .Rows > 4 Then
    For nCol = 2 To .Cols - 1
        For nRow = 4 To .Rows - 2
            nTotal = nTotal + Val(.TextMatrix(nRow, nCol))
        Next
        .TextMatrix(.Rows - 1, nCol) = Myvalue(nTotal)
        nTotal = 0
    Next
    For nRow = 4 To .Rows - 1
        For nCol = 2 To .Cols - 2
            nTotal = nTotal + Val(.TextMatrix(nRow, nCol))
        Next
        .TextMatrix(nRow, .Cols - 1) = Myvalue(nTotal)
        nTotal = 0
    Next
    .ColAlignment(.Cols - 1) = flexAlignRightCenter
End If
End With
End Sub
Private Sub FillItem_INCASE(cFieldas, cField, pFilter As String)
Dim GRDTABLE As New ADODB.Recordset, cString As String
' „·∆ «·ÃœÊ·
cString = "Select c_color as [—ﬁ„ «··Ê‰] ,color as [«··Ê‰] " & turn(cFieldas, ",") & cFieldas & _
          " From " & _
          " (Select c_color,Color,scal,item from file1_10 " & turn(pFilter) & pFilter & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(item)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  " & _
          " order by pvt.C_color"

GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grdModel.Cols = GRDTABLE.Fields.Count: grdModelCodes.Cols = GRDTABLE.Fields.Count

For nCol = 2 To GRDTABLE.Fields.Count - 1
    grdModel.TextMatrix(0, nCol) = GRDTABLE.Fields(nCol).Name
Next

Do Until GRDTABLE.EOF
    grdModelCodes.AddItem ""
    grdModel.AddItem ""
    For nCol = 0 To GRDTABLE.Fields.Count - 1
        If nCol <= 1 Then
            grdModel.TextMatrix(grdModelCodes.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
        Else
            grdModelCodes.TextMatrix(grdModelCodes.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
            grdModel.TextMatrix(grdModel.Rows - 1, nCol) = ""
        End If
    Next
    GRDTABLE.MoveNext
Loop
With grdModel
.AddItem ""
.TextMatrix(.Rows - 1, 1) = "≈Ã„«·Ì „ﬁ«”"
.Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HE0E0E0
.ColWidth(.Cols - 1) = 1000

.Cols = .Cols + 1
.TextMatrix(0, .Cols - 1) = "≈Ã„«·Ì ·Ê‰"
.Cell(flexcpBackColor, 0, .Cols - 1, .Rows - 1, .Cols - 1) = &HE0E0E0
End With
GRDTABLE.Close
Set GRDTABLE = Nothing
End Sub


