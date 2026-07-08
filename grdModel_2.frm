VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form grdModel_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ·»Ì«  «·›—Ê⁄"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   13860
   Begin VB.Frame Frame1 
      Caption         =   "«·—’Ìœ"
      Height          =   3075
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   3510
      Width           =   10185
      Begin VSFlex7Ctl.VSFlexGrid grdModel2 
         Height          =   2730
         Left            =   90
         TabIndex        =   8
         Top             =   225
         Width           =   9960
         _cx             =   17568
         _cy             =   4815
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
         WordWrap        =   -1  'True
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
   Begin VB.CheckBox XISSEND 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FFFF&
      Caption         =   "IsSend"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3195
      Width           =   1275
   End
   Begin VB.CommandButton cmdExit 
      Height          =   555
      Left            =   2115
      Picture         =   "grdModel_2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4230
      Width           =   1365
   End
   Begin VB.CommandButton cmdSave 
      Height          =   600
      Left            =   2115
      MaskColor       =   &H00FFFFFF&
      Picture         =   "grdModel_2.frx":246C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Õ›Ÿ"
      Top             =   3555
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.Frame Frame4 
      Caption         =   "„ﬁ«”«  Ê«·Ê«‰"
      Height          =   3435
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   10185
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         MaxLength       =   100
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2925
         Width           =   8880
      End
      Begin VSFlex7Ctl.VSFlexGrid grdModel 
         Height          =   2550
         Left            =   90
         TabIndex        =   1
         Top             =   270
         Width           =   9915
         _cx             =   17489
         _cy             =   4498
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
         WordWrap        =   -1  'True
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
         Height          =   195
         Left            =   9270
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   2970
         Width           =   660
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2205
      Top             =   495
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
   Begin MSAdodcLib.Adodc data20 
      Height          =   330
      Left            =   2205
      Top             =   135
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
Attribute VB_Name = "grdModel_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public pString As String, pString_BAL  As String, pstore As String, pCaption As String, pFieldValue As String, pModelNo As String

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
    Dim nItem As Double
    With grdModel
        For nRow = 4 To .Rows - 2
            For nCol = 2 To .Cols - 2
                cString = " SELECT file1_10.item FROM   FILE1_10 WHERE MODELNO = " & MyParn(pModelNo) & _
                        " and scal = " & MyParn(grdModel.TextMatrix(0, nCol)) & _
                        " and color = " & MyParn(grdModel.TextMatrix(nRow, 1))
                nItem = Val(GetDesca(cString, con) & "")
                If nItem > 0 Then
                    If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                        If GetDesca("SELECT ITEM FROM Model_Dem WHERE STORE = " & MyParn(pstore) & " AND  ITEM = " & nItem, con) = "" Then
                            con.Execute " INSERT INTO Model_Dem (MODELNO , ITEM ,  STORE , QUANT , [date] ) VALUES ( " & addstring(pModelNo) & ", " & nItem & ", " & addstring(pstore) & ", " & Val(.TextMatrix(nRow, nCol)) & "  , " & addDate(Date) & " ) "
                        Else
                            con.Execute "  UPDATE Model_Dem SET QUANT = " & Val(.TextMatrix(nRow, nCol)) & " , ISNEW = 1 , DATE = " & addDate(Date) & "  WHERE ITEM = " & nItem & " AND STORE = " & MyParn(pstore)
                        End If
                    Else
                        con.Execute " DELETE FROM Model_Dem WHERE ITEM = " & nItem & " AND STORE  = " & MyParn(pstore), nRec
    '                   If nRec = 1 Then Inform "  „ Õ–› ⁄œœ  ”Ã· "
                    End If
                End If
            Next nCol
        Next nRow
        con.Execute "  UPDATE Model_Dem SET remark = " & addstring(xNotes.text) & " WHERE modelno = " & MyParn(pModelNo) & " AND STORE = " & MyParn(pstore)
    End With
    Inform "  „ Õ›Ÿ «·ÿ·»Ì… "
    Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me


End Sub
Private Sub Form_Load()
Dim cString As String
openCon con
myloadGrdModel " FILE1_10.modelno = " & MyParn(pModelNo)
myload
myload_BAL

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
.ColWidth(1) = 1300
nColWidth = (.Width - 200 - .ColWidth(0) - .ColWidth(1)) / .Cols
If nColWidth < 500 Then nColWidth = 500
If nColWidth > 1200 Then nColWidth = 1200
For nCol = 2 To .Cols - 1
    .ColWidth(nCol) = nColWidth
    .ColAlignment(nCol) = flexAlignCenterCenter
Next
.RowHidden(1) = (cBranch <> "00")
'.RowHidden(2) = True
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
grdModel.RowHidden(1) = (cBranch <> "00")
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
Dim cString As String, cWhere As String
FieldTable.Open "Select SCAL from file1_10 " & turn(pFilter) & pFilter & " GROUP BY SCAL,C_SCAL order by c_scal", con, adOpenStatic, adLockReadOnly
Do Until FieldTable.EOF
    If Not IsNull(FieldTable!scal) Then
        cFieldas = cFieldas & turn(cField, ",") & "[" & FieldTable!scal & "]" & " as " & "[" & FieldTable!scal & "]"
        cField = cField & turn(cField, ",") & "[" & FieldTable!scal & "]"
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
Private Sub FillItem(cFieldas, cField, pFilter As String)
Dim GrdTable As New ADODB.Recordset, cString As String
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

GrdTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grdModel.Cols = GrdTable.Fields.Count

For nCol = 2 To GrdTable.Fields.Count - 1
    grdModel.TextMatrix(0, nCol) = GrdTable.Fields(nCol).Name & ""
Next

Do Until GrdTable.EOF
    grdModel.AddItem ""
    For nCol = 0 To 1
        grdModel.TextMatrix(grdModel.Rows - 1, nCol) = GrdTable.Fields(nCol).Value & ""
    Next
    GrdTable.MoveNext
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
GrdTable.Close
Set GrdTable = Nothing
End Sub
Private Sub GRID1_Click()
myload
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
    .RowHidden(.Rows - 1) = True
End If
End With
End Sub
Private Sub myload()
For nRow = 4 To grdModel.Rows - 1
    For nCol = 2 To grdModel.Cols - 1
        grdModel.TextMatrix(nRow, nCol) = ""
    Next
Next
xNotes.text = GetDesca("select remark from model_dem where STORE = " & MyParn(pstore) & " AND modelno = " & MyParn(pModelNo), con)
XISSEND.Value = IIf(GetBoolean("select ISSEND from model_dem where STORE = " & MyParn(pstore) & " AND modelno = " & MyParn(pModelNo), con) = 1, 1, 0)
cmdSave.Enabled = (XISSEND.Value = 0) Or (cBranch = "00")
Dim loctable As New ADODB.Recordset, cString As String
loctable.Open pString, con, adOpenStatic, adLockReadOnly
Do Until loctable.EOF
    For nRow = 4 To grdModel.Rows - 1
        If grdModel.TextMatrix(nRow, 1) = loctable!Color Then
            For nCol = 2 To grdModel.Cols - 1
                If grdModel.TextMatrix(0, nCol) = loctable!scal Then
                    grdModel.TextMatrix(nRow, nCol) = loctable(pFieldValue) & ""
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
Private Sub grdModel_EnterCell()
    grdModel.Editable = flexEDKbdMouse
End Sub
Private Sub MYLOAD2()
    With grdModel
        grdModel2.Cols = .Cols
        grdModel2.Rows = .Rows
        grdModel.FixedCols = .FixedCols
        grdModel.FixedRows = .FixedRows
        For nRow = 0 To .Rows - 1
            For nCol = 0 To .Cols - 1
                grdModel2.TextMatrix(nRow, nCol) = .TextMatrix(nRow, nCol)
            Next nCol
        Next nRow
    End With
    With grdModel2
    For nRow = 4 To .Rows - 1
        For nCol = 2 To .Cols - 1
            .TextMatrix(nRow, nCol) = Val(GetDesca("SELECT SUM([IN]-[OUT]) FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_11.ITEM = FILE1_10.ITEM  WHERE STORE = " & MyParn(pstore) & " AND MODEL = " & MyParn(pModelNo) & " AND SCAL = " & MyParn(.TextMatrix(0, nCol)) & " AND COLOR = " & MyParn(.TextMatrix(nRow, 1)), con) & "")
        Next nCol
    Next nRow
    End With
    CalcTotals2
End Sub
Private Sub CalcTotals2()
With grdModel2
.ColHidden(0) = True
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
    .RowHidden(.Rows - 1) = True
End If

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
End Sub

Private Sub myload_BAL()

With grdModel
    grdModel2.Cols = .Cols
    grdModel2.Rows = .Rows
    grdModel2.FixedCols = .FixedCols
    grdModel2.FixedRows = .FixedRows
    For nRow = 0 To .Rows - 1
        For nCol = 0 To .Cols - 1
            grdModel2.ColWidth(nCol) = .ColWidth(nCol)
            grdModel2.TextMatrix(nRow, nCol) = .TextMatrix(nRow, nCol)
        Next nCol
    Next nRow
End With

For nRow = 4 To grdModel2.Rows - 1
    For nCol = 2 To grdModel2.Cols - 1
        grdModel2.TextMatrix(nRow, nCol) = ""
    Next
Next
Dim loctable As New ADODB.Recordset, cString As String
loctable.Open pString_BAL, con, adOpenStatic, adLockReadOnly
Do Until loctable.EOF
    For nRow = 4 To grdModel2.Rows - 1
        If grdModel2.TextMatrix(nRow, 1) = loctable!Color Then
            For nCol = 2 To grdModel2.Cols - 1
                If grdModel2.TextMatrix(0, nCol) = loctable!scal Then
                    grdModel2.TextMatrix(nRow, nCol) = loctable(pFieldValue) & ""
                    Exit For
                End If
            Next
        End If
    Next
    loctable.MoveNext
Loop


CalcTotals2
loctable.Close
Set loctable = Nothing
End Sub

Private Sub grdModel_KeyPress(KeyAscii As Integer)
    With grdModel
        If KeyAscii = 13 Then
'            If .Row = .Rows - 2 And .Col = .Cols - 2 Then
                cmdSave_Click
'            End If
        End If
    End With
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 19 Then
    cmdSave_Click
End If
End Sub

