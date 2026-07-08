VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form barcodefrm_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»«— ﬂÊœ „ƒﬁ "
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Simplified Arabic"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "≈Œ Ì«— „ÊœÌ· ··ÿ»«⁄…"
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   180
      Width           =   4665
      Begin VB.TextBox xmodelfact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   270
         Width           =   1950
      End
      Begin VB.Label Label4 
         Caption         =   "—ﬁ„ „ÊœÌ·"
         Height          =   390
         Left            =   2385
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   270
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   6225
      Begin VB.CommandButton cmdPrint 
         Height          =   540
         Left            =   1305
         Picture         =   "BARCODE_2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   540
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE_2.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelPrinted 
         CausesValidation=   0   'False
         Height          =   540
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE_2.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   540
         Left            =   3735
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE_2.frx":7130
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4950
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE_2.frx":96A9
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8295
      Left            =   90
      TabIndex        =   11
      Top             =   945
      Width           =   15135
      _cx             =   26696
      _cy             =   14631
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   9270
      Width           =   3855
      Begin VB.TextBox xCol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox xRow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "«·⁄„Êœ :"
         Height          =   300
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "«·’›:"
         Height          =   390
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
Attribute VB_Name = "barcodefrm_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim tCard As New ADODB.Recordset
Dim oSearchItemDesca As New Search
Dim NCOLS As Double
Dim NROWS As Double
Private Sub CmdDelPrinted_Click()
    myDelete "Õ–› «·ﬂ·"
End Sub
Private Sub CmdExit_Click()
    Unload Me
    Set barcodefrm = Nothing
End Sub
Private Sub cmdPrint_Click()
'If Val(xRow.Text) > 3 Then
'    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
'    Exit Sub
'End If
'
'If Val(xCol.Text) > 6 Then
'    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
'    Exit Sub
'End If

'myreplace
'If Not myvalidRowCol Then Exit Sub
DoprintStr2
Set myForm = Me
CardPrintNew.Show
End Sub
Private Sub cmdSave_Click()
If myreplaceGrd Then Inform " „ «·Õ›Ÿ"
End Sub
Private Sub cmduno_Click()
myloadgrd
End Sub
Private Sub CmdDel_Click()
myDelete "Õ–› „«  „  ÿ»«⁄ Â „‰ «·„” ‰œ", "ISPRINT = 1"
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
'Grid1.LoadGrid App.Path & "\addPrint.grd", flexFileData
frmReturn.Show

End Sub
Private Sub CmdDelAll_Click()
If MsgBox("Õ–› ﬂ· «·”Ã·« ", vbYesNo + vbDefaultButton1) = vbYes Then
    grid1.Rows = 1
    grid1.AddItem ""
    con.Execute "DELETE FROM ADDPRINT WHERE PRINT = 0"
    myloadgrd
End If
End Sub
Private Sub Form_Load()
openCon con
Set grid1.DataSource = data1
data1.ConnectionString = strCon
myloadgrd
'grid1.Row = 1
'grid1.Col = 0
End Sub
Sub myloadgrd()
With grid1
 '                   0               1               2       3                   4               5                   6               7             8               9           10              11              12
cString = "Select ADDPRINT.ITEM,FILE1_10.MOSM,fact.desca ,FILE1_10.MODElfaCT,file1_10.desca , FILE1_10.SUPP , FILE1_10.COLOR ,FILE1_10.SCAL , ADDPRINT.QUANT,FILE1_10.PRICE,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,ADDPRINT.ID " & _
          " FROM (FILE1_10 inner join fact on fact.code = file1_10.fact ) INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM  "
data1.RecordSource = cString
data1.Refresh
grid1.AddItem ""
MakeSerial

cmdSave.Enabled = grid1.Rows > 2
cmdDelPrinted.Enabled = grid1.Rows > 2
cmdPrint.Enabled = grid1.Rows > 2

'MakeSerial
'CalcTotals

Fixgrd
End With
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    myEdit grid1.Row, grid1.Col
End Sub
Private Sub GRID1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'ItemsLookupAll Me, oSearchItemDesca, , True
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 2 Or grid1.Col = 3 Or grid1.Col = 7 Or grid1.Col = 10 Or grid1.Col = 11 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 Then
    grid1.RemoveItem grid1.Row
'    CalcTotals
End If
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
        'CalcTotals
    End If
End If
End With
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean) As Boolean
With grid1
If (.TextMatrix(Row, 3)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub delcheck()
For i = 1 To grid1.Rows - 2
   If Val(grid1.TextMatrix(i, 8)) <> 0 Then
        grid1.RowHidden(i) = True
   End If
Next
myreplaceGrd
End Sub
Sub myProc()
    Dim bNew As Boolean:     bNew = grid1.Row = grid1.Rows - 1
    Dim nCol As Long: nCol = grid1.Col
    
    If grid1.Col = 3 Then
        grid1.TextMatrix(grid1.Row, 2) = oSearchItemDesca.grid1.TextMatrix(oSearchItemDesca.grid1.Row, 0)
    End If
    
    Grid1_AfterEdit grid1.Row, grid1.Col
    
    
    If Not bNew Then
        If nCol = 3 Then Unload oSearchitem Else oSearchItemDesca.Hide
    Else
        grid1.Col = nCol
    End If
End Sub
Private Function DoprintStr2() As Boolean
nLastMargin = 0.1
SettingArray(cUpMargin) = MyMeasure(-0.2)
SettingArray(cLeftMargin) = MyMeasure(0.1)
SettingArray(cCardWidth) = MyMeasure(5)
SettingArray(cCardHeight) = MyMeasure(2.5)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 1
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select ADDPRINT.ITEM,FILE1_10.MOSM,fact.desca as descafact ,FILE1_10.MODElfaCT,file1_10.desca , FILE1_10.SUPP,FILE1_10.COLOR,FILE1_10.SCAL,ADDPRINT.QUANT,FILE1_10.PRICE,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,ADDPRINT.ID " & _
          " FROM (FILE1_10 inner join fact on fact.code = file1_10.fact ) INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM  "
cString = cString & "where addPrint.isPrint = 1 order by file1_10.mosm , file1_10.fact , file1_10.modelfact , file1_10.color , file1_10.c_scal "

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * NCOLS) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = NCOLS, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    BLASTROW = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next
'«‰ Â«¡

nCard = 0
nRow = 0
nCard = 0
nCol = 0
NCOLS = SettingArray(cCols)
NROWS = SettingArray(cRows)
nFixLeft = MyMeasure(0.1)
nFixlast = MyMeasure(0.18)
nFixlastbar = MyMeasure(0.1)

With sourcetable
Do Until .EOF
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        
        nHeight = 0
        For nCount = 1 To 1
            BLASTROW = (nRow = NROWS And nCount = 2)
            
            temptable.AddNew
            temptable!Top = MyMeasure(0.1) + nHeight
            temptable!Left = MyMeasure(0.2) - nFixLeft
            temptable!Width = 0
            temptable!Height = 0
            temptable!FontName = "Tahoma"
            temptable!FontBold = True
            temptable!FontSize = 14
            temptable!Text = sourcetable!descafact
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(1.8) - nFixLeft
            temptable!Top = MyMeasure(0.2) + nHeight
            temptable!Width = MyMeasure(3)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 10
            temptable!TextAlign = taRightTop
            temptable!Text = sourcetable!doc_no
            temptable!CardNo = nCard
            temptable.Update
            
'            temptable.AddNew
'            temptable!Left = MyMeasure(0.1) - nFixLeft
'            temptable!Top = MyMeasure(2.5) + nHeight
'            temptable!Width = MyMeasure(2)
'            temptable!Height = 0
'            temptable!FontName = "Arial"
'            temptable!TextAngle = 900
'            temptable!FontBold = True
'            temptable!TextAlign = taCenterTop
'            temptable!FontSize = 12
'            temptable!Text = sourcetable!descafact
'            temptable!ForeColor = vbBlack
'            temptable!CardNo = nCard
'            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.7) - nFixLeft
            temptable!Top = MyMeasure(0.67) + nHeight
            temptable!Width = MyMeasure(3)
            temptable!Height = MyMeasure(0.6)
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 8
            temptable!isBarcode = True
            temptable!Text = sourcetable!doc_no
            temptable!CardNo = nCard
            temptable.Update
           
            temptable.AddNew
            temptable!Left = MyMeasure(3.6) - nFixLeft
            temptable!Top = MyMeasure(0.68) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = MyMeasure(0.5)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            temptable!TextAlign = taRightTop
            temptable!Text = DelZero(sourcetable!modelfact)
            temptable!CardNo = nCard
            temptable.Update
           
'            temptable.AddNew
'            temptable!Left = MyMeasure(3.8) - nFixLeft
'            temptable!Top = MyMeasure(0.95) + nHeight
'            temptable!Width = MyMeasure(1)
'            temptable!Height = MyMeasure(0.5)
'            temptable!FontName = "arial"
'            temptable!FontBold = False
'            temptable!FontSize = 8
'            temptable!Text = sourcetable!SUPP
'            temptable!CardNo = nCard
'            temptable.Update
           
           
            temptable.AddNew
            temptable!Left = MyMeasure(0.7) - nFixLeft
            temptable!Top = MyMeasure(1.4) + nHeight
            temptable!Width = MyMeasure(4)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!ForeColor = vbBlack
            temptable!FontSize = 10
            temptable!Text = sourcetable!DESCA
            temptable!TextAlign = taCenterTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
'            temptable.AddNew
'            temptable!Left = MyMeasure(0.7) - nFixLeft
'            temptable!Top = MyMeasure(1.55) + nHeight
'            temptable!Width = MyMeasure(2)
'            temptable!Height = 0
'            temptable!FontName = "arial"
'            temptable!FontBold = True
'            temptable!ForeColor = vbBlack
'            temptable!FontSize = 8
'            temptable!Text = sourcetable!Color
'            temptable!TextAlign = taLeftTop
'            temptable!ForeColor = vbBlack
'            temptable!CardNo = nCard
'            temptable.Update
    
'            temptable.AddNew
'            temptable!Left = MyMeasure(2.9) - nFixLeft
'            temptable!Top = MyMeasure(1.55) + nHeight
'            temptable!Width = MyMeasure(1.7)
'            temptable!Height = 0
'            temptable!FontName = "arial"
'            temptable!FontBold = False
'            temptable!ForeColor = vbBlack
'            temptable!FontSize = 8
'            temptable!Text = DelZero(sourcetable!modelfact)
'            temptable!TextAlign = taRightTop
'            temptable!ForeColor = vbBlack
'            temptable!CardNo = nCard
'            temptable.Update
    
'            temptable.AddNew
'            temptable!Left = MyMeasure(0.8) - nFixLeft
'            temptable!Top = MyMeasure(1.8) + nHeight
'            temptable!Width = MyMeasure(3)
'            temptable!Height = 0
'            temptable!FontName = "arial"
'            temptable!FontBold = True
'            temptable!FontUnderline = True
'            temptable!FontSize = 14
'            temptable!TextAlign = taLeftTop
'            temptable!Text = "L.E." & sourcetable!price
'            temptable!ForeColor = vbBlack
'            temptable!CardNo = nCard
'            temptable.Update
'
'            temptable.AddNew
'            temptable!Left = MyMeasure(3) - nFixLeft
'            temptable!Top = MyMeasure(1.8) + nHeight
'            temptable!Width = MyMeasure(1.7)
'            temptable!Height = 0
'            temptable!FontName = "arial"
'            temptable!FontBold = False
'            temptable!FontUnderline = True
'            temptable!TextAlign = taRightTop
'            temptable!FontSize = 14
'            temptable!ForeColor = vbBlack
'            temptable!Text = sourcetable!scal
'            temptable!CardNo = nCard
'            temptable.Update
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
End Function
Private Function myvalidRowCol() As Boolean
'If Val(xRow.Text) > SettingArray(cRows) Then
If Val(xRow.Text) > NROWS Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Function
End If

If Val(xCol.Text) > NCOLS Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Function
End If
myvalidRowCol = True
End Function
Private Sub Fixgrd()
With grid1
   .Cols = 13
'                   0               1           2       3               4       5               6       7           8           9           10              11
    .FormatString = "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "—ﬁ„ „ÊœÌ·|" & "«·„ÊœÌ·|" & "«·„Ê—œ|" & "«··Ê‰|" & "«·„ﬁ«”|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "»«—ﬂÊœ ÃœÌœ|" & "«·ÿ»«⁄…|"
    .FixedCols = 0
    .ColWidth(0) = 1000
    .ColWidth(1) = 0
    .ColWidth(2) = 2000
    .ColWidth(3) = 1500
    .ColWidth(5) = 0
    .ColWidth(4) = 3000
    .ColWidth(6) = 0
    .ColWidth(7) = 0
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 2000
    
    .ColDataType(11) = flexDTBoolean

    .ColHidden(.Cols - 1) = True
    For i = 0 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 3) = DelZero(.TextMatrix(nRow, 3))
    Next nRow
 
 ' .ColDataType(8) = flexDTBoolean
End With
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
'For i = 1 To grid1.Rows - 1
'    grid1.TextMatrix(i, 0) = i
'Next
End Sub
Private Sub myEdit(Row As Long, Col As Long)
Dim bNew As Boolean
With grid1

'If Col = 0 Then GrdDesc Row, Col

If Not validRow(Row) Then Exit Sub

'If Row = .Rows - 1 Then
'    .AddItem ""
'    .TextMatrix(.Rows - 1, 0) = .Rows - 1
'End If

If myreplaceGrd Then
    bNew = grid1.TextMatrix(Row, .Cols - 1) = ""
End If

If Row = grid1.Rows - 2 And bNew Then
    myloadgrd
    grid1.ShowCell grid1.Rows - 1, 2
    grid1.Select .Rows - 1, 2
End If
End With
End Sub
Private Function myreplaceGrd(Optional nRow As Long = -1) As Boolean
Dim aInsert(4, 1)
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert(0, 0) = "model"
        aInsert(0, 1) = addstring(grid1.TextMatrix(i, 3))
        
        aInsert(1, 0) = "doc_no"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 10))
        
        aInsert(2, 0) = "item"
        aInsert(2, 1) = addstring(grid1.TextMatrix(i, 0))
                
        aInsert(3, 0) = "quant"
        aInsert(3, 1) = Val(.TextMatrix(i, 8))
       
        aInsert(4, 0) = "isprint"
        aInsert(4, 1) = IIf(.TextMatrix(i, 11) = "True" Or Val(.TextMatrix(i, 11)) = -1, 1, 0)
        
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "ADDPRINT")
        Else
            con.Execute CreateUpdate(aInsert, "ADDPRINT", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
myreplaceGrd = True
End Function
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem (grid1.Row)
        CalcTotals
        MakeSerial grid1.Row
    End If
End If
End Sub
Private Function RemoveItem(Row As Long) As Boolean
con.BeginTrans
On Error GoTo myerror
If grid1.TextMatrix(Row, grid1.Cols - 1) <> "" Then
    con.Execute "Delete  From ADDPRINT where id = " & grid1.TextMatrix(Row, grid1.Cols - 1)
End If
grid1.RemoveItem Row
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Sub myDelete(cMsg As String, Optional cFilter As String)
If MsgBox(cMsg, vbYesNo + vbDefaultButton2) = vbYes Then
    con.BeginTrans
    On Error GoTo myerror
    con.Execute "DELETE FROM ADDPRINT"
    con.CommitTrans
    myloadgrd
End If
Exit Sub
myerror:
    con.RollbackTrans
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub CalcTotals()

End Sub
Private Sub xmodelfact_KeyPress(KeyAscii As Integer)
Dim nRow As Long
If KeyAscii = 13 Then
    With grid1
        If xmodelfact.Text = "" Then
            For nRow = 1 To .Rows - 1
                .TextMatrix(nRow, 11) = False
                myreplaceGrd nRow
            Next nRow
        Else
            For nRow = 1 To .Rows - 1
                If .TextMatrix(nRow, 3) = xmodelfact.Text Then
                    .TextMatrix(nRow, 11) = True
                End If
                myreplaceGrd nRow
            Next nRow
        End If
    End With
End If
End Sub
