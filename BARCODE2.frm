VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form barcodefrm2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»«— ﬂÊœ"
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "≈Œ›«¡ «·„’‰⁄"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   4410
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   450
      Width           =   1410
   End
   Begin VB.Frame Frame4 
      Height          =   870
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   3930
      Begin VB.CheckBox XALL 
         Alignment       =   1  'Right Justify
         Caption         =   "≈Œ Ì«— «·ﬂ·"
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox xmodelfact 
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
         Height          =   360
         Left            =   1575
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   " ÕœÌœ „ÊœÌ·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   405
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   5895
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   4200
      Begin VB.CommandButton cmdPrint1 
         Caption         =   "ÿ»«⁄… 1"
         CausesValidation=   0   'False
         Height          =   585
         Left            =   2745
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   1410
      End
      Begin VB.CommandButton cmdPrint2 
         Caption         =   "ÿ»«⁄… 2"
         CausesValidation=   0   'False
         Height          =   585
         Left            =   1305
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   225
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.CommandButton cmdPrint96 
         BackColor       =   &H00FF8080&
         Caption         =   "ÿ»«⁄… 96"
         Height          =   585
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   225
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   10125
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   540
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelAll 
         CausesValidation=   0   'False
         Height          =   540
         Left            =   1305
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE2.frx":246C
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
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE2.frx":4D06
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
         Left            =   3735
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE2.frx":727F
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
      Left            =   180
      TabIndex        =   10
      Top             =   900
      Width           =   15045
      _cx             =   26538
      _cy             =   14631
      _ConvInfo       =   1
      Appearance      =   0
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
      Cols            =   12
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
Attribute VB_Name = "barcodefrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearchItem As New Search3
Dim con As New ADODB.Connection
Dim temptable As New ADODB.Recordset
Dim NCOLS As Double
Dim NROWS As Double

Private Sub cmd_bar96_Click()

End Sub
Private Sub CmdDelPrinted_Click()
If MsgBox("Õ–› „«  „  ÿ»«⁄ Â", vbOKCancel + vbDefaultButton2 + vbCritical, "Õ–› „«  „  ÿ»«⁄ Â") <> vbOK Then Exit Sub
Dim cError As String, cFilter As String
cFilter = "ADDPRINT.ISPRINT = 1"
cError = myDelete(cFilter)
If cError = "ok" Then Inform " „ Õ–› ﬂ· «·”Ã·«  «·„ÿ»Ê⁄… »‰Ã«Õ" Else MsgBox cError
End Sub

Private Sub cmd_str_Click()
If Val(xRow.Text) > 4 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If
If Val(xCol.Text) > 6 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
myreplaceGrd
myloadgrd
If Not DoprintOkazStr Then Exit Sub
Set myForm = Me
CardPrintNew.Show 1
CmdDelPrinted_Click

End Sub

Private Sub CmdDelAll_Click()
If MsgBox("Õ–› «·ﬂ· ", vbOKCancel + vbDefaultButton2 + vbCritical, "Õ–› ﬂ· «·”Ã·« ") <> vbOK Then Exit Sub
Dim cError As String
cError = myDelete
If cError = "ok" Then Inform " „ Õ–› ﬂ· «·”Ã·«  »‰Ã«Õ" Else MsgBox cError
End Sub

Private Sub cmdExit_Click()
Unload Me
Set barcodefrm = Nothing
End Sub
Private Sub cmdPrint1_Click()
If Val(xRow.Text) > 16 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If
If Val(xCol.Text) > 6 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
myreplaceGrd
myloadgrd
If Not doprint1 Then Exit Sub
Set myForm = Me
aPrinterBar = AddFlag(Empty, "name", RetPrinterByType("2"))
CardPrintNew.Show 1
CmdDelPrinted_Click
End Sub
Private Sub cmdPrint96_Click()
If Val(xRow.Text) > 16 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If
If Val(xCol.Text) > 6 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
myreplaceGrd
myloadgrd
If Not Doprint96 Then Exit Sub
Set myForm = Me
CardPrintNew.Show 1
CmdDelPrinted_Click
End Sub

Private Sub cmdPrintSale_Click()

End Sub

Private Sub cmdPrin96_Click()

End Sub

Private Sub cmdSave_Click()
    If myreplaceGrd Then Inform " „ «·Õ›Ÿ"
    myloadgrd
End Sub
Private Sub cmduno_Click()
myloadgrd
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
'Grid1.LoadGrid App.Path & "\addPrint.grd", flexFileData
frmReturn.Show
End Sub

Private Sub CmdPrint2_Click()
If Val(xRow.Text) > 8 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.Text) > 12 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
myreplaceGrd
myloadgrd

If Not doprint2 Then Exit Sub
Set myForm = Me
aPrinterBar = AddFlag(Empty, "name", RetPrinterByType("2"))
CardPrintNew.Show 1
CmdDelPrinted_Click
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
'                   0                      1           2              3               4               5              6          7              8                 9               10              11
cString = "Select ADDPRINT.ITEM,FILE1_10.MODElFACT0,FILE1_10.MODEL,file1_10.desca,FILE1_10.COLOR,FILE1_10.SCAL,ADDPRINT.QUANT,FILE1_10.PRICE,FILE1_10.PRICE2,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,ADDPRINT.ID " & _
          " FROM FILE1_10 INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM "
data1.RecordSource = cString
data1.Refresh
myaddItem
Fixgrd
End With
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
Dim bNew As Boolean
If Col = 0 Then GrdDesc Row
If Not validRow(Row) Then Exit Sub
With grid1
If Row = .Rows - 1 Then
    myaddItem
End If
CalcTotals

If myreplaceGrd(Row) Then
   bNew = grid1.TextMatrix(Row, .Cols - 1) = ""
End If
myloadgrd
If bNew Then
    'grid1.Row = grid1.Rows - 1
    grid1.ShowCell grid1.Rows - 1, 1
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_DBLClick()
With grid1
    If .Col = 1 Then
        xmodelfact.Text = .TextMatrix(.Row, 1)
        xmodelfact_KeyPress (13)
    End If
End With
End Sub

Private Sub grid1_EnterCell()
If grid1.Col = 0 Or grid1.Col = 6 Or grid1.Col = 10 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub

Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
End Sub

Private Sub Grid1_Validate(Cancel As Boolean)
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
End If
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then .RemoveItem OldRow
End If
End With
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean) As Boolean
With grid1
If (.TextMatrix(Row, 0)) = "" Then Exit Function
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
Dim bNew As Boolean
bNew = grid1.Row = grid1.Rows - 1
grid1.TextMatrix(grid1.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
Grid1_AfterEdit grid1.Row, grid1.Col
If Not bNew Then
    Unload oSearchItem
    CellPos 13, grid1.Row, grid1.Col
Else
    grid1.Row = grid1.Rows - 1
    grid1.Col = 0
End If
End Sub
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
'                   0                      1           2              3               4               5              6          7              8                 9               10              11
'cString = "Select ADDPRINT.ITEM,FILE1_10.MODElFACT,FILE1_10.MODEL,file1_10.desca,FILE1_10.COLOR,FILE1_10.SCAL,ADDPRINT.QUANT,FILE1_10.PRICE,FILE1_10.PRICE2,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,ADDPRINT.ID "
    .FormatString = "«·ﬂÊœ|" & "„ÊœÌ· «·„’‰⁄|" & "—ﬁ„ «·„ÊœÌ·|" & "≈”„ «·„ÊœÌ·|" & "«··Ê‰|" & "«·„ﬁ«”|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "”⁄— «·«Êﬂ«“ÌÊ‰|" & "—ﬁ„ «·„” ‰œ|" & "«·ÿ»«⁄…|"
    .FixedCols = 0
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1500
    .ColWidth(3) = 4500
    .ColWidth(5) = 1000
    .ColWidth(4) = 1000
    .ColWidth(6) = 800
    .ColWidth(7) = 800
    .ColWidth(8) = 900
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColHidden(9) = True
    .ColDataType(10) = flexDTBoolean
    .ColHidden(11) = True
    .ColHidden(2) = True
    .ColHidden(.Cols - 1) = True
    .RowHeight(0) = 700
    For i = 0 To grid1.Cols - 1
        If i <> 3 Then .ColAlignment(i) = flexAlignRightCenter
    Next
    .Cell(flexcpFontSize, 1, 3, .Rows - 1, 3) = 10
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 3) = DelZero(.TextMatrix(nRow, 3))
    Next nRow
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
If Not validRow(Row) Then Exit Sub

If Row = .Rows - 1 Then
    myaddItem
End If

If myreplaceGrd Then
    bNew = grid1.TextMatrix(Row, .Cols - 1) = ""
End If

If bNew Then
    myloadgrd
    grid1.ShowCell grid1.Rows - 1, 2
    grid1.Select .Rows - 1, 2
End If
End With
End Sub
Private Function myreplaceGrd(Optional nRow As Long = -1) As Boolean
Dim aInsert(3, 1)
With grid1
For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
     aInsert(0, 0) = "doc_no"
     aInsert(0, 1) = addstring(grid1.TextMatrix(i, 9))
     
     aInsert(1, 0) = "item"
     aInsert(1, 1) = addstring(grid1.TextMatrix(i, 0))
             
     aInsert(2, 0) = "quant"
     aInsert(2, 1) = Val(.TextMatrix(i, 6))
    
     aInsert(3, 0) = "isprint"
     aInsert(3, 1) = IIf(Abs(Val(.TextMatrix(i, 10))) <> 0, 1, 0)
     
     If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
         con.Execute CreateInsert(aInsert, "ADDPRINT")
     Else
         con.Execute CreateUpdate(aInsert, "ADDPRINT", " where ID = " & grid1.TextMatrix(i, 11), Array(-1))
     End If
Next
End With
myreplaceGrd = True
End Function
'Private Sub GrdDesc(ByVal Row As Long, Col As Long)
'Dim loctable  As ADODB.Recordset
'If IsNumeric(grid1.TextMatrix(Row, 2)) Then
'    Set loctable = RetItemRow(grid1.TextMatrix(Row, 2), con)
'    If Not (loctable.EOF And loctable.BOF) Then
'        grid1.TextMatrix(Row, 1) = loctable!Item & ""
'        grid1.TextMatrix(Row, 3) = loctable!Model & ""
'
'        Dim ModelTable As New ADODB.Recordset
'        Set ModelTable = ModelByModel(loctable!Model & "", con, "ModelByModelAll")
'        If Not (ModelTable.EOF And ModelTable.BOF) Then
'            grid1.TextMatrix(Row, 4) = ModelTable!X1 & ""
'        End If
'
'        Set ModelTable = Nothing
'        grid1.TextMatrix(Row, 5) = loctable!SCAL
'        grid1.TextMatrix(Row, 6) = loctable!Color & ""
'        grid1.TextMatrix(Row, 7) = "1"
'        grid1.TextMatrix(Row, 8) = Format(loctable!PRICE, "#0.00")
'        grid1.TextMatrix(Row, 9) = Format(loctable!Price2, "#0.00")
'        grid1.TextMatrix(Row, 11) = -1
'    End If
'End If
'Set loctable = Nothing
'End Sub
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem (grid1.Row)
        CalcTotals
        MakeSerial grid1.Row
    End If
ElseIf KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
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
Private Function myDelete(Optional cFilter As String) As String
Dim cString As String
cString = "DELETE FROM ADDPRINT"
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
con.BeginTrans
On Error GoTo myerror
con.Execute cString
con.CommitTrans
myDelete = "ok"
myloadgrd
Exit Function
myerror:
    myDelete = Err.Description
    con.RollbackTrans
    Err.Clear
End Function
Private Sub CalcTotals()

End Sub

Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
    If Col = 0 Then
        If Not IsNumeric(.EditText) Then
            Cancel = True
            Exit Sub
        End If
        If Not validItem(.EditText, con) Then
            MsgBox ("ﬂÊœ «·„ÊœÌ· €Ì— ”·Ì„")
            Cancel = True
        End If
    End If
End With
End Sub

Private Sub XALL_Click()
With grid1
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 10) = XALL.Value
    Next nRow
    cmdSave_Click
End With
End Sub

Private Sub xmodelfact_KeyPress(KeyAscii As Integer)
Dim nRow As Long
If KeyAscii = 13 Then
    With grid1
        If xmodelfact.Text = "" Then
            For nRow = 1 To .Rows - 1
                .TextMatrix(nRow, 10) = 1
                myreplaceGrd nRow
            Next nRow
        Else
            For nRow = 1 To .Rows - 1
                If .TextMatrix(nRow, 1) = xmodelfact.Text Then
                    .TextMatrix(nRow, 10) = 1
                End If
                myreplaceGrd nRow
            Next nRow
        End If
    End With
End If
End Sub
Private Sub GrdDesc(Row)
With grid1
grid1.TextMatrix(Row, 1) = ""
grid1.TextMatrix(Row, 2) = ""
grid1.TextMatrix(Row, 5) = ""
grid1.TextMatrix(Row, 3) = ""
grid1.TextMatrix(Row, 4) = ""
grid1.TextMatrix(Row, 7) = ""
grid1.TextMatrix(Row, 6) = ""
grid1.TextMatrix(Row, 9) = ""
grid1.TextMatrix(Row, 8) = ""
If Not IsNumeric(grid1.TextMatrix(Row, 0)) Then Exit Sub
Dim ItemTable As New ADODB.Recordset
Set ItemTable = ItemFind(grid1.TextMatrix(Row, 0), con)
If Not ItemTable.EOF Then
    grid1.TextMatrix(Row, 1) = ItemTable!MOSM
    grid1.TextMatrix(Row, 2) = ItemTable!Fact
    grid1.TextMatrix(Row, 5) = ItemTable!Supp & ""
    grid1.TextMatrix(Row, 3) = ItemTable!MODELFACT0
    grid1.TextMatrix(Row, 4) = ItemTable!DESCA
    grid1.TextMatrix(Row, 7) = ItemTable!scal
    grid1.TextMatrix(Row, 6) = ItemTable!Color
    grid1.TextMatrix(Row, 9) = Format(ItemTable!price, "#0.00")
    grid1.TextMatrix(Row, 8) = 1
End If
End With
End Sub
Private Sub cmdPrintBar_Click()
If Not myreplaceGrd Then
    MsgBox "Œÿ√ «À‰«¡«·Õ›Ÿ"
    Exit Sub
End If
If Val(xRow.Text) > 3 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.Text) > 6 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
doprint
Set myForm = Me
CardPrintNew.Show 1
End Sub
Private Function doprint() As Boolean
nLastMargin = 0.1
SettingArray(cUpMargin) = MyMeasure(0.3)
SettingArray(cLeftMargin) = MyMeasure(0)
SettingArray(cCardWidth) = MyMeasure(3.4)
SettingArray(cCardHeight) = MyMeasure(6.79)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 3
SettingArray(cCols) = 6
SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select FILE1_10.MODELFACT,ADDPRINT.ITEM,FILE1_10.MODEL,FILE1_10.[SEC]" & _
          ",dbo.f_min_scal(FILE1_10.MODEL) as minScal,dbo.f_max_scal(FILE1_10.MODEL) as maxScal" & _
          ",FILE1_10.DESCA,FILE1_10.PRICE,FILE1_10.PRICE2,FILE1_10.COLOR,ADDPRINT.QUANT,FILE1_10.SCAL" & _
          " FROM FILE1_10 INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM  "
cString = cString & turn(cString) & " ADDPRINT.ISPRINT = 1"
cString = cString & "order by file1_10.model, file1_10.c_scal,file1_10.color "

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

Dim nCard As Long, nRow As Long, nCol As Long, NCOLS As Long, NROWS As Long
nCard = 0
nRow = 0
nCard = 0
nCol = 0
NCOLS = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * NCOLS) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = NCOLS, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next

Do Until sourcetable.EOF
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        
        temptable.AddNew
        temptable!Left = MyMeasure(3.5)
        temptable!Top = MyMeasure(0.4)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 9
        temptable!Text = IIf(nCol = NCOLS, Null, "|")
        temptable!CardNo = nCard
        temptable.Update
        
        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(0.3)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 9
        temptable!Text = "TRES CHIC"
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(0.7)
        temptable!Width = MyMeasure(2.3)
        temptable!Height = MyMeasure(0.6)
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 9
        temptable!Text = sourcetable!Item
        temptable!isBarcode = True
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

'/*****
        temptable.AddNew
        temptable!Left = MyMeasure(3)
        temptable!Top = MyMeasure(0.85)
        temptable!Width = MyMeasure(0)
        temptable!Height = MyMeasure(0.6)
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 9
        temptable!Text = sourcetable!sec
        temptable!isBarcode = False
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(1.3)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!FontBold = True
        temptable!FontUnderline = True
        temptable!FontSize = 9
        temptable!Text = sourcetable!ModelFACT
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(1.9)
        temptable!Top = MyMeasure(1.3)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!FontBold = True
        temptable!FontSize = 8
        temptable!Text = sourcetable!Item
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(1.8)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 9
        temptable!Text = sourcetable!model
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(2.3)
        temptable!Width = MyMeasure(2.5)
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 6
        temptable!Text = sourcetable!DESCA
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(2.7)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontUnderline = True
        temptable!FontSize = 10
        temptable!ForeColor = vbBlack
        temptable!Text = sourcetable!scal
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(1.25)
        temptable!Top = MyMeasure(2.7)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontUnderline = False
        temptable!FontSize = 12
        temptable!Text = "L.E." & Format(sourcetable!price, "Fixed")
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
    
       
        temptable.AddNew
        temptable!Left = MyMeasure(2.5)
        temptable!Top = MyMeasure(0.4)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!FontSize = 5
        temptable!Text = sourcetable!MinScal & "-" & sourcetable!MaxScal
        temptable!isBarcode = False
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
       
       
       ' ----------
        temptable.AddNew
        temptable!Left = MyMeasure(3.5)
        temptable!Top = MyMeasure(3.7) - IIf(blastrow, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 9
        temptable!Text = IIf(nCol = NCOLS, Null, "|")
        temptable!CardNo = nCard
        temptable.Update
        

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(3.7) - IIf(blastrow, MyMeasure(nLastMargin), 0)
        temptable!Width = MyMeasure(2.3)
        temptable!Height = MyMeasure(0.6)
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 9
        temptable!Text = sourcetable!Item
        temptable!isBarcode = True
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(3)
        temptable!Top = MyMeasure(3.7) - IIf(blastrow, MyMeasure(nLastMargin), 0)
        temptable!Width = MyMeasure(0)
        temptable!Height = MyMeasure(0.6)
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 9
        temptable!Text = sourcetable!sec
        temptable!isBarcode = False
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(4.3) - IIf(blastrow, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!FontBold = True
        temptable!FontUnderline = True
        temptable!FontSize = 9
        temptable!Text = sourcetable!ModelFACT
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(1.9)
        temptable!Top = MyMeasure(4.3) - IIf(blastrow, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!FontBold = True
        temptable!FontSize = 8
        temptable!Text = sourcetable!Item
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(4.7) - IIf(blastrow, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 9
        temptable!Text = sourcetable!model
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(5.07) - IIf(blastrow, MyMeasure(nLastMargin), 0)
        temptable!Width = MyMeasure(2.5)
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 6
        temptable!Text = sourcetable!DESCA
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(5.35) - IIf(blastrow, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontUnderline = True
        temptable!FontSize = 9
        temptable!ForeColor = vbBlack
        temptable!Text = sourcetable!scal
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(1.25)
        temptable!Top = MyMeasure(5.35) - IIf(blastrow, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontUnderline = False
        temptable!FontSize = 12
        temptable!Text = "L.E." & Format(sourcetable!price, "Fixed")
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
         
        temptable.AddNew
        temptable!Left = MyMeasure(1)
        temptable!Top = MyMeasure(6)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "times new Roman"
        temptable!FontBold = False
        temptable!FontUnderline = False
        temptable!FontSize = 11
        temptable!Text = IIf(blastrow, Null, "------------------")
        temptable!CardNo = nCard
        temptable.Update
                     
        temptable.AddNew
        temptable!Left = MyMeasure(3.5)
        temptable!Top = MyMeasure(6)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 10
        temptable!Text = IIf(blastrow Or nCol = NCOLS, Null, "|")
        temptable!CardNo = nCard
        temptable.Update
    
    ' «··Ê‰
        temptable.AddNew
        temptable!Left = MyMeasure(2.9)
        temptable!Top = MyMeasure(2.2) - IIf(blastrow, MyMeasure(0.2), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!TextAngle = 900
        temptable!FontBold = True
        temptable!FontUnderline = False
        temptable!FontSize = 8
        If Not IsNull(sourcetable!Color) Then
            temptable!Text = IIf(sourcetable!Color = "„ÊÕœ" Or sourcetable!Color = "√·Ê«‰" Or sourcetable!Color = "---", Null, sourcetable!Color)
        End If
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
    
    ' «··Ê‰
        temptable.AddNew
        temptable!Left = MyMeasure(2.9)
        temptable!Top = MyMeasure(5.1) - IIf(blastrow, MyMeasure(0.1), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!TextAngle = 900
        temptable!FontBold = True
        temptable!FontUnderline = False
        temptable!FontSize = 8
        If Not IsNull(sourcetable!Color) Then
            temptable!Text = IIf(sourcetable!Color = "„ÊÕœ" Or sourcetable!Color = "√·Ê«‰" Or sourcetable!Color = "---", Null, sourcetable!Color)
        End If
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
    Next
    sourcetable.MoveNext
Loop
Set sourcetable = Nothing
Set temptable = Nothing
doprint = True
End Function
Private Sub myaddItem()
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 10) = -1
MakeSerial
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 4 Then
    If Col < 6 Then
        grid1.Col = 6
    Else
        grid1.Col = Col + 1
    End If
ElseIf Row < grid1.Rows - 1 Then
    grid1.Row = grid1.Row + 1
    grid1.Col = 0
End If
End Sub
Private Function doprint1() As Boolean
nLastMargin = 0.1
SettingArray(cUpMargin) = MyMeasure(0)
SettingArray(cLeftMargin) = MyMeasure(0.1)
SettingArray(cCardWidth) = MyMeasure(3.5)
SettingArray(cCardHeight) = MyMeasure(2.5)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 1
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select ADDPRINT.ITEM,FILE1_10.MOSM,fact.desca as descafact ,FILE1_10.MODElfaCT,file1_10.desca , FILE1_10.SUPP,FILE1_10.COLOR,FILE1_10.SCAL,ADDPRINT.QUANT,FILE1_10.PRICE,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,FILE1_10.CODE,ADDPRINT.ID " & _
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
    blastrow = (nRow = NROWS)
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
            blastrow = (nRow = NROWS And nCount = 2)
            
            temptable.AddNew
            temptable!Top = MyMeasure(0.1) + nHeight
            temptable!Left = MyMeasure(0.2) - nFixLeft
            temptable!Width = MyMeasure(2.5)
            temptable!Height = 0
            temptable!TextAlign = taCenterTop
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontSize = 12
            temptable!Text = TurnValue(Secondtitle)
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(2.3) - nFixLeft
            temptable!Top = MyMeasure(0.1) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 12
            temptable!TextAlign = taRightTop
            temptable!Text = sourcetable!Item
            temptable!CardNo = nCard
            temptable.Update
            
'            If Check1.Value = 0 Then
'                temptable.AddNew
'                temptable!Left = MyMeasure(0.1) - nFixLeft
'                temptable!Top = MyMeasure(2.5) + nHeight
'                temptable!Width = MyMeasure(2)
'                temptable!Height = 0
'                temptable!FontName = "Arial"
'                temptable!TextAngle = 900
'                temptable!FontBold = True
'                temptable!TextAlign = taCenterTop
'                temptable!FontSize = 12
'                temptable!Text = sourcetable!descafact
'                temptable!ForeColor = vbBlack
'                temptable!CardNo = nCard
'                temptable.Update
'            End If
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.6) - nFixLeft
            temptable!Top = MyMeasure(0.67) + nHeight
            temptable!Width = MyMeasure(2)
            temptable!Height = MyMeasure(0.6)
            temptable!TextAlign = taRightTop
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 8
            temptable!isBarcode = True
            temptable!Text = sourcetable!Item
            temptable!CardNo = nCard
            temptable.Update
           
            temptable.AddNew
            temptable!Left = MyMeasure(2.5) - nFixLeft
            temptable!Top = MyMeasure(0.68) + nHeight
            temptable!Width = MyMeasure(0.8)
            temptable!Height = MyMeasure(0.5)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            temptable!TextAlign = taRightTop
            temptable!Text = sourcetable!MOSM
            temptable!CardNo = nCard
            temptable.Update
           
            temptable.AddNew
            temptable!Left = MyMeasure(2.5) - nFixLeft
            temptable!Top = MyMeasure(1.1) + nHeight
            temptable!Width = MyMeasure(0.8)
            temptable!Height = MyMeasure(0.5)
            temptable!TextAlign = taRightTop
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            temptable!Text = sourcetable!code
            temptable!CardNo = nCard
            temptable.Update
           
           
            temptable.AddNew
            temptable!Left = MyMeasure(0.7) - nFixLeft
            temptable!Top = MyMeasure(1.25) + nHeight
            temptable!Width = MyMeasure(2.5)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!ForeColor = vbBlack
            temptable!FontSize = 9
            temptable!Text = sourcetable!DESCA
            temptable!TextAlign = taCenterTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(0.7) - nFixLeft
            temptable!Top = MyMeasure(1.55) + nHeight
            temptable!Width = MyMeasure(1.5)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!ForeColor = vbBlack
            temptable!FontSize = 8
            temptable!Text = DelZero(sourcetable!ModelFACT)
            temptable!TextAlign = taLeftTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(2) - nFixLeft
            temptable!Top = MyMeasure(1.55) + nHeight
            temptable!Width = MyMeasure(1.3)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!ForeColor = vbBlack
            temptable!FontSize = 8
            temptable!Text = sourcetable!Color
            temptable!TextAlign = taRightTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(0.6) - nFixLeft
            temptable!Top = MyMeasure(1.8) + nHeight
            temptable!Width = MyMeasure(2)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontUnderline = True
            temptable!FontSize = 12
            temptable!TextAlign = taLeftTop
            temptable!Text = "L.E." & Format(sourcetable!price, "#0.00")
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(2.3) - nFixLeft
            temptable!Top = MyMeasure(1.95) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!TextAlign = taRightTop
            temptable!FontSize = 8
            temptable!ForeColor = vbBlack
            
            temptable!Text = sourcetable!scal
            temptable!CardNo = nCard
            temptable.Update
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
doprint1 = True
End Function
Private Function doprint2() As Boolean
nLastMargin = 0.1
SettingArray(cUpMargin) = MyMeasure(-0.3)
SettingArray(cLeftMargin) = MyMeasure(0.1)
SettingArray(cCardWidth) = MyMeasure(3.5)
SettingArray(cCardHeight) = MyMeasure(1.25)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 2
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
    blastrow = (nRow = NROWS)
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
            blastrow = (nRow = NROWS And nCount = 2)
'            If Check1.Value = 0 Then
'                temptable.AddNew
'                temptable!Top = MyMeasure(0.25) + nHeight
'                temptable!Left = MyMeasure(0.2) - nFixLeft
'                temptable!Width = MyMeasure(2)
'                temptable!Height = 0
'                temptable!FontName = "Arial"
'                temptable!FontBold = False
'                temptable!FontSize = 8
'                temptable!TextAlign = taCenterTop
'                temptable!Text = sourcetable!descafact
'                temptable!CardNo = nCard
'                temptable.Update
'            End If
            temptable.AddNew
            temptable!Left = MyMeasure(2.2) - nFixLeft
            temptable!Top = MyMeasure(0.3) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 8
            temptable!TextAlign = taRightTop
            temptable!Text = sourcetable!Item
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.2) - nFixLeft
            temptable!Top = MyMeasure(0.57) + nHeight
            temptable!Width = MyMeasure(2)
            temptable!Height = MyMeasure(0.3)
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 8
            temptable!isBarcode = True
            temptable!TextAlign = taRightTop
            temptable!Text = sourcetable!Item
            temptable!CardNo = nCard
            temptable.Update
           
            temptable.AddNew
            temptable!Left = MyMeasure(2.2) - nFixLeft
            temptable!Top = MyMeasure(0.6) + nHeight
            temptable!Width = MyMeasure(0.5)
            temptable!Height = MyMeasure(0)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 6
            temptable!TextAlign = taRightTop
            temptable!Text = sourcetable!MOSM
            temptable!CardNo = nCard
            temptable.Update
           
            temptable.AddNew
            temptable!Left = MyMeasure(2.75) - nFixLeft
            temptable!Top = MyMeasure(0.6) + nHeight
            temptable!Width = MyMeasure(0.4)
            temptable!Height = MyMeasure(0)
            temptable!TextAlign = taRightTop
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 6
            temptable!Text = sourcetable!Supp
            temptable!CardNo = nCard
            temptable.Update
           
           
            temptable.AddNew
            temptable!Left = MyMeasure(0.2) - nFixLeft
            temptable!Top = MyMeasure(0.85) + nHeight
            temptable!Width = MyMeasure(2.3)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!ForeColor = False
            temptable!FontSize = 7
            temptable!Text = sourcetable!DESCA
            temptable!TextAlign = taCenterTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(2.2) - nFixLeft
            temptable!Top = MyMeasure(0.85) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!ForeColor = vbBlack
            temptable!FontSize = 8
            temptable!Text = DelZero(sourcetable!ModelFACT)
            temptable!TextAlign = taRightTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(0.2) - nFixLeft
            temptable!Top = MyMeasure(1.1) + nHeight
            temptable!Width = MyMeasure(2)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontUnderline = True
            temptable!FontSize = 8
            temptable!TextAlign = taLeftTop
            temptable!Text = "L.E." & Format(sourcetable!price, "#0.00")
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(2.2) - nFixLeft
            temptable!Top = MyMeasure(1.12) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontUnderline = False
            temptable!TextAlign = taRightTop
            temptable!FontSize = 7
            temptable!ForeColor = vbBlack
            
            temptable!Text = sourcetable!scal
            temptable!CardNo = nCard
            temptable.Update
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
doprint2 = True
End Function
Private Function DoprintOkazStr() As Boolean
nLastMargin = 0.1
SettingArray(cUpMargin) = MyMeasure(0.1)
SettingArray(cLeftMargin) = MyMeasure(2.8)
SettingArray(cCardWidth) = MyMeasure(2.6)
SettingArray(cCardHeight) = MyMeasure(3.35)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 4
SettingArray(cCols) = 6
SettingArray(cPageWidth) = MyMeasure(16.8)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select FILE1_10.MODELFACT,ADDPRINT.ITEM,FILE1_10.MODEL,FILE1_10.[SEC]" & _
          ",dbo.f_min_scal(FILE1_10.MODEL) as minScal,dbo.f_max_scal(FILE1_10.MODEL) as maxScal" & _
          ",FILE1_10.DESCA,FILE1_10.PRICE,FILE1_10.PRICE2,FILE1_10.COLOR,ADDPRINT.QUANT,FILE1_10.SCAL" & _
          " FROM FILE1_10 INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM  "
cString = cString & turn(cString) & " ADDPRINT.ISPRINT = 1"
cString = cString & "order by file1_10.model, file1_10.c_scal,file1_10.color "

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

Dim nCard As Long, nRow As Long, nCol As Long, NCOLS As Long, NROWS As Long
nCard = 0
nRow = 0
nCard = 0
nCol = 0
NCOLS = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * NCOLS) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = NCOLS, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next

Do Until sourcetable.EOF
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = NCOLS)
        
        nHeight = 0
        nLast = 0
        nLastCol = 0
        
        For nCount = 1 To 2
        temptable.AddNew
        temptable!Left = MyMeasure(0.15)
        temptable!Top = MyMeasure(0.3) + nHeight - IIf(blastrow And nCount = 2, nLast, 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Book Antiqua"
        temptable!FontBold = False
        temptable!FontItalic = False
        temptable!FontSize = 8
        temptable!Text = sourcetable!model
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

      
        temptable.AddNew
        temptable!Left = MyMeasure(1)
        temptable!Top = MyMeasure(0.65) + nHeight - IIf(blastrow And nCount = 2, nLast, 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 8
        temptable!Text = sourcetable!Item
        temptable!CardNo = nCard
        temptable.Update

    
        temptable.AddNew
        temptable!Left = MyMeasure(0.2)
        temptable!Top = MyMeasure(0.65) + nHeight - IIf(blastrow And nCount = 2, nLast, 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!FontBold = False
        temptable!FontSize = 8
        temptable!FontUnderline = True
        temptable!Text = sourcetable!ModelFACT
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(1)
        temptable!Top = MyMeasure(1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontUnderline = False
        temptable!FontSize = 11
        temptable!Text = Format(sourcetable!PRICE2, "Fixed")
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
        
        temptable.AddNew
        temptable!Left = MyMeasure(0.2)
        temptable!Top = MyMeasure(1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!FontUnderline = True
        temptable!FontSize = 9
        temptable!ForeColor = vbBlack
        temptable!Text = sourcetable!scal
        temptable!CardNo = nCard
        temptable.Update
        nHeight = SettingArray(cCardHeight) / 2
        Next nCount
    Next i
    sourcetable.MoveNext
Loop
Set sourcetable = Nothing
Set temptable = Nothing
DoprintOkazStr = True
End Function
Private Function Doprint96() As Boolean
nLastMargin = 0
NCOLS = 6
SettingArray(cUpMargin) = MyMeasure(0)
SettingArray(cLeftMargin) = MyMeasure(0.3)
SettingArray(cCardWidth) = MyMeasure(3.5)
SettingArray(cCardHeight) = MyMeasure(1.852)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 16
SettingArray(cCols) = 6

SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable


cString = "Select ADDPRINT.ITEM,FILE1_10.MOSM,fact.desca as descafact ,FILE1_10.MODElfaCT,file1_10.desca , FILE1_10.SUPP,FILE1_10.COLOR,FILE1_10.SCAL,ADDPRINT.QUANT,FILE1_10.PRICE2,FILE1_10.PRICE,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,ADDPRINT.ID " & _
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
    blastrow = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next
'«‰ Â«¡
If sourcetable.RecordCount = 0 Then Exit Function
Do
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = NCOLS)
        
        nHeight = 0
        nLast = MyMeasure(0.2)
        nLastCol = MyMeasure(0.2)
        For nCount = 1 To 1
            temptable.AddNew
            temptable!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(1.4)
            temptable!Height = MyMeasure(0.3)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            temptable!TextAlign = taRightTop
            temptable!Text = sourcetable!Item
            temptable!isBarcode = True
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update

            temptable.AddNew
            temptable!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.45) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = 0
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = False
            temptable!FontUnderline = True
            temptable!FontSize = 7
            temptable!TextAlign = taLeftTop
            temptable!Text = DelZero(sourcetable!ModelFACT)
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(1.8) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(2)
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!TextAngle = 900
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!FontSize = 8
            temptable!Text = Secondtitle
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update

            temptable.AddNew
            temptable!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.45) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(2)
            temptable!TextAlign = taRightTop
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 8
            temptable!Text = sourcetable!Item
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(2.8) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = 0
            temptable!Height = MyMeasure(0.35)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 7
            
            temptable!Text = sourcetable!Supp
            temptable!TextAlign = taRightTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
        
            temptable.AddNew
            temptable!Left = MyMeasure(1.7) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(1)
            temptable!Height = MyMeasure(0.35)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 7
            
            temptable!Text = sourcetable!MOSM
            temptable!TextAlign = taRightTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
        
            temptable.AddNew
            temptable!Left = MyMeasure(1.3) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(1)
            temptable!Top = MyMeasure(0.45) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Height = 0
            temptable!TextAlign = taCenterTop
            temptable!FontName = "Simplified arabic"
            temptable!FontBold = False
            temptable!FontSize = 7
            If Not IsNull(sourcetable!Color) Then
                temptable!Text = IIf(sourcetable!Color = "„ÊÕœ" Or sourcetable!Color = "√·Ê«‰", Null, sourcetable!Color)
            End If
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(1.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(1.5)
            temptable!Height = 0
            temptable!FontName = "Simplified arabic"
            temptable!TextAlign = taLeftTop
            temptable!FontBold = False
            temptable!ForeColor = vbBlack
            temptable!FontSize = 7
            temptable!Text = sourcetable!DESCA
            temptable!CardNo = nCard
            temptable.Update
    
'            temptable.AddNew
'            temptable!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
'            temptable!Top = MyMeasure(1.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
'            temptable!Width = MyMeasure(2)
'            temptable!Height = 0
'            temptable!FontName = "Simplified arabic"
'            temptable!TextAlign = taRightTop
'            temptable!FontBold = False
'            temptable!ForeColor = vbBlack
'            temptable!FontSize = 7
'            temptable!Text = sourcetable!descafact
'            temptable!CardNo = nCard
'            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(1.45) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = 0
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontUnderline = False
            temptable!FontSize = 8
            temptable!Text = "L.E." & Format(sourcetable!price, "#0.00")
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(1.45) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(2)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!TextAlign = taRightTop
            temptable!FontUnderline = True
            temptable!FontSize = 8
            temptable!ForeColor = vbBlack
            temptable!Text = sourcetable!scal
            temptable!CardNo = nCard
            temptable.Update
            
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop Until sourcetable.EOF
Doprint96 = True
End Function


