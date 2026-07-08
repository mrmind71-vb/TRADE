VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form BarCodeCust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÿ»«⁄… »«—ﬂÊœ ⁄„·«¡ «·ﬂ«— "
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   9270
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   9270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   -90
      Width           =   2535
      Begin VB.CommandButton cmdPrintStr 
         Caption         =   "ÿ»«⁄… ” Ìﬂ—"
         CausesValidation=   0   'False
         Height          =   540
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   2160
      End
   End
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   4905
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   -90
      Width           =   3885
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   540
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BarCodeCust.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "BarCodeCust.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "BarCodeCust.frx":4D06
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8295
      Left            =   180
      TabIndex        =   9
      Top             =   855
      Width           =   8610
      _cx             =   15187
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
Attribute VB_Name = "BarCodeCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Dim oSearchItem As New Search3
Dim con As New ADODB.Connection
Dim temptable As New ADODB.Recordset
Dim NCOLS As Double
Dim NROWS As Double
Private Sub CmdDelPrinted_Click()
    If MsgBox("Õ–› „«  „  ÿ»«⁄ Â", vbOKCancel + vbDefaultButton2 + vbCritical, "Õ–› „«  „  ÿ»«⁄ Â") <> vbOK Then Exit Sub
    Dim cError As String, cFilter As String
    cFilter = "ADDPRINT3.ISPRINT = 1"
    cError = myDelete(cFilter)
    If cError = "ok" Then Inform " „ Õ–› ﬂ· «·”Ã·«  «·„ÿ»Ê⁄… »‰Ã«Õ" Else MsgBox cError
End Sub
Private Sub CmdDelAll_Click()
    If MsgBox("Õ–› «·ﬂ· ", vbOKCancel + vbDefaultButton2 + vbCritical, "Õ–› ﬂ· «·”Ã·« ") <> vbOK Then Exit Sub
    Dim cError As String
    cError = myDelete
    If cError = "ok" Then Inform " „ Õ–› ﬂ· «·”Ã·«  »‰Ã«Õ" Else MsgBox cError
End Sub
Private Sub cmdexit_Click()
    Unload Me
    Set barcodefrm = Nothing
End Sub
Private Sub cmdPrint_Click()
If Val(xRow.Text) > 16 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If
If Val(xCol.Text) > 6 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
If Not doprint Then Exit Sub
Set myForm = Me
CardPrintNew.Show 1
CmdDelPrinted_Click
End Sub
Private Sub cmdPrintStr_Click()

If Val(xRow.Text) > 8 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.Text) > 12 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
If Not doprintStr Then Exit Sub
Set myForm = Me
CardPrintNew.Show 1

'doprint

CmdDelPrinted_Click


End Sub
Private Sub Form_Load()
openCon con
Set grid1.DataSource = data1
data1.ConnectionString = strCon
myloadgrd
End Sub
Sub myloadgrd()
With grid1
'
cString = "Select ADDPRINT3.BARCODE ,ADDPRINT3.DESCA " & _
          " FROM ADDPRINT3 "
data1.RecordSource = cString
data1.Refresh
Fixgrd
End With
End Sub
Private Sub Fixgrd()
With grid1
    .FormatString = "«·ﬂÊœ|" & "«·⁄„Ì·"
    .FixedCols = 0
    .ColWidth(0) = 2000
    .ColWidth(1) = 5000
End With
End Sub
Private Function RemoveItem(Row As Long) As Boolean
con.BeginTrans
On Error GoTo myerror
If grid1.TextMatrix(Row, grid1.Cols - 1) <> "" Then
    con.Execute "Delete  From ADDPRINT3 where BARCODE = " & grid1.TextMatrix(Row, 0)
End If
grid1.RemoveItem Row
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function doprintStr() As Boolean
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(0)
SettingArray(cLeftMargin) = MyMeasure(0.25)
SettingArray(cCardWidth) = MyMeasure(3.5)
SettingArray(cCardHeight) = MyMeasure(1.848)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 16
SettingArray(cCols) = 6
SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select ADDPRINT3.BARCODE ,ADDPRINT3.DESCA " & _
          " FROM ADDPRINT3 "
cString = cString & "order by DESCA "

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
For I = 1 To nBegin - 1
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
    For I = 1 To 1
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = NCOLS)
        
        nHeight = 0
        nLast = MyMeasure(0.3)
        nLastCol = MyMeasure(0.4)
        For nCount = 1 To 1
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.7) + nHeight + IIf((nRow = 1 And nCount = 1), MyMeasure(0.2), MyMeasure(0))
            temptable!Width = MyMeasure(2.8)
            temptable!Height = MyMeasure(1)
            temptable!FontName = "Arial"
            temptable!FontBold = False
            temptable!FontSize = 12
            temptable!TextAlign = taCenterTop
            temptable!Text = sourcetable!desca
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.5) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) + IIf((nRow = 1 And nCount = 1), MyMeasure(0.2), MyMeasure(0))
            temptable!Width = MyMeasure(2.3)
            temptable!Height = MyMeasure(0.6)
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 9
            temptable!Text = sourcetable!Barcode
            temptable!isBarcode = True
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
Set sourcetable = Nothing
Set temptable = Nothing
doprintStr = True
End Function



Private Function myDelete(Optional cFilter As String) As String
Dim cString As String
cString = "DELETE FROM ADDPRINT3"
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


Private Function doprint()
contemp.Execute "DELETE * FROM TEMP"
If temptable.State = adStateOpen Then temptable.Close
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Dim nCount As Double


With grid1
nCount = (6 * (Val(xRow.Text) - 1) + Val(xCol.Text)) - 1
For I = 1 To nCount
    temptable.AddNew
    temptable!str8 = "."
    temptable.Update
Next I

For nRow = 1 To .Rows - 1
    temptable.AddNew
    temptable!str6 = .TextMatrix(nRow, 0)
    temptable!str3 = .TextMatrix(nRow, 1)
    temptable.Update
Next nRow
End With

contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\BARCUST.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Function

