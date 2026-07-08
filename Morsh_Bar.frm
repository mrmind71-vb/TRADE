VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Begin VB.Form Morsh_Bar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
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
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox xIndex 
      Alignment       =   1  'Right Justify
      Caption         =   " — Ì» »«—ﬂÊœ"
      Height          =   375
      Left            =   7425
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   300
      Width           =   2040
   End
   Begin VB.CommandButton Command5 
      Caption         =   "≈÷«›… —’Ìœ „ÊœÌ·«  „Ê—œ"
      Height          =   540
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   225
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   9765
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   -75
      Width           =   1875
      Begin VB.CommandButton Command3 
         Caption         =   "«÷«›… „— Ã⁄"
         Height          =   540
         Left            =   4125
         RightToLeft     =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ÿ»«⁄… «” Ìﬂ—“"
         Height          =   540
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   300
         Width           =   1650
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
         Height          =   540
         Left            =   2850
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   300
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   6540
      Left            =   75
      TabIndex        =   0
      Top             =   975
      Width           =   11565
      _cx             =   20399
      _cy             =   11536
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
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
      AutoResize      =   -1  'True
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
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   5175
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   7425
      Width           =   6315
      Begin VB.CommandButton cmduno 
         Caption         =   " —«Ã⁄"
         Height          =   540
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ"
         Height          =   540
         Left            =   4875
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   300
         Width           =   1290
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Õ–› «·„ÿ»Ê⁄"
         Height          =   540
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelAll 
         Caption         =   "Õ–› «·ﬂ·"
         Height          =   540
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   300
         Width           =   1140
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Œ—ÊÃ"
         Height          =   540
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   1125
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   7425
      Width           =   3990
      Begin VB.TextBox xCol 
         Alignment       =   1  'Right Justify
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
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   375
         Width           =   915
      End
      Begin VB.TextBox xRow 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2175
         RightToLeft     =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   375
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄„Êœ :"
         Height          =   390
         Left            =   1275
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   375
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "«·’›:"
         Height          =   390
         Left            =   3225
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   375
         Width           =   615
      End
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "≈Ã„«·Ï ⁄œœ «·ﬁÿ⁄ ··ÿ»«⁄…"
      Height          =   390
      Left            =   4950
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   225
      Width           =   2040
   End
   Begin VB.Label xTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   150
      Width           =   1065
   End
End
Attribute VB_Name = "Morsh_Bar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim tAddPrint As New ADODB.Recordset
Dim FlagTable As New ADODB.Recordset
Dim NCOLS As Double
Dim NROWS As Double
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
'Grid1.SaveGrid App.Path & "\addPrint.grd", flexFileData
MyReplace
MsgBox " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
MyLoad
End Sub
Private Sub cmduno_Click()
    MyLoad
End Sub
Private Sub Command1_Click()
If MsgBox("Õ–› „«  „  ÿ»«⁄ Â „‰ «·„” ‰œ", vbYesNo + vbDefaultButton2) = vbYes Then
    delcheck
End If
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub CmdDelAll_Click()
If MsgBox("Õ–› ﬂ· «·”Ã·« ", vbYesNo + vbDefaultButton1) = vbYes Then
    Grid1.Rows = 1
    Grid1.AddItem ""
    con.Execute "delete from addprint"
End If
End Sub
Private Sub Command4_Click()
If Val(xRow.Text) > 16 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.Text) > 6 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If

MyReplace
'If Not myvalidRowCol Then Exit Sub
DoprintStr
Set Myform = Me
CardPrintNew.Show 1
If MsgBox("Õ–› „«  „  ÿ»«⁄ Â „‰ «·„” ‰œ", vbYesNo + vbDefaultButton2) = vbYes Then
    delcheck
End If
End Sub
Private Sub Command5_Click()
    BalBar.Show 1
    tAddPrint.Requery
    MyLoad
End Sub
Private Sub Form_Load()
openCon strCon
cStr1 = "Select file1_10.ModelFact,FILE1_10.CODE,addPrint.Item,addPrint.Model,addPrint.DescA,addPrint.Scal,addPrint.Price,addPrint.Quant,AddPrint.price2,addprint.isPrint,AddPrint.Doc_No From (addPrint Inner join File1_10 on AddPrint.Item = File1_10.item ) inner join fact on fact.code = file1_10.fact order by file1_10.code , file1_10.fact , file1_10.modelfact , file1_10.color ,  file1_10.scal "
With Grid1
    .Cols = 10
    .Rows = 2
    .FormatString = "»«—ﬂÊœ|" & "—ﬁ„ „ÊœÌ· |" & "«·„Ê—œ|" & "≈”„ «·„ÊœÌ·|" & "«·„ﬁ«”|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "”⁄— √Êﬂ«“ÌÊ‰|" & "«·ÿ»«⁄…|"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1300
    .ColWidth(2) = 800
    .ColWidth(3) = 3500
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColHidden(9) = True
    
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignLeftCenter
    .ColAlignment(3) = flexAlignLeftCenter
    .ColAlignment(4) = flexAlignRightCenter
    .ColAlignment(5) = flexAlignRightCenter
    .ColAlignment(6) = flexAlignRightCenter
    .ColDataType(8) = flexDTBoolean
End With
MyLoad
Grid1.Row = 1
Grid1.Col = 0
End Sub
Sub MyLoad()
Dim nTot As Double
nTot = 0
With Grid1
Grid1.Rows = 1
tAddPrint.Requery
If tAddPrint.RecordCount = 0 Then
    .AddItem ""
    Exit Sub
End If
tAddPrint.Requery
If tAddPrint.RecordCount = 0 Then Exit Sub
tAddPrint.MoveFirst
Do
   .AddItem ""
    If Not IsNull(tAddPrint!Item) And tAddPrint!Item <> "" Then
   .TextMatrix(Grid1.Rows - 1, 0) = TurnValue(tAddPrint!Item, Null, "")
   .TextMatrix(Grid1.Rows - 1, 1) = DelZero(TurnValue(retitem(tAddPrint!Item, "MODELFACT"), Null, ""))
   .TextMatrix(Grid1.Rows - 1, 2) = retitem(tAddPrint!Item, "CODE")
   .TextMatrix(Grid1.Rows - 1, 3) = SayCode(FlagTable, 3, ItemTable.FACT) & " / " & TurnValue(retitem(tAddPrint!Item, "DESCA"), Null, "")
   .TextMatrix(Grid1.Rows - 1, 4) = TurnValue(retitem(tAddPrint!Item, "SCAL"), Null, "")
   .TextMatrix(Grid1.Rows - 1, 5) = TurnValue(tAddPrint!Quant, Null, "")
   .TextMatrix(Grid1.Rows - 1, 6) = TurnValue(retitem(tAddPrint!Item, "PRICE"), Null, "")
   .TextMatrix(Grid1.Rows - 1, 7) = TurnValue(retitem(tAddPrint!Item, "PRICE2"), Null, "")
   .TextMatrix(Grid1.Rows - 1, 8) = IIf(tAddPrint!isPrint, "-1", "0")
   .TextMatrix(Grid1.Rows - 1, 9) = TurnValue(tAddPrint!doc_no, Null, "")
    nTot = nTot + Val(.TextMatrix(Grid1.Rows - 1, 5))
    End If
    tAddPrint.MoveNext
Loop Until tAddPrint.EOF
'CalcTotal
.AddItem ""
End With
Me.xTotal.Caption = Format(nTot, "#0")
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Grid1.Col = 0 Then GridRow (Row)
End Sub

Private Sub Grid1_EnterCell()
If Grid1.Col = 0 Or Grid1.Col = 5 Or Grid1.Col = 8 Then
    Grid1.Editable = flexEDKbdMouse
Else
    Grid1.Editable = flexEDNone
End If

End Sub
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And Grid1.Row <> Grid1.Rows - 1 Then
    If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        Grid1.RemoveItem Grid1.Row
        ' calcTotal
    End If
End If
If KeyCode = 112 And Grid1.Col = 0 Then
    CardLookup
End If
End Sub
Private Sub Grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Grid1.Col = 0 Then
    CardLookup
End If
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

If Grid1.Row = Grid1.Rows - 1 Then
    Grid1.AddItem ""
End If
End Sub
Private Sub MyReplace()
Dim DPrint As Recordset
Set DPrint = mydb.OpenRecordset("addprint")
mydb.Execute "Delete * From addprint"
' Õ–› Õ—ﬂ… √’‰«› «·„” ‰œ
With Grid1
For I = 1 To .Rows - 2
    If Not Grid1.RowHidden(I) Then
        If .TextMatrix(I, 0) <> "" Then
        DPrint.AddNew
        DPrint.Item = .TextMatrix(I, 0)
        DPrint.Quant = Val(.TextMatrix(I, 5))
        DPrint.isPrint = TurnValue(.TextMatrix(I, 8), "", False)
        DPrint.doc_no = TurnValue(.TextMatrix(I, 9), "", Null)
        DPrint.Update
        End If
    End If
Next
End With
MyLoad
End Sub
Private Function MYVALID() As Boolean
With Grid1
For I = 1 To Grid1.Rows - 2
    If Val(.TextMatrix(I, 5)) = 0 Then
        MsgBox "«·ﬂ„Ì… €Ì— „”Ã·…"
        Exit Function
    End If
Next
MYVALID = True
End With
End Function
Private Sub delcheck()
For I = 1 To Grid1.Rows - 2
   If Val(Grid1.TextMatrix(I, 8)) <> 0 Then
        Grid1.RowHidden(I) = True
   End If
Next
MyReplace
End Sub
Sub ItemsLookup()
Dim Generalarray(4)
Dim GrdArray(7)
Set Generalarray(1) = Me
If Grid1.TextMatrix(Grid1.Row, 0) <> "" Then
    Generalarray(2) = "Select ITEM , MODEL as [ﬂÊœ «·„ÊœÌ· ] ,  DescA as [«”„ «·’‰›] , SCAL AS „ﬁ«” , COLOR AS  ·Ê‰ , PRICE AS [”⁄—»Ì⁄] , PRICE2 AS [”⁄— √Êﬂ«“ÌÊ‰],COST AS [«· ﬂ·›…] From file1_10 WHERE MODEL LIKE " & MyParn("*" & Grid1.TextMatrix(Grid1.Row, 0) & "*")
Else
    Generalarray(2) = "Select ITEM , MODEL as [ﬂÊœ «·„ÊœÌ· ] ,  DescA as [«”„ «·’‰›] , SCAL AS „ﬁ«” , COLOR AS  ·Ê‰ , PRICE AS [”⁄—»Ì⁄] , PRICE2 AS [”⁄— √Êﬂ«“ÌÊ‰],COST AS [«· ﬂ·›…] From file1_10 WHERE MODEL LIKE " & MyParn("*" & Grid1.TextMatrix(Grid1.Row, 0) & "*")
End If
Generalarray(3) = " MODEL Like('%cFilter%') "
Generalarray(4) = " MODEL Like('%cFilter%') "

GrdArray(1) = 1500
GrdArray(2) = 1800
GrdArray(3) = 2700
GrdArray(4) = 900
GrdArray(5) = 900
GrdArray(6) = 900
GrdArray(7) = 900

Lookupdata = Array(Generalarray, GrdArray)
Load Search1
Search1.Text1 = Grid1.TextMatrix(Grid1.Row, 0)
Search1.Caption = "«” ⁄·«„ "
Search1.Show 1
End Sub
Sub myProc()
Grid1.TextMatrix(Grid1.Row, 0) = Search.Grid1.TextMatrix(Search.Grid1.Row, 0)
GridRow (Grid1.Row)
If Grid1.Row = Grid1.Rows - 1 Then Grid1.AddItem ""
Unload Search
End Sub
Private Sub GridRow(nRow)
On Error Resume Next
With Grid1
ItemTable.Index = "nItem"
ItemTable.Seek "=", Grid1.TextMatrix(nRow, 0)

If ItemTable.NoMatch Then
    Grid1.TextMatrix(nRow, 1) = ""
    Grid1.TextMatrix(nRow, 2) = ""
    Grid1.TextMatrix(nRow, 3) = ""
    Grid1.TextMatrix(nRow, 4) = ""
    Grid1.TextMatrix(nRow, 5) = ""
    Grid1.TextMatrix(nRow, 6) = ""
    Grid1.TextMatrix(nRow, 7) = ""
    Grid1.TextMatrix(nRow, 8) = ""
Else
    Grid1.TextMatrix(nRow, 1) = TurnValue(retitem(.TextMatrix(nRow, 0), "ModelFact"), Null, "")
    Grid1.TextMatrix(nRow, 2) = TurnValue(retitem(.TextMatrix(nRow, 0), "CODE"), Null, "")
    Grid1.TextMatrix(nRow, 3) = SayCode(FlagTable, 3, ItemTable.FACT) & " / " & TurnValue(retitem(.TextMatrix(nRow, 0), "Desca"), Null, "")
    Grid1.TextMatrix(nRow, 4) = TurnValue(retitem(.TextMatrix(nRow, 0), "Scal"), Null, "")
    Grid1.TextMatrix(nRow, 5) = ""
    Grid1.TextMatrix(nRow, 6) = TurnValue(retitem(.TextMatrix(nRow, 0), "Price"), Null, "")
    Grid1.TextMatrix(nRow, 7) = TurnValue(retitem(.TextMatrix(nRow, 0), "Price2"), Null, "")
    Grid1.TextMatrix(nRow, 8) = "-1"
End If
End With
End Sub
Private Sub DoprintStr()
Dim tCard As Recordset
Dim tPrint As Recordset
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(-0.15)
SettingArray(cLeftMargin) = MyMeasure(0.4)
SettingArray(cCardWidth) = MyMeasure(3.5)
SettingArray(cCardHeight) = MyMeasure(1.852)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 16
SettingArray(cCols) = 6
SettingArray(cPageWidth) = MyMeasure(21)

mydb.Execute "delete * From Card"
Set tCard = mydb.OpenRecordset("Select * From card")
If xIndex.Value Then
    Set tPrint = mydb.OpenRecordset("Select file1_10.model,file1_10.price2 as price2 ,file1_10.ModelFact,FILE1_10.SUPP ,FILE1_10.MOSM,addPrint.Item,addPrint.DescA,addPrint.Scal,addPrint.Price,addPrint.Quant,AddPrint.price2 From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint  ORDER BY FILE1_10.ITEM")
Else
    Set tPrint = mydb.OpenRecordset("Select file1_10.model,file1_10.price2 as price2 ,file1_10.ModelFact,FILE1_10.SUPP ,FILE1_10.MOSM,addPrint.Item,addPrint.DescA,addPrint.Scal,addPrint.Price,addPrint.Quant,AddPrint.price2 From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint  ORDER BY FILE1_10.FACT , FILE1_10.MODEL , FILE1_10.C_COLOR , FILE1_10.C_SCAL ")
End If
With tCard
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
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡
If tPrint.RecordCount = 0 Then Exit Sub
nadd = MyMeasure(0.2)
Do
'************
    For I = 1 To tPrint.Quant
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
            tCard.AddNew
            tCard!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(1.8)
            tCard!Height = MyMeasure(0.3) + nadd
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontSize = 8
            tCard!TextAlign = taRightTop
            tCard!Text = tPrint!Item
            tCard!isBarcode = True
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update


            tCard.AddNew
            tCard!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = 0
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = False
            tCard!FontUnderline = True
            tCard!FontSize = 7
            tCard!TextAlign = taLeftTop
            tCard!Text = DelZero(retitem(tPrint!Item, "MODELFACT"))
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.8) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = MyMeasure(2)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!TextAngle = 900
            tCard!FontBold = True
            tCard!FontUnderline = False
            tCard!FontSize = 7
            tCard!Text = "«”Ê«ﬁ «·„’ÿ›Ï"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(2)
            tCard!TextAlign = taRightTop
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!FontSize = 8
            tCard!Text = tPrint!Item
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(2) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = 0
            tCard!TextAlign = taRightTop
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontSize = 5
            tCard!Text = Format(Int(tPrint!PRICE2) + 5, "#")
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(2)
            tCard!TextAlign = taRightTop
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!FontSize = 8
            tCard!Text = tPrint!Item
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(2.8) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.8) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = 0
            tCard!Height = MyMeasure(0.35)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontSize = 7
            'tCard!Text = tPrint!Model
            tCard!Text = tPrint!SUPP
            tCard!TextAlign = taRightTop
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
        
        
            tCard.AddNew
            tCard!Left = MyMeasure(2.8) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = 0
            tCard!TextAlign = taRightTop
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontSize = 5
            tCard!Text = Format(Int(retitem(tPrint!Item, "COST")) + 20, "#")
            tCard!CardNo = nCard
            tCard.Update

            tCard.AddNew
            tCard!Left = MyMeasure(2.8) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = 0
            tCard!Height = MyMeasure(0.35)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontSize = 6
            'tCard!Text = tPrint!Model
            tCard!Text = Format(Int(tPrint!price) + 10, "#")
            tCard!TextAlign = taRightTop
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
        
            tCard.AddNew
            tCard!Left = MyMeasure(1.7) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(1)
            tCard!Height = MyMeasure(0.35)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontSize = 7
            'tCard!Text = tPrint!Model
    '       tCard!Text = tPrint!MOSM
            tCard!TextAlign = taRightTop
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
        
            tCard.AddNew
            tCard!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(1)
            tCard!Top = MyMeasure(0.45) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Height = 0
            tCard!TextAlign = taCenterTop
            tCard!FontName = "Simplified arabic"
            tCard!FontBold = False
            tCard!FontSize = 7
            If Not IsNull(retitem(tPrint.Item, "Color")) Then
                tCard!Text = Mid(IIf(retitem(tPrint.Item, "Color") = "„ÊÕœ" Or retitem(tPrint.Item, "Color") = "√·Ê«‰", Null, retitem(tPrint.Item, "Color")), 1, 10)
            End If
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
    
            tCard.AddNew
            tCard!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.27) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = MyMeasure(2.5)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!TextAlign = taCenterTop
            tCard!FontBold = Bold
            tCard!ForeColor = vbBlack
            tCard!FontSize = 7
            tCard!Text = Trim(retitem(tPrint!Item, "desca"))
            tCard!CardNo = nCard
            tCard.Update
    
'            tCard.AddNew
'            tCard!Left = MyMeasure(2) - IIf(blastcol, nLastCol, 0)
'            tCard!Top = MyMeasure(1.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
'            tCard!Width = MyMeasure(1.3)
'            tCard!Height = 0
'            tCard!FontName = "Simplified arabic"
'            tCard!TextAlign = taRightTop
'            tCard!FontBold = False
'            tCard!ForeColor = vbBlack
'            tCard!FontSize = 7
'            tCard!Text = Trim(SayCode(FlagTable, 3, RetItem(tPrint!Item, "FACT")))
'            tCard!CardNo = nCard
'            tCard.Update
    
            tCard.AddNew
            tCard!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = 0
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!FontUnderline = False
            tCard!FontSize = 8
            tCard!Text = Format(retitem(tPrint!Item, "price"), "Fixed") & "Ã‰Ì…"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = MyMeasure(2)
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!TextAlign = taRightTop
            tCard!FontUnderline = True
            tCard!FontSize = 7
            tCard!ForeColor = vbBlack
            tCard!Text = retitem(tPrint!Item, "scal")
            tCard!CardNo = nCard
            tCard.Update
            
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    tPrint.MoveNext
Loop Until tPrint.EOF
End With
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
Private Sub CardLookup()
    Dim cModelFact As String
    Dim Generalarray(3)
    Dim GrdArray(8)
    Set Generalarray(1) = Me
    cModelFact = RetZero(Grid1.TextMatrix(Grid1.Row, 0))
    
    Generalarray(2) = "Select FILE1_10.ITEM AS [»«— ﬂÊœ]  , FILE1_10.MODELFACT as „ÊœÌ·, FILE1_10.desca as ’‰› , FILE1_70.desca as „’‰⁄ , FILE1_10.mosm as „Ê”„ , FILE1_10.code as „Ê—œ  , scal as [„ﬁ«”] , color as  [·Ê‰] FROM FILE1_10 LEFT JOIN FILE1_70 ON FILE1_10.Fact = FILE1_70.CODE WHERE flag = 3 and modelfact LIKE " & MyParn(cModelFact & "%")
    Generalarray(3) = " AND (file1_10.desca Like('%cFilter%') OR file1_70.desca Like('%cFilter%') )"
    GrdArray(1) = 1000
    GrdArray(2) = 1300
        
    GrdArray(3) = 1800
    GrdArray(4) = 1800
    GrdArray(5) = 1800
    GrdArray(6) = 1600
    GrdArray(7) = 800
    GrdArray(8) = 800
    
    Lookupdata = Array(Generalarray, GrdArray)
    Load Search
    Search.Caption = "«” ⁄·«„ "
    Search.Show 1
        
End Sub

