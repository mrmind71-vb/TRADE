VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form branchbr_codesfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· ÊﬂÌ·« "
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12690
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   12690
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   6165
      MaskColor       =   &H00FFFFFF&
      Picture         =   "branchbr_code.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   4860
      UseMaskColor    =   -1  'True
      Width           =   1455
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
   Begin MSAdodcLib.Adodc DATA2 
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
   Begin MSAdodcLib.Adodc DATA3 
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4740
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   12480
      _cx             =   22013
      _cy             =   8361
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
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      RightToLeft     =   0   'False
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
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   7650
      TabIndex        =   2
      Top             =   4770
      Width           =   4920
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   4110
      End
      Begin VB.Label label2 
         Caption         =   "«·›—⁄"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   4275
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   585
      End
   End
End
Attribute VB_Name = "branchbr_codesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Dim con As New ADODB.Connection
Dim clist1 As String, cList2 As String
Public sCaption As String, sCode As String, sdesca As String, sGroupDesca As String, sTable As String, sTableGroup As String, sGroupCaption As String, nZero, nZeroGroup
Private Sub myload()
'                   0                   1                       2                   3               4                               5
cString = "SELECT CODE as [«·ﬂÊœ], DESCA as [«· ÊﬂÌ·], IPSQL AS [IP Number] , ISSTOP AS [„ Êﬁ›] , BRANCH_GR " & _
          " FROM BRANCH_FR "

If Trim(xDesca.text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(xDesca.text, "desca")
End If
cString = cString & " order by CODE"
data1.RecordSource = cString
data1.Refresh
myaddItem
Fixgrd
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Set FlagGroupFrm = Nothing
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
   myaddItem
End If

Dim aInsert As Variant
aInsert = AddFlag(Empty, "DESCA", addstring(.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "[IPSQL]", addstring(.TextMatrix(Row, 2)))
aInsert = AddFlag(aInsert, "[ISSTOP]", IIf(.TextMatrix(Row, 3), 1, 0))
aInsert = AddFlag(aInsert, "[BRANCH_GR]", addstring(.TextMatrix(Row, 4)))

On Error GoTo myerror
con.BeginTrans
If Not IsNumeric(.TextMatrix(Row, 0)) Then
    .TextMatrix(Row, 0) = RetZero(Newflag("BRANCH", "CODE", con), 2)
    aInsert = AddFlag(aInsert, "CODE", addstring(.TextMatrix(Row, 0)))
    con.Execute addInsert(aInsert, "BRANCH_FR")
Else
    con.Execute addUpdate(aInsert, "BRANCH_FR", "code = " & MyParn(.TextMatrix(Row, 0)))
End If
con.CommitTrans
End With
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myload
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    If MsgBox("⁄—÷ »Ì«‰«  «·›—⁄", vbYesNo) = vbYes Then
'        LoadConString_shop grid1.Row
'    End If
End Sub
Private Sub grid1_EnterCell()
If (grid1.Col >= 1) Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Form_Load()
    openCon con
    If cBranch = "00" Then con.Execute "UPDATE BRANCH_FR SET ISSTOP = 0 WHERE ISSTOP IS NULL "
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    With grid1
    myload
    CellPos 13, 0, .Cols - 1
    End With
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox sdesca & " „ÿ·Ê»"
        Cancel = True
    End If
ElseIf Col = 2 Then
     If Trim(grid1.EditText) = "" Then grid1.EditSelText = ""
End If
End Sub
Private Sub Fixgrd()
With grid1
.Cols = 5
.ColComboList(2) = clist1
.ColWidth(0) = 1000
.ColWidth(1) = 3000
.ColWidth(2) = 3000
.ColWidth(3) = 1500
.ColWidth(4) = 3000
.ColComboList(4) = StrListA("SELECT CODE , DESCA FROM BRANCH_GR ORDER BY CODE ", con)
.ColDataType(3) = flexDTBoolean
.RowHeight(0) = 1000
.WordWrap = True
.Cell(flexcpFontSize, 1, 3, .Rows - 1, .Cols - 1) = 9
.Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = &H8000000F
End With
End Sub
Private Sub xDesca_Change()
myload
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Trim(.TextMatrix(nRow, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, 0) = "" Then
    If Not validRow(OldRow) Then
 '       .RemoveItem OldRow
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
With grid1
If (Not validRow(.Row)) And .Row <> .Rows - 1 And .Row <> 0 And grid1.TextMatrix(OldRow, 0) = "" Then
'    .RemoveItem .Row
End If
End With
End Sub
Private Sub myaddItem()
Exit Sub
With grid1
    .AddItem ""
    grid1.Cell(flexcpBackColor, 1, 0, grid1.Rows - 1, 0) = &H8000000F
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 1 Then
    grid1.Select Row, IIf(Col + 1 < 3, 3, Col + 1)
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 3)
    grid1.ShowCell Row + 1, 1
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 And Col <> 2 Then CellPos KeyCode, Row, Col
End Sub
Function NextEmpty(pGrid As Object, Row As Long, Optional nBegincol As Long = -1, Optional nEndCol As Long = -1) As Long
Dim nLast
For i = IIf(nBegincol = -1, 0, nBegincol) To IIf(nEndCol = -1, pGrid.Cols - 1, IIf(nEndCol > pGrid.Cols - 1, pGrid.Cols - 1, nEndCol))
    If Trim(pGrid.TextMatrix(Row, i)) = "" And pGrid.ColHidden(i) = False Then
        NextEmpty = i
        Exit Function
    End If
Next
NextEmpty = IIf(nEndCol = -1, pGrid.Cols - 1, nEndCol)
End Function
