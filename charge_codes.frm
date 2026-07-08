VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form chargeCodesfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10185
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5880
   ScaleWidth      =   10185
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4740
      Left            =   270
      TabIndex        =   0
      Top             =   0
      Width           =   9825
      _cx             =   17330
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
   Begin Threed.SSCommand cmdExcel 
      Height          =   600
      Left            =   2790
      TabIndex        =   7
      Top             =   4815
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1058
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "charge_codes.frx":0000
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdExit 
      Height          =   600
      Left            =   270
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4815
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1058
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "charge_codes.frx":252B
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdSql 
      Height          =   600
      Left            =   1530
      TabIndex        =   9
      Top             =   4815
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   1058
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "charge_codes.frx":484E
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   4005
      TabIndex        =   3
      Top             =   4725
      Width           =   6090
      Begin VB.CommandButton cmdGroup 
         BackColor       =   &H00FFFFFF&
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   330
      End
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
         TabIndex        =   2
         Top             =   540
         Width           =   4110
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   330
         Left            =   450
         TabIndex        =   1
         Top             =   180
         Width           =   3750
         _ExtentX        =   6615
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„’—Ê› «·—∆Ì”Ì"
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
         TabIndex        =   5
         Top             =   225
         Width           =   1770
      End
      Begin VB.Label label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·»Ì«‰"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   675
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -450
      Top             =   720
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
      Left            =   -495
      Top             =   1080
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
      Left            =   -1665
      Top             =   1215
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
Attribute VB_Name = "chargeCodesfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Integer
Public sFileGroup As String, sFile As String, sdesca As String
Public bedit As Boolean
Dim sTable As String
Dim CLIST As String, cListString As String
Dim con As New ADODB.Connection
Private Sub myload(Optional bSql As Boolean = False)
Dim cString As String, cWhere As String

cString = "SELECT CODE AS [«·ﬂÊœ], DESCA as " & IIf(myPublic = 1, "[«·„’—Ê›]", "[«·«Ì—«œ]") & ",[MAINGROUP] as [«·„Ã„Ê⁄… «·—∆Ì”Ì…],[ISASSET] as [«’Ê·],IS_BRANCH AS " & IIf(myPublic = 1, "[„’—Ê› ›—Ê⁄]", "[«Ì—«œ ›—Ê⁄]") & _
          " FROM " & IIf(myPublic = 1, "FILE8_51", "FILE8_61")

If xGroup.MatchedWithList Then
    cWhere = cWhere & Tr(Where) & " MAINGROUP LIKE " & MyParn(xGroup.BoundText)
End If

If Trim(xDesca.text) <> "" Then
    cWhere = cWhere & Tr(Where) & MyParnAnd(xDesca.text, "desca")
End If

If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " order by CODE"

If bSql Then
    Clipboard.Clear
    Clipboard.SetText cString
    Exit Sub
End If

Set data1.Recordset = mycmd(cString, con)
myAddItem

If xGroup.MatchedWithList Then grid1.TextMatrix(grid1.Rows - 1, 2) = xGroup.BoundText
fixGrd
End Sub

Private Sub cmdExcel_Click()
Me.MousePointer = 11
'Dim i As Long
Dim aRow As Variant
'If grid1.Rows > 1 Then
'    aSub = AddFlag(Empty, "row", 1)
'    aSub = AddFlag(aSub, "bold", True)
'    aSub = AddFlag(aSub, "word_wrap", False)
'    aSub = AddFlag(aSub, "back_color", 19)
'    aRow = AddFlag(aRow, aSub)
'End If
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption)
Me.MousePointer = 0
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdGroup_Click()
Dim sBound As String
sBound = xGroup.BoundText
Dim oFlagfrm As New flag_mainfrm
oFlagfrm.sFieldCode = "[CODE]"
oFlagfrm.sFieldDesca = "[DESCA]"
oFlagfrm.sFieldName1 = "«·ﬂÊœ"
oFlagfrm.sFieldName2 = IIf(myPublic = 1, "«·„’—Ê›", "«·«Ì—«œ")
oFlagfrm.sCaption = IIf(myPublic = 1, "„Ã„Ê⁄… „’—Ê›« ", "„Ã„Ê⁄… «·«Ì—«œ« ")
oFlagfrm.sTable = IIf(myPublic = 1, "FILE8_52", "FILE8_62")
oFlagfrm.nZero = 3
oFlagfrm.bedit = bedit
oFlagfrm.Show 1

Set DATA2.Recordset = mycmd(cListString, con)
xGroup.BoundText = sBound
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
CLIST = StrList(cListString, con)
fixGrd
End Sub

Private Sub cmdSql_Click()
myload True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Set chargecodefrm = Nothing
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
With grid1

If Not validRow(Row) Then Exit Sub

If grid1.Row = grid1.Rows - 2 Then
    grid1.TextMatrix(grid1.Rows - 1, 2) = grid1.TextMatrix(grid1.Rows - 2, 2)
End If

If Row = .Rows - 1 Then
   myAddItem
End If

Dim aInsert As Variant
aInsert = AddFlag(Empty, "desca", addstring(.TextMatrix(Row, 1)))
aInsert = AddFlag(aInsert, "[MAINGROUP]", addstring(.TextMatrix(Row, 2)))
aInsert = AddFlag(aInsert, "[ISASSET]", IIf(.ValueMatrix(Row, 3) = 0, 0, 1))
aInsert = AddFlag(aInsert, "[IS_BRANCH]", IIf(.ValueMatrix(Row, 4) = 0, 0, 1))

On Error GoTo myerror
If .TextMatrix(Row, 0) = "" Then
    .TextMatrix(Row, 0) = RetZero(Newflag(IIf(myPublic = 1, "FILE8_51", "fILE8_61"), "CODE", con), 3)
    aInsert = AddFlag(aInsert, "CODE", addstring(.TextMatrix(Row, 0)))
    con.Execute addInsert(aInsert, IIf(myPublic = 1, "FILE8_51", "FILE8_61"))
Else
    con.Execute addUpdate(aInsert, IIf(myPublic = 1, "FILE8_51", "FILE8_61"), "CODE = " & addstring(.TextMatrix(Row, 0)))
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
myload
End Sub
Private Sub grid1_EnterCell()
If Not bedit Then
    grid1.Editable = flexEDNone
ElseIf grid1.col = 0 Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Form_Load()
openCon con

Me.Caption = IIf(myPublic = 1, "«ﬂÊ«œ „’«—Ì›", "«ﬂÊ«œ «Ì—«œ« ")

cListString = "SELECT CODE,DESCA FROM " & IIf(myPublic = 1, "FILE8_52", "FILE8_62") & " ORDER BY DESCA"
Set DATA2.Recordset = mycmd(cListString, con)
Set xGroup.RowSource = DATA2
xGroup.ListField = "DESCA"
xGroup.BoundColumn = "CODE"

Set grid1.DataSource = data1
With grid1


CLIST = StrList(cListString, con)

myload
CellPos 13, grid1.Rows - 2, grid1.Cols - 1
End With
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
If col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox sdesca & " „ÿ·Ê»"
        Cancel = True
    End If
ElseIf col = 2 Then
     If Trim(grid1.EditText) = "" Then grid1.EditSelText = ""
End If
End Sub
Private Sub fixGrd()
With grid1
.ColComboList(2) = CLIST
.ColWidth(0) = 1000
.ColWidth(1) = 3500
.ColWidth(2) = 2500
.ColWidth(3) = 800
.ColWidth(4) = 1200
.ColDataType(3) = flexDTBoolean

For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = &H8000000F
End With
End Sub
Private Sub xDesca_Change()
myload
End Sub
Private Function validRow(Row) As Boolean
With grid1
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
If bRequired Then If .TextMatrix(Row, 2) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub myAddItem()
With grid1
    .AddItem ""
    If .Rows > 2 Then
        grid1.TextMatrix(.Rows - 1, 2) = grid1.TextMatrix(.Rows - 2, 2)
    End If
    If xGroup.MatchedWithList And grid1.TextMatrix(.Rows - 1, 2) = "" Then
        grid1.TextMatrix(.Rows - 1, 2) = xGroup.BoundText
    End If
    grid1.Cell(flexcpBackColor, 1, 0, grid1.Rows - 1, 0) = &H8000000F
End With
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If col < grid1.Cols - 1 Then
    grid1.Select Row, col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 1)
    grid1.ShowCell Row + 1, 1
Else
    grid1.Select Row, col
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
If KeyCode = 46 And bedit Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Õ–›?? Â· √‰  „Ê«›ﬁ", vbYesNo + vbDefaultButton2 + vbCritical, "Õ–›") = vbYes Then
            If IsNumeric(grid1.TextMatrix(grid1.Row, 0)) Then
                con.BeginTrans
                con.Execute "Delete From " & IIf(myPublic = 1, "FILE8_51", "FILE8_61") & " where code =  " & MyParn(grid1.TextMatrix(grid1.Row, 0))
                con.CommitTrans
            End If
            grid1.RemoveItem grid1.Row
        End If
    End If
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
myload
End Sub
Private Sub xGroup_Change()
If xGroup.MatchedWithList Or Trim(xGroup.BoundText) = "" Then
    myload
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grid1.col = 2 And grid1.TextMatrix(grid1.Row, grid1.col) = "" Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If col = 2 And grid1.TextMatrix(Row, col) = "" Then Exit Sub
    CellPos KeyCode, Row, col
End If
End Sub
Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal col As Long, KeyAscii As Integer)
If KeyAscii = 46 And grid1.ColComboList(col) <> "" Then
    grid1.TextMatrix(Row, col) = ""
    grid1.EditText = ""
End If
End Sub
Private Sub xDescA_GotFocus()
myGotFocus xDesca
End Sub
Private Sub xDesca_LostFocus()
myLostFocus xDesca
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
