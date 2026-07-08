VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form itemsGroupFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„Ã„Ê⁄… «·„ÊœÌ·«  «·—∆Ì”Ì…"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4740
   ScaleWidth      =   6180
   Begin VB.CommandButton Command2 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   1530
      TabIndex        =   2
      Top             =   3690
      Width           =   4515
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   540
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo XGROUP 
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   180
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ã„Ê⁄… «·—∆Ì”Ì… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2685
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "≈”„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2685
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   555
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3630
      Left            =   45
      TabIndex        =   6
      Top             =   45
      Width           =   6000
      _cx             =   10583
      _cy             =   6403
      _ConvInfo       =   1
      Appearance      =   0
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
End
Attribute VB_Name = "itemsGroupFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nRaw, bEdit As Boolean
Dim con As New ADODB.Connection
Dim clist1 As String
Dim CardTable As New ADODB.Recordset
Private Sub MyLoad()
cString = "SELECT CODE as [«·ﬂÊœ], DESCA as [«·»Ì«‰],[GROUP] as [«·„Ã„Ê⁄…]" & _
          " FROM FILE1_50 WHERE ISRAW = " & nRaw

If IsNumeric(xGroup.BoundText) Then
    cString = cString & turn(cString) & "[GROUP] = " & xGroup.BoundText
End If

If Trim(xDescA.Text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(xDescA.Text, "desca")
End If

cString = cString & " order by CODE"
data1.RecordSource = cString
data1.Refresh
Grid1.AddItem ""
If xGroup.BoundText <> "" Then
   Grid1.TextMatrix(Grid1.Rows - 1, 2) = xGroup.BoundText
End If
FixGrd
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
Err.Clear
Set FlagFrm = Nothing
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not bEdit Then Exit Sub
If Not validRow(Row) Then Exit Sub
If Row = Grid1.Rows - 1 Then Grid1.AddItem ""

Dim aInsert(2, 1)
aInsert(0, 0) = "Code"
aInsert(0, 1) = addstring(Grid1.TextMatrix(Row, 0))

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(Grid1.TextMatrix(Row, 1))

aInsert(2, 0) = "[Group]"
aInsert(2, 1) = addstring(Grid1.TextMatrix(Row, 2))

On Error GoTo myerror
con.BeginTrans
If Not IsNumeric(Grid1.TextMatrix(Row, 0)) Then
    cString = "Select Max(Code) From File1_50 WHERE ISRAW = 0"
    Grid1.TextMatrix(Row, 0) = RetZero(Val(GetDesca(cString)) + 1)
    aInsert(0, 1) = addstring(Grid1.TextMatrix(Row, 0))
    con.Execute CreateInsert(aInsert, "FILE1_50")
Else
    con.Execute CreateUpdate(aInsert, "FILE1_50", " WHERE FILE1_50.code = " & MyParn(Grid1.TextMatrix(Row, 0)))
End If
con.CommitTrans
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
MyLoad
End Sub

Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow And OldRow <> Grid1.Rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then Grid1.RemoveItem OldRow
End If
End Sub

Private Sub grid1_EnterCell()
If (Grid1.Col = 0) Then
    Grid1.Editable = flexEDNone
Else
    Grid1.Editable = flexEDKbdMouse
    
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
If KeyCode = 46 And bEdit Then
    If Trim(Grid1.TextMatrix(Grid1.Row, 0)) <> "" Then
        If MsgBox("Õ–›?? Â· √‰  „Ê«›ﬁ", vbYesNo + vbCritical, "Õ–›") = vbYes Then
            If IsNumeric(Grid1.TextMatrix(Grid1.Row, 0)) Then
                con.BeginTrans
                con.Execute "Delete From file1_50 where code =  " & MyParn(Grid1.TextMatrix(Grid1.Row, 0))
                con.CommitTrans
            End If
            Grid1.RemoveItem Grid1.Row
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
MyLoad
End Sub
Private Sub Form_Load()
openCon con
DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM FILE1_50G WHERE ISRAW = " & nRaw

Set xGroup.RowSource = DATA2
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

Set Grid1.DataSource = data1
data1.ConnectionString = strCon
With Grid1
clist1 = StrList("Select code,desca from file1_50G where israw = " & nRaw & " order by desca")
MyLoad
End With
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Grid1.Row = Grid1.Rows - 20 Then
If xGroup.BoundText <> "" Then
   Grid1.TextMatrix(Grid1.Rows - 1, 2) = xGroup.BoundText
End If
End Sub
Private Sub xcountry_code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then MyLoad
End Sub

Private Sub grid1_Validate(Cancel As Boolean)
If Not validRow(Grid1.Row) And Grid1.Row <> Grid1.Rows - 1 And Grid1.Row <> 0 Then Grid1.RemoveItem Grid1.Row
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(Grid1.EditText) = "" Then
        MsgBox "ﬂÊœ «·’‰› „ÿ·Ê»"
        Cancel = True
        Exit Sub
    End If
End If
If Col = 1 Then
    If Trim(Grid1.EditText) = "" Then
        MsgBox "Ê’› «·„Ã„Ê⁄… „ÿ·Ê»"
        Cancel = True
    End If
End If
End Sub
Private Sub FixGrd()
With Grid1
.ColComboList(2) = clist1
.ColWidth(0) = 700
.ColWidth(1) = 2300
.ColWidth(2) = 2300
'.ColHidden(0) = True
For i = 1 To Grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Function StrList(cString)
Dim listTable As New ADODB.Recordset
listTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until listTable.EOF
    StrList = StrList & "|#" & listTable.Fields(0) & ";" & listTable.Fields(1)
    listTable.MoveNext
Loop
End Function
Private Sub xDesca_Change()
MyLoad
End Sub
Private Sub xGroup_Click(Area As Integer)
If Area = 2 Then MyLoad
End Sub
Private Sub xgroup_Validate(Cancel As Boolean)
MyLoad
End Sub
Private Function validRow(nRow) As Boolean
If Trim(Grid1.TextMatrix(nRow, 1)) = "" Then Exit Function
validRow = True
End Function
