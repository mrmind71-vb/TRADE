VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form sectionfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·«ﬁ”«„"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6180
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4365
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
      Top             =   3825
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   1530
      TabIndex        =   2
      Top             =   3735
      Width           =   4515
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   135
         Width           =   2445
      End
      Begin VB.Label Label1 
         Caption         =   "«·»Ì«‰ :"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1635
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3690
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   6000
      _cx             =   10583
      _cy             =   6509
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
   Begin MSAdodcLib.Adodc DATA1 
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
Attribute VB_Name = "sectionfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean, nRaw As Integer
Dim con As New ADODB.Connection
Dim clist1 As String, cList2 As String
Dim CardTable As New ADODB.Recordset
Private Sub myload()

cString = "SELECT CODE as [«·ﬂÊœ], DESCA as [«·»Ì«‰],CODE AS [ID]" & _
          " FROM FILE1_10SC "

If Trim(xDesca.text) <> "" Then
    cString = cString & turn(cString) & MyParnAnd(xDesca.text, "desca")
End If

cString = cString & " order by CODE"
data1.RecordSource = cString
data1.Refresh
grid1.AddItem ""
Fixgrd
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
If Not bedit Then Exit Sub
Dim aInsert(1, 1)
aInsert(0, 0) = "Code"
aInsert(0, 1) = addstring(grid1.TextMatrix(Row, 0))

aInsert(1, 0) = "desca"
aInsert(1, 1) = addstring(grid1.TextMatrix(Row, 1))

On Error GoTo myerror
con.BeginTrans
If grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
    con.Execute CreateInsert(aInsert, "FILE1_10SC")
    grid1.TextMatrix(Row, grid1.Cols - 1) = grid1.TextMatrix(grid1.Row, 0)
Else
    con.Execute CreateUpdate(aInsert, "FILE1_10SC", " WHERE FILE1_10SC.code = " & MyParn(grid1.TextMatrix(Row, grid1.Cols - 1)))
End If
con.CommitTrans
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myload
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
    grid1.Editable = flexEDNone
ElseIf grid1.Col = 1 And grid1.TextMatrix(grid1.Row, 0) = "" Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
If grid1.Row = grid1.Rows - 1 Then Exit Sub
If KeyCode = 46 And bedit Then
        If MsgBox("Õ–›?? Â· √‰  „Ê«›ﬁ", vbYesNo + vbCritical, "Õ–›") = vbYes Then
            If Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) <> "" Then
                con.BeginTrans
                con.Execute "Delete From FILE1_10SC where code =  " & MyParn(grid1.TextMatrix(grid1.Row, 0))
                con.CommitTrans
            End If
            grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
myload
End Sub
Private Sub Form_Load()
openCon con
Set grid1.DataSource = data1
data1.ConnectionString = strCon
myload
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Grid1.Row = Grid1.Rows - 20 Then
If grid1.Row = grid1.Rows - 1 Then grid1.Rows = grid1.Rows + 1
End Sub
Private Sub xcountry_code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then myload
End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "ﬂÊœ «·ﬁ”„ „ÿ·Ê»"
        Cancel = True
        Exit Sub
    End If
End If
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "Ê’› «·ﬁ”„ „ÿ·Ê»"
        Cancel = True
    End If
End If
grid1.EditText = Trim(UCase(grid1.EditText))
End Sub
Private Sub Fixgrd()
With grid1
.ColWidth(0) = 1500
.ColWidth(1) = 4000
.ColHidden(.Cols - 1) = True
For i = 1 To grid1.Cols - 1
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
myload
End Sub
