VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form FlagFrm2 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   5400
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
      Left            =   135
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1455
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   3645
      Left            =   135
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   5190
      _cx             =   9155
      _cy             =   6429
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
      Cols            =   2
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
      TabBehavior     =   0
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
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   90
      TabIndex        =   3
      Top             =   3690
      Width           =   5190
      Begin VB.TextBox xname 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   4380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   555
      End
   End
End
Attribute VB_Name = "FlagFrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nZeros As Integer
Dim con As New ADODB.Connection
Dim GrdTable As New ADODB.Recordset
Public myPublic, bedit As Boolean
Private Sub myload()
Dim cFilter As String
GrdTable.Requery
If Trim(xname.text) <> "" Then cFilter = MyParnAnd(xname.text, myPublic(2))
GrdTable.Filter = cFilter
i = 0
Grid1.Rows = 1
Do Until GrdTable.EOF
    Grid1.AddItem ""
    Grid1.TextMatrix(Grid1.Rows - 1, 0) = GrdTable(myPublic(1)) & ""
    Grid1.TextMatrix(Grid1.Rows - 1, 1) = GrdTable(myPublic(2)) & ""
    GrdTable.MoveNext
Loop
Grid1.AddItem ""
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Set FlagFrm = Nothing
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Not bedit Then Exit Sub
Dim aInsert(1, 1)

aInsert(0, 0) = myPublic(1)
aInsert(0, 1) = addstring(Grid1.TextMatrix(Row, 0))

aInsert(1, 0) = myPublic(2)
aInsert(1, 1) = addstring(Grid1.TextMatrix(Row, 1))

On Error GoTo myerror
con.BeginTrans
If Not IsNumeric(Grid1.TextMatrix(Row, 0)) Then
    If myPublic(0) = "FILE0_40" Then
        Grid1.TextMatrix(Row, 0) = IncRec(GetDesca("SELECT MAX(CODE) FROM FILE0_40 WHERE LEN(CODE) = 3 ", con))
    Else
        Grid1.TextMatrix(Row, 0) = RetZero(Newflag(myPublic(0), myPublic(1), con), nZeros)
    End If
    aInsert(0, 1) = addstring(Grid1.TextMatrix(Row, 0))
    con.Execute CreateInsert(aInsert, myPublic(0))
Else
    con.Execute CreateUpdate(aInsert, myPublic(0), " WHERE " & myPublic(1) & " = " & Grid1.TextMatrix(Row, 0))
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
If Grid1.Row > Grid1.Rows - 1 Or Not bedit Then
    Grid1.Editable = flexEDNone
Else
    Grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
If KeyCode = 46 And bedit Then
    If Trim(Grid1.TextMatrix(Grid1.Row, 0)) <> "" Then
        If MsgBox("Õ–›?? Â· √‰  „Ê«›ﬁ", vbYesNo + vbCritical, "Õ–›") = vbYes Then
            If IsNumeric(Grid1.TextMatrix(Grid1.Row, 0)) Then
                con.BeginTrans
                con.Execute "Delete From " & myPublic(0) & " where code =  " & Grid1.TextMatrix(Grid1.Row, 0)
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
myload
End Sub
Private Sub Form_Load()
nZeros = myPublic(6)
If nZeros = 0 Then
    nZeros = 2
ElseIf nZeros = -1 Then
    nZeros = 0
End If
openCon con
If myPublic(0) = "FILE0_40" Then
    cString = "SELECT  *  from " & myPublic(0) & " WHERE LEN(CODE) = 3 order by code"
Else
    cString = "SELECT  *  from " & myPublic(0) & " order by code"
End If
GrdTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
With Grid1
.ColWidth(0) = 1200
.ColWidth(1) = Grid1.Width - 1600
.TextMatrix(0, 0) = myPublic(3)
.TextMatrix(0, 1) = myPublic(4)
Me.Caption = myPublic(5)
.ColAlignment(0) = flexAlignRightCenter
.ColAlignment(1) = flexAlignRightCenter
myload
End With
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Grid1.Row = Grid1.Rows - 20 Then
If Grid1.Row = Grid1.Rows - 1 Then Grid1.Rows = Grid1.Rows + 1
End Sub
Private Sub xcountry_code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then myload
End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If Trim(Grid1.EditText) = "" Then
        MsgBox "description required"
        Cancel = True
    End If
End If
End Sub
Private Sub xname_Change()
myload
End Sub

