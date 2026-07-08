VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form sectionCode 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·√ﬁ”«„"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8370
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4785
   ScaleWidth      =   8370
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
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   8160
      _cx             =   14393
      _cy             =   6429
      _ConvInfo       =   1
      Appearance      =   1
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
      Cols            =   4
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
      Left            =   945
      TabIndex        =   3
      Top             =   3690
      Width           =   7305
      Begin VB.TextBox xname 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   4965
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
         Left            =   5310
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   555
      End
   End
End
Attribute VB_Name = "sectionCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public aPublic, bedit As Boolean
Dim nZeros As Integer
Dim GrdTable As New ADODB.Recordset
Dim CTABLE As String, cGroup As String
Private Sub myload()
Dim cFilter As String
GrdTable.Requery
If Trim(xname.text) <> "" Then cFilter = MyParnAnd(xname.text, "desca")
GrdTable.Filter = cFilter
i = 0
grid1.Rows = 1
Do Until GrdTable.EOF
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = GrdTable!CODE
    grid1.TextMatrix(grid1.Rows - 1, 1) = GrdTable!DESCA & ""
    grid1.TextMatrix(grid1.Rows - 1, 2) = GrdTable!BARCODE_GS1_SEC & ""
    grid1.TextMatrix(grid1.Rows - 1, 3) = IIf(GrdTable!ISDISC, 1, 0)
    GrdTable.MoveNext
Loop
grid1.AddItem ""
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GrdTable.Close
Set GrdTable = Nothing
Set FlagFrm = Nothing
    closeCon con
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'On Error GoTo myerror
With grid1
con.BeginTrans
If grid1.TextMatrix(Row, 0) <> "" Then
    con.Execute "update FILE1_10SC " & _
                "  SET DESCA = " & addstring(grid1.TextMatrix(Row, 1)) & _
                "  , BARCODE_GS1_SEC = " & addstring(grid1.TextMatrix(Row, 2)) & _
                ", ISDISC = " & Val(grid1.TextMatrix(Row, 3)) & _
                " WHERE code = " & MyParn(grid1.TextMatrix(Row, 0))
Else
    If Trim(grid1.TextMatrix(Row, 1)) <> "" Then
        NNEW = IncRec(GetDesca("Select max(cast(code as int)) from FILE1_10SC", con))
        grid1.TextMatrix(Row, 0) = RetZero(NNEW, 2)
        con.Execute "INSERT INTO FILE1_10SC ( CODE , DESCA  , BARCODE_GS1_SEC , ISDISC )" & _
                    "VALUES( " & _
                    addstring(grid1.TextMatrix(Row, 0)) & "," & _
                    addstring(grid1.TextMatrix(Row, 1)) & "," & _
                    addstring(grid1.TextMatrix(Row, 2)) & "," & _
                    IIf(Val(grid1.TextMatrix(Row, 2)) <> 0, "1", "0") & _
                    ")"
    End If
End If
End With
con.CommitTrans
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myload
End Sub
Private Sub grid1_EnterCell()
If grid1.Row > grid1.Rows - 1 Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Remove Row ?? Are you sure", vbYesNo) = vbYes Then
            con.BeginTrans
            con.Execute "Delete from FILE1_10SC  where CODE = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
            con.CommitTrans
            grid1.RemoveItem grid1.Row
        End If
    End If
End If
End Sub
Private Sub Form_Load()
con.CursorLocation = adUseClient
con.Open strCon

'Fix_Data

cString = "SELECT  *  from FILE1_10SC order by code"
GrdTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
With grid1

.Cols = 4
.ColWidth(0) = 800
.ColWidth(1) = 4000
.ColWidth(2) = 2500
.ColWidth(3) = 700
.RowHeight(0) = 1000
.WordWrap = True
.ColHidden(3) = True

.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·ﬁ”„"
.TextMatrix(0, 2) = "»«—ﬂÊœ "
.TextMatrix(0, 3) = "⁄—÷ 2+1"
.ColAlignment(0) = flexAlignRightCenter
.ColAlignment(1) = flexAlignRightCenter
.ColDataType(3) = flexDTBoolean
myload
End With
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then grid1.Rows = grid1.Rows + 1
End Sub
Private Sub xcountry_code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then myload
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "description required"
        Cancel = True
    End If
End If
End Sub
Private Sub xname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then myload
End Sub
