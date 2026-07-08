VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form chargecodefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«ﬂÊ«œ «·„’«—Ì› "
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   7425
   Begin VB.TextBox XTEXT1 
      Alignment       =   1  'Right Justify
      Height          =   1050
      Left            =   45
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   5220
      Width           =   4470
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   90
      TabIndex        =   2
      Top             =   3735
      Width           =   5910
      Begin VB.TextBox xname 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   3615
      End
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   360
         Left            =   315
         TabIndex        =   4
         Top             =   540
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄… :"
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
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "«·Ê’› :"
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
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdExit 
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
      Top             =   4770
      Width           =   1455
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   3690
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   7215
      _cx             =   12726
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
      Cols            =   3
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1905
      _ExtentX        =   3360
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
   Begin Threed.SSCommand cmd_excel 
      Height          =   465
      Left            =   3645
      TabIndex        =   7
      Top             =   4725
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   13822956
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "chargecode.frx":0000
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin ComctlLib.ProgressBar prog1 
      Height          =   330
      Left            =   1575
      TabIndex        =   8
      Top             =   4815
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "chargecodefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection, cSub1 As String
Public myPublic As Integer, bedit As Boolean
Dim GrdTable As New ADODB.Recordset
Dim CTABLE As String, CGROUP As String
Private Sub cmd_excel_Click()
    grid1.ColHidden(4) = True
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
    grid1.ColHidden(4) = False
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Command2_Click()
myload
End Sub
Private Sub myload()
Dim cFilter As String
GrdTable.Requery
If Trim(xName.text) <> "" Then cFilter = MyParnAnd(xName.text, "desca")
If Trim(xGroup.BoundText) <> "" Then cFilter = cFilter & turnFound2(cFilter, "") & " MainGroup = " & MyParn(xGroup.BoundText)
GrdTable.Filter = cFilter
i = 0
grid1.Rows = 1
Do Until GrdTable.EOF
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = GrdTable!CODE & ""
    grid1.TextMatrix(grid1.Rows - 1, 1) = GrdTable!DESCA & ""
    grid1.TextMatrix(grid1.Rows - 1, 2) = GrdTable!MainGroup & ""
    grid1.TextMatrix(grid1.Rows - 1, 3) = GrdTable!ISasset & ""
    If myPublic = 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 4) = GetDesca("SELECT DESCA FROM FILE8_52 WHERE CODE = " & MyParn(GrdTable!MainGroup), con)
    Else
        grid1.TextMatrix(grid1.Rows - 1, 4) = GetDesca("SELECT DESCA FROM FILE8_62 WHERE CODE = " & MyParn(GrdTable!MainGroup), con)
    End If
    GrdTable.MoveNext
Loop

grid1.AddItem ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
GrdTable.Close
Set GrdTable = Nothing
Set chargecodefrm = Nothing
closeCon con
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error Resume Next
If Not bedit Then Exit Sub
con.BeginTrans
If grid1.TextMatrix(Row, 0) <> "" Then
    con.Execute "update " & CTABLE & _
                " set DESCA = " & addstring(grid1.TextMatrix(Row, 1)) & "," & _
                " [MainGroup] = " & addstring(grid1.TextMatrix(Row, 2)) & "," & _
                " [ISasset] = " & IIf(grid1.TextMatrix(Row, 3), 1, 0) & _
                " WHERE CODE = " & MyParn(grid1.TextMatrix(Row, 0))
Else
    If Trim(grid1.TextMatrix(Row, 1)) <> "" Then
        grid1.TextMatrix(Row, 0) = RetZero(Newflag(CTABLE, "code", con), 3)
        For nTry = 1 To 10
            con.Execute "INSERT INTO " & CTABLE & "(CODE,DESCA,ISasset,MainGroup) " & _
                        "VALUES( " & _
                        addstring(grid1.TextMatrix(Row, 0)) & "," & _
                        addstring(grid1.TextMatrix(Row, 1)) & "," & _
                        IIf(grid1.TextMatrix(Row, 3), 1, 0) & "," & _
                        addstring(grid1.TextMatrix(Row, 2)) & _
                        ")"
            
            If Err.Number = -2147467259 And nTry < 10 Then
                Err.Clear
                grid1.TextMatrix(Row, 0) = RetZero(Val(grid1.TextMatrix(Row, 0)) + 1, 3)
            End If
            If Err.Number = 0 Then Exit For
            If Err.Number <> 0 Then GoTo myerror
        Next
    End If
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
If grid1.Row > grid1.Rows - 1 Or Not bedit Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And bedit Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Remove Row ?? Are you sure", vbYesNo) = vbYes Then
        If GetDesca("SELECT CHARGE FROM FILE8_50 WHERE CHARGE = " & MyParn(grid1.TextMatrix(grid1.Row, 0)), con) = "" Then
'                con.BeginTrans
                con.Execute "Delete  from " & CTABLE & " where code = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
'                con.CommitTrans
                grid1.RemoveItem grid1.Row
                grid1_EnterCell
        Else
            MsgBox "ÌÊÃœ Õ—ﬂ… ⁄·Ï «·»‰œ "
        End If
        End If
    End If
End If
End Sub
Private Sub Form_Load()

XTEXT1.Visible = bSupermode Or lSupperVisor

openCon con
If myPublic = 1 Then
    cSub1 = "file8_50"
    CTABLE = "file8_51"
    CGROUP = "file8_52"
    Me.Caption = "«ﬂÊ«œ «·„’«—Ì›"
Else
    cSub1 = "file8_60"
    CTABLE = "file8_61"
    CGROUP = "file8_62"
    Me.Caption = "«ﬂÊ«œ «·«Ì—«œ« "
End If

data1.ConnectionString = strCon
data1.RecordSource = CGROUP
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
If myPublic = 1 Then
    cString = "SELECT FILE8_51.code,FILE8_51.desca,FILE8_51.MainGroup ,  FILE8_51.ISasset ,  FILE8_52.DESCA from FILE8_51 INNER JOIN FILE8_52 ON FILE8_51.MainGroup = FILE8_52.CODE order by FILE8_51.CODE "
Else
    cString = "SELECT FILE8_61.code,FILE8_61.desca,FILE8_61.MainGroup ,  FILE8_61.ISasset ,  FILE8_62.DESCA from FILE8_61 INNER JOIN FILE8_62 ON FILE8_61.MainGroup = FILE8_62.CODE order by FILE8_61.CODE "
End If
GrdTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
XTEXT1.text = cString
With grid1
.Cols = 5
.ColWidth(0) = 1000
.ColWidth(1) = 2000
.ColWidth(2) = 2000
.ColWidth(3) = 1000
.TextMatrix(0, 0) = "«·ﬂÊœ"
.TextMatrix(0, 1) = "«·»Ì«‰"
.TextMatrix(0, 2) = "«·„Ã„Ê⁄…"
.TextMatrix(0, 3) = "«’Ê·"
.TextMatrix(0, 4) = "«·„Ã„Ê⁄…"

.ColAlignment(1) = flexAlignRightCenter
.ColAlignment(2) = flexAlignRightCenter
.ColComboList(2) = StrList_2
.ColDataType(3) = flexDTBoolean
.ColHidden(4) = True
myload
End With
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'If Grid1.Row = Grid1.Rows - 20 Then
If grid1.Row = grid1.Rows - 1 Then grid1.Rows = grid1.Rows + 1
End Sub
Private Sub xcountry_code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then myload
End Sub
Private Sub xcountry_code_LostFocus()
xCountry.text = ""
If Trim(xCountry.text) <> "" Then xCountry.text = GetDesca("select country_name from country where country_code = " & Val(xcountry_code.text), con)

myload
End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 1 Then
    If grid1.TextMatrix(Row, 3) = "" Then grid1.TextMatrix(Row, 3) = 0
    If Trim(grid1.EditText) = "" Then
        MsgBox "description required"
        Cancel = True
    End If
End If
End Sub

Private Sub xGroup_Click(Area As Integer)
If Area = 2 Then myload
End Sub

Private Sub xname_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then myload
End Sub
Private Function StrList_2()
Dim listTable As New ADODB.Recordset
listTable.Open "SELECT * FROM " & CGROUP & " ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
StrList_2 = "#  " & ";       "
Do Until listTable.EOF
    StrList_2 = StrList_2 & "|#" & listTable!CODE & ";" & listTable!DESCA
    listTable.MoveNext
Loop
End Function

Private Sub XTEXT1_GotFocus()
XTEXT1.SelStart = 0
XTEXT1.SelLength = Len(XTEXT1.text)
End Sub

