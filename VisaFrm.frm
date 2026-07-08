VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form VisaFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·›Ì“«"
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10110
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7125
   ScaleWidth      =   10110
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   5490
      Visible         =   0   'False
      Width           =   4110
      Begin VB.OptionButton xstop 
         Alignment       =   1  'Right Justify
         Caption         =   "›Ï «·⁄„·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   1005
      End
      Begin VB.OptionButton xstop 
         Alignment       =   1  'Right Justify
         Caption         =   "„ Êﬁ›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   915
      End
      Begin VB.OptionButton xstop 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬂ·"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   645
      End
   End
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
      Height          =   555
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6480
      Width           =   1500
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   5355
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   9915
      _cx             =   17489
      _cy             =   9446
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
      Height          =   1005
      Left            =   4230
      TabIndex        =   3
      Top             =   5490
      Width           =   5775
      Begin VB.TextBox xname 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   4560
      End
      Begin MSDataListLib.DataCombo XBRANCH 
         Height          =   315
         Left            =   1215
         TabIndex        =   5
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "»Ì«‰ : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   825
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   555
      End
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
Attribute VB_Name = "VisaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public aPublic, bedit As Boolean
Dim nZeros As Integer
Dim GrdTable As New ADODB.Recordset
Dim CTABLE As String, CGROUP As String
Private Sub myload()
Dim cFilter As String
GrdTable.Requery
If Trim(xname.text) <> "" Then cFilter = MyParnAnd(xname.text, "desca")
If cFilter = "" Then
    If Trim(XBRANCH.BoundText) <> "" Then cFilter = " and BRANCH = " & MyParn(XBRANCH.BoundText)
Else
    If Trim(XBRANCH.BoundText) <> "" Then cFilter = cFilter & " AND BRANCH = " & MyParn(XBRANCH.BoundText)
End If

If cFilter <> "" Then cFilter = Mid(cFilter, 5)
GrdTable.Filter = cFilter

i = 0
Grid1.Rows = 1
Do Until GrdTable.EOF
    Grid1.AddItem ""
    Grid1.TextMatrix(Grid1.Rows - 1, 0) = GrdTable!CODE
    Grid1.TextMatrix(Grid1.Rows - 1, 1) = GrdTable!DESCA & ""
    Grid1.TextMatrix(Grid1.Rows - 1, 2) = GrdTable!branch & ""
    Grid1.TextMatrix(Grid1.Rows - 1, 3) = IIf(GrdTable!ISSTOP, -1, "")
    Grid1.TextMatrix(Grid1.Rows - 1, 4) = GrdTable!Rate & ""
    GrdTable.MoveNext
Loop
Grid1.AddItem ""
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
With Grid1
'con.BeginTrans
If Grid1.TextMatrix(Row, 0) <> "" Then
    con.Execute "update VISA" & _
                " SET DESCA = " & addstring(Grid1.TextMatrix(Row, 1)) & _
                " , BRANCH = " & addstring(Grid1.TextMatrix(Row, 2)) & _
                " , IsSTOP = " & IIf(Val(Grid1.TextMatrix(Row, 3)) <> 0, "1", "0") & _
                " WHERE code = " & MyParn(Grid1.TextMatrix(Row, 0))
Else
    If Trim(Grid1.TextMatrix(Row, 1)) <> "" Then
        Grid1.TextMatrix(Row, 0) = RetZero(Newflag("VISA", "code", con), 4)
        Grid1.TextMatrix(Row, 2) = XBRANCH.BoundText
        con.Execute "INSERT INTO VISA ( CODE , DESCA , BRANCH , ISSTOP )" & _
                    "VALUES( " & _
                    addstring(Grid1.TextMatrix(Row, 0)) & "," & _
                    addstring(Grid1.TextMatrix(Row, 1)) & "," & _
                    addstring(Grid1.TextMatrix(Row, 2)) & "," & _
                    IIf(Val(Grid1.TextMatrix(Row, 3)) <> 0, "1", "0") & _
                    ")"
    End If
End If
End With
'con.CommitTrans
Exit Sub
myerror:
MsgBox Err.Description
'con.RollbackTrans
Err.Clear
myload
End Sub
Private Sub grid1_EnterCell()
If Grid1.Row > Grid1.Rows - 1 Then
    Grid1.Editable = flexEDNone
Else
    Grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then
    If Trim(Grid1.TextMatrix(Grid1.Row, 0)) <> "" Then
        If MsgBox("Remove Row ?? Are you sure", vbYesNo) = vbYes Then
'            con.BeginTrans
            con.Execute "Delete from VISA where CODE = " & MyParn(Grid1.TextMatrix(Grid1.Row, 0))
'            con.CommitTrans
            Grid1.RemoveItem Grid1.Row
        End If
    End If
End If
End Sub
Private Sub Form_Load()
con.CursorLocation = adUseClient
con.Open strCon

'Fix_Data


cString = "SELECT  * from VISA  Order by code"
GrdTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText


DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select Code,DescA From BRANCH order by Desca"
Set XBRANCH.RowSource = DATA2
XBRANCH.ListField = "Desca"
XBRANCH.BoundColumn = "Code"

With Grid1

.Cols = 5
.ColWidth(0) = 800
.ColWidth(1) = 4000
.ColWidth(2) = 1500
.ColWidth(3) = 700
.ColWidth(4) = 1200
.ColHidden(4) = True
.RowHeight(0) = 1000
.WordWrap = True

.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·»Ì«‰"
.TextMatrix(0, 2) = "«·›—⁄"
.TextMatrix(0, 3) = "„ Êﬁ›"
.TextMatrix(0, 4) = "ﬂÊœ „ÊŸ›"

'.ColHidden(2) = (cBranch > "60")
'.ColHidden(4) = (cBranch > "60")
.ColAlignment(0) = flexAlignRightCenter
.ColAlignment(1) = flexAlignRightCenter
.ColDataType(3) = flexDTBoolean
.ColComboList(2) = StrListA("SELECT CODE , DESCA FROM BRANCH ", con)
myload
End With
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
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
Private Sub XBRANCH_Change()
    myload
End Sub
Private Sub xname_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then myload
End Sub
Private Sub xid_emp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then myload
End Sub

Private Sub xStore_Click(Area As Integer)
myload
End Sub
Private Sub xstop_Click(Index As Integer)
    myload
End Sub
