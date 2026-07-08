VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ItemsGrdFrm 
   Caption         =   "»Ì«‰«  «·«’‰«›"
   ClientHeight    =   10530
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10530
   ScaleWidth      =   15240
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8715
      Left            =   135
      TabIndex        =   0
      Top             =   45
      Width           =   15045
      _cx             =   26538
      _cy             =   15372
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   13040125
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
      RowHeightMin    =   400
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
      Height          =   1500
      Left            =   11700
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   8775
      Width           =   2625
      Begin VB.CommandButton Command4 
         Caption         =   " ⁄œÌ· «·„Ã„Ê⁄… «·—∆Ì”Ì…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1035
         Width           =   2445
      End
      Begin VB.CommandButton Command3 
         Caption         =   " ⁄œÌ· «·„Ã„Ê⁄… "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   2445
      End
      Begin VB.CommandButton Command1 
         Caption         =   " ⁄œÌ· «·√ﬁ”«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   2445
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
      Height          =   420
      Left            =   135
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   8865
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   1620
      TabIndex        =   6
      Top             =   8775
      Width           =   7980
      Begin VB.TextBox XITEM 
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4185
         TabIndex        =   3
         Top             =   630
         Width           =   2445
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   630
         Width           =   2445
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   4185
         TabIndex        =   2
         Top             =   225
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬂÊœ :"
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
         Left            =   6795
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… «·’‰› :"
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
         Left            =   2700
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ”„ :"
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
         Left            =   6780
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   615
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   630
         Width           =   555
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
Attribute VB_Name = "itemsgrdFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public aPublic, bEdit As Boolean
Dim clist1 As String, cList2 As String, cList3 As String
Dim CardTable As New ADODB.Recordset
Private Sub MyLoad()
Dim cFilter As String
On Error GoTo myerror
cString = "SELECT ITEM as [«·ﬂÊœ], DESCA as [«·»Ì«‰],[GROUP] as [«·„Ã„Ê⁄…],[SECTION] as [«·ﬁ”„],SUPLER as [«·„Ê—œ],REORDER AS [Õœ «·ÿ·»],SHELF AS [«·—›],PRICE AS [”⁄— «·Ã„·…],PRICE2 AS [”⁄— «·„” Â·ﬂ],COST AS [«· ﬂ·›…],PACKAGE AS [«·⁄»Ê…],ROW,ITEM AS ITEMCODE" & _
          " FROM FILE1_10"

If IsNumeric(XGROUP.BoundText) Then
    cString = cString & turnFound(cString) & " [GROUP] = " & XGROUP.BoundText
End If

If IsNumeric(xSection.BoundText) Then
    cString = cString & turnFound(cString) & " [SECTION] = " & xSection.BoundText
End If

If Trim(xdesca.Text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(xdesca.Text, "desca")
End If

If Trim(XITEM.Text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(XITEM.Text, "ITEM")
End If

If Trim(xdesca.Text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(xdesca.Text, "desca")
End If

cString = cString & " order by Row"
data1.RecordSource = cString
data1.Refresh
grid1.AddItem ""
If XGROUP.BoundText <> "" And grid1.Rows = 2 Then
   grid1.TextMatrix(grid1.Rows - 1, 2) = XGROUP.BoundText
End If
If xSection.BoundText <> "" And grid1.Rows = 2 Then
   grid1.TextMatrix(grid1.Rows - 1, 3) = xSection.BoundText
End If
FixGrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub Command1_Click()
ReDim aPublic(5)
aPublic(0) = "FILE1_10SC"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "«·ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "√ﬁ”«„ «·«’‰«›"
FlagFrm.bEdit = bEdit
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
cList2 = StrList("select * from file1_10SC order by desca")
grid1.ColComboList(3) = cList2
data3.Refresh
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
itemsGroupFrm.bEdit = True
itemsGroupFrm.Show 1
clist1 = StrList("select * from file1_50 order by desca")
grid1.ColComboList(2) = clist1
data2.Refresh
End Sub

Private Sub Command4_Click()
ReDim aPublic(5)
aPublic(0) = "FILE1_50G"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "«·ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "«·„Ã„Ê⁄… «·—∆Ì”Ì…"
FlagFrm.bEdit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
SetKbLayout Lang_AR
DefAdd Me.Name, "GROUP", XGROUP.BoundText
DefAdd Me.Name, "SECTION", xSection.BoundText
CardTable.Close
Set CardTable = Nothing
Set itemsgrdFrm = Nothing
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
con.BeginTrans
If Trim(grid1.TextMatrix(Row, grid1.Cols - 1)) = "" Then
    CardTable.AddNew
    CardTable!Item = grid1.TextMatrix(Row, 0)
Else
    CardTable.Seek grid1.TextMatrix(Row, grid1.Cols - 1), adSeekFirstEQ
    If CardTable.EOF Then
        MsgBox "·„ Ì „ «· ⁄œÌ· !! «·”Ã· „Õ–Ê› «Ê „⁄·ﬁ"
        GoTo myerror
    End If
End If
If Col = 7 And grid1.TextMatrix(Row, 8) = "" Then
    grid1.TextMatrix(Row, 8) = Val(grid1.TextMatrix(Row, 7)) * 1.3
    If Val(grid1.TextMatrix(Row, 8)) - Int(Val(grid1.TextMatrix(Row, 8))) > 0 Then
        If Val(grid1.TextMatrix(Row, 8)) = Int(Val(grid1.TextMatrix(Row, 8))) <= 0.5 Then
            grid1.TextMatrix(Row, 8) = Int(Val(grid1.TextMatrix(Row, 8))) + 0.5
        Else
            grid1.TextMatrix(Row, 8) = Int(Val(grid1.TextMatrix(Row, 8))) + 1
        End If
    End If
End If
CardTable!DESCA = TurnValue(grid1.TextMatrix(Row, 1))
CardTable!Group = TurnValue(grid1.TextMatrix(Row, 2))
CardTable!Section = TurnValue(grid1.TextMatrix(Row, 3))
CardTable!SUPLER = TurnValue(grid1.TextMatrix(Row, 4))
CardTable!reOrder = TurnValue(grid1.TextMatrix(Row, 5))
CardTable!Shelf = TurnValue(grid1.TextMatrix(Row, 6))
CardTable!price = TurnValue(grid1.TextMatrix(Row, 7))
CardTable!price2 = TurnValue(grid1.TextMatrix(Row, 8))
CardTable!cost = TurnValue(grid1.TextMatrix(Row, 9))
CardTable!package = TurnValue(grid1.TextMatrix(Row, 10))
CardTable!Row = Row
CardTable.Update
con.CommitTrans
grid1.TextMatrix(Row, grid1.Cols - 1) = grid1.TextMatrix(Row, 0)
If Col = 0 Then grid1_EnterCell
Exit Sub
myerror:
MsgBox Err.Description
On Error Resume Next
con.RollbackTrans
CardTable.CancelUpdate
Err.Clear
MyLoad
End Sub
Private Sub grid1_EnterCell()
If (grid1.Col = 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "") Or (Not bEdit) Or (grid1.Col <> 0 And grid1.TextMatrix(grid1.Row, 0) = "") Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
    SetKbLayout IIf(grid1.Col = 0, Lang_EN, Lang_AR)
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo myerror
If KeyCode = 46 Then
    If Trim(grid1.TextMatrix(grid1.Row, 0)) <> "" Then
        If MsgBox("Remove Row ?? Are you sure", vbYesNo) = vbYes Then
            CardTable.Seek grid1.TextMatrix(grid1.Row, grid1.Cols - 1), adSeekFirstEQ
            If Not CardTable.EOF Then
                CardTable.Delete
                grid1.RemoveItem grid1.Row
                grid1_EnterCell
            Else
                GoTo myerror
            End If
        End If
    End If
End If
Exit Sub
myerror:
If Err.Number <> 0 Then MsgBox Err.Description
CardTable.CancelUpdate
con.RollbackTrans
MyLoad
End Sub
Private Sub Form_Load()
data2.ConnectionString = strCon
data2.RecordSource = "FILE1_50"
Set XGROUP.RowSource = data2
XGROUP.ListField = "Desca"
XGROUP.BoundColumn = "Code"

data3.ConnectionString = strCon
data3.RecordSource = "FILE1_10SC"
Set xSection.RowSource = data3
xSection.ListField = "Desca"
xSection.BoundColumn = "Code"

XGROUP.BoundText = DefGet(Me.Name, "GROUP")
xSection.BoundText = DefGet(Me.Name, "SECTION")

Set grid1.DataSource = data1
data1.ConnectionString = strCon
CardTable.CursorLocation = adUseServer
CardTable.Index = "NDXITEM"
CardTable.Open "FILE1_10", con, adOpenKeyset, adLockOptimistic, adCmdTableDirect
With grid1
clist1 = StrList("Select code,desca from file1_50 order by desca")
cList2 = StrList("Select code,desca from file1_10sc order by desca")
cList3 = StrList("Select code,desca from FILE4_10 order by desca")
MyLoad
End With
End Sub
Private Sub grid1_LostFocus()
SetKbLayout Lang_AR
End Sub

Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then
    grid1.AddItem ""
    If Row <> 2 Then
        grid1.TextMatrix(Row, 2) = grid1.TextMatrix(Row - 1, 2)
        grid1.TextMatrix(Row, 3) = grid1.TextMatrix(Row - 1, 3)
    Else
        If XGROUP.BoundText <> "" Then
           grid1.TextMatrix(Row, 2) = XGROUP.BoundText
        End If
        If xSection.BoundText <> "" Then
           grid1.TextMatrix(Row, 3) = xSection.BoundText
        End If
    End If
End If
End Sub
Private Sub xcountry_code_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then MyLoad
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "ﬂÊœ «·’‰› „ÿ·Ê»"
        Cancel = True
        Exit Sub
    End If
End If
If Col = 1 Then
    If Trim(grid1.EditText) = "" Then
        MsgBox "Ê’› «·’‰› „ÿ·Ê»"
        Cancel = True
    End If
End If
On Error GoTo myerror
    If Col = 2 Then
        If Trim(GetDesca("Select Desca From FILE1_50 where Desca = " & MyParn(grid1.EditText))) = "" Then
            If MsgBox("«÷«›… „Ã„Ê⁄… ÃœÌœ…", vbYesNo + vbDefaultButton1) = vbYes Then
                con.BeginTrans
                On Error GoTo myerror
                nCode = Newflag("FILE1_50", "code")
                con.Execute "insert into FILE1_50(CODE,Desca) " & _
                "values(" & _
                addvalue(nCode) & "," & _
                addstring(grid1.EditText) & _
                ")"
                con.CommitTrans
                grid1.EditText = nCode
                clist1 = StrList("select * from file1_50 order by desca")
                grid1.ColComboList(2) = clist1
                data2.Refresh
            Else
                Cancel = True
            End If
        End If
    End If

    If Col = 3 Then
        If Trim(GetDesca("Select Desca From FILE1_10SC where Desca = " & MyParn(grid1.EditText))) = "" Then
            If MsgBox("«÷«›… ﬁ”„ ÃœÌœ", vbYesNo + vbDefaultButton1) = vbYes Then
                con.BeginTrans
                On Error GoTo myerror
                nCode = Newflag("FILE1_10SC", "code")
                con.Execute "insert into FILE1_10SC(CODE,Desca) " & _
                "values(" & _
                addvalue(nCode) & "," & _
                addstring(grid1.EditText) & _
                ")"
                con.CommitTrans
                grid1.EditText = nCode
                cList2 = StrList("select * from file1_10SC order by desca")
                grid1.ColComboList(3) = cList2
                data3.Refresh
            Else
                Cancel = True
            End If
        End If
    End If

Exit Sub
myerror:
On Error Resume Next
If Err.Number <> 0 Then MsgBox Err.Description
CardTable.CancelUpdate
con.RollbackTrans
MyLoad
Err.Clear
End Sub
Private Sub FixGrd()
With grid1
.ColComboList(2) = clist1
.ColComboList(3) = cList2
.ColComboList(4) = cList3
.ColWidth(0) = 2000
.ColWidth(1) = 3300
.ColWidth(2) = 2100
.ColWidth(3) = 2100
.ColWidth(4) = 1500
.ColWidth(5) = 800
.ColWidth(6) = 1100
.ColWidth(7) = 1100
.ColWidth(8) = 1100
.ColWidth(9) = 800
.ColHidden(5) = True
.ColHidden(6) = True
.ColHidden(9) = True
.ColHidden(11) = True
.ColHidden(12) = True
.RowHeight(0) = 1000
.WordWrap = True
For i = 1 To grid1.Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub

Private Sub xDesca_Change()
MyLoad
End Sub
Private Sub xITEM_Change()
MyLoad
End Sub

Private Sub xGroup_Click(Area As Integer)
If Area = 2 Then MyLoad
End Sub
Private Sub xgroup_Validate(Cancel As Boolean)
MyLoad
End Sub

Private Sub XITEM_GotFocus()
SetKbLayout Lang_EN
End Sub

Private Sub xitem_LostFocus()
SetKbLayout Lang_AR
End Sub

Private Sub xSection_Click(Area As Integer)
If Area = 2 Then MyLoad
End Sub
Private Sub xSection_Validate(Cancel As Boolean)
MyLoad
End Sub
