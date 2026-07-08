VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form GrItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«·„Ã„Ê⁄« "
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   9285
   Begin VB.TextBox XTEXT1 
      Alignment       =   1  'Right Justify
      Height          =   1050
      Left            =   90
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   6750
      Width           =   4470
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
      Left            =   45
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6255
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   1890
      TabIndex        =   2
      Top             =   5535
      Width           =   7305
      Begin VB.TextBox xname 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   6405
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   3060
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   585
         Width           =   570
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
         Left            =   6570
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   555
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   180
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
   Begin MSAdodcLib.Adodc DATA1 
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5205
      Left            =   90
      TabIndex        =   6
      Top             =   225
      Width           =   9120
      _cx             =   16087
      _cy             =   9181
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
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
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   3
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
      ScrollBars      =   2
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin Threed.SSCommand cmd_excel 
      Height          =   465
      Left            =   6750
      TabIndex        =   7
      Top             =   6795
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
      Picture         =   "GrItem.frx":0000
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin ComctlLib.ProgressBar prog1 
      Height          =   330
      Left            =   4680
      TabIndex        =   8
      Top             =   6885
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "GrItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public aPublic, bedit As Boolean
Dim nZeros As Integer
Dim GrdTable As New ADODB.Recordset
Dim CTABLE As String, CGROUP As String
Private Sub myloadOLD()
Dim cFilter As String
GrdTable.Requery
If Trim(xName.text) <> "" Then cFilter = MyParnAnd(xName.text, "DESCA")
If Trim(XSECTION.BoundText) <> "" Then cFilter = " [GROUP] = " & Val(XSECTION.BoundText)
GrdTable.Filter = cFilter
I = 0
grid1.Rows = 1
Do Until GrdTable.EOF
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = GrdTable!CODE
    grid1.TextMatrix(grid1.Rows - 1, 1) = GrdTable!DESCA & ""
    grid1.TextMatrix(grid1.Rows - 1, 2) = GrdTable!descaA & ""
    grid1.TextMatrix(grid1.Rows - 1, 3) = GrdTable!Group & ""
    GrdTable.MoveNext
Loop
grid1.AddItem ""
End Sub



Private Sub cmd_excel_Click()
    grid1.ColHidden(4) = True
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
    grid1.ColHidden(4) = True
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
With grid1
con.BeginTrans
If Trim(grid1.TextMatrix(Row, 0)) <> "" Then
    con.Execute "update FILE1_50 " & _
                " SET DESCA = " & addstring(grid1.TextMatrix(Row, 1)) & _
                ",DESCAA = " & addstring(grid1.TextMatrix(Row, 2)) & _
                ",[GROUP]= " & addstring(grid1.TextMatrix(Row, 3)) & _
                " WHERE code = " & MyParn(grid1.TextMatrix(Row, 0))
ElseIf Trim(grid1.TextMatrix(Row, 1)) <> "" Then
        grid1.TextMatrix(Row, 0) = myField("SELECT dbo.RetZero(max(file1_50.code),3,1) From file1_50", con) & ""
        con.Execute "INSERT INTO FILE1_50 ( CODE ,CODE_0, DESCA , DESCAA, [GROUP] )" & _
                    "VALUES( " & _
                    addstring(grid1.TextMatrix(Row, 0)) & "," & _
                    addstring(grid1.TextMatrix(Row, 0)) & "," & _
                    addstring(grid1.TextMatrix(Row, 1)) & "," & _
                    addstring(grid1.TextMatrix(Row, 2)) & "," & _
                    addstring(grid1.TextMatrix(Row, 3)) & _
                    ")"
End If
End With
con.CommitTrans
myload
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
            con.Execute "Delete from FILE1_50 where CODE = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
            con.CommitTrans
            grid1.RemoveItem grid1.Row
        End If
    End If
End If
End Sub
Private Sub Form_Load()
con.CursorLocation = adUseClient
con.Open strCon

XTEXT1.Visible = bSupermode Or lSupperVisor

con.Execute " UPDATE FILE1_50 SET CODE_0 = CODE WHERE CODE_0 IS NULL "
con.Execute " UPDATE FILE1_50 SET CODE_0 = '0' + CODE WHERE LEN(CODE) = 2 "
con.Execute " UPDATE FILE1_50 SET CODE_0 = '00' + CODE WHERE LEN(CODE) = 1 "

Set DATA2.Recordset = myRecordSet("SELECT * FROM file1_10sc ORDER BY desca ", con)
Set XSECTION.RowSource = DATA2
XSECTION.ListField = "DESCA"
XSECTION.BoundColumn = "CODE"

Set grid1.DataSource = data1

'cString = "SELECT * from FILE1_50 order by code"
'GrdTable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
myload
FIXGRID
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
Private Sub xSection_LostFocus()
    myload
End Sub
Private Sub myload()
Dim cString As String
cString = " SELECT FILE1_50.CODE , FILE1_50.DESCA , FILE1_50.DESCAA , FILE1_50.[GROUP] , file1_10sc.desca FROM FILE1_50 left join file1_10sc on file1_10sc.code = file1_50.[group] WHERE FILE1_50.CODE IS NOT NULL  "
If Trim(xName.text) <> "" Then cString = cString & " AND " & MyParnAnd(xName.text, "DESCA")
If Trim(XSECTION.BoundText) <> "" Then cString = cString & " AND  [GROUP] = " & Val(XSECTION.BoundText)
cString = cString & " ORDER BY FILE1_50.CODE_0 "
Set data1.Recordset = myRecordSet(cString, con)
grid1.AddItem ""
FIXGRID
XTEXT1.text = cString
End Sub
Sub FIXGRID()
With grid1
    .Cols = 5
    .ColWidth(0) = 800
    .ColWidth(1) = 2500
    .ColWidth(2) = 2500
    .ColWidth(3) = 2000
    .ColWidth(4) = 0
    .ColComboList(3) = StrList("SELECT CODE , DESCA FROM FILE1_10SC ORDER BY DESCA ", con)
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "NAME"
    .TextMatrix(0, 2) = "«·»Ì«‰"
    .TextMatrix(0, 3) = "«·ﬁ”„"
    .ColHidden(4) = True
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .ColAlignment(2) = flexAlignRightCenter
    .ColAlignment(3) = flexAlignRightCenter
End With
End Sub

Private Sub XTEXT1_GotFocus()
XTEXT1.SelStart = 0
XTEXT1.SelLength = Len(XTEXT1.text)
End Sub
