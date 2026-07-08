VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form StoreData_BR 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«· ÊﬂÌ·« "
   ClientHeight    =   3345
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   7185
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   7185
   Begin VB.TextBox xipsql 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   855
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   1890
      Width           =   4650
   End
   Begin VB.CheckBox xisstop 
      Alignment       =   1  'Right Justify
      Caption         =   "›—⁄  „ Êﬁ›"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   855
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2340
      Width           =   1410
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   7185
      TabIndex        =   5
      Top             =   2880
      Width           =   7185
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   4425
         Picture         =   "StoreData_BR.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   3510
         Picture         =   "StoreData_BR.frx":25C0
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   2355
         Picture         =   "StoreData_BR.frx":4B93
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   1485
         Picture         =   "StoreData_BR.frx":726D
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   45
         Width           =   870
      End
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4500
      MaxLength       =   3
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   585
      Width           =   1005
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   855
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1000
      Width           =   4650
   End
   Begin VB.PictureBox SSPanel2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   7185
      TabIndex        =   2
      Top             =   0
      Width           =   7185
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   180
         MaskColor       =   &H00FFFFFF&
         Picture         =   "StoreData_BR.frx":994C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Picture         =   "StoreData_BR.frx":9A96
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Õ›Ÿ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         Height          =   420
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         Picture         =   "StoreData_BR.frx":9ED8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   1320
         MaskColor       =   &H00FFFFFF&
         Picture         =   "StoreData_BR.frx":C772
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   420
         Left            =   4740
         MaskColor       =   &H00FFFFFF&
         Picture         =   "StoreData_BR.frx":ECEB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   5880
         Picture         =   "StoreData_BR.frx":11297
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   45
         Width           =   1140
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2190
      _ExtentX        =   3863
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
   Begin MSDataListLib.DataCombo XBRANCH_GR 
      Height          =   390
      Left            =   855
      TabIndex        =   17
      Top             =   1415
      Width           =   4650
      _ExtentX        =   8202
      _ExtentY        =   688
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   855
      Top             =   585
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   2070
      Top             =   585
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
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«· IP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1935
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«·„‰ÿﬁ… : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   1530
      Width           =   750
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«· ÊﬂÌ· :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   645
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   5625
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   660
      Width           =   315
   End
End
Attribute VB_Name = "StoreData_BR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer, bedit As Boolean
Dim con As New ADODB.Connection
Dim formMode As Byte, cTableName As String, cGroupname As String
Dim CardTable As New ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub Form_Load()
openCon con
bedit = True
CmdDel.Visible = False


DATA3.ConnectionString = strCon
DATA3.RecordSource = "Select * From BRANCH_GR order by CODE "
Set XBRANCH_GR.RowSource = DATA3
XBRANCH_GR.ListField = "Desca"
XBRANCH_GR.BoundColumn = "Code"


CardTable.Open "SELECT * FROM BRANCH_FR  ORDER BY CODE", con, adOpenStatic, adLockReadOnly, adCmdText
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    mydefine
End If
End Sub
Private Sub CmdAdd_Click()
mydefine
xCode.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.BeginTrans
    con.Execute "Delete  From BRANCH_FR Where code = " & MyParn(xCode.text)
    con.CommitTrans
    CardTable.Requery
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(xCode.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        mydefine
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub

Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
CardTable.Requery
If xCode.Enabled Then
    CmdAdd_Click
Else
    CardTable.Find "code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
End If
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    mydefine
Else
    If xCode.Enabled Then
        mydefine
    Else
        CardTable.Find "code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub CmdFirst_Click()
    CardTable.MoveFirst
    myload
End Sub
Private Sub CmdInform_Click()
    CardLookup
End Sub
Private Sub CmdLast_Click()
    CardTable.MoveLast
    myload
End Sub
Private Sub CmdNext_Click()
    CardTable.MoveNext
    If CardTable.EOF Then
        CardTable.MovePrevious
    Else
        myload
    End If
End Sub
Private Sub CmdPrevious_Click()
    CardTable.MovePrevious
    If CardTable.BOF Then
        CardTable.MoveNext
    Else
        myload
    End If
End Sub
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode)
CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xCode.Enabled = Not (nMode = LoadMode)
CmdSave.Enabled = bedit
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From BRANCH_FR "
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Sub mydefine()
    xCode.text = RetZero(Val(Newflag("BRANCH_FR", "CODE", con)), 2)
    xDescA.text = ""
    xipsql.text = ""
    XBRANCH_GR.BoundText = ""
    xisstop.Value = 0
    Handlecontrols DefineMode
End Sub
Sub myload()
xCode.text = CardTable!CODE & ""
xDescA.text = CardTable!DESCA & ""
xipsql.text = CardTable!IPsql & ""
xisstop.Value = IIf(CardTable!ISSTOP, 1, 0)
XBRANCH_GR.BoundText = CardTable!branch_GR & ""
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert(4, 1)

aInsert(0, 0) = "CODE"
aInsert(0, 1) = addstring(xCode.text)

aInsert(1, 0) = "DESCA"
aInsert(1, 1) = addstring(xDescA.text)

aInsert(2, 0) = "ISSTOP"
aInsert(2, 1) = IIf(xisstop.Value, 1, 0)

aInsert(3, 0) = "BRANCH_GR"
aInsert(3, 1) = addvalue(XBRANCH_GR.BoundText)

aInsert(4, 0) = "IPSQL"
aInsert(4, 1) = addstring(xipsql.text)

On Error GoTo myerror
con.BeginTrans
If xCode.Enabled Then
    con.Execute CreateInsert(aInsert, "BRANCH_FR")
Else
    con.Execute CreateUpdate(aInsert, "BRANCH_FR", " where CODE = " & addstring(xCode.text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
   CardTable.Find "CODE = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
   myload
   Search3.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Unload Search3
Set Search3 = Nothing
Err.Clear
closeCon con
End Sub
Private Sub xCode_LostFocus()
If xCode.text = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 2)
CardTable.Find "CODE = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xDescA.text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

MYVALID = True
End Function
Private Sub xCode_GotFocus()
    xCode.SelStart = 0
    xCode.SelLength = Len(xCode.text)
End Sub
Private Sub xDescA_GotFocus()
    xDescA.SelStart = 0
    xDescA.SelLength = Len(xDescA.text)
End Sub
