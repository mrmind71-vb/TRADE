VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form SubUser 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "’·«ÕÌ«  «·≈œ«—… ··„”ƒ·Ì‰ ›Ï «·›—Ê⁄"
   ClientHeight    =   3210
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
   ScaleHeight     =   3210
   ScaleWidth      =   7185
   Begin VB.CheckBox xIsReturn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Caption         =   "«·”„«Õ  »„— Ã⁄ »œÊ‰ ›« Ê—…"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2205
      Width           =   2310
   End
   Begin VB.CheckBox xstop2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      Caption         =   "«Ìﬁ«›"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1575
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   1800
      Width           =   960
   End
   Begin VB.CheckBox xstop1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      Caption         =   "«Ìﬁ«›"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1575
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   1440
      Width           =   960
   End
   Begin VB.TextBox XPASSWORD2 
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
      Left            =   2700
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1764
      Width           =   3390
   End
   Begin VB.TextBox XPASSWORD 
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
      Left            =   4320
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1371
      Width           =   1770
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
      Top             =   2745
      Width           =   7185
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   4425
         Picture         =   "SubUser.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   3510
         Picture         =   "SubUser.frx":25C0
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   2355
         Picture         =   "SubUser.frx":4B93
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   1485
         Picture         =   "SubUser.frx":726D
         Style           =   1  'Graphical
         TabIndex        =   14
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
      Left            =   5085
      MaxLength       =   4
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
      Left            =   1440
      MaxLength       =   50
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   978
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
         Picture         =   "SubUser.frx":994C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "SubUser.frx":9A96
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Õ›Ÿ"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         Height          =   420
         Left            =   2460
         MaskColor       =   &H00FFFFFF&
         Picture         =   "SubUser.frx":9ED8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Picture         =   "SubUser.frx":C772
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   " —«Ã⁄"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   420
         Left            =   4770
         MaskColor       =   &H00FFFFFF&
         Picture         =   "SubUser.frx":ECEB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   5880
         Picture         =   "SubUser.frx":11297
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   45
         Width           =   1140
      End
   End
   Begin MSDataListLib.DataCombo xbranch 
      Height          =   315
      Left            =   2700
      TabIndex        =   18
      Top             =   2160
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777088
      Text            =   ""
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ«—   :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6210
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   1845
      Width           =   510
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "«·›—⁄ :"
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
      Left            =   6255
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2205
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ·„… «·”— : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6210
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«·≈”„ : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6210
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂÊœ : "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6210
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   630
      Width           =   405
   End
End
Attribute VB_Name = "SubUser"
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
data1.ConnectionString = strCon
data1.RecordSource = "Select S_BRANCH ,DescA From STORE_BR order by code "
Set xbranch.RowSource = data1
xbranch.ListField = "Desca"
xbranch.BoundColumn = "S_BRANCH"

CardTable.Open "SELECT * FROM SUB_USER ORDER BY CODE", con, adOpenStatic, adLockReadOnly, adCmdText
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdAdd_Click()
myDefine
xCode.SetFocus
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    con.BeginTrans
    con.Execute "Delete  From SUB_USER  Where code = " & MyParn(xCode.text)
    con.CommitTrans
    CardTable.Requery
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(xCode.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        myDefine
    End If
End If
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
    con.RollbackTrans
End Sub

Private Sub CmdExit_Click()
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
    myDefine
Else
    If xCode.Enabled Then
        myDefine
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
cmdSave.Enabled = True
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select   SUB_USER.code ,SUB_USER.DescA , STORE_BR.DESCA , STOP1 , STOP2 , IsReturn From SUB_USER INNER JOIN STORE_BR ON STORE_BR.S_BRANCH = SUB_USER.BRANCH "
Generalarray(2) = "Order by SUB_USER.code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%SUB_USER.DESCA%%)"

listarray(1, 0) = "«·›—⁄"
listarray(1, 1) = "(%%STORE_BR.DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 5000

GrdArray(2, 0) = "«·›—⁄"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "Êﬁ› ﬂ·„…"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "Êﬁ› ﬂ«— "
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "„— Ã⁄ "
GrdArray(5, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ «·Œ“‰"
Search3.Show 1
End Sub
Sub myDefine()

xCode.text = RetZero(Val(Newflag("SUB_USER", "CODE", con)), 4)
xDesca.text = ""
xPassword.text = ""
XPASSWORD2.text = ""
XSTOP1.Value = 0
XSTOP2.Value = 0
xbranch.BoundText = ""
xIsReturn.Value = 0
Handlecontrols DefineMode
End Sub
Sub myload()
xCode.text = CardTable!CODE & ""
xDesca.text = CardTable!DESCA & ""
xPassword.text = CardTable!PassWord & ""
XPASSWORD2.text = CardTable!PassWord2 & ""
XSTOP1.Value = IIf(CardTable!Stop1, 1, 0)
XSTOP2.Value = IIf(CardTable!Stop2, 1, 0)
xIsReturn.Value = IIf(CardTable!IsReturn, 1, 0)
xbranch.BoundText = CardTable!branch & ""
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert(7, 1)
aInsert(0, 0) = "CODE"
aInsert(0, 1) = addstring(xCode.text)

aInsert(1, 0) = "DESCA"
aInsert(1, 1) = addstring(xDesca.text)

aInsert(2, 0) = "PASSWORD"
aInsert(2, 1) = addstring(xPassword.text)

aInsert(3, 0) = "PASSWORD2"
aInsert(3, 1) = addstring(XPASSWORD2.text)

aInsert(4, 0) = "BRANCH"
aInsert(4, 1) = addstring(xbranch.BoundText)

aInsert(5, 0) = "STOP1"
aInsert(5, 1) = IIf(XSTOP1.Value, 1, 0)

aInsert(6, 0) = "STOP2"
aInsert(6, 1) = IIf(XSTOP2.Value, 1, 0)

aInsert(7, 0) = "IsReturn"
aInsert(7, 1) = IIf(xIsReturn.Value, 1, 0)

On Error GoTo myerror
con.BeginTrans
If xCode.Enabled Then
    con.Execute CreateInsert(aInsert, "SUB_USER")
Else
    con.Execute CreateUpdate(aInsert, "SUB_USER", " where CODE = " & addstring(xCode.text))
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
xCode.text = xCode.text
CardTable.Find "CODE = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xDesca.text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

MYVALID = True
End Function

Private Sub xf_Date_GotFocus()
xF_date.SelStart = 0
xF_date.SelLength = Len(xF_date.text)
End Sub
Private Sub xf_BAL_GotFocus()
xF_Bal.SelStart = 0
xF_Bal.SelLength = Len(xF_Bal.text)
End Sub
Private Sub xCode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.text)
End Sub
Private Sub xDescA_GotFocus()
xDesca.SelStart = 0
xDesca.SelLength = Len(xDesca.text)
End Sub

