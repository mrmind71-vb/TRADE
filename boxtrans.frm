VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form boxtransfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕÊÌ·«  ‰ﬁœÌ… „‰ Œ“‰… ≈·Ï  √Œ—Ï"
   ClientHeight    =   3240
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   7920
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
   ScaleHeight     =   3240
   ScaleWidth      =   7920
   Begin VB.Frame FRM_CLOSED 
      Height          =   1095
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1665
      Width           =   2085
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·„” ‰œ „€·ﬁ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   135
         Visible         =   0   'False
         Width           =   1410
      End
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         TabIndex        =   28
         Top             =   405
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   1058
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "boxtrans.frx":0000
         Alignment       =   4
         PictureAlignment=   9
      End
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   5085
      MaxLength       =   6
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   585
      Width           =   1365
   End
   Begin VB.TextBox xDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   5085
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   945
      Width           =   1365
   End
   Begin VB.TextBox xValue 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   4905
      MaxLength       =   10
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2385
      Width           =   1545
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
      ScaleWidth      =   7920
      TabIndex        =   11
      Top             =   2775
      Width           =   7920
      Begin VB.CommandButton Command1 
         Caption         =   "fix doc"
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   90
         Width           =   1005
      End
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   4425
         Picture         =   "boxtrans.frx":25CC
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   3510
         Picture         =   "boxtrans.frx":4B8C
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   2355
         Picture         =   "boxtrans.frx":715F
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   45
         Width           =   915
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   1485
         Picture         =   "boxtrans.frx":9839
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   45
         Width           =   870
      End
      Begin VB.Label xRecordNumber 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   45
         TabIndex        =   15
         Top             =   150
         Width           =   1890
      End
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   2205
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2025
      Width           =   4245
   End
   Begin VB.PictureBox SSPanel2 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
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
      ScaleWidth      =   7920
      TabIndex        =   7
      Top             =   0
      Width           =   7920
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   5760
         Picture         =   "boxtrans.frx":BF18
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   45
         Width           =   1140
      End
      Begin VB.CommandButton CmdAdd 
         Height          =   420
         Left            =   4605
         MaskColor       =   &H00FFFFFF&
         Picture         =   "boxtrans.frx":E6EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   1185
         MaskColor       =   &H00FFFFFF&
         Picture         =   "boxtrans.frx":10C97
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdDel 
         Height          =   420
         Left            =   2325
         MaskColor       =   &H00FFFFFF&
         Picture         =   "boxtrans.frx":13210
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   3465
         MaskColor       =   &H00FFFFFF&
         Picture         =   "boxtrans.frx":15AAA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "boxtrans.frx":15EEC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   1140
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2565
      Top             =   630
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
   Begin MSDataListLib.DataCombo xNo1 
      Height          =   315
      Left            =   3555
      TabIndex        =   2
      Top             =   1305
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DataCombo1"
      RightToLeft     =   -1  'True
   End
   Begin MSDataListLib.DataCombo xNo2 
      Height          =   315
      Left            =   3555
      TabIndex        =   3
      Top             =   1665
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   "DataCombo1"
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "„”·”·"
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
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   675
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "≈·Ï Œ“‰…"
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
      Index           =   1
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬁÌ„…"
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
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2475
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ"
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
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1080
      Width           =   390
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì«‰"
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
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2115
      Width           =   315
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   "„‰ Œ“‰…"
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
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1425
      Width           =   660
   End
End
Attribute VB_Name = "boxtransfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Dim formMode As Byte
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode And bedit)
CmdDel.Enabled = (nMode = LoadMode And (xClosed.Value = 0))
CmdSave.Enabled = (xClosed.Value = 0) Or lSupperVisor
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xCode.Enabled = Not (nMode = LoadMode)
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me

Generalarray(1) = "SELECT  top 200 CODE,DESCA,CONVERT(VARCHAR(10),[DATE],111) FROM FILE0_51 ORDER BY DATE DESC "
Generalarray(2) = ""
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰ √Ê «· «—ÌŒ"
listarray(0, 1) = "(%%DESCA%% or ##date##)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 3000

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1200


searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "≈” ⁄·«„ «·Œ“‰"
Search3.Show 1
End Sub
Sub myDefine()
    cmd_closed.BackColor = &H8080FF
    cmd_closed.Caption = "› Õ „” ‰œ"
    xClosed.Visible = False
    xClosed.Value = ssCBUnchecked

xCode.text = RetZero(Val(Newflag("FILE0_51", "CODE", con)), 6)
xDescA.text = ""
If cBranch = "00" Then
    xDate.text = ""
Else
    xDate.text = Format(dSalesDate, "DD-MM-YYYY")
End If
xNo1.BoundText = ""
xNo2.BoundText = ""
xValue.text = ""
Handlecontrols DefineMode
End Sub
Sub myProc()
xCode.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
myUndo
Unload Search3
End Sub
Sub myload()
xCode.text = CardTable!CODE
xDescA.text = TurnValue(CardTable!Desca, Null, "")
xNo1.BoundText = TurnValue(CardTable!no1, Null, "")
xNo2.BoundText = TurnValue(CardTable!no2, Null, "")
xDate.text = Format(CardTable!Date, "DD-MM-YYYY")
xValue.text = Format(CardTable!Value, "Fixed")
    If CardTable!ISCLOSED Then
        cmd_closed.BackColor = &H8080FF
        cmd_closed.Caption = "› Õ „” ‰œ"
        xClosed.Value = ssCBChecked
        xClosed.Visible = True
    Else
        cmd_closed.BackColor = vbGreen
        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
        xClosed.Value = ssCBUnchecked
        xClosed.Visible = False
    End If
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim aInsert(7, 1)
aInsert(0, 0) = "CODE"
aInsert(0, 1) = addstring(xCode.text)

aInsert(1, 0) = "DESCA"
aInsert(1, 1) = addstring(xDescA.text)

aInsert(2, 0) = "Date"
aInsert(2, 1) = addDate(xDate.text)

aInsert(3, 0) = "NO1"
aInsert(3, 1) = addstring(xNo1.BoundText)

aInsert(4, 0) = "NO2"
aInsert(4, 1) = addstring(xNo2.BoundText)

aInsert(5, 0) = "[VALUE]"
aInsert(5, 1) = Val(xValue.text)

aInsert(6, 0) = "[ISNEW]"
aInsert(6, 1) = 1

aInsert(7, 0) = "ISClosed"
aInsert(7, 1) = 0

On Error GoTo myerror
con.BeginTrans
If xCode.Enabled Then
    con.Execute CreateInsert(aInsert, "FILE0_51")
Else
    con.Execute CreateUpdate(aInsert, "FILE0_51", " where CODE = " & addstring(xCode.text))
End If
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox " ”ÃÌ· „”·”· "
    Exit Function
End If
If xDate.text = "" Then
    MsgBox " ”ÃÌ·  «—ÌŒ "
    Exit Function
End If
If cBranch <> "00" And cBranch < "60" Then
    If DateValue(xDate.text) < dMaxDate And Not lSupperVisor Then
        MsgBox " «—ÌŒ €Ì— ’«·Õ "
        Exit Function
    End If
End If
MYVALID = True
End Function
Private Sub cmd_closed_Click()
If CardTable!ISCLOSED Then
    con.Execute " update FILE0_51 set isclosed = 0 where code = " & MyParn(xCode.text)
Else
    con.Execute " update FILE0_51 set isclosed = 1 where code = " & MyParn(xCode.text)
End If
myUndo
End Sub
Private Sub CmdAdd_Click()
    myDefine
    On Error Resume Next
    xCode.SetFocus
    Err.Clear
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE0_51 Where Code = " & MyParn(xCode.text)
    con.CommitTrans
    CardTable.Requery
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "Code < " & MyParn(xCode.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    Else
        myDefine
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdInform_Click()
    CardLookup
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
    myUndo
End Sub
Private Sub Command1_Click()
With CardTable
    .MoveFirst
    Do While Not .EOF
        cDoc = RetZero(!CODE, 6)
        con.Execute " UPDATE FILE0_51 SET CODE = " & addstring(cDoc) & " WHERE CODE = " & MyParn(!CODE)
        .MoveNext
    Loop
End With
    MsgBox "OK"
End Sub
Private Sub Form_Load()
openCon con
CmdDel.Visible = (cBranch = "00" Or cBranch > "60")
FRM_CLOSED.Visible = lSupperVisor Or (bopt2 And (cBranch = "00" Or cBranch > "60"))

Set data1.Recordset = myRecordSet("Select * From file0_50 WHERE ISSTOP = 0 ORDER BY CODE ", con)
Set xNo1.RowSource = data1
xNo1.ListField = "Desca"
xNo1.BoundColumn = "Code"

Set xNo2.RowSource = data1
xNo2.ListField = "Desca"
xNo2.BoundColumn = "Code"
xValue.Locked = (cBranch <> "00")
myUndo
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub

Private Sub xCode_LostFocus()
xCode.text = RetZero(xCode.text, 6)
CardTable.Find "Code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub xCode_GotFocus()
xCode.SelStart = 0
xCode.SelLength = Len(xCode.text)
End Sub
Private Sub xdate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.text)
End Sub
Private Sub xValue_DblClick()
    If Val(xValue.text) = 0 And cBranch <> "00" Then
        xValue.text = Round(Val(GetDesca("select sum(plus-minus) from BOXMOVE where box = " & MyParn(xNo1.BoundText) & " and date <= " & DateSq(xDate.text), con) & ""), 2)
    End If
End Sub

Private Sub xValue_GotFocus()
xValue.SelStart = 0
xValue.SelLength = Len(xValue.text)
End Sub
Private Sub xDescA_GotFocus()
xDescA.SelStart = 0
xDescA.SelLength = Len(xDescA.text)
End Sub
Private Sub myUndo()
If (xCode.text) <> "" Then
    openCardTable xCode.text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then myDefine Else myload
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Function openCardTable(Optional pCode As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset

cString = "SELECT TOP 1 * from FILE0_51 "
If pSign = "=" Then
    If pCode <> "" Then cWhere = " code " & pSign & addstring(pCode)
Else
    If pCode <> "" Then cWhere = " CODE " & pSign & addstring(pCode)
End If
If pCode <> "" Then cFilter = cFilter & turn(cFilter, " and ") & "CODE = " & MyParn(pCode)

If cWhere <> "" Then cString = cString & " WHERE " & cWhere
If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by CODE desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by CODE ASC"
End If
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function
Private Sub CmdNext_Click()
openCardTable xCode.text, ">"
If CardTable.EOF Then openCardTable xCode.text
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xCode.text, "<"
If CardTable.EOF Then openCardTable xCode.text
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub

