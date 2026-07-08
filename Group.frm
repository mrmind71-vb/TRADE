VERSION 5.00
Begin VB.Form flag 
   ClientHeight    =   1980
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   5820
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1980
   ScaleWidth      =   5820
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   5820
      TabIndex        =   11
      Top             =   1470
      Width           =   5820
      Begin VB.CommandButton cmdFirst 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2070
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   90
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2505
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   90
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2940
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   90
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3375
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move Last"
         Top             =   90
         Width           =   435
      End
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   180
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1035
      Width           =   4515
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3690
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   630
      Width           =   1005
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   5820
      TabIndex        =   0
      Top             =   0
      Width           =   5820
      Begin VB.CommandButton CmdInform 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«” ⁄·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4725
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         Width           =   915
      End
      Begin VB.CommandButton CmdUndo 
         BackColor       =   &H00C0FFFF&
         Caption         =   " —«Ã⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2895
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«÷«›…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3810
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdDel 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ–›"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1050
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Õ›Ÿ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1980
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Œ—ÊÃ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   150
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   915
      End
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "«·ﬂ‹‹Êœ"
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
      Left            =   5025
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "«·»Ì‹‹«‰ "
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
      Left            =   4995
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1080
      Width           =   570
   End
End
Attribute VB_Name = "flag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myFlag As Integer, bEdit As Boolean
Dim formMode As Byte
Dim CardTable As New ADODB.Recordset
Dim RecordCountTable As Recordset
Const LoadMode = 1, DefineMode = 2
Private Sub Form_Load()
Select Case myFlag
Case 1
    flag.Caption = "»Ì«‰«  «·„Œ«“‰"
Case 2
    flag.Caption = "„Ã„Ê⁄«  «’‰«› —∆Ì”Ì…"
Case 3
    flag.Caption = "„Ã„Ê⁄«  «·⁄„·«¡"
Case 4
    flag.Caption = "„Ã„Ê⁄«  «·„Ê—œÌ‰"
Case 5
    flag.Caption = " «·‘—ﬂ«¡"
Case 6
    flag.Caption = "«·»«∆⁄Ì‰"
Case 7
    flag.Caption = "√ﬂÊ«œ „’«—Ì›"
Case 8
    flag.Caption = "«·„‰œÊ»Ì‰"
Case 10
    flag.Caption = "«·„‰œÊ»Ì‰"
Case 11
    flag.Caption = "«·„‰œÊ»Ì‰"
Case 12
    flag.Caption = "«·„‰œÊ»Ì‰"
Case 13
    flag.Caption = "≈Ì—«œ« "
Case 14
    flag.Caption = "„Ã„Ê⁄«  √’Ê·"
Case 15
    flag.Caption = "„Ã„Ê⁄«  Œ’Ê„"
End Select

CardTable.Open "Select * From File1_70 Where Flag = " & myFlag & " order by code", con, adOpenKeyset, adLockReadOnly, adCmdText
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    MyLoad
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
    con.Execute "Delete * From file1_70  Where code = " & MyParn(xCode.Text) & " and flag = " & myFlag
    con.CommitTrans
    CardTable.Requery
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.Find "code < " & MyParn(xCode.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        MyLoad
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
Inform " „ Õ›Ÿ »Ì«‰«  «·’‰› »‰Ã«Õ"
CardTable.Requery
CardTable.Find "code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
MyLoad
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    mydefine
Else
    If xCode.Enabled Then
        CardTable.MoveLast
    Else
        CardTable.Find "code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    MyLoad
End If
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
MyLoad
End Sub
Private Sub CmdInform_Click()
CardLookup
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
MyLoad
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    MyLoad
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    MyLoad
End If
End Sub
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode And bEdit)
CmdDel.Enabled = (nMode = LoadMode And bEdit)
CmdSave.Enabled = bEdit
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
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From file1_70 where flag = " & myFlag
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
If CardTable.EOF And CardTable.BOF Then
    xCode.Text = RetZero("1", 2)
Else
    CardTable.MoveLast
    xCode.Text = RetZero(Val(CardTable!CODE & "") + 1, 2)
End If
xDescA.Text = ""
Handlecontrols DefineMode
End Sub
Sub MyLoad()
xCode.Text = CardTable!CODE
xDescA.Text = CardTable!DESCA & ""
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Function myreplace() As Boolean
Dim nTry As Integer
'On Error Resume Next
For i = 1 To 1
    con.BeginTrans
    If xCode.Enabled Then
        con.Execute "insert into FILE1_70(CODE,DESCA,[FLAG])" & _
        "VALUES(" & _
        addstring(xCode.Text) & "," & _
        addstring(xDescA.Text) & "," & _
        myFlag & _
        ")"
    Else
        con.Execute "update FILE1_70 Set " & _
        "[DESCA] = " & addstring(xDescA.Text) & _
        " WHERE CODE = " & MyParn(xCode.Text) & " AND FLAG = " & myFlag
    End If
    
    If Err.Number = 0 Then Exit For
    If Err.Number = -2147467259 Then
        Err.Clear
        xCode.Text = RetZero(Val(xCode.Text) + 1, 2)
        con.RollbackTrans
    End If
    If Err.Number <> 0 Then GoTo myerror
Next
con.CommitTrans
myreplace = True
Exit Function
myerror:
con.RollbackTrans
If Err.Number <> 0 Then MsgBox Err.Description
Err.Clear
End Function
Sub myproc()
   CardTable.Find "CODE = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
   MyLoad
   Search3.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
Unload Search3
Set Search3 = Nothing
Err.Clear
End Sub
Private Sub xcode_LostFocus()
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 2)
CardTable.Find "CODE = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad
End Sub
Function MYVALID() As Boolean
If xCode.Text = "" Then
    MsgBox "«·ﬂÊœ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

If xDescA.Text = "" Then
    MsgBox "«·≈”„ ·« Ì„ﬂ‰ «‰ ÌﬂÊ‰ Œ«·Ì«"
    Exit Function
End If

MYVALID = True
End Function

