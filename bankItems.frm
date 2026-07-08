VERSION 5.00
Begin VB.Form bankitemsfrm 
   Caption         =   "ÈäæÏ ÍÑßÉ ÇáÈäß"
   ClientHeight    =   1905
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   6450
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1905
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox SSPanel2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      Negotiate       =   -1  'True
      ScaleHeight     =   525
      ScaleWidth      =   6390
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   6390
      Begin VB.CommandButton CmdSave 
         Caption         =   "ÍÝÙ"
         Height          =   390
         Left            =   2175
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "ÍÐÝ"
         Height          =   390
         Left            =   1125
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   "ÊÑÇÌÚ"
         Height          =   390
         Left            =   3225
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton CmdInform 
         Caption         =   "ÅÓÊÚáÇã"
         Height          =   390
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "ÅÖÇÝÉ"
         Height          =   390
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "ÎÑæÌ"
         Height          =   390
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
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
      ScaleWidth      =   6450
      TabIndex        =   4
      Top             =   1440
      Width           =   6450
      Begin VB.CommandButton cmdfirst 
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
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Ãæá"
         Top             =   45
         Width           =   435
      End
      Begin VB.CommandButton cmdPrevious 
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
         Left            =   3285
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "ÓÇÈÞ"
         Top             =   45
         Width           =   435
      End
      Begin VB.CommandButton cmdNext 
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
         Left            =   2835
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "ÊÇáí"
         Top             =   45
         Width           =   435
      End
      Begin VB.CommandButton cmdLast 
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
         Left            =   2385
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "ÇÎíÑ"
         Top             =   45
         Width           =   435
      End
      Begin VB.Label xRecordNumber 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2325
         TabIndex        =   9
         Top             =   150
         Width           =   1890
      End
   End
   Begin VB.TextBox xCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   3825
      MaxLength       =   3
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   630
      Width           =   735
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
      Left            =   945
      MaxLength       =   40
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1035
      Width           =   3615
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ÈíÇä ÇáÈäÏ "
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
      Left            =   4770
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ßæÏ ÇáÈäÏ "
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
      Left            =   4770
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   675
      Width           =   690
   End
End
Attribute VB_Name = "bankitemsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim formMode As Byte
Dim CardTable As ADODB.Recordset
Dim sFileName As String
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
CmdAdd.Enabled = (nMode = LoadMode And bEdit)
cmdDel.Enabled = (nMode = LoadMode And bEdit)
CmdSave.Enabled = bEdit
CmdInform.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
xCode.Enabled = Not (nMode = LoadMode)
End Sub
Sub CardLookup()
Dim Generalarray(3)
Dim GrdArray(2)
Set Generalarray(1) = Me
Generalarray(2) = "Select Code as ÇáßæÏ,DescA as [ÈíÇä ÇáãÌãæÚÉ] From " & sFileName
Generalarray(3) = " Where DescA Like '%cFilter%'"
       
GrdArray(1) = 1200
GrdArray(2) = 4000
    
Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search3.Caption = "ÇÓÊÚáÇã "
Search3.Show 1
End Sub
Sub myDefine()
xDescA.Text = ""
Handlecontrols DefineMode
'If myRecordCount = 0 Then
'    xRecordNumber = ""
'Else
'    xRecordNumber = "ÓÌá " & myRecordCount + 1 & " ãä " & myRecordCount + 1
'End If
End Sub
Sub myProc()
CardTable.Find "Code = " & MyParn(GrdText(Search3.grid1, 0)), , adSearchForward, adBookmarkFirst
myload
End Sub
Sub myload()
xCode.Text = CardTable!CODE
xDescA.Text = CardTable!desca
'xMainGroup.BoundText = TurnValue(CardTable!MainGroup, Null, "")
'xCenter.BoundText = TurnValue(CardTable!Center, Null, "")
'xRecordNumber = "ÓÌá " & CARDTABLE.AbsolutePosition + 1 & " ãä " & myRecordCount
Handlecontrols LoadMode
End Sub
Sub myreplace()
CardTable.Find "Code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.AddNew
CardTable!CODE = xCode.Text
CardTable!desca = xDescA.Text
'CardTable!MainGroup = TurnValue(xMainGroup.BoundText, "", Null)
'CardTable!Center = TurnValue(xCenter.BoundText, "", Null)
CardTable.Update
End Sub
Function MYVALID() As Boolean
If xCode.Text = "" Then
    MsgBox "ÊÓÌíá ãÓáÓá "
    Exit Function
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
CardTable.MoveLast
xCode.Text = RetZero(IncRec(CardTable!CODE), 2)
myDefine
xCode.SetFocus
End Sub
Private Sub CmdDel_Click()
If MsgBox("ÇáÛÇÁ ÇáÓÌá ÇáÍÇáì : åá ÇäÊ ãæÇÝÞ ¿", 4) = 6 Then
con.Execute "Delete * From " & sFileName & "  Where Code = " & MyParn(xCode.Text)
CardTable.Requery
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "Code < " & MyParn(xCode.Text), , adSearchBackward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveFirst
    myload
Else
    myDefine
End If
End If
End Sub
Private Sub CmdExit_Click()
    Unload Me
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
Private Sub cmdSave_Click()
msgBoxStr = IIf(addmove, "ÇÖÇÝÉ ÓÌá : åá ÇäÊ ãæÇÝÞ ¿", "ÍÝÙ ÇáÊÛííÑÇÊ ! åá ÃäÊ ãæÇÝÞ ¿")
If Not MYVALID Then Exit Sub

If Not MsgBox(msgBoxStr, 4) = 6 Then
    CmdUndo_Click
    Exit Sub
End If
myreplace
CardTable.Requery
If xCode.Enabled Then
    CmdAdd_Click
Else
    CardTable.Find "code = " & MyParn(xCode.Text), adSearchForward, adBookmarkFirst
    myload
End If
End Sub
Private Sub CmdUndo_Click()
If (CardTable.EOF And CardTable.BOF) Then
    myDefine
Else
    If xCode.Enabled Then
        CardTable.MoveLast
        myload
    Else
        myload
    End If
End If
End Sub
Private Sub Form_Load()
openCon con
sFileName = "File5_00"
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM " & sFileName & " ORDER BY CODE", con, adOpenDynamic, adLockPessimistic, adCmdText

'data1.ConnectionString = strCon
'data1.RecordSource = "Select * From file1_70 where flag = " & IIf(publicFlag = 0, 7, 13)
'Set xMainGroup.RowSource = data1
'xMainGroup.ListField = "Desca"
'xMainGroup.BoundColumn = "Code"

'DATA2.ConnectionString = strCon
'DATA2.RecordSource = "Qfile1_70_2"
'Set xCenter.RowSource = DATA2
'xCenter.ListField = "Desca"
'xCenter.BoundColumn = "Code"

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    xCode.Text = RetZero(1, 2)
    myDefine
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub xCode_LostFocus()
xCode.Text = RetZero(xCode.Text, 2)
CardTable.Find "Code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function myRecordCount() As Integer
'If RecordCountTable.RecordCount = 0 Then Exit Function
'RecordCountTable.MoveLast
'myRecordCount = RecordCountTable.RecordCount
End Function
