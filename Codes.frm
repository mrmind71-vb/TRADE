VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form CodesFrm 
   ClientHeight    =   1965
   ClientLeft      =   420
   ClientTop       =   1470
   ClientWidth     =   6390
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
   ScaleHeight     =   1965
   ScaleWidth      =   6390
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   75
      Top             =   675
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
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
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   6390
      TabIndex        =   11
      Top             =   1500
      Width           =   6390
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
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "«ŒÌ—"
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
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   " «·Ì"
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
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "”«»ﬁ"
         Top             =   45
         Width           =   435
      End
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
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   45
         Width           =   435
      End
      Begin VB.Label xRecordNumber 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   2325
         TabIndex        =   12
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
      Left            =   3780
      MaxLength       =   3
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   675
      Width           =   690
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
      Left            =   1575
      MaxLength       =   40
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1035
      Width           =   2895
   End
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   6390
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   390
         Left            =   75
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "≈÷«›…"
         Height          =   390
         Left            =   4275
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton CmdInform 
         Caption         =   "≈” ⁄·«„"
         Height          =   390
         Left            =   5325
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   75
         Width           =   990
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   390
         Left            =   3225
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton CmdDel 
         Caption         =   "Õ–›"
         Height          =   390
         Left            =   1125
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ"
         Height          =   390
         Left            =   2175
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   75
         UseMaskColor    =   -1  'True
         Width           =   990
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   75
      Top             =   1050
      Visible         =   0   'False
      Width           =   1380
      _ExtentX        =   2434
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
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1125
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Left            =   4590
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   720
      Width           =   1485
   End
End
Attribute VB_Name = "CodesFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bEdit As Boolean
Dim formMode As Byte
Dim CardTable As ADODB.Recordset
Dim sFileName As String
Const LoadMode = 1, DefineMode = 2
Sub Handlecontrols(nMode)
cmdAdd.Enabled = (nMode = LoadMode And bEdit)
CmdDel.Enabled = (nMode = LoadMode And bEdit)
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
Generalarray(2) = "Select Code as [" & aPublic(2) & "],DescA as [" & aPublic(3) & "] From " & sFileName
Generalarray(3) = " Where DescA Like '%cFilter%'"
       
GrdArray(1) = 1200
GrdArray(2) = 4000
    
Lookupdata = Array(Generalarray, GrdArray)
Load Search
Search3.Caption = "«” ⁄·«„ "
Search3.Show 1
End Sub
Sub mydefine()
If CardTable.EOF And CardTable.BOF Then
    xCode.Text = "01"
Else
    CardTable.MoveLast
    xCode.Text = RetZero(IncRec(CardTable!CODE), 2)
End If
xdesca.Text = ""
Handlecontrols DefineMode
End Sub
Sub myproc()
CardTable.Find "Code = " & MyParn(GrdText(Search3.Grid1, 0)), , adSearchForward, adBookmarkFirst
MyLoad
End Sub
Sub MyLoad()
xCode.Text = CardTable!CODE
xdesca.Text = CardTable!DESCA
Handlecontrols LoadMode
End Sub
Sub myreplace()
CardTable.Find "Code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.AddNew
CardTable!CODE = xCode.Text
CardTable!DESCA = TurnValue(xdesca.Text)
CardTable.Update
End Sub
Function MYVALID() As Boolean
If xCode.Text = "" Then
    MsgBox " ”ÃÌ· " & aPublic(5)
    Exit Function
End If
MYVALID = True
End Function
Private Sub CmdAdd_Click()
mydefine
xCode.SetFocus
End Sub
Private Sub CmdDel_Click()
If MsgBox("«·€«¡ «·”Ã· «·Õ«·Ï : Â· «‰  „Ê«›ﬁ ø", 4) = 6 Then
con.Execute "Delete * From " & sFileName & "  Where Code = " & MyParn(xCode.Text)
CardTable.Requery
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.Find "Code < " & MyParn(xCode.Text), , adSearchBackward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveFirst
    MyLoad
Else
    mydefine
End If
End If
End Sub
Private Sub cmdExit_Click()
    Unload Me
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
Private Sub cmdSave_Click()
msgBoxStr = IIf(addmove, "«÷«›… ”Ã· : Â· «‰  „Ê«›ﬁ ø", "Õ›Ÿ «· €ÌÌ—«  ! Â· √‰  „Ê«›ﬁ ø")
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
    CardTable.Find "code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
    MyLoad
End If
End Sub
Private Sub CmdUndo_Click()
If (CardTable.EOF And CardTable.BOF) Then
    mydefine
Else
    If xCode.Enabled Then
        CardTable.MoveLast
        MyLoad
    Else
        MyLoad
    End If
End If
End Sub
Private Sub Form_Load()
Me.Caption = aPublic(1)
Label1.Caption = aPublic(2)
Label2.Caption = aPublic(3)
sFileName = aPublic(0)
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM " & sFileName & " ORDER BY CODE", con, adOpenKeyset, adLockOptimistic, adCmdText

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    MyLoad
Else
    xCode.Text = RetZero(1, 2)
    mydefine
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
End Sub
Private Sub xcode_LostFocus()
xCode.Text = RetZero(xCode.Text, 2)
CardTable.Find "Code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then MyLoad
End Sub
Private Function myRecordCount() As Integer
'If RecordCountTable.RecordCount = 0 Then Exit Function
'RecordCountTable.MoveLast
'myRecordCount = RecordCountTable.RecordCount
End Function
