VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form boxtrans_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕÊÌ·«  «·Œ“‰ «·Œ«’… »«·›—Ê⁄"
   ClientHeight    =   4110
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   6960
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
   ScaleHeight     =   4110
   ScaleWidth      =   6960
   Begin VB.Frame FRM_CLOSED 
      Height          =   1095
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   2430
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
         TabIndex        =   28
         Top             =   135
         Visible         =   0   'False
         Width           =   1410
      End
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         TabIndex        =   29
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
         Picture         =   "boxtrans_2.frx":0000
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
      Left            =   4005
      Locked          =   -1  'True
      MaxLength       =   10
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
      Left            =   4005
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
      Left            =   3825
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
      ScaleWidth      =   6960
      TabIndex        =   11
      Top             =   3645
      Width           =   6960
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
         Picture         =   "boxtrans_2.frx":25CC
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
         Picture         =   "boxtrans_2.frx":4B8C
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
         Picture         =   "boxtrans_2.frx":715F
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
         Picture         =   "boxtrans_2.frx":9839
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
      Left            =   1125
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
      ScaleWidth      =   6960
      TabIndex        =   7
      Top             =   0
      Width           =   6960
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   5760
         Picture         =   "boxtrans_2.frx":BF18
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
         Picture         =   "boxtrans_2.frx":E6EB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   1185
         MaskColor       =   &H00FFFFFF&
         Picture         =   "boxtrans_2.frx":10C97
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
         Picture         =   "boxtrans_2.frx":13210
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   45
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   3465
         MaskColor       =   &H00FFFFFF&
         Picture         =   "boxtrans_2.frx":15AAA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   45
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "boxtrans_2.frx":15EEC
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
      Left            =   225
      Top             =   750
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
   Begin MSDataListLib.DataCombo xNo1 
      Height          =   315
      Left            =   2475
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
      Left            =   2475
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
   Begin VB.Label XBRANCH 
      Alignment       =   1  'Right Justify
      Caption         =   "Label5"
      Height          =   375
      Left            =   3375
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   585
      Width           =   555
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
      Left            =   5565
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
      Left            =   5550
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
      Left            =   5535
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
      Left            =   5565
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
      Left            =   5550
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
      Left            =   5565
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1425
      Width           =   660
   End
End
Attribute VB_Name = "boxtrans_2"
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

CmdAdd.Enabled = (nMode = LoadMode)

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
Dim GrdArray(6, 1)

Set Generalarray(0) = Me
'                                   0           1                   2                           3               4               5                       6
Generalarray(1) = "SELECT file0_52.CODE,file0_52.DESCA,CONVERT(VARCHAR(10),[DATE],111) , file0_52.BRANCH , file0_50.desca , file0_50_2.desca , value FROM FILE0_52 inner join file0_50 on file0_52.no1 = file0_50.code inner join file0_50 as file0_50_2 on file0_50_2.code = file0_52.no2 "
If lIsBranchStore Then Generalarray(1) = Generalarray(1) & " AND file0_52.BRANCH = " & MyParn(cBranch)
Generalarray(2) = ""
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰ √Ê «· «—ÌŒ"
listarray(0, 1) = "(%%DESCA%% or ##date##)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 3000

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1200

GrdArray(3, 0) = "›—⁄"
GrdArray(3, 1) = 700

GrdArray(4, 0) = "„‰ "
GrdArray(4, 1) = 1500

GrdArray(5, 0) = "≈·Ï "
GrdArray(5, 1) = 1500

GrdArray(6, 0) = "«·„»·€"
GrdArray(6, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "≈” ⁄·«„  ÕÊÌ·«  Œ“‰"
Search3.Show 1
End Sub
Sub myDefine()

    cmd_closed.BackColor = &H8080FF
    cmd_closed.Caption = "› Õ „” ‰œ"
    xClosed.Visible = False
    xClosed.Value = ssCBUnchecked

xCode.text = RetZero(Val(NewflagBranch("FILE0_52", "CODE", cBranch, con)), 8)
XBRANCH.Caption = cBranch
XDESCA.text = ""
xDate.text = Format(dSalesDate, "DD-MM-YYYY")
xNo1.BoundText = ""
xNo2.BoundText = ""
XVALUE.text = ""

Handlecontrols DefineMode
End Sub
Sub myProc()
CardTable.Find "CODE = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
Unload Search3
End Sub
Sub myload()
xCode.text = CardTable!CODE
XDESCA.text = TurnValue(CardTable!DESCA, Null, "")
XBRANCH.Caption = CardTable!branch & ""
xNo1.BoundText = TurnValue(CardTable!no1, Null, "")
xNo2.BoundText = TurnValue(CardTable!no2, Null, "")
xDate.text = Format(CardTable!Date, "DD-MM-YYYY")
XVALUE.text = Format(CardTable!Value, "Fixed")

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
Dim aInsert(8, 1)
aInsert(0, 0) = "CODE"
aInsert(0, 1) = addstring(xCode.text)

aInsert(1, 0) = "DESCA"
aInsert(1, 1) = addstring(XDESCA.text)

aInsert(2, 0) = "Date"
aInsert(2, 1) = addDate(xDate.text)

aInsert(3, 0) = "NO1"
aInsert(3, 1) = addstring(xNo1.BoundText)

aInsert(4, 0) = "NO2"
aInsert(4, 1) = addstring(xNo2.BoundText)

aInsert(5, 0) = "[VALUE]"
aInsert(5, 1) = Val(XVALUE.text)

aInsert(6, 0) = "[ISNEW]"
aInsert(6, 1) = 1

aInsert(7, 0) = "BRANCH"
aInsert(7, 1) = addstring(XBRANCH.Caption)

aInsert(8, 0) = "ISClosed"
aInsert(8, 1) = 0


On Error GoTo myerror
con.BeginTrans
If xCode.Enabled Then
    con.Execute CreateInsert(aInsert, "FILE0_52")
Else
    con.Execute CreateUpdate(aInsert, "FILE0_52", " where CODE = " & addstring(xCode.text))
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
MYVALID = True
End Function
Private Sub cmd_closed_Click()
    If CardTable!ISCLOSED Then
        con.Execute " update FILE0_52 set isclosed = 0 where code = " & MyParn(xCode.text)
        xClosed.Value = 0
    Else
        con.Execute " update FILE0_52 set isclosed = 1 where code = " & MyParn(xCode.text)
        xClosed.Value = 1
    End If
    myload
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
    con.Execute "Delete  From FILE0_52 Where Code = " & MyParn(xCode.text)
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
Private Sub cmdExit_Click()
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
If CardTable.EOF And CardTable.BOF Then
    myDefine
Else
    If xCode.Enabled Then
        CardTable.MoveLast
    Else
        CardTable.Find "code = " & MyParn(xCode.text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    End If
    myload
End If
End Sub
Private Sub Command1_Click()
'''''With CardTable
'''''    .MoveFirst
'''''    Do While Not .EOF
'''''        cDoc = RetZero(!CODE, 6)
'''''        con.Execute " UPDATE FILE0_52 SET CODE = " & addstring(cDoc) & " WHERE CODE = " & MyParn(!CODE)
'''''        .MoveNext
'''''    Loop
'''''End With
'''''    MsgBox "OK"
End Sub
Private Sub Form_Load()
openCon con
FRM_CLOSED.Visible = lSupperVisor Or (bopt2 And (cBranch = "00" Or cBranch > "60"))

If lIsBranchStore Then
    CmdSave.Visible = True
    CmdDel.Visible = True
    cmdNext.Visible = True
    CmdAdd.Visible = True
Else
    CmdSave.Visible = False
    CmdDel.Visible = False
    cmdNext.Visible = False
    CmdAdd.Visible = False
End If

Set CardTable = New ADODB.Recordset
If lIsBranchStore Then
    CardTable.Open "SELECT * FROM FILE0_52 WHERE BRANCH = " & MyParn(cBranch) & " ORDER BY CODE", con, adOpenKeyset, adLockOptimistic, adCmdText
Else
    CardTable.Open "SELECT * FROM FILE0_52 ORDER BY CODE", con, adOpenKeyset, adLockOptimistic, adCmdText
End If

If cBranch = "00" Then
    Set DATA1.Recordset = myRecordSet("Select * From file0_50", con)
Else
    Set DATA1.Recordset = myRecordSet("Select * From file0_50 WHERE BRANCH = " & MyParn(cBranch), con)
End If
Set xNo1.RowSource = DATA1
xNo1.ListField = "Desca"
xNo1.BoundColumn = "Code"

Set xNo2.RowSource = DATA1
xNo2.ListField = "Desca"
xNo2.BoundColumn = "Code"

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
End If
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
Private Sub xValue_GotFocus()
XVALUE.SelStart = 0
XVALUE.SelLength = Len(XVALUE.text)
End Sub
Private Sub xDescA_GotFocus()
XDESCA.SelStart = 0
XDESCA.SelLength = Len(XDESCA.text)
End Sub
Private Sub xValue_DblClick()
    If lIsBranchStore Then
        If Val(XVALUE.text) = 0 And cBranch <> "00" Then
            XVALUE.text = Round(Val(GetDesca("select sum(plus-minus) from BOXMOVE where box = " & MyParn(xNo1.BoundText) & " and date <= " & DateSq(xDate.text), con) & ""), 2)
        End If
    End If
End Sub

