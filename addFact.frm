VERSION 5.00
Begin VB.Form addFactfrm 
   Caption         =   "«÷«›… „’«‰⁄"
   ClientHeight    =   1110
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1110
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   6855
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   4875
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3690
         MaxLength       =   3
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label Label3 
         Caption         =   "≈”„ «·„’‰⁄ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "«·ﬂÊœ «·„Œ ’— :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   1275
      End
   End
End
Attribute VB_Name = "addFactfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public sCode As String, sdesca As String
Private Function myreplace()
Dim aInsert(1, 1)
aInsert(0, 0) = "Code"
aInsert(0, 1) = addstring(UCase(xCode.text))

aInsert(1, 0) = "Desca"
aInsert(1, 1) = addstring(xDesca.text)

On Error GoTo myerror
con.BeginTrans
If GetDesca("Select Code from fact where Code = " & MyParn(xCode.text), con) = "" Then
   con.Execute CreateInsert(aInsert, "fact")
Else
   con.Execute CreateUpdate(aInsert, "Fact", " WHERE CODE = " & MyParn(xCode.text), Array(-1))
End If
con.CommitTrans
Inform " „ «· ⁄œÌ· »‰Ã«Õ"
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Form_Load()
    openCon con
    myload
End Sub
Private Sub myload()
xCode.text = sCode
xDesca.text = sdesca

Dim aRet As Variant
aRet = aGetDesca("Select Code,desca from Fact where code = " & MyParn(sCode), con)

If UBound(aRet) > 0 Then
    If Trim(xDesca.text) = "" Then xDesca.text = sdesca
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set addFactfrm = Nothing
closeCon con
End Sub
Private Sub xcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MYVALID Then
        myreplace
        Unload Me
    End If
End If
End Sub
Private Function MYVALID() As Boolean
If Len(Trim(xCode.text)) <> 3 Then
    MsgBox "«·ﬂÊœ ·« Ì’·Õ"
    Exit Function
End If
Dim aRet As Variant
aRet = aGetDesca("Select Code,desca from Fact where code = " & MyParn(xCode.text), con)
If UBound(aRet) > 0 Then
    If MsgBox("«·„’‰⁄ „ÊÃÊœ „‰ ﬁ»· Â·  Êœ «· ⁄œÌ·", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
End If
MYVALID = True
End Function
Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If MYVALID Then
        myreplace
        Unload Me
    End If
End If
End Sub
