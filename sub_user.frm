VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form sub_userfrm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬂ·„… «·”—"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1845
   ScaleWidth      =   5865
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   -45
      Width           =   5595
      Begin VB.TextBox xPassword 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   180
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   360
         Width           =   4020
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "ﬂ·„… «·”—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   405
         Width           =   1185
      End
   End
   Begin Threed.SSCommand cmdExit 
      Cancel          =   -1  'True
      Height          =   645
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   1138
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "sub_user.frx":0000
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
   Begin Threed.SSCommand cmdOk 
      Default         =   -1  'True
      Height          =   645
      Left            =   1530
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1138
      _Version        =   196610
      ForeColor       =   0
      BackColor       =   16777215
      PictureFrames   =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "sub_user.frx":2323
      Alignment       =   8
      ButtonStyle     =   3
      PictureAlignment=   11
      BevelWidth      =   0
      ShapeSize       =   1
   End
End
Attribute VB_Name = "sub_userfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public bReturn As Boolean
Public sField As String
Public myForm As Form

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim loctable As New ADODB.RecordSet
Dim cString As String
cString = "SELECT SUB_USER.*" & _
          " FROM SUB_USER " & _
          " WHERE " & _
          " (PASSWORD = " & MyParn(xPassword.text) & " AND STOP1 = 0) OR " & _
          " (PASSWORD2 = " & MyParn(xPassword.text) & " AND STOP2 = 0)"

Set loctable = cmd(cString, con).Execute
If loctable.EOF Then
    MsgBox "ﬂ·„… ”— €Ì— ’ÕÌÕ…"
    myGotFocus xPassword
    Exit Sub
End If
If sField <> "" Then
    If Not loctable(sField) Then
        MsgBox "·Ì”  ·œÌﬂ «·’·«ÕÌ…"
        myGotFocus xPassword
        Exit Sub
    End If
End If
myForm.subUserCode = loctable!code & ""
myForm.subUserName = loctable!DESCA
Unload Me
End Sub
Private Sub Form_Load()
myForm.subUserCode = ""
myForm.subUserName = ""
OpenCon con
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set sub_userfrm = Nothing
End Sub
Private Sub xPassword_Change()
CmdOk.Enabled = Trim(xPassword.text) <> ""
End Sub
Private Sub xPassword_GotFocus()
myGotFocus xPassword
End Sub
Private Sub xPassword_LostFocus()
myLostFocus xPassword
End Sub
