VERSION 5.00
Begin VB.Form ManagPass 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1335
   ClientLeft      =   4020
   ClientTop       =   3465
   ClientWidth     =   4590
   ControlBox      =   0   'False
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   1335
   ScaleWidth      =   4590
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   780
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   225
      Width           =   3570
      Begin VB.TextBox xPass 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   180
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   2310
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ·„… «·”— :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   1005
      End
   End
End
Attribute VB_Name = "ManagPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pPassWord As String
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Unload Me
End Sub
Private Sub xPass_GotFocus()
    Clipboard.Clear
End Sub
Private Sub xPass_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        pPassWord = Trim(xPass.text)
        Unload Me
    End If
End Sub
Private Sub xPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
End Sub
