VERSION 5.00
Begin VB.Form InformOkfrm 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1050
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5055
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   1050
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   585
      Top             =   -90
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004BA6D6&
      Height          =   780
      Left            =   225
      TabIndex        =   0
      Top             =   135
      Width           =   4605
   End
End
Attribute VB_Name = "InformOkfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        KeyCode = 0
        Unload Me
    End If
End Sub

