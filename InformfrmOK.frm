VERSION 5.00
Begin VB.Form InformfrmOK 
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5835
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_ok 
      Caption         =   "O K "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   45
      TabIndex        =   2
      Top             =   1665
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Left            =   30
      Top             =   -15
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ENTER  ≈÷€ÿ ··≈” „—«—"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   1215
      TabIndex        =   1
      Top             =   1305
      Width           =   3390
   End
   Begin VB.Label lbl_inform 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   15
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1140
      Left            =   270
      TabIndex        =   0
      Top             =   90
      Width           =   5550
   End
End
Attribute VB_Name = "InformfrmOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_ok_Click()
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Unload Me
End Sub

Private Sub Form_Load()
'Timer1.Interval = 900

End Sub
Private Sub Timer1_Timer()
'Timer1.Enabled = False
'Unload Me
End Sub
