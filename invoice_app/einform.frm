VERSION 5.00
Begin VB.Form eInformfrm 
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5730
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   915
   ScaleWidth      =   5730
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Left            =   585
      Top             =   -90
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Enter ··«” „—«— «÷€ÿ "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   270
      TabIndex        =   2
      Top             =   990
      Visible         =   0   'False
      Width           =   5220
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   2835
      TabIndex        =   1
      Top             =   510
      Width           =   90
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2535
      TabIndex        =   0
      Top             =   45
      Width           =   120
   End
End
Attribute VB_Name = "eInformfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public sLabel1 As String, sLabel2 As String, nInterval As Integer
Dim nMargin As Integer
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Unload Me
End Sub
Private Sub Form_Load()
Dim FileName As String
Label1.Caption = sLabel1
Label2.Caption = sLabel2
If nInterval = -1 Then
    Timer1.Enabled = False
    If sLabel2 = "" Then
        Label2.Caption = Label3.Caption
        Label2.ForeColor = Label3.ForeColor
    Else
        Me.Height = Me.Height + Label3.Height + 100
        Label3.Visible = True
    End If
Else
    Timer1.Interval = nInterval
End If
If Label2.Caption = "" Then Me.Height = Me.Height - (Label2.Height + 100)

nMargin = 1000
If Label1.Width < Label2.Width Then Label1.Width = Label2.Width
If Label3.Visible And Label1.Width < Label3.Width Then Label1.Width = Label3.Width
Me.Width = Label1.Width + nMargin + 100
Label1.Left = nMargin / 2

Label2.Width = Label1.Width
Label2.Left = Label1.Left

Label3.Width = Label1.Width
Label3.Left = Label1.Left
End Sub
Private Sub Timer1_Timer()
Timer1.Enabled = False
Unload Me
End Sub
