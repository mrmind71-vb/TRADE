VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14490
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   14490
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   180
      TabIndex        =   3
      Top             =   9450
      Width           =   2580
   End
   Begin VB.TextBox Text3 
      Height          =   5010
      Left            =   45
      TabIndex        =   2
      Top             =   4320
      Width           =   20130
   End
   Begin VB.TextBox Text2 
      Height          =   2085
      Left            =   45
      TabIndex        =   1
      Top             =   2205
      Width           =   20130
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   20130
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text3.text = Serialize(Text1.text)
MsgBox Len(Text2.text)
MsgBox Len(Text3.text)
End Sub
