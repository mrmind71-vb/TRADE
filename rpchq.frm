VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form rpChq 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2775
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   9960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   2775
   ScaleWidth      =   9960
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   4785
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   5
         Left            =   150
         TabIndex        =   7
         Top             =   225
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   688
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Ê—«ﬁ œ›⁄ „” Õﬁ… Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   6
         Left            =   150
         TabIndex        =   8
         Top             =   675
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   688
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Ê—«ﬁ œ›⁄ „Õ’·… Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   7
         Left            =   135
         TabIndex        =   9
         Top             =   1125
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   688
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Ê—«ﬁ œ›⁄ „— œ… Œ·«· › —…"
         ButtonStyle     =   3
      End
   End
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C00000&
      Caption         =   "Œ—ÊÃ"
      Height          =   465
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2250
      Width           =   1890
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      Left            =   4905
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   4830
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   1
         Left            =   150
         TabIndex        =   1
         Top             =   225
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   688
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "√Ê—«ﬁ ﬁ»÷  „” Õﬁ… Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   2
         Left            =   150
         TabIndex        =   2
         Top             =   675
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   688
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "√Ê—«ﬁ ﬁ»÷ „Õ’·… Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   3
         Left            =   135
         TabIndex        =   4
         Top             =   1125
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   688
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Ê—«ﬁ ﬁ»÷ „— œ… Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   4
         Left            =   135
         TabIndex        =   5
         Top             =   1575
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   688
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Ê—«ﬁ ﬁ»÷ „Êœ⁄… Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   8
         Left            =   135
         TabIndex        =   10
         Top             =   2025
         Width           =   4545
         _ExtentX        =   8017
         _ExtentY        =   688
         _Version        =   196610
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "«Ê—«ﬁ ﬁ»÷ „‰ «·⁄„·«¡ » «—ÌŒ «·«” ·«„"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "rpChq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOption As Integer
Private Sub CmdGo_Click(Index As Integer)
publicFlag = Index
Select Case Index
Case 1, 2, 3, 4, 8
    rpChq1.Show 1
Case 5, 6, 7
    rpChq2.Show 1
End Select
End Sub
Private Sub cmdgo_MouseEnter(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdgo(Index).ForeColor = &HC00000
End Sub
Private Sub cmdgo_MouseExit(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdgo(Index).ForeColor = &H80000008
End Sub
Private Sub CmdOk_Click()
Unload Me
End Sub

