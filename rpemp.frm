VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form rpempfrm 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3690
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   8160
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
   ScaleHeight     =   3690
   ScaleWidth      =   8160
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C00000&
      Caption         =   "Œ—ÊÃ"
      Height          =   465
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3195
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
      Height          =   3015
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   150
      Width           =   7965
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   2
         Left            =   4455
         TabIndex        =   2
         Top             =   675
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   " ›’Ì·Ì „’—Ê› Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   4
         Left            =   4455
         TabIndex        =   3
         Top             =   1575
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   "≈Ã„«·Ì «Ì—«œ«  Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   5
         Left            =   4455
         TabIndex        =   4
         Top             =   2025
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   " ›’Ì·Ì «Ì—«œ Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   3
         Left            =   4455
         TabIndex        =   5
         Top             =   1125
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   "„ «»⁄… ÌÊ„Ì… ··„’«—Ì›"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   1
         Left            =   4455
         TabIndex        =   0
         Top             =   225
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   688
         _Version        =   196610
         CaptionStyle    =   1
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
         Caption         =   " ›’Ì·Ì »Ì«‰«  „ÊŸ›Ì‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   6
         Left            =   4455
         TabIndex        =   7
         Top             =   2475
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   "„ «»⁄… ÌÊ„Ì… ··«Ì—«œ« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   7
         Left            =   135
         TabIndex        =   8
         Top             =   225
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   " ›’Ì·Ì  ÕÊÌ·«  „‰ Œ“‰… ·Œ“‰…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   8
         Left            =   150
         TabIndex        =   9
         Top             =   675
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   " ›’Ì·Ï „”ÕÊ»«  ‘—ﬂ«¡"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   9
         Left            =   150
         TabIndex        =   10
         Top             =   1125
         Width           =   3405
         _ExtentX        =   6006
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
         Caption         =   "≈Ã„«·Ï „”ÕÊ»«  ‘—ﬂ«¡"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "rpempfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOption As Integer
Private Sub CmdApply_Click()
End Sub
Private Sub CmdGo_Click(Index As Integer)
publicFlag = Index
Select Case Index
Case 1
    rpCharge1.Show 1
Case 2
   rpCharge2.Show 1
Case 3
    rpCharge3.Show 1
Case 4
    rpCharge1.Show 1
Case 5
    rpCharge2.Show 1
Case 6
    rpCharge3.Show 1
Case 7
    rpBox1.Show 1
Case 8
    rpCharge8.Show 1
Case 9
    rpCharge9.Show 1
End Select
End Sub
Private Sub cmdgo_MouseEnter(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdGo(Index).ForeColor = &HC00000
End Sub
Private Sub cmdgo_MouseExit(Index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdGo(Index).ForeColor = &H80000008
End Sub
Private Sub CmdOk_Click()
Unload Me
End Sub

