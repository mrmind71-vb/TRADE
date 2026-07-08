VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form rpCharge_br 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4170
   ClientLeft      =   30
   ClientTop       =   315
   ClientWidth     =   4815
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
   ScaleHeight     =   4170
   ScaleWidth      =   4815
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C00000&
      Caption         =   "Œ—ÊÃ"
      Height          =   465
      Left            =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3600
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
      Height          =   3375
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   150
      Width           =   4410
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   675
         Width           =   4035
         _ExtentX        =   7117
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
         Caption         =   " ›’Ì·Ì „’—Ê› ›—Ê⁄ Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   3
         Left            =   135
         TabIndex        =   3
         Top             =   1125
         Width           =   4035
         _ExtentX        =   7117
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
         Caption         =   "„ «»⁄… ÌÊ„Ì… ··„’«—Ì› ›—Ê⁄  ÊﬂÌ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   1
         Left            =   135
         TabIndex        =   0
         Top             =   225
         Width           =   4035
         _ExtentX        =   7117
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
         Caption         =   "≈Ã„«·Ì „’«—Ì› ›—Ê⁄  ÊﬂÌ·«  Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   4
         Left            =   135
         TabIndex        =   5
         Top             =   1620
         Width           =   4035
         _ExtentX        =   7117
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
         Caption         =   "≈Ã„«·Ï „’«—Ì› Õ”» «· ”ÃÌ·"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   5
         Left            =   135
         TabIndex        =   6
         Top             =   2115
         Width           =   4035
         _ExtentX        =   7117
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
         Caption         =   "≈Ã„«·Ï „Êﬁ› Õ”«»  ÊﬂÌ· "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   6
         Left            =   135
         TabIndex        =   7
         Top             =   2610
         Width           =   4035
         _ExtentX        =   7117
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
         Caption         =   "≈Ã„«·Ï „Êﬁ› Õ”«»  ÊﬂÌ·  - ·›—⁄ ·› —…"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "rpCharge_br"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOption As Integer
Private Sub CmdApply_Click()
End Sub
Private Sub cmdGo_Click(Index As Integer)
publicFlag = Index
Select Case Index
Case 1
    rpcharge_br1.Show 1
Case 2
   rpcharge_BR2.Show 1
Case 3
    rpcharge_BR3.Show 1
Case 4
    rpcharge_br4.Show 1
Case 5
    rpcharge_br5.Show 1
Case 6
    rpcharge_br6.Show 1
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
