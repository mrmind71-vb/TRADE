VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form rpSup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   3765
   ClientLeft      =   30
   ClientTop       =   435
   ClientWidth     =   9630
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3765
   ScaleWidth      =   9630
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C00000&
      Caption         =   "Œ—ÊÃ"
      Height          =   465
      Left            =   90
      TabIndex        =   2
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
      Height          =   3105
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   8460
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   1
         Left            =   4410
         TabIndex        =   0
         Top             =   225
         Width           =   3945
         _ExtentX        =   6959
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
         Caption         =   "«—’œ… «·„Ê—œÌ‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   4
         Left            =   4410
         TabIndex        =   3
         Top             =   675
         Width           =   3945
         _ExtentX        =   6959
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
         Caption         =   " «Ã„«·Ï ›Ê« Ì— „Ê—œÌ‰ Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   5
         Left            =   4410
         TabIndex        =   4
         Top             =   1125
         Width           =   3945
         _ExtentX        =   6959
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
         Caption         =   "≈Ã„«·Ì ›Ê« Ì— „—œÊœ „‘ —Ì«  Œ·«· › —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   8
         Left            =   225
         TabIndex        =   5
         Top             =   225
         Width           =   3945
         _ExtentX        =   6959
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
         Caption         =   " ›’Ì·Ì ›Ê« Ì— „‘ —Ì«  „Ê—œÌ‰ "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   9
         Left            =   225
         TabIndex        =   6
         Top             =   675
         Width           =   3945
         _ExtentX        =   6959
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
         Caption         =   " ›’Ì·Ï „—œÊœ „‘ —Ì«  „Ê—œÌ‰ "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   11
         Left            =   225
         TabIndex        =   7
         Top             =   1125
         Width           =   3945
         _ExtentX        =   6959
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
         Caption         =   " ﬁ—Ì— œ›⁄«  ‰ﬁœÌ…-‘Ìﬂ« -ﬁÌÊœ  ”ÊÌ« - ”ÊÌ« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   31
         Left            =   4410
         TabIndex        =   8
         Top             =   2070
         Width           =   3945
         _ExtentX        =   6959
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
         Caption         =   " «Ã„«·Ï ›Ê« Ì— „Ê—œÌ‰ Œ·«· › —… ·· ÊﬂÌ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   32
         Left            =   4410
         TabIndex        =   9
         Top             =   2520
         Width           =   3945
         _ExtentX        =   6959
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
         Caption         =   "≈Ã„«·Ì ›Ê« Ì— „—œÊœ „‘ —Ì«  Œ·«· › —… ·· ÊﬂÌ·« "
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "rpSup"
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
    rpSup1.Show 1
Case 4
    rpSup4.Show 1
Case 5
    rpSup5.Show 1

Case 8
    rpSup8.Show 1
Case 9
    rpSup9.Show 1
Case 11
    rpSup11.Show 1
Case 31
    rpSup31.Show 1
Case 32
    rpSup32.Show 1

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

Private Sub Form_Load()
'cmdGo(1).Visible = False
'
'cmdGo(4).Visible = False
'cmdGo(5).Visible = False
'cmdGo(11).Visible = False

End Sub
