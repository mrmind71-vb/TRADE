VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form rpItem 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6345
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6345
   ScaleWidth      =   8160
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C00000&
      Caption         =   "Œ—ÊÃ"
      Height          =   465
      Left            =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5310
      Width           =   1485
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
      Height          =   5055
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   45
      Width           =   7980
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   2
         Left            =   675
         TabIndex        =   1
         Top             =   690
         Width           =   6510
         _ExtentX        =   11483
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
         Caption         =   " ›’Ì·Ï „‘ —Ì«  ° „»Ì⁄«  ° —’Ìœ ( ⁄œœ - ﬁÌ„… ) ·„Ê—œ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   5
         Left            =   675
         TabIndex        =   2
         Top             =   2085
         Width           =   6510
         _ExtentX        =   11483
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
         Caption         =   "≈Ã„«·Ï —’Ìœ „ÊœÌ·«  „Ê—œ  ⁄œœ /  ﬁÌ„…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   1
         Left            =   675
         TabIndex        =   0
         Top             =   225
         Width           =   6510
         _ExtentX        =   11483
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
         Caption         =   "√—’œ… Ê  ﬁÌ„ «·„ÊœÌ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   6
         Left            =   675
         TabIndex        =   5
         Top             =   3585
         Visible         =   0   'False
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   688
         _Version        =   196610
         Enabled         =   0   'False
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
         Caption         =   "—’Ìœ & ﬁÌ„… „ÊœÌ·«  ( „‘ —Ì«  & „— Ã⁄«  & —’Ìœ ) ·„ÊœÌ·«  „Ê—œ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   7
         Left            =   675
         TabIndex        =   6
         Top             =   4035
         Visible         =   0   'False
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   767
         _Version        =   196610
         Enabled         =   0   'False
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
         Caption         =   " Ê“Ì⁄ ﬁÌ„… «·—’Ìœ »Ì‰ «·„Ê«”„ & «·„Ê—œÌ‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   8
         Left            =   675
         TabIndex        =   7
         Top             =   4545
         Visible         =   0   'False
         Width           =   6510
         _ExtentX        =   11483
         _ExtentY        =   767
         _Version        =   196610
         Enabled         =   0   'False
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
         Caption         =   " Ê“Ì⁄ ﬁÌ„… «·—’Ìœ »Ì‰ «·„Ê«”„ & «·„Õ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   3
         Left            =   675
         TabIndex        =   8
         Top             =   1155
         Width           =   6510
         _ExtentX        =   11483
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
         Caption         =   " ›’Ì·Ï „»Ì⁄«  „ÊœÌ·«  „Ê—œ ·› —…"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   4
         Left            =   675
         TabIndex        =   9
         Top             =   1620
         Width           =   6510
         _ExtentX        =   11483
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
         Caption         =   "√—’œ… «·„ÊœÌ·«  ( ›’Ì·Ï „ﬁ«”«  & √·√Ê«‰ )  &   ﬁÌÌ„ «·—’Ìœ"
         ButtonStyle     =   3
      End
   End
End
Attribute VB_Name = "rpItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim nOption As Integer
Private Sub CmdApply_Click()
End Sub
Private Sub CmdGo_Click(index As Integer)
publicFlag = index
Select Case index
Case 1
    rpitem1.Show 1
Case 2
    rpitem2_2.Show 1
Case 3
    rpitem3.Show 1
Case 4
    rpitem4.Show 1
Case 5
    rpitem5.Show 1
Case 6
    rpitem4.Show 1

End Select
End Sub
Private Sub cmdgo_MouseEnter(index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdGo(index).ForeColor = &HC00000
End Sub
Private Sub cmdgo_MouseExit(index As Integer, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
cmdGo(index).ForeColor = &H80000008
End Sub
Private Sub CmdOk_Click()
Unload Me
End Sub
Private Sub Form_Load()
    openCon con
'    cmdgo(7).Visible = bopt1
'    cmdgo(8).Visible = bopt1
'    cmdgo(9).Visible = bopt1
'    CMD_FIX.Visible = bopt1
End Sub
