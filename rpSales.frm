VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form rpItem 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7575
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
   ScaleHeight     =   7575
   ScaleWidth      =   8160
   Begin VB.CommandButton CmdOk 
      BackColor       =   &H00C00000&
      Caption         =   "Œ—ÊÃ"
      Height          =   465
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6975
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
      Height          =   7170
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   7980
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   2
         Left            =   3780
         TabIndex        =   1
         Top             =   675
         Width           =   4080
         _ExtentX        =   7197
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
         Caption         =   " ›’Ì·Ï √—’œ… & Õ—ﬂ… «·„ÊœÌ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   3
         Left            =   3780
         TabIndex        =   2
         Top             =   1125
         Width           =   4080
         _ExtentX        =   7197
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
         Caption         =   "≈Ã„«·Ï ﬁÌ„… —’Ìœ „Õ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   4
         Left            =   3780
         TabIndex        =   3
         Top             =   1575
         Width           =   4080
         _ExtentX        =   7197
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
         Caption         =   " ﬁÌ„ √—’œ… „ÊœÌ·«  „Ê—œÌ‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   7
         Left            =   360
         TabIndex        =   4
         Top             =   2475
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
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
         Caption         =   " ›’Ì·Ï ÌÊ„Ì… „»Ì⁄«  „Õ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   1
         Left            =   3780
         TabIndex        =   0
         Top             =   225
         Width           =   4080
         _ExtentX        =   7197
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
         Caption         =   "√—’œ… «·„ÊœÌ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   6
         Left            =   3780
         TabIndex        =   7
         Top             =   2475
         Width           =   4080
         _ExtentX        =   7197
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
         Caption         =   "—’Ìœ & ﬁÌ„… „ÊœÌ·«  ( „‘ —Ì«  & „— Ã⁄«  & —’Ìœ ) „ÊÌœ·«  „Ê—œ"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   390
         Index           =   5
         Left            =   3780
         TabIndex        =   8
         Top             =   2025
         Width           =   4080
         _ExtentX        =   7197
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
         Caption         =   "⁄—÷ «—’œ… »«·”«·»"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   8
         Left            =   90
         TabIndex        =   9
         Top             =   2790
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
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
         Caption         =   "≈Ã„«·Ï ÌÊ„Ï ·„»Ì⁄«  „Õ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   9
         Left            =   -315
         TabIndex        =   10
         Top             =   3105
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
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
         Caption         =   "≈Ã„«·Ï ‘Â—Ï „»Ì⁄«  „Õ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   0
         Left            =   405
         TabIndex        =   11
         Top             =   3375
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
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
         Caption         =   "„ﬁ«—‰… „»Ì⁄«  „Õ·«  ÌÊ„Ï"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   10
         Left            =   4275
         TabIndex        =   12
         Top             =   3015
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
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
         Caption         =   " Ê“Ì⁄ ﬁÌ„… «·—’Ìœ »Ì‰ «·„Ê«”„ & «·„Ê—œÌ‰"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   11
         Left            =   585
         TabIndex        =   13
         Top             =   5040
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
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
         Caption         =   " ›’Ì·Ï ⁄„Ê·… »«∆⁄"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   12
         Left            =   90
         TabIndex        =   14
         Top             =   3780
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
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
         Caption         =   "„ﬁ«—‰… „»Ì⁄«  & —»Õ & „’«—Ì› ··„Õ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   13
         Left            =   4140
         TabIndex        =   15
         Top             =   3375
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
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
         Caption         =   " Ê“Ì⁄ ﬁÌ„… «·—’Ìœ »Ì‰ «·„Ê«”„ & «·„Õ·« "
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   435
         Index           =   14
         Left            =   720
         TabIndex        =   16
         Top             =   4455
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   767
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
         Caption         =   "≈Ã„«·Ï ﬁÌ„… ⁄„Ê·«  »Ì‰ «·»«∆⁄Ì‰"
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
Private Sub CmdGo_Click(Index As Integer)
publicFlag = Index
Select Case Index
Case 1
    rpitem1.Show 1
Case 2
    rpitem2.Show 1
Case 3
    rpItem3.Show 1
Case 4
    rpItem4.Show 1
Case 5
    rpItem5.Show 1
Case 6
    rpitem6.Show 1
Case 7
    rpSales1.Caption = cmdgo(Index).Caption
    rpSales1.Show 1
Case 8
    rpSales2.Caption = cmdgo(Index).Caption
    rpSales2.Show 1
Case 9
    rpSales3.Caption = cmdgo(Index).Caption
    rpSales3.Show 1
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
Private Sub Form_Load()
    openCon con
    cmdgo(7).Visible = bopt1
    cmdgo(8).Visible = bopt1
    cmdgo(9).Visible = bopt1
    CMD_FIX.Visible = bopt1
End Sub
