VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FixPriceOkaz 
   Caption         =   " ÕœÌœ ”⁄— «·»Ì⁄"
   ClientHeight    =   2370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "«·»Ì⁄ »√”⁄«— «·«Êﬂ«“ÊÌ‰ ··„ÊœÌ·«  «·„Õœœ… ›ﬁÿ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   465
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1620
      Width           =   6585
   End
   Begin VB.CommandButton PRICE1 
      BackColor       =   &H00E0E0E0&
      Caption         =   " ”⁄Ì— ”⁄— »Ì⁄ "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3555
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   3195
   End
   Begin VB.CommandButton PRICE2 
      BackColor       =   &H00E0E0E0&
      Caption         =   " ”⁄Ì— ”⁄— »Ì⁄ √Êﬂ«“ÌÊ‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   3150
   End
   Begin Threed.SSCommand xHead 
      Height          =   540
      Left            =   135
      TabIndex        =   2
      Top             =   90
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   953
      _Version        =   196610
      Font3D          =   2
      ForeColor       =   192
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1"
      ButtonStyle     =   4
   End
End
Attribute VB_Name = "FixPriceOkaz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub Check1_Click()
If Check1.Value <> 0 Then
    con.Execute " UPDATE FILE0_00 SET OKAZ_ITEM2 = 1 "
Else
    con.Execute " UPDATE FILE0_00 SET OKAZ_ITEM2 = 0 "
End If
End Sub

Private Sub Form_Load()
    con.CursorLocation = adUseClient
    con.Open strCon
    Dim cRet As String
    Dim cRet2 As String
    cRet2 = GetDesca(" SELECT OKAZ_ITEM2 FROM FILE0_00 ", con)
    If cRet2 = "True" Then
        Check1.Value = 1
    Else
        Check1.Value = 0
    End If
    If lokaz Then
        xhead.Caption = " «· ”⁄Ì— «·Õ«·Ï ÂÊ ”⁄— »Ì⁄ «·√Êﬂ«“ÊÌ‰ "
        PRICE1.Enabled = True
        PRICE2.Enabled = False
    Else
        xhead.Caption = " «· ”⁄Ì— «·Õ«·Ï ÂÊ ”⁄— »Ì⁄ «·„” Â·ﬂ  "
        PRICE1.Enabled = False
        PRICE2.Enabled = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    closeCon con
    Set con = Nothing
End Sub

Private Sub PRICE1_Click()
lokaz = False
con.Execute " UPDATE FILE0_00 SET PRICE = 0 "
Unload Me
End Sub
Private Sub PRICE2_Click()
lokaz = True
con.Execute " UPDATE FILE0_00 SET PRICE = 1 "
Unload Me
End Sub

