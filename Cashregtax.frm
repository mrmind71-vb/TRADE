VERSION 5.00
Begin VB.Form Cashregtax 
   Caption         =   " "
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3645
   ScaleWidth      =   4215
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox xVisa 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1890
      Width           =   2175
   End
   Begin VB.TextBox xPay 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   360
      MaxLength       =   15
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1215
      Width           =   2175
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D8FCFE&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÝíÒÇ"
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2655
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1980
      Width           =   1275
   End
   Begin VB.Label xRest 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   915
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2610
      Width           =   2175
   End
   Begin VB.Label xTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   660
      Width           =   2175
   End
   Begin VB.Label xtotalQuant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   90
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D8FCFE&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÚÏÏ ÇáÞØÚ ÇáãÈÇÚÉ"
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2655
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   90
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D8FCFE&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÅÌãÇáì ÞíãÉ ÇáÈæä"
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2655
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   675
      Width           =   1275
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D8FCFE&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÇáãÏÝæÚ"
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2655
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1260
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00D8FCFE&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ÇáÈÇÞì"
      BeginProperty Font 
         Name            =   "Traditional Arabic"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   465
      Left            =   2655
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2790
      Width           =   1275
   End
End
Attribute VB_Name = "Cashregtax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lExit As Boolean
Private Sub Form_Load()
xTotal.Caption = Myvalue(SalesFrm_tax.xTotal.Text, "Fixed")
xPay.Text = Myvalue(SalesFrm_tax.xPay.Caption, "Fixed")
xtotalQuant.Caption = Myvalue(SalesFrm_tax.xtQuant.Caption, "Fixed")
CalcTotals
SalesFrm_tax.lSave = False
End Sub
Private Sub xPay_Change()
    CalcTotals
End Sub
Private Sub xPay_GotFocus()
    xPay.SelStart = 0
    xPay.SelLength = Len(xPay.Text)
End Sub
Private Sub xPay_KeyPress(KeyAscii As Integer)
    If Trim(xPay.Text) = "" Then Exit Sub
    If KeyAscii = 13 And IsNumeric(xPay.Text) Then
        If Val(xRest.Caption) >= 0 Then
            SalesFrm_tax.xPay = xPay.Text
            SalesFrm_tax.xRest.Caption = xRest.Caption
            SalesFrm_tax.lSave = True
            Unload Me
        Else
            xPay.Text = ""
        End If
    End If
End Sub
Private Sub xRet_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Exit Sub
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub
Private Sub CalcTotals()
    xRest.Caption = Format(Val(xPay.Text) + Val(xVisa.Text) - Val(xTotal.Caption), "#0.00")
End Sub
Private Sub xVisa_Change()
    CalcTotals
End Sub
Private Sub xVisa_GotFocus()
'    If Val(xTotal.Caption) - Val(xPay.Text) > 0 Then
'        xVisa.Text = Format((Val(xTotal.Caption) - Val(xPay.Text)), "#0.00")
'    End If
End Sub
Private Sub xVisa_KeyPress(KeyAscii As Integer)
    If Trim(xVisa.Text) = "" Then Exit Sub
    If KeyAscii = 13 And IsNumeric(xVisa.Text) Then
        If Val(xRest.Caption) >= 0 Then
            SalesFrm_tax.xVisa = xVisa.Text
            SalesFrm_tax.xRest.Caption = xRest.Caption
            SalesFrm_tax.lSave = True
            Unload Me
        Else
            xVisa.Text = ""
        End If
    End If
End Sub
