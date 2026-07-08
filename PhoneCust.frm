VERSION 5.00
Begin VB.Form PhoneCust 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·⁄„·«¡"
   ClientHeight    =   1620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8565
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1620
   ScaleWidth      =   8565
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox XPHONE 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4560
      MaxLength       =   11
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   2190
   End
   Begin VB.TextBox XEMAIL 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   990
      Width           =   6210
   End
   Begin VB.TextBox xDescA 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   540
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   555
      Width           =   6210
   End
   Begin VB.Label xDescA2 
      Alignment       =   1  'Right Justify
      Caption         =   "Label1"
      Height          =   240
      Left            =   585
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "—ﬁ„ «· ·Ì›Ê‰ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   6885
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   210
      Width           =   1200
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "«·«”„ :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   6885
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   630
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "E MAIL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   6885
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1050
      Width           =   645
   End
End
Attribute VB_Name = "PhoneCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public lAdd As Boolean
Dim SubCustTable As New ADODB.Recordset
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
    openCon con
    XPHONE.text = SalesFrm.XPHONE.Caption
    If XPHONE.text <> "" Then
'        xDesca.text = GetDesca("SELECT DESCA FROM SUBCUST WHERE PHONE = " & MyParn(xphone.text), con)
        xDescA.text = Ret_Phone(XPHONE.text, con)
        xDescA2.Caption = xDescA.text
    End If
End Sub
Private Sub XPHONE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Len(XPHONE.text) = 11 Then
'        xDesca.text = GetDesca("SELECT DESCA FROM SUBCUST WHERE PHONE = " & MyParn(xphone.text), con)
        xDescA.text = Ret_Phone(XPHONE.text, con)
        xDescA2.Caption = xDescA.text
        On Error Resume Next
        xDescA.SetFocus
        Err.Clear
    End If
End Sub
Private Sub xDesca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(XPHONE.text) = 11 Then
'            aRet = aGetDesca("SELECT DESCA , BRANCH FROM SUBCUST WHERE PHONE = " & MyParn(XPHONE.text), con)
            If xDescA2.Caption = "" Then
                con.Execute " INSERT INTO [SUBCUST] (desca,phone,f_date,branch,isnew) VALUES ( " & addstring(xDescA.text) & " , " & addstring(XPHONE.text) & " , " & addDate(Date) & " , " & addstring(cBranch) & " , 1 ) "
            Else
                con.Execute " update [SUBCUST] set desca = " & addstring(xDescA.text) & ", isnew = 1 where phone = " & MyParn(XPHONE.text)
            End If
            con.Execute " update file6_20h set phone = " & addstring(XPHONE.text) & " where doc_no = " & MyParn(SalesFrm.xDoc_No.text)
            SalesFrm.XPHONE.Caption = XPHONE.text
            Unload Me
        End If
    End If
End Sub
