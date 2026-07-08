VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FrmOpenMosm 
   Caption         =   "„Ê”„ ÃœÌœ"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5700
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar xBar 
      Height          =   285
      Left            =   180
      TabIndex        =   7
      Top             =   2115
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.TextBox xDesca 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   135
      Width           =   3555
   End
   Begin VB.TextBox xMosm 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2205
      MaxLength       =   3
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   615
      Width           =   1530
   End
   Begin VB.TextBox xDate 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1095
      Width           =   1530
   End
   Begin Threed.SSCommand CMD_SAVE 
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   1575
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   873
      _Version        =   196610
      Font3D          =   5
      ForeColor       =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "› Õ „Ê”„ ÃœÌœ ·ﬂ·  «·„Ê—œÌ‰"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "≈”„ «·„Ê”„"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   180
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ √Ê· «·„œ… "
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„Ê”„"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3915
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   660
      Width           =   405
   End
End
Attribute VB_Name = "FrmOpenMosm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CMD_SAVE_Click()
    Dim SuppTable As New ADODB.Recordset, nCount As Double
    XBAR.Max = Val(GetDesca("SELECT COUNT(CODE) FROM FILE4_10", con) & "")
    XBAR.Min = 0
    XBAR.Value = 0
    If Not (xMosm.text <> "" And xDesca.text <> "" And IsDate(xdate.text)) Then
        MsgBox " —«Ã⁄ »Ì«‰«  «·„Ê”„ «·ÃœÌœ "
        Exit Sub
    End If
    xMosm.text = UCase(xMosm.text)
    If MsgBox("”Ê› Ì „ › Õ „Ê”„ ÃœÌœ ·ﬂ· «·„Ê—œÌ‰ " & xMosm.text, vbOKCancel) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        
        If GetDesca("SELECT MOSM FROM MOSM WHERE MOSM = " & MyParn(xMosm.text), con) = "" Then
            con.Execute " INSERT INTO MOSM (MOSM,DESCA,DATE) " & _
                   "Values(" & _
                   addstring(xMosm.text) & "," & _
                   addstring(xDesca.text) & "," & _
                   addDate(xdate.text) & _
                   ")"
        Else
            con.Execute " UPDATE MOSM SET DESCA = " & addstring(xDesca.text) & " ,DATE = " & addDate(xdate.text) & " WHERE MOSM = " & MyParn(xMosm.text)
        End If
        
        cStr1 = "SELECT * FROM FILE4_10 "
        SuppTable.Open "FILE4_10", con, adOpenStatic, adLockReadOnly, adCmdTable
        With SuppTable
            Do While Not .EOF
                nCount = nCount + 1
                XBAR.Value = nCount
                If GetDesca("SELECT CODE FROM FILE4_22 WHERE MOSM = " & MyParn(xMosm.text) & " AND CODE = " & MyParn(!CODE), con) = "" Then
                    con.Execute " INSERT INTO FILE4_22 (MOSM,CODE,DATE) " & _
                           "Values(" & _
                           addstring(xMosm.text) & "," & _
                           addstring(!CODE) & "," & _
                           addDate(xdate.text) & _
                           ")"
                End If
                .MoveNext
            Loop
        End With
    End If
    con.CommitTrans
    Unload Me
    Exit Sub
myerror:
    con.RollbackTrans
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub Form_Load()
    con.CursorLocation = adUseClient
    con.Open strCon
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub xMosm_Validate(Cancel As Boolean)
    If Len(xMosm.text) <> 3 Then
        MsgBox "«·„Ê”„ €Ì— ’ÕÌÕ"
        Cancel = True
        Exit Sub
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo) Then SendKeys "{tAB}"
End Sub
