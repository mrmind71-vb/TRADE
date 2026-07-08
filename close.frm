VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form closefrm 
   Caption         =   "√€·«ﬁ › —…"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2145
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar bar1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   7
      Top             =   1680
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "close.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1500
   End
   Begin VB.CommandButton cmdApply 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   1620
      Picture         =   "close.frx":246C
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   3075
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1860
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1860
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·Ï  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdApply 
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   1620
      Picture         =   "close.frx":487C
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "⁄—÷ «·»Ì«‰« "
      Top             =   1080
      Width           =   1590
   End
End
Attribute VB_Name = "closefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nMode As Integer
Public sFile As String, sFieldClose As String, sFieldDate As String, sCaption As String
Dim con As New ADODB.Connection

Private Sub CmdApply_Click(Index As Integer)
Dim nRecord As Long, nRecordAft
If MsgBox(IIf(nMode = 0, "≈€·«ﬁ „” ‰œ« ", "› Õ „” ‰œ« "), vbOKCancel) <> vbOK Then Exit Sub

cString = "SELECT COUNT(*) FROM " & sFile
If IsDate(xdate1.text) Then cString = cString & turn(cString) & sFieldDate & " >= " & DateSq(xdate1.text)
If IsDate(xDate2.text) Then cString = cString & turn(cString) & sFieldDate & " <= " & DateSq(xDate2.text)
nRecord = Val(GetDesca(cString, con))


cString = "UPDATE " & sFile & " SET " & sFieldClose & " = " & IIf(nMode = 0, 1, 0) & "  FROM " & sFile
If IsDate(xdate1.text) Then cString = cString & turn(cString) & sFieldDate & " >= " & DateSq(xdate1.text)
If IsDate(xDate2.text) Then cString = cString & turn(cString) & sFieldDate & " <= " & DateSq(xDate2.text)
con.BeginTrans
On Error GoTo myerror
con.Execute cString, nRercodAft
con.CommitTrans
If nRercodAft > 0 Then
    Inform " „ " & IIf(nMode = 0, "≈€·«ﬁ ", "› Õ ") & nRecord & " „” ‰œ" & turn(sCaption, " ") & sCaption
    bar1.Panels(1).text = "⁄œœ «·”Ã·«  " & IIf(nMode = 0, "«·„€·ﬁ… ", "«·„› ÊÕ… ") & nRecord
Else
    bar1.Panels(1).text = "·„ Ì „ " & IIf(nMode = 0, "≈€·«ﬁ", "› Õ") & " «Ì ”Ã·« "
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
If sFieldClose = "" Then sFieldClose = "[CLOSED]"
If sFieldDate = "" Then sFieldDate = "[DATE]"
If nMode = 0 Then
    If Not IsDate(xdate1.text) Then
        Dim cString As String
        cString = "Select min(" & sFieldDate & ") FROM " & sFile
        cString = cString & turn(cString) & sFieldClose & " = " & nMode
        xdate1.text = Format(GetDesca(cString, con), "DD-MM-YYYY")
        xDate2.text = Format(Date, "DD-MM-YYYY")
    End If
Else
    xdate1.text = Format(Date, "DD-MM-YYYY")
    xDate2.text = Format(Date, "DD-MM-YYYY")
End If
CmdApply(0).Visible = nMode = 0
CmdApply(1).Visible = nMode = 1

CmdApply(0).ToolTipText = IIf(nMode = 0, "≈€·«ﬁ «·„·›« ", "› Õ «·„·›« ")
Me.Caption = IIf(nMode = 0, "≈€·«ﬁ ", "› Õ ") & "„” ‰œ« " & turn(sCaption, " ", "") & sCaption
openCon con
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set closefrm = Nothing
End Sub

Private Sub xDate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xdate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
End Sub


Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xdate1
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xDate2
End Sub

