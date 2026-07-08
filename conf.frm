VERSION 5.00
Begin VB.Form confFrm 
   Caption         =   "Connection Configuration"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2625
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1860
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   6450
      Begin VB.TextBox XSERVERPICTNEW 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   14
         Top             =   1395
         Width           =   4200
      End
      Begin VB.TextBox xpasswordNew 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   990
         Width           =   4200
      End
      Begin VB.TextBox xuseidNew 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   585
         Width           =   4200
      End
      Begin VB.TextBox xServerName 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   180
         Width           =   4200
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Server Pict :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   405
         TabIndex        =   13
         Top             =   1395
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Password :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   405
         TabIndex        =   9
         Top             =   990
         Width           =   1545
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "User Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   405
         TabIndex        =   8
         Top             =   630
         Width           =   1545
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Server Name :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         TabIndex        =   7
         Top             =   225
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1980
      Width           =   3300
      Begin VB.CommandButton Command2 
         Caption         =   "Exit"
         Height          =   420
         Left            =   1665
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   135
         Width           =   1590
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Save"
         Height          =   420
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   135
         Width           =   1590
      End
   End
   Begin VB.Label XSERVERPICT 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   1665
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1890
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label xPassword 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   855
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1890
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label xuserid 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1890
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "confFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConFileName As String
Public bCheck As Boolean
Private Sub Command1_Click()
    If InputBox("", "") <> "654321" Then Exit Sub
    If Not MYVALID Then Exit Sub
    addSetting "Server", xServerName.text
    If Trim(Me.xuseidNew.text) <> "" Then
        addSetting "userid", crypt(Me.xuseidNew.text, "dr")
    End If
    
    If Trim(xpasswordNew.text) <> "" Then
        addSetting "password", crypt(xpasswordNew.text, "dr")
    End If
    
    If Trim(XSERVERPICTNEW.text) <> "" Then
        addSetting "SERVERPICT", crypt(XSERVERPICTNEW.text, "dr")
    End If
    Inform " OK "
    End
    Unload Me
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    ConFileName = App.Path & "\conf.txt"
    myload
End Sub
Private Sub myload()
    xServerName.text = RetSetting("server", ConFileName)
    xuserid.Caption = RetSetting("userid", ConFileName)
    xPassword.Caption = RetSetting("password", ConFileName)
'    XSERVERPICT.Caption = RetSetting("SERVERPICT", ConFileName)
End Sub
Private Function MYVALID() As Boolean
If Trim(xServerName.text) = "" Then
    MsgBox "«”„ «·”Ì—›Ì— €Ì— „ÊÃÊœ"
End If

If Trim(xuseidNew.text) = "" Then
    MsgBox "«”„ «·”Ì—›Ì— €Ì— „ÊÃÊœ"
End If

If Trim(xpasswordNew.text) = "" Then
    MsgBox "ﬂ·„… «·”— €Ì— „ÊÃÊœ"
End If

'If cBranch = "00" Then
'    If Trim(XSERVERPICTNEW.Text) = "" Then
'        MsgBox "”Ì—›— «·’Ê— €Ì— „ÊÃÊœ"
'    End If
'End If
'
If bCheck Then
'    Dim con As New ADODB.Connection
'    cServerName = MyParn(cServerName & turn(cExpress, "\") & cExpress)
'    cServerNamePICT = MyParn(cServerName & turn(cExpress, "\") & cExpress)
'    cString = "provider=SQLOLEDB;data source = " & cServerName & ";initial " _
'            & "catalog=" & sCatalog & ";user id = " & xuseidNew.Text & ";" & "password = " & xpasswordNew.Text
'    On Error GoTo myerror
'    con.Open cString
'    closeCon con
End If
MYVALID = True
Set con = Nothing
Exit Function
myerror:
MsgBox Err.Description
End Function

