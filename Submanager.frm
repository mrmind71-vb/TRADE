VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Submanager 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬂ·„… ”— «·’·«ÕÌ« "
   ClientHeight    =   1785
   ClientLeft      =   4050
   ClientTop       =   3825
   ClientWidth     =   5535
   ControlBox      =   0   'False
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   1785
   ScaleWidth      =   5535
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   600
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Submanager.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   " ‘€Ì·"
      Top             =   1170
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   5280
      Begin VB.TextBox xPass 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xUser 
         Height          =   360
         Left            =   135
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   225
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ·„… «·”— :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "≈”„ «·„” Œœ„ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   315
         Width           =   1305
      End
   End
   Begin Threed.SSCommand cmdApply 
      Height          =   600
      Left            =   1485
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1170
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1058
      _Version        =   196610
      PictureFrames   =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Submanager.frx":246C
      Caption         =   "œŒÊ·"
      PictureAlignment=   10
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Submanager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Form, sString As String, sFilter As String, sFlag As String
Dim con As New ADODB.Connection
Dim nTry As Integer, nTries As Integer
Private Sub CmdApply_Click()
    checkUser
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
nTries = 3
aUser = Empty
openCon con
Dim cString As String
cString = "SELECT * FROM USERS"
If sFilter <> "" Then cString = cString & turn(cString) & sFilter
data1.RecordSource = cString
data1.ConnectionString = strCon
Set xUser.RowSource = data1
xUser.ListField = "Desca"
xUser.BoundColumn = "Code"
data1.Refresh
If data1.Recordset.RecordCount = 1 Then
    xUser.BoundText = data1.Recordset!CODE
    xUser.Enabled = False
Else
    LoadText Me, , sFlag
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveText Me, , Array(xUser.Name), sFlag
closeCon con
Set Managerfrm = Nothing
End Sub
Private Sub checkUser()
If xPass.text = "" Then Exit Sub
Dim aRet
aUser = retUser()
If IsEmpty(aUser) Then
    Inform "ﬂ·„… ”— €Ì— ’ÕÌÕ…"
    nTry = nTry + 1
    If nTry < nTries Then Exit Sub
Else
    SalesFrm.sUserDisc = xUser.text
End If
Unload Me
End Sub

Private Sub xPass_GotFocus()
Clipboard.Clear
End Sub

Private Sub xPass_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    aUser = Empty
    Unload Me
ElseIf KeyCode = 13 And xPass.text <> "" And xUser.MatchedWithList Then
    KeyCode = 0
    checkUser
End If
End Sub
Private Function retUser() As Variant
Dim loctable As New ADODB.Recordset
If Trim(xPass.text) <> "" Then sString = sString & turn(sString) & "password = " & MyParn(xPass.text)
If xUser.MatchedWithList Then sString = sString & turn(sString) & "code = " & MyParn(xUser.BoundText)
loctable.Open sString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.EOF And loctable.BOF) Then
    Dim aRet()
    ReDim aRet(loctable.Fields.Count - 1)
    For i = 0 To UBound(aRet)
        aRet(i) = loctable.Fields(i).Value
    Next
    retUser = aRet
End If
loctable.Close
Set loctable = Nothing
End Function
Private Sub xPass_Change()
cmdApply.Enabled = xUser.MatchedWithList And Trim(xPass.text) <> ""
End Sub
Private Sub xPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Clipboard.Clear
End Sub
Private Sub xUser_Change()
    cmdApply.Enabled = xUser.MatchedWithList And Trim(xPass.text) <> ""
End Sub
Private Sub xUser_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If Trim(xPass.text) <> "" Then
        xPass_KeyUp KeyCode, Shift
    Else
        KeyCode = 0
        On Error Resume Next
        xPass.SetFocus
        Err.Clear
    End If
End If
End Sub
