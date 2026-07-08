VERSION 5.00
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form copyFlashfrm 
   Caption         =   "⁄„· ‰”Œ… «Õ Ì«ÿÌ…"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4665
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2895
   ScaleWidth      =   4665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRestoreFrom 
      Caption         =   "«” —Ã«⁄ «·‰”Œ…"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   90
      TabIndex        =   4
      Top             =   855
      Visible         =   0   'False
      Width           =   4515
   End
   Begin VB.Frame Frame3 
      Caption         =   "Drive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2250
      Width           =   1050
      Begin VB.TextBox xDrive 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdExit 
      Height          =   690
      Left            =   90
      Picture         =   "copyFlash.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1530
      Width           =   4515
   End
   Begin Threed.SSCommand cmdCopyTo 
      Height          =   735
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   1296
      _Version        =   196610
      ForeColor       =   0
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "copyFlash.frx":246C
      Caption         =   "⁄„· ‰”Œ… «Õ Ì«ÿÌ…"
      ButtonStyle     =   1
      PictureAlignment=   1
      BevelWidth      =   10
      ShapeSize       =   1
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "copyFlashfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bFirst As Boolean
Private Sub cmdCopyto_Click()
    On Error GoTo myerror
    Me.MousePointer = 11
    If copyToFlash Then MsgBox " „ ⁄„· ‰”Œ… «Õ Ì«ÿÌ… »‰Ã«Õ", , "⁄„· ‰”Œ… «Õ Ì«ÿÌ…"
lastsub:
    Me.MousePointer = 1
    Exit Sub
myerror:
    MsgBox Err.Description, , "Error description"
    Err.Clear
    Me.MousePointer = 1
    GoTo lastsub
End Sub
Private Function copyToFlash() As Boolean
'On erorr GoTo myerror
Dim fs As FileSystemObject, cDir As String, cFileName As String, aDrive As Variant
Set fs = CreateObject("Scripting.FileSystemObject")
If Trim(xDrive.text) = "" Then
    aDrive = aLastDrive(False)
    cDir = retFlag(aDrive, "LETTER") & ":\DataBackup"
Else
    aDrive = aLastDrive(False)
    cDir = xDrive.text & ":\DataBackup"
End If
MyCreateFolder cDir
'FixFiles (cDir)
'If Val(aDrive(5)) = 1 Then
'    cFileName = retFlag(aDrive, "LETTER") & ":\DataBackup\shop_bak.bak"
'Else
    cFileName = cDir & "\" & sCatalog & "_" & Format(Date, "yyyymmdd") & ".bak"
'End If
Inform App.Path & "\DataBackup\back_data.bak"
Inform cFileName

If createBackUp(App.Path & "\DataBackup\back_data.bak") Then DoBackUp = True
fs.CopyFile App.Path & "\DataBackup\back_data.bak", cFileName

Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdRestoreFrom_Click()
Set fs = CreateObject("Scripting.FileSystemObject")
Dim cFile As String
Dim cNewFile As String
On Error GoTo myerror
    Common1.FileName = ""
    Common1.InitDir = App.Path
    Common1.Filter = "*.BAK"
    Common1.ShowOpen
    If Common1.FileTitle <> "" Then
        cFile = Common1.FileName
        If cFile <> "" Then
            If RestoreFromFlash(cFile) Then Inform " „ «” —Ã«⁄ «·»Ì«‰«  »‰Ã«Õ"
             
        End If
        
    End If
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear

End Sub
Private Function FixFiles(pDir As String) As Boolean
Dim fs As New FileSystemObject
Dim aRet As Variant, nDelete As Long
On Error Resume Next
aRet = retFArray(pDir, "bak")
nDelete = (UBound(aRet) + 1) - nMaxFiles
For i = 0 To (nDelete)
    fs.DeleteFile pDir & "\" & aRet(i)
Next
Err.Clear
End Function
Private Function createBackUp(pFileName) As Boolean
Dim cFile As String
Dim con As New ADODB.Connection
openCon con

Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText

Set FS1.ActiveConnection = con
cString = "BACKUP DATABASE " & sCatalog & " TO  DISK = N'" & pFileName & "' WITH  RETAINDAYS = 1, NOFORMAT, INIT,  NAME = N'over-Full Database Backup', SKIP,  NOREWIND, NOUNLOAD, STATS = 10"
FS1.CommandText = cString

FS1.Execute
Set FS1 = Nothing
closeCon con
createBackUp = True
End Function
Private Function RestoreFromFlash(pFile) As Boolean
On erorr GoTo myerror
Dim fs As FileSystemObject, cDir As String, cFileName As String, cDrive As String, cLastFile As String
Set fs = CreateObject("Scripting.FileSystemObject")

    cFileName = pFile
    MsgBox cFileName
    If Not fs.FileExists(cFileName) Then
        MsgBox "·« ÌÊÃœ „·› ·«” —Ã«⁄ «·»Ì«‰«  „‰Â"
        End
    End If

    If RestoreBackUp(cFileName) Then RestoreFromFlash = True

Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function LastFile(pDir As String) As String
Dim fs As New FileSystemObject
Dim aRet As Variant
On Error Resume Next
aRet = retFArray(pDir, "bak")
If UBound(aRet) > 0 Then
    LastFile = aRet(UBound(aRet)) & ""
End If
Err.Clear
End Function
Private Function RestoreBackUp(pFileName) As Boolean
Dim conMaster As New ADODB.Connection
Dim cString As String, cServerName As String
cServerName = RetSetting("server", App.Path & "\conf.txt")
If cServerName = "" Then cServerName = "."
'cServerName = MyParn(cServerName & turn(cExpress, "\") & cExpress)
cString = "provider=SQLOLEDB;data source= " & cServerName & "  ;initial " _
        & "catalog=master;Trusted_Connection=yes"
conMaster.Open cString

Dim FS1 As New ADODB.Command
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = conMaster
cString = "alter database  [" & sCatalog & "] set offline with rollback immediate"
cString = cString & turn(cString, vbCrLf) & "alter database [" & sCatalog & "]"
cString = cString & turn(cString, vbCrLf) & "set online"
'MsgBox cString
FS1.CommandText = cString
FS1.Execute

cString = "RESTORE DATABASE [" & sCatalog & "] FROM  DISK = N'" & pFileName & "' WITH  FILE = 1,  NOUNLOAD,  REPLACE,  STATS = 10"
'MsgBox cString
FS1.CommandText = cString
FS1.Execute
Set FS1 = Nothing
closeCon conMaster
RestoreBackUp = True
End Function
Private Sub Form_Load()
If UCase(RetSetting("COPY", App.Path & "\COPY.txt")) = "COPY" Then
    cmdCopyTo.Visible = False
    cmdRestoreFrom.Visible = True
Else
    cmdCopyTo.Visible = True
    cmdRestoreFrom.Visible = False
End If
aDrive = aLastDrive(False)

xDrive.text = retFlag(aDrive, "LETTER")
If cCopyDrive <> "" Then xDrive.text = cCopyDrive
End Sub
