VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form show_model_photofrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ÕæÑ ÇáÕäÝ"
   ClientHeight    =   10290
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   13770
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   8505
      Width           =   3795
      Begin VB.CheckBox chkFit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fit Width"
         Height          =   195
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   315
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.ComboBox cboZoomBy 
         Height          =   315
         ItemData        =   "show_model_photo.frx":0000
         Left            =   1125
         List            =   "show_model_photo.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zoom By %"
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   315
         Width           =   930
      End
   End
   Begin ImgXCtrl6.ImgXCtrl ImgX1 
      Height          =   10095
      Left            =   45
      TabIndex        =   6
      Top             =   90
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   17806
      BackColor       =   16777215
      BorderStyle     =   4
      SelectionLineType=   4
      Center          =   -1  'True
      ImageBorderThickness=   1
      ScrollSmallChange=   0.1
      AntialiasDisplay=   -1  'True
      WMFPlaceable    =   0   'False
      SelectionEnableMove=   -1  'True
      AutoZoomType    =   2
      DoubleBuffer    =   -1  'True
      LicenseUserName =   "mrmind"
      LicenseRegCode  =   "’§Ò»»¥²³½­°Ò±²§«´©®¯OOHH-FAOOYNJB-EQCF6gI"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   17595
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   8505
      Width           =   2580
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "show_model_photo.frx":0004
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1305
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "show_model_photo.frx":2327
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "show_model_photo.frx":469D
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   -2175
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   4140
      Top             =   10530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   16425
      Top             =   9630
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin ImgXCtrl6.ImgXCtrl imgx2 
      Height          =   8565
      Left            =   -855
      TabIndex        =   2
      Top             =   9765
      Visible         =   0   'False
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   15108
      BackColor       =   16777215
      BorderStyle     =   4
      AutoZoom        =   -1  'True
      SelectionLineType=   4
      Center          =   -1  'True
      ImageBorderThickness=   1
      DoubleBuffer    =   -1  'True
      LicenseUserName =   "mrmind"
      LicenseRegCode  =   "’§Ò»»¥²³½­°Ò±²§«´©®¯OOHH-FAOOYNJB-EQCF6gI"
   End
   Begin VB.Label xPhoto 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   9630
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   8730
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "show_model_photofrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents twain As ImgXTwain
Attribute twain.VB_VarHelpID = -1
Public bedit As Boolean
Private WithEvents MyPrinter As ImgXPrint
Attribute MyPrinter.VB_VarHelpID = -1
Dim bIg As Boolean
Dim ImageTable As ADODB.Recordset
Dim fs As New FileSystemObject
Private Sub cboZoomBy_Click()
If cboZoomBy.text = "" Then
    If chkFit.Value = 1 Then
        chkFit_Click
    Else
        ImgX1.Zoom = 1
    End If
Else
    ImgX1.Zoom = Val(cboZoomBy.text) / 100
End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkFit_Click()
If Not bIg Then
    ImgX1.AutoZoom = chkFit.Value = 1
    If chkFit.Value = 0 Then ImgX1.Zoom = 1
    'LoadPhoto
End If
'ImgX1.Update
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set ImgXTwain = Nothing

'addText chkFit.Name, chkFit.Value
'addText cboZoomBy.Name, cboZoomBy.ListIndex

Set show_model_photofrm = Nothing
Err.Clear
End Sub

Private Sub Form_Load()
fixZoom

'chkFit.Value = Val(RetTxt(chkFit.Name))
cboZoomBy.ListIndex = 2
End Sub

Private Sub ImgX1_MouseWheelScroll(ByVal ScrollValue As Long, ByVal Keys As Long)
'If chkFit.Value = 1 Then chkFit.Value = 0
ImgX1.Zoom = ImgX1.Zoom + (ScrollValue / 100)
End Sub
Private Sub myloadgrd2()
'Dim aPrm As Variant
'aPrm = AddFlag(aPrm, "ID_CASH", sId_cash)
End Sub
Private Sub fixZoom()
LoadingCboZoomBy = True
With cboZoomBy
    .Clear
    .AddItem ""
    .AddItem "80"
    .AddItem "90"
    'Some fiddly code to ensure we always have 100% in the list in its
    'correct position even if our loop doesn't hit I = 100.  Also
    'record its index for use in Zoom100Percent.
    Zoom100PctCboZoomIndex = -1
    For I = 100 To 150 Step 10
        If I > 100 And Zoom100PctCboZoomIndex < 0 Then
            .AddItem I
            Zoom100PctCboZoomIndex = .NewIndex
        End If
        .AddItem CStr(I)
        If I = 100 Then Zoom100PctCboZoomIndex = .NewIndex
    Next
End With
End Sub
Function RetTxt(cSearch As String, Optional cFile As String) As String
Dim TextLine As String
On Error GoTo myerror

If cFile = "" Then
    cFile = App.Path & "\CONF.TXT"
End If

Dim FileNumber
FileNumber = FreeFile()

Open cFile For Input As #FileNumber   ' Open file.

Do While Not EOF(FileNumber)   ' Loop until end of file.
   Line Input #FileNumber, TextLine   ' Read line into variable.
   If InStr(FileNumber, LCase(TextLine), LCase(cSearch) & "=") > 0 Then
       RetTxt = Mid(TextLine, Len(cSearch) + 2)
       Exit Do
   End If
Loop

Close #FileNumber   ' Close file.
Exit Function
myerror:
Err.Clear
RetTxt = ""
End Function

Function addText(cField, cValue, Optional cFile As String) As Boolean
Dim TextLine As String, cText As String, aLocal, nFoundTimes As Integer
Dim fs As New FileSystemObject
If cFile = "" Then cFile = App.Path & "\CONF.TXT"
Dim FileNumber
On Error GoTo myerror
If fs.FileExists(cFile) Then
    FileNumber = FreeFile()
    Open cFile For Input As #FileNumber   ' Open file.
    Do While Not EOF(FileNumber)   ' Loop until end of file.
       Line Input #FileNumber, TextLine   ' Read line into variable.
       cText = cText & Tr(cText, vbCrLf) & TextLine
    Loop
    Close #FileNumber   ' Close file.
End If

aLocal = Split(cText, vbCrLf)

FileNumber = FreeFile()
Open cFile For Output As #FileNumber   ' Open file.
For I = 0 To UBound(aLocal)
    nFound = InStr(1, LCase(Trim(aLocal(I))), LCase(Trim(cField)) & "=")
    If nFound > 0 Then
        nFoundTimes = nFoundTimes + 1
        If nFoundTimes = 1 Then
            If Trim(cValue) <> "" Then
                Print #FileNumber, UCase(Trim(cField)) & "=" & Trim(cValue)
            End If
        End If
    Else
        Print #FileNumber, aLocal(I)
    End If
Next

If nFoundTimes = 0 Then
    If Trim(cValue) <> "" Then
        Print #FileNumber, Trim(UCase(cField)) & "=" & Trim(cValue)
    End If
End If
Close #FileNumber
addText = True
Exit Function
myerror:
Err.Clear
End Function

