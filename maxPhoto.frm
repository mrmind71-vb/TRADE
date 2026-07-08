VERSION 5.00
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form maxPhoto 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9780
   DrawStyle       =   5  'Transparent
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   9780
   StartUpPosition =   1  'CenterOwner
   Begin ImgXCtrl6.ImgXCtrl photo1 
      Height          =   7200
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   12700
      BackColor       =   16777215
      BorderStyle     =   0
      AutoSize        =   -1  'True
      LicenseUserName =   "mrvb71"
      LicenseRegCode  =   "挧一Д布东报构东掸ZQEH-AOZOOOZT-EFLF6gI"
   End
   Begin VB.Label xdesca 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   9600
   End
End
Attribute VB_Name = "maxPhoto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cFile As String, cDesca As String
Private Sub Form_Keyup(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub
Sub myapplyskin(pSkin)
xdesca.Width = Frame1.Width - 200
End Sub
Private Sub Form_Load()
nwidth = photo1.Width
nHeight = photo1.Height
'On Error Resume Next
photo1.Import.FromFile cFile
photo1.Width = nwidth
photo1.Height = nHeight
Me.photo1.AutoZoomType = ixazBestFit
Me.photo1.AutoSize = False
Me.photo1.AutoZoom = True
Me.Width = photo1.Left + nwidth + photo1.Left + 100
Me.Height = photo1.Height + photo1.Top + 700
xdesca.Width = photo1.Width
xdesca.Caption = cDesca
Err.Clear
End Sub
