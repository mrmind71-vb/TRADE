VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form TestDisplay 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   465
      Left            =   3150
      TabIndex        =   4
      Top             =   1935
      Width           =   1635
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "TestDisplay.frx":0000
      Left            =   0
      List            =   "TestDisplay.frx":0037
      TabIndex        =   3
      Text            =   "COM3"
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CLEAR"
      Height          =   600
      Left            =   1080
      TabIndex        =   2
      Top             =   1890
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      Height          =   510
      Left            =   1080
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1305
      Width           =   3570
   End
   Begin VB.TextBox Text1 
      Height          =   600
      Left            =   1080
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   450
      Width           =   3570
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   270
      Top             =   495
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      Handshaking     =   2
      RTSEnable       =   -1  'True
   End
End
Attribute VB_Name = "TestDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OSSTab1 As Integer
Dim IntrCh23$, IntrCh24$, IntrCh40$, IntrCh5B$, IntrCh5C$, IntrCh5D$, IntrCh5E$
Dim IntrCh60$, IntrCh7B$, IntrCh7C$, IntrCh7D$, IntrCh7E$
Private Sub Command1_Click()
    opencom
    ClearDisplay

End Sub
Sub opencom()
Dim i As Integer
Dim comFlag As Boolean
  comFlag = False
    mOpenUSBpd
  IntrCh23$ = Chr$(&H23)
  IntrCh24$ = Chr$(&H24)
  IntrCh40$ = Chr$(&H40)
  IntrCh5B$ = Chr$(&H5B)
  IntrCh5C$ = Chr$(&H5C)
  IntrCh5D$ = Chr$(&H5D)
  IntrCh5E$ = Chr$(&H5E)
  IntrCh60$ = Chr$(&H60)
  IntrCh7B$ = Chr$(&H7B)
  IntrCh7C$ = Chr$(&H7C)
  IntrCh7D$ = Chr$(&H7D)
  IntrCh7E$ = Chr$(&H7E)
End Sub
Sub ClearDisplay()
    Dim i, Pagei, nn As Integer
    Dim ss As String
    Dim TmpTime$
    Dim TmpMM, TmpHH, TmpSS
'    lstReturn.SetFocus
'    lstReturn.AddItem "**" + Str(Index + 1) + ". State1: " & Hex(mPdState())
    DoEvents
    Call Epson_Init
'    LstTest.AddItem (CmdEpson(Index).Caption)
'    LstTest.AddItem ("CLEAR LINE1")
'    LstTest.AddItem ("and   LINE2")
    ss = "CLEAR LINE1" + Chr$(&HA)
    ss = ss + "and LINE2"
'    lstReturn.AddItem Hex(mWritePD(ss, Len(ss)))
    DoEvents
    Sleep (1000)
    ss = Chr$(&HC)
'    lstReturn.AddItem Hex(mWritePD(ss, Len(ss)))
    DoEvents

End Sub

Private Sub Epson_Init()
Dim ss As String
    ss = Chr$(&H1B) + Chr$(&H40)
    DoEvents
    mWritePD ss, Len(ss)
    DoEvents
End Sub
Function mWritePD(mdata As String, ByVal mlength As Long) As Long
    If Combo1.Text = "USB" Then
       mWritePD = WritePD(mdata, mlength)
    ElseIf Left(Combo1.Text, 3) = "COM" Then
       MSComm1.Output = mdata
    End If
End Function
Function mOpenUSBpd() As Long
    MSComm1.CommPort = Mid$(CCOMPP  Combo1.Text, 4, 2)
    MSComm1.PortOpen = True
End Function
Private Sub Command2_Click()
    Dim ss As String
    ss = "Cursor Off."
    ss = ss + Chr$(&H1F) + Chr$(&H43) + Chr$(&H0)
    mWritePD ss, Len(ss)
    DoEvents
    ss = vbCrLf + "Cursor ON."
    ss = ss + Chr$(&H1F) + Chr$(&H43) + Chr$(&H1)
    mWritePD ss, Len(ss)
    DoEvents
End Sub
