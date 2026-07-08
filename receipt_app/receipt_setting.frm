VERSION 5.00
Begin VB.Form receipt_settingfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "÷»ÿ «⁄œ«œ   ÿ»Ìﬁ «·„Õ„Ê·"
   ClientHeight    =   6570
   ClientLeft      =   -15
   ClientTop       =   435
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   4410
      Width           =   11445
      Begin VB.TextBox xActivityCode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2925
         MaxLength       =   200
         TabIndex        =   11
         Tag             =   "t"
         Top             =   180
         Width           =   2940
      End
      Begin VB.TextBox xSubType 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2925
         MaxLength       =   200
         TabIndex        =   12
         Tag             =   "t"
         Top             =   540
         Width           =   2940
      End
      Begin VB.TextBox xTaxRate 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2925
         MaxLength       =   200
         TabIndex        =   13
         Tag             =   "t"
         Top             =   900
         Width           =   2940
      End
      Begin VB.Label Label6 
         Caption         =   "taxpayerActivityCode"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   225
         Width           =   2310
      End
      Begin VB.Label Label6 
         Caption         =   "Tax Subtype"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   585
         Width           =   1320
      End
      Begin VB.Label Label6 
         Caption         =   "Tax Rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   945
         Width           =   1275
      End
   End
   Begin VB.Frame Frame3 
      Height          =   2040
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   2385
      Width           =   11445
      Begin VB.TextBox xdate_ref 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   10
         Top             =   1620
         Width           =   8430
      End
      Begin VB.TextBox xQRUrl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   8
         Top             =   900
         Width           =   8430
      End
      Begin VB.TextBox xVersion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   9
         Top             =   1260
         Width           =   8430
      End
      Begin VB.TextBox xapiBaseUrl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   6
         Top             =   180
         Width           =   8430
      End
      Begin VB.TextBox xSubmitUrl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   7
         Top             =   540
         Width           =   8430
      End
      Begin VB.Label Label12 
         Caption         =   "Ignore Refrence Date from"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   32
         Top             =   1665
         Width           =   2625
      End
      Begin VB.Label Label9 
         Caption         =   "QRCode URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   31
         Top             =   945
         Width           =   1860
      End
      Begin VB.Label Label4 
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   24
         Top             =   1305
         Width           =   870
      End
      Begin VB.Label Label2 
         Caption         =   "API Submit Invoice"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   21
         Top             =   585
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "API Base URL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         TabIndex        =   20
         Top             =   225
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2400
      Left            =   90
      TabIndex        =   16
      Top             =   0
      Width           =   11445
      Begin VB.TextBox xPososVersion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   5
         Top             =   1980
         Width           =   8430
      End
      Begin VB.TextBox xPosSerial 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   4
         Top             =   1620
         Width           =   8430
      End
      Begin VB.TextBox xTokenUrl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   3
         Top             =   1260
         Width           =   8430
      End
      Begin VB.TextBox xidSrvBaseUrl 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   2
         Top             =   900
         Width           =   8430
      End
      Begin VB.TextBox xclient_id 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   0
         Top             =   180
         Width           =   8430
      End
      Begin VB.TextBox xclient_secret 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   1
         Top             =   540
         Width           =   8430
      End
      Begin VB.Label Label8 
         Caption         =   "Pos Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1980
         Width           =   1185
      End
      Begin VB.Label Label7 
         Caption         =   "Pos Serial"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Token Url"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1260
         Width           =   1005
      End
      Begin VB.Label Label10 
         Caption         =   "Id Base"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         TabIndex        =   22
         Top             =   945
         Width           =   780
      End
      Begin VB.Label Label11 
         Caption         =   "Client Key"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label5 
         Caption         =   "API Secret"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   585
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdExit 
      Height          =   555
      Left            =   10170
      MaskColor       =   &H00FFFFFF&
      Picture         =   "receipt_setting.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5805
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.CommandButton cmdSave 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   8685
      MaskColor       =   &H00FFFFFF&
      Picture         =   "receipt_setting.frx":4AB6
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Õ›Ÿ"
      Top             =   5805
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
End
Attribute VB_Name = "receipt_settingfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
If myreplace Then Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
End Sub
Private Function myreplace() As Boolean
On Error GoTo myerror
Dim xml As New ChilkatXml
Dim success As Long
success = xml.LoadXmlFile(tempPath & "\receipt.xml")
xml.Tag = "root"

xml.UpdateChildContent "setting|idSrvBaseUrl", xidSrvBaseUrl.text
xml.UpdateChildContent "setting|tokenUrl", xTokenUrl.text
xml.UpdateChildContent "setting|client_id", xclient_id
xml.UpdateChildContent "setting|client_secret", xclient_secret.text
xml.UpdateChildContent "setting|posSerial", xPosSerial.text
xml.UpdateChildContent "setting|pososVersion", xPososVersion.text
xml.UpdateChildContent "setting|apiBaseUrl", xapiBaseUrl.text
xml.UpdateChildContent "setting|submitUrl", xSubmitUrl.text
xml.UpdateChildContent "setting|QRUrl", xQRUrl.text
xml.UpdateChildContent "setting|date_ref", myFormat(xdate_ref.text)
xml.UpdateChildContent "setting|version", xVersion.text
xml.UpdateChildContent "setting|activityCode", xActivityCode.text
xml.UpdateChildContent "setting|subType", xSubType.text
xml.UpdateChildContent "setting|taxRate", xTaxRate.text
success = xml.SaveXml(tempPath & "\receipt.xml")
If success = 0 Then
    MsgBox xml.LastErrorXml
End If
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_Load()
MFocus Me
myload
End Sub
Private Sub myload()
Dim xml As New ChilkatXml

Dim success As Integer
success = xml.LoadXmlFile(tempPath & "\receipt.xml")
If success = 0 Then
    MsgBox xml.LastErrorXml
    Exit Sub
End If

xidSrvBaseUrl.text = xml.GetChildContent("setting|idSrvBaseUrl")
xTokenUrl.text = xml.GetChildContent("setting|tokenUrl")
xclient_id.text = xml.GetChildContent("setting|client_id")
xclient_secret.text = xml.GetChildContent("setting|client_secret")
xPosSerial.text = xml.GetChildContent("setting|posSerial")
xPososVersion.text = xml.GetChildContent("setting|pososVersion")
xdate_ref.text = Format(xml.GetChildContent("setting|date_ref"), "d/m/yyyy")
xapiBaseUrl.text = xml.GetChildContent("setting|apiBaseUrl")
xSubmitUrl.text = xml.GetChildContent("setting|submitUrl")
xQRUrl.text = xml.GetChildContent("setting|QRUrl")
xVersion.text = xml.GetChildContent("setting|version")
xActivityCode.text = xml.GetChildContent("setting|activityCode")
xSubType.text = xml.GetChildContent("setting|subType")
xTaxRate.text = xml.GetChildContent("setting|taxRate")
End Sub
Private Sub Form_Unload(Cancel As Integer)
'DefineUrl con
'closeCon con
End Sub
Private Sub xapiBaseUrl_GotFocus()
myGotFocus xapiBaseUrl
End Sub
Private Sub xapiBaseUrl_LostFocus()
myLostFocus xapiBaseUrl
End Sub

Private Sub xPososVersion_GotFocus()
myGotFocus xPososVersion
End Sub
Private Sub xPososVersion_LostFocus()
myLostFocus xPososVersion
End Sub
Private Sub xPosSerial_GotFocus()
myGotFocus xPosSerial
End Sub
Private Sub xPosSerial_LostFocus()
myLostFocus xPosSerial
End Sub
Private Sub xSubmitUrl_GotFocus()
myGotFocus xSubmitUrl
End Sub
Private Sub xSubmitUrl_LostFocus()
myLostFocus xSubmitUrl
End Sub
Private Sub xTokenUrl_GotFocus()
myGotFocus xTokenUrl
End Sub
Private Sub xTokenUrl_LostFocus()
myLostFocus xTokenUrl
End Sub
Private Sub xidSrvBaseUrl_GotFocus()
myGotFocus xidSrvBaseUrl
End Sub
Private Sub xidSrvBaseUrl_LostFocus()
myLostFocus xidSrvBaseUrl
End Sub
Private Sub xclient_id_GotFocus()
myGotFocus xclient_id
End Sub
Private Sub xclient_id_LostFocus()
myLostFocus xclient_id
End Sub
Private Sub xclient_secret_GotFocus()
myGotFocus xclient_secret
End Sub
Private Sub xclient_secret_LostFocus()
myLostFocus xclient_secret
End Sub
Private Sub xVersion_GotFocus()
myGotFocus xVersion
End Sub
Private Sub xVersion_LostFocus()
myLostFocus xVersion
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyAscii = 0
    End If
ElseIf KeyAscii = 19 And cmdSave.Enabled Then
    cmdSave_Click
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then
        KeyCode = 0
        SendKeys "{TAB}"
    End If
End If
End Sub
Private Sub xSubType_GotFocus()
myGotFocus xSubType
End Sub
Private Sub xSubType_LostFocus()
myLostFocus xSubType
End Sub
Private Sub xActivityCode_GotFocus()
myGotFocus xActivityCode
End Sub
Private Sub xActivityCode_LostFocus()
myLostFocus xActivityCode
End Sub
Private Sub xTaxRate_GotFocus()
myGotFocus xTaxRate
End Sub
Private Sub xTaxRate_LostFocus()
myLostFocus xTaxRate
End Sub
Private Sub xdate_ref_GotFocus()
myGotFocus xdate_ref
End Sub
Private Sub xdate_ref_LostFocus()
myLostFocus xdate_ref
myValidDate_en xdate_ref
End Sub
Private Sub xQRUrl_GotFocus()
myGotFocus xQRUrl
End Sub
Private Sub xQRUrl_LostFocus()
myLostFocus xQRUrl
End Sub
