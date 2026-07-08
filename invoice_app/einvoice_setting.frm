VERSION 5.00
Begin VB.Form einvoice_settingfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "÷»ÿ «⁄œ«œ   ÿ»Ìﬁ «·„Õ„Ê·"
   ClientHeight    =   5595
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
   ScaleHeight     =   5595
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   2970
      Width           =   11445
      Begin VB.CheckBox xShowSign 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Show Signature"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1305
         Width           =   1860
      End
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
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         TabIndex        =   25
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
         TabIndex        =   24
         Top             =   585
         Width           =   2310
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
         Height          =   240
         Index           =   10
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   945
         Width           =   2310
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1320
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1665
      Width           =   11445
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
         TabIndex        =   6
         Top             =   900
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
         TabIndex        =   4
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
         TabIndex        =   5
         Top             =   540
         Width           =   8430
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
         TabIndex        =   21
         Top             =   945
         Width           =   1950
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
         TabIndex        =   18
         Top             =   585
         Width           =   1950
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
         TabIndex        =   17
         Top             =   225
         Width           =   2445
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1680
      Left            =   90
      TabIndex        =   13
      Top             =   0
      Width           =   11445
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
         TabIndex        =   20
         Top             =   1260
         Width           =   1455
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
         TabIndex        =   19
         Top             =   945
         Width           =   1455
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
         TabIndex        =   15
         Top             =   270
         Width           =   1455
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
         TabIndex        =   14
         Top             =   585
         Width           =   1275
      End
   End
   Begin VB.CommandButton CmdExit 
      Height          =   555
      Left            =   10170
      MaskColor       =   &H00FFFFFF&
      Picture         =   "einvoice_setting.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4860
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
      Picture         =   "einvoice_setting.frx":4AB6
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Õ›Ÿ"
      Top             =   4860
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
End
Attribute VB_Name = "einvoice_settingfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdSave_Click()
myreplace
End Sub
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "idSrvBaseUrl", addstring(xidSrvBaseUrl.text))
aInsert = AddFlag(aInsert, "TokenUrl", addstring(xTokenUrl.text))
aInsert = AddFlag(aInsert, "client_id", addstring(xclient_id.text))
aInsert = AddFlag(aInsert, "client_secret", addstring(xclient_secret.text))
aInsert = AddFlag(aInsert, "apiBaseUrl", addstring(xapiBaseUrl.text))
aInsert = AddFlag(aInsert, "SubmitUrl", addstring(xSubmitUrl.text))
aInsert = AddFlag(aInsert, "Version", addstring(xVersion.text))
aInsert = AddFlag(aInsert, "ActivityCode", addstring(xActivityCode.text))
aInsert = AddFlag(aInsert, "subType", addstring(xSubType.text))
aInsert = AddFlag(aInsert, "taxRate", mRound(xTaxRate.text, 0))
aInsert = AddFlag(aInsert, "ShowSign", xShowSign.Value)
On Error GoTo myerror
con.BeginTrans
If IsEmpty(myField("Select id from SettingInvoice", con)) Then
   con.Execute addInsert(aInsert, "SettingInvoice")
Else
   con.Execute addUpdate(aInsert, "SettingInvoice", "Id = 1")
End If
con.CommitTrans
Inform " „ «· ⁄œÌ·"
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con
myload
End Sub
Private Sub myload()
Dim loctable As New ADODB.Recordset
Set loctable = myCmd("Select * from SettingInvoice", con)
If Not loctable.EOF Then
    xidSrvBaseUrl.text = loctable!idSrvBaseUrl & ""
    xTokenUrl.text = loctable!TokenUrl & ""
    xclient_id.text = loctable!client_id & ""
    xclient_secret.text = loctable!client_secret & ""
    
    xapiBaseUrl.text = loctable!apiBaseUrl & ""
    xSubmitUrl.text = loctable!SubmitUrl & ""
    xVersion.text = loctable!Version & ""
    
    xActivityCode.text = loctable!ActivityCode & ""
    xSubType.text = loctable!subType & ""
    xTaxRate.text = Myvalue(loctable!taxRate)
    xShowSign.Value = IIf(loctable!ShowSign, 1, 0)
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
DefineUrl con
closeCon con
End Sub
Private Sub xapiBaseUrl_GotFocus()
myGotFocus xapiBaseUrl
End Sub
Private Sub xapiBaseUrl_LostFocus()
myLostFocus xapiBaseUrl
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



