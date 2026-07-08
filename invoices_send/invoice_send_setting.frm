VERSION 5.00
Begin VB.Form invoice_send_settingfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "÷»ÿ «⁄œ«œ   ÿ»Ìﬁ «·„Õ„Ê·"
   ClientHeight    =   2745
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
   ScaleHeight     =   2745
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2085
      Left            =   90
      TabIndex        =   7
      Top             =   -45
      Width           =   11445
      Begin VB.TextBox xMissingUrl 
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
      Begin VB.TextBox xOwnerName 
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
      Begin VB.TextBox xTenantName 
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
      Begin VB.TextBox xBaseUrl 
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
      Begin VB.TextBox xAgentToken 
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
      Begin VB.Label Label2 
         Caption         =   "Missing URL"
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
         TabIndex        =   12
         Top             =   1665
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Owner Name"
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
         TabIndex        =   11
         Top             =   1305
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Tenant Name"
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
         TabIndex        =   10
         Top             =   945
         Width           =   1455
      End
      Begin VB.Label Label11 
         Caption         =   "Base URL"
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
         TabIndex        =   9
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Agent Token"
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
         TabIndex        =   8
         Top             =   585
         Width           =   1275
      End
   End
   Begin VB.CommandButton CmdExit 
      Height          =   555
      Left            =   10170
      MaskColor       =   &H00FFFFFF&
      Picture         =   "invoice_send_setting.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2070
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
      Picture         =   "invoice_send_setting.frx":4AB6
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Õ›Ÿ"
      Top             =   2070
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
End
Attribute VB_Name = "invoice_send_settingfrm"
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
aInsert = AddFlag(Empty, "baseUrl", addstring(xBaseUrl.text))
aInsert = AddFlag(aInsert, "agentToken", addstring(xAgentToken.text))
aInsert = AddFlag(aInsert, "tenantName", addstring(xTenantName.text))
aInsert = AddFlag(aInsert, "ownerName", addstring(xOwnerName.text))
aInsert = AddFlag(aInsert, "MissingUrl", addstring(xMissingUrl.text))
On Error GoTo myerror
con.BeginTrans
If IsEmpty(myField("Select id from SettingInvoiceSend", con)) Then
   con.Execute addInsert(aInsert, "SettingInvoiceSend")
Else
   con.Execute addUpdate(aInsert, "SettingInvoiceSend", "Id = 1")
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
Set loctable = myCmd("Select * from SettingInvoiceSend", con)
If Not loctable.EOF Then
    xBaseUrl.text = loctable!baseUrl & ""
    xAgentToken.text = loctable!agentToken & ""
    xTenantName.text = loctable!tenantName & ""
    xOwnerName.text = loctable!ownerName & ""
    xMissingUrl.text = loctable!missingUrl & ""
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
DefineUrl con
closeCon con
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
