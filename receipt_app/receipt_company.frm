VERSION 5.00
Begin VB.Form receipt_companyfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "÷»ÿ «⁄œ«œ   ÿ»Ìﬁ «·„Õ„Ê·"
   ClientHeight    =   4410
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
   ScaleHeight     =   4410
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   3840
      Left            =   90
      TabIndex        =   12
      Top             =   0
      Width           =   11445
      Begin VB.TextBox xBuildingNumber 
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
         Height          =   330
         Left            =   2880
         MaxLength       =   100
         TabIndex        =   9
         Top             =   3420
         Width           =   8430
      End
      Begin VB.TextBox xStreet 
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
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   8
         Top             =   3060
         Width           =   8430
      End
      Begin VB.TextBox xRegionCity 
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
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   7
         Top             =   2700
         Width           =   8430
      End
      Begin VB.TextBox xGovernate 
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
         Height          =   330
         Left            =   2880
         MaxLength       =   100
         TabIndex        =   6
         Top             =   2340
         Width           =   8430
      End
      Begin VB.TextBox xCountry 
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
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "EG"
         Top             =   1980
         Width           =   8430
      End
      Begin VB.TextBox xActivityCode 
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
         Height          =   330
         Left            =   2880
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1620
         Width           =   8430
      End
      Begin VB.TextBox xDeviceSerialNumber 
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
         Height          =   330
         Left            =   2880
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1260
         Width           =   8430
      End
      Begin VB.TextBox xBranchCode 
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
         Height          =   330
         Left            =   2880
         MaxLength       =   50
         TabIndex        =   2
         Top             =   900
         Width           =   8430
      End
      Begin VB.TextBox xCompanyTradeName 
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
         Height          =   330
         Left            =   2880
         MaxLength       =   200
         TabIndex        =   1
         Top             =   540
         Width           =   8430
      End
      Begin VB.TextBox xRin 
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
         Height          =   330
         Left            =   2880
         MaxLength       =   30
         TabIndex        =   0
         Top             =   180
         Width           =   8430
      End
      Begin VB.Label Label6 
         Caption         =   "buildingNumber"
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
         Index           =   4
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   3465
         Width           =   2310
      End
      Begin VB.Label Label6 
         Caption         =   "street"
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
         Index           =   3
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   3105
         Width           =   1500
      End
      Begin VB.Label Label6 
         Caption         =   "regionCity"
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
         Index           =   2
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2745
         Width           =   1050
      End
      Begin VB.Label Label6 
         Caption         =   "Governate"
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
         Index           =   1
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2385
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Country"
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
         Index           =   0
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2025
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "ActivityCode"
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
         TabIndex        =   17
         Top             =   1665
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "DeviceSerialNumber"
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
         TabIndex        =   16
         Top             =   1305
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "BranchCode"
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
         Top             =   945
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "CompanyTradeName"
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
         TabIndex        =   14
         Top             =   585
         Width           =   2310
      End
      Begin VB.Label Label5 
         Caption         =   "Rin"
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
         TabIndex        =   13
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.CommandButton CmdExit 
      Height          =   465
      Left            =   10170
      MaskColor       =   &H00FFFFFF&
      Picture         =   "receipt_company.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3870
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
      Height          =   465
      Left            =   8685
      MaskColor       =   &H00FFFFFF&
      Picture         =   "receipt_company.frx":4AB6
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Õ›Ÿ"
      Top             =   3870
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
End
Attribute VB_Name = "receipt_companyfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If myreplace Then Inform " „ «· ⁄œÌ· »‰Ã«Õ"
End Sub
Private Function MYVALID() As Boolean
If Trim(xRin.text) = "" Then
    MsgBox "Registeration Number Required"
    Exit Function
End If

If Trim(xCompanyTradeName.text) = "" Then
    MsgBox "Company Trade Name Required"
    Exit Function
End If

If Trim(xBranchCode.text) = "" Then
    MsgBox "Branch Code Required"
    Exit Function
End If

If Trim(xDeviceSerialNumber.text) = "" Then
    MsgBox "Device Serial Number Required"
    Exit Function
End If

If Trim(xActivityCode.text) = "" Then
    MsgBox "Activity Code Required"
    Exit Function
End If

If Trim(xCountry.text) = "" Then
    MsgBox "Country Required"
    Exit Function
End If

If Trim(xGovernate.text) = "" Then
    MsgBox "Governate required"
    Exit Function
End If

If Trim(xRegionCity.text) = "" Then
    MsgBox "Governate required"
    Exit Function
End If

If Trim(xStreet.text) = "" Then
    MsgBox "Street required"
    Exit Function
End If

If Trim(xBuildingNumber.text) = "" Then
    MsgBox "Building Number required"
    Exit Function
End If
MYVALID = True
End Function
Private Function myreplace() As Boolean
Dim xml As New ChilkatXml
Dim success As Long
success = xml.LoadXmlFile(tempPath & "\receipt.xml")
xml.Tag = "root"

xml.UpdateChildContent "company|rin", xRin.text
xml.UpdateChildContent "company|country", xCountry.text
xml.UpdateChildContent "company|governate", xGovernate.text
xml.UpdateChildContent "company|regionCity", xRegionCity.text
xml.UpdateChildContent "company|street", xStreet.text
xml.UpdateChildContent "company|buildingNumber", xBuildingNumber.text
xml.UpdateChildContent "company|companyTradeName", xCompanyTradeName.text
xml.UpdateChildContent "company|branchCode", xBranchCode.text
xml.UpdateChildContent "company|deviceSerialNumber", xDeviceSerialNumber.text
xml.UpdateChildContent "company|activityCode", xActivityCode.text
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
openCon con
cmdSave.Enabled = cBranch <> ""
myload
End Sub
Private Sub myload()
Dim xml As New ChilkatXml
Dim child As New ChilkatXml
Dim success As Integer

success = xml.LoadXmlFile(tempPath & "\receipt.xml")
If success = 0 Then
    MsgBox xml.LastErrorXml
End If

xRin.text = xml.GetChildContent("company|rin")
xCompanyTradeName.text = xml.GetChildContent("company|companyTradeName")
xBranchCode.text = xml.GetChildContent("company|branchCode")
xDeviceSerialNumber.text = xml.GetChildContent("company|deviceSerialNumber")
xActivityCode.text = xml.GetChildContent("company|activityCode")
xCountry.text = xml.GetChildContent("company|country")
xGovernate.text = xml.GetChildContent("company|governate")
xRegionCity.text = xml.GetChildContent("company|regionCity")
xStreet.text = xml.GetChildContent("company|street")
xBuildingNumber.text = xml.GetChildContent("company|buildingNumber")
    
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
Private Sub xBuildingNumber_GotFocus()
myGotFocus xBuildingNumber
End Sub
Private Sub xBuildingNumber_LostFocus()
myLostFocus xBuildingNumber
End Sub
Private Sub xStreet_GotFocus()
myGotFocus xStreet
End Sub
Private Sub xStreet_LostFocus()
myLostFocus xStreet
End Sub
Private Sub xRegionCity_GotFocus()
myGotFocus xRegionCity
End Sub
Private Sub xRegionCity_LostFocus()
myLostFocus xRegionCity
End Sub
Private Sub xGovernate_GotFocus()
myGotFocus xGovernate
End Sub
Private Sub xGovernate_LostFocus()
myLostFocus xGovernate
End Sub
Private Sub xCountry_GotFocus()
myGotFocus xCountry
End Sub
Private Sub xCountry_LostFocus()
myLostFocus xCountry
End Sub
Private Sub xActivityCode_GotFocus()
myGotFocus xActivityCode
End Sub
Private Sub xActivityCode_LostFocus()
myLostFocus xActivityCode
End Sub
Private Sub xDeviceSerialNumber_GotFocus()
myGotFocus xDeviceSerialNumber
End Sub
Private Sub xDeviceSerialNumber_LostFocus()
myLostFocus xDeviceSerialNumber
End Sub
Private Sub xBranchCode_GotFocus()
myGotFocus xBranchCode
End Sub
Private Sub xBranchCode_LostFocus()
myLostFocus xBranchCode
End Sub
Private Sub xCompanyTradeName_GotFocus()
myGotFocus xCompanyTradeName
End Sub
Private Sub xCompanyTradeName_LostFocus()
myLostFocus xCompanyTradeName
End Sub
Private Sub xRin_GotFocus()
myGotFocus xRin
End Sub
Private Sub xRin_LostFocus()
myLostFocus xRin
End Sub
