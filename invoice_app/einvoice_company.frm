VERSION 5.00
Begin VB.Form einvoice_companyfrm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "÷»ÿ «⁄œ«œ   ÿ»Ìﬁ «·„Õ„Ê·"
   ClientHeight    =   5895
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
   ScaleHeight     =   5895
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5280
      Left            =   90
      TabIndex        =   16
      Top             =   0
      Width           =   11445
      Begin VB.TextBox xName 
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
         TabIndex        =   13
         Top             =   4860
         Width           =   8430
      End
      Begin VB.TextBox xUser_id 
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
         TabIndex        =   12
         Top             =   4500
         Width           =   8430
      End
      Begin VB.TextBox xType 
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
         TabIndex        =   11
         Top             =   4140
         Width           =   8430
      End
      Begin VB.TextBox xAdditional 
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
         Top             =   3780
         Width           =   8430
      End
      Begin VB.TextBox xLandMark 
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
         Top             =   3420
         Width           =   8430
      End
      Begin VB.TextBox xRoom 
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
         Top             =   3060
         Width           =   8430
      End
      Begin VB.TextBox xFloor 
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
         Top             =   2700
         Width           =   8430
      End
      Begin VB.TextBox xPostal 
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
         Top             =   2340
         Width           =   8430
      End
      Begin VB.TextBox xBuilding 
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
      Begin VB.TextBox xStreet 
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
      Begin VB.TextBox xRegionCity 
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
      Begin VB.TextBox xGovernate 
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
         Locked          =   -1  'True
         MaxLength       =   200
         TabIndex        =   1
         Text            =   "EG"
         Top             =   540
         Width           =   8430
      End
      Begin VB.TextBox xBranchId 
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
      Begin VB.Label Label6 
         Caption         =   "Name"
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
         Index           =   7
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   4905
         Width           =   960
      End
      Begin VB.Label Label6 
         Caption         =   "ID"
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
         Index           =   6
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   4545
         Width           =   960
      End
      Begin VB.Label Label6 
         Caption         =   "Type"
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
         Index           =   5
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   4185
         Width           =   960
      End
      Begin VB.Label Label6 
         Caption         =   "Additional Information"
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
         TabIndex        =   27
         Top             =   3825
         Width           =   2310
      End
      Begin VB.Label Label6 
         Caption         =   "Land Mark"
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
         TabIndex        =   26
         Top             =   3465
         Width           =   1500
      End
      Begin VB.Label Label6 
         Caption         =   "Room"
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
         TabIndex        =   25
         Top             =   3105
         Width           =   1050
      End
      Begin VB.Label Label6 
         Caption         =   "Floor"
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
         TabIndex        =   24
         Top             =   2745
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Postal Code"
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
         TabIndex        =   23
         Top             =   2385
         Width           =   1860
      End
      Begin VB.Label Label4 
         Caption         =   "Building Number"
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
         TabIndex        =   22
         Top             =   2025
         Width           =   1860
      End
      Begin VB.Label Label3 
         Caption         =   "Street"
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
         TabIndex        =   21
         Top             =   1665
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Region City"
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
         Top             =   1305
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Governrate"
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
         TabIndex        =   19
         Top             =   945
         Width           =   1455
      End
      Begin VB.Label Label10 
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
         Left            =   225
         TabIndex        =   18
         Top             =   585
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Branch ID"
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
         Top             =   225
         Width           =   1275
      End
   End
   Begin VB.CommandButton CmdExit 
      Height          =   465
      Left            =   10170
      MaskColor       =   &H00FFFFFF&
      Picture         =   "einvoice_company.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   5355
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
      Picture         =   "einvoice_company.frx":4AB6
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Õ›Ÿ"
      Top             =   5355
      UseMaskColor    =   -1  'True
      Width           =   1455
   End
End
Attribute VB_Name = "einvoice_companyfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
myreplace
End Sub
Private Function MYVALID() As Boolean
If Trim(xBranchId.text) = "" Then
    MsgBox "Branch required"
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

If Trim(xStreet.text) = "" Then
    MsgBox "Building Number required"
    Exit Function
End If


If Trim(xBuilding.text) = "" Then
    MsgBox "Building Number required"
    Exit Function
End If
MYVALID = True
End Function
Private Function myreplace() As Boolean
Dim aInsert As Variant
aInsert = AddFlag(Empty, "BranchId", addstring(xBranchId.text))
aInsert = AddFlag(aInsert, "Country", addstring(xCountry.text))
aInsert = AddFlag(aInsert, "Governate", addstring(xGovernate.text))
aInsert = AddFlag(aInsert, "RegionCity", addstring(xRegionCity.text))
aInsert = AddFlag(aInsert, "Street", addstring(xStreet.text))
aInsert = AddFlag(aInsert, "Building", addstring(xBuilding.text))
aInsert = AddFlag(aInsert, "Postal", addstring(xPostal.text))
aInsert = AddFlag(aInsert, "Floor", addstring(xFloor.text))
aInsert = AddFlag(aInsert, "Room", addstring(xRoom.text))
aInsert = AddFlag(aInsert, "LandMark", addstring(xLandMark.text))
aInsert = AddFlag(aInsert, "Additional", addstring(xAdditional.text))
aInsert = AddFlag(aInsert, "Type", addstring(xType.text))
aInsert = AddFlag(aInsert, "user_id", addstring(xUser_id.text))
aInsert = AddFlag(aInsert, "name", addstring(xname.text))

con.BeginTrans
On Error GoTo myerror
If IsEmpty(myField("Select id from SettingCompany", con)) Then
   con.Execute addInsert(aInsert, "SettingCompany")
Else
   con.Execute addUpdate(aInsert, "SettingCompany", "id = 1")
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
Dim loctable As ADODB.Recordset
Set loctable = myCmd("Select * from SettingCompany", con)
If Not loctable.EOF Then
    xBranchId.text = loctable!branchId & ""
    xCountry.text = loctable!Country & ""
    xGovernate.text = loctable!GOVERNATE & ""
    xRegionCity.text = loctable!RegionCity & ""
    xStreet.text = loctable!Street & ""
    xBuilding.text = loctable!building & ""
    xPostal.text = loctable!Postal & ""
    xFloor.text = loctable!Floor & ""
    xRoom.text = loctable!Room & ""
    xLandMark.text = loctable!LandMark & ""
    xAdditional.text = loctable!Additional & ""
    xType.text = loctable!Type & ""
    xname.text = loctable!Name & ""
    xUser_id.text = loctable!user_id & ""
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
Private Sub xName_GotFocus()
myGotFocus xname
End Sub
Private Sub xname_LostFocus()
myLostFocus xname
End Sub
Private Sub xID_GotFocus()
myGotFocus xID
End Sub
Private Sub xId_LostFocus()
myLostFocus xID
End Sub
Private Sub xType_GotFocus()
myGotFocus xType
End Sub
Private Sub xType_LostFocus()
myLostFocus xType
End Sub
Private Sub xAdditional_GotFocus()
myGotFocus xAdditional
End Sub
Private Sub xAdditional_LostFocus()
myLostFocus xAdditional
End Sub
Private Sub xLandMark_GotFocus()
myGotFocus xLandMark
End Sub
Private Sub xLandMark_LostFocus()
myLostFocus xLandMark
End Sub
Private Sub xRoom_GotFocus()
myGotFocus xRoom
End Sub
Private Sub xRoom_LostFocus()
myLostFocus xRoom
End Sub
Private Sub xFloor_GotFocus()
myGotFocus xFloor
End Sub
Private Sub xFloor_LostFocus()
myLostFocus xFloor
End Sub
Private Sub xPostal_GotFocus()
myGotFocus xPostal
End Sub
Private Sub xPostal_LostFocus()
myLostFocus xPostal
End Sub
Private Sub xBuilding_GotFocus()
myGotFocus xBuilding
End Sub
Private Sub xBuilding_LostFocus()
myLostFocus xBuilding
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
Private Sub xBranchId_GotFocus()
myGotFocus xBranchId
End Sub
Private Sub xBranchId_LostFocus()
myLostFocus xBranchId
End Sub
