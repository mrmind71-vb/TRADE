VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form SettingFrm 
   Caption         =   "»Ì«‰«  «·„Õ·"
   ClientHeight    =   6330
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   9675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   9675
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FRM_DISPLAY 
      Caption         =   "»Ì«‰«  ‘«‘…  DISPLAY"
      Height          =   780
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   5220
      Width           =   9285
      Begin VB.CommandButton Command1 
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
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Setting.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   960
      End
      Begin VB.TextBox XDISPLAY 
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
         Left            =   1125
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   270
         Width           =   5100
      End
      Begin VB.TextBox xPort 
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
         Left            =   7155
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   270
         Width           =   1050
      End
      Begin VB.Label Label13 
         Caption         =   "Display"
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
         Left            =   6300
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   360
         Width           =   870
      End
      Begin VB.Label Label5 
         Caption         =   "COMM"
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
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   360
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Œ’„ Œ«’ ··›—⁄"
      Height          =   825
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   4320
      Width           =   9375
      Begin VB.CommandButton CMD_SAVEDISC 
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
         Left            =   135
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Setting.frx":2363
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.TextBox XDISC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2295
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   360
         Width           =   825
      End
      Begin VB.TextBox XDATE2_DISC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4185
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox XDATE1_DISC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5850
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   1500
      End
      Begin VB.Label Label12 
         Caption         =   "«·Œ’„ : "
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   405
         Width           =   960
      End
      Begin VB.Label Label11 
         Caption         =   "«·› —… «·Œ«’… »«·Œ’„"
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
         Left            =   7470
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   405
         Width           =   1725
      End
   End
   Begin VB.Frame Framebox 
      Caption         =   "Ÿ»ÿ «·Œ“‰… Ê «·„Œ“‰ ··›—⁄ "
      Height          =   690
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3600
      Width           =   6540
      Begin VB.CommandButton cmd_okbox 
         Caption         =   "Õ›Ÿ"
         Height          =   465
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin MSDataListLib.DataCombo xbox_pc 
         Height          =   315
         Left            =   3870
         TabIndex        =   16
         Top             =   270
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xstore_pc 
         Height          =   315
         Left            =   990
         TabIndex        =   17
         Top             =   270
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label8 
         Caption         =   "„Œ“‰ :"
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
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Label7 
         Caption         =   "Œ“‰… :"
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
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   270
         Width           =   555
      End
   End
   Begin VB.CommandButton CmdExit 
      Height          =   465
      Left            =   135
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Setting.frx":46C6
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2925
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
      Left            =   1530
      MaskColor       =   &H00FFFFFF&
      Picture         =   "Setting.frx":6AE4
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Õ›Ÿ"
      Top             =   2925
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   9465
      Begin VB.TextBox XBACK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   2430
         Width           =   870
      End
      Begin VB.TextBox XID_NO2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4545
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   2385
         Width           =   3480
      End
      Begin VB.TextBox XID_NO1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   4545
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2025
         Width           =   3480
      End
      Begin VB.TextBox XCOUNTPRINT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2025
         Width           =   870
      End
      Begin VB.TextBox xhead2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3420
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2385
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.TextBox xhead1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   690
         Left            =   135
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1260
         Width           =   7890
      End
      Begin VB.TextBox xPhone 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   7890
      End
      Begin VB.TextBox xAddress 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   7890
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   7890
      End
      Begin VB.Label Label14 
         Caption         =   "BAKUP"
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
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   2475
         Width           =   1725
      End
      Begin VB.Label Label10 
         Caption         =   "”Ã·  Ã«—Ï "
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
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2430
         Width           =   1185
      End
      Begin VB.Label Label9 
         Caption         =   "—ﬁ„  ”ÃÌ· ÷"
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
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2070
         Width           =   1185
      End
      Begin VB.Label Label6 
         Caption         =   "⁄œœ „—«  ÿ»«⁄… «·»Ê‰"
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
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2115
         Width           =   1725
      End
      Begin VB.Label Label4 
         Caption         =   "„·«ÕŸ«  «·»Ê‰ "
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
         Left            =   8055
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1485
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "«· ·Ì›Ê‰ :"
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
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   945
         Width           =   1185
      End
      Begin VB.Label Label2 
         Caption         =   "«·⁄‰Ê«‰ :"
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
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "≈”„ «·„Õ· :"
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
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   1005
      End
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
   Begin MSAdodcLib.Adodc data2 
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
Attribute VB_Name = "SettingFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmd_okbox_Click()
    addSetting "BOX", xbox_pc.BoundText, "C:\TempMrshd\BRANCH_J.txt"
    addSetting "STORE", xstore_pc.BoundText, "C:\TempMrshd\BRANCH_J.txt"
    addSetting "DATETIME", Now(), "C:\TempMrshd\BRANCH_J.txt"
    addSetting "USER", cusername, "C:\TempMrshd\BRANCH_J.txt"
    
    Inform " „ «· ⁄œÌ·"
    End
End Sub

Private Sub CMD_SAVEDISC_Click()
    con.Execute " UPDATE ADDRESS SET DATE1_DISC = " & addDate(XDATE1_DISC.text) & " , DATE2_DISC = " & addDate(XDATE2_DISC.text) & " , DISC = " & Val(XDISC.text) & " WHERE BRANCH = " & MyParn(cBranch), nRec
    If nRec > 0 Then MsgBox "  „ «· ⁄œÌ· "
End Sub

Private Sub cmdSave_Click()
Dim nId As Double
Dim aInsert(10, 1)
If lIsBranchStore Then
    cId = GetDesca("SELECT MAX(ID)  FROM ADDRESS WHERE BRANCH = " & MyParn(cBranch), con)
Else
    cId = GetDesca("SELECT MAX(ID)  FROM ADDRESS ", con)
End If
If cId <> "" Then nId = Val(cId)
aInsert(0, 0) = "ID"
aInsert(0, 1) = addvalue(nId)

aInsert(1, 0) = "Desca"
aInsert(1, 1) = addstring(xDesca.text)

aInsert(2, 0) = "Address"
aInsert(2, 1) = addstring(xAddress.text)

aInsert(3, 0) = "Phone"
aInsert(3, 1) = addstring(xPhone.text)

aInsert(4, 0) = "head1"
aInsert(4, 1) = addstring(xhead1.text)

aInsert(5, 0) = "head2"
aInsert(5, 1) = addstring(xhead2.text)

aInsert(6, 0) = "COUNTPRINT"
aInsert(6, 1) = Val(XCOUNTPRINT.text)

aInsert(7, 0) = "branch"
aInsert(7, 1) = addstring(cBranch)

aInsert(8, 0) = "ID_NO1"
aInsert(8, 1) = addstring(XID_NO1.text)

aInsert(9, 0) = "ID_NO2"
aInsert(9, 1) = addstring(XID_NO2.text)

aInsert(10, 0) = "BACK"
aInsert(10, 1) = addstring(XBACK.text)

On Error GoTo myerror
''con.BeginTrans
If cId = "" Then
    nId = Val(GetDesca("SELECT MAX(ID) FROM Address", con) & "") + 1
    aInsert(0, 1) = addvalue(nId)
    con.Execute CreateInsert(aInsert, "Address")
Else
   con.Execute CreateUpdate(aInsert, "Address", " WHERE ID = " & nId)
End If

'con.CommitTrans
con.Execute " UPDATE Address SET HEAD1 = " & addstring(xhead1.text) ' & " , ID_NO1 = " & addstring(XID_NO1.text) & " , ID_NO2 = " & addstring(XID_NO2.text)
Inform " „ «· ⁄œÌ· »‰Ã«Õ"

Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
'con.RollbackTrans
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub Command1_Click()
    If lIsBranchStore Then
        If GetDesca("SELECT ID FROM COMP WHERE BRANCH = " & MyParn(cBranch), con) = "" Then
            con.Execute "  INSERT INTO COMP (BRANCH ,  PORT , DISPLAY , ADDRESS ) VALUES  (" & addstring(cBranch) & " , " & addstring(xPort.text) & " , " & addstring(XDISPLAY.text) & " , " & addstring(xDesca.text) & " )"
            MsgBox " «„ «÷«›… «·»Ê—   "
        Else
            con.Execute " UPDATE COMP SET PORT = " & addstring(xPort.text) & " , DISPLAY = " & addstring(XDISPLAY.text) & " WHERE BRANCH = " & MyParn(cBranch)
        End If
    End If
End Sub

Private Sub Form_Load()
    openCon con
    myload
    
    FRM_DISPLAY.Visible = (lIsBranchStore And bopt2) Or (cusername = "Aymen Junior")
    
    Framebox.Visible = bopt2 Or (cusername = "Aymen Junior")
    Frame2.Visible = bSupermode Or (cusername = "Aymen Junior")
    
    data1.ConnectionString = strCon
    If lIsBranchStore Then
        data1.RecordSource = "SELECT * FROM FILE0_50 where branch = " & MyParn(cBranch)
    Else
        data1.RecordSource = "SELECT * FROM FILE0_50"
    End If
    
    Set xbox_pc.RowSource = data1
    xbox_pc.ListField = "Desca"
    xbox_pc.BoundColumn = "Code"
    xbox_pc.BoundText = RetSetting("BOX", "C:\TempMrshd\BRANCH_J.txt")
    
    data2.ConnectionString = strCon
    If lIsBranchStore Then
        data2.RecordSource = "SELECT * FROM FILE0_40 where branch = " & MyParn(cBranch)
    Else
        data2.RecordSource = "SELECT * FROM FILE0_40  where branch = " & MyParn(cBranch)
    End If
    Set xstore_pc.RowSource = data2
    xstore_pc.ListField = "Desca"
    xstore_pc.BoundColumn = "Code"
    xstore_pc.BoundText = RetSetting("store", "C:\TempMrshd\BRANCH_J.txt")

    aRet = aGetDesca("SELECT DATE1_DISC , DATE2_DISC , DISC FROM ADDRESS  WHERE BRANCH = " & MyParn(cBranch), con)
    If UBound(aRet) > 0 Then
        XDATE1_DISC.text = Format(aRet(1), "DD-MM-YYYY") & ""
        XDATE2_DISC.text = Format(aRet(2), "DD-MM-YYYY") & ""
        XDISC.text = aRet(3) & ""
    End If
      
    aRet = aGetDesca("SELECT ID_NO1 , ID_NO2 , head1 FROM ADDRESS WHERE BRANCH = '00'", con)
    If UBound(aRet) > 0 Then
'        XID_NO1.text = aRet(1) & ""
'        XID_NO2.text = aRet(2) & ""
        xhead1.text = aRet(3) & ""
    End If
    
'    XID_NO1.Locked = (cBranch <> "00")
'    XID_NO2.Locked = (cBranch <> "00")
    xhead1.Locked = (cBranch <> "00")

    If lIsBranchStore Then
        xPort.text = GetDesca("SELECT PORT FROM COMP WHERE BRANCH = " & MyParn(cBranch), con)
        XDISPLAY.text = GetDesca("SELECT DISPLAY FROM COMP WHERE BRANCH = " & MyParn(cBranch), con)
        If xPort.text = "" Then
            xPort.text = "COM7"
            FRM_DISPLAY.Caption = FRM_DISPLAY.Caption & "  " & " «·÷⁄ÿ ⁄·Ï Õ›Ÿ  · ”ÃÌ· «·»Ê—  «·Œ«’ »«·ÃÂ«“ "
        End If
        If XDISPLAY.text = "" Then XDISPLAY.text = "WELCOM TO JUNIOR"
        XBACK.Visible = False
        Label14.Visible = False
    End If


End Sub
Private Sub myload()
Dim aRet As Variant
If lIsBranchStore Then
    aRet = aGetDesca("Select desca,address,Phone,head1,head2 , COUNTPRINT   , LANG , ID_NO1 , ID_NO2 , '' from Address where branch = " & MyParn(cBranch), con)
Else
    aRet = aGetDesca("Select desca,address,Phone,head1,head2 , COUNTPRINT   , LANG , ID_NO1 , ID_NO2 , back from Address", con)
End If
If UBound(aRet) > 0 Then
    xDesca.text = aRet(1) & ""
    xAddress.text = aRet(2) & ""
    xPhone.text = aRet(3) & ""
    xhead1.text = aRet(4) & ""
    xhead2.text = aRet(5) & ""
    XCOUNTPRINT.text = Val(aRet(6) & "")
    XID_NO1.text = aRet(8) & ""
    XID_NO2.text = aRet(9) & ""
    XBACK.text = aRet(10) & ""
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
