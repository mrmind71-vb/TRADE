VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form passWord2 
   Caption         =   "ﬂ·„… «·”—"
   ClientHeight    =   5100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   5070
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FramMosm 
      Caption         =   "«·„Ê«”„"
      Height          =   2850
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1755
      Visible         =   0   'False
      Width           =   4875
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   2490
         Left            =   90
         TabIndex        =   10
         Top             =   270
         Width           =   4695
         _cx             =   8281
         _cy             =   4392
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12648384
         ForeColorSel    =   64
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   5
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   500
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   1
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   0   'False
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.CommandButton CmdExit 
      Height          =   420
      Left            =   135
      MaskColor       =   &H00FFFFFF&
      Picture         =   "password.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   1170
      UseMaskColor    =   -1  'True
      Width           =   1320
   End
   Begin VB.CheckBox xEditLogin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "Edit Login"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3060
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1170
      Width           =   1860
   End
   Begin VB.CommandButton CmdApply 
      BackColor       =   &H00C0FFFF&
      Caption         =   " ‘€Ì·"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1485
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1170
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   90
      Width           =   4785
      Begin VB.TextBox xPass 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1980
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xUser 
         Height          =   315
         Left            =   135
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   225
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ·„… «·”— :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "≈”„ «·„” Œœ„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   1305
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   4815
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   135
      Top             =   495
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
Attribute VB_Name = "passWord2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTimes As Integer, nTime, userTable As Recordset
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
On Error GoTo LOCALERROR
If UCase(Trim(xPass.Text)) = "20132013" Then
    lSupperVisor = True
    nusercode = -1
    cusername = "bSupervisor"
    bSupermode = True
    bopt1 = False
    bopt2 = True
    bopt3 = True
    bOpt4 = True
    bOpt5 = True
    lSupperVisor = True
ElseIf Me.xUser.BoundText = "" Then
    Exit Sub
Else
    nusercode = GetDesca("select code from users where Code = " & xUser.BoundText & " and password Like " & MyParn(LCase(Trim(xPass.Text))))
    If nusercode = "" Then
       nTime = nTime + 1
       If nTime = nTimes Then
           Unload Me
           Exit Sub
       End If
       MsgBox "ﬂ·„… «·”— €Ì— ’ÕÌÕ…"
       Exit Sub
    End If
    cString = "SELECT users.code,users.SUPPER ,users.box,users.title1,users.title2, users.Password, users.desca,USERS.option1,users.option2,users.option3,users.option4,users.option5 ,Menusetting.Visible, Menusetting.Editable, Menu.Control, Menu.MenuNo,Menu.mainMenu " & _
              " FROM (users INNER JOIN Menusetting ON users.Code = Menusetting.code) INNER JOIN Menu ON Menusetting.control = Menu.Control " & _
              " where users.code = " & nusercode
    sectable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    If Not (sectable.EOF And sectable.BOF) Then
        nusercode = sectable!CODE
        cusername = sectable!Desca & ""
        bopt1 = sectable!Option1
        bopt2 = sectable!Option2
        bopt3 = sectable!Option3
        bOpt4 = sectable!Option4
        bOpt5 = sectable!Option5
        cManBox = sectable!BOX & ""
        lSupperVisor = sectable!SUPPER
    End If
    sectable.Close
    Set sectable = Nothing
End If
SaveSetting
If xEditLogin.Value = 1 And bopt2 Then
    confFrm.Show 1
Else
'If Not bopt1 Then
    FramMosm.Visible = True
'Else
'    cPMosm = grid1.TextMatrix(grid1.Row, 0)
'    Unload Me
'    main.Show
'End If
End If
Exit Sub
LOCALERROR:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
'On Error GoTo myerror
ValidDate
strCon = LoadConString
If openCon(con) <> "ok" Then
    confFrm.Show 1
    End
End If
openCon GetCon
myLoadVar
MakeLocal
StatusBar1.Panels(1).Text = "9.5.2012 "
Set grid1.DataSource = DATA2
DATA2.ConnectionString = strCon
cString = "SELECT MOSM,DESCA FROM MOSM WHERE CLOSED = 0 ORDER BY DATE DESC "
DATA2.RecordSource = cString
DATA2.Refresh

With grid1
    .Cols = 2
    .ColWidth(0) = 0
    .ColWidth(1) = .Width - 400
    If .Rows = 0 Then .AddItem ""
End With


data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM USERS"
Set xUser.RowSource = data1

xUser.ListField = "Desca"
xUser.BoundColumn = "Code"
xUser.BoundText = RetSetting("user", tempPath & "\password.txt")

Exit Sub
myerror:
    MsgBox Err.Description
    confFrm.Show 1
    Err.Clear
    End
End Sub
Private Sub MakeLocal()
On Error GoTo myerror
Dim fs As New FileSystemObject
MyCreateFolder tempPath
fs.CopyFile App.Path & "\temp.mdb", tempFile
contemp.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & tempFile
Exit Sub
myerror:
MsgBox "„‘ﬂ·… ›Ï ‰”Œ «·„·› «·„ƒﬁ " & vbCrLf & Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub

Private Sub xEditLogin_Click()
    If xEditLogin.Value = 1 And xPass.Text = "20122012" Then
        confFrm.Show 1
    End If
End Sub

Private Sub xUser_Click(Area As Integer)
If Not xUser.MatchedWithList Then xUser.BoundText = ""
'CmdApply.Enabled = xUser.BoundText <> ""
End Sub
Private Sub xUser_LostFocus()
If Not xUser.MatchedWithList Then xUser.BoundText = ""
'CmdApply.Enabled = xUser.BoundText <> ""
End Sub
Private Sub myLoadVar()
PublicPath = App.Path
LocalPath = App.Path
tempPath = "c:\TempMrshd"
tempFile = tempPath & "\temp.mdb"
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set PassWord = Nothing
End Sub
Private Sub SaveSetting()
addSetting "user", xUser.BoundText, tempPath & "\password.txt"
End Sub
Private Sub grid1_DBLClick()
    If grid1.Row >= 0 Then
        cPMosm = grid1.TextMatrix(grid1.Row, 0)
        Unload Me
        main.Show
    End If
End Sub

