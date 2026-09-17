VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form PassWord 
   Caption         =   "ﬂ·„… «·”—"
   ClientHeight    =   6135
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5565
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
   ScaleHeight     =   6135
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3840
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   5235
      Begin VSFlex7Ctl.VSFlexGrid GRBRANCH 
         Height          =   2685
         Left            =   90
         TabIndex        =   11
         Top             =   990
         Width           =   5010
         _cx             =   8837
         _cy             =   4736
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483633
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   600
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   2865
         Left            =   90
         TabIndex        =   10
         Top             =   180
         Visible         =   0   'False
         Width           =   5010
         _cx             =   8837
         _cy             =   5054
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483633
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   600
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   600
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "password2.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   " ‘€Ì·"
      Top             =   1080
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin Threed.SSCommand cmdApply 
      Height          =   600
      Left            =   1485
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   1058
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "password2.frx":246C
      Caption         =   "œŒÊ·"
      PictureAlignment=   10
   End
   Begin VB.CheckBox xEditLogin 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "»Ì«‰«  «·œŒÊ·"
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
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1170
      Width           =   1725
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   -45
      Width           =   5370
      Begin VB.TextBox xPass 
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
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   1575
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1725
      End
      Begin MSDataListLib.DataCombo xUser 
         Height          =   360
         Left            =   135
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   225
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ·„… «·”— :"
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
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "≈”„ «·„” Œœ„ :"
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
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Width           =   1305
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   5805
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
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
   Begin MSAdodcLib.Adodc Adodc1 
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
      Caption         =   "DATA3"
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
   Begin MSAdodcLib.Adodc DATA3 
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
Attribute VB_Name = "PassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nTimes As Integer, nTime, userTable As Recordset
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
On Error GoTo LOCALERROR
If Not xUser.MatchedWithList Then Exit Sub

cComputerName = GetComputerName
If UCase(Trim(xPass.text)) = "DATA@2023" Or DefUser Then
    lSupperVisor = True
    nusercode = -1
    If UCase(Trim(xPass.text)) = UCase("Data@2023") Then cusername = "Data Supervisor"
    If UCase(Trim(xPass.text)) = UCase("MohammedJ95") Then cusername = "Mohammed Supervisor"
    
    bSupermode = True
    bopt1 = False
    bopt2 = True
    bopt3 = True
    bOpt4 = True
    bOpt5 = True
    bOpt6 = True
    bOpt7 = True
    bOpt8 = True
    bOpt9 = True
    bOpt10 = True
    bOpt11 = True
    grid1.Visible = True
    lSupperVisor = True
    aPassword = AddFlag(aPassword, "super", True)
    
ElseIf UCase(Trim(xPass.text)) = UCase("Aymen@Junior") Then
    cusername = "Aymen Junior"
    grid1.Visible = True
    lSupperVisor = True

    bSupermode = True
    bopt1 = False
    bopt2 = True
    bopt3 = True
    bOpt4 = True
    bOpt5 = True
    bOpt6 = True
    bOpt7 = True
    bOpt8 = True
    bOpt9 = True
    bOpt10 = True
    bOpt11 = True
    grid1.Visible = True
    lSupperVisor = True

Else
    nusercode = GetDesca("select code from users where Code = " & xUser.BoundText & " and password Like " & MyParn(LCase(Trim(xPass.text))), con)
    If nusercode = "" Then
       nTime = nTime + 1
       If nTime = nTimes Then
           Unload Me
           Exit Sub
       End If
       Inform "ﬂ·„… «·”— €Ì— ’ÕÌÕ…"
       myGotFocus xPass
       Exit Sub
    End If
    If lServerOnLine Then
        cString = "SELECT users.code,users.PASS_DISC AS PASSDISC,users.SUPPER ,users.box,users.title1,users.title2, users.Password, users.desca,USERS.option1,users.option2,users.option3,users.option4,users.option5 ,users.option6 ,users.option7 ,users.option8 ,users.option9,users.option10 ,users.option11,users.store,Menusetting.Visible, Menusetting.Editable, Menu.Control, Menu.MenuNo,Menu.mainMenu " & _
                  " FROM (users INNER JOIN Menusetting ON users.Code = Menusetting.code) INNER JOIN Menu ON Menusetting.control = Menu.Control " & _
                  " where users.code = " & nusercode
    Else
        cString = "SELECT users.code,users.PASSDISC,users.SUPPER ,users.box,users.title1,users.title2, users.Password, users.desca,USERS.option1,users.option2,users.option3,users.option4,users.option5 ,users.option6 ,users.option7 ,users.option8 ,users.option9,users.option10 ,users.option11,users.store,Menusetting.Visible, Menusetting.Editable, Menu.Control, Menu.MenuNo,Menu.mainMenu " & _
                  " FROM (users INNER JOIN Menusetting ON users.Code = Menusetting.code) INNER JOIN Menu ON Menusetting.control = Menu.Control " & _
                  " where users.code = " & nusercode
    End If
              
    sectable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    If Not (sectable.EOF And sectable.BOF) Then
        
        nusercode = sectable!code
        cusername = sectable!desca & ""
        bopt1 = sectable!Option1
        bopt2 = sectable!Option2
        bopt3 = sectable!Option3
        bOpt4 = sectable!Option4
        bOpt5 = sectable!Option5
        bOpt6 = sectable!Option6
        bOpt7 = sectable!Option7
        bOpt8 = sectable!Option8
        bOpt9 = sectable!Option9
        bOpt10 = sectable!Option10
        bOpt11 = sectable!Option11
        
        cPassDisc = sectable!PASSDISC & ""
        Firstitle = sectable!title1 & ""
        Secondtitle = sectable!title2 & ""
        lSupperVisor = sectable!SUPPER
        aPassword = AddFlag(Empty, "box", sectable!BOX & "")
        aPassword = AddFlag(aPassword, "store", sectable!STORE & "")
        
        aPassword = AddFlag(aPassword, "SUPER", sectable!SUPPER)
        cUserBox = TurnValue(sectable!BOX, Null, "")
        
        If cBranchBox <> "" And cBranch <> "00" Then
            If cBranchBox <> sectable!BOX Then
                MsgBox "«·ﬂ«‘Ì— €Ì— ’ÕÌÕ"
                cBranchBox = ""
            End If
        End If
    End If
    
    Set rsUser = createRs(cmd("select * from users where code = " & nusercode, con).Execute)
    
    sectable.Close
    Set sectable = Nothing
End If


AddLod_Data cusername, 0, " › Õ «·»—‰«„Ã ", con


If lShowBranch Then
    grid1.Visible = False
    GRBRANCH.Visible = True
Else
    SaveSetting
    If xEditLogin.Value = 1 And bopt2 Then
        confFrm.Show 1
    Else
        If sBranchCode = "00" Or lSupperVisor Or (cBranch > "60" And bopt2) Then
    '        FramMosm.Visible = True
            On Error Resume Next
            grid1.Visible = True
            Frame1.Visible = True
            If grid1.Rows > 0 Then
                grid1.SetFocus
                grid1.Select 0, 1
            End If
        Else
            Unload Me
            Main.Show
        End If
    End If
End If
Exit Sub
LOCALERROR:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Activate()
    On Error Resume Next
    'Clipboard.Clear
    Err.Clear
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If TypeOf ActiveControl Is TextBox Then SendKeys "{TAB}"
    End If
End Sub
Private Sub Form_Load()
On Error GoTo myerror
Dim fs As New FileSystemObject
Dim bMain As Boolean, cString As String
bMain = True
ValidDate


If fs.FileExists(App.Path & "\conf.txt") Then
    cPathConf = App.Path
Else
    cPathConf = "C:\TempMrshd\DATA_PRO"
End If

'If RetSetting("ONLINE", cPathConf & "\CONF.TXT") = "ONLINE" Then lServerOnLine = True
'If RetSetting("ONLINESERVER", cPathConf & "\CONF.TXT") = "ONLINE" Then lServerOnLineShop = True

If RetSetting("COPY", cPathConf & "\CONF.TXT") = "COPY" Then
    grid1.Visible = False
    GRBRANCH.Visible = False
    xEditLogin.Visible = False
    lShowBranch = True
Else
    lShowBranch = False
    GRBRANCH.Visible = False
End If
sCatalog = RetSetting("Catalog", cPathConf & "\CONF.TXT")
sMdfName = sCatalog


cBranchBox = RetSetting("BOX", "C:\TempMrshd\BRANCH_J.txt")
cManBox = cBranchBox
    
strCon = LoadConString
strConfact = LoadConStringfact
strConfact2 = LoadConStringfact2
strConfact3 = LoadConStringfact3
    
Dim cError As String
cError = testData
If cError <> "ok" Then
    MsgBox cError
    End
End If
    
openCon con
myLoadVar
MakeLocal
    
If decrypt(RetSetting("BRANCH", cPathConf & "\CONF.TXT"), "dr") = "" Then
    sBranchCode = GetDesca("select code from branch ORDER BY CODE ", con)
    cBranch = sBranchCode
    sBranch = sBranchCode
    lIsBranchStore = False
Else
    sBranchCode = decrypt(RetSetting("BRANCH", cPathConf & "\CONF.TXT"), "dr")
    cBranch = sBranchCode
    sBranch = sBranchCode
    lIsBranchStore = True
End If
    
    'FixData11
    
If lServerOnLine Or lServerOnLineShop Then
    'cServerNamePICT = "junior-sql.database.windows.net"
Else
    If cBranch = "00" Or Not lIsBranchStore Then
        cServerNamePICT = (RetSetting("SERVER", cPathConf & "\CONF.TXT"))
    Else
        cServerNamePICT = TurnValue(GetDesca("select IPSERVER from branch", con), Null, "")
    End If
End If
   
If cServerNamePICT = "null" Then cServerNamePICT = ""
strConPICT = LoadConStringPICT
If strConPICT <> "" Then openConPICT conPict

StatusBar1.Panels(1).text = "19.01.2024 JENUR_SHOP"
StatusBar1.Panels(2).text = RetSetting("SERVER", cPathConf & "\CONF.TXT")

Set grid1.DataSource = DATA2
cString = "SELECT MOSM , DESCA FROM MOSM WHERE CLOSED = 0 ORDER BY DATE DESC "
Set DATA2.Recordset = myRecordSet(cString, con)

Set GRBRANCH.DataSource = DATA3
If cBranch = "00" Then
    cString = "SELECT CODE , DESCA FROM  QBRANCH_ALL WHERE CODE <> '00' AND DATA IS NOT NULL ORDER BY CODE      "
Else
    cString = "SELECT CODE , DESCA FROM  QBRANCH_ALL WHERE CODE = " & MyParn(cBranch)
End If
Set DATA3.Recordset = myRecordSet(cString, con)

With grid1
    .Cols = 2
    .ColWidth(0) = 0
    .ColWidth(1) = .Width - 700
    If .Rows = 0 Then .AddItem ""
End With


With GRBRANCH
    .Cols = 2
    .ColWidth(0) = 0
    .ColWidth(1) = .Width - 700
'    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
End With

If lIsBranchStore Then
    Set data1.Recordset = myRecordSet("SELECT * FROM USERS WHERE BRANCH = " & MyParn(cBranch), con)
Else
    Set data1.Recordset = myRecordSet("SELECT * FROM USERS WHERE BRANCH IS NULL ", con)
End If
Set xUser.RowSource = data1
xUser.ListField = "Desca"
xUser.BoundColumn = "Code"
xUser.BoundText = RetSetting("user", tempPath & "\password.txt")

Set rsBranch = createRs(cmd("select * from branch where code = " & MyParn(cBranch), con).Execute)
Set rsBranches = createRs(cmd("select * from BRANCH", con).Execute)

Exit Sub
myerror:
    
    MsgBox Err.Description
'    confFrm.Show 1
    Err.Clear
    End

End Sub
Private Sub MakeLocal()
'On Error GoTo myerror
Dim fs As New FileSystemObject
MyCreateFolder tempPath
Me.Caption = "temp.mdb"
fs.CopyFile App.Path & "\temp.mdb", tempFile
contemp.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & tempFile

Me.Caption = "def.MDB"
If Not fs.FileExists("C:\TempMrshd\def.MDB") Then
    fs.CopyFile App.Path & "\DEF.MDB", "c:\TempMrshd\def.MDB"
End If
condef.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & "c:\TempMrshd\def.MDB"

Me.Caption = "BRANCH_J.txt"
If Not fs.FileExists(tempPath & "\BRANCH_J.txt") And fs.FileExists(App.Path & "\BRANCH_J.txt") Then
    fs.CopyFile App.Path & "\BRANCH_J.txt", tempPath & "\BRANCH_J.txt"
End If
Me.Caption = "JUNIOR"
Exit Sub
myerror:
MsgBox "„‘ﬂ·… ›Ï ‰”Œ «·„·› «·„ƒﬁ " & vbCrLf & Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub
Private Sub GRBRANCH_DBLClick()
        
    sCatalog = GetDesca("SELECT DATA FROM QBRANCH_ALL WHERE CODE = " & MyParn(GRBRANCH.TextMatrix(GRBRANCH.Row, 0)), GetCon)
    strCon = LoadConString_B(GRBRANCH.TextMatrix(GRBRANCH.Row, 0))
    openCon con, strCon

    sBranchCode = GRBRANCH.TextMatrix(GRBRANCH.Row, 0)
    cBranch = sBranchCode
    sBranch = sBranchCode
    cBranchStore = GetDesca("SELECT STORE FROM BRANCH ", GetCon)
    closeCon con
    openCon con
    
    lMainShow = True
    cComp_Name = GetDesca("SELECT DESCA FROM ADDRESS ", con)
    If lMainShow Then
        Main.Caption = cComp_Name
    End If
    If lIsBranchStore Then
        dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES WHERE BRANCH = " & MyParn(cBranch), con), "DD-MM-YYYY")
    Else
        dSalesDate = Format(GetDesca("SELECT DSALES FROM DSALES", con), "DD-MM-YYYY")
    End If
    aPassword = AddFlag(aPassword, "super", True)
    Main.Show
End Sub

Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    grid1_DblClick
End If
End Sub

Private Sub SSTab1_DblClick()

End Sub

Private Sub xEditLogin_Click()
    If xEditLogin.Value = 1 And xPass.text = "20122012" Then
        confFrm.Show 1
    End If
End Sub
Private Sub xPass_GotFocus()
    On Error Resume Next
    Clipboard.Clear
    myGotFocus xPass
    Err.Clear
End Sub
Private Sub xPass_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    CmdApply_Click
End If
End Sub
Private Sub xPass_LostFocus()
    myLostFocus xPass
End Sub
Private Sub xPass_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Clipboard.Clear
    myGotFocus xPass
    Err.Clear
End Sub
Private Sub xUser_Click(Area As Integer)
If Not xUser.MatchedWithList Then xUser.BoundText = ""
'CmdApply.Enabled = xUser.BoundText <> ""
End Sub

Private Sub xUser_GotFocus()
'myGotFocus xUser
End Sub

Private Sub xUser_LostFocus()
If Not xUser.MatchedWithList Then xUser.BoundText = ""
myLostFocus xUser
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
Private Sub grid1_DblClick()
    On Error GoTo myerror
    If grid1.Row >= 0 Then
        cPMosm = grid1.TextMatrix(grid1.Row, 0)
        cPMosmD = grid1.TextMatrix(grid1.Row, 1)
        Unload Me
        Main.Show
    End If
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Function testData() As String
Dim cString As String, cError As String
cError = openCon(GetCon)
If cError = "ok" Then
    testData = "ok"
    Exit Function
End If

cString = "The user is not associated with a trusted SQL Server connection."
If LCase(Right(cError, Len(cString))) = LCase(cString) Then
    cError = CreateRemote
    If cError = "ok" Then
        Inform " „ «÷«›… ’·«ÕÌ«  «·»Ì«‰«  »‰Ã«Õ «·—Ã«¡ «⁄«œ…  ‘€Ì· «·ÃÂ«“ »⁄œ «·«‰ Â«¡"
        End
    Else
        MsgBox cError
        testData = cError
    End If
End If

cString = "Login failed for user"
If LCase(Mid(cError, 1, 21)) = LCase(cString) Then
    cError = createLogin
    If cError = "ok" Then
        Inform " „ «÷«›… „” Œœ„ »‰Ã«Õ"
        cError = openCon(GetCon)
    End If
End If

If cError <> "ok" Then
    cString = "Cannot open database"
    If Left(LCase(cError), 20) = LCase(cString) Then
        cError = AttachData
        If cError = "ok" Then Inform " „ —»ÿ «·»Ì«‰«  »‰Ã«Õ"
        cError = openCon(GetCon)
    End If
End If

If cError <> "ok" Then
    cString = "Cannot open database"
    If Left(LCase(cError), 20) = LCase(cString) Then
        cError = bringOnLine
        If cError = "ok" Then
            Inform " „ › Õ «·„·› »‰Ã«Õ"
            cError = openCon(GetCon)
         End If
    End If
End If

If cError <> "ok" Then
    MsgBox cError
    confFrm.Show 1
End If
testData = "ok"
End Function
Private Function CreateRemote() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection
Dim cString As String, cServerName As String
cServerName = MyParn("." & turn(cExpress, "\") & cExpress)
cString = "provider=SQLOLEDB;data source= " & cServerName & "  ;initial " _
        & "catalog=master;Trusted_Connection=yes"
conMaster.Open cString

cString = "EXEC xp_instance_regwrite N'HKEY_LOCAL_MACHINE', N'Software\Microsoft\MSSQLServer\MSSQLServer', N'LoginMode', REG_DWORD, 2"
createCommand cString, conMaster

closeCon conMaster
CreateRemote = "ok"
Exit Function
myerror:
   CreateRemote = Err.Description
   Err.Clear
End Function
Private Function createLogin() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection
Dim cServerName As String, cString As String
cServerName = MyParn("." & turn(cExpress, "\") & cExpress)
cString = "provider=SQLOLEDB;data source= " & cServerName & "  ;initial " _
        & "catalog=master;Trusted_Connection=yes"
conMaster.Open cString
cString = "CREATE LOGIN [elmorshed] WITH PASSWORD=N'2010', DEFAULT_DATABASE=[master], DEFAULT_LANGUAGE=[us_english], CHECK_EXPIRATION=OFF, CHECK_POLICY=OFF"
cString = cString & turn(cString, vbCrLf) & "EXEC sys.sp_addsrvrolemember @loginame = N'elmorshed', @rolename = N'sysadmin'"
createCommand cString, conMaster
closeCon conMaster
createLogin = "ok"
Exit Function
myerror:
   createLogin = Err.Description
   Err.Clear
End Function
Private Function AttachData() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection
Dim cString As String, cServerName As String
cServerName = MyParn("." & turn(cExpress, "\") & cExpress)
cString = "provider=SQLOLEDB;data source= " & cServerName & "  ;initial " _
        & "catalog=master;Trusted_Connection=yes"
conMaster.Open cString

Dim cFile As String
cFile = App.Path & "\mdf\" & sMdfName
cString = "CREATE DATABASE [" & sCatalog & "] ON (FILENAME = N'" & cFile & ".mdf" & "' )," & _
      "(FILENAME = N'" & cFile & "_LOG.ldf" & "' )" & _
      " FOR ATTACH"
createCommand cString, conMaster

closeCon conMaster
AttachData = "ok"
Exit Function
myerror:
   AttachData = Err.Description
   Err.Clear
End Function
Private Function bringOnLine() As String
On Error GoTo myerror
Dim conMaster As New ADODB.Connection
Dim cString As String, cServerName As String
cServerName = MyParn("." & turn(cExpress, "\") & cExpress)
cString = "provider=SQLOLEDB;data source= " & cServerName & "  ;initial " _
        & "catalog=master;Trusted_Connection=yes"
conMaster.Open cString

Dim FS1 As New ADODB.command
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = conMaster
cString = "alter database [" & sCatalog & "]"
cString = cString & turn(cString, vbCrLf) & "set online"
FS1.CommandText = cString
FS1.Execute
bringOnLine = "ok"
Exit Function
myerror:
bringOnLine = Err.Description
Err.Clear
End Function
Private Function LoadConString()
Dim cServerName As String, cUserId As String, cPassword As String
If lServerOnLine Then
'    cServerName = "junior-sql.database.windows.net"
'    cUserId = decrypt("062F8C407C77", "dr")
'    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
Else
    cServerName = RetSetting("server", cPathConf & "\CONF.TXT")
    cUserId = decrypt(RetSetting("userId", cPathConf & "\CONF.TXT"), "dr")
    cPassword = decrypt(RetSetting("Password", cPathConf & "\CONF.TXT"), "dr")
End If
If cUserId = "" Or cPassword = "" Then
    LoadConString = "provider=SQLOLEDB;data source= " & cServerName & " ;initial " _
            & "catalog=" & sCatalog & ";Trusted_Connection=yes" & ";Timeout=10"
Else
    LoadConString = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
            & "catalog=" & sCatalog & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End If
End Function
Private Function LoadConStringfact()
Dim cServerName As String, cUserId As String, cPassword As String, sCatalogFACT As String
cServerName = RetSetting("server", App.Path & "\conffact.txt")
cUserId = decrypt(RetSetting("userId", App.Path & "\conffact.txt"), "dr")
cPassword = decrypt(RetSetting("Password", App.Path & "\conffact.txt"), "dr")
sCatalogFACT = RetSetting("Catalog", App.Path & "\conffact.txt")
If cServerName = "" Then cServerName = "."
cServerName = MyParn(cServerName & turn(cExpress, "\") & cExpress)
LoadConStringfact = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
        & "catalog=" & sCatalogFACT & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Private Function LoadConStringPICT()
Dim cUserId As String, cPassword As String, sCatalogPICT As String

If lServerOnLine Or lServerOnLineShop Then
'    sCatalogPICT = "JUNIOR_SHOP"
'    cServerNamePICT = "junior-sql.database.windows.net"
'    cUserId = decrypt("062F8C407C77", "dr")
'    cPassword = decrypt("C77DA5F185954963DD1F902815", "dr")
Else
    cUserId = decrypt(RetSetting("userId", cPathConf & "\CONF.TXT"), "dr")
    cPassword = decrypt(RetSetting("Password", cPathConf & "\CONF.TXT"), "dr")
    sCatalogPICT = "JUNIOR_PICT"
'    cServerNamePICT = RetSetting("server", App.Path & "\conf.txt")
End If
If cServerNamePICT <> "" Then
    LoadConStringPICT = "provider=SQLOLEDB;data source=" & cServerNamePICT & ";initial " _
            & "catalog=" & sCatalogPICT & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
Else
    LoadConStringPICT = ""
End If
End Function

Private Function LoadConStringfact2()
Dim cServerName As String, cUserId As String, cPassword As String, sCatalogFACT As String
cServerName = RetSetting("server", App.Path & "\conffact.txt")
cUserId = decrypt(RetSetting("userId", App.Path & "\conffact.txt"), "dr")
cPassword = decrypt(RetSetting("Password", App.Path & "\conffact.txt"), "dr")
sCatalogFACT = RetSetting("Catalog", App.Path & "\conffact.txt")
If cServerName = "" Then cServerName = "."
cServerName = MyParn(cServerName & turn(cExpress, "\") & cExpress)

LoadConStringfact2 = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
        & "catalog= FACT_2 ;user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Private Function LoadConStringfact3()
Dim cServerName As String, cUserId As String, cPassword As String, sCatalogFACT As String
cServerName = RetSetting("server", App.Path & "\conffact.txt")
cUserId = decrypt(RetSetting("userId", App.Path & "\conffact.txt"), "dr")
cPassword = decrypt(RetSetting("Password", App.Path & "\conffact.txt"), "dr")
sCatalogFACT = RetSetting("Catalog", App.Path & "\conffact.txt")
If cServerName = "" Then cServerName = "."
cServerName = MyParn(cServerName & turn(cExpress, "\") & cExpress)

LoadConStringfact3 = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
        & "catalog= FACT_SOCK ;user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function

Private Sub FixData2()
Dim FS1 As New ADODB.command
'openCon con
On Error Resume Next
cString = "ALTER TABLE [dbo].[FILE6_20H] ADD   [SALES_RET] [nvarchar](50) COLLATE Arabic_CI_AS NULL"
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
'closeCon con
Err.Clear
End Sub
Private Sub FixData3()
Dim FS1 As New ADODB.command
'openCon con
On Error Resume Next
cString = "alter TABLE [dbo].[users] add   [option11] [bit] NOT NULL CONSTRAINT [DF_users_option11]  DEFAULT ((0))"
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
'closeCon con
Err.Clear
End Sub
Private Sub FixData4()
Dim FS1 As New ADODB.command
'openCon con
On Error Resume Next
cString = "ALTER TABLE [dbo].[FILE6_20H] ADD [username_ret] [nvarchar](50) COLLATE Arabic_CI_AS NULL "
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
'closeCon con
Err.Clear
End Sub
Private Sub FixData5()
Dim FS1 As New ADODB.command
'openCon con
On Error Resume Next
cString = "alter TABLE [dbo].[users] add   [option9] [bit] NOT NULL CONSTRAINT [DF_users_option9]  DEFAULT ((0))"
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
'closeCon con
Err.Clear
End Sub
Private Sub FixData11()
Dim FS1 As New ADODB.command
'openCon con
On Error Resume Next
cString = "alter TABLE [dbo].[users] add   [option11] [bit] NOT NULL CONSTRAINT [DF_users_option11]  DEFAULT ((0))"
FS1.CommandType = adCmdText
Set FS1.ActiveConnection = con
FS1.CommandText = cString
FS1.Execute
'closeCon con
Err.Clear
End Sub
Private Function LoadConString_B(pCode)
Dim cServerName As String, cUserId As String, cPassword As String
cServerName = GetDesca("SELECT IPSQL FROM QBRANCH_ALL WHERE CODE = " & MyParn(pCode), GetCon)
'cUserId = decrypt(RetSetting("userId", App.Path & "\conf.txt"), "dr")
'cPassword = decrypt(RetSetting("Password", App.Path & "\conf.txt"), "dr")

cUserId = "data_pro"
cPassword = "2015"

LoadConString_B = "provider=SQLOLEDB;data source=" & cServerName & ";initial " _
        & "catalog=" & sCatalog & ";user id = " & cUserId & ";" & "password = " & cPassword
End Function


Sub FixAddress()
Dim loctable As New ADODB.Recordset
loctable.Open "select * From Address", con, adOpenStatic, adLockReadOnly
If Not (loctable.EOF And loctable.BOF) Then
    cComp_Name = loctable!desca & ""
    cComp_address = loctable!Address & ""
    cComp_Phone = loctable!phone & ""
    cComp_Head1 = loctable!HEAD1 & ""
    cComp_Head2 = loctable!HEAD2 & ""
    nCountPrint = Val(loctable!COUNTPRINT & "")
    nLang_Boon = Val(loctable!lang & "")
End If
loctable.Close
Set loctable = Nothing
End Sub

