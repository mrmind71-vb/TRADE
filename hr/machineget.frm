VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{FE9DED34-E159-408E-8490-B720A5E632C7}#1.0#0"; "zkemkeeper.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form machineGetFrm 
   Caption         =   "”Õ» «·»Ì«‰« "
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10545
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGetAll 
      Caption         =   "”Õ» «·›—Ê⁄"
      Height          =   600
      Left            =   14625
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   225
      Width           =   1410
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "«⁄«œ… ”Õ»"
      Enabled         =   0   'False
      Height          =   600
      Left            =   13230
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   225
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Frame Frame2 
      Height          =   825
      Left            =   16065
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   0
      Width           =   4245
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   360
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   270
         Visible         =   0   'False
         Width           =   1365
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   390
         Left            =   1755
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "„‰  «—ÌŒ"
         Height          =   285
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   315
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   0
      TabIndex        =   4
      Top             =   10260
      Visible         =   0   'False
      Width           =   3345
      Begin VB.OptionButton optNetconncet 
         Appearance      =   0  'Flat
         Caption         =   "NetConnect"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   225
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   1890
      End
      Begin VB.OptionButton optCom 
         Appearance      =   0  'Flat
         Caption         =   "ComConnect"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   225
         TabIndex        =   5
         Top             =   765
         Width           =   1785
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   10230
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   90
      TabIndex        =   7
      Top             =   11475
      Visible         =   0   'False
      Width           =   3345
      Begin VB.TextBox txtRate 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1665
         TabIndex        =   13
         Text            =   "115200"
         Top             =   1395
         Width           =   1185
      End
      Begin VB.TextBox txtPort 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1665
         TabIndex        =   12
         Text            =   "4370"
         Top             =   990
         Width           =   1185
      End
      Begin VB.TextBox txtMachNum 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   1665
         TabIndex        =   10
         Text            =   "1"
         Top             =   630
         Width           =   555
      End
      Begin VB.TextBox txtComNum 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   1665
         TabIndex        =   8
         Text            =   "1"
         Top             =   225
         Width           =   510
      End
      Begin VB.Label labPort 
         Alignment       =   1  'Right Justify
         Caption         =   "Port :"
         Height          =   255
         Left            =   975
         TabIndex        =   19
         Top             =   1035
         Width           =   600
      End
      Begin VB.Label labRare 
         Alignment       =   1  'Right Justify
         Caption         =   "Rate :"
         Height          =   255
         Left            =   900
         TabIndex        =   14
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label labMachNum 
         Alignment       =   1  'Right Justify
         Caption         =   "Machine No :"
         Height          =   375
         Left            =   150
         TabIndex        =   11
         Top             =   675
         Width           =   1425
      End
      Begin VB.Label labCom 
         Alignment       =   1  'Right Justify
         Caption         =   "ComNumber :"
         Height          =   345
         Left            =   60
         TabIndex        =   9
         Top             =   270
         Width           =   1515
      End
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   4275
      Top             =   180
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   12645
      Top             =   1800
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
   Begin Threed.SSCommand cmdConnect 
      Height          =   375
      Left            =   90
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   9360
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   661
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Connect Server"
      ButtonStyle     =   3
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   22
      Top             =   10080
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   8475
      Left            =   15705
      TabIndex        =   23
      Top             =   855
      Width           =   4560
      _cx             =   8043
      _cy             =   14949
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   -1  'True
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
   Begin Threed.SSCommand cmdExit 
      Height          =   600
      Left            =   10530
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   225
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   1058
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "machineget.frx":0000
      Alignment       =   8
      PictureAlignment=   6
   End
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Height          =   8475
      Left            =   10575
      TabIndex        =   18
      Top             =   855
      Width           =   5100
      _cx             =   8996
      _cy             =   14949
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   2
      ScrollTips      =   -1  'True
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
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   1395
      Top             =   225
      Visible         =   0   'False
      Width           =   1590
      _ExtentX        =   2805
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
   Begin zkemkeeperCtl.CZKEM CZKEM1 
      Height          =   870
      Left            =   2835
      OleObjectBlob   =   "machineget.frx":242E
      TabIndex        =   25
      Top             =   2070
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label labSDK 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   630
      TabIndex        =   3
      Top             =   1260
      Visible         =   0   'False
      Width           =   1695
   End
End
Attribute VB_Name = "machineGetFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection, con2 As New ADODB.Connection, bActiviated As Boolean
Dim sMachine As String, sMachineIp As String
Dim sDate_First As String
Private Sub cmdClear_Click()
MachineFix.Show 1
End Sub

Private Sub cmdConnect_Click()
ConnectServer con, Me
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub cmdGetAll_Click()
Me.MousePointer = vbHourglass
If GetBranchData Then
    Inform " „ ”Õ» «·Õ÷Ê— Ê«·«‰’—«› »‰Ã«Õ "
    myloadgrd2
End If
Me.MousePointer = vbNormal
End Sub
Private Sub cmdLoad_Click()
'Dim i As Long
'For i = 1 To grdBranch.Rows - 1
'    getIpData i
'Next
'If Not myValidGet Then Exit Sub
'myLoadGrd2
'If GetData Then MyLoadgrd
End Sub

Private Sub Command1_Click()
LoadConString
End Sub

Private Sub Command2_Click()
LoadConStringHr
End Sub

Private Sub Form_Load()
Dim sError As String
On Error GoTo myerror
sError = openConHr(con)

If sError <> "ok" Then
    MsgBox sError
    Exit Sub
End If

Dim loctable As New ADODB.Recordset
Set loctable = myCmd("select *  from hr.machines where store = " & MyParn(sStore), con)
If Not loctable.EOF Then
    sMachine = loctable!CODE & ""
    sMachineIp = loctable!Ip & ""
End If

If sMachine = "" Or sMachineIp = "" Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ﬂ«›Ì… ·”Õ» «·»Ì«‰« "
    Exit Sub
End If
    


xDate1.text = myFormat_p(myField("SELECT TOP 1 hr.SALARY_H.DATE1 FROM hr.SALARY_H WHERE CLOSED = 0 ORDER BY DOC_NO DESC", con))
xDate1.Tag = myFormat(xDate1.text)

xDate2.text = myFormat(Date)

Set GRID2.DataSource = DATA11
Set grid1.DataSource = data10

myload

Exit Sub
myerror:
bError = True
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myload()
Me.MousePointer = 11
myloadgrd2
myloadgrd
Me.MousePointer = 0
End Sub
Private Function myloadgrd2() As Boolean
Dim cString As String, aret As Variant, cWhere As String
Dim loctable  As ADODB.Recordset

Dim aPrm As Variant
aPrm = AddFlag(aPrm, "USERID", Null)
aPrm = AddFlag(aPrm, "FTIME1", myBeginDate(xDate1.text, 3))
aPrm = AddFlag(aPrm, "FTIME2", myLastDate(xDate2.text, 2))

Set DATA11.Recordset = myCmd("hr.sp_check_time", con, adStoredProc, aPrm, 300)
prog1.Visible = False
myloadgrd2 = True
Fixgrd2
End Function
Private Function myloadgrd() As Boolean
Dim cString As String, aret As Variant, cWhere As String
Dim loctable  As ADODB.Recordset

Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DATE1", myBeginDate(xDate1.text, 3))
aPrm = AddFlag(aPrm, "DATE2", myLastDate(xDate2.text, 2))

Set data10.Recordset = myCmd("hr.SP_ARRIVE_DAY_TOTAL", con, adStoredProc, aPrm, 300)
myloadgrd = True
Fixgrd
End Function
Private Function myreplace() As Boolean
'Dim cTime As String, cId As String, cString As String
'Dim nRecord As Long
'On Error GoTo myerror
'con.BeginTrans
'With grid1
'prog1.Visible = True
'prog1.Value = 0
'
'cString = "DELETE FROM MACHINE"
'For i = 1 To .Rows - 1
'    prog1.Value = Round(i / (grid1.Rows - 1), 2) * 100
'    cTime = Format(.TextMatrix(i, 4), "YYYY-MM-DD HH:NN")
'    If IsDate(cTime) And IsNumeric(.TextMatrix(i, 1)) Then
'        con.Execute cString
'
'        cString = "INSERT INTO MACHINE(CODE,[DATE])"
'        cString = cString & "Values("
'        cString = cString & addvalue(.TextMatrix(i, 1))
'        cString = cString & "," & addstring(cTime)
'        cString = cString & ")"
'        con.Execute cString
'    End If
'Next
'End With
'con.CommitTrans
'prog1.Visible = False
'
'myLoadGrd
'
'myreplace = myreplaceGrd
'Exit Function
'myerror:
'prog1.Visible = False
'MsgBox Err.Description
'con.RollbackTrans
'Err.Clear
End Function
Private Sub Fixgrd2()
With GRID2
.Cols = 5
.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "—ﬁ„ «·„ÊŸ›"
.TextMatrix(0, 2) = "«· «—ÌŒ"
.TextMatrix(0, 3) = "id"
.TextMatrix(0, 4) = "«·ÌÊ„"
.ColHidden(3) = True

.RowHeight(0) = 600
.ColWidth(0) = 700
.ColWidth(1) = 1200
.ColWidth(2) = 1300
.ColWidth(3) = 900
.ColWidth(4) = 900

For i = 1 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

Dim sDate As String
For i = 1 To .Rows - 1
    If sDate <> myFormat(.TextMatrix(i, 2)) Then
        If sDate <> "" Then
            .Select i - 1, 0, i - 1, .Cols - 1
            .CellBorder vbBlack, 0, 0, 0, 2, 0, 1
        End If
        sDate = myFormat(.TextMatrix(i, 2))
    End If
Next

For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
    .TextMatrix(i, 4) = arbDay(GRID2.TextMatrix(i, 2))
Next
End With
End Sub
Private Sub Fixgrd()
With grid1
.TextMatrix(0, 0) = "„”·”·"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "⁄œœ «·„ÊŸ›Ì‰"
.TextMatrix(0, 3) = "⁄œœ «·»’„« "

.RowHeight(0) = 600
.ColWidth(0) = 700
.ColWidth(1) = 1300
.ColWidth(2) = 1300
.ColWidth(3) = 1300

For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
Next

.SubtotalPosition = flexSTAbove
.Subtotal flexSTSum, -1, 2, "#0", &HC0FFC0, vbBlack, True, "«Ã„«·Ì"
.Subtotal flexSTSum, -1, 3, "#0", , &HC0FFC0, vbBlack, True, "«Ã„«·Ì"

For i = 1 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
Set machineGetFrm = Nothing
End Sub
Private Sub grdBranch_DblClick()
End Sub
Private Sub grdBranch_EnterCell()
With grdBranch
    If .Col = 4 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub

Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'grid1.SaveGrid App.Path & "\" & Me.Name & ".grd", flexFileData
End Sub
Private Sub xDate1_Validate(Cancel As Boolean)
myValidDate xDate1
If myFormat(xDate1.text) < myFormat(sDate_Period) Then
    xDate1.text = myFormat_p(sDate_Period)
End If
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
myValidDate xDate2
If myFormat(xDate2.text) < myFormat(xDate1.Tag) Then
    xDate2.text = myFormat_p(xDate2.Tag)
End If
End Sub
Private Sub MakeSerial(Optional nBeginRow As Long = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Function MYVALID() As Boolean
If GRID2.Rows = 1 Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·Õ›ŸÂ«"
    Exit Function
End If
MYVALID = True
End Function
Private Function myValidGet(Optional bCheck As Boolean = False) As Boolean
myValidGet = True
End Function
Private Function getIpData() As Boolean
On Error GoTo myerror
Me.MousePointer = 11
Inform "Ì „ «·« ’«· »«·›—⁄"
If Not connect(sMachineIp) Then GoTo myerror
Inform "Ì „ ”Õ» «·»Ì«‰«  „‰ «·›—⁄"
If Not GetDataIp(sMachine) Then GoTo myerror
Inform " „ ”Õ» «·»Ì«‰«  „‰ «·›—⁄ »‰Ã«Õ"
Me.MousePointer = 0
myload
getIpData = True
Exit Function
myerror:
Me.MousePointer = 0
End Function
Private Function connect(pIp) As Boolean
Dim bconn As Boolean
Dim mint As Integer
Dim SDKVersion As String
Dim strVersion As String
On Error GoTo myerror:

CZKEM1.BASE64 = 1
bconn = False
CZKEM1.GetSDKVersion SDKVersion
labSDK.Caption = "SDKVersion:" + SDKVersion
   
If txtPort.text = "" Then Exit Function
bconn = CZKEM1.Connect_Net(CStr(pIp), CLng(txtPort.text))
If bconn Then
    Inform "« ’«· ‰«ÃÕ"
    CZKEM1.Beep 150
Else
    Inform "« ’«· €Ì— ‰«ÃÕ"
    Exit Function
End If
connect = True
myerror:
MsgBox Err.Description
End Function
Private Function GetDataIp(pMachine As String) As Boolean
Dim dwEnrollNmber As Long
Dim dwEnrollString As String
Dim dwverifyMode As Long
Dim dwInOutMode As Long
Dim dwYear As Long
Dim dwMonth As Long
Dim dwDay As Long
Dim dwHour As Long
Dim dwMinute As Long
Dim dwSecond As Long
Dim dwWorkCode As Long
Dim dwDate As Date
Dim iDate As Date
Dim dwTime As Date
Dim i As Long
Dim sDate As String
Dim nRecord As Long
Dim pMrd As Long
Dim sString As String
Dim cInsert As String
sCaption = Me.Caption
Dim nMaxRows As Long

On Error GoTo myerror
nMaxRows = 100
Inform "”Ê› Ì „  Õ„Ì· «·»Ì«‰« "
If CZKEM1.ReadAllGLogData(1) Then
    Inform " „  Õ„Ì· «·»Ì«‰« "
    If CZKEM1.IsTFTMachine(1) Then
        Inform "Ì „  ”ÃÌ· «·»Ì«‰« "
        Do While CZKEM1.SSR_GetGeneralLogData(1, dwEnrollString, dwverifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute, dwSecond, dwWorkCode)
            i = i + 1
            dwDate = DateSerial(dwYear, dwMonth, dwDay)
            dwTime = TimeSerial(dwHour, dwMinute, 0)
            sDate = Format(dwYear & "-" & dwMonth & "-" & dwDay & " " & dwHour & ":" & dwMinute, "yyyy-mm-dd hh:nn")
            If myFormat(sDate) >= myFormat(xDate1.text) Then
                If nRow >= nMaxRows Then
                    nRow = 0
                    aString = AddFlag(aString, cInsert)
                    cInsert = ""
                End If
                nRow = nRow + 1
                                
                pKey = Format(sDate, "YYYY-MM-DD HH:NN") & "-" & pMachine & "-" & dwEnrollString
                
                cInsert = cInsert & "INSERT INTO hr.CHECKTIME(USERID,[FTIME],MRD,FDATE)" & _
                        " SELECT " & _
                        addvalue(dwEnrollString & "") & _
                        "," & addDate(sDate) & _
                        "," & pMachine & _
                        "," & Format(sDate, "YYYYMMDD") & _
                        " Where hr.fn_checktime_found(" & addstring(pKey) & ") = 0" & _
                        ";"
            End If
        Loop
        If Trim(cInsert) <> "" Then
            aString = AddFlag(aString, cInsert)
        End If
    Else
        Do While CZKEM1.GetAllGLogData(1, 1, dwEnrollNmber, 1, dwverifyMode, dwInOutMode, dwYear, dwMonth, dwDay, dwHour, dwMinute)
            i = i + 1
            dwDate = DateSerial(dwYear, dwMonth, dwDay)
            dwTime = TimeSerial(dwHour, dwMinute, 0)
            sDate = Format(dwYear & "-" & dwMonth & "-" & dwDay & " " & dwHour & ":" & dwMinute, "yyyy-mm-dd hh:nn")
            If myFormat(sDate) >= myFormat(xDate1.text) Then
                If Not Is_Found_check(dwEnrollNmber & "", sDate, Val(pMachine), con) Then
                    pKey = addDate(sDate) & "-" & pMachine & "-" & dwEnrollString
                    cString = "INSERT INTO hr.CHECKTIME(USERID,[FTIME],MRD,FDATE)" & _
                               "SELECT " & addvalue(dwEnrollNmber) & _
                               "," & addDate(sDate) & _
                               "," & addvalue(pMachine) & _
                               "," & Format(sDate, "YYYYMMDD") & _
                               " WHERE hr.fn_ftime_mrd(" & MyParn(pKey) & ") = 0"
                    con.Execute cString, nRecord
                End If
            End If
            Me.Caption = sCaption & "-" & i
        Loop
    End If
End If
If Not IsEmpty(aString) Then
    prog1.Visible = True
    prog1.Value = 0
    con.BeginTrans
    On Error GoTo myerror
    For i = 0 To UBound(aString)
        prog1.Value = Round((i + 1) / (UBound(aString) + 1), 2) * 100
        con.Execute aString(i)
    Next
    con.CommitTrans
    prog1.Visible = False
End If
GetDataIp = True
Me.Caption = sCaption
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
Me.Caption = sCaption
prog1.Visible = False
End Function
Private Sub CZKEM1_OnAttTransaction(ByVal EnrollNumber As Long, ByVal IsInValid As Long, ByVal AttState As Long, ByVal VerifyMethod As Long, ByVal Year As Long, ByVal Month As Long, ByVal Day As Long, ByVal Hour As Long, ByVal Minute As Long, ByVal Second As Long)
    'MsgBox "It is " & EnrollNumber & ",welcome!", vbOKOnly, "Even"
    DoEvents
End Sub

'Private Sub CZKEM1_OnEnrollFinger(ByVal EnrollNumber As Long, ByVal FingerIndex As Long)
'    'MsgBox "user " & EnrollNumber & " enroll number of " & FingerIndex + 1 & " finger"
'    DoEvents
'End Sub

Private Sub CZKEM1_OnKeyPress(ByVal key As Long)
    'MsgBox "Key:" & Key & ""
    DoEvents
End Sub

Private Sub CZKEM1_OnNewUser(ByVal EnrollNumber As Long)
    'MsgBox "You are enrolling" & EnrollNumber & ""
    DoEvents
End Sub

Private Sub CZKEM1_OnVerify(ByVal UserID As Long)
    'MsgBox "" & UserID & " is verify"
    DoEvents
End Sub
Private Function MissingRow(Row As Long) As Boolean
With grid1
If (Trim(.TextMatrix(Row, 8)) <> "" And Trim(.TextMatrix(Row, 11)) <> "") Or (Trim(.TextMatrix(Row, 9)) <> "" And Trim(.TextMatrix(Row, 10)) <> "") Then Exit Function
If Trim(.TextMatrix(Row, 8)) = "" And Trim(.TextMatrix(Row, 11)) = "" And Trim(.TextMatrix(Row, 9)) = "" And Trim(.TextMatrix(Row, 10)) = "" Then Exit Function
End With
MissingRow = True
End Function
Private Function GetBranchData() As Boolean
With grdBranch
If (Not getIpData()) Then
    Exit Function
End If
GetBranchData = True
End With
End Function
Public Function ConnectServer(ByRef con As ADODB.Connection, myForm As Form) As Boolean
Dim sError As String

myForm.MousePointer = vbHourglass
sError = openConHr(con)

myForm.MousePointer = vbNormal
If sError = "ok" Then
    MsgBox " „ «·« ’«· »ﬁ«⁄œ… «·»Ì«‰« "
Else
    MsgBox cError
End If
End Function
Private Sub myValidDate(ByRef pControl As Variant)
If IsDate(pControl.text) Then pControl.text = myFormat_p(pControl.text) Else pControl.text = ""
End Sub
Function Is_Found_check(pUserId As String, pTime As String, pMrd As Long, pCon As ADODB.Connection) As Boolean
Dim rdTable As New ADODB.Recordset
Dim cmdTable As New ADODB.Command
Set cmdTable.ActiveConnection = pCon
cmdTable.CommandType = adCmdStoredProc
cmdTable.Parameters.Append cmdTable.CreateParameter("USERID", adInteger, adParamInput, 6, pUserId)
cmdTable.Parameters.Append cmdTable.CreateParameter("FTIME", adVarChar, adParamInput, 16, pTime)
cmdTable.Parameters.Append cmdTable.CreateParameter("MRD", adInteger, adParamInput, 10, pMrd)
cmdTable.CommandText = "hr.FOUND_CHECK"
Set rdTable = cmdTable.Execute
If Not (rdTable.EOF And rdTable.BOF) Then Is_Found_check = True
Set rdTable = Nothing
Set cmdTable = Nothing
End Function
Function myFormat(sDate As Variant, Optional bLong As Boolean = False) As String
myFormat = Format(sDate, "YYYY-MM-DD") & IIf(bLong, Format(sDate & "", " HH:NN"), "")
End Function
Function DateSq(ByVal X As Variant) As String
If Not IsDate(X) Then
    DateSq = X
    Exit Function
End If
DateSq = MyParn(Format(X, "YYYY-MM-DD HH:NN"))
End Function
Function addDate(pValue) As String
If Not IsDate(pValue) Then
    addDate = "NULL"
Else
    addDate = DateSq(pValue)
End If
End Function
Function myLastDate(sDate As String, Optional nHour As Integer = 2)
myLastDate = myFormat(DateAdd("d", 1, myFormat(sDate))) & " " & RetZero(nHour, 2) & ":" & "59"
End Function
Function myBeginDate(sDate As String, Optional nHour As Integer = 2)
myBeginDate = myFormat(sDate) & " " & RetZero(nHour, 2) & ":" & "00"
End Function

