VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsTBalSupl 
   Caption         =   "√—’œ… Ê  ⁄«„·«  «·„Ê—œÌ‰"
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1365
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   3840
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2475
         Picture         =   "VsTBalSupl.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "VsTBalSupl.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "VsTBalSupl.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   15
         Top             =   630
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   820
         _Version        =   196610
         BackColor       =   13822956
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "VsTBalSupl.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   45
         TabIndex        =   16
         Top             =   1125
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7620
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGrCust 
         Height          =   315
         Left            =   6000
         TabIndex        =   7
         Top             =   615
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   1125
         TabIndex        =   9
         Top             =   630
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   675
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄… „Ê—œÌ‰:"
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
         Index           =   0
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   9555
         TabIndex        =   5
         Top             =   270
         Width           =   675
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "≈·Ï  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   4
         Top             =   285
         Width           =   735
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTBalSupl.frx":9583
      Height          =   8190
      Left            =   90
      TabIndex        =   6
      Top             =   1395
      Width           =   15000
      _cx             =   26458
      _cy             =   14446
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      BackColorSel    =   16777088
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
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2475
      Top             =   75
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   150
      Top             =   75
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
Attribute VB_Name = "VsTBalSupl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cString As String
Dim cStr1 As String, cStr2 As String
Dim con As New ADODB.Connection

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ √—’œ… Ê  ⁄«„·«  «·„Ê—œÌ‰"
    If xGrCust.text <> "" Then cHead1 = cHead1 & xGrCust.text
    cHead2 = " ·„Ê”„ " & xMosm.text
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.text, "DD-MM-YYYY") & " ≈·Ï  «—ÌŒ " & Format(xdate2.text, "DD-MM-YYYY")
    
'    Set printGrd3.myForm = Me
'    printGrd3.doprint 0.9, -2, cHead1, cHead2, , False, True, 8
'    printGrd3.Show 1


    Load PrintGrd
    PrintGrd.doprint grid1, 1, , cHead1, cHead2, , , True, 8
    PrintGrd.Show 1

End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    If xMosm.BoundText = "" Or Not IsDate(xDate1.text) Or Not IsDate(xdate2.text) Then
        MsgBox "ÌÃ»  ÕœÌœ „Ê”„ - «· «—ÌŒ "
        Exit Sub
    End If
    myload
End Sub
Private Sub Form_Load()
    openCon con
    xDate1.text = GetDesca("select min(date) from file4_11 where mosm =  " & MyParn(cPMosm), con)
    xdate2.text = Format(Date, "dd-mm-yyyy")
    
    data1.ConnectionString = strCon
    data1.RecordSource = "SELECT * FROM FILE4_50 "
    Set xGrCust.RowSource = data1
    xGrCust.ListField = "Desca"
    xGrCust.BoundColumn = "Code"
    
    data3.ConnectionString = strCon
    data3.RecordSource = "SELECT * FROM MOSM"
    Set xMosm.RowSource = data3
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "mosm"
    xMosm.BoundText = cPMosm
    
    Set grid1.DataSource = DATA2
    DATA2.ConnectionString = strCon
    
    FIXGRID
    grid1.Rows = 1
End Sub
Private Sub myload()
Dim cString As String

cWhere = " MOSM =  " & MyParn(xMosm.BoundText)

cField2 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '1' )", "[SAL]") & " AS F_BAL "

'cField2 = "(select sum(sal - pay) from file4_11 as file4_112 where code = file4_10.code and mosm = " & MyParn(xMosm.BoundText) & " and date < " & DateSq(xDate1.Text) & " )"

cWhere = "file4_11.date >= " & DateSq(xDate1.text)

cField3 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '4' )", "[SAL]") & " AS T_SALES"
cField4 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '5' )", "(PAY)") & " AS T_RET"
cField5 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '4' OR TYPE = '5' )", "[sal]- [PAY]") & " AS net_sal"

cField6 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '7' or TYPE = '8'  )", "[PAY] - [SAL]") & " AS T_CASH"
cField7 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = 'A' OR TYPE = 'C' )", "[PAY]") & " AS T_CHQ"
cField8 = myiif(cWhere & turn(cWhere, " And ") & " (TYPE = '9')", "[PAY]") & " AS T_CASH0 "

cWhere = ""
cField9 = myiif(cWhere, "[sal]- [pay] ") & " AS C_BAL"

cField10 = "(select sum(value) from file5_21 WHERE code1 = file4_10.code  and mosm = " & MyParn(xMosm.BoundText) & " and CLOSED = '0' )"
cField11 = "(SELECT  sum(Items_Sales.TOTAL * ((100 - Items_Sales.discount) / 100) )  from Items_Sales inner join file1_10 on file1_10.item = Items_Sales.item  where file1_10.code = file4_10.code and file1_10.mosm = " & MyParn(xMosm.BoundText) & " and date  >= " & DateSq(xDate1.text) & " and date <= " & DateSq(xdate2.text) & " ) "
cField12 = "(SELECT  sum(([in]-[out]) * file1_10.cost )   from file1_11 inner join file1_10 on file1_10.item = file1_11.item  where file1_10.code = file4_10.code and file1_10.mosm = " & MyParn(xMosm.BoundText) & " and date <= " & DateSq(xdate2.text) & " ) "



With grid1
'                           0                           1
    cString = "  select FILE4_10.CODE AS CODE , FILE4_10.DESCA AS DESCA , " & _
                cField2 & " , " & cField3 & " , " & cField4 & " , " & cField5 & " , " & _
                cField6 & " , " & cField7 & " , " & cField8 & " , " & cField9 & " , " & cField10 & " , " & cField11 & " , " & cField12 & _
                " FROM (FILE4_11 inner JOIN file4_10 ON FILE4_11.CODE = FILE4_10.CODE) LEFT JOIN file3_50 ON FILE4_10.[group] = file3_50.CODE WHERE FILE4_10.CODE IS NOT NULL "
    If xGrCust.BoundText <> "" Then cString = cString & turn(cString) & " file4_10.[group]  = " & MyParn(xGrCust.BoundText)
    If xMosm.BoundText <> "" Then cString = cString & turn(cString) & " FILE4_11.[MOSM] = " & MyParn(xMosm.BoundText)
'   If IsDate(xDate1.Text) Then cString = cString & turn(cString) & " FILE4_11.date >= " & DateSq(xDate1.Text)
    If IsDate(xdate2.text) Then cString = cString & turn(cString) & " FILE4_11.date <= " & DateSq(xdate2.text)
    cString = cString & " GROUP BY FILE4_10.DESCA , FILE4_10.CODE "
    DATA2.RecordSource = cString
    DATA2.Refresh
End With
FIXGRID

End Sub
Sub FIXGRID()
    With grid1
    .Cols = 13
    .RowHeight(0) = 1000
    .WordWrap = True
    .FrozenCols = 2
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    
    .TextMatrix(0, 2) = "—’Ìœ √Ê· "
    .TextMatrix(0, 3) = "Ã. „‘ —Ì« "
    .TextMatrix(0, 4) = "Ã. „— Ã⁄«  "
    .TextMatrix(0, 5) = "’«›Ï „‘ —Ì« "
    .TextMatrix(0, 6) = " ”œ«œ ‰ﬁœÏ"
    .TextMatrix(0, 7) = "”œ«œ √.ﬁ "
    .TextMatrix(0, 8) = " ”ÊÌ« "
    .TextMatrix(0, 9) = "—’Ìœ Õ«·Ï"
    .TextMatrix(0, 10) = "‘Ìﬂ«  €Ì— „”œœ…"
    .TextMatrix(0, 12) = "ﬁÌ„… —’Ìœ «·«’‰«›"
    .TextMatrix(0, 11) = "ﬁÌ„… «·„»Ì⁄« "
    
    
    .ColWidth(0) = 800
    .ColWidth(1) = 2000
    .ColWidth(2) = 1200
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    .ColWidth(12) = 1200
    
    For nCol = 2 To .Cols - 1
        .ColDataType(nCol) = flexDTDouble
        .ColFormat(nCol) = "#0.00"
    Next nCol
    
    .SubtotalPosition = flexSTAbove
    For i = 2 To .Cols - 2
        .Subtotal flexSTSum, -1, i, "#0", vbRed, vbYellow, True, "  "
    Next i
'    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
Private Sub grid1_DblClick()
    If grid1.Row > 1 Then
        Load supMovefrm
        supMovefrm.xCode.text = grid1.TextMatrix(grid1.Row, 0)
        supMovefrm.XDESCA.Caption = grid1.TextMatrix(grid1.Row, 1)
        supMovefrm.xMosm.Caption = xMosm.BoundText
        supMovefrm.Fillgrd
        supMovefrm.LoadSuplData
        supMovefrm.Show
    End If
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
xCustName.Caption = ""
If xCode.text = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 6)

xCustName.Caption = GetDesca("select desca from FILE4_10 where code = " & MyParn(xCode.text), con) & ""
End Sub
Sub myProc()
ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Unload Search3
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From FILE4_10"
Generalarray(2) = "Order by file4_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub
Private Sub xMosm_LostFocus()
    xDate1.text = GetDesca("select min(date) from file4_11 where mosm =  " & MyParn(xMosm.BoundText), con)

End Sub
