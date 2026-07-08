VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Target_Month_Branch 
   Caption         =   "ÅÌãÇáì ÇáãÓÊåÏÝ ÇáÔåÑì ááÝÑÚ ÎáÇá ÝÊÑÉ"
   ClientHeight    =   9405
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   13380
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9405
   ScaleWidth      =   13380
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "ÚÏÏ ÇáÓÌáÇÊ ÇáãØÇÈÞÉ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   315
      Width           =   2040
      Begin VB.Label xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   1860
      End
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   2205
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   3615
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "Target_Month_Branch.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Enabled         =   0   'False
         Height          =   510
         Left            =   1230
         Picture         =   "Target_Month_Branch.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   2385
         Picture         =   "Target_Month_Branch.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "ÚÑÖ"
         Top             =   135
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   5850
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   9060
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6300
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1635
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6300
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1635
      End
      Begin MSDataListLib.DataCombo xbranch 
         Height          =   315
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ÇáÝÑÚ :"
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
         Index           =   5
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   225
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ÍÊì ÊÇÑíÎ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   8010
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ãä ÊÇÑíÎ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   8010
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   2520
      Top             =   405
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   2040
      Top             =   -15
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
      Left            =   840
      Top             =   -15
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   60
      Top             =   120
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8205
      Left            =   135
      TabIndex        =   11
      Top             =   1080
      Width           =   14775
      _cx             =   26061
      _cy             =   14473
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   20
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   12
      Top             =   9165
      Visible         =   0   'False
      Width           =   13380
      _ExtentX        =   23601
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "Target_Month_Branch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cString As String
Private Sub cmdPrint_Click()
Dim cHead1 As String
Dim cHead2 As String
cHead1 = "äÞÑíÑ ÇáãÓÊåÏÝ ÇáÔåÑì áÝÑÚ " & XBRANCH.text
If IsDate(xdate1.text) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "ãä : " & Format(xdate1.text, "dd-mm-yyyy")
If IsDate(xDate2.text) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "ÍÊí : " & Format(xDate2.text, "dd-mm-yyyy")
Load PrintGrd
PrintGrd.doprint Grid1, 1.02, 0, cHead1, cHead2, "", False, False, 10
PrintGrd.Show 1
End Sub
Private Sub cmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
openCon con

data2.ConnectionString = strCon
Set data2.Recordset = myRecordSet("Select Code,DescA From STORE_BR WHERE " & cUserStore, con)
Set XBRANCH.RowSource = data2
XBRANCH.ListField = "Desca"
XBRANCH.BoundColumn = "code"

Set Grid1.DataSource = DATA1
DATA1.ConnectionString = strCon

Fixgrd
End Sub
Private Sub myload()
Dim cFiled1 As String, cFiled2 As String
Dim dDate11 As Date, dDate22 As Date

cFiled1 = ",(SELECT  SUM(T_TOTAL - DISCOUNT)   FROM T_SALESDOC_ALL WHERE T_SALESDOC_ALL.STORE =  " & MyParn(XBRANCH.BoundText) & " AND DATE >= FILE6_80H.DATE1  AND DATE <= FILE6_80H.DATE2 )"

With Grid1
'                                                   0                           1       2         3
    cString = " SELECT  YEAR(FILE6_80H.DATE1) ,  MONTH( FILE6_80H.DATE1),DATE1,DATE2, FILE6_80.VALUE_TARGET " & cFiled1 & _
              " FROM    FILE6_80H INNER JOIN FILE6_80 ON FILE6_80H.DOC_NO = FILE6_80.DOC_NO WHERE BRANCH = " & MyParn(XBRANCH.BoundText)
    If IsDate(xdate1.text) Then cString = cString & " AND FILE6_80H.DATE1 >= " & DateSq(xdate1.text)
    If IsDate(xDate2.text) Then cString = cString & " AND FILE6_80H.DATE2 <= " & DateSq(xDate2.text)
    If Not bOpt5 Then cString = cString & " AND FILE6_80.[BRANCH] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    cString = cString & " Order by DATE1 "
    DATA1.RecordSource = cString
    DATA1.Refresh
End With
Fixgrd
End Sub
Sub Fixgrd()
    With Grid1
    .RowHeight(0) = 800
    .WordWrap = True
    .Cols = 8
    .TextMatrix(0, 0) = "ÇáÓäÉ"
    .TextMatrix(0, 1) = "ÇáÔåÑ"
    .TextMatrix(0, 4) = "ÇáãÓÊåÏÝ ÇáÔåÑì"
    .TextMatrix(0, 5) = "ÇáãÈíÚÇÊ ÇáÔåÑíÉ"
    .TextMatrix(0, 6) = "ÇáÝÑÞ"
    .TextMatrix(0, 7) = "äÓÈÉ "
    
    .ColHidden(2) = True
    .ColHidden(3) = True
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    
    For i = 4 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .MergeCells = flexMergeFixedOnly
    
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0", vbRed, vbYellow, True, "  "
    
    For i = 1 To .Rows - 1
        .TextMatrix(i, 5) = Round(Val(.TextMatrix(i, 5)), 0)
        .TextMatrix(i, 6) = Val(.TextMatrix(i, 5)) - Val(.TextMatrix(i, 4))
        nRate = 0
        If Val(.TextMatrix(i, 4)) <> 0 Then nRate = Round((Val(.TextMatrix(i, 5)) / Val(.TextMatrix(i, 4))) * 100, 2)
        .TextMatrix(i, 7) = nRate
    Next i
    
    .ColFormat(4) = "#,###,##0"
    .ColFormat(5) = "#,###,##0"
    .ColFormat(6) = "#,###,##0"
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub


