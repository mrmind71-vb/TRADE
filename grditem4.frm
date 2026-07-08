VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form grditemfrm4 
   Caption         =   "ãÊÇÈÚÉ ÝÑæÚ ÎáÇá ÝÊÑÉ"
   ClientHeight    =   9405
   ClientLeft      =   90
   ClientTop       =   465
   ClientWidth     =   14955
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
   ScaleWidth      =   14955
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
         Picture         =   "grditem4.frx":0000
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
         Picture         =   "grditem4.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   2385
         Picture         =   "grditem4.frx":4896
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
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   225
         TabIndex        =   13
         Top             =   450
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   495
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   300
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
      Rows            =   1
      Cols            =   7
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
      Width           =   14955
      _ExtentX        =   26379
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "grditemfrm4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchItem As New Search31
Dim LastSalTable As New ADODB.Recordset
Dim LastImpTable As New ADODB.Recordset
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "ãÊÇÈÚÉ ÇáÝÑæÚ"
    If IsDate(xDate1.Text) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "ãä : " & Format(xDate1.Text, "dd-mm-yyyy")
    If IsDate(xDate2.Text) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "ÍÊí : " & Format(xDate2.Text, "dd-mm-yyyy")
'        Set printGrd3.grid1 = grid1
    Load PrintGrd
    PrintGrd.doprint grid1, 1, 0, cHead1, cHead2, "", False, False, 11
    PrintGrd.Show 1
    
'    Set printGrd3.myform = Me
'    printGrd3.doprint 0.8, 0, cHead1, cHead2, "", False, False, 9
'    printGrd3.Show 1
    
    'PrintGrd.doprint grid1, 1, -2, cHead1, cHead2, , False, True, 8, , Array(1), Array(1), 11
    'PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
Unload Me
Set TSalItem = Nothing
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
openCon con

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select Code,DescA From File0_40 ORDER BY DESCA"
Set XSTORE.RowSource = DATA2
XSTORE.ListField = "Desca"
XSTORE.BoundColumn = "Code"

Set grid1.DataSource = data10
data10.ConnectionString = strCon

Fixgrd
End Sub
Private Sub myload()
Dim sMyBox As String
Dim cString As String, cwhere As String
With grid1
If XSTORE.BoundText = "" Then
    MsgBox " ÊÍÏíÏ ÇáÝÑÚ "
    Exit Sub
End If
sMyBox = GetDesca("SELECT BOX FROM FILE0_40 WHERE CODE = " & MyParn(XSTORE.BoundText))
cStr1 = " SELECT file4_10.CODE , file4_10.DESCA " & _
        " , (SELECT SUM(([IN]-[OUT]) * FILE1_10.COST ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE (TYPE = '2' OR TYPE = '7' OR TYPE = 'F' OR TYPE = 'T' ) AND FILE1_10.CODE = FILE4_10.CODE AND FILE1_11.STORE = " & MyParn(XSTORE.BoundText) & " ) AS T_PURCH  " & _
        " , (SELECT SUM(T_SALESDOC.T_TOTAL - T_SALESDOC.DISCOUNT ) FROM T_SALESDOC WHERE T_SALESDOC.SUPP = FILE4_10.CODE AND T_SALESDOC.STORE = " & MyParn(XSTORE.BoundText) & " )  AS T_SALES " & _
        " , (SELECT SUM(FILE8_20.VALUE) FROM FILE8_20 INNER JOIN FILE8_20H ON FILE8_20.DOC_NO = FILE8_20H.DOC_NO WHERE FILE8_20.CODE = FILE4_10.CODE AND FILE8_20.BOX = " & MyParn(sMyBox) & " )  AS T_PAY " & _
        " , '  ' AS N4 " & _
        " ,  (SELECT SUM(([IN]-[OUT]) * FILE1_10.COST ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE  FILE1_10.CODE  = FILE4_10.CODE AND FILE1_11.STORE = " & MyParn(XSTORE.BoundText) & " )   AS COST_ITEM " & _
        " FROM FILE4_10 ORDER BY FILE4_10.CODE "

'
'select CODE , DESCA
', (SELECT SUM(([IN]-[OUT]) * FILE1_10.COST ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE (TYPE = '2' OR TYPE = '7' OR TYPE = 'F' OR TYPE = 'T' )  AND FILE1_11.STORE = FILE0_40.CODE )
', (SELECT SUM(T_SALESDOC.T_TOTAL - T_SALESDOC.DISCOUNT ) FROM T_SALESDOC WHERE T_SALESDOC.STORE  = FILE0_40.CODE )
', (SELECT SUM(FILE8_50.VALUE) FROM FILE8_50 INNER JOIN FILE8_50H ON FILE8_50.DOC_NO = FILE8_50H.DOC_NO WHERE FILE8_50.BOX = FILE0_40.box)
'From FILE0_40


''''cwhere = "FILE6_20H.STORE = FILE0_40.CODE"
''''If IsDate(xDate1.Text) Then cwhere = cwhere & turn(cwhere, " AND ") & "FILE6_20H.DATE >= " & DateSq(xDate1.Text)
''''If IsDate(xDate2.Text) Then cwhere = cwhere & turn(cwhere, " AND ") & "FILE6_20H.DATE <= " & DateSq(xDate2.Text)
''''
''''
''''cField1 = " (SELECT SUM(([IN]-[OUT]) * FILE1_10.COST ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE (TYPE = '2' OR TYPE = '7' OR TYPE = 'F' OR TYPE = 'T' )  AND FILE1_11.STORE = FILE0_40.CODE ) "
''''cField2 = " (SELECT SUM(T_SALESDOC.T_TOTAL - T_SALESDOC.DISCOUNT ) FROM T_SALESDOC WHERE T_SALESDOC.STORE  = FILE0_40.CODE ) "
''''cField3 = " (SELECT SUM(FILE8_50.VALUE) FROM FILE8_50 INNER JOIN FILE8_50H ON FILE8_50.DOC_NO = FILE8_50H.DOC_NO WHERE FILE8_50.BOX = FILE0_40.box)"
''''cField4 = " (SELECT SUM(([IN]-[OUT]) * FILE1_10.COST ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE  FILE1_11.STORE = FILE0_40.CODE ) "
''''
''''cString = "SELECT CODE , DESCA," & cField1 & "," & cField2 & "," & cField3 & ", '  ' AS N4 , " & cField4 & _
''''          " FROM FILE0_40"
''''cString = cString & " ORDER BY FILE0_40.CODE"

Set data10.Recordset = myRecordSet(cStr1, con)

'DATA10.Refresh

'Dim loctable As New adodb.Recordset
'loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
''
'.Rows = grid1.FixedRows
'Do Until loctable.EOF
'    .AddItem ""
'    For I = 0 To grid1.Cols - 1
'        .TextMatrix(.Rows - 1, I) = loctable.Fields(I)
'    Next
''    .AddItem ""
''    .TextMatrix(.Rows - 1, 0) = loctable!Item
''    .TextMatrix(.Rows - 1, 1) = loctable!desca & ""
''    .TextMatrix(.Rows - 1, 2) = Myvalue(Val(loctable!field1 & ""))
''    .TextMatrix(.Rows - 1, 3) = Myvalue(Val(loctable!field2 & ""), "FIXED")
''    .TextMatrix(.Rows - 1, 4) = Myvalue(Val(loctable!field3 & ""))
''    .TextMatrix(.Rows - 1, 5) = Myvalue(Val(loctable!field4 & ""), "FIXED")
''    .TextMatrix(.Rows - 1, 5) = Myvalue(Val(loctable!field5 & ""))
''    .TextMatrix(.Rows - 1, 6) = Myvalue(Val(loctable!Field6 & ""), "FIXED")
''
'    loctable.MoveNext
'Loop


End With
Fixgrd
Handlecontrols
End Sub
Sub Fixgrd()
Dim nTotal_Sales As Double, nTotal_in As Double
    With grid1
    .RowHeight(0) = 800
    .WordWrap = True
    
'                                           0       1           2                   3       4           5           6
    .FormatString = retFormatString(Array("ßæÏ", "ÇáÝÑÚ", "ÅÌãÇáì ÇáæÇÑÏ", "ÅÌãÇáì ãÈíÚÇÊ", "ÓÏÇÏ", "ÏÝÚÉ ãÓÊÍÞÉ", "ÑÕíÏ ÈÖÇÚÉ", "äÓÈÉ ÇáÈíÚ"))
    
    .ColWidth(1) = 1000
    .ColWidth(2) = 2000
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColHidden(0) = True
    For i = 2 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    For i = 1 To .Rows - 1
        .TextMatrix(i, 5) = Val(.TextMatrix(i, 3)) - Val(.TextMatrix(i, 4))
        nRate = Val(.TextMatrix(i, 2)) - Val(.TextMatrix(i, 6))
        If Val(.TextMatrix(i, 2)) <> 0 Then
            nRate = Round((nRate / Val(.TextMatrix(i, 2))) * 100, 2)
        Else
            nRate = 0
        End If
        .TextMatrix(i, 7) = nRate
        For nCol = 2 To .Cols - 1
            .TextMatrix(i, nCol) = Round(Val(.TextMatrix(i, nCol)), 2)
        Next nCol
    Next i
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 3, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0.00", vbRed, vbYellow, True, "  "
    If .Rows > 1 Then
        .TextMatrix(1, 1) = "ÇáÅÌãÇáì"
        
    End If
    
    xTotal.Caption = IIf(grid1.Rows > 1, grid1.Rows - 2, "")
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set VsTItem = Nothing
End Sub
Private Sub Handlecontrols()
    cmdPrint.Enabled = grid1.Rows > 1
End Sub


