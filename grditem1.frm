VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form grditemfrm1 
   Caption         =   "„ «»⁄… ›—Ê⁄ Œ·«· › —…"
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
      Caption         =   "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ…"
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
         Picture         =   "grditem1.frx":0000
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
         Picture         =   "grditem1.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   2385
         Picture         =   "grditem1.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "⁄—÷"
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ :"
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
         Caption         =   "„‰  «—ÌŒ :"
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
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
      FixedRows       =   2
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
Attribute VB_Name = "grditemfrm1"
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
cHead1 = "„ «»⁄… «·›—Ê⁄"
If IsDate(xDate1.text) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "„‰ : " & Format(xDate1.text, "dd-mm-yyyy")
If IsDate(xDate2.text) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "Õ Ì : " & Format(xDate2.text, "dd-mm-yyyy")
'Set printGrd3.myform = Me
Load PrintGrd
PrintGrd.doprint grid1, 1.02, 0, cHead1, cHead2, "", False, True, 9
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
    myload
End Sub
Private Sub Form_Load()
openCon con

Set grid1.DataSource = DATA10
DATA10.ConnectionString = strCon

Fixgrd
End Sub
Private Sub myload()
Dim cString As String, cWhere As String
With grid1
If IsDate(xDate1.text) Then
    cWhere = "DATE < " & DateSq(xDate1.text)
    cField1 = myiif(cWhere, "[IN] - [OUT]") & " AS FIELD1"
    cField2 = myiif(cWhere, "Round(([IN]- OUT) * FILE1_10.COST,0)") & " AS FIELD2"
Else
    cField1 = "0 AS FIELD1"
    cField2 = "0 AS FIELD2"
End If
cWhere = ""
If IsDate(xDate1.text) Then cWhere = "DATE >= " & DateSq(xDate1.text)
If IsDate(xDate2.text) Then cWhere = cWhere & turn(cWhere, " AND") & " DATE <= " & DateSq(xDate2.text)

cField3 = myiif(cWhere & turn(cWhere, " AND ") & "(TYPE = '2' OR TYPE = '7')", "[IN] - OUT") & " AS PUR"
cField4 = myiif(cWhere & turn(cWhere, " AND ") & "(TYPE = '2' OR TYPE = '7')", "Round(([IN] - OUT) * FILE1_10.COST,0)") & " AS PUR_COST"

cField5 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'T'", "[IN]") & " AS TRANS_IN"
cField6 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'T'", "ROUND([IN] * FILE1_10.COST,0)") & " AS TRANS_IN_COST"

cField7 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'F'", "[OUT]") & " AS TRANS_OUT"
cField8 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'F'", "Round([OUT] * FILE1_10.COST,0)") & " AS TRANS_OUT_COST"

cField9 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'Z'", "[OUT]") & " AS STOCK"
cField10 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = 'Z'", "Round([OUT] * FILE1_10.COST,0)") & " AS STOCK_COST"

cField11 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = '6'", "[OUT]") & " AS SALES"
cField12 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = '6'", "Round([OUT] * FILE1_11.COST,0)") & " AS SALES_COST"
cField13 = myiif(cWhere & turn(cWhere, " AND ") & "TYPE = '6'", "Round( TOTAL ,0)") & " AS SALES_PRICE"
cField14 = "'' AS SALES_PRF"
cField15 = myiif("", "([IN]- OUT)") & " AS BAL_LAST"
cField16 = myiif("", "Round(([IN]- OUT) * FILE1_10.COST,0)") & " AS BAL_LAST_COST"
cField17 = "'' AS SALES_RATE"
cField18 = myiif(cWhere & turn(cWhere, " AND ") & "[TYPE] <> '6'", "Round(([IN]-OUT)  * FILE1_10.COST,0)") & " AS TOTAL_IN_COST"

cField = cField1 & "," & cField2 & "," & cField3 & "," & cField4 & "," & cField5 & "," & cField6 & "," & cField7 & _
         "," & cField8 & "," & cField9 & "," & cField10 & "," & cField11 & "," & cField12 & "," & cField13 & "," & cField14 & _
         "," & cField15 & "," & cField16 & "," & cField17 & "," & cField18

cString = "SELECT FILE1_11.STORE,FILE0_40.DESCA," & cField & _
          " FROM (FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM) INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE"
If IsDate(xDate2.text) Then cString = cString & turn(cString) & " [DATE] <= " & DateSq(xDate2.text)
cString = cString & " GROUP BY FILE1_11.STORE,FILE0_40.DESCA"
cString = cString & " ORDER BY FILE1_11.STORE"
Set DATA10.Recordset = myRecordSet(cString, con)
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
    .FrozenCols = 2
    .WordWrap = True
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·›—⁄"
    .TextMatrix(1, 0) = "ﬂÊœ"
    .TextMatrix(1, 1) = "«·›—⁄"

    .TextMatrix(0, 2) = "—’Ìœ «Ê·"
    .TextMatrix(0, 3) = "—’Ìœ «Ê·"
    .TextMatrix(1, 2) = "⁄œœ"
    .TextMatrix(1, 3) = "ﬁÌ„…"

    .TextMatrix(0, 4) = "„‘ —Ì« "
    .TextMatrix(0, 5) = "„‘ —Ì« "
    .TextMatrix(1, 4) = "⁄œœ"
    .TextMatrix(1, 5) = "ﬁÌ„…"

    .TextMatrix(0, 4 + 2) = "Ê«—œ  ÕÊÌ·« "
    .TextMatrix(0, 5 + 2) = "Ê«—œ  ÕÊÌ·« "
    .TextMatrix(1, 4 + 2) = "⁄œœ"
    .TextMatrix(1, 5 + 2) = "ﬁÌ„…"

    .TextMatrix(0, 6 + 2) = "’«œ—  ÕÊÌ·« "
    .TextMatrix(0, 7 + 2) = "’«œ—  ÕÊÌ·« "
    .TextMatrix(1, 6 + 2) = "⁄œœ"
    .TextMatrix(1, 7 + 2) = "ﬁÌ„…"

    .TextMatrix(0, 8 + 2) = " ”ÊÌ«  Ã—œ"
    .TextMatrix(0, 9 + 2) = " ”ÊÌ«  Ã—œ"
    .TextMatrix(1, 8 + 2) = "⁄œœ"
    .TextMatrix(1, 9 + 2) = "ﬁÌ„…"

    .TextMatrix(0, 10 + 2) = "„»Ì⁄« "
    .TextMatrix(0, 11 + 2) = "„»Ì⁄« "
    .TextMatrix(0, 12 + 2) = "„»Ì⁄« "
    .TextMatrix(0, 13 + 2) = "„»Ì⁄« "
    .TextMatrix(1, 10 + 2) = "⁄œœ"
    .TextMatrix(1, 11 + 2) = "ﬁÌ„…  ﬂ·›…"
    .TextMatrix(1, 12 + 2) = "ﬁÌ„… »Ì⁄"
    .TextMatrix(1, 13 + 2) = "—»Õ"

    .TextMatrix(0, 14 + 2) = "—’Ìœ Õ«·Ì"
    .TextMatrix(0, 15 + 2) = "—’Ìœ Õ«·Ì"
    .TextMatrix(1, 14 + 2) = "⁄œœ"
    .TextMatrix(1, 15 + 2) = "ﬁÌ„…"

    .TextMatrix(0, 16 + 2) = "„»Ì⁄« "
    .TextMatrix(1, 16 + 2) = "‰”»…"
    
    '.ColFormat(15) = "#.##%"
    .ColFormat(18) = "#%"
    .ColWidth(1) = 1200
    .ColWidth(2) = 700
    .ColWidth(3) = 900
    .ColWidth(4) = 700
    .ColWidth(5) = 1000
    .ColWidth(6) = 700
    .ColWidth(7) = 1000
    .ColWidth(8) = 700
    .ColWidth(9) = 1000
    .ColWidth(10) = 700
    .ColWidth(11) = 1100
    .ColWidth(12) = 700
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000
    .ColWidth(15) = 1000
    .ColWidth(16) = 800
    .ColWidth(17) = 1000
    .ColWidth(18) = 0
    
    .ColHidden(0) = True
    .ColHidden(.Cols - 1) = True
    For i = 2 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    For i = 2 To .Rows - 1
        For nCol = 2 To .Cols - 1
            .TextMatrix(i, nCol) = Round(Val(.ValueMatrix(i, nCol)), 2)
        Next nCol
        .TextMatrix(i, 15) = Val(.TextMatrix(i, 14)) - Val(.TextMatrix(i, 13))
        If Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 19)) <> 0 Then
            .TextMatrix(i, 18) = (Val(.TextMatrix(i, 13)) / (Val(.TextMatrix(i, 2)) + Val(.TextMatrix(i, 19))))
        End If
    Next i
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 3, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 8, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 9, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 10, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 11, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 12, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 13, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 14, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 15, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 16, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 17, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 19, "#0", vbRed, vbYellow, True, "  "
    If .Rows > 2 Then
        .TextMatrix(2, 1) = "«·≈Ã„«·Ï"
        If Val(.TextMatrix(2, 2)) + Val(.TextMatrix(2, 19)) <> 0 Then
            .TextMatrix(2, 18) = (Val(.TextMatrix(2, 13)) / (Val(.TextMatrix(2, 2)) + Val(.TextMatrix(2, 19))))
        End If
    End If
    xTotal.Caption = IIf(grid1.Rows > 2, grid1.Rows - 2, "")
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set VsTItem = Nothing
End Sub

Private Sub grid1_DblClick()
    Dim sHead2 As String
    With grid1
    If grid1.Col = 2 + 2 Or grid1.Col = 10 + 2 And grid1.Row > 1 Then
        StoreMove.sType = 2
        StoreMove.sitem = grid1.TextMatrix(grid1.Row, 0 + 2)
        StoreMove.sDate1 = xDate1.text
        StoreMove.Show
    End If
    
    If grid1.Row < 1 Then Exit Sub
    
    If .Col = 3 + 2 Or .Col = 4 + 2 Then
        cString = "SELECT FILE0_10H.DOC_NO AS [—ﬁ„ «·„” ‰œ],CONVERT(VARCHAR(10),FILE0_10H.[DATE],111) AS [«· «—ÌŒ],SUM(FILE0_10.DIFFER) AS [≈Ã„«·Ì «·ﬂ„Ì…]" & _
                  ",FILE0_10.COST AS [«· ﬂ·›…], SUM(FILE0_10.DIFFER * FILE0_10.COST) as [«·«Ã„«·Ì] " & _
                  " FROM FILE0_10 INNER JOIN FILE0_10H ON FILE0_10.DOC_NO = FILE0_10H.DOC_NO"
        cString = cString & turn(cString) & " FILE0_10H.CLOSED = 1"
        If Trim(.TextMatrix(.Row, 0 + 2)) <> "" Then cString = cString & turn(cString) & " FILE0_10.ITEM = " & MyParn(.TextMatrix(.Row, 0 + 2))
        Showfrm1.sHead1 = "„ «»⁄…  ”ÊÌ… Ã—œ" & turn(.TextMatrix(.Row, 0 + 2), " : " & .TextMatrix(grid1.Row, 1 + 2))
        If IsDate(xDate1.text) Then
            cString = cString & turn(cString) & " FILE0_10H.DATE >= " & DateSq(xDate1.text)
            sHead2 = "„‰ : " & xDate1.text
        End If
        If IsDate(xDate2.text) Then
            cString = cString & turn(cString) & " FILE0_10H.DATE <= " & DateSq(xDate2.text)
            sHead2 = sHead2 & turn(sHead2, " ") & "Õ Ï : " & xDate2.text
        End If
        Showfrm1.sHead2 = sHead2
        cString = cString & " Group by FILE0_10H.DOC_NO,FILE0_10H.[DATE],FILE0_10.COST"
        cString = cString & " ORDER BY FILE0_10H.DATE DESC,FILE0_10H.DOC_NO DESC"
        Showfrm1.sSql = cString
        
        Showfrm1.sCol = "1000-1300-1100-1100-1100"
        Showfrm1.sTotal = "2-3-4"
        Showfrm1.Show 1
    End If
    
    If .Col = 5 + 2 Or .Col = 6 + 2 Then
        cString = "SELECT FILE1_80H.DOC_NO AS [—ﬁ„ «·„” ‰œ],CONVERT(VARCHAR(10),FILE1_80H.[DATE],111) AS [«· «—ÌŒ],SUM(FILE1_80.QUANT) AS [≈Ã„«·Ì «·ﬂ„Ì…]" & _
                  ",FILE1_80.COST AS [«· ﬂ·›…], SUM(FILE1_80.QUANT * FILE1_80.COST) as [«·«Ã„«·Ì] " & _
                  " FROM FILE1_80 INNER JOIN FILE1_80H ON FILE1_80.DOC_NO = FILE1_80H.DOC_NO"
        
        If Trim(.TextMatrix(.Row, 0 + 2)) <> "" Then cString = cString & turn(cString) & " FILE1_80.ITEM = " & MyParn(.TextMatrix(.Row, 0 + 2))
        Showfrm1.sHead1 = "„ «»⁄… «‰ «Ã" & turn(.TextMatrix(.Row, 0 + 2), " : " & .TextMatrix(grid1.Row, 1 + 2))
        
        If IsDate(xDate1.text) Then
            cString = cString & turn(cString) & " FILE1_80H.DATE >= " & DateSq(xDate1.text)
            sHead2 = "„‰ : " & xDate1.text
        End If
        If IsDate(xDate2.text) Then
            cString = cString & turn(cString) & " FILE1_80H.DATE <= " & DateSq(xDate2.text)
            sHead2 = sHead2 & turn(sHead2, " ") & "Õ Ï : " & xDate2.text
        End If
        cString = cString & " Group by FILE1_80H.DOC_NO,FILE1_80H.[DATE],FILE1_80.COST"
        cString = cString & " ORDER BY FILE1_80H.DATE DESC,FILE1_80H.DOC_NO DESC"
        Showfrm1.sHead2 = sHead2
        
        Showfrm1.sSql = cString
        Showfrm1.sCol = "1000-1300-1100-1100-1100"
        Showfrm1.sTotal = "2-3-4"
        Showfrm1.Show 1
    ElseIf .Col = 7 + 2 Or .Col = 8 + 2 Then
        cString = "SELECT FILE1_82H.DOC_NO AS [—ﬁ„ «·„” ‰œ],CONVERT(VARCHAR(10),FILE1_82H.[DATE],111) AS [«· «—ÌŒ],SUM(FILE1_82.QUANT) AS [≈Ã„«·Ì «·ﬂ„Ì…]" & _
                  ",FILE1_82.COST AS [«· ﬂ·›…], SUM(FILE1_82.QUANT * FILE1_82.COST) as [«·«Ã„«·Ì] " & _
                  " FROM FILE1_82 INNER JOIN FILE1_82H ON FILE1_82.DOC_NO = FILE1_82H.DOC_NO"
        
        If Trim(.TextMatrix(.Row, 0 + 2)) <> "" Then cString = cString & turn(cString) & " FILE1_82.ITEM = " & MyParn(.TextMatrix(.Row, 0 + 2))
        Showfrm1.sHead1 = "„ «»⁄… Â«·ﬂ" & turn(.TextMatrix(.Row, 0 + 2), " : " & .TextMatrix(grid1.Row, 1 + 2))
        
        If IsDate(xDate1.text) Then
            cString = cString & turn(cString) & " FILE1_82H.DATE >= " & DateSq(xDate1.text)
            sHead2 = "„‰ : " & xDate1.text
        End If
        If IsDate(xDate2.text) Then
            cString = cString & turn(cString) & " FILE1_82H.DATE <= " & DateSq(xDate2.text)
            sHead2 = sHead2 & turn(sHead2, " ") & "Õ Ï : " & xDate2.text
        End If
        Showfrm1.sHead2 = sHead2
        cString = cString & " Group by FILE1_82H.DOC_NO,FILE1_82H.[DATE],FILE1_82.COST"
        cString = cString & " ORDER BY FILE1_82H.DATE DESC,FILE1_82H.DOC_NO DESC"
        Showfrm1.sSql = cString
        Showfrm1.sCol = "1000-1300-1100-1100-1100"
        Showfrm1.sTotal = "2-3-4"
        Showfrm1.Show 1
    ElseIf .Col = 9 + 2 Then
        cString = "SELECT FILE1_90H.DOC_NO AS [—ﬁ„ «·„” ‰œ],CONVERT(VARCHAR(10),FILE1_90H.[DATE],111) AS [«· «—ÌŒ],FILE0_50.DESCA AS [«·ﬂ«‘Ì—],SUM(FILE1_90.QUANT) AS [≈Ã„«·Ì «·ﬂ„Ì…]" & _
                  " FROM ((FILE1_90 INNER JOIN FILE1_90H ON FILE1_90.DOC_NO = FILE1_90H.DOC_NO) INNER JOIN CLOSEDAY ON FILE1_90H.DOC_NO = CLOSEDAY.CODE) INNER JOIN FILE0_50 ON CLOSEDAY.BOX1 = FILE0_50.CODE"
        
        If Trim(.TextMatrix(.Row, 0 + 2)) <> "" Then cString = cString & turn(cString) & " FILE1_90.ITEM = " & MyParn(.TextMatrix(.Row, 0 + 2))
        Showfrm1.sHead1 = "„ «»⁄… „»Ì⁄« " & turn(.TextMatrix(.Row, 0 + 2), " : " & .TextMatrix(grid1.Row, 1 + 2))
        
        If IsDate(xDate1.text) Then
            cString = cString & turn(cString) & " FILE1_90H.DATE >= " & DateSq(xDate1.text)
            sHead2 = "„‰ : " & xDate1.text
        End If
        If IsDate(xDate2.text) Then
            cString = cString & turn(cString) & " FILE1_90H.DATE <= " & DateSq(xDate2.text)
            sHead2 = sHead2 & turn(sHead2, " ") & "Õ Ï : " & xDate2.text
        End If
        Showfrm1.sHead2 = sHead2
        cString = cString & " Group by FILE1_90H.DOC_NO,FILE1_90H.[DATE],FILE0_50.DESCA"
        cString = cString & " ORDER BY FILE1_90H.DATE DESC,FILE1_90H.DOC_NO DESC"
        Showfrm1.sSql = cString
        Showfrm1.sCol = "1000-1300-2500-1100"
        Showfrm1.sTotal = "3"
        Showfrm1.Show 1
    End If

    
    End With
    
'        StoreMove.Show
'    Else
'        cRepItem = grid1.TextMatrix(grid1.Row, 0)
'        DRepDate1 = xdate1.Text
'        DRepDate2 = xDate2.Text
'        ShowPrdItem.Show
'    End If
End Sub
Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xDesca.text, 1
End If
End Sub
Private Sub xitem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    ItemsLookupAll Me, oSearchItem
End If
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, XITEM.text, 0
End If
End Sub
Sub myProc()
XITEM.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
xDesca.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
Unload oSearchItem
End Sub
Private Sub xDate1_Validate(Cancel As Boolean)
With xDate1
If Not IsDate(.text) Then
    .text = ""
Else
    .text = Format(.text, "dd-mm-yyyy")
End If
End With
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
With xDate2
If Not IsDate(.text) Then
    .text = ""
Else
    .text = Format(.text, "dd-mm-yyyy")
End If
End With
End Sub
Private Sub Handlecontrols()
cmdPrint.Enabled = grid1.Rows > 1
End Sub


