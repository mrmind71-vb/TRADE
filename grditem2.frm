VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form grditemfrm2 
   Caption         =   "≈Ã„«·Ï «—»«Õ «·›—Ê⁄ ·› —…"
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
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "query "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7290
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   45
      Width           =   915
   End
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
      Width           =   6045
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "grditem2.frx":0000
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
         Picture         =   "grditem2.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   2385
         Picture         =   "grditem2.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   3825
         TabIndex        =   13
         Top             =   135
         Width           =   2055
         _ExtentX        =   3625
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
         Picture         =   "grditem2.frx":78EF
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   8370
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   7755
      Begin VB.TextBox textquery 
         Alignment       =   1  'Right Justify
         Height          =   780
         Left            =   45
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   135
         Width           =   4875
      End
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
         Left            =   4995
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
         Left            =   4995
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
         Left            =   6705
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
         Left            =   6705
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
      Width           =   15990
      _cx             =   28205
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
Attribute VB_Name = "grditemfrm2"
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
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "„ «»⁄… «·›—Ê⁄"
    If IsDate(xDate1.text) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "„‰ : " & Format(xDate1.text, "dd-mm-yyyy")
    If IsDate(xDate2.text) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "Õ Ì : " & Format(xDate2.text, "dd-mm-yyyy")
'        Set printGrd3.grid1 = grid1
    Load PrintGrd
    PrintGrd.doprint grid1, 1.1, 0, cHead1, cHead2, "", False, True, 11
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
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
openCon con

Set grid1.DataSource = data10
data10.ConnectionString = strCon

Fixgrd
End Sub
Private Sub myload()
Dim cString As String, cWhere As String
With grid1

cWhere = "FILE6_20H.STORE = FILE0_40.CODE"
If IsDate(xDate1.text) Then cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.DATE >= " & DateSq(xDate1.text)
If IsDate(xDate2.text) Then cWhere = cWhere & turn(cWhere, " AND ") & "FILE6_20H.DATE <= " & DateSq(xDate2.text)


cField1 = "SELECT SUM(FILE6_20.TOTAL) FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO"
cField1 = "(" & cField1 & turn(cWhere) & cWhere & ")"
cField1 = "COALESCE(" & cField1 & ",0)"

cField2 = "SELECT SUM(FILE6_20H.DISCOUNT) FROM FILE6_20H WHERE FILE6_20H.DOC_NO IN (SELECT DOC_NO FROM FILE6_20)"
cField2 = "(" & cField2 & " AND " & cWhere & ")"
cField2 = "COALESCE(" & cField2 & ",0)"

cField3 = "SELECT SUM(FILE6_20H.tax ) FROM FILE6_20H WHERE FILE6_20H.DOC_NO IN (SELECT DOC_NO FROM FILE6_20)"
cField3 = "(" & cField3 & " AND " & cWhere & ")"
cField3 = "COALESCE(" & cField3 & ",0)"

cField1 = cField1 & " - " & cField2 & " + " & cField3 & " AS SALES"

cField2 = "SELECT SUM(FILE6_20.COST * FILE6_20.QUANT) FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO"
cField2 = "(" & cField2 & turn(cWhere) & cWhere & ") AS COST"
cField3 = "' ' AS PROFIT"

cWhere = " AND FILE0_40_1.CODE = FILE0_40.CODE "
If IsDate(xDate1.text) Then cWhere = cWhere & " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xDate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xDate2.text)

cField4 = "SELECT SUM(acc_charge.VALUE) FROM acc_charge INNER JOIN FILE0_40 AS FILE0_40_1 ON FILE0_40_1.BRANCH = acc_charge.BRANCH INNER JOIN File8_51 ON File8_51.CODE = ACC_CHARGE.CHARGE WHERE File8_51.ISasset = 0 and acc_charge.branch is not null "
cField4 = "(" & cField4 & cWhere & ") AS TOTAL_CHARGES"

cWhere = " WHERE FILE0_40.CODE = FILE8_00.STORE "
If IsDate(xDate1.text) Then cWhere = cWhere & " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xDate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xDate2.text)
cField5 = "SELECT SUM(FILE8_00.VALUE) FROM FILE8_00 INNER JOIN FILE8_00H ON FILE8_00.DOC_NO = FILE8_00H.DOC_NO  "
cField5 = "(" & cField5 & cWhere & ") AS TOTAL_DISC "

cString = "SELECT FILE0_40.CODE,FILE0_40.DESCA," & cField1 & "," & cField2 & "," & cField3 & "," & cField4 & "," & cField5 & _
          " FROM FILE0_40 WHERE FILE0_40.BRANCH IS NOT NULL "

cString = cString & " UNION ALL "

cString = cString & " SELECT BRANCH.CODE ,  BRANCH.DESCA , 0 , 0 , 0 , SUM(VALUE) , 0 "
cString = cString & " FROM acc_charge INNER JOIN BRANCH ON BRANCH.CODE = ACC_CHARGE.BRANCH WHERE BRANCH IN (SELECT CODE FROM BRANCH WHERE STORE IS NULL)"
If IsDate(xDate1.text) Then cString = cString & " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xDate2.text) Then cString = cString & " AND DATE <= " & DateSq(xDate2.text)
cString = cString & " GROUP BY BRANCH.CODE ,  BRANCH.DESCA"

If Check1.Value = 0 Then
    Set data10.Recordset = myRecordSet(cString, con)
Else
    textquery.text = cString
End If
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
Dim nTotal_Sales As Double, nTotal_in As Double, nTotRow As Double
    With grid1
    .Cols = 12
    .RowHeight(0) = 1000
    .WordWrap = True
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·›—⁄"
    
    .TextMatrix(0, 2) = "„»Ì⁄« "
    .TextMatrix(0, 3) = " ﬂ·›…"
    .TextMatrix(0, 4) = "«—»«Õ „»Ì⁄« "
    
    .TextMatrix(0, 5) = "„’«—Ì›"
    .TextMatrix(0, 6) = " ”ÊÌ«  „Ê—œÌ‰"
    .TextMatrix(0, 7) = "’«›Ì «—»«Õ"
    
    .TextMatrix(0, 8) = "‰”»… —»Õ ··„»Ì⁄« "
    .TextMatrix(0, 9) = "«·‰”»… „»Ì⁄«  ··≈Ã„«·Ï"
    .TextMatrix(0, 10) = "‰”»… „’«—Ì› ··„»Ì⁄« "
    .TextMatrix(0, 11) = "‰”»… —»Õ ·’«›Ï «·„»Ì⁄« "
    
    .ColWidth(1) = 2000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColWidth(6) = 1500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    
    .ColHidden(0) = True
    For i = 2 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    
    .Select 0, 0
    .Sort = flexSortGenericAscending
 
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    .ColFormat(2) = "#,###,##0"
    .ColFormat(3) = "#,###,##0"
    .ColFormat(4) = "#,###,##0"
    .ColFormat(5) = "#,###,##0"
    .ColFormat(6) = "#,###,##0"
    .ColFormat(7) = "#,###,##0"
    For i = 1 To .Rows - 1
        .TextMatrix(i, 4) = Round(Val(.TextMatrix(i, 2)) - Val(.TextMatrix(i, 3)), 2)
        .TextMatrix(i, 7) = Round(Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 5)) + Val(.TextMatrix(i, 6)), 2)
        If Val(.TextMatrix(i, 2)) <> 0 Then .TextMatrix(i, 8) = Round(Val(.TextMatrix(i, 4)) / Val(.TextMatrix(i, 2)) * 100, 2)
        If Val(.TextMatrix(i, 2)) <> 0 Then .TextMatrix(i, 10) = Round(Val(.TextMatrix(i, 5)) / Val(.TextMatrix(i, 2)) * 100, 2)
        If Val(.TextMatrix(i, 2)) <> 0 Then .TextMatrix(i, 11) = Round(Val(.TextMatrix(i, 7)) / Val(.TextMatrix(i, 2)) * 100, 2)
    Next i
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#,###,##0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 3, "#,###,##0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 4, "#,###,##0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#,###,##0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#,###,##0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 7, "#,###,##0", vbRed, vbYellow, True, "  "
    If .Rows > 1 Then
        .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
    If Val(.TextMatrix(1, 2)) <> 0 Then .TextMatrix(1, 8) = Round(Val(.TextMatrix(1, 4)) / Val(.TextMatrix(1, 2)) * 100, 2)
    If Val(.TextMatrix(1, 2)) <> 0 Then .TextMatrix(1, 10) = Round(Val(.TextMatrix(1, 5)) / Val(.TextMatrix(1, 2)) * 100, 2)
    If Val(.TextMatrix(1, 2)) <> 0 Then .TextMatrix(1, 11) = Round(Val(.TextMatrix(1, 7)) / Val(.TextMatrix(1, 2)) * 100, 2)
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
    For i = 2 To .Rows - 1
        If Val(.TextMatrix(1, 4)) <> 0 Then .TextMatrix(i, 9) = Round(Val(.TextMatrix(i, 4)) / Val(.TextMatrix(1, 4)) * 100, 2)
    Next i
    .Subtotal flexSTSum, -1, 9, "#0.00", vbRed, vbYellow, True, "  "
    
    xTotal.Caption = IIf(grid1.Rows > 1, grid1.Rows - 2, "")
    
    For nRow = 1 To .Rows - 1
        nTotRow = 0
        For nCol = 2 To .Cols - 1
            nTotRow = nTotRow + Abs(Val(.TextMatrix(nRow, nCol)))
        Next nCol
        If nTotRow = 0 Then .RowHidden(nRow) = True
    Next nRow
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set VsTItem = Nothing
End Sub
Private Sub grid1_DblClick()
    Dim cBr As String
    If grid1.Col = 5 Then
    If grid1.TextMatrix(grid1.Row, 0) <> "" Then
        cBr = GetDesca("SELECT BRANCH FROM  FILE0_40 WHERE CODE = " & MyParn(grid1.TextMatrix(grid1.Row, 0)), con)
    End If
    publicFlag = 1
    rpCharge1.XBRANCH.BoundText = cBr
    rpCharge1.xDate1.text = xDate1.text
    rpCharge1.xDate2.text = xDate2.text
    rpCharge1.Show 1
    End If
End Sub
Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, XDESCA.text, 1
End If
End Sub
Private Sub xitem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    ItemsLookupAll Me, oSearchItem
End If
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xitem.text, 0
End If
End Sub
Sub myProc()
xitem.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
XDESCA.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
Unload oSearchItem
End Sub

Private Sub textquery_GotFocus()
textquery.SelStart = 0
textquery.SelLength = Len(textquery.text)
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


