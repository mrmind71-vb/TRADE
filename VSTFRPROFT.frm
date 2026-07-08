VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VSTFRPROFT 
   Caption         =   "≈Ã„«·Ï „»Ì⁄«  & —»Õ „»Ì⁄«  ·· ÊﬂÌ·« "
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
         Picture         =   "VSTFRPROFT.frx":0000
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
         Picture         =   "VSTFRPROFT.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   510
         Left            =   2385
         Picture         =   "VSTFRPROFT.frx":4896
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
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   90
         TabIndex        =   13
         Top             =   450
         Width           =   2595
         _ExtentX        =   4577
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
         Picture         =   "VSTFRPROFT.frx":78EF
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
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
         Size            =   9.75
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
Attribute VB_Name = "VSTFRPROFT"
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
    cHead1 = "„ «»⁄… «·›—Ê⁄ «· ÊﬂÌ·«  "
    If IsDate(xdate1.Text) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "„‰ : " & Format(xdate1.Text, "dd-mm-yyyy")
    If IsDate(xDate2.Text) Then cHead2 = cHead2 & turn(cHead2, Space(3)) & "Õ Ì : " & Format(xDate2.Text, "dd-mm-yyyy")
'        Set printGrd3.grid1 = grid1
    Load PrintGrd
    PrintGrd.doprint grid1, 1.02, 0, cHead1, cHead2, "", False, True, 10
    PrintGrd.Show 1
    
'    Set printGrd3.myform = Me
'    printGrd3.doprint 0.8, 0, cHead1, cHead2, "", False, False, 9
'    printGrd3.Show 1
    
    'PrintGrd.doprint grid1, 1, -2, cHead1, cHead2, , False, True, 8, , Array(1), Array(1), 11
    'PrintGrd.Show 1
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

Set grid1.DataSource = data10
data10.ConnectionString = strCon

Fixgrd
End Sub
Private Sub myload()
Dim cString As String, cwhere As String
With grid1

cwhere = "FR6_20H.STORE = BRANCH_FR.CODE"
If IsDate(xdate1.Text) Then cwhere = cwhere & turn(cwhere, " AND ") & "FR6_20H.DATE >= " & DateSq(xdate1.Text)
If IsDate(xDate2.Text) Then cwhere = cwhere & turn(cwhere, " AND ") & "FR6_20H.DATE <= " & DateSq(xDate2.Text)


cField1 = "SELECT SUM(FR6_20.TOTAL) FROM FR6_20 INNER JOIN FR6_20H ON FR6_20.DOC_NO = FR6_20H.DOC_NO"
cField1 = "(" & cField1 & turn(cwhere) & cwhere & ")"
cField1 = "COALESCE(" & cField1 & ",0)"

cField2 = "SELECT SUM(FR6_20H.DISCOUNT) FROM FR6_20H"
cField2 = "(" & cField2 & turn(cwhere) & cwhere & ")"
cField2 = "COALESCE(" & cField2 & ",0)"

cField1 = cField1 & " - " & cField2 & " AS SALES"

cField2 = "SELECT SUM(FR6_20.COST * FR6_20.QUANT) FROM FR6_20 INNER JOIN FR6_20H ON FR6_20.DOC_NO = FR6_20H.DOC_NO"
cField2 = "(" & cField2 & turn(cwhere) & cwhere & ") AS COST"
cField3 = "' ' AS PROFIT"


cString = "SELECT BRANCH_FR.CODE,BRANCH_FR.DESCA," & cField1 & "," & cField2 & "," & cField3 & _
          " FROM BRANCH_FR"
cString = cString & " ORDER BY BRANCH_FR.CODE"
Set data10.Recordset = myRecordSet(cString, con)
End With
Fixgrd
Handlecontrols
End Sub
Sub Fixgrd()
Dim nTotal_Sales As Double, nTotal_in As Double
    With grid1
    .RowHeight(0) = 800
    .WordWrap = True
    
'                                           0       1           2       3       4               5           6               7           8
    .FormatString = retFormatString(Array("ﬂÊœ", "«·›—⁄", "„»Ì⁄« ", " ﬂ·›…", "«—»«Õ „»Ì⁄« ", "„’«—Ì›", "’«›Ì «—»«Õ", "‰”»… «·—»Õ", "«·‰”»… ··≈Ã„«·Ï"))
    
    .ColWidth(1) = 4000
    .ColWidth(2) = 1700
    .ColWidth(3) = 1700
    .ColWidth(4) = 1700
    .ColWidth(5) = 0
    .ColWidth(6) = 0
    .ColWidth(7) = 1700
    .ColWidth(8) = 1700
    
    .ColHidden(0) = True
    For i = 2 To .Cols - 1
        .ColDataType(i) = flexDTDouble
    Next
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    For i = 1 To .Rows - 1
        .TextMatrix(i, 4) = Val(.TextMatrix(i, 2)) - Val(.TextMatrix(i, 3))
        .TextMatrix(i, 6) = Val(.TextMatrix(i, 4)) - Val(.TextMatrix(i, 5))
        If Val(.TextMatrix(i, 3)) <> 0 Then .TextMatrix(i, 7) = Round(Val(.TextMatrix(i, 4)) / Val(.TextMatrix(i, 3)) * 100, 2)
    Next i
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 3, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0.00", vbRed, vbYellow, True, "  "
    If .Rows > 1 Then
        .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
        If Val(.TextMatrix(1, 5)) <> 0 Then .TextMatrix(1, 7) = Round(Val(.TextMatrix(1, 4)) / Val(.TextMatrix(1, 3)) * 100, 2)
    End If
    For i = 2 To .Rows - 1
        If Val(.TextMatrix(1, 4)) <> 0 Then .TextMatrix(i, 8) = Round(Val(.TextMatrix(i, 4)) / Val(.TextMatrix(1, 4)) * 100, 2)
    Next i
    .Subtotal flexSTSum, -1, 8, "#0.00", vbRed, vbYellow, True, "  "
    
    xtotal.Caption = IIf(grid1.Rows > 1, grid1.Rows - 2, "")
    
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set VsTItem = Nothing
End Sub

Private Sub xDesca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    FilterGrd grid1, xDescA.Text, 1
End If
End Sub
Private Sub xitem_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    ItemsLookupAll Me, oSearchItem
End If
End Sub
Sub myProc()
xItem.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
xDescA.Text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
Unload oSearchItem
End Sub
Private Sub xDate1_Validate(Cancel As Boolean)
With xdate1
If Not IsDate(.Text) Then
    .Text = ""
Else
    .Text = Format(.Text, "dd-mm-yyyy")
End If
End With
End Sub
Private Sub xDate2_Validate(Cancel As Boolean)
With xDate2
If Not IsDate(.Text) Then
    .Text = ""
Else
    .Text = Format(.Text, "dd-mm-yyyy")
End If
End With
End Sub
Private Sub Handlecontrols()
cmdPrint.Enabled = grid1.Rows > 1
End Sub


