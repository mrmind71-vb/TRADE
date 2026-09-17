VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form sales_returnsfrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "„— Ã⁄«  «’‰«›"
   ClientHeight    =   10200
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   8370
      Width           =   3390
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   180
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "sales_returns2.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand CMD_SEND 
         Height          =   555
         Left            =   1620
         TabIndex        =   4
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "sales_returns2.frx":2323
         Caption         =   " ”ÃÌ· «·„— Ã⁄"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_returns2.frx":498A
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4110
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   20220
      _cx             =   35666
      _cy             =   7250
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
      Cols            =   14
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
      WordWrap        =   -1  'True
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   4110
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4230
      Width           =   20220
      _cx             =   35666
      _cy             =   7250
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
      Cols            =   13
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
      WordWrap        =   -1  'True
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -765
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
      Top             =   -765
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   18135
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   8325
      Width           =   2220
      Begin Threed.SSCommand cmdSelectAll 
         Height          =   555
         Left            =   45
         TabIndex        =   6
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "sales_returns2.frx":6C80
         Caption         =   "«Œ Ì«— ﬂ· «·›« Ê—…"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "sales_returns2.frx":9156
      End
   End
End
Attribute VB_Name = "sales_returnsfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Public sDoc_no_ret As String
Public sOnline_doc As String
Public sid_lading

Public myform As Form
Dim bStopCell As Boolean
Dim con As New ADODB.Connection
Dim conServer As New ADODB.Connection
Private Sub CMD_SEND_Click()
If Not myValid Then Exit Sub
Dim sNewDoc As String
sNewDoc = myReplacereturn
If sNewDoc = "" Then Exit Sub
MsgBox " „ «· —ÃÌ· «·Ì „” ‰œ „»Ì⁄«  —ﬁ„ " & sNewDoc
myform.myproc_return sNewDoc
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub cmdSelectAll_Click()
'Dim i As Long
'grid1.Rows = 1
'myloadgrd2
'For i = 1 To grid2.Rows - 2
'    addRow i
'Next
'
'Fixgrd2
'fixGrd
'
'CalcTotals
grid1.Rows = 1
myLoadGrd
addRow
End Sub

Private Sub Form_Load()
OpenCon con

Set grid1.DataSource = data1
Set grid2.DataSource = data2

myloadgrd2
fixGrd
CalcTotals

CellPos 13, 0, grid1.Cols - 1
End Sub
Private Sub myLoadGrd()
'Dim cString As String
'cString = "SELECT FILE6_20.ITEM," & _
'          "FILE1_10.MOSM," & _
'          "FACT.DESCA," & _
'          "FILE1_10.MODELFACT0," & _
'          "FILE1_10.DESCA," & _
'          "FILE1_10.SCAL," & _
'          "FILE1_10.COLOR," & _
'          "FILE6_20.PRICE_C," & _
'          "FILE6_20.Quant," & _
'          "CASE WHEN RATE > 100 THEN 0 ELSE ROUND(FILE6_20.PRICE * (1 - (FILE6_20H.RATE/100)),2) END ," & _
'          "FILE6_20.TOTAL," & _
'          "FILE6_25.DESCA," & _
'          "FILE6_20.ID" & _
'          " FROM FILE6_20 " & _
'          " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20.DOC_NO " & _
'          " LEFT JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM" & _
'          " LEFT JOIN FACT ON FACT.CODE = FILE1_10.FACT" & _
'          " LEFT JOIN FILE6_25 ON FILE6_20.MAN  = FILE6_25.CODE" & _
'cString = cString & " WHERE FILE6_20.DOC_NO = " & MyParn(sDoc_no)
'cString = cString & " ORDER BY FILE6_20.ID"
'
'Set data1.Recordset = myCmd(cString, con)
'fixGrd
End Sub
Private Sub myloadgrd2()
Dim aPrm As Variant
aPrm = AddFlag(aPrm, "DOC_NO", sDoc_no_ret)
aPrm = AddFlag(aPrm, "DOC_NO_IG", sDoc_no)
Set data2.RecordSet = mycmd("dbo.sp_SALES_RETRUN", con, adStoredProc, aPrm)
Fixgrd2
End Sub
Sub Fixgrd2()
With grid2
    .RowHeight(0) = 600

    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê”„"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "„ÊœÌ· «·„’‰⁄"
    
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«··Ê‰"
    .TextMatrix(0, 7) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(0, 8) = "«·ﬂ„Ì…"
    .TextMatrix(0, 9) = "«·”⁄—"
    .TextMatrix(0, 10) = "«·≈Ã„«·Ì"
    .TextMatrix(0, 11) = "«·»«∆⁄"
    
    .TextMatrix(0, 12) = "„— Ã⁄"
    
    .TextMatrix(0, 13) = "«·»«ﬁÌ"
    .TextMatrix(0, 14) = "«· ﬂ·›…"
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1400
    .ColWidth(4) = 3000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    .ColWidth(13) = 1000
    .ColWidth(14) = 1000

    .ColHidden(.Cols - 4) = True
    .ColHidden(.Cols - 3) = True
    .ColHidden(.Cols - 2) = True
    .ColHidden(.Cols - 1) = True

    .ColWidth(grid2.Cols - 1) = 1000
    
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub CalcTotals()


grid1.Subtotal flexSTClear
grid2.Subtotal flexSTClear

With grid2
For i = 1 To .Rows - 1
    nFound = grid1.FindRow(.TextMatrix(i, 0), , 0)
    If nFound <> -1 Then
         .TextMatrix(i, 12) = .ValueMatrix(i, .Cols - 1) + grid1.ValueMatrix(nFound, 8)
         .TextMatrix(i, 13) = .ValueMatrix(i, 8) - .ValueMatrix(i, 12)
    Else
         .TextMatrix(i, 12) = .ValueMatrix(i, .Cols - 1)
         .TextMatrix(i, 13) = .ValueMatrix(i, 8) - .ValueMatrix(i, 12)
    End If
    
    If .ValueMatrix(i, 13) = 0 Then
        .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0C0FF
    ElseIf .ValueMatrix(i, 13) > 0 And .ValueMatrix(i, 13) < .ValueMatrix(i, 8) Then
        .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H80FFFF
    ElseIf .ValueMatrix(i, 12) = .ValueMatrix(i, .Cols - 1) Then
        .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbWhite
    End If
Next

.SubtotalPosition = flexSTBelow
.Subtotal flexSTSum, -1, 8, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
.Subtotal flexSTSum, -1, 10, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
.Subtotal flexSTSum, -1, 11, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
.Subtotal flexSTSum, -1, 12, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
.Subtotal flexSTSum, -1, 13, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
End With

With grid1
    .Subtotal flexSTClear
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 8, "#", &HE0E0E0, , True, "«·≈Ã„«·Ì"
    .Subtotal flexSTSum, -1, 10, "#.##", &HE0E0E0, , True, "«·≈Ã„«·Ì"
End With
End Sub
Sub fixGrd()
With grid1
    .RowHeight(0) = 600
    '.Cols = 13
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê”„"
    .TextMatrix(0, 2) = "«·„’‰⁄"
    .TextMatrix(0, 3) = "„ÊœÌ· «·„’‰⁄"
    
    .TextMatrix(0, 4) = "«·’‰›"
    .TextMatrix(0, 5) = "„ﬁ«”"
    .TextMatrix(0, 6) = "«··Ê‰"
    .TextMatrix(0, 7) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(0, 8) = "«·ﬂ„Ì…"
    .TextMatrix(0, 9) = "«·”⁄—"
    .TextMatrix(0, 10) = "«·≈Ã„«·Ì"
    
    .TextMatrix(0, 11) = "«· ﬂ·›…"
    .TextMatrix(0, 12) = "DESCA2"
    .TextMatrix(0, 13) = "SKU"
        
    .ColWidth(0) = 1000
    .ColWidth(1) = 1000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1400
    .ColWidth(4) = 3000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1200
    .ColWidth(8) = 750
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    
    .ColHidden(.Cols - 3) = True
    .ColHidden(.Cols - 2) = True
    .ColHidden(.Cols - 1) = True
        
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set sales_returnsfrm = Nothing
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
CalcTotals
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 8 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Val(grid1.EditText) <= 0 Then
    Cancel = True
    Exit Sub
End If

Dim nFound As Long
nFound = grid2.FindRow(grid1.TextMatrix(Row, 0), , 0)
If grid2.ValueMatrix(nFound, 8) - grid2.ValueMatrix(nFound, grid1.Cols - 1) < Val(grid1.EditText) Then
    Cancel = True
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If bStopCell Then
    bStopCell = False
ElseIf Col < 8 Then
    grid1.Col = 8
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 8, 8)
    grid1.ShowCell grid1.Row, 8
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    'If Col = 4 Then Exit Sub
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    'If grid1.Col = 4 Then Exit Sub
    KeyAscii = 0
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", vbOKCancel) = vbOK Then
        myRemove grid1.Row
        CalcTotals
        bStopCell = True
    End If
End If
End Sub
Private Sub myRemove(Row As Long)
grid1.RemoveItem Row
End Sub
Private Function myValid() As Boolean
Dim i As Long
Dim nFound As Long

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«› „”Ã·…"
    Exit Function
End If

With grid2
For i = 1 To .Rows - 2
    nFound = grid1.FindRow(.TextMatrix(i, 0), , 0)
    If nFound = -1 Then
        If .ValueMatrix(i, 13) > 0 Then
            If MsgBox("«·’‰› " & .TextMatrix(i, 0) & " ·„ Ì”Ã· ›Ï «·›« Ê—…", vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
        End If
    ElseIf .TextMatrix(i, 13) < 0 Then
        MsgBox "«·’‰› " & .TextMatrix(i, 0) & " «ﬂ»— „‰ «·›« Ê—… » " & Abs(.ValueMatrix(i, 13))
        Exit Function
    ElseIf .TextMatrix(i, 13) > 0 Then
        If MsgBox("«·’‰› " & .TextMatrix(i, 0) & "  „ »ﬁÌ »ﬂ„Ì… " & .ValueMatrix(i, 13), vbOKCancel + vbDefaultButton2) <> vbOK Then Exit Function
    End If
Next
myValid = True
End With
End Function
Private Sub Grid2_dblClick()
With grid2
If grid2.Row = 0 Or grid2.Row = grid2.Rows - 1 Then Exit Sub
addRow grid2.Row

End With
End Sub
Private Sub addRow(Optional pRow As Long = -1)
Dim i As Long
Dim Row As Long

'grid1.Subtotal flexSTClear
For Row = IIf(pRow = -1, 1, pRow) To IIf(pRow = -1, grid2.Rows - 2, pRow)
    If grid2.ValueMatrix(Row, 13) > 0 Then
        grid1.AddItem ""
        For i = 0 To 10
            grid1.TextMatrix(grid1.Rows - 1, i) = grid2.TextMatrix(Row, i)
        Next
        grid1.TextMatrix(grid1.Rows - 1, 11) = grid2.TextMatrix(Row, 14)
        grid1.TextMatrix(grid1.Rows - 1, 12) = grid2.TextMatrix(Row, 15)
        grid1.TextMatrix(grid1.Rows - 1, 13) = grid2.TextMatrix(Row, 16)
    End If
Next

CalcTotals
End Sub
Private Function myReplacereturn() As String
    Dim sDoc_New As String
    Dim sDate As String
    Dim cAddress As String
    
    Dim cString As New ChilkatStringBuilder
    cString.Append "SELECT DSALES FROM DSALES"
    cString.Append " WHERE BRANCH = " & MyParn(cBranch)
    sDate = myFormat(myField(cString.GetAsString(), con))
    
    cString.Clear
    
    Dim LadingTable As New ADODB.RecordSet
    cString.Append "Select * from lading_bill where id = " & sid_lading
    Set LadingTable = mycmd(cString.GetAsString, con)
    If LadingTable.EOF Then
        MsgBox "·« ÌÊÃœ »Ê·Ì’… ‘Õ‰ „”»ﬁ…"
        Exit Function
    End If
    
    
    cString.Clear
        
    cString.Append " SELECT DOC_NO," & _
            " DATE," & _
            " NAME," & _
            " PHONE," & _
            "Shipping_City," & _
            "CITY," & _
            "STREET," & _
            "Payment_Method," & _
            "ADDRESS" & _
            " FROM FILE6_20H "
    
    cString.Append " WHERE FILE6_20H.DOC_NO = " & MyParn(sDoc_no_ret)
           
    Dim loctable As New ADODB.RecordSet
    Set loctable = mycmd(cString.GetAsString, con)
    
    If loctable.EOF And loctable.BOF Then
        MsgBox "«·„” ‰œ €Ì— „”Ã·"
        Exit Function
    End If
    
    aInsert = AddFlag(Empty, "SALES_RET", addstring(sDoc_no_ret))
    aInsert = AddFlag(aInsert, "CODE", addstring("0000"))
    aInsert = AddFlag(aInsert, "[Date]", addDate(sDate))
    aInsert = AddFlag(aInsert, "STORE", addstring(cBranchStore))
    aInsert = AddFlag(aInsert, "BOX", addstring(cManBox))
    aInsert = AddFlag(aInsert, "userName", addstring(cusername))
    aInsert = AddFlag(aInsert, "branch", addstring(cBranch))
    aInsert = AddFlag(aInsert, "MAN", addstring("0001"))
    aInsert = AddFlag(aInsert, "PHONE", addstring(loctable!phone))
    aInsert = AddFlag(aInsert, "isNew", "1")
    aInsert = AddFlag(aInsert, "ONLINE", "2")
    aInsert = AddFlag(aInsert, "USER_IP", addstring(GetComputerName))
    aInsert = AddFlag(aInsert, "online_Doc", addstring(sOnline_doc))
    aInsert = AddFlag(aInsert, "ADDRESS", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "[NAME]", addstring(loctable!Name))
    aInsert = AddFlag(aInsert, "[city]", addstring(loctable!city))
    aInsert = AddFlag(aInsert, "[Shipping_City]", addstring(loctable!Shipping_City))
    aInsert = AddFlag(aInsert, "[street]", addstring(loctable!Street))
    aInsert = AddFlag(aInsert, "[Payment_Method]", addstring(loctable!Payment_Method))
    aInsert = AddFlag(aInsert, "[ship]", addstring(LadingTable!SHIP))
    aInsert = AddFlag(aInsert, "[ship_no]", addstring(LadingTable!SHIP_NO))
    aInsert = AddFlag(aInsert, "[Charge1]", mRound(LadingTable!CHARGE1))
    aInsert = AddFlag(aInsert, "[Charge2]", mRound(LadingTable!charge2))
'    aInsert = AddFlag(aInsert, "[isclosed]", "1")
'    aInsert = AddFlag(aInsert, "[PRINTED]", "1")

    sDoc_New = NewflagDoc(sDate, cManBox, con)
    aInsert = AddFlag(aInsert, "[doc_no]", addstring(sDoc_New))
    
    con.BeginTrans
    con.Execute addInsert(aInsert, "FILE6_20H")
            
    cString.Clear
    cString.Append "insert into FILE6_20(" & _
                    "DOC_NO," & _
                    "ITEM," & _
                    "QUANT," & _
                    "PRICE," & _
                    "[ROW]," & _
                    "[MAN]," & _
                    "[USER_IP]," & _
                    "[PRICE_C]," & _
                    "[COST]," & _
                    "[DESCA2]," & _
                    "[SKU]" & _
                    ")"
    
    With grid1
    cString.Append " VALUES "
    
    For i = 1 To .Rows - 2
        If .ValueMatrix(i, 8) > 0 Then
            cString.Append "("
            cString.Append addstring(sDoc_New) & ","
            cString.Append addstring(.TextMatrix(i, 0)) & ","
            cString.Append -1 * .ValueMatrix(i, 8) & ","
            cString.Append .ValueMatrix(i, 9) & ","
            cString.Append i & ","
            cString.Append addstring("0001") & ","
            cString.Append addstring(GetComputerNamecIpName) & ","
            cString.Append .ValueMatrix(i, 7) & ","
            cString.Append .ValueMatrix(i, 11) & ","
            cString.Append addstring(.TextMatrix(i, 12)) & ","
            cString.Append addstring(.TextMatrix(i, 13)) & ")"
            cString.Append ","
        End If
    Next
    cString.Shorten 1
    End With
    
    con.Execute cString.GetAsString()
    
    UpdateInvTotal sDoc_New, con
    
    con.Execute " UPDATE LADING_BILL SET DOC_NO_RET = " & addstring(sDoc_New) & " WHERE ID = " & sid_lading, nAffect
    If nAffect = 0 Then
        MsgBox "·„ Ì „ Õ›Ÿ ›« Ê—… «·„— Ã⁄ ›Ï «·»Ê«·’ «·„”»ﬁ…"
        con.RollbackTrans
        Exit Function
    End If

    con.CommitTrans
    
    myReplacereturn = sDoc_New
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
