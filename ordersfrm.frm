VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Ordersfrm 
   Caption         =   "ÿ·»Ì«  „ÊœÌ·"
   ClientHeight    =   7950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15195
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   15195
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   2835
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   90
      Width           =   12255
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ «·„ﬁ«”« "
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7695
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label xcounts 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6705
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   585
         Width           =   945
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   8685
      End
      Begin VB.Label xScal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   585
         Width           =   2070
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "«·„ﬁ«”«  :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10980
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   1050
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·„ÊœÌ· :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   10890
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label xModelNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   2070
      End
   End
   Begin VB.Frame Frame4 
      Height          =   780
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   6975
      Width           =   3615
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "ordersfrm.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1680
      End
      Begin Threed.SSCommand cmdPrint2 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   1755
         TabIndex        =   9
         Top             =   135
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1058
         _Version        =   196610
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
         Picture         =   "ordersfrm.frx":241E
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2295
      Top             =   315
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   4455
      Top             =   1575
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   2475
      Top             =   1125
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   270
      Top             =   1350
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5775
      Left            =   90
      TabIndex        =   0
      Top             =   1170
      Width           =   15000
      _cx             =   26458
      _cy             =   10186
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
      Rows            =   50
      Cols            =   10
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "Ordersfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sModelNo As String
Dim clist1 As String
Dim CardTable As New ADODB.Recordset
Dim con As New ADODB.Connection
Private Sub myload()
Dim cFilter As String
On Error GoTo myerror
Dim cField As String

cField = myiif("type = '2' ", "[IN]  ") & " AS purch"

cField = cField & turn(cField, ",") & _
        myiif("type = '7' ", "[OUT] ") & " AS purch_ret"

cField = cField & turn(cField, ",") & _
        myiif("type = '2' OR type = '7'", "([IN] - [OUT]) ") & " AS purch_net"
        
cField = cField & turn(cField, ",") & _
        myiif("type = '6'", "([OUT])") & " AS sales"

cField = cField & turn(cField, ",") & _
        myiif("", "([IN] - [OUT])") & " AS BALANCE"

cField = cField & turn(cField, ",") & _
          "''" & " AS SALES_RATE"

   
cString = " SELECT file1_10.CODE , FILE4_10.DESCA , " & _
        cField & _
        ", ORDERS1.QUANT" & _
        ", CONVERT(VARCHAR(10),ORDERS1.[DATE],111)" & _
        ", ORDERS2.QUANT" & _
        ", CONVERT(VARCHAR(10),ORDERS2.[DATE],111)" & _
        " FROM  FILE1_10 INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE LEFT JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM" & _
        " LEFT JOIN ORDERS1 ON (FILE1_10.MODELNO =  ORDERS1.MODELNO AND FILE1_10.CODE = ORDERS1.CODE) " & _
        " LEFT JOIN ORDERS2 ON (FILE1_10.MODELNO =  ORDERS2.MODELNO AND FILE1_10.CODE = ORDERS2.CODE)"
cString = cString & turn(cString) & " FILE1_10.MODELNO = " & MyParn(xmodelno.Caption)
cString = cString & " GROUP BY FILE1_10.CODE, FILE4_10.DESCA,FILE1_10.MODELNO,ORDERS1.QUANT,ORDERS1.[DATE],ORDERS2.QUANT,ORDERS2.[DATE]"
cString = cString & " order by FILE1_10.code"
data1.RecordSource = cString
data1.Refresh
Fixgrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub CmdPrint2_Click()
Dim cHead1 As String
Dim cHead2 As String
cHead1 = "ÿ»«⁄… ÿ·»Ì«  «·„ÊœÌ· " & xmodelno.Caption
cHead2 = xDesca.Caption
For I = 2 To 7
    grid1.ColHidden(I) = True
Next
PrintGrd.doprint grid1, 0.9, 0, cHead1, , , False, False, 8
PrintGrd.Show 1
For I = 2 To 7
    grid1.ColHidden(I) = False
Next
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim nCode As Integer
'If Col = 8 Then GRID1.TextMatrix(Row, 9) = Format(Date, "DD-MM-YYYY")
'If Col = 10 Then GRID1.TextMatrix(Row, 10) = Format(Date, "DD-MM-YYYY")
'GRID1.Refresh
If Not validRow(Row) Then Exit Sub

If Not myReplace(Row) Then
   Inform "·„ Ì „ Õ›Ÿ «·»Ì«‰« "
   myload
Else
    If Col = 8 Or Col = 10 And Val(grid1.TextMatrix(Row, Col)) = 0 Then grid1.TextMatrix(Row, Col + 1) = ""
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myload
End Sub
Private Sub grid1_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 Then
    'If Not validRow(OldRow) Then grid1.RemoveItem OldRow
    If Not validRow(OldRow) Then myload
End If
End Sub
Private Sub grid1_EnterCell()
If (grid1.Col < 8) Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim cCode As String
On Error GoTo myerror
If KeyCode = 46 And grid1.Row <> 0 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› !! „Ê«›ﬁ", vbOKCancel + vbDefaultButton2) = vbOK Then
        If Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) <> "" Then
            con.BeginTrans
            con.Execute "delete from file1_10 where item = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
If Err.Number <> 0 Then MsgBox Err.Description
con.RollbackTrans
myload
End Sub
Private Sub Form_Load()
bEdit = True
openCon con

Set grid1.DataSource = data1
data1.ConnectionString = strCon

With grid1
xmodelno.Caption = sModelNo
Dim aRet As Variant, cString
cString = "SELECT TOP (1) desca, dbo.min_max_scal_modelNo(modelno) , COUNT(DISTINCT SCAL) FROM         FILE1_10 "
cString = cString & turn(cString) & " FILE1_10.MODELNO = " & MyParn(xmodelno.Caption)
cString = cString & " GROUP BY desca, dbo.min_max_scal_modelNo(modelno) "
aRet = aGetDesca(cString)
If UBound(aRet) > 0 Then
    xDesca.Caption = aRet(1) & ""
    xscal.Caption = aRet(2) & ""
    xcounts.Caption = aRet(3) & ""
End If

myload
If grid1.Rows > 1 Then
    grid1.Row = 1
    grid1.Col = 8
End If

'grid1.Row = grid1.Rows - 1
'grid1.Col = 1
'grid1.ShowCell grid1.Rows - 1, 1
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 Then myload
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 8 Or Col = 10 Then
    If Trim(grid1.EditText) <> "" And Not IsNumeric(grid1.EditText) Then
        MsgBox "«·ﬂ„Ì… €Ì— ’ÕÌÕ…"
        Cancel = True
    End If
'    If Col = 8 Then GRID1.TextMatrix(Row, 9) = Format(Date, "DD-MM-YYYY")
'    If Col = 10 Then GRID1.TextMatrix(Row, 9) = Format(Date, "DD-MM-YYYY")
    grid1.Refresh
End If
If Col = 9 Or Col = 11 Then
    If Not IsDate(grid1.EditText) Then
        MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
        Cancel = True
    Else
        grid1.EditText = Format(grid1.EditText, "DD-MM-YYYY")
    End If
End If

'If Col = 2 Then
'    If Trim(grid1.EditText) = "" Then
'        MsgBox "„Ã„Ê⁄… «·’‰› „ÿ·Ê»…"
'        Cancel = True
'    End If
'End If
'If Col = 3 Then
'    If (Not IsNumeric(grid1.EditText)) Then
'        Cancel = True
'    ElseIf Val(grid1.EditText) = 0 Then
'        Cancel = True
'    End If
'End If
Exit Sub
myerror:
On Error Resume Next
If Err.Number <> 0 Then MsgBox Err.Description
con.RollbackTrans
myload
Err.Clear
End Sub
Private Sub Fixgrd()
With grid1
.ColWidth(0) = 800
.ColWidth(1) = 3900
.ColWidth(2) = 900
.ColWidth(3) = 900
.ColWidth(4) = 900
.ColWidth(5) = 900
.ColWidth(6) = 900
.ColWidth(7) = 900
.ColWidth(8) = 900
.ColWidth(9) = 1400
.ColWidth(10) = 900
.ColWidth(11) = 1400
.ColFormat(7) = "#%"

.TextMatrix(0, 0) = "ﬂÊœ"
.TextMatrix(0, 1) = "«·„Ê—œ"
.TextMatrix(0, 2) = "„‘ —Ì« "
.TextMatrix(0, 3) = "„— Ã⁄« "
.TextMatrix(0, 4) = "’«›Ï"
.TextMatrix(0, 5) = "„»Ì⁄« "
.TextMatrix(0, 6) = "—’Ìœ"
.TextMatrix(0, 7) = "‰”»… »Ì⁄"
.TextMatrix(0, 8) = "ﬂ„Ì… ÿ·»Ì…"
.TextMatrix(0, 9) = " «—ÌŒ «·ÿ·»Ì…"
.TextMatrix(0, 10) = "ﬂ„Ì… ‘ﬂÊÌ"
.TextMatrix(0, 11) = " «—ÌŒ «·‘ﬂÊÌ"
.ColHidden(0) = True
.RowHeight(0) = 600
.WordWrap = True
For I = 1 To grid1.Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next

.Cell(flexcpBackColor, 1, 0, .Rows - 1, 7) = &H8000000F
Dim nRow As Long
For nRow = 1 To .Rows - 1
    If Val(.TextMatrix(nRow, 4)) <> 0 Then
        .TextMatrix(nRow, 7) = Val(.TextMatrix(nRow, 5)) / Val(.TextMatrix(nRow, 4))
    End If
Next
End With
End Sub
Private Function validRow(Row) As Boolean
'If Val(GRID1.TextMatrix(Row, 8)) > 0 Then
'    If Not IsDate(GRID1.TextMatrix(Row, 9)) Then Exit Function
'End If
'
'If Val(GRID1.TextMatrix(Row, 10)) > 0 Then
'    If Not IsDate(GRID1.TextMatrix(Row, 11)) Then Exit Function
'End If
If IsDate(grid1.TextMatrix(Row, 9)) Then
    If (Not IsNumeric(grid1.TextMatrix(Row, 8))) Or Val(grid1.TextMatrix(Row, 8)) < 0 Then Exit Function
End If

If IsDate(grid1.TextMatrix(Row, 11)) Then
    If (Not IsNumeric(grid1.TextMatrix(Row, 10))) Or Val(grid1.TextMatrix(Row, 10)) < 0 Then Exit Function
End If
validRow = True
End Function
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
     CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    If Not (Col = 2) Then CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 1 Then
    If Col < 8 Then
        grid1.Select Row, 8
    ElseIf Col >= 8 Then
        grid1.Select Row, Col + 1
    End If
ElseIf Row < grid1.Rows - 1 Then
    grid1.Row = Row + 1
    grid1.Select Row + 1, 8
    If Not grid1.RowIsVisible(Row + 1) Then grid1.ShowCell Row + 1, 8
End If
End Sub
Private Sub myaddItem()
With grid1
    .AddItem ""
End With
End Sub
Private Function myReplace(Row As Long) As Boolean
Dim aInsert(5, 1), cString As String

aInsert(0, 0) = "MODELNO"
aInsert(0, 1) = addstring(xmodelno.Caption)

aInsert(1, 0) = "CODE"
aInsert(1, 1) = addstring(grid1.TextMatrix(Row, 0))

aInsert(4, 0) = "SCAL"
aInsert(4, 1) = addstring(xscal.Caption)

aInsert(5, 0) = "C_SCAL"
aInsert(5, 1) = addvalue(xcounts.Caption)

aInsert(2, 0) = "QUANT"
aInsert(2, 1) = Val(grid1.TextMatrix(Row, 8))

aInsert(3, 0) = "DATE"
If grid1.TextMatrix(Row, 9) = "" Then
    aInsert(3, 1) = "GETDATE()"
Else
    aInsert(3, 1) = addDate(grid1.TextMatrix(Row, 9))
End If

cString = "DELETE FROM ORDERS1"
cString = cString & turn(cString) & " CODE = " & MyParn(grid1.TextMatrix(Row, 0))
cString = cString & turn(cString) & " MODELNO = " & MyParn(xmodelno)

con.BeginTrans
On Error GoTo myerror
con.Execute cString
If Val(grid1.TextMatrix(Row, 8)) > 0 Then con.Execute CreateInsert(aInsert, "ORDERS1")

aInsert(2, 0) = "QUANT"
aInsert(2, 1) = Val(grid1.TextMatrix(Row, 10))

aInsert(3, 0) = "DATE"
If grid1.TextMatrix(Row, 11) = "" Then
    aInsert(3, 1) = "GETDATE()"
Else
    aInsert(3, 1) = addDate(grid1.TextMatrix(Row, 11))
End If

cString = "DELETE FROM ORDERS2"
cString = cString & turn(cString) & " CODE = " & MyParn(grid1.TextMatrix(Row, 0))
cString = cString & turn(cString) & " MODELNO = " & MyParn(xmodelno)
con.Execute cString
If Val(grid1.TextMatrix(Row, 10)) > 0 Then con.Execute CreateInsert(aInsert, "ORDERS2")
con.CommitTrans
myReplace = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Function

Private Sub Label2_Click()

End Sub
