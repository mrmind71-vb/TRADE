VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form FIXINV2 
   Caption         =   " ⁄œÌ· »Ì«‰«  «·„ÊœÌ·"
   ClientHeight    =   8070
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD_PrintBar 
      Caption         =   " ⁄œÌ· »Ì«‰«  «·„ÊœÌ·« "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1980
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   7380
      Width           =   2085
   End
   Begin VB.CommandButton CMD_EXIT 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   7380
      Width           =   1770
   End
   Begin Threed.SSCommand xNameCode 
      Height          =   390
      Left            =   8250
      TabIndex        =   0
      Top             =   135
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   688
      _Version        =   196610
      BackColor       =   13822956
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "-"
      ButtonStyle     =   3
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   45
      Top             =   180
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   2655
      Top             =   135
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
      Height          =   6630
      Left            =   135
      TabIndex        =   1
      Top             =   630
      Width           =   14910
      _cx             =   26300
      _cy             =   11695
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16777152
      ForeColorSel    =   128
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
Attribute VB_Name = "FIXINV2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub

Private Sub CMD_PrintBar_Click()
Dim cModel As String, cFact As String, cMosm As String, cModelFact As String, cModelNo As String
'On Error GoTo myerror
con.BeginTrans
If MsgBox("Õ›Ÿ «· ⁄œÌ·« ", vbOKCancel) = vbOK Then
    With grid1
        For i = 2 To .Rows - 1
            If Trim(.TextMatrix(i, 9)) <> "" Then
                cModel = .TextMatrix(i, .Cols - 1)
                cFact = .TextMatrix(i, 7)
                cMosm = .TextMatrix(i, 6)
                cModelFact = DelZero(.TextMatrix(i, 8))
                cModelNo = cFact & cMosm & RetZero(cModelFact, 10)
                con.Execute " UPDATE FILE1_10 SET MODEL = " & MyParn(cModel) & " , MODELNO = " & MyParn(cModelNo) & " , FACT = " & MyParn(cFact) & " , MODELFACT0 = " & MyParn(DelZero(cModelFact)) & " , MOSM = " & MyParn(cMosm) & " , MODELFACT = " & MyParn(RetZero(cModelFact, 10)) & " WHERE MODEL = " & MyParn(.TextMatrix(i, 0))
                
                con.Execute " DELETE FROM FILE1_10H WHERE MODEL = " & MyParn(.TextMatrix(i, 0))
                cStr1 = " INSERT INTO FILE1_10H  (MODEL, desca, [GROUP], OKAZ, FACT, RATE, MOSM, MODELNO, MODELFACT, code, SUPP, MODELFACT0) " & _
                        " SELECT     MODEL, DESCA, [GROUP], OKAZ, FACT, RATE, MOSM, MODELNO, MODELFACT, code, SUPP, MODELFACT0 FROM         FILE1_10H_V WHERE MODEL = " & MyParn(cModel)
                            
                con.Execute cStr1
            End If
        Next i
    End With
End If
con.CommitTrans
Unload Me
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
MsgBox "·„ Ì „ «· ⁄œÌ·"
Err.Clear
End Sub

Private Sub Form_Load()
con.CursorLocation = adUseClient
con.Open strCon
Set grid1.DataSource = data1
data1.ConnectionString = strCon

With grid1
grid1.Rows = 2

grid1.FixedRows = 2
'                           0               1               2               3                   4               5               6           7               8                       9
    cString = "SELECT FILE1_10.MODEL , FILE1_10.mosm,FILE1_10.fact , FILE1_10.modelfact , FILE1_10.desca ,SUM(Quant)  , FILE1_10.mosm,FILE1_10.fact,FILE1_10.modelfact , ' ' AS NEWMPDEL " & _
          " FROM FILE7_20 LEFT JOIN FILE1_10 ON FILE7_20.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(Purchasefrm.xDoc_No.Text) & " GROUP BY FILE1_10.MOSM,FILE1_10.FACT,FILE1_10.MODELFACT , FILE1_10.DESCA , FILE1_10.MODEL order by FILE1_10.MOSM,FILE1_10.FACT,FILE1_10.MODELFACT "
    data1.RecordSource = cString
    data1.Refresh
End With
Fixgrd
End Sub
Private Sub Fixgrd()
With grid1
'                   0          1        2            3              4               5          6             7             8           9
.MergeCells = flexMergeFree
.MergeRow(0) = True

.TextMatrix(0, 0) = "ﬁ»· «· ⁄œÌ·"
.TextMatrix(0, 1) = "ﬁ»· «· ⁄œÌ·"
.TextMatrix(0, 2) = "ﬁ»· «· ⁄œÌ·"
.TextMatrix(0, 3) = "ﬁ»· «· ⁄œÌ·"
.TextMatrix(0, 5) = "ﬁ»· «· ⁄œÌ·"

.TextMatrix(0, 6) = "»⁄œ «· ⁄œÌ·"
.TextMatrix(0, 7) = "»⁄œ «· ⁄œÌ·"
.TextMatrix(0, 8) = "»⁄œ «· ⁄œÌ·"
.TextMatrix(0, 9) = "»⁄œ «· ⁄œÌ·"

.TextMatrix(1, 0) = "„ÊœÌ·"
.TextMatrix(1, 1) = "„Ê”„"
.TextMatrix(1, 2) = "„’‰⁄"
.TextMatrix(1, 3) = "≈”„ «·„ÊœÌ·"
.TextMatrix(1, 5) = "ﬂ„Ì…"

.TextMatrix(1, 6) = "„Ê”„"
.TextMatrix(1, 7) = "„’‰⁄"
.TextMatrix(1, 8) = "—ﬁ„ „ÊœÌ·"
.TextMatrix(1, 9) = "«·„ÊœÌ· «·ÃœÌœ"


.WordWrap = True
.Cols = 10
.ColWidth(0) = 0
.ColWidth(1) = 700
.ColWidth(2) = 2000
.ColWidth(3) = 1000
.ColWidth(4) = 3000
.ColWidth(5) = 1000
.ColWidth(6) = 700
.ColWidth(7) = 2000
.ColWidth(8) = 1000
.ColWidth(9) = 1500
.ColComboList(9) = "..."

.ColHidden(7) = (Purchasefrm.xSupp.Value = 0)
.ColComboList(1) = StrList("SELECT MOSM, MOSM FROM MOSM")
.ColComboList(2) = StrList("SELECT CODE , DESCA FROM FACT ")
.ColComboList(6) = StrList("SELECT MOSM, MOSM FROM MOSM")
.ColComboList(7) = StrList("SELECT CODE , DESCA FROM FACT ")
.SubtotalPosition = flexSTBelow
.Subtotal flexSTSum, -1, 5, "#0", vbYellow, vbRed, True, "≈Ã„«·Ï "
If .Rows > 1 Then
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .Cell(flexcpBackColor, 2, 0, .Rows - 1, .Cols - 1) = vbYellow
End If
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub Grid1_Click()
With grid1
    If .Col = 6 Or .Col = 7 Or .Col = 8 Or .Col = 9 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub GRID1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With grid1
If MsgBox(" €ÌÌ— —ﬁ„ «·„ÊœÌ·", vbYesNo + vbDefaultButton2) = vbYes Then
    For i = 2 To .Rows - 1
        UpdateModel i
    Next i
End If
End With
End Sub
Sub UpdateModel(nRow)
    Dim cModel As String
    With grid1
        If .TextMatrix(nRow, 6) = "" Or .TextMatrix(nRow, 7) = "" Or .TextMatrix(nRow, 8) = "" Then
            .Select nRow, 0, nRow, .Cols - 1
            .ShowCell nRow, 1
            .Cell(flexcpFontStrikethru, nRow, 0, nRow, .Cols - 1) = True
            Exit Sub
        End If
        If Purchasefrm.xSupp.Value = 1 Then
            cModel = .TextMatrix(nRow, 7) & .TextMatrix(nRow, 6) & RetZero(.TextMatrix(nRow, 8), 10) & Purchasefrm.xCode.Text
        Else
            cModel = .TextMatrix(nRow, 7) & .TextMatrix(nRow, 6) & RetZero(.TextMatrix(nRow, 8), 10)
        End If
        If cModel <> .TextMatrix(nRow, 0) Then
            If GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(cModel)) = "" Then
                .TextMatrix(nRow, .Cols - 1) = cModel
            Else
                .Select nRow, 0, nRow, .Cols - 1
                .ShowCell nRow, 1
                .Cell(flexcpFontStrikethru, nRow, 0, nRow, .Cols - 1) = True
                Exit Sub
            End If
        End If
    End With
End Sub

