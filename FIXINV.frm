VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FIXINV 
   Caption         =   " ⁄œÌ· »Ì«‰«  «·„ÊœÌ·"
   ClientHeight    =   8070
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   11280
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_ok 
      Caption         =   " €Ì— «·„Ê—œ"
      Height          =   375
      Left            =   7785
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   135
      Width           =   2445
   End
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
      TabIndex        =   2
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
      TabIndex        =   1
      Top             =   7380
      Width           =   1770
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   45
      Top             =   855
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
      Left            =   270
      Top             =   810
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
      TabIndex        =   0
      Top             =   630
      Width           =   14910
      _cx             =   26300
      _cy             =   11695
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorFixed  =   16761024
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
   Begin MSDataListLib.DataCombo xsupp 
      Height          =   315
      Left            =   10575
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   135
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Text            =   ""
      RightToLeft     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo xMosm 
      Height          =   315
      Left            =   4995
      TabIndex        =   6
      Top             =   135
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   360
      Top             =   450
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   45
      Top             =   630
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
   Begin MSDataListLib.DataCombo xfact 
      Height          =   315
      Left            =   180
      TabIndex        =   8
      Top             =   135
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      RightToLeft     =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„’‰⁄"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   2805
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   180
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„Ê”„"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   7020
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "«·„Ê—œ :"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   14130
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   135
      Width           =   615
   End
End
Attribute VB_Name = "FIXINV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub
Private Sub CMD_OK_Click()
If xSupp.BoundText <> Purchasefrm.xCode.text Then
    If MsgBox(" €Ì— «·„Ê—œ  ·«’‰«› «·›« Ê—…", vbYesNo) = vbYes Then
        con.Execute " update " & Purchasefrm.cFileHeader & " set code = " & addstring(xSupp.BoundText) & " where doc_no = " & MyParn(Purchasefrm.xdoc_no.text)
        con.Execute " update file1_10 set code = " & addstring(xSupp.BoundText) & "  where item in ( select item from " & Purchasefrm.cFile & " where doc_no = " & MyParn(Purchasefrm.xdoc_no.text) & " ) "
        Purchasefrm.xCode.text = xSupp.BoundText
        
        
            Dim cString As String, cWhere As String
            cString = "select code,desca, SUPP , DISC ,SUBCODE  FROM FILE4_10"
            cWhere = cWhere & turn(cWhere, " OR ") & " CODE = " & MyParn(Purchasefrm.xCode.text)
            cWhere = cWhere & turn(cWhere, " OR ") & " SUBCODE = " & MyParn(Purchasefrm.xCode.text)
            cString = cString & turn(cWhere) & cWhere
            aRet = aGetDesca(cString, con)
            If UBound(aRet) > 0 Then
                Purchasefrm.xCode.text = aRet(1)
                Purchasefrm.xCodeDesca.Caption = aRet(2)
                Purchasefrm.xSupp.Value = IIf(aRet(3), 1, 0)
                Purchasefrm.xsup_desca.Caption = IIf(aRet(3), "„ﬂ » Ã„·…", "„’‰⁄")
                Purchasefrm.xr_Discount.text = Val(aRet(4) & "")
                Purchasefrm.xRateSup.Caption = Val(aRet(4) & "")
                Purchasefrm.xsubcode.Caption = aRet(5) & ""
            End If
        With grid1
            For nRow = 1 To .Rows - 1
                UpdateModel nRow
            Next nRow
        End With
    End If
    CMD_PrintBar_Click
End If
End Sub
Private Sub CMD_PrintBar_Click()
Dim cModel As String, cFact As String, cMosm As String, cModelFact As String, cModelNo As String
Dim MyItemTable As New ADODB.Recordset
'On Error GoTo myerror
con.BeginTrans
If MsgBox("Õ›Ÿ «· ⁄œÌ·« ", vbOKCancel) = vbOK Then
    With grid1
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 9)) <> "" Then
                cModel = .TextMatrix(i, .Cols - 1)
                cFact = .TextMatrix(i, 7)
                cMosm = .TextMatrix(i, 6)
                cModelFact = DelZero(.TextMatrix(i, 8))
                cModelNo = cFact & cMosm & RetZero(cModelFact, 10)
                
                If MyItemTable.State = adStateOpen Then MyItemTable.Close
                MyItemTable.Open "SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(.TextMatrix(i, 0)), con, adOpenStatic, adLockReadOnly, adCmdText
                Do While Not MyItemTable.EOF
                    For nRow = 3 To nCountBranch
                        sStore = RetZero(nRow, 2)
                        con.Execute "  INSERT INTO NEW_ITEM (ITEM, BRANCH) VALUES        ( " & addvalue(MyItemTable!Item) & " , " & addstring(sStore) & " )  "
                    Next nRow
                    
                    For nRow = 61 To nCountBranch_fr
                        sStore = RetZero(nRow, 2)
                        con.Execute "  INSERT INTO NEW_ITEM_fr (ITEM, BRANCH) VALUES        ( " & addvalue(MyItemTable!Item) & " , " & addstring(sStore) & " )  "
                    Next nRow
                    
                    MyItemTable.MoveNext
                Loop
                
                con.Execute " UPDATE FILE1_10 SET isdate = getdate() , isnew1 = 1 , isnew2 = 1 , isnew3 = 1 , isnew4 = 1 , isnew5 = 1 , MODEL = " & MyParn(cModel) & " , MODELNO = " & MyParn(cModelNo) & " , FACT = " & MyParn(cFact) & " , MODELFACT0 = " & MyParn(DelZero(cModelFact)) & " , MOSM = " & MyParn(cMosm) & " , MODELFACT = " & MyParn(RetZero(cModelFact, 10)) & " WHERE MODEL = " & MyParn(.TextMatrix(i, 0))
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

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM file4_10 order by desca "
Set xSupp.RowSource = DATA2
xSupp.ListField = "DESCA"
xSupp.BoundColumn = "code"

data3.ConnectionString = strCon
data3.RecordSource = "SELECT * FROM MOSM"
Set xMosm.RowSource = data3
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"

data5.ConnectionString = strCon
data5.RecordSource = "SELECT * FROM fact order by desca"
Set xFact.RowSource = data5
xFact.ListField = "DESCA"
xFact.BoundColumn = "CODE"

With grid1
.Rows = 2
.FixedRows = 2
'                           0               1               2               3                   4               5               6           7               8                       9
    cString = "SELECT FILE1_10.MODEL , FILE1_10.mosm,FILE1_10.fact,FILE1_10.modelfact0,FILE1_10.desca ,SUM(Quant)  , FILE1_10.mosm,FILE1_10.fact,FILE1_10.modelfact0 , ' ' AS NEWMPDEL " & _
          " FROM " & Purchasefrm.cFile & " LEFT JOIN FILE1_10 ON " & Purchasefrm.cFile & ".ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(Purchasefrm.xdoc_no.text) & " GROUP BY FILE1_10.MOSM,FILE1_10.FACT,FILE1_10.MODELFACT0 , FILE1_10.DESCA , FILE1_10.MODEL order by FILE1_10.MOSM,FILE1_10.FACT,FILE1_10.MODELFACT0 "
    data1.RecordSource = cString
    data1.Refresh
End With
Fixgrd
End Sub
Private Sub Fixgrd()
With grid1
'                   0          1        2            3           4               5          6             7             8              9
'.FormatString = "„ÊœÌ·|" & "„Ê”„|" & "„’‰⁄|" & "—ﬁ„ „ÊœÌ·|" & "≈”„ «·„ÊœÌ·|" & " ﬂ„Ì…|" & "„Ê”„|" & " „’‰⁄|" & "—ﬁ„ „ÊœÌ·|" & "„ÊœÌ· ÃœÌœ|"

.TextMatrix(1, 0) = ""
.TextMatrix(1, 1) = "„Ê”„"
.TextMatrix(1, 2) = "„’‰⁄"
.TextMatrix(1, 3) = "„ÊœÌ· «·„’‰⁄"
.TextMatrix(1, 4) = "≈”„ «·„ÊœÌ·"
.TextMatrix(1, 5) = "«·ﬂ„Ì…"
.TextMatrix(1, 6) = "„Ê”„"
.TextMatrix(1, 7) = "„’‰⁄"
.TextMatrix(1, 8) = "„ÊœÌ· «·„’‰⁄"
.TextMatrix(1, 9) = "MODEL"


.TextMatrix(0, 0) = "ﬁ»· «· ⁄œÌ·"
.TextMatrix(0, 1) = "ﬁ»· «· ⁄œÌ·"
.TextMatrix(0, 2) = "ﬁ»· «· ⁄œÌ·"
.TextMatrix(0, 3) = "ﬁ»· «· ⁄œÌ·"
.TextMatrix(0, 4) = "ﬁ»· «· ⁄œÌ·"
.TextMatrix(0, 5) = "ﬁ»· «· ⁄œÌ·"

.TextMatrix(0, 5) = "»⁄œ «· ⁄œÌ·"
.TextMatrix(0, 6) = "»⁄œ «· ⁄œÌ·"
.TextMatrix(0, 7) = "»⁄œ «· ⁄œÌ·"
.TextMatrix(0, 8) = "»⁄œ «· ⁄œÌ·"
.TextMatrix(0, 9) = "»⁄œ «· ⁄œÌ·"

.WordWrap = True
.Cols = 10
.RowHeight(0) = 400
.RowHeight(1) = 400

.MergeCells = flexMergeFree
.MergeRow(0) = True
.ColWidth(0) = 0
.ColWidth(1) = 700
.ColWidth(2) = 2200
.ColWidth(3) = 1300
.ColWidth(4) = 3500
.ColWidth(5) = 1000
.ColWidth(6) = 700
.ColWidth(7) = 2200
.ColWidth(8) = 1300
.ColWidth(9) = 1000
.ColComboList(9) = "..."

'.ColHidden(7) = (Purchasefrm.xsupp.Value = 0)
.ColComboList(1) = StrList("SELECT MOSM, MOSM FROM MOSM", con)
.ColComboList(2) = StrList("SELECT CODE , DESCA FROM FACT ", con)
.ColComboList(6) = StrList("SELECT MOSM, MOSM FROM MOSM", con)
.ColComboList(7) = StrList("SELECT CODE , DESCA FROM FACT ", con)
.SubtotalPosition = flexSTBelow
.Subtotal flexSTSum, -1, 5, "#0", vbYellow, vbRed, True, "≈Ã„«·Ï "
.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4

End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub

Private Sub GRID1_Click()
With grid1
    If .Col = 6 Or .Col = 7 Or .Col = 8 Or .Col = 9 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
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
            cModel = .TextMatrix(nRow, 7) & .TextMatrix(nRow, 6) & RetZero(.TextMatrix(nRow, 8), 10) & Purchasefrm.xCode.text
        Else
            cModel = .TextMatrix(nRow, 7) & .TextMatrix(nRow, 6) & RetZero(.TextMatrix(nRow, 8), 10)
        End If
        If cModel <> .TextMatrix(nRow, 0) Then
            If GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(cModel), con) = "" Then
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
Private Sub xMosm_LostFocus()
With grid1
    For nRow = 2 To .Rows - 1
        .TextMatrix(nRow, 6) = xMosm.BoundText
    Next nRow
End With
End Sub
Private Sub xFACT_LostFocus()
With grid1
    For nRow = 2 To .Rows - 1
        .TextMatrix(nRow, 7) = xFact.BoundText
    Next nRow
End With
End Sub

