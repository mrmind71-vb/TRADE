VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form SalesDoc_COPY 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   7590
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   15405
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   11880
      Picture         =   "SalesDoc_COPY.frx":0000
      ScaleHeight     =   2445
      ScaleWidth      =   3210
      TabIndex        =   3
      Top             =   45
      Width           =   3210
      Begin ImgXCtrl6.ImgXCtrl imgx1 
         Height          =   2175
         Left            =   45
         TabIndex        =   4
         Top             =   45
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   3836
         BackColor       =   16777215
         BorderStyle     =   0
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "amr123"
         LicenseRegCode  =   "íß“µª∫≥ºΩ∫æ“±æß´¥Ω∏•OANH-FAZOHIZB-EFTP6gI"
      End
   End
   Begin VB.CommandButton CMD_EXIT 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1890
      Width           =   2400
   End
   Begin VB.CommandButton cmd_retall 
      Caption         =   "„— Ã⁄ ··»Ê‰ »«·ﬂ«„·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1350
      Width           =   2400
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   45
      Top             =   -630
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
      Left            =   45
      Top             =   -1035
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
      Height          =   4965
      Left            =   90
      TabIndex        =   0
      Top             =   2430
      Width           =   15000
      _cx             =   26458
      _cy             =   8758
      _ConvInfo       =   1
      Appearance      =   0
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
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
      ScrollBars      =   2
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
   Begin VB.Label xNOTE_DISC 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5310
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1350
      Width           =   5145
   End
   Begin VB.Label xDOC_DISC 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2655
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1350
      Width           =   2535
   End
   Begin VB.Label xdate 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1890
      Width           =   2535
   End
   Begin VB.Label xbranch 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5310
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1890
      Width           =   2535
   End
   Begin VB.Label xdoc_no 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2655
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1890
      Width           =   2535
   End
End
Attribute VB_Name = "SalesDoc_COPY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cInv_No As String, nDiscInvoice As Double, cNote_Disc As String, cMan As String
Dim lOtherBranch As Boolean, lSalesDisc2 As Boolean
Dim lConServer As Boolean
Public sDoc_no As String
Dim con As New ADODB.Connection
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub
Private Sub cmd_retall_Click()
If xdoc_no.Caption <> "" Then
    If GetDesca("SELECT DOC_NO  FROM FILE6_20H WHERE ISRET = 1 AND DOC_NO = " & MyParn(sDoc_no), con) = "" Then
        MsgBox " „ ⁄„· „— Ã⁄ „‰ ﬁ»· ··›« Ê—…"
        Unload Me
    End If
End If
With grid1
    If lOtherBranch Then
        If cBranch > "60" Then
            cMan = GetDesca("SELECT CODE FROM FILE6_25 ORDER BY CODE", con)
        Else
            cMan = "0001"
        End If
        aRet = aGetDesca("SELECT INV_NO , DISCOUNT , NOTE_DISC , MAN FROM QFILE6_20H WHERE DOC_NO = " & MyParn(sDoc_no), conShop)
        cInv_No = aRet(1) & ""
        nDiscInvoice = Val(aRet(2) & "") * -1
        cNote_Disc = aRet(3) & ""
    Else
        aRet = aGetDesca("SELECT INV_NO , DISCOUNT , NOTE_DISC , MAN FROM QFILE6_20H WHERE DOC_NO = " & MyParn(sDoc_no), con)
        If UBound(aRet) > 0 Then
            cInv_No = aRet(1) & ""
            nDiscInvoice = Val(aRet(2) & "") * -1
            cNote_Disc = aRet(3) & ""
            xMan = aRet(4) & ""
        End If
    End If
    SalesFrm.XSALES_RET.Caption = sDoc_no
    SalesFrm.xMan.BoundText = cMan
    If cInv_No <> "" Then
        SalesFrm.XISRETS.Value = 1
        SalesFrm.xinv_no.text = cInv_No
        SalesFrm.xDiscount.text = nDiscInvoice
        SalesFrm.xNOTE_DISC.Caption = cNote_Disc
    End If
    
    If SalesFrm.grid1.Rows > 2 Then
        MsgBox "ÌÊÃœ «’‰«› „”Ã·… ⁄·Ï «·»Ê‰"
        Exit Sub
    End If
'   If cInv_No <> "" Then Inform " „— Ã⁄ ⁄—÷  " & cNote_Disc
    For nRow = 2 To .Rows - 1
        If Val(.TextMatrix(nRow, 6)) <> 0 Then
        If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & Val(.TextMatrix(nRow, 6)), con) = "" Then
            sendItems Val(.TextMatrix(nRow, 6))
        End If
        If IsNumeric(SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 1)) Then
            SalesFrm.grid1.AddItem ""
        End If
        .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = vbRed
        
        SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 10) = Val(.TextMatrix(nRow, 8)) * -1
        SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 1) = .TextMatrix(nRow, 6)
        SalesFrm.GrdDesc .TextMatrix(nRow, 6), SalesFrm.grid1.Rows - 1
        If lOtherBranch Then
            If cBranch > "60" Then
                SalesFrm.xMan.BoundText = "00"
            Else
                SalesFrm.xMan.BoundText = "0001"
            End If
        Else
            If SalesFrm.xMan.BoundText = "" Then SalesFrm.xMan.BoundText = .TextMatrix(nRow, 3)
        End If
        SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 16) = .TextMatrix(nRow, 3)
        SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 15) = .TextMatrix(nRow, 3)
        If cInv_No = "" Then
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 11) = Val(.TextMatrix(nRow, 9))
        Else
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 11) = Val(.TextMatrix(nRow, 10))
        End If
'       SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 12) = Val(.TextMatrix(nRow, 10))
        SalesFrm.grid1_AfterEdit SalesFrm.grid1.Rows - 1, 1
        End If
    Next nRow
    Unload Me
End With
End Sub
Private Sub Form_Load()
    openCon con
    If sDoc_no = "" Then Exit Sub
'    If Not lIsBranchStore Then
'        If Len(sDoc_no) = 11 Then
'            If Mid(sDoc_no, 7, 1) = Val(cBranch) Then
'                lOtherBranch = False
'            Else
'                lOtherBranch = True
'            End If
'        Else
'            If Mid(sDoc_no, 7, 2) = Val(cBranch) Then
'                lOtherBranch = False
'            Else
'                lOtherBranch = True
'            End If
'        End If
'    Else
'        lOtherBranch = False
'    End If
    If myload Then
        Fixgrd
    End If
End Sub
Private Sub Fixgrd()
With grid1
.Cols = 12
.TextMatrix(0, 0) = "—ﬁ„ „” ‰œ"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "«·Êﬁ "
.TextMatrix(0, 3) = "«·»«∆⁄"
.TextMatrix(0, 4) = "«·„ÊœÌ·"
.TextMatrix(0, 5) = "«·„’‰⁄"
.TextMatrix(0, 6) = "»«—ﬂÊœ"
.TextMatrix(0, 7) = "«·’‰›"
.TextMatrix(0, 8) = "⁄œœ"
.TextMatrix(0, 9) = "«·”⁄—"
'.TextMatrix(0, 10) = "«·Œ’„"
.TextMatrix(0, 11) = "«·≈Ã„«·Ï"
.MergeCells = flexMergeFree
.MergeCol(0) = True
.MergeCol(1) = True
.ColComboList(3) = StrList("SELECT CODE , DESCA FROM FILE6_25")
.ColWidth(0) = 0
.ColWidth(1) = 1300
.ColWidth(2) = 1000
.ColWidth(3) = 1300
.ColWidth(4) = 1200
.ColWidth(5) = 1500
.ColWidth(6) = 1000
.ColWidth(7) = 2400
.ColWidth(8) = 700
.ColWidth(9) = 700
.ColWidth(10) = 0
.ColWidth(11) = 1000
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Function myload() As Boolean
Dim cWhere As String, cString As String, aReT2 As Variant
Dim nRateDisc As Double
    If lOtherBranch Then
        Inform " «·„»Ì⁄«  „‰ ›—⁄ «Œ— "
        strConShop = LoadConStringshop
        If openConShop(conShop) <> "ok" Then
            MsgBox "·« ÌÊÃœ « ’«· »«·”Ì—›—"
        Else
            lConServer = True
            Inform "  „ «·≈ ’«· »«·”Ì—›— "
        End If
    End If

Dim loctable As New ADODB.Recordset
If lOtherBranch Then
    If GetDesca("SELECT doc_no FROM QFILE6_20 WHERE DOC_NO = " & MyParn(sDoc_no), conShop) <> "" Then
'                                   1       2               3      4       5
        aRet = aGetDesca("SELECT INV_NO , DISCOUNT , NOTE_DISC , MAN ,  DOC_DISC FROM QFILE6_20H WHERE DOC_NO = " & MyParn(sDoc_no), conShop)
        If UBound(aRet) = 0 Then Exit Function
        cInv_No = aRet(1) & ""
        nDiscInvoice = Val(aRet(2) & "")
        cNote_Disc = aRet(3) & ""
        xNOTE_DISC.Caption = aRet(3) & ""
        xNOTE_DISC.Caption = cNote_Disc
        If cInv_No = "" Then
            xDOC_DISC.Caption = aRet(5) & ""
        Else
            xDOC_DISC.Caption = cInv_No & ""
        End If
        aReT2 = aGetDesca("SELECT QTY1 , QTY2 FROM FILE0_90H WHERE DOC_NO = " & MyParn(xDOC_DISC.Caption), con)
        If UBound(aReT2) > 0 Then
            If Val(aReT2(1) & "") > 1 Or Val(aReT2(2) & "") <> 0 Then SalesFrm.lDiscModelRet = True
        End If
        nRateDisc = Val(GetDesca("SELECT Rate FROM QFILE6_20 WHERE DOC_NO = " & MyParn(sDoc_no), conShop) & "")

'        lSalesDisc2 = GetBoolean("SELECT ISDISC2_1 FROM QFILE6_20 WHERE DOC_NO = " & MyParn(sDoc_no), conShop)
'        If XDOC_DISC.Caption <> "" Then
'            Inform_OK " »Ê‰ „»Ì⁄«  ⁄—÷  ÌÃ» «—Ã⁄«⁄ «·»Ê‰ »«·ﬂ«„· "
'        End If
        
        If xDOC_DISC.Caption <> "" Then Inform_OK " »Ê‰ „»Ì⁄«  Œ’„ ⁄—Ê÷ " & cNote_Disc
        
        If Len(sDoc_no) = 12 And Mid(sDoc_no, 7, 3) > "600" Then
            cString = "SELECT SALESDTL_FR.* , file1_10.modelfact " & _
                      " FROM SALESDTL_FR left join file1_10 on file1_10.item = SALESDTL_FR.item WHERE DOC_NO = " & MyParn(sDoc_no)
            cString = cString & " ORDER BY ITEM "
        Else
            cString = "SELECT SALESDTL.* , file1_10.modelfact " & _
                      " FROM SALESDTL left join file1_10 on file1_10.item = SALESDTL.item WHERE DOC_NO = " & MyParn(sDoc_no)
            cString = cString & " ORDER BY ITEM "
        End If
        loctable.Open cString, conShop, adOpenKeyset, adLockReadOnly, adCmdText
    Else
        MsgBox "—ﬁ„ »Ê‰ €Ì— „ÊÃÊœ"
        myload = False
        Exit Function
    End If
Else
'                               1       2           3           4       5
    aRet = aGetDesca("SELECT INV_NO , DISCOUNT , NOTE_DISC , MAN , DOC_DISC FROM FILE6_20H WHERE DOC_NO = " & MyParn(sDoc_no), con)
    If UBound(aRet) = 0 Then Exit Function
    nDiscInvoice = Val(aRet(2) & "")
    cInv_No = aRet(1) & ""
    cNote_Disc = aRet(3) & ""
    cMan = aRet(4) & ""
    If cInv_No = "" Then
        xDOC_DISC.Caption = aRet(5) & ""
    Else
        xDOC_DISC.Caption = cInv_No & ""
    End If
    xNOTE_DISC.Caption = cNote_Disc
    aReT2 = aGetDesca("SELECT QTY1 , QTY2 FROM FILE0_90H WHERE DOC_NO = " & MyParn(xDOC_DISC.Caption), con)
    If UBound(aReT2) > 0 Then
        If Val(aReT2(1) & "") > 1 Or Val(aReT2(2) & "") <> 0 Then SalesFrm.lDiscModelRet = True
    End If
    
    nRateDisc = Val(GetDesca("SELECT    Rate FROM FILE6_20H WHERE DOC_NO = " & MyParn(sDoc_no)) & "")
'    lSalesDisc2 = GetBoolean("SELECT    ISDISC2_1 FROM FILE6_20 WHERE DOC_NO = " & MyParn(sDoc_no))
'    If XDOC_DISC.Caption <> "" And cInv_No = "" Then
'        Inform_OK " »Ê‰ „»Ì⁄«  ⁄—÷ ÌÃ» «—Ã⁄«⁄ «·»Ê‰ »«·ﬂ«„· "
'    End If
    If xDOC_DISC.Caption <> "" Then Inform_OK " »Ê‰ „»Ì⁄«  Œ’„ ⁄—Ê÷ " & cNote_Disc
    cString = "SELECT SALESDTL.* , file1_10.modelfact " & _
              " FROM SALESDTL left join file1_10 on file1_10.item = SALESDTL.item WHERE DOC_NO = " & MyParn(sDoc_no)
    cString = cString & " ORDER BY ITEM "
    loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
End If


aRet = aGetDesca("select doc_no , store , date  from file6_20h where  SALES_RET = " & MyParn(sDoc_no), con)
If UBound(aRet) > 0 Then
    xdate.Caption = Format(aRet(3), "dd-mm-yyyy")
    xdoc_no.Caption = aRet(1)
    xbranch.Caption = GetDesca("select desca from  STORE_BR where code = " & MyParn(aRet(2)))
    If GetDesca("SELECT DOC_NO  FROM FILE6_20H WHERE ISRET = 1 AND DOC_NO = " & MyParn(sDoc_no)) = "" Then
        MsgBox "  „ ⁄„· „— Ã⁄ ··»Ê‰ „‰ ﬁ»·  "
        cmd_retall.Enabled = False
    End If
Else
    If lConServer Then
        aRet = aGetDesca("select doc_no , store , date from qfile6_20 where  SALES_RET = " & MyParn(sDoc_no), conShop)
        If UBound(aRet) > 0 And SalesFrm.XISRET.Value = 0 Then
            xdate.Caption = Format(aRet(3), "dd-mm-yyyy")
            xdoc_no.Caption = aRet(1)
            xbranch.Caption = GetDesca("select desca from  STORE_BR where code = " & MyParn(aRet(2)), conShop)
            If GetDesca("SELECT DOC_NO  FROM FILE6_20H WHERE ISRET = 1 AND DOC_NO = " & MyParn(sDoc_no), con) = "" Then
                cmd_retall.Enabled = False
                MsgBox "  „ ⁄„· „— Ã⁄ ··»Ê‰ „‰ ﬁ»·  "
            End If
        End If
    End If
End If
With grid1
.Rows = 1
.Cols = 12
.ColComboList(3) = StrList("SELECT CODE , DESCA FROM FILE6_25")
If loctable.RecordCount <> 0 Then loctable.MoveFirst
Do Until loctable.EOF
    If loctable!Flag = 0 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "dd-mm-yyyy")
        .TextMatrix(.Rows - 1, 0) = Mid(loctable!doc_no, 7, 4)
        .TextMatrix(.Rows - 1, 2) = Format(loctable!Time, "HH:NN")
        .TextMatrix(.Rows - 1, 3) = "000"
        .TextMatrix(.Rows - 1, 4) = DelZero(loctable!modelfact & "")
        .TextMatrix(.Rows - 1, 5) = loctable!fact_Desca & ""
        .TextMatrix(.Rows - 1, 6) = loctable!Item & ""
        .TextMatrix(.Rows - 1, 7) = loctable!DESCA & ""
        .TextMatrix(.Rows - 1, 8) = loctable!Quant
        .TextMatrix(.Rows - 1, 9) = Round(loctable!price * ((100 - nRateDisc) / 100), 2)
        .TextMatrix(.Rows - 1, 10) = loctable!price & ""
        .TextMatrix(.Rows - 1, 11) = Round(loctable!TOTAL * ((100 - nRateDisc) / 100), 2)
    ElseIf loctable!Flag = 1 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "dd-mm-yyyy")
        .TextMatrix(.Rows - 1, 0) = DelZero(Mid(loctable!doc_no, 7, 4))
        For i = 0 To 9
            .TextMatrix(.Rows - 1, i) = "«·Œ’„"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 11) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
    ElseIf loctable!Flag = 3 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "dd-mm-yyyy")
        .TextMatrix(.Rows - 1, 0) = DelZero(Mid(loctable!doc_no, 7, 4))
        For i = 0 To 9
            .TextMatrix(.Rows - 1, i) = "«·«Ã„«·Ì"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 11) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
    ElseIf loctable!Flag = 5 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "dd-mm-yyyy")
        .TextMatrix(.Rows - 1, 0) = DelZero(Mid(loctable!doc_no, 7, 4))
        For i = 0 To 9
            .TextMatrix(.Rows - 1, i) = "”œ«œ ›Ì“«"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 11) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HFF&
    End If
    loctable.MoveNext
Loop
End With
Fixgrd
myload = True
End Function
Private Sub grid1_DblClick()
'    cString = "SELECT FILE6_20.ITEM, FILE1_10.MOSM, FILE1_10.FACT, FILE1_10.SUPP, FILE1_10.MODELFACT0, FILE1_10.DESCA, FILE1_10.SCAL , FILE1_10.COLOR , FILE6_20.PRICE_C, file6_20.Quant, file6_20.Price, file6_20.discount ,FILE6_20.TOTAL ,FILE6_20.PRICE_C2, FILE1_10.price2 ,FILE6_20.man ,FILE1_10.MODELNO ,ID "
    
    If xdoc_no.Caption <> "" Then
        If GetDesca("SELECT DOC_NO  FROM FILE6_20H WHERE ISRET = 1 AND DOC_NO = " & MyParn(sDoc_no)) = "" Then
            MsgBox " „ ⁄„· „— Ã⁄ „‰ ﬁ»· ··›« Ê—…"
            Unload Me
        End If
    End If
    
    If xdoc_no.Caption <> "" Then Exit Sub
    If SalesFrm.lDiscModelRet Then
        MsgBox " €Ì— „”„ÊÕ »„— Ã⁄ ’‰› Ê«Õœ ›Ï „— Ã⁄ «·⁄—÷ "
        Exit Sub
    End If
    With grid1
    If lOtherBranch Then
        If cBranch > "60" Then
            cMan = "00"
        Else
            cMan = "0001"
        End If
    Else
        cMan = GetDesca("SELECT MAN FROM FILE6_20H WHERE DOC_NO = " & MyParn(sDoc_no))
    End If
        
    SalesFrm.XSALES_RET.Caption = sDoc_no
    SalesFrm.xMan.BoundText = cMan
    If IsNumeric(SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 1)) Then
        SalesFrm.grid1.AddItem ""
    End If
    .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = vbRed
    SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 10) = Val(.TextMatrix(.Row, 8)) * -1
    SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 1) = .TextMatrix(.Row, 6)
    SalesFrm.GrdDesc .TextMatrix(.Row, 6), SalesFrm.grid1.Rows - 1
    SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 16) = .TextMatrix(.Row, 3)
    SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 15) = .TextMatrix(.Row, 3)
    SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 11) = Val(.TextMatrix(.Row, 9))
'    SalesFrm.Grid1.TextMatrix(SalesFrm.Grid1.Rows - 1, 12) = Val(.TextMatrix(.Row, 10))
    SalesFrm.XSALES_RET.Caption = sDoc_no
    SalesFrm.grid1_AfterEdit SalesFrm.grid1.Rows - 1, 1
'   Unload Me
End With
End Sub
Private Function sendItems(pItem) As Long
Dim cFile As String
cFile = "FILE1_10"
cString = "SELECT * FROM FILE1_10 WHERE  ITEM  = " & Val(pItem)
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    con.Execute " DELETE FROM FILE1_10 where item = " & loctable!Item
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!MODEL))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable!MODELNO))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "FACT", addstring(loctable!Fact))
    aInsert = AddFlag(aInsert, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "MODELFACT", addstring(loctable!modelfact))
    aInsert = AddFlag(aInsert, "MODELFACT0", addstring(loctable!modelfact0))
    aInsert = AddFlag(aInsert, "SUPP", addstring(loctable!SUPP))
    aInsert = AddFlag(aInsert, "OKAZ", Val(loctable!okaz & ""))
    aInsert = AddFlag(aInsert, "OKAZ_2", Val(loctable!okaz_2 & ""))
    aInsert = AddFlag(aInsert, "RATE", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "[SECTION]", addvalue(loctable!Section))
    aInsert = AddFlag(aInsert, "[ITEM]", addvalue(loctable!Item))
    aInsert = AddFlag(aInsert, "[COST]", Val(loctable!cost & ""))
    If cBranch > "60" Then
        aInsert = AddFlag(aInsert, "[COSTITEM]", Val(loctable!costITEM_FR & ""))
    Else
        aInsert = AddFlag(aInsert, "[COSTITEM]", Val(loctable!costITEM & ""))
    End If
    aInsert = AddFlag(aInsert, "[COST2]", Val(loctable!cost2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE]", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "[PRICE2]", Val(loctable!PRICE2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE_2]", Val(loctable!PRICE_2 & ""))
    aInsert = AddFlag(aInsert, "[SCAL]", addstring(loctable!scal))
    aInsert = AddFlag(aInsert, "[C_SCAL]", addvalue(loctable!C_SCAL))
    aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable!Color))
    aInsert = AddFlag(aInsert, "[C_COLOR]", addvalue(loctable!c_Color))
    aInsert = AddFlag(aInsert, "[BARCODE]", addstring(loctable!BARCODE))
    aInsert = AddFlag(aInsert, "[BARCODE2]", addstring(loctable!BARCODE2))
    aInsert = AddFlag(aInsert, "[BARCODE13]", addstring(loctable!BARCODE13))
    aInsert = AddFlag(aInsert, "[ISOKAZITEM]", IIf(loctable!ISOKAZITEM, 1, 0))
    aInsert = AddFlag(aInsert, "[MOSM2]", addstring(loctable!MOSM2))
    aInsert = AddFlag(aInsert, "[ISNOITEM]", IIf(loctable!ISNOITEM, 1, 0))
    aInsert = AddFlag(aInsert, "[ISNODEM]", IIf(loctable!ISNODEM, 1, 0))
    con.Execute addInsert(aInsert, "FILE1_10"), nAffect
    Inform "  „ ≈÷«›… »«—ﬂÊœ " & loctable!Item & "  " & loctable!DESCA
    loctable.MoveNext
Loop
End Function
Private Sub myLoadPicture(pModel As String)
On Error Resume Next
    If Not lServerPict Then Exit Sub
    imgx1.Images.Clear
    If pModel = "" Then Exit Sub
    If conPict Is Nothing Then Exit Sub
    Dim loctable As Recordset
    Dim bytes() As Byte, nSize As Long
    imgx1.Images.Clear
    Set loctable = myRecordSet("select PICT,size from PICT where MODELNO = " & MyParn(pModel), conPict)
    If loctable.EOF Then Exit Sub
    nSize = CLng(loctable("size").Value)
    bytes = loctable("PICT").GetChunk(nSize)
    imgx1.Import.FromMemoryFile bytes, ixmfJPG
    Err.Clear
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
    If OldRow <> NewRow Then
        myLoadPicture GetDesca("SELECT MODELNO FROM FILE1_10 WHERE ITEM = " & Val(.TextMatrix(NewRow, 6)))
    End If
End With
End Sub




