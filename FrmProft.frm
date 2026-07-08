VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmProft 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ﬁ«∆„… «·œŒ·"
   ClientHeight    =   10230
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   9030
   Begin ComctlLib.ProgressBar prog1 
      Height          =   285
      Left            =   180
      TabIndex        =   8
      Top             =   1125
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   503
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmd_print 
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   945
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   675
      Width           =   1455
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "⁄—÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   945
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   180
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1410
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6315
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   562
         Width           =   2490
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   2490
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   1575
         TabIndex        =   9
         Top             =   990
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   135
         TabIndex        =   11
         Top             =   180
         Width           =   2190
         _ExtentX        =   3863
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
         Picture         =   "FrmProft.frx":0000
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1035
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "«·Ï :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   390
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ :"
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
         Left            =   5175
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   765
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   225
      Top             =   135
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7290
      Left            =   135
      TabIndex        =   7
      Top             =   1440
      Width           =   8700
      _cx             =   15346
      _cy             =   12859
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      SelectionMode   =   0
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   180
      Top             =   765
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
End
Attribute VB_Name = "frmProft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_OK_Click()
If Not (IsDate(xdate1.text) And IsDate(xDate2.text)) Then Exit Sub
Dim nFBal As Double, nCountC As Double
Dim nCBal As Double, nTcomp As Double
Dim nRate As Double
Dim nTAllSales As Double
Dim nTPurch As Double
Dim nNetSales As Double
Dim nTSales As Double
Dim nTSales2 As Double
Dim nTCharge1 As Double
Dim nTACCCharge1 As Double
Dim nProft As Double
Dim nTPart1 As Double
Dim nTPart2 As Double
Dim nTCharge2 As Double
Dim nTAccCharge2 As Double
Dim nTAcc1 As Double
Dim nTAcc2 As Double

Dim ItemTable As New ADODB.Recordset
Dim Charge1Table As New ADODB.Recordset, AddestTable As New ADODB.Recordset
Dim Charge2Table As New ADODB.Recordset, nTDICSUP As Double, nTDICCUST As Double
Dim PartTable As New ADODB.Recordset

Dim Acc1Table As New ADODB.Recordset
Dim Acc2Table As New ADODB.Recordset

ItemTable.Open "SELECT * FROM FILE1_10 ", con, adOpenKeyset, adLockReadOnly, adCmdText
Dim nRNetSales As Double
Dim nRTSales As Double

Dim nTCharge As Double
Dim nBal As Double

cStr1 = "SELECT SUM(T_TOTAL-DISCOUNT+TAX) FROM T_SALESDOC WHERE DOC_NO IS NOT NULL "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
If XSTORE.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(XSTORE.BoundText)
nTAllSales = Val(GetDesca(cStr1, con) & "")

cStr1 = "select sum( FILE7_20.QUANT * FILE7_20.COST_INV ) from FILE7_20 INNER JOIN FILE7_20H ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO WHERE FILE7_20H.DOC_NO IS NOT NULL "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
If XSTORE.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(XSTORE.BoundText)
nTPurch = Val(GetDesca(cStr1, con) & "")

cStr1 = "select sum( FILE7_10.QUANT * FILE7_10.COST_INV ) from FILE7_10 INNER JOIN FILE7_10H ON FILE7_10H.DOC_NO = FILE7_10.DOC_NO WHERE FILE7_10H.DOC_NO IS NOT NULL "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
If XSTORE.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(XSTORE.BoundText)
nTPurch = nTPurch - Val(GetDesca(cStr1, con) & "")

If XSTORE.BoundText <> "" Then
    cStr1 = "select sum( FILE1_60.COST * FILE1_60.QUANT ) from FILE1_60 INNER JOIN FILE1_60H ON FILE1_60H.DOC_NO = FILE1_60.DOC_NO WHERE STORE2 = " & MyParn(XSTORE.BoundText)
    cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
    nTTRANS = Val(GetDesca(cStr1, con) & "")

    cStr1 = "select sum( FILE1_60.COST * FILE1_60.QUANT ) from FILE1_60 INNER JOIN FILE1_60H ON FILE1_60H.DOC_NO = FILE1_60.DOC_NO WHERE STORE1 = " & MyParn(XSTORE.BoundText)
    cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
    nTTRANS = nTTRANS - Val(GetDesca(cStr1, con) & "")
    nTPurch = nTPurch + nTTRANS
End If

cStr1 = "select sum( T_COST ) from T_SALESDOC WHERE DOC_NO IS NOT NULL "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
If XSTORE.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(XSTORE.BoundText)
nTCOST = Val(GetDesca(cStr1, con) & "")

cStr1 = "select sum( file0_10.cost * differ ) from file0_10 inner join file0_10h on file0_10h.doc_no = file0_10.doc_no where file0_10h.DOC_NO IS NOT NULL "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
If XSTORE.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(XSTORE.BoundText)
nTcomp = Val(GetDesca(cStr1, con) & "")

cStr1 = " SELECT FILE8_52.DESCA AS DESCA , SUM(VALUE) AS TCharge FROM  ACC_CHARGE INNER JOIN FILE8_51 ON ACC_CHARGE.CHARGE = FILE8_51.CODE   inner join file8_52 on file8_52.code = file8_51.MAINGROUP     where VALUE <> 0 AND File8_51.ISasset = 0 "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
If XSTORE.BoundText <> "" Then
    pBra = GetDesca("SELECT BRANCH FROM FILE0_40 WHERE code = " & MyParn(XSTORE.BoundText), con)
    cStr1 = cStr1 & " AND branch = " & MyParn(pBra)
End If
cStr1 = cStr1 & " GROUP BY FILE8_52.DESCA  "
Charge1Table.Open cStr1, con, adOpenKeyset, adLockReadOnly, adCmdText

cStr1 = " SELECT FILE8_51.DESCA AS DESCA , SUM(VALUE) AS TCharge FROM  ACC_CHARGE INNER JOIN FILE8_51 ON ACC_CHARGE.CHARGE = FILE8_51.CODE   where VALUE <> 0 AND File8_51.ISasset = 1 "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
If XSTORE.BoundText <> "" Then
    pBra = GetDesca("SELECT BRANCH FROM FILE0_40 WHERE code = " & MyParn(XSTORE.BoundText), con)
    cStr1 = cStr1 & " AND branch = " & MyParn(pBra)
End If
cStr1 = cStr1 & " GROUP BY FILE8_51.DESCA  "
AddestTable.Open cStr1, con, adOpenKeyset, adLockReadOnly, adCmdText


cStr1 = " SELECT   FILE8_62.DESCA AS DESCA , SUM(VALUE) AS TCharge FROM  ACC_INCOM INNER JOIN FILE8_61 ON ACC_INCOM.CHARGE = FILE8_61.CODE   INNER JOIN FILE8_62 on file8_62.code = file8_61.MAINGROUP   where VALUE <> 0 "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
If XSTORE.BoundText <> "" Then
    pBra = GetDesca("SELECT BRANCH FROM FILE0_40 WHERE code = " & MyParn(XSTORE.BoundText), con)
    cStr1 = cStr1 & " AND branch = " & MyParn(pBra)
End If
cStr1 = cStr1 & " GROUP BY FILE8_62.DESCA "
Charge2Table.Open cStr1, con, adOpenKeyset, adLockReadOnly, adCmdText

If XSTORE.BoundText = "" Then
    cStr1 = " SELECT FILE8_71.DESCA AS DESCA , SUM(VALUE_M-VALUE_P) AS TVALUE FROM QFILE8_70 INNER JOIN FILE8_71 ON QFILE8_70.CODE = FILE8_71.CODE     where FILE8_71.DESCA IS NOT NULL "
    cStr1 = cStr1 & " AND DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
    cStr1 = cStr1 & " GROUP BY FILE8_71.DESCA "
    PartTable.Open cStr1, con, adOpenKeyset, adLockReadOnly, adCmdText
End If

cStr1 = " SELECT   SUM(VALUE) FROM            Q_CASH0 where VALUE <> 0 "
cStr1 = cStr1 & " and DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
If XSTORE.BoundText <> "" Then
    cStr1 = cStr1 & " AND STORE = " & MyParn(XSTORE.BoundText)
End If
nTDICSUP = Val(GetDesca(cStr1, con) & "")

If cBranch = "00" Then
    cStr1 = " SELECT   SUM(VALUE) FROM  Q_CASH11 where VALUE <> 0 "
    cStr1 = cStr1 & " and DATE >= " & DateSq(xdate1.text) & " AND DATE <= " & DateSq(xDate2.text)
    If XSTORE.BoundText <> "" Then
        cStr1 = cStr1 & " AND STORE = " & MyParn(XSTORE.BoundText)
    End If
    nTDICCUST = Val(GetDesca(cStr1, con) & "")
End If
nFBal = TCostItem1
nCBal = TCostItem2

With grid1
    .Cols = 2
    .ColWidth(0) = 6000
    .ColWidth(1) = 2000
    .FixedRows = 1
    .TextMatrix(0, 0) = "»Ì«‰"
    .TextMatrix(0, 1) = "«·ﬁÌ„…"
    
    .TextMatrix(1, 0) = "„»Ì⁄«  "
'   .TextMatrix(2, 0) = "„»Ì⁄«  Ã„·…"

    .TextMatrix(4, 0) = "„—œÊœ „»Ì⁄«  "
    .TextMatrix(5, 0) = "Œ’„ „—œÊœ „»Ì⁄«  "
    
    .RowHidden(1) = True
    .RowHidden(2) = True
    .RowHidden(3) = True
    .RowHidden(4) = True
    .RowHidden(5) = True


   .RowHidden(4) = True
   .RowHidden(5) = True
   .RowHidden(6) = True
    
    .TextMatrix(7, 0) = "’«›Ï „»Ì⁄«  "
    .TextMatrix(9, 0) = " ﬂ·›… »÷«⁄… √Ê· "
    .TextMatrix(10, 0) = "’«›Ï „‘ —Ì« "
    .TextMatrix(11, 0) = " ﬂ·›… »÷«⁄… √Œ— ··„œÂ"
    .TextMatrix(12, 0) = " ﬂ·›… „»Ì⁄«  "
    .TextMatrix(13, 0) = "—»Õ „»Ì⁄«  "
    .TextMatrix(14, 0) = "⁄Ã“ Ã—œ"
    .TextMatrix(15, 0) = "›—ﬁ  ﬂ·›… »÷«⁄…"
    .TextMatrix(16, 0) = " ﬂ·›… «·»÷«⁄… «·„»«⁄…"
    
    
    .Rows = 17
    nTSales = nTAllSales
    nTSales2 = 0
    .TextMatrix(1, 1) = Format(nTSales, "#,###,##0")
'   .TextMatrix(2, 1) = Format(nTSales2, "#,###,##0")

'    .TextMatrix(4, 1) = Format(nRTSales, "#,###,##0")
'    .TextMatrix(5, 1) = Format(nRTSales - nRNetSales, "#,###,##0")

    .TextMatrix(7, 1) = Format(nTSales, "#,###,##0")
    
    .TextMatrix(9, 1) = Format(nFBal, "#,###,##0")
    .TextMatrix(10, 1) = Format(nTPurch, "#,###,##0")
    .TextMatrix(11, 1) = Format(nCBal, "#,###,##0")
    
    .TextMatrix(12, 1) = Format(nTCOST, "#,###,##0")
    
    .TextMatrix(13, 1) = Format(nTSales + nTSales2 - nTCOST, "#,###,##0")
    .TextMatrix(14, 1) = Format(nTcomp, "#,###,##0")
    .TextMatrix(15, 1) = Format((nFBal - nCBal + nTPurch) - nTCOST + nTcomp, "#,###,##0")
    .TextMatrix(16, 1) = Format((nFBal - nCBal + nTPurch), "#,###,##0")
    nProft = nTSales - (nFBal - nCBal + nTPurch)
    
    nCountC = 0
    .AddItem ""
    nTCharge1 = 0
    Do While Not Charge1Table.EOF
        nCountC = nCountC + 1
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = Charge1Table!DESCA & ""
        .TextMatrix(.Rows - 1, 1) = Format(Val(Charge1Table!TCharge & ""), "#,###,##0")
        nTCharge1 = nTCharge1 + Val(Charge1Table!TCharge & "")
        Charge1Table.MoveNext
    Loop
    .AddItem ""
    If nCountC > 1 Then
        .TextMatrix(.Rows - 1, 0) = "≈Ã„«·Ï „’«—Ì›"
        .TextMatrix(.Rows - 1, 1) = Format(nTCharge1, "#,###,##0")
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbYellow
    End If
    
    nCountC = 0
    .AddItem ""
    nTCharge2 = 0
    Do While Not Charge2Table.EOF
        nCountC = nCountC + 1
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = Charge2Table!DESCA & ""
        .TextMatrix(.Rows - 1, 1) = Format(Val(Charge2Table!TCharge & ""), "#,###,##0")
        nTCharge2 = nTCharge2 + Val(Charge2Table!TCharge & "")
        Charge2Table.MoveNext
    Loop
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = "≈Ã„«·Ï ≈Ì—«œ«  "
    .TextMatrix(.Rows - 1, 1) = Format(nTCharge2, "#,###,##0")
    .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbYellow
    
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = "≈Ã„«·Ï Œ’„ „Ê—œÌ‰"
    .TextMatrix(.Rows - 1, 1) = Format(nTDICSUP, "#,###,##0")
    
    If cBranch = "00" Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = "≈Ã„«·Ï Œ’„ ⁄„·«¡ "
        .TextMatrix(.Rows - 1, 1) = Format(nTDICCUST, "#,###,##0")
    End If
    .AddItem ""
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = "’«›Ï «·—»Õ"
    .TextMatrix(.Rows - 1, 1) = Format(nProft - nTCharge1 + nTCharge2 + nTDICSUP - nTDICSUP, "#,###,##0")
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4

    If XSTORE.BoundText = "" Then
        .AddItem ""
        nTCharge1 = 0
        Do While Not PartTable.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = PartTable!DESCA & ""
            .TextMatrix(.Rows - 1, 1) = Format(Val(PartTable!TVALUE & ""), "#,###,##0")
            nTPART = nTPART + Val(PartTable!TVALUE & "")
            PartTable.MoveNext
        Loop
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = "≈Ã„«·Ï Ã«—Ï ‘—ﬂ«¡ "
        .TextMatrix(.Rows - 1, 1) = Format(nTPART, "#,###,##0")
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbYellow
    End If
    
    
    .AddItem ""
    nTPART = 0
    Do While Not AddestTable.EOF
        .AddItem ""
        .TextMatrix(.Rows - 1, 0) = AddestTable!DESCA & ""
        .TextMatrix(.Rows - 1, 1) = Format(Val(AddestTable!TCharge & ""), "#,###,##0")
        nTPART = nTPART + Val(AddestTable!TCharge & "")
        AddestTable.MoveNext
    Loop
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = "≈Ã„«·Ï „’«—Ì› «’Ê·"
    .TextMatrix(.Rows - 1, 1) = Format(nTPART, "#,###,##0")
    .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbYellow
    .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
End With
End Sub
Private Sub CMD_PRINT_Click()
Load PrintGrd
    PrintGrd.doprint grid1, 1.2, , "ﬁ«∆„… «·œŒ· " & XSTORE.text, " „‰  «—ÌŒ " & xdate1.text & " ≈·Ï  «—ÌŒ " & xDate2.text
    PrintGrd.Show 1
End Sub

Private Sub Form_Load()
    con.CursorLocation = adUseClient
    con.Open strCon

    data1.ConnectionString = strCon
    data1.RecordSource = "Select Code,DescA From File0_40 ORDER BY CODE"
    Set XSTORE.RowSource = data1
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    If cBranch <> "00" Then
        XSTORE.BoundText = cBranchStore
        XSTORE.Enabled = False
    End If

End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Function TCostItem1()
On Error Resume Next
Dim cStr1 As String
    cStr1 = " SELECT SUM(FILE1_11.[IN] - FILE1_11.OUT) AS Balance " & _
            " , COALESCE ((SELECT TOP (1) (COST_INV) AS COST FROM  FILE7_20 INNER JOIN FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO WHERE (FILE7_20.ITEM = FILE1_10.ITEM) AND (FILE7_20.PRICE <> 0) AND (FILE7_20H.date < " & DateSq(xdate1.text) & " ) " & _
             " ORDER BY FILE7_20H.date DESC), 0) AS COST INTO          BALITEM FROM         FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM " & _
            " WHERE DATE < " & DateSq(xdate1.text)
    If XSTORE.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(XSTORE.BoundText)
    cStr1 = cStr1 & " GROUP BY FILE1_10.ITEM "
    con.Execute " Drop Table BALITEM"
    con.Execute cStr1
    TCostItem1 = Val(GetDesca("SELECT SUM(Balance * COST ) FROM BALITEM ", con) & "")
End Function
Function TCostItem2()
On Error Resume Next
con.Execute " Drop Table BALITEM"
Dim cStr1 As String
    cStr1 = " SELECT  SUM(FILE1_11.[IN] - FILE1_11.OUT) AS Balance " & _
            " , COALESCE ((SELECT     TOP (1) (COST_INV)  FROM  FILE7_20 INNER JOIN FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO WHERE (FILE7_20.ITEM = FILE1_10.ITEM) AND (FILE7_20.PRICE <> 0) AND (FILE7_20H.date <= " & DateSq(xDate2.text) & " ) " & _
            " ORDER BY FILE7_20H.date DESC), 0) AS COST INTO  BALITEM FROM  FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM " & _
            " WHERE DATE <= " & DateSq(xDate2.text)
    If XSTORE.BoundText <> "" Then cStr1 = cStr1 & " AND STORE = " & MyParn(XSTORE.BoundText)
    cStr1 = cStr1 & " GROUP BY FILE1_10.ITEM "
    con.Execute cStr1
    TCostItem2 = Val(GetDesca("SELECT SUM(Balance * COST ) FROM  BALITEM ", con) & "")
End Function


