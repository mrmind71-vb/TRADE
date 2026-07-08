VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmTot 
   Caption         =   "ﬁ«∆„… «·œŒ·"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10785
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ProgressBar xItem 
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
      Top             =   585
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
      Top             =   90
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1050
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6270
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Visible         =   0   'False
         Width           =   2490
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
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
         Top             =   540
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ï  «—ÌŒ :"
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
         Left            =   5055
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   885
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
End
Attribute VB_Name = "frmTot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CMD_OK_Click()
If Not IsDate(xdate1.text) Then Exit Sub
Dim nCBal As Double
Dim nBALCUST As Double
Dim nBALBank   As Double
Dim nBALBox    As Double
Dim nBALSupp   As Double
Dim ItemTable As New ADODB.Recordset
Dim nBALchq As Double
ItemTable.Open "SELECT * FROM FILE1_10 ", con, adOpenKeyset, adLockReadOnly, adCmdText

cStr1 = "select sum( value ) from FILE5_21  WHERE   date_3 is null "
nBALchq = Val(GetDesca(cStr1, con) & "")

cStr1 = "select sum( [sal] - [PAY] ) from FILE3_11  WHERE FILE3_11.CODE <> '0000'  "
cStr1 = cStr1 & " AND DATE <= " & DateSq(xdate1.text)
nBALCUST = Val(GetDesca(cStr1, con) & "")

cStr1 = "select sum( [sal] - [PAY] ) FROM  FILE4_11 INNER JOIN FILE4_10 ON FILE4_11.code = FILE4_10.CODE INNER JOIN file4_22 ON FILE4_11.code = file4_22.code AND FILE4_11.mosm = file4_22.mosm  "
cStr1 = cStr1 & " where file4_11.DATE <= " & DateSq(xdate1.text)
nBALSupp = Val(GetDesca(cStr1, con) & "")

cStr1 = "select sum( [VALUE2] - [VALUE1] ) from BANKMOVE WHERE   "
cStr1 = cStr1 & " DATE <= " & DateSq(xdate1.text)
nBALBank = Val(GetDesca(cStr1, con) & "")

cStr1 = "select sum( [PLUS] - [MINUS] ) from boxmove INNER JOIN FILE0_50 ON FILE0_50.CODE = boxmove.BOX WHERE ISSTOP = 0 "
cStr1 = cStr1 & " AND DATE <= " & DateSq(xdate1.text)
nBALBox = Val(GetDesca(cStr1, con) & "")

nCBal = TCostItem1

With grid1
    .Cols = 2
    .Rows = 11
    .ColWidth(0) = 6000
    .ColWidth(1) = 2000
    .ColFormat(1) = "#,###,##0"
    .FixedRows = 1
    .TextMatrix(0, 0) = "»Ì«‰"
    .TextMatrix(0, 1) = "«·ﬁÌ„…"
    
    .TextMatrix(1, 0) = "ﬁÌ„…  ﬂ·›… «·√’‰«›"
    .TextMatrix(2, 0) = "—’Ìœ «·⁄„·«¡ "
    .TextMatrix(3, 0) = "—’Ìœ «·‰ﬁœÌ…"
    .TextMatrix(4, 0) = "—’Ìœ «·»‰Êﬂ"
    .TextMatrix(6, 0) = "≈Ã„«·Ï "
    .TextMatrix(8, 0) = "—’Ìœ „Ê—œÌ‰"
    .TextMatrix(9, 0) = "«Ê—«ﬁ œ›⁄ "
    .TextMatrix(10, 0) = "«·’«›Ï"
    
    .TextMatrix(1, 1) = Format(nCBal, "#,###,##0")
    .TextMatrix(2, 1) = Format(nBALCUST, "#,###,##0")
    .TextMatrix(3, 1) = Format(nBALBox, "#,###,##0")
    .TextMatrix(4, 1) = Format(nBALBank, "#,###,##0")
    
    .TextMatrix(6, 1) = Format(nCBal + nBALBank + nBALBox + nBALCUST, "#,###,##0")
    .TextMatrix(8, 1) = Format(nBALSupp, "#,###,##0")
    .TextMatrix(9, 1) = Format(nBALchq, "#,###,##0")

    .TextMatrix(10, 1) = Format((nCBal + nBALBank + nBALBox + nBALCUST) - nBALSupp - nBALchq, "#,###,##0")
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
End With
End Sub
Private Sub CMD_PRINT_Click()
Load PrintGrd
PrintGrd.doprint grid1, 1.2, , "≈Ã„«·Ï «·„Êﬁ› Õ Ï  «—ÌŒ " & xdate1.text
PrintGrd.Show 1
End Sub
Private Sub Form_Load()
    con.CursorLocation = adUseClient
    con.Open strCon
    xdate1.text = Format(Date, "DD-MM-YYYY")
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


