VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form transdataFR_ip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”Õ» «·»Ì«‰«  „‰ «·Êﬂ·«¡"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   9150
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5325
   ScaleWidth      =   9150
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "”Õ» «·»Ì«‰«  „‰ «·Êﬂ·«¡"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   585
      TabIndex        =   4
      Top             =   2295
      Width           =   5730
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   585
      TabIndex        =   2
      Top             =   3510
      Width           =   5730
   End
   Begin VB.Frame Frame11 
      Height          =   600
      Left            =   585
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4455
      Width           =   5730
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   2040
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   8835
      _cx             =   15584
      _cy             =   3598
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
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
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
Attribute VB_Name = "transdataFR_ip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lAddAll As Boolean
Dim con As New ADODB.Connection, pBranch As String
Dim conShop As New ADODB.Connection
Dim conShop_Fr As New ADODB.Connection
Dim cFilesave As String
Dim cDataFolder As String, cDataFile As String
Dim dPostDate As String

Dim cUser_Fr  As String
Dim cPass_Fr As String
Dim nBranchPost As Double
Private Function ToBranch() As Boolean
'On Error GoTo myerror
Dim lMyPrice As Boolean, lMyPriceItem As Boolean

nRecordCount = sendItems
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = SendGroup
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ „Ã„Ê⁄«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = SendMAN
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ «·»«∆⁄Ì‰", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = sendCode("FILE1_10SC", True)
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«ﬁ”«„", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = Sendfact
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„’«‰⁄", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = sendCode("FILE3_50")
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  „Ã„Ê⁄«  «·⁄„·«¡", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = sendCode("FILE4_50")
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  „Ã„Ê⁄«  „Ê—œÌ‰ ", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = sendCode("FILE4_10")
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„Ê—œÌ‰ ", "»‰Ã«Õ" Else GoTo myerror
con.Execute " update file4_10 set [group] = '01'"

nRecordCount = sendClients
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·⁄„·«¡ «·œ«∆„Ì‰", "»‰Ã«Õ" Else GoTo myerror

If lAddAll Then
    nRecordCount = SendPurchaseHeader("SELECT FILE7_20H.* , FILE0_40.BRANCH FROM FILE7_20H LEFT OUTER JOIN FILE0_40 ON FILE7_20H.STORE = FILE0_40.CODE ", "file7_20h", "file7_20")
    nRecordCount = SendPurchase("SELECT     FILE7_20.* , FILE0_40.BRANCH FROM         FILE7_20H INNER JOIN FILE7_20 ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO LEFT OUTER JOIN FILE0_40 ON FILE7_20H.STORE = FILE0_40.CODE  ", "file7_20")
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„‘ —Ì«  ", "»‰Ã«Õ" Else GoTo myerror
Else
    nRecordCount = SendPurchaseHeader("SELECT FILE7_20H.* , FILE0_40.BRANCH FROM FILE7_20H LEFT OUTER JOIN FILE0_40 ON FILE7_20H.STORE = FILE0_40.CODE  where ISNEW" & nBranchPost & " = 1 ", "file7_20h", "file7_20")
    nRecordCount = SendPurchase("SELECT     FILE7_20.* , FILE0_40.BRANCH FROM         FILE7_20H INNER JOIN FILE7_20 ON FILE7_20H.DOC_NO = FILE7_20.DOC_NO LEFT OUTER JOIN FILE0_40 ON FILE7_20H.STORE = FILE0_40.CODE  where ISNEW" & nBranchPost & "  = 1 ", "file7_20")
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„‘ —Ì«  ", "»‰Ã«Õ" Else GoTo myerror
End If

If lAddAll Then
    nRecordCount = SendPurchaseHeader("SELECT     FILE7_10H.* , FILE0_40.BRANCH FROM         FILE7_10H LEFT OUTER JOIN                       FILE0_40 ON FILE7_10H.STORE = FILE0_40.CODE  ", "file7_10h", "file7_10")
    nRecordCount = SendPurchase("SELECT     FILE7_10.* , FILE0_40.BRANCH FROM  FILE7_10H INNER JOIN FILE7_10 ON FILE7_10H.DOC_NO = FILE7_10.DOC_NO LEFT OUTER JOIN FILE0_40 ON FILE7_10H.STORE = FILE0_40.CODE   ", "file7_10")
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  „—œÊœ «·„‘ —Ì« ", "»‰Ã«Õ" Else GoTo myerror
Else
    nRecordCount = SendPurchaseHeader("SELECT     FILE7_10H.* , FILE0_40.BRANCH FROM         FILE7_10H LEFT OUTER JOIN                       FILE0_40 ON FILE7_10H.STORE = FILE0_40.CODE  where ISNEW" & nBranchPost & "  = 1 ", "file7_10h", "file7_10")
    nRecordCount = SendPurchase("SELECT     FILE7_10.* , FILE0_40.BRANCH FROM  FILE7_10H INNER JOIN FILE7_10 ON FILE7_10H.DOC_NO = FILE7_10.DOC_NO LEFT OUTER JOIN FILE0_40 ON FILE7_10H.STORE = FILE0_40.CODE   where ISNEW" & nBranchPost & "  = 1 ", "file7_10")
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  „—œÊœ «·„‘ —Ì« ", "»‰Ã«Õ" Else GoTo myerror
End If

Me.Caption = " ÕÊÌ·« 1"
nRecordCount = sendTransHeader
Me.Caption = " ÕÊÌ·« 2"
nRecordCount = sendTrans
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «· ÕÊÌ·« ", "»‰Ã«Õ" Else GoTo myerror
Me.Caption = "«·Ã—œ1"
nRecordCount = sendStockHeader
Me.Caption = "«·Ã—œ2"
nRecordCount = sendSTOCK
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  Ã—œ «·„Œ«“‰ ", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = sendbox
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·Œ“‰", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = sendstores
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„Œ«“‰", "»‰Ã«Õ" Else GoTo myerror

'nRecordCount = sendCode("FILE6_25")
'If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·»«∆⁄Ì‰", "»‰Ã«Õ" Else GoTo myerror

'nRecordCount = sendBALstores
'If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ «—’œ… «·„Œ“‰", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = sendmosm
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„Ê«”„ ", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = SendDiscount
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ «·Œ’Ê„«  ", "»‰Ã«Õ" Else GoTo myerror


nRecordCount = NewPrice
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ «”⁄«—  „  ⁄œÌ·Â«", "»‰Ã«Õ" Else GoTo myerror

lMyPrice = (GetDesca("SELECT ISOKAZ FROM BRANCH WHERE CODE = " & MyParn(pBranch)) = "True")
lMyPriceItem = (GetDesca("SELECT ISOKAZMODEL FROM BRANCH WHERE CODE = " & MyParn(pBranch)) = "True")

If lMyPrice Then
    conShop.Execute " update FILE0_00 SET PRICE = 1 "
Else
    conShop.Execute " update FILE0_00 SET PRICE = 0 "
End If

If lMyPriceItem Then
    conShop.Execute " update FILE0_00 SET OKAZ_ITEM = 1 "
Else
    conShop.Execute " update FILE0_00 SET OKAZ_ITEM = 0 "
End If

conShop.Execute " DELETE FROM MOSM      WHERE MOSM NOT IN (SELECT MOSM FROM FILE1_10)"
conShop.Execute " DELETE FROM FACT      WHERE CODE NOT IN (SELECT FACT FROM FILE1_10)"
conShop.Execute " DELETE FROM FILE4_10  WHERE CODE NOT IN (SELECT CODE FROM FILE1_10)"

ToBranch = True
MsgBox " „ ‰ﬁ· «·»Ì«‰« "
ToBranch = True
Exit Function
ToBranch = False
myerror:
MsgBox "·„ Ì „ ‰ﬁ· ﬂ«›… «·»Ì«‰«  »‰Ã«Õ"
Err.Clear
End Function

Private Sub cmd_allitem_Click()
If MsgBox(" „  ÕœÌÀ ﬂ· «·«’‰«› »«·ﬂ«„· ··›—Ê⁄", vbYesNo + vbDefaultButton2) = vbYes Then
    con.Execute " update file1_10 set isnew1 = 0 "
End If
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
End Sub
Private Sub Command1_Click()
    ToMain
End Sub
Private Sub Form_Load()
    cDataFolder = App.Path & "\mdb"
    cDataFile = "data"
    openCon con
    myload
    dPostDate = DateAdd("M", -1, Date)
End Sub
Private Function CopyData() As Boolean
Dim fs As New FileSystemObject
On Error GoTo myerror:
If fs.FileExists(cDataFolder & "\" & cDataFile & ".mdb") Then
    fs.DeleteFile cDataFolder & "\" & cDataFile & ".mdb"
End If
If fs.FileExists(cDataFolder & "\" & cDataFile & "_" & "blk.mdb") Then
    fs.CopyFile cDataFolder & "\" & cDataFile & "_" & "blk.mdb", cDataFolder & "\" & cDataFile & ".mdb"
End If
CopyData = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function SendGroup() As Long
'On Error GoTo myerror

Dim cFile As String
Dim loctable As New ADODB.Recordset

cFile = "FILE1_50"
cString = "SELECT * FROM " & cFile
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
conShop.Execute " DELETE FROM FILE1_50 "

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
            
    conShop.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    SendGroup = SendGroup + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
SendGroup = -1
Err.Clear
GoTo lastsub
End Function
Private Function sendItems() As Long
'On Error GoTo myerror
Dim cFile As String
cFile = "FILE1_10"

''''If lAddAll Then
''''    cString = "SELECT * FROM FILE1_10 WHERE ( ITEM IN (SELECT ITEM FROM ITEM_SEND ) OR ITEM IN (SELECT ITEM FROM FILE1_11 INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE WHERE  FILE0_40.BRANCH = " & MyParn(pBranch) & ") ) "
''''Else
''''    cString = "SELECT * FROM FILE1_10 WHERE ( ITEM IN (SELECT ITEM FROM ITEM_SEND ) OR ITEM IN (SELECT ITEM FROM FILE1_11 INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE WHERE FILE0_40.BRANCH = " & MyParn(pBranch) & ") ) AND ISNEW" & nBranchPost & "  = 1  "
''''End If

If lAddAll Then
    cString = "SELECT * FROM FILE1_10 WHERE  ITEM IN (SELECT ITEM FROM FILE1_11 INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE WHERE  FILE0_40.BRANCH = " & MyParn(pBranch) & ")  "
Else
    cString = "SELECT * FROM FILE1_10 WHERE  ITEM IN (SELECT ITEM FROM FILE1_11 INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE WHERE FILE0_40.BRANCH = " & MyParn(pBranch) & ") AND ISNEW" & nBranchPost & "  = 1  "
End If

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
sCaption = Me.Caption
Do Until loctable.EOF
    conShop.Execute " DELETE FROM FILE1_10 where item = " & loctable!Item
    
    Me.Caption = loctable!Item
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!Model))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable!MODELNO))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "FACT", addstring(loctable!Fact))
    aInsert = AddFlag(aInsert, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "MODELFACT", addstring(loctable!modelfact))
    aInsert = AddFlag(aInsert, "MODELFACT0", addstring(loctable!MODELFACT0))
    aInsert = AddFlag(aInsert, "SUPP", addstring(loctable!SUPP))
    aInsert = AddFlag(aInsert, "OKAZ", Val(loctable!okaz & ""))
    aInsert = AddFlag(aInsert, "RATE", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "[SECTION]", addvalue(loctable!Section))
    aInsert = AddFlag(aInsert, "[ITEM]", addvalue(loctable!Item))
    aInsert = AddFlag(aInsert, "[COST]", Val(loctable!COST & ""))
    aInsert = AddFlag(aInsert, "[COSTITEM]", Val(loctable!costITEM & ""))
    aInsert = AddFlag(aInsert, "[COST2]", Val(loctable!cost2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE]", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "[PRICE2]", Val(loctable!PRICE2 & ""))
    aInsert = AddFlag(aInsert, "[SCAL]", addstring(loctable!SCAL))
    aInsert = AddFlag(aInsert, "[C_SCAL]", addvalue(loctable!C_SCAL))
    aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable!Color))
    aInsert = AddFlag(aInsert, "[C_COLOR]", addvalue(loctable!c_Color))
    aInsert = AddFlag(aInsert, "[BARCODE]", addstring(loctable!BARCODE))
    aInsert = AddFlag(aInsert, "[ISOKAZITEM]", IIf(loctable!ISOKAZITEM, 1, 0))
    conShop.Execute addInsert(aInsert, "FILE1_10"), nAffect
    Me.Caption = loctable!Item
    con.Execute " update file1_10 set isNEW" & nBranchPost & "  = 0 WHERE ITEM = " & loctable!Item
    loctable.MoveNext
    sendItems = sendItems + nAffect
Loop

lastsub:
Me.Caption = sCaption
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendItems = -1
GoTo lastsub
End Function
Private Function sendCode(cFile, Optional isNumber As Boolean = False) As Long
On Error GoTo myerror
cString = "SELECT * FROM " & cFile
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
conShop.Execute " delete from " & cFile
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100

    If isNumber Then
        aInsert = AddFlag(Empty, "CODE", addvalue(loctable!code))
    Else
        aInsert = AddFlag(Empty, "CODE", addstring(loctable!code))
    End If
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    conShop.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendCode = sendCode + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendCode = -1
GoTo lastsub
End Function
Private Function sendstores()
'On Error GoTo myerror
cFile = "FILE0_40"
cString = "SELECT * FROM FILE0_40 "

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop.Execute " delete from file0_40"
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "SERVER", addstring(loctable!SERVER))
    aInsert = AddFlag(aInsert, "USERNAME", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "PASSWORD", addstring(loctable!PassWord))
    
    conShop.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendstores = sendstores + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendstores = -1
GoTo lastsub
End Function
Private Function sendClients() As Long
On Error GoTo myerror
Dim cFile As String
cFile = "FILE3_10"
cString = "SELECT * FROM " & cFile
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop.Execute " delete from file3_10"
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "MANAGER", addstring(loctable!Manager))
    aInsert = AddFlag(aInsert, "ADDRESS", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "PHONE1", addstring(loctable!PHONE1))
    aInsert = AddFlag(aInsert, "FAX", addstring(loctable!Fax))
    aInsert = AddFlag(aInsert, "F_BALANCE", Val(loctable!F_Balance & ""))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "[F_DATE]", addstring(Format(loctable!F_DATE, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[EMAIL]", addstring(loctable!email))
    aInsert = AddFlag(aInsert, "[REMARK]", addstring(loctable!remark))
    aInsert = AddFlag(aInsert, "DISC", Val(loctable!disc & ""))
    aInsert = AddFlag(aInsert, "[CASH]", IIf(loctable!CASH, 1, 0))
           
    conShop.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendClients = sendClients + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendClients = -1
GoTo lastsub
End Function
Private Function SendPurchaseHeader(cString, cFile, cFile2) As Long
'On Error GoTo myerror
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    conShop.Execute " delete from " & cFile & " where doc_no = " & MyParn(loctable!doc_no)
    conShop.Execute " delete from " & cFile2 & " where doc_no = " & MyParn(loctable!doc_no)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!branch = "03" Then
'        MsgBox loctable!branch
    End If
    
    If loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "CODE", addstring(loctable!code))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!store))
        aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
        aInsert = AddFlag(aInsert, "[TAX]", Val(loctable!tax & ""))
        aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM))
        aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
        conShop.Execute addInsert(aInsert, cFile), nAffect
    End If
    loctable.MoveNext
    SendPurchaseHeader = SendPurchaseHeader + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendPurchaseHeader = -1
GoTo lastsub
End Function
Private Function SendPurchase(cString, cFile) As Long
'On Error GoTo myerror
Dim loctable As New ADODB.Recordset

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100

    If loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "ITEM", addvalue(loctable!Item))
        aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant & ""))
        aInsert = AddFlag(aInsert, "PRICE", Val(loctable!price & ""))
        conShop.Execute addInsert(aInsert, cFile), nAffect
    End If
    loctable.MoveNext
    SendPurchase = SendPurchase + nAffect
Loop
con.Execute " update " & cFile & "H  SET ISNEW" & nBranchPost & "  = 0 "
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendPurchase = -1
GoTo lastsub
End Function
Private Function sendTransHeader() As Long
cFile = "FILE1_60H"

If lAddAll Then
    cString = "SELECT  FILE0_40.branch AS branch_1 ,FILE0_40_1.branch AS BRANCH , file1_60h.* FROM         file1_60h INNER JOIN                       FILE0_40 ON file1_60h.store1 = FILE0_40.CODE INNER JOIN                       FILE0_40 AS FILE0_40_1 ON file1_60h.store2 = FILE0_40_1.CODE "
Else
    cString = "SELECT  FILE0_40.branch AS branch_1 ,FILE0_40_1.branch AS BRANCH , file1_60h.* FROM         file1_60h INNER JOIN                       FILE0_40 ON file1_60h.store1 = FILE0_40.CODE INNER JOIN                       FILE0_40 AS FILE0_40_1 ON file1_60h.store2 = FILE0_40_1.CODE where ISNEW" & nBranchPost & "  = 1 "
End If
Dim loctable As New ADODB.Recordset

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    conShop.Execute " delete from file1_60  where doc_no = " & MyParn(loctable!doc_no)
    conShop.Execute " delete from file1_60h where doc_no = " & MyParn(loctable!doc_no)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!BRANCH_1 = pBranch Or loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
        aInsert = AddFlag(aInsert, "[STORE1]", addstring(loctable!STORE1))
        aInsert = AddFlag(aInsert, "[STORE2]", addstring(loctable!STORE2))
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
        conShop.Execute addInsert(aInsert, cFile), nAffect
    End If
    loctable.MoveNext
    sendTransHeader = sendTransHeader + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTransHeader = -1
GoTo lastsub
End Function
Private Function sendTrans() As Long
'On Error GoTo myerror
cFile = "FILE1_60"
If lAddAll Then
    cString = "SELECT  FILE1_60.* , FILE0_40.branch AS BRANCH_1 , FILE0_40_1.branch  AS BRANCH FROM  FILE1_60 INNER JOIN file1_60h ON FILE1_60.DOC_NO = file1_60h.doc_no INNER JOIN  FILE0_40 ON file1_60h.store1 = FILE0_40.CODE INNER JOIN FILE0_40 AS FILE0_40_1 ON file1_60h.store2 = FILE0_40_1.CODE   "
Else
    cString = "SELECT  FILE1_60.* , FILE0_40.branch AS BRANCH_1 , FILE0_40_1.branch  AS BRANCH FROM  FILE1_60 INNER JOIN file1_60h ON FILE1_60.DOC_NO = file1_60h.doc_no INNER JOIN  FILE0_40 ON file1_60h.store1 = FILE0_40.CODE INNER JOIN FILE0_40 AS FILE0_40_1 ON file1_60h.store2 = FILE0_40_1.CODE   where ISNEW" & nBranchPost & "  = 1 "
End If
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    Me.Caption = " ÕÊÌ·« "
    If loctable!BRANCH_1 = pBranch Or loctable!branch = pBranch Then
        Me.Caption = loctable!doc_no
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "ITEM", addvalue(loctable!Item))
        aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant & ""))
        aInsert = AddFlag(aInsert, "COST", Val(loctable!COST & ""))
        conShop.Execute addInsert(aInsert, cFile), nAffect
    End If
    con.Execute " UPDATE FILE1_60H SET ISNEW" & nBranchPost & "  = 0 where doc_no = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    sendTrans = sendTrans + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTrans = -1
GoTo lastsub
End Function
Private Function sendSTOCK() As Long
cFile = "FILE0_10"
If lAddAll Then
    cString = "SELECT   FILE0_10.* , FILE0_40.BRANCH  FROM  FILE0_10 INNER JOIN file0_10h ON FILE0_10.DOC_NO = file0_10h.doc_no INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE0_10H.STORE WHERE FILE0_10H.branch IS NULL "
Else
    cString = "SELECT   FILE0_10.* , FILE0_40.BRANCH  FROM  FILE0_10 INNER JOIN file0_10h ON FILE0_10.DOC_NO = file0_10h.doc_no INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE0_10H.STORE where FILE0_10H.branch IS NULL  AND isnew" & nBranchPost & " = 1  "
End If
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    If loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "[ITEM]", addstring(loctable!Item))
        aInsert = AddFlag(aInsert, "[RealBal]", Val(loctable!RealBal & ""))
        aInsert = AddFlag(aInsert, "[ComputerBal]", Val(loctable!ComputerBal & ""))
        aInsert = AddFlag(aInsert, "[Differ]", Val(loctable!Differ & ""))
        conShop.Execute addInsert(aInsert, cFile), nAffect
        conShop.Execute addInsert(aInsert, "FILE0_100"), nAffect
    End If
    loctable.MoveNext
    sendSTOCK = sendSTOCK + nAffect
Loop
con.Execute " UPDATE FILE0_10H set isnew" & nBranchPost & "  = 0 "
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendSTOCK = -1
GoTo lastsub
End Function

Private Function sendmosm() As Long
'On Error GoTo myerror
cFile = "mosm"
cString = "SELECT * FROM MOSM "

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop.Execute " delete from mosm "
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
    conShop.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendmosm = sendmosm + nAffect
Loop
lastsub:
prog1.Visible = False

'If lokaz Then
'    conShop.Execute " update file0_00 set price = 1 "
'Else
'    conShop.Execute " update file0_00 set price = 0 "
'End If

Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendmosm = -1
GoTo lastsub
End Function
Private Function sendbox()
'On Error GoTo myerror
cFile = "FILE0_50"
cString = "SELECT * FROM FILE0_50 where branch = " & MyParn(pBranch)

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop.Execute " delete from file0_50"
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[F_DATE]", addDate(loctable!F_DATE))
    aInsert = AddFlag(aInsert, "[ISVISA]", IIf(loctable!ISVISA, 1, 0))
    aInsert = AddFlag(aInsert, "[BRANCH]", addDate(loctable!branch))
    
    conShop.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendbox = sendbox + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendbox = -1
GoTo lastsub
End Function


Private Function Sendfact() As Long
On Error GoTo myerror
Dim cFile As String
cFile = "fact"
cString = "SELECT FACT.CODE , FACT.DESCA FROM FACT "
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
conShop.Execute " delete from fact "
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
            
    conShop.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    Sendfact = Sendfact + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Sendfact = -1
Err.Clear
GoTo lastsub
End Function
Private Sub myload()
openCon con
Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT * FROM BRANCH_FR  "
cString = cString & " ORDER BY CODE"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With grid1
.Cols = 7
Do Until loctable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = loctable(0) & ""
    .TextMatrix(.Rows - 1, 1) = loctable!DESCA & ""
    .TextMatrix(.Rows - 1, 2) = loctable!ipsql & ""
    .TextMatrix(.Rows - 1, 3) = loctable!data & ""
    .TextMatrix(.Rows - 1, 4) = loctable!UserName & ""
    .TextMatrix(.Rows - 1, 5) = loctable!PassWord & ""
    If (loctable!ipsql & "") = "" Then
        .TextMatrix(.Rows - 1, 6) = False
    Else
        .TextMatrix(.Rows - 1, 6) = True
    End If
    loctable.MoveNext
Loop
Fixgrd
End With
End Sub
Private Function ToMain() As Boolean
'On Error GoTo myerror
With grid1
For nRow = 1 To .Rows - 1
    If .TextMatrix(nRow, 2) <> "" And TurnValue(.TextMatrix(nRow, 6), "True", True) Then
        IpShop = .TextMatrix(nRow, 2)
        cDataShop = .TextMatrix(nRow, 3)
        cPass_Fr = .TextMatrix(nRow, 5)
        cUser_Fr = .TextMatrix(nRow, 4)
        strConShop_Fr = LoadConStringshop_fr(IpShop, cDataShop, cUser_Fr, cPass_Fr)
        If openConShop(conShop_Fr, strConShop_Fr) = "ok" Then
            .Cell(flexcpBackColor, nRow, 1, nRow, .Cols - 1) = vbYellow
            If Not getData_FR(.TextMatrix(nRow, 0)) Then
                Inform "„‘ﬂ·… «À‰«¡ ‰ﬁ· „·›«  «·›—⁄ : " & .TextMatrix(nRow, 1)
    '            Exit Function
            Else
                Inform " „ ‰”Œ „·›«  ›—⁄ : " & .TextMatrix(nRow, 1), " »‰Ã«Õ", 1200
            End If
        Else
            Inform "„‘ﬂ·… »«·≈ ’«· »«·›—⁄ : " & .TextMatrix(nRow, 1)
            .Cell(flexcpBackColor, nRow, 1, nRow, .Cols - 1) = vbRed
        End If
    End If
Next
End With
ToMain = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function getData_FR(sBranch As String) As Boolean
Dim nRecordCount As Long
'    On Error GoTo myerror
    
'    Inform "„Ì⁄« "
'    nRecordCount = GetSales(sBranch)
'    If nRecordCount >= 0 Then
'        Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ Õ—ﬂ«  «·»Ì⁄", "»‰Ã«Õ"
'    Else
'        GoTo myerror
'    End If

    Inform "„‘ —Ì« "
    nRecordCount = GetPurches(sBranch)
    If nRecordCount >= 0 Then
        Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ Õ—ﬂ«  «·„‘ —Ì«  ", "»‰Ã«Õ"
    Else
        GoTo myerror
    End If

    Inform "„—œÊœ „‘ —Ì« "
    nRecordCount = GetPurchesR(sBranch)
    If nRecordCount >= 0 Then
        Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ Õ—ﬂ«  „—œÊœ «·„‘ —Ì«  ", "»‰Ã«Õ"
    Else
        GoTo myerror
    End If

getData_FR = True
Exit Function
MsgBox Err.Description
Err.Clear
myerror:
End Function
Private Function GetSales(sBranch As String) As Long
Dim loctable As New ADODB.Recordset
Dim cDoc As String
cString = "Select * from FILE6_20H WHERE ISNEW = 1 AND PRINTED = 1 "
If loctable.State = adStateOpen Then loctable.Close
loctable.Open cString, conShop_Fr, adOpenStatic, adLockReadOnly, adCmdText

'con.Execute " delete from FR6_20  where DOC_NO IN (SELECT DOC_NO FROM FR6_20H WHERE BRANCH = " & MyParn(sBranch) & " ) "
'con.Execute " delete from FR6_20H where BRANCH = " & MyParn(sBranch)

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
    cDoc = Mid(loctable!doc_no, 1, 6) & sBranch & Mid(loctable!doc_no, 7, 4)
    
    con.Execute " delete from FR6_20 where doc_no = " & MyParn(cDoc)
    con.Execute " delete from FR6_20h where doc_no = " & MyParn(cDoc)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(cDoc))
    aInsert = AddFlag(aInsert, "CODE", addstring("0000"))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[TIME]", addTime(Format(loctable!Time, "SHORT TIME")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(sBranch))
    aInsert = AddFlag(aInsert, "[NOTES]", addstring(loctable!NOTES))
    aInsert = AddFlag(aInsert, "[RATE]", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[CASH]", Val(loctable!CASH & ""))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(sBranch))
    aInsert = AddFlag(aInsert, "[PAY]", Val(loctable!PAY & ""))
    aInsert = AddFlag(aInsert, "[VISA]", Val(loctable!Visa & ""))
    aInsert = AddFlag(aInsert, "[LATE]", Val(loctable!late & ""))
    aInsert = AddFlag(aInsert, "[REST]", Val(loctable!Rest & ""))
    aInsert = AddFlag(aInsert, "[MAN]", addstring("01"))
    aInsert = AddFlag(aInsert, "[PRINTED]", IIf(loctable!PRINTED, "1", "0"))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranch))
                                                                                
    con.Execute addInsert(aInsert, "FR6_20H"), nAffect
    
    Dim loctable2 As New ADODB.Recordset
    cString = "Select FILE6_20.* from FILE6_20 WHERE FILE6_20.doc_no = " & MyParn(loctable!doc_no)
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open cString, conShop_Fr, adOpenStatic, adLockReadOnly, adCmdText
    Dim sitem As String
    Dim aitem As Variant
    Do Until loctable2.EOF
        sitem = loctable2!Item
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(cDoc))
        aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
        aInsert2 = AddFlag(aInsert2, "MAN", addstring("01"))
        aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
        aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
        aInsert2 = AddFlag(aInsert2, "DISCOUNT", Val(loctable2!discount & ""))
        aInsert2 = AddFlag(aInsert2, "PRICE_C", Val(loctable2!PRICE_C & ""))
        aInsert2 = AddFlag(aInsert2, "ROW", Val(loctable2!Row & ""))
        aInsert2 = AddFlag(aInsert2, "COST", Val(loctable2!COST & ""))
        aInsert2 = AddFlag(aInsert2, "S_okaz", Val(loctable2!S_OKAZ & ""))
        con.Execute addInsert(aInsert2, "FR6_20"), nAffect
        loctable2.MoveNext
    Loop

    conShop_Fr.Execute " update file6_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
GetSales = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GetSales = -1
GoTo lastsub
End Function
Private Sub Fixgrd()
With grid1
.Cols = 7
.ColWidth(0) = 500
.ColWidth(1) = 1500
.ColWidth(2) = 2800
.ColWidth(3) = 1300
.ColWidth(4) = 900
.ColWidth(5) = 900
.ColWidth(6) = 700
.TextMatrix(0, 0) = "ﬂÊœ «·›—⁄"
.TextMatrix(0, 1) = "≈”„ «·›—⁄"
.TextMatrix(0, 2) = "IP"
.TextMatrix(0, 3) = "DATA"
.TextMatrix(0, 4) = "USERNAME"
.TextMatrix(0, 5) = "PASSWORD"

.ColDataType(6) = flexDTBoolean

.ColAlignment(0) = flexAlignRightCenter
.ColAlignment(1) = flexAlignRightCenter

End With
End Sub
Private Function validData(sBranch As String) As Boolean
Dim fs As New FileSystemObject
If Not fs.FileExists(cDataFolder & "\" & cDataFile & "_" & sBranch & ".mdb") Then Exit Function
validData = True
End Function
Private Function sendStockHeader() As Long
cFile = "FILE0_10H"

If lAddAll Then
    cString = "SELECT  FILE0_10H.* , FILE0_40.BRANCH FROM  file0_10h INNER JOIN FILE0_40 ON FILE0_10H.STORE = FILE0_40.CODE WHERE FILE0_10H.branch IS NULL "
Else
    cString = "SELECT  FILE0_10H.* , FILE0_40.BRANCH FROM  file0_10h INNER JOIN FILE0_40 ON FILE0_10H.STORE = FILE0_40.CODE where FILE0_10H.branch IS NULL AND isnew" & nBranchPost & "  = 1 "
End If
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    conShop.Execute " delete from file0_10 where doc_no = " & MyParn(loctable!doc_no)
    conShop.Execute " delete from file0_100 where doc_no = " & MyParn(loctable!doc_no)
    conShop.Execute " delete from file0_10h where doc_no = " & MyParn(loctable!doc_no)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!branch = pBranch Then
'        Me.Caption = loctable!doc_no & ""
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!store))
        aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
        aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
'        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
        conShop.Execute addInsert(aInsert, cFile), nAffect
    End If
    loctable.MoveNext
    sendStockHeader = sendStockHeader + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendStockHeader = -1
GoTo lastsub
End Function


Private Function LoadConStringshop_fr(pIp, pData, pUser, pPass)
LoadConStringshop_fr = "provider=SQLOLEDB;data source=" & pIp & ";initial " _
            & "catalog=" & pData & ";user id = " & pUser & ";" & "password = " & pPass & ";Timeout=10"
End Function

Private Sub grid1_EnterCell()
With grid1
    If .Col >= 4 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub

Private Function sendBALstores()
'On Error GoTo myerror
cString = "SELECT * FROM Q_BAL1"

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop.Execute " delete from BAL_1"
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "ITEM", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "BAL1", Val(loctable!BAL1))
    conShop.Execute addInsert(aInsert, "BAL_1"), nAffect
    loctable.MoveNext
    sendBALstores = sendBALstores + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendBALstores = -1
GoTo lastsub
End Function
Private Function getStock(pBranch) As Long
cFile = "FILE0_10H"

cString = "SELECT  FILE0_10H.* FROM  file0_10h where isnew  = 1 AND BRANCH = " & MyParn(pBranch)
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    con.Execute " delete from file0_10 where doc_no = " & MyParn(loctable!doc_no)
    con.Execute " delete from file0_100 where doc_no = " & MyParn(loctable!doc_no)
    con.Execute " delete from file0_10h where doc_no = " & MyParn(loctable!doc_no)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!store))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
'    aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
    con.Execute addInsert(aInsert, cFile), nAffect
    
    Dim loctable2 As New ADODB.Recordset
    loctable2.Open "SELECT * FROM FILE0_10 WHERE DOC_NO = " & MyParn(loctable!doc_no), conShop, adOpenStatic, adLockReadOnly, adCmdText
    Do Until loctable2.EOF
        Me.Caption = loctable2!Item
        If GetDesca("select item from a1_10 where item = " & loctable2!Item) = "" Then
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "[ITEM]", addstring(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "[RealBal]", Val(loctable2!RealBal & ""))
            aInsert2 = AddFlag(aInsert2, "[ComputerBal]", Val(loctable2!ComputerBal & ""))
            aInsert2 = AddFlag(aInsert2, "[Differ]", Val(loctable2!Differ & ""))
            con.Execute addInsert(aInsert2, "file0_10"), nAffect
            con.Execute addInsert(aInsert2, "FILE0_100"), nAffect
            Me.Caption = loctable2!doc_no & " " & loctable2!Item
        End If
        loctable2.MoveNext
    Loop
    conShop.Execute " UPDATE FILE0_10H SET ISNEW = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    getStock = getStock + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getStock = -1
GoTo lastsub
End Function
Private Function SendMAN() As Long
'On Error GoTo myerror

Dim cFile As String
Dim loctable As New ADODB.Recordset

cFile = "FILE6_25"
cString = "SELECT * FROM FILE6_25 "
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If IsNull(loctable!branch) Or loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "CODE", addstring(loctable!code))
        aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
        aInsert = AddFlag(aInsert, "BRANCH", addstring(loctable!branch))
                
        conShop.Execute " DELETE FROM FILE6_25 WHERE CODE =  " & MyParn(loctable!code)
        conShop.Execute addInsert(aInsert, cFile), nAffect
    End If
    loctable.MoveNext
    SendMAN = SendMAN + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
SendMAN = -1
Err.Clear
GoTo lastsub
End Function
Private Function NewPrice() As Long
On Error GoTo myerror
cString = "SELECT * FROM NEWPRICEITEM WHERE ISNEW" & nBranchPost & "  = 1  "
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    conShop.Execute " delete from NEWPRICEITEM where model = " & MyParn(loctable!Model)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!Model))
    conShop.Execute addInsert(aInsert, "NEWPRICEITEM"), nAffect
    con.Execute " update    NEWPRICEITEM set    ISNEW" & nBranchPost & "  = 0 where model = " & MyParn(loctable!Model)
    con.Execute " delete    from NEWPRICEITEM  where ISNEW1 = 0 and ISNEW2 = 0 and ISNEW3 = 0 and ISNEW4 = 0 and ISNEW5 = 0 and model = " & MyParn(loctable!Model)
    loctable.MoveNext
    NewPrice = NewPrice + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
NewPrice = -1
GoTo lastsub
End Function


Private Function SendDiscount() As Long
'On Error GoTo myerror
cString = "SELECT * FROM DISCOUNT "

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop.Execute " delete from DISCOUNT "
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "ROW", addstring(loctable!Row))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[DATE1]", addDate(Format(loctable!Date1, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[DATE2]", addDate(Format(loctable!date2, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[DISC]", Val(loctable!disc & ""))
    conShop.Execute addInsert(aInsert, "DISCOUNT"), nAffect
    loctable.MoveNext
    SendDiscount = SendDiscount + nAffect
Loop
lastsub:
prog1.Visible = False

Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendDiscount = -1
GoTo lastsub
End Function

Private Function GetPurches(sBranch As String) As Long
Set loctable = New ADODB.Recordset
Dim cDoc As String

cString = "Select * from FILE7_20H  WHERE ISNEW = 1"
loctable.Open cString, conShop_Fr, adOpenStatic, adLockReadOnly, adCmdText

'con.Execute " delete from FR7_20  where DOC_NO IN (SELECT DOC_NO FROM FR7_20H WHERE BRANCH = " & MyParn(sBranch) & " ) "
'con.Execute " delete from FR7_20H where BRANCH = " & MyParn(sBranch)

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
    cDoc = loctable!doc_no
    con.Execute " delete from FR7_20 where doc_no = " & MyParn(cDoc)
    con.Execute " delete from FR7_20h where doc_no = " & MyParn(cDoc)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(cDoc))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(sBranch))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranch))
    aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM & ""))
    aInsert = AddFlag(aInsert, "[REMARK]", Val(loctable!remark & ""))
    aInsert = AddFlag(aInsert, "[TAX]", Val(loctable!tax & ""))
    aInsert = AddFlag(aInsert, "[R_TAX]", Val(loctable!R_TAX & ""))
    aInsert = AddFlag(aInsert, "[Received]", 1)

    con.Execute addInsert(aInsert, "FR7_20H"), nAffect
    
    Set loctable2 = New ADODB.Recordset
    
    cString = "Select FILE7_20.* from FILE7_20 WHERE FILE7_20.doc_no = " & MyParn(loctable!doc_no)
    loctable2.Open cString, conShop_Fr, adOpenStatic, adLockReadOnly, adCmdText
    
    Dim sitem As String
    Dim aitem As Variant
    Do Until loctable2.EOF
        sitem = loctable2!Item
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(cDoc))
        aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
        aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
        aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
        aInsert2 = AddFlag(aInsert2, "DISCOUNT", Val(loctable2!discount & ""))
        aInsert2 = AddFlag(aInsert2, "R_TAX", Val(loctable2!R_TAX & ""))
        aInsert2 = AddFlag(aInsert2, "ROW", Val(loctable2!Row & ""))
        con.Execute addInsert(aInsert2, "FR7_20"), nAffect
        loctable2.MoveNext
    Loop

    conShop_Fr.Execute " update file7_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
GetPurches = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GetPurches = -1
GoTo lastsub
End Function

Private Function GetPurchesR(sBranch As String) As Long
Set loctable = New ADODB.Recordset
Dim cDoc As String

cString = "Select * from FILE7_10H " ' WHERE ISNEW = 1"
loctable.Open cString, conShop_Fr, adOpenStatic, adLockReadOnly, adCmdText

'con.Execute " delete from FR7_20  where DOC_NO IN (SELECT DOC_NO FROM FR7_20H WHERE BRANCH = " & MyParn(sBranch) & " ) "
'con.Execute " delete from FR7_20H where BRANCH = " & MyParn(sBranch)

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
    cDoc = loctable!doc_no
    con.Execute " delete from FR7_10 where doc_no = " & MyParn(cDoc)
    con.Execute " delete from FR7_10h where doc_no = " & MyParn(cDoc)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(cDoc))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(sBranch))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranch))
    aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM & ""))
    aInsert = AddFlag(aInsert, "[REMARK]", Val(loctable!remark & ""))
    aInsert = AddFlag(aInsert, "[TAX]", Val(loctable!tax & ""))
    aInsert = AddFlag(aInsert, "[R_TAX]", Val(loctable!R_TAX & ""))
    aInsert = AddFlag(aInsert, "[Received]", 1)
    con.Execute addInsert(aInsert, "FR7_10H"), nAffect
    
    Set loctable2 = New ADODB.Recordset
    
    cString = "Select FILE7_10.* from FILE7_10 WHERE FILE7_10.doc_no = " & MyParn(loctable!doc_no)
    loctable2.Open cString, conShop_Fr, adOpenStatic, adLockReadOnly, adCmdText
    
    Dim sitem As String
    Dim aitem As Variant
    Do Until loctable2.EOF
        sitem = loctable2!Item
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(cDoc))
        aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
        aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
        aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
        aInsert2 = AddFlag(aInsert2, "DISCOUNT", Val(loctable2!discount & ""))
        aInsert2 = AddFlag(aInsert2, "ROW", Val(loctable2!Row & ""))
        aInsert2 = AddFlag(aInsert2, "R_TAX", Val(loctable2!R_TAX & ""))
        con.Execute addInsert(aInsert2, "FR7_10"), nAffect
        loctable2.MoveNext
    Loop

    conShop_Fr.Execute " update file7_10h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
GetPurchesR = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GetPurchesR = -1
GoTo lastsub
End Function

