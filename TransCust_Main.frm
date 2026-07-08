VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TransCust_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÓÍÈ ÇáÈíÇäÇÊ ãä ÇáÑÆíÓí"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   5910
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5640
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox xaddall 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      Caption         =   "ÊÑÍíá ßá ÇáÈíÇäÇÊ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5130
      Width           =   3165
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ÊÍÏíË ãÈíÚÇÊ ãä ÇáÚãáÇÁ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   90
      TabIndex        =   5
      Top             =   2970
      Visible         =   0   'False
      Width           =   5730
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "ÎÑæÌ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   90
      TabIndex        =   3
      Top             =   3780
      Width           =   5730
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "ÖÈØ ãáÝ ÈíÇäÇÊ ááÚãáÇÁ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   90
      TabIndex        =   0
      Top             =   2160
      Width           =   5730
   End
   Begin VB.Frame Frame11 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4455
      Width           =   5730
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   2
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
      TabIndex        =   4
      Top             =   45
      Visible         =   0   'False
      Width           =   5730
      _cx             =   10107
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
Attribute VB_Name = "TransCust_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim cFilesave As String
Dim cDataFolder As String, cDataFile As String
Dim dPostDate As String
Private Function ToBranch() As Boolean
'On Error GoTo myerror

If Not CopyData Then
    MsgBox "ãáÝ äÞá ÇáÈíÇäÇÊ ÛíÑ ãæÌæÏ"
    Exit Function
End If

nRecordCount = SendDiscount
If nRecordCount >= 0 Then Inform "Êã äÞá " & nRecordCount & " ÓÌá ãä ÈíÇäÇÊ ÇáÎÕæãÇÊ ", "ÈäÌÇÍ" Else GoTo myerror

nRecordCount = SendGroup
If nRecordCount >= 0 Then Inform "Êã äÞá " & nRecordCount & " ÓÌá ãä ãÌãæÚÇÊ ÇáÇÕäÇÝ", "ÈäÌÇÍ" Else GoTo myerror

nRecordCount = sendItems
If nRecordCount >= 0 Then Inform "Êã äÞá " & nRecordCount & " ÓÌá ãä ÈíÇäÇÊ ÇáÇÕäÇÝ", "ÈäÌÇÍ" Else GoTo myerror

nRecordCount = sendCode("FILE1_10SC", True)
If nRecordCount >= 0 Then Inform "Êã äÞá " & nRecordCount & " ÓÌá ãä ÈíÇäÇÊ ÇáÇÞÓÇã", "ÈäÌÇÍ" Else GoTo myerror

nRecordCount = Sendfact
If nRecordCount >= 0 Then Inform "Êã äÞá " & nRecordCount & " ÓÌá ãä ÈíÇäÇÊ ÇáãÕÇäÚ", "ÈäÌÇÍ" Else GoTo myerror

nRecordCount = sendCode("FILE3_50")
If nRecordCount >= 0 Then Inform "Êã äÞá " & nRecordCount & " ÓÌá ãä ÈíÇäÇÊ ãÌãæÚÇÊ ÇáÚãáÇÁ", "ÈäÌÇÍ" Else GoTo myerror

nRecordCount = SendSupler
If nRecordCount >= 0 Then Inform "Êã äÞá " & nRecordCount & " ÓÌá ãä ÈíÇäÇÊ ÇáãæÑÏíä ", "ÈäÌÇÍ" Else GoTo myerror

nRecordCount = sendmosm
If nRecordCount >= 0 Then Inform "Êã äÞá " & nRecordCount & " ÓÌá ãä ÈíÇäÇÊ ÇáãæÇÓã ", "ÈäÌÇÍ" Else GoTo myerror


ToBranch = True
MsgBox "Êã äÞá ÇáÈíÇäÇÊ"
Exit Function
myerror:
MsgBox "áã íÊã äÞá ßÇÝÉ ÇáÈíÇäÇÊ ÈäÌÇÍ"
Err.Clear
End Function

Private Sub cmd_allitem_Click()
If MsgBox("Êã ÊÍÏíË ßá ÇáÇÕäÇÝ ÈÇáßÇãá ááÝÑæÚ", vbYesNo + vbDefaultButton2) = vbYes Then
    con.Execute " update file1_10 set isnew1 = 0 "
End If
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    Me.Caption = "äÞá ÇáÈíÇäÇÊ ãä ÇáÅÏÇÑÉ Åáì ÇáÚãáÇÁ "
    If Not ToBranch Then
        MsgBox "áã íÊãßä ÇáäÙÇã ãä äÞá ÇáÈíÇäÇÊ Åáì ÇáæßáÇÁ"
    Else
        MsgBox "Êãßä ÇáäÙÇã ãä äÞá ÇáÈíÇäÇÊ Åáì ÇáæßáÇÁ"
        Unload Me
    End If
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
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

Dim cFile As String
cFile = "FILE1_50"
cString = "SELECT * FROM " & cFile
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

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
            
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    SendGroup = SendGroup + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
SendGroup = -1
Err.Clear
GoTo lastsub
End Function
Private Function sendItems() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
Dim cFile As String
cFile = "FILE1_10"

'cString = "SELECT * FROM FILE1_10 WHERE ITEM IN (SELECT ITEM FROM ITEM_SEND) "
cString = "SELECT * FROM FILE1_10 "

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
sCaption = Me.Caption
Do Until loctable.EOF
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
    aInsert = AddFlag(aInsert, "[COST]", Val(loctable!cost2 & ""))
    aInsert = AddFlag(aInsert, "[COST2]", Val(loctable!cost2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE]", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "[PRICE2]", Val(loctable!PRICE2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE_2]", Val(loctable!PRICE_2 & ""))
    aInsert = AddFlag(aInsert, "[OKAZ_2]", Val(loctable!okaz_2 & ""))
    aInsert = AddFlag(aInsert, "[SCAL]", addstring(loctable!SCAL))
    aInsert = AddFlag(aInsert, "[C_SCAL]", addvalue(loctable!C_SCAL))
    aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable!Color))
    aInsert = AddFlag(aInsert, "[C_COLOR]", addvalue(loctable!c_Color))
    aInsert = AddFlag(aInsert, "[code_COLOR]", addvalue(loctable!code_Color))
    aInsert = AddFlag(aInsert, "[code_scal]", addvalue(loctable!CODE_SCAL))
    aInsert = AddFlag(aInsert, "[BARCODE]", addstring(loctable!BARCODE))
    aInsert = AddFlag(aInsert, "[BARCODE2]", addstring(loctable!BARCODE2))
    aInsert = AddFlag(aInsert, "[ISOKAZITEM]", IIf(loctable!ISOKAZITEM, 1, 0))
    
    conmdb.Execute addInsert(aInsert, "FILE1_10"), nAffect
    Me.Caption = loctable!Item
    loctable.MoveNext
    sendItems = sendItems + nAffect
Loop
lastsub:
Me.Caption = sCaption
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendItems = -1
GoTo lastsub
End Function
Private Function sendCode(cFile, Optional isNumber As Boolean = False) As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
cString = "SELECT * FROM " & cFile

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100

    If isNumber Then
        aInsert = AddFlag(Empty, "CODE", addvalue(loctable!code))
    Else
        aInsert = AddFlag(Empty, "CODE", addstring(loctable!code))
    End If
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendCode = sendCode + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendCode = -1
GoTo lastsub
End Function
Private Function sendstores()
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
cFile = "FILE0_40"
cString = "SELECT * FROM FILE0_40 "

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
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
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendstores = sendstores + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendstores = -1
GoTo lastsub
End Function
Private Function sendClients() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
Dim cFile As String
cFile = "FILE3_10"
cString = "SELECT * FROM " & cFile
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

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
           
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendClients = sendClients + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendClients = -1
GoTo lastsub
End Function


Private Function SendPurchaseHeader(cString, cFile) As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100

    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "[DATE]", addstring(Format(loctable!Date, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!store))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[TAX]", Val(loctable!tax & ""))
    aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    SendPurchaseHeader = SendPurchaseHeader + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendPurchaseHeader = -1
GoTo lastsub
End Function
Private Function SendPurchase(cString, cFile) As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100

    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "ITEM", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant & ""))
    aInsert = AddFlag(aInsert, "PRICE", Val(loctable!price & ""))
       
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    SendPurchase = SendPurchase + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendPurchase = -1
GoTo lastsub
End Function
Private Function sendTransHeader() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, cFile As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

cFile = "FILE1_60H"
cString = "SELECT     file1_60h.* FROM         file1_60h where isdate >= " & DateSq(dPostDate)
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addstring(Format(loctable!Date, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[STORE1]", addstring(loctable!STORE1))
    aInsert = AddFlag(aInsert, "[STORE2]", addstring(loctable!STORE2))
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendTransHeader = sendTransHeader + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTransHeader = -1
GoTo lastsub
End Function
Private Function sendTrans() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, cFile As String
'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
cFile = "FILE1_60"
cString = "SELECT  FILE1_60.* from file1_60h inner join file1_60 ON FILE1_60.DOC_NO = file1_60h.doc_no where isdate >= " & DateSq(dPostDate)

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "ITEM", addvalue(loctable!Item))
    aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant & ""))
    aInsert = AddFlag(aInsert, "COST", Val(loctable!Quant & ""))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendTrans = sendTrans + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTrans = -1
GoTo lastsub
End Function

Private Function sendSTOCK() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, cFile As String
'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
cFile = "FILE0_10"
cString = "SELECT     FILE0_10.* FROM  FILE0_10 INNER JOIN file0_10h ON FILE0_10.DOC_NO = file0_10h.doc_no where isdate >= " & DateSq(dPostDate)
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[ITEM]", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "[RealBal]", Val(loctable!RealBal & ""))
    aInsert = AddFlag(aInsert, "[ComputerBal]", Val(loctable!ComputerBal & ""))
    aInsert = AddFlag(aInsert, "[Differ]", Val(loctable!Differ & ""))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendSTOCK = sendSTOCK + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendSTOCK = -1
GoTo lastsub
End Function

Private Function sendmosm() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
cFile = "mosm"
cString = "SELECT * FROM MOSM "

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
'    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendmosm = sendmosm + nAffect
Loop
lastsub:
prog1.Visible = False

'If lokaz Then
'    conmdb.Execute " update file0_00 set price = 1 "
'Else
'    conmdb.Execute " update file0_00 set price = 0 "
'End If

conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendmosm = -1
GoTo lastsub
End Function




Private Function sendbox()
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
cFile = "FILE0_50"
cString = "SELECT * FROM FILE0_50 "

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
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
    aInsert = AddFlag(aInsert, "[F_DATE]", addstring(Format(loctable!F_DATE, "dd-mm-yyyy")))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendbox = sendbox + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendbox = -1
GoTo lastsub
End Function


Private Function Sendfact() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

Dim cFile As String
cFile = "fact"
cString = "SELECT * FROM " & cFile
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

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
            
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    Sendfact = Sendfact + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
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
cString = "SELECT CODE ,DESCA   FROM branch  "
cString = cString & turn(cString) & " (CODE <> '00') "
cString = cString & " ORDER BY CODE"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With grid1
Do Until loctable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = loctable(0) & ""
    .TextMatrix(.Rows - 1, 1) = loctable!DESCA & ""
' .TextMatrix(.Rows - 1, 2) = retBranchData(loctable(0))
'   If IsDate(.TextMatrix(.Rows - 1, 2)) Then .TextMatrix(.Rows - 1, 3) = -1
    loctable.MoveNext
Loop
Fixgrd
End With
End Sub
Private Function ToMain() As Boolean
'On Error GoTo myerror
With grid1

For i = 1 To .Rows - 1
    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = &HE0E0E0
    If MsgBox("äÞá ãä ÝÑÚ " & .TextMatrix(i, 0), vbYesNo) = vbYes Then
        If Not getData(.TextMatrix(i, 0)) Then
            MsgBox "ãÔßáÉ ÇËäÇÁ äÞá ãáÝÇÊ ÝÑÚ : " & .TextMatrix(i, 1)
            Exit Function
        Else
            Inform "Êã äÓÎ ãáÝÇÊ ÝÑÚ : " & .TextMatrix(i, 1), " ÈäÌÇÍ", 1200
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
Private Function getData(sBranch As String) As Boolean
If Not validData(sBranch) Then
    MsgBox "ãáÝ äÞá ÇáÈíÇäÇÊ ÛíÑ ãæÌæÏ"
    Exit Function
End If
    
Dim nRecordCount As Long

'con.BeginTrans
If MsgBox("ÊÍÏíË ÇáãÈíÚÇÊ", vbYesNo) = vbYes Then
    nRecordCount = GetSales(sBranch)
    If nRecordCount >= 0 Then
        Inform "Êã ÓÍÈ " & nRecordCount & " ÓÌá ãä ÍÑßÇÊ ÇáÈíÚ", "ÈäÌÇÍ"
    Else
        GoTo myerror
    End If
End If

If MsgBox("ÊÍÏíË ÇáãÕÇÑíÝ ", vbYesNo) = vbYes Then
    nRecordCount = getCharge(sBranch, "FILE8_50")
    If nRecordCount >= 0 Then
        Inform "Êã ÓÍÈ " & nRecordCount & " ÓÌá ãä ÍÑßÇÊ ÇáãÕÇÑíÝ", "ÈäÌÇÍ"
    Else
        GoTo myerror
    End If
End If
'con.CommitTrans
getData = True
Exit Function
myerror:
'con.RollbackTrans
End Function
Private Function GetSales(sBranch As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & sBranch & ".mdb"

cFile = "FILE6_20"



Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile & "H"
cString = cString & turn(cString) & " Branch = " & MyParn(sBranch)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Do Until loctable.EOF
    con.Execute " delete from file6_20 where doc_no = " & MyParn(loctable!doc_no)
    con.Execute " delete from file6_20h where doc_no = " & MyParn(loctable!doc_no)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!store))
    aInsert = AddFlag(aInsert, "[NOTES]", addstring(loctable!NOTES))
    aInsert = AddFlag(aInsert, "[RATE]", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[CASH]", Val(loctable!CASH & ""))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[PAY]", Val(loctable!PAY & ""))
    aInsert = AddFlag(aInsert, "[VISA]", Val(loctable!Visa & ""))
    aInsert = AddFlag(aInsert, "[LATE]", Val(loctable!late & ""))
    aInsert = AddFlag(aInsert, "[REST]", Val(loctable!Rest & ""))
    aInsert = AddFlag(aInsert, "[MAN]", addstring(loctable!MAN))
    aInsert = AddFlag(aInsert, "[PRINTED]", IIf(loctable!PRINTED, "1", "0"))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
                                                                                
    con.Execute addInsert(aInsert, cFile & "H"), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
GetSales = nAffectTotal


Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0
Dim sitem As String
Dim aitem As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    sitem = loctable!Item
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "ITEM", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "MAN", addstring(loctable!MAN))
    aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant & ""))
    aInsert = AddFlag(aInsert, "PRICE", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "DISCOUNT", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "PRICE_C", Val(loctable!PRICE_C & ""))
    aInsert = AddFlag(aInsert, "ROW", Val(loctable!Row & ""))
    aInsert = AddFlag(aInsert, "COST", Val(loctable!COST & ""))
    con.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
Loop

Set loctable = New ADODB.Recordset
cString = "Select * FROM SUBCUST "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0


Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "PHONE", addstring(loctable!phone))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "E_MAIL", addstring(loctable!E_MAIL))
    aInsert = AddFlag(aInsert, "branch", addstring(loctable!branch))
    
    con.Execute addInsert(aInsert, "SUBCUST"), nAffect
    loctable.MoveNext
Loop

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
Private Function getCash(sBranch As String, cFile As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & sBranch & ".mdb"

cString = "Delete from " & cFile
cString = cString & turn(cString) & " DOC_NO IN (SELECT DOC_NO FROM " & cFile & "H" & " WHERE BRANCH = " & MyParn(sBranch) & ")"
con.Execute cString

cString = "Delete from " & cFile & "H"
cString = cString & turn(cString) & " Branch = " & MyParn(sBranch)
con.Execute cString


Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile & "H"
cString = cString & turn(cString) & " Branch = " & MyParn(sBranch)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
Dim aInsert As Variant
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranch))
                                            
    con.Execute addInsert(aInsert, cFile & "H"), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
getCash = nAffectTotal


Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

ReDim aInsert(6, 1)
prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[CODE]", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(loctable!Value & ""))
    aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "[ROW]", Val(loctable!Row & ""))
                    
    con.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
Loop

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getCash = -1
GoTo lastsub
End Function
Private Function getCharge(sBranch As String, cFile As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & sBranch & ".mdb"

cString = "Delete from file8_80 "
cString = cString & turn(cString) & " DOC_NO IN (SELECT DOC_NO FROM file8_80h WHERE BRANCH = " & MyParn(sBranch) & ")"
con.Execute cString

cString = "Delete from file8_80h "
cString = cString & turn(cString) & " Branch = " & MyParn(sBranch)
con.Execute cString

Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile & "H"
cString = cString & turn(cString) & " Branch = " & MyParn(sBranch)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    cDocno = sBranch & loctable!doc_no
    aInsert = AddFlag(Empty, "DOC_NO", addstring(cDocno))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranch))
                                        
    con.Execute addInsert(aInsert, "file8_80h"), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
getCharge = nAffectTotal


Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
nRecordCount = loctable.RecordCount
nRecord = 0
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    cDocno = sBranch & loctable!doc_no
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(cDocno))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[CHARGE]", addstring(loctable!CHARGE))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(loctable!Value & ""))
    aInsert = AddFlag(aInsert, "[ROW]", Val(loctable!Row & ""))
    
    con.Execute addInsert(aInsert, "file8_80"), nAffect
    loctable.MoveNext
Loop
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getCharge = -1
GoTo lastsub
End Function




Private Sub Fixgrd()
With grid1
.Cols = 2
.ColWidth(0) = 1200
.ColWidth(1) = 2500
.TextMatrix(0, 0) = "ßæÏ ÇáÝÑÚ"
.TextMatrix(0, 1) = "ÅÓã ÇáÝÑÚ"
.ColAlignment(0) = flexAlignRightCenter
.ColAlignment(1) = flexAlignRightCenter
End With
End Sub


Private Function validData(sBranch As String) As Boolean
Dim fs As New FileSystemObject
If Not fs.FileExists(cDataFolder & "\" & cDataFile & "_" & sBranch & ".mdb") Then Exit Function
validData = True
End Function
Private Function getStock(sBranch As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & sBranch & ".mdb"

cFile = "FILE0_10"
cString = "Delete from " & cFile
cString = cString & turn(cString) & " DOC_NO IN (SELECT DOC_NO FROM " & cFile & "H" & " WHERE BRANCH = " & MyParn(sBranch) & ")"
con.Execute cString

cString = "Delete from FILE0_100 "
cString = cString & turn(cString) & " DOC_NO IN (SELECT DOC_NO FROM FILE0_10H WHERE BRANCH = " & MyParn(sBranch) & ")"
con.Execute cString

cString = "Delete from " & cFile & "H"
cString = cString & turn(cString) & " Branch = " & MyParn(sBranch)
con.Execute cString


Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile & "H"
'cString = cString & turn(cString) & " Branch = " & MyParn(sBranch)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!store))
    aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, "1", "0"))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranch))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
                                                                                
    con.Execute addInsert(aInsert, cFile & "H"), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
getStock = nAffectTotal


'FILE0_10

Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0
Dim sitem As String
Dim aitem As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    sitem = loctable!Item
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[ITEM]", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "[RealBal]", Val(loctable!RealBal & ""))
    aInsert = AddFlag(aInsert, "[ComputerBal]", Val(loctable!ComputerBal & ""))
    aInsert = AddFlag(aInsert, "[Differ]", Val(loctable!Differ & ""))
    con.Execute addInsert(aInsert, "file0_10"), nAffect
    loctable.MoveNext
Loop


Set loctable = New ADODB.Recordset
cString = "Select * from FILE0_100"
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    sitem = loctable!Item
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[ITEM]", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "[RealBal]", Val(loctable!RealBal & ""))
    aInsert = AddFlag(aInsert, "[ComputerBal]", Val(loctable!ComputerBal & ""))
    aInsert = AddFlag(aInsert, "[Differ]", Val(loctable!Differ & ""))
    con.Execute addInsert(aInsert, "FILE0_100"), nAffect
    loctable.MoveNext
Loop

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getStock = -1
GoTo lastsub
End Function


Private Function sendStockHeader() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, cFile As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

cFile = "FILE0_10H"

cString = "SELECT  * FROM         file0_10h where isdate >= " & DateSq(dPostDate)
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addstring(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!store))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranchCode))
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendStockHeader = sendStockHeader + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendStockHeader = -1
GoTo lastsub
End Function

Private Function getpart(sBranch As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String

On Error GoTo myerror
cFile = "file8_70"
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & sBranch & ".mdb"

cString = "Delete from file8_70 "
cString = cString & turn(cString) & " DOC_NO IN (SELECT DOC_NO FROM " & cFile & "H" & " WHERE BRANCH = " & MyParn(sBranch) & ")"
con.Execute cString

cString = "Delete from " & cFile & "H"
cString = cString & turn(cString) & " Branch = " & MyParn(sBranch)
con.Execute cString

Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile & "H"
cString = cString & turn(cString) & " Branch = " & MyParn(sBranch)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranch))
                                        
    con.Execute addInsert(aInsert, cFile & "H"), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
getpart = nAffectTotal


Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile & " WHERE MID(DOC_NO,1,2) = " & MyParn(sBranch)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
nRecordCount = loctable.RecordCount
nRecord = 0
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[CODE]", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[VALUE_P]", Val(loctable!Value_P & ""))
    aInsert = AddFlag(aInsert, "[VALUE_M]", Val(loctable!Value_M & ""))
    aInsert = AddFlag(aInsert, "[ROW]", Val(loctable!Row & ""))
    
    con.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
Loop

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getpart = -1
GoTo lastsub
End Function
Private Function SendSupler() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

Dim cFile As String
cFile = "FILE4_10"
cString = "SELECT * FROM " & cFile
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

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
    aInsert = AddFlag(aInsert, "Manager", addstring(loctable!Manager))
    aInsert = AddFlag(aInsert, "Address", addstring(Mid(loctable!Address, 1, 200)))
    aInsert = AddFlag(aInsert, "phone1", addstring(loctable!PHONE1))
    aInsert = AddFlag(aInsert, "fax", addstring(Mid(loctable!Fax, 200)))
    aInsert = AddFlag(aInsert, "[Group]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "disc", addvalue(loctable!disc))
    aInsert = AddFlag(aInsert, "RATE", addvalue(loctable!Rate))
    aInsert = AddFlag(aInsert, "FACT", IIf(loctable!Fact, 1, 0))
    aInsert = AddFlag(aInsert, "SUPP", IIf(loctable!SUPP, 1, 0))
    aInsert = AddFlag(aInsert, "remark", addstring(loctable!remark))
    aInsert = AddFlag(aInsert, "SUBCODE", addstring(loctable!SUBCODE))
            
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    SendSupler = SendSupler + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
SendSupler = -1
Err.Clear
GoTo lastsub
End Function

Private Function SendDiscount() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset

'On Error GoTo myerror

'lokaz = IIf(cRet = "True", True, False)
'lokazItem = IIf(cRet = "True", True, False)

    conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

    If lokaz Then
        conmdb.Execute " update file0_00 set price = true "
    Else
        conmdb.Execute " update file0_00 set price = false "
    End If
    If lokazItem Then
        conmdb.Execute " update file0_00 set OKAZ_ITEM = true "
    Else
        conmdb.Execute " update file0_00 set OKAZ_ITEM = false "
    End If

cFile = "Discount"
cString = "SELECT * FROM Discount"

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "ROW", addvalue(loctable!Row))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[DATE1]", addDate(Format(loctable!Date1, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[DATE2]", addDate(Format(loctable!date2, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "DISC", Val(loctable!disc))
    aInsert = AddFlag(aInsert, "ISPRICE_2", IIf(loctable!ISPRICE_2, 1, 0))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    SendDiscount = SendDiscount + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendDiscount = -1
GoTo lastsub
End Function

