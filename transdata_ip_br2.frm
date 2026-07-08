VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form transdata_ip_br2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕœÌÀ Ê ”Õ» »Ì«‰«  «· ÊﬂÌ·« "
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   6825
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
   ScaleHeight     =   3630
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CMD_IP 
      Caption         =   "Ÿ»ÿ —ﬁ„ IP «·”Ì—›—"
      Height          =   405
      Left            =   585
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Visible         =   0   'False
      Width           =   2625
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   " ÕœÌÀ ﬂ· «·»Ì«‰« "
      Height          =   375
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«—”«· «·„»Ì⁄«  ··”Ì—›—"
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
      Left            =   585
      TabIndex        =   4
      Top             =   1350
      Visible         =   0   'False
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
      Height          =   780
      Left            =   585
      TabIndex        =   3
      Top             =   2160
      Width           =   5730
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "”Õ» »Ì«‰«  „‰ «·”Ì—›—"
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
      Left            =   585
      TabIndex        =   0
      Top             =   540
      Width           =   5730
   End
   Begin VB.Frame Frame11 
      Height          =   600
      Left            =   585
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2835
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
End
Attribute VB_Name = "transdata_ip_br2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lAddAll As Boolean
Dim con As New ADODB.Connection, pBranch As String
Dim conShop As New ADODB.Connection
Dim cFilesave As String
Dim cDataFolder As String, cDataFile As String
Dim dPostDate As String
Dim nBranchPost As Double

Private Function ToBranch() As Boolean
'On Error GoTo myerror
Dim lMyPrice As Boolean, lMyPriceItem As Boolean

'FixItemDel
Inform "»Ì«‰«  «·«’‰«›"

nRecordCount = sendItems
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror
'sendSUB_USER
GetGrDiscout

nRecordCount = SendGroup
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ „Ã„Ê⁄«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = SendSection
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ «·√ﬁ”«„ ", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = Sendfact
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„’«‰⁄", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = sendCode("FILE4_10")
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„Ê—œÌ‰ ", "»‰Ã«Õ" Else GoTo myerror
'con.Execute " update file4_10 set [group] = '01'"

nRecordCount = sendCode("FILE0_81", True)
nRecordCount = sendCode("FILE0_82", True)
nRecordCount = sendCode("SHIP", True)

''''''''''''''''''''''''
'''''''If Check1.Value <> 0 Then
'''''''    nRecordCount = SendPurchaseHeader("SELECT * FROM FR7_20H WHERE STORE = " & MyParn(cBranchStore), "FR7_20h", "FR7_20", "FILE7_20H", "FILE7_20")
'''''''    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„‘ —Ì«  ", "»‰Ã«Õ" Else GoTo myerror
'''''''Else
'''''''    nRecordCount = SendPurchaseHeader("SELECT * FROM FR7_20H WHERE ISNEW1 = 1 AND STORE = " & MyParn(cBranchStore), "FR7_20h", "FR7_20", "FILE7_20H", "FILE7_20")
'''''''    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„‘ —Ì«  ", "»‰Ã«Õ" Else GoTo myerror
'''''''End If
'''''''
''''''' If Check1.Value <> 0 Then
'''''''    nRecordCount = SendPurchaseHeader("SELECT * FROM FR7_10H WHERE STORE = " & MyParn(cBranchStore), "FR7_10h", "FR7_10", "FILE7_10H", "FILE7_10")
'''''''    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„‘ —Ì«  ", "»‰Ã«Õ" Else GoTo myerror
'''''''Else
'''''''    nRecordCount = SendPurchaseHeader("SELECT * FROM FR7_10H WHERE ISNEW1 = 1 AND  STORE = " & MyParn(cBranchStore), "FR7_10h", "FR7_10", "FILE7_10H", "FILE7_10")
'''''''    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„‘ —Ì«  ", "»‰Ã«Õ" Else GoTo myerror
'''''''End If
'''''''
'''''''nRecordCount = sendmosm
'''''''If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„Ê«”„ ", "»‰Ã«Õ" Else GoTo myerror
'''''''
'''''''nRecordCount = SendDiscount
'''''''If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ «·Œ’Ê„«  ", "»‰Ã«Õ" Else GoTo myerror
'''''''
'''''''nRecordCount = NewPrice
'''''''If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ «”⁄«—  „  ⁄œÌ·Â«", "»‰Ã«Õ" Else GoTo myerror
'''''''
'''''''lMyPrice = GetBoolean("SELECT ISOKAZ FROM BRANCH WHERE CODE = " & MyParn(pBranch), conShop)
'''''''lMyPriceItem = GetBoolean("SELECT ISOKAZMODEL FROM BRANCH WHERE CODE = " & MyParn(pBranch), conShop)
'''''''
'''''''If lMyPrice Then
'''''''    con.Execute " update FILE0_00 SET PRICE = 1 "
'''''''    Inform " «”⁄«— «Êﬂ«“ÊÌ‰"
'''''''Else
'''''''    con.Execute " update FILE0_00 SET PRICE = 0 "
'''''''End If
'''''''
'''''''If lMyPriceItem Then
'''''''    con.Execute " update FILE0_00 SET OKAZ_ITEM = 1 "
'''''''    Inform " «”⁄«— ⁄—÷ ·› —… "
'''''''Else
'''''''    con.Execute " update FILE0_00 SET OKAZ_ITEM = 0 "
'''''''End If

con.Execute " DELETE FROM MOSM      WHERE MOSM NOT IN (SELECT MOSM FROM FILE1_10)"
con.Execute " DELETE FROM FACT      WHERE CODE NOT IN (SELECT FACT FROM FILE1_10)"

ToBranch = True
'MsgBox " „ ‰ﬁ· «·»Ì«‰« "
ToBranch = True
Exit Function
ToBranch = False
myerror:
'MsgBox "·„ Ì „ ‰ﬁ· ﬂ«›… «·»Ì«‰«  »‰Ã«Õ"
Err.Clear
End Function

Private Sub cmd_allitem_Click()
If MsgBox(" „  ÕœÌÀ ﬂ· «·«’‰«› »«·ﬂ«„· ··›—Ê⁄", vbYesNo + vbDefaultButton2) = vbYes Then
    con.Execute " update file1_10 set isnew1 = 0 "
End If
End Sub

Private Sub CMD_IP_Click()
'    cIp = InputBox(" ", " ", GetDesca("SELECT IPSERVER FROM BRANCH"))
'    con.Execute " UPDATE BRANCH SET IPSERVER = " & addstring(COP)
'    Inform " „  ⁄œÌ· —ﬁ„ IP «·”Ì—›— "
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
'   On Error GoTo myerror
    Dim pServer As String
    Dim pData As String
'   pServer = GetDesca("SELECT IPSERVER FROM BRANCH")
'   pData = GetDesca("SELECT DATASERVER FROM BRANCH")
    Inform "”Õ» »Ì«‰«  „‰ «·”Ì—›—"
    strConShop = LoadConStringshop
    nBranchPost = Val(cBranch)
    pBranch = cBranch
'   Inform strConShop
    If openConShop(conShop) = "ok" Then
        Inform " „ «·≈ ’«· »‰Ã«Õ »«·”Ì—›—"
        If Not ToBranch() Then
            MsgBox "·„ Ì „ ”Õ» «·»Ì«‰«  „‰ «·”Ì—›—"
            If conShop.State = adStateOpen Then closeCon conShop
        End If
    Else
        MsgBox "›‘· «·≈ ’«·"
    End If
    MsgBox "  „ ”Õ» «·»Ì«„«  "
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Command1_Click()
'    Dim pServer As String
'    Dim pData As String
'    pServer = GetDesca("SELECT IPSERVER FROM BRANCH")
'    pData = GetDesca("SELECT DATASERVER FROM BRANCH")
'    Me.Caption = "«—”«· «·„»Ì⁄«  ··”Ì—›—"
'    strConShop = LoadConStringshop
'    nBranchPost = Val(cBranch)
'    pBranch = cBranch
'    If openConShop(conShop) = "ok" Then
'        Inform " „ «·≈ ’«· »‰Ã«Õ »«·”Ì—›—"
'        If Not ToMain() Then
'            MsgBox "·„ Ì „ ”Õ» «·»Ì«‰«  „‰ «·”Ì—›—"
'            If conShop.State = adStateOpen Then closeCon conShop
'        End If
'    End If
End Sub
Private Sub Form_Load()
    
    Check1.Visible = bSupermode
    CMD_IP.Visible = bSupermode
    cDataFolder = App.Path & "\mdb"
    cDataFile = "data"
    openCon con
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
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
con.Execute " DELETE FROM FILE1_50 "

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
            
    con.Execute addInsert(aInsert, cFile), nAffect
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
Dim cFile As String, nCount As Double
cFile = "FILE1_10"

If Check1.Value <> 0 Then
    cString = "SELECT * FROM FILE1_10 WHERE  ITEM IN (SELECT ITEM FROM FR1_11 WHERE     FR1_11.STORE = " & MyParn(pBranch) & ")  "
Else
    cString = "SELECT * FROM FILE1_10 WHERE  ITEM IN (SELECT ITEM FROM FR1_11 WHERE     FR1_11.STORE = " & MyParn(pBranch) & ") AND ITEM IN ( SELECT ITEM FROM NEW_ITEM_FR WHERE BRANCH = " & MyParn(pBranch) & " ) "
End If
Dim cStrall As String, cStrDEl As String
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
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
    
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!MODEL))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable!MODELNO))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
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
    aInsert = AddFlag(aInsert, "[COSTITEM]", Val(loctable!costITEM & ""))
    aInsert = AddFlag(aInsert, "[COST2]", Val(loctable!cost2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE]", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "[PRICE_2]", Val(loctable!PRICE_2 & ""))
    aInsert = AddFlag(aInsert, "[SCAL]", addstring(loctable!SCAL))
    aInsert = AddFlag(aInsert, "[C_SCAL]", addvalue(loctable!C_SCAL))
    aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable!Color))
    aInsert = AddFlag(aInsert, "[C_COLOR]", addvalue(loctable!c_Color))
    aInsert = AddFlag(aInsert, "[BARCODE]", addstring(loctable!BARCODE))
    aInsert = AddFlag(aInsert, "[BARCODE2]", addstring(loctable!BARCODE2))
    aInsert = AddFlag(aInsert, "[BARCODE13]", addstring(loctable!BARCODE13))
    aInsert = AddFlag(aInsert, "[ISOKAZITEM]", IIf(loctable!ISOKAZITEM, 1, 0))
    aInsert = AddFlag(aInsert, "[ISNOITEM]", IIf(loctable!ISNOITEM, 1, 0))
    aInsert = AddFlag(aInsert, "[ISNODEM]", IIf(loctable!ISNODEM, 1, 0))
    aInsert = AddFlag(aInsert, "[MOSM2]", addstring(loctable!MOSM2))
    aInsert = AddFlag(aInsert, "[PRICE2]", Val(loctable!PRICE2 & ""))
    aInsert = AddFlag(aInsert, "[COST]", Val(loctable!cost & ""))
    aInsert = AddFlag(aInsert, "MODELSEX", addvalue(loctable!MODELSEX))
    aInsert = AddFlag(aInsert, "MODELAGE", addvalue(loctable!MODELAGE))
    aInsert = AddFlag(aInsert, "[ISSTOP]", IIf(loctable!ISSTOP, 1, 0))
    
    
'   con.Execute addInsert(aInsert, "FILE1_10"), nAffect
    Me.Caption = loctable!Item
    nCount = nCount + 1
    cStrall = cStrall & " DELETE FROM FILE1_10 where item = " & loctable!Item & " ; "
    cStrDEl = cStrDEl & " DELETE FROM NEW_ITEM_FR WHERE ITEM = " & loctable!Item & " AND BRANCH = " & MyParn(cBranch) & " ; "
    cStrall = cStrall & addInsert(aInsert, "FILE1_10") & " ; "
    If nCount = 100 Then
        con.Execute cStrall
        conShop.Execute cStrDEl
        nCount = 0
    End If
    
    loctable.MoveNext
    sendItems = sendItems + 1
Loop
If cStrall <> "" Then
    con.Execute cStrall
    conShop.Execute cStrDEl
End If

lastsub:
Me.Caption = sCaption
prog1.Visible = False
sendItems = nRecord
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
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
con.Execute " delete from " & cFile
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100

    If isNumber Then
        aInsert = AddFlag(Empty, "CODE", addvalue(loctable!CODE))
    Else
        aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    End If
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
    con.Execute addInsert(aInsert, cFile), nAffect
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
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
con.Execute " delete from file0_40"
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
'    aInsert = AddFlag(aInsert, "SERVER", addstring(loctable!SERVER))
'    aInsert = AddFlag(aInsert, "USERNAME", addstring(loctable!UserName))
'    aInsert = AddFlag(aInsert, "PASSWORD", addstring(loctable!PassWord))
    
    con.Execute addInsert(aInsert, cFile), nAffect
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

Private Function sendvisa()
'On Error GoTo myerror
cFile = "VISA"
cString = "SELECT * FROM VISA WHERE BRANCH = " & MyParn(pBranch)

Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
con.Execute " delete from visa "
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "BRANCH", addstring(loctable!branch))
    con.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendvisa = sendvisa + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendvisa = -1
GoTo lastsub
End Function

Private Function sendClients() As Long
On Error GoTo myerror
Dim cFile As String
cFile = "FILE3_10"
cString = "SELECT * FROM " & cFile
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
con.Execute " delete from file3_10"
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
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
           
    con.Execute addInsert(aInsert, cFile), nAffect
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
Private Function SendPurchaseHeader(cString, cFile, cFile2, cFileBranch, cFileBranch2) As Long
'On Error GoTo myerror
Dim loctable As New ADODB.Recordset
Dim loctable2 As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    
    con.Execute " delete from " & cFileBranch2 & " where doc_no = " & MyParn(loctable!doc_no)
    con.Execute " delete from " & cFileBranch & " where doc_no = " & MyParn(loctable!doc_no)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!STORE = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
        aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
        aInsert = AddFlag(aInsert, "[TAX]", Val(loctable!tax & ""))
        aInsert = AddFlag(aInsert, "[r_tax]", Val(loctable!R_TAX & ""))
        aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM))
        aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
        con.Execute addInsert(aInsert, cFileBranch), nAffect
    
        Dim aInsert2 As Variant
        If loctable2.State = adStateOpen Then loctable2.Close
        loctable2.Open "SELECT * FROM " & cFile2 & " WHERE DOC_NO = " & MyParn(loctable!doc_no), conShop, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not loctable2.EOF
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
            aInsert2 = AddFlag(aInsert2, "DISCOUNT", Val(loctable2!discount & ""))
            aInsert2 = AddFlag(aInsert2, "r_tax", Val(loctable2!R_TAX & ""))
            con.Execute addInsert(aInsert2, cFileBranch2), nAffect
            loctable2.MoveNext
        Loop
    End If
    conShop.Execute " UPDATE " & cFile & " SET ISNEW1 = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
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
'con.Execute " update " & cFile & "H  SET ISNEW" & nBranchPost & "  = 0 "
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
If Check1.Value <> 0 Then
    cString = "SELECT  FILE0_40.branch AS branch_1 ,FILE0_40_1.branch AS BRANCH , file1_60h.* FROM         file1_60h INNER JOIN   FILE0_40 ON file1_60h.store1 = FILE0_40.CODE INNER JOIN FILE0_40 AS FILE0_40_1 ON file1_60h.store2 = FILE0_40_1.CODE where ISNOITEM = 0 AND  ( store1 = " & MyParn(cBranchStore) & " OR STORE2 = " & MyParn(cBranchStore) & " ) "
Else
    cString = "SELECT  FILE0_40.branch AS branch_1 ,FILE0_40_1.branch AS BRANCH , file1_60h.* FROM         file1_60h INNER JOIN   FILE0_40 ON file1_60h.store1 = FILE0_40.CODE INNER JOIN FILE0_40 AS FILE0_40_1 ON file1_60h.store2 = FILE0_40_1.CODE where ISNOITEM = 0 AND DOC_NO IN (SELECT DOC_NO FROM NEW1_60H WHERE BRANCH = " & MyParn(pBranch) & " ) "
End If
Dim loctable As New ADODB.Recordset
Dim loctable2 As New ADODB.Recordset

loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    con.Execute " delete from file1_60  where doc_no = " & MyParn(loctable!doc_no)
    con.Execute " delete from file1_60h where doc_no = " & MyParn(loctable!doc_no)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!BRANCH_1 = pBranch Or loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
        aInsert = AddFlag(aInsert, "[STORE1]", addstring(loctable!STORE1))
        aInsert = AddFlag(aInsert, "[STORE2]", addstring(loctable!STORE2))
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
        con.Execute addInsert(aInsert, cFile), nAffect
        
        Dim aInsert2 As Variant
        If loctable2.State = adStateOpen Then loctable2.Close
        loctable2.Open "SELECT FILE1_60.* from file1_60 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_60.ITEM where FILE1_10.ISNOITEM = 0 AND doc_no = " & MyParn(loctable!doc_no), conShop, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not loctable2.EOF
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
            aInsert2 = AddFlag(aInsert2, "COST", Val(loctable2!cost & ""))
            con.Execute addInsert(aInsert2, "file1_60"), nAffect
            loctable2.MoveNext
        Loop
    End If
    conShop.Execute " DELETE FROM NEW1_60H WHERE DOC_NO = " & MyParn(loctable!doc_no) & " AND BRANCH = " & MyParn(pBranch)
    loctable.MoveNext
    sendTransHeader = sendTransHeader + nAffect
Loop
conShop.Execute " DELETE FROM NEW1_60H WHERE BRANCH = " & MyParn(pBranch)
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTransHeader = -1
GoTo lastsub
End Function

Private Function sendmosm() As Long
'On Error GoTo myerror
cFile = "mosm"
cString = "SELECT * FROM MOSM "

Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
con.Execute " delete from mosm "
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[closed]", IIf(loctable!CLOSED, 1, 0))
    con.Execute addInsert(aInsert, cFile), nAffect
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
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
con.Execute " delete from file0_50"
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[F_DATE]", addDate(loctable!F_DATE))
    aInsert = AddFlag(aInsert, "[ISVISA]", IIf(loctable!ISVISA, 1, 0))
    aInsert = AddFlag(aInsert, "[BRANCH]", addDate(loctable!branch))
    aInsert = AddFlag(aInsert, "[type]", addstring(loctable!Type))
    
    con.Execute addInsert(aInsert, cFile), nAffect
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
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
con.Execute " delete from fact "
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
            
    con.Execute addInsert(aInsert, cFile), nAffect
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
Private Function ToMain() As Boolean
'On Error GoTo myerror
    If Not getData(cBranch) Then
        MsgBox "„‘ﬂ·… «À‰«¡ «—”«· «·„»Ì⁄«  ··”Ì—›— "
        Exit Function
    Else
        Inform " „ «—”«· «·„»Ì⁄«  ··”Ì—›— : ", 1200
    End If
ToMain = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function getData(sBranch As String) As Boolean
Dim nRecordCount As Long
    Inform "„»Ì⁄« "
    nRecordCount = GetSales(sBranch)
    If nRecordCount >= 0 Then
        Inform " „ «—”«· " & nRecordCount & " ”Ã· „‰ Õ—ﬂ«  «·»Ì⁄", "»‰Ã«Õ"
    Else
        GoTo myerror
    End If

    Inform "«·Ã—œ"
    nRecordCount = getStock(sBranch)
    If nRecordCount >= 0 Then
        Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ „” ‰œ«  «·Ã—œ ", "»‰Ã«Õ"
    Else
        GoTo myerror
    End If

    Me.Caption = "≈” ·«„ „” ‰œ« "
    nRecordCount = Get_Purch_R("FR7_20H")
    nRecordCount = Get_Purch_R("FILE7_10H")
    nRecordCount = Get_Purch_R("FILE1_60H")


    Inform " ÕÊÌ·«  ‰ﬁœÌ…"
    nRecordCount = GetTransBox(sBranch)
    Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰  ÕÊÌ·«  ‰ﬁœÌ… ", "»‰Ã«Õ"

    nRecordCount = getCharge(cBranch)
    If nRecordCount >= 0 Then
        Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ Õ—ﬂ«  «·≈Ì—«œ«  ", "»‰Ã«Õ"
    Else
        GoTo myerror
    End If

getData = True
Exit Function
myerror:
End Function
Private Function GetSales(sBranch As String) As Long

cFile = "FILE6_20"
Set loctable = New ADODB.Recordset

If Check1.Value <> 0 Then
    cString = "Select * from FILE6_20H where PRINTED = 1 "
Else
    cString = "Select * from FILE6_20H WHERE PRINTED =1 and isnew = 1"
End If
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
    conShop.Execute " delete from file6_20 where doc_no = " & MyParn(loctable!doc_no)
    conShop.Execute " delete from file6_20h where doc_no = " & MyParn(loctable!doc_no)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[TIME]", addTime(Format(loctable!Time, "SHORT TIME")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "[NOTES]", addstring(loctable!NOTES))
    aInsert = AddFlag(aInsert, "[RATE]", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[CASH]", Val(loctable!CASH & ""))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[BOXVISA]", addstring(loctable!BOXVISA))
    aInsert = AddFlag(aInsert, "[PAY]", Val(loctable!PAY & ""))
    aInsert = AddFlag(aInsert, "[VISA]", Val(loctable!Visa & ""))
    aInsert = AddFlag(aInsert, "[LATE]", Val(loctable!late & ""))
    aInsert = AddFlag(aInsert, "[REST]", Val(loctable!Rest & ""))
    aInsert = AddFlag(aInsert, "[MAN]", addstring(loctable!MAN))
    aInsert = AddFlag(aInsert, "[PRINTED]", IIf(loctable!PRINTED, "1", "0"))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "[phone]", addstring(loctable!phone))
                                                                                
    aInsert = AddFlag(aInsert, "[username_disc]", addstring(loctable!username_disc))
    aInsert = AddFlag(aInsert, "[username_ret]", addstring(loctable!username_ret))
    aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(loctable!SALES_RET))
    aInsert = AddFlag(aInsert, "CODEVISA", addvalue(loctable!CODEVISA))
    aInsert = AddFlag(aInsert, "[NOTE_DISC]", addstring(loctable!NOTE_DISC))
                                                                                
    aInsert = AddFlag(aInsert, "CHARGE1", Val(loctable!CHARGE1 & ""))
    aInsert = AddFlag(aInsert, "CHARGE2", Val(loctable!charge2 & ""))
    aInsert = AddFlag(aInsert, "[NAME]", addstring(loctable!Name))
    aInsert = AddFlag(aInsert, "[SHIP]", addstring(loctable!SHIP))
    aInsert = AddFlag(aInsert, "[ADDRESS]", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "[SHIP_NO]", addstring(loctable!SHIP_NO))
    aInsert = AddFlag(aInsert, "[DATE_PAY]", addDate(loctable!DATE_PAY))
    aInsert = AddFlag(aInsert, "[BANK]", addstring(loctable!BANK))
    
    conShop.Execute addInsert(aInsert, cFile & "H"), nAffect
    
    
    Set loctable2 = New ADODB.Recordset
    cString = "Select FILE6_20.* from FILE6_20 WHERE FILE6_20.doc_no = " & MyParn(loctable!doc_no)
    loctable2.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    Dim sitem As String
    Dim aitem As Variant
    Do Until loctable2.EOF
        sitem = loctable2!Item
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
        aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
        aInsert2 = AddFlag(aInsert2, "MAN", addstring(loctable2!MAN))
        aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
        aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
        aInsert2 = AddFlag(aInsert2, "DISCOUNT", Val(loctable2!discount & ""))
        aInsert2 = AddFlag(aInsert2, "PRICE_C", Val(loctable2!PRICE_C & ""))
        aInsert2 = AddFlag(aInsert2, "ROW", Val(loctable2!Row & ""))
        aInsert2 = AddFlag(aInsert2, "COST", Val(loctable2!cost & ""))
        aInsert2 = AddFlag(aInsert2, "S_OKAZ", Val(loctable2!S_OKAZ & ""))
        aInsert2 = AddFlag(aInsert2, "ISDISC2_1", IIf(loctable2!ISDISC2_1, 1, 0))
        conShop.Execute addInsert(aInsert2, "file6_20"), nAffect
        loctable2.MoveNext
    Loop

    con.Execute " update file6_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
GetSales = nAffectTotal


Inform "”Õ» «· Ì·›Ê‰« "
Set loctable = New ADODB.Recordset
cString = "Select * from SUBCUST WHERE isnew = 1 "
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0


Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "PHONE", addstring(loctable!phone))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "E_MAIL", addstring(loctable!E_MAIL))
    aInsert = AddFlag(aInsert, "branch", addstring(loctable!branch))
    
    conShop.Execute addInsert(aInsert, "SUBCUST"), nAffect
    loctable.MoveNext
Loop
con.Execute " UPDATE SUBCUST SET isnew = 0 "
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
.Cols = 6
.ColWidth(0) = 1000
.ColWidth(1) = 1500
.ColWidth(2) = 2000
.ColWidth(3) = 0
.ColWidth(4) = 600
.ColWidth(5) = 600
.TextMatrix(0, 0) = "ﬂÊœ «·›—⁄"
.TextMatrix(0, 1) = "≈”„ «·›—⁄"
.TextMatrix(0, 2) = "IP SQL"
.TextMatrix(0, 4) = "„ ’·"
.TextMatrix(0, 5) = "«·ﬂ·"

.ColDataType(4) = flexDTBoolean
.ColDataType(5) = flexDTBoolean

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
        aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
        aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!Desca))
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
Private Function Get_Purch_R(pFile) As Long
Set loctable = New ADODB.Recordset

cString = "Select * from " & pFile & " WHERE Received = 1 "
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
    conShop.Execute " UPDATE " & pFile & " SET Received = 1 , TIMEACCEPT = " & addDate(loctable!TimeAccept) & " , USERACCEPT = " & addstring(loctable!UserAccept) & " where doc_no = " & MyParn(loctable!doc_no)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
Get_Purch_R = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
Get_Purch_R = -1
GoTo lastsub
End Function
Private Function GetTransBox(pBranch) As Long
Set loctable = New ADODB.Recordset
Dim cDoc As String

If Check1.Value <> 0 Then
    cString = "Select * from FILE0_51 "
Else
    cString = "Select * from FILE0_51 WHERE isnew = 1"
End If
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Do Until loctable.EOF
    cDoc = pBranch & loctable!CODE
    conShop.Execute " delete from file0_52  where CODE = " & MyParn(cDoc)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(cDoc))
    aInsert = AddFlag(aInsert, "NO1", addstring(loctable!no1))
    aInsert = AddFlag(aInsert, "NO2", addstring(loctable!no2))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(loctable!Value))
                                                                                
    conShop.Execute addInsert(aInsert, "FILE0_52"), nAffect
    con.Execute " update FILE0_51 set isnew = 0 where CODE = " & MyParn(loctable!CODE)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
GetTransBox = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GetTransBox = -1
GoTo lastsub
End Function
Private Function getStock(pBranch) As Long
cFile = "FILE0_10H"

cString = "SELECT  FILE0_10H.* FROM  file0_10h where isnew  = 1 AND BRANCH = " & MyParn(pBranch)
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    conShop.Execute " delete from file0_10 where SUBSTRING(DOC_NO,1,2) = " & MyParn(cBranch) & " AND  doc_no = " & MyParn(loctable!doc_no)
    conShop.Execute " delete from file0_100 where SUBSTRING(DOC_NO,1,2) = " & MyParn(cBranch) & " AND doc_no = " & MyParn(loctable!doc_no)
    conShop.Execute " delete from file0_10h where SUBSTRING(DOC_NO,1,2) = " & MyParn(cBranch) & " AND doc_no = " & MyParn(loctable!doc_no)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
'    aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
    conShop.Execute addInsert(aInsert, cFile), nAffect
    
    Dim loctable2 As New ADODB.Recordset
    loctable2.Open "SELECT * FROM FILE0_10 WHERE DOC_NO = " & MyParn(loctable!doc_no), con, adOpenStatic, adLockReadOnly, adCmdText
    Do Until loctable2.EOF
        Me.Caption = loctable2!Item
        If GetDesca("select item from a1_10 where item = " & loctable2!Item, conShop) = "" Then
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "[ITEM]", addstring(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "[RealBal]", Val(loctable2!RealBal & ""))
            aInsert2 = AddFlag(aInsert2, "[ComputerBal]", Val(loctable2!ComputerBal & ""))
            aInsert2 = AddFlag(aInsert2, "[Differ]", Val(loctable2!Differ & ""))
            conShop.Execute addInsert(aInsert2, "file0_10"), nAffect
            conShop.Execute addInsert(aInsert2, "FILE0_100"), nAffect
            Me.Caption = loctable2!doc_no & " " & loctable2!Item
        End If
        loctable2.MoveNext
    Loop
    con.Execute " UPDATE FILE0_10H SET ISNEW = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
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

con.Execute " DELETE FROM FILE6_25 "
cFile = "FILE6_25"
cString = "SELECT * FROM FILE6_25 "
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If IsNull(loctable!branch) Or loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
        aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
        aInsert = AddFlag(aInsert, "BRANCH", addstring(loctable!branch))
        aInsert = AddFlag(aInsert, "ID_EMP", addstring(loctable!ID_EMP))
                
        con.Execute addInsert(aInsert, cFile), nAffect
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
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    con.Execute " delete from NEWPRICEITEM where model = " & MyParn(loctable!MODEL)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!MODEL))
    con.Execute addInsert(aInsert, "NEWPRICEITEM"), nAffect
    conShop.Execute " update    NEWPRICEITEM set    ISNEW" & nBranchPost & "  = 0 where model = " & MyParn(loctable!MODEL)
    conShop.Execute " delete    from NEWPRICEITEM  where ISNEW1 = 0 and ISNEW2 = 0 and ISNEW3 = 0 and ISNEW4 = 0 and ISNEW5 = 0 and model = " & MyParn(loctable!MODEL)
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
Me.Caption = "Œ’„"
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
con.Execute " delete from DISCOUNT "
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "ROW", addstring(loctable!Row))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[DATE1]", addDate(Format(loctable!Date1, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[DATE2]", addDate(Format(loctable!date2, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[DISC]", Val(loctable!disc & ""))
    aInsert = AddFlag(aInsert, "[MAXDISC]", Val(loctable!MaxDisc & ""))
    aInsert = AddFlag(aInsert, "[ISPRICE_2]", IIf(loctable!ISPRICE_2, 1, 0))
    aInsert = AddFlag(aInsert, "[ISDISC2]", IIf(loctable!ISdisc2, 1, 0))
    
    con.Execute addInsert(aInsert, "DISCOUNT"), nAffect
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

Private Function getCharge(sBranch As String) As Long
cFile = "FILE8_60"
Dim loctable   As New ADODB.Recordset
Dim loctable2  As New ADODB.Recordset
cString = "Select * from FILE8_60H WHERE ISNEW = 1 "
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount
    
Dim aInsert As Variant
Do Until loctable.EOF
    
    conShop.Execute " delete from file8_60 where DOC_NO = " & MyParn(loctable!doc_no) & " AND DOC_NO IN ( SELECT DOC_NO FROM FILE8_60H WHERE BRANCH = " & MyParn(sBranch) & " ) "
    conShop.Execute " delete from file8_60h where DOC_NO = " & MyParn(loctable!doc_no) & " AND BRANCH = " & MyParn(sBranch)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(cBranch))
    aInsert = AddFlag(aInsert, "ISCLOSED", 1)
    conShop.Execute addInsert(aInsert, cFile & "H"), nAffect

'''''''''''''
    Dim aInsert2 As Variant
    cString = "Select * FROM FILE8_60 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    Do Until loctable2.EOF
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
        aInsert2 = AddFlag(aInsert2, "[BOX]", addstring(loctable2!BOX))
        aInsert2 = AddFlag(aInsert2, "[CHARGE]", addstring(loctable2!CHARGE))
        aInsert2 = AddFlag(aInsert2, "[DESCA]", addstring(loctable2!Desca))
        aInsert2 = AddFlag(aInsert2, "[BRANCH]", addstring(loctable2!branch))
        aInsert2 = AddFlag(aInsert2, "[VALUE]", Val(loctable2!Value & ""))
        aInsert2 = AddFlag(aInsert2, "[ROW]", Val(loctable2!Row & ""))
        conShop.Execute addInsert(aInsert2, "FILE8_60"), nAffect
        loctable2.MoveNext
    Loop
'''''''''''''
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
getCharge = nAffectTotal
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

Private Function SendSection() As Long
'On Error GoTo myerror

Dim cFile As String
Dim loctable As New ADODB.Recordset

con.Execute " DELETE FROM FILE1_10SC"
cFile = "FILE1_10SC"
cString = "SELECT * FROM FILE1_10SC"
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "ISDISC", IIf(loctable!ISDISC, 1, 0))
    con.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    SendSection = SendSection + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
SendSection = -1
Err.Clear
GoTo lastsub
End Function


Private Function FixItemDel() As Long
'On Error GoTo myerror
cString = "SELECT * FROM BARCODE_EDIT "
Me.Caption = " ⁄œÌ· »«—ﬂÊœ"
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    con.Execute " UPDATE FILE6_20 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
    con.Execute " UPDATE FILE7_10 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
    con.Execute " UPDATE FILE7_20 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
    'con.Execute " UPDATE FILE1_60 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
    'con.Execute " UPDATE FILE1_61 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
    con.Execute " DELETE FROM FILE1_10 WHERE ITEM = " & (loctable!ITEMOLD)
    
    loctable.MoveNext
    FixItemDel = FixItemDel + nAffect
Loop
lastsub:
prog1.Visible = False

Exit Function
myerror:
MsgBox Err.Description
Err.Clear
FixItemDel = -1
GoTo lastsub
End Function




Private Function sendSUB_USER()
'On Error GoTo myerror
cString = "SELECT * FROM SUB_USER "

Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
con.Execute " delete from SUB_USER"
Dim aInsert As Variant
cStr1 = ""
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "[PassWord]", addstring(loctable!PassWord))
    cStr1 = cStr1 & addInsert(aInsert, "SUB_USER") & " ; "
    loctable.MoveNext
Loop
If cStr1 <> "" Then con.Execute cStr1
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendSUB_USER = -1
GoTo lastsub
End Function

Private Function GetGrDiscout() As Long
'On Error GoTo myerror
Dim cFile As String
cString = "SELECT * FROM FILE0_90H "
Dim loctable As New ADODB.Recordset, cStr1 As String
Dim loctable2 As New ADODB.Recordset, nCountR As Double
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!Desca))
    aInsert = AddFlag(aInsert, "DATE1", addDate(loctable!Date1))
    aInsert = AddFlag(aInsert, "DATE2", addDate(loctable!date2))
    aInsert = AddFlag(aInsert, "DATE", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "QTY1", Val(loctable!QTY1))
    aInsert = AddFlag(aInsert, "QTY2", Val(loctable!QTY2))
    aInsert = AddFlag(aInsert, "DISC", Val(loctable!disc))
    aInsert = AddFlag(aInsert, "ISSTOP", IIf(loctable!ISSTOP, 1, 0))
    
    con.Execute " DELETE FROM  FILE0_90  WHERE DOC_NO = " & MyParn(loctable!doc_no)
    con.Execute " DELETE FROM  FILE0_90H WHERE DOC_NO = " & MyParn(loctable!doc_no)
    con.Execute addInsert(aInsert, "FILE0_90H")
    
    
    cStr2 = ""
    Dim aInsert2 As Variant
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open "SELECT * FROM FILE0_90 WHERE DOC_NO = " & MyParn(loctable!doc_no), conShop, adOpenStatic, adLockReadOnly, adCmdText
    nCountR = 0
    Do While Not loctable2.EOF
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
        aInsert2 = AddFlag(aInsert2, "MODELNO", addstring(loctable2!MODELNO))
        aInsert2 = AddFlag(aInsert2, "DISC2", Val(loctable2!disc2))
        cStr2 = cStr2 & addInsert(aInsert2, "FILE0_90") & " ; "
        nCountR = nCountR + 1
        If nCountR = 100 Then
            If cStr2 <> "" Then con.Execute cStr2
            cStr2 = ""
            nCountR = 0
        End If
        loctable2.MoveNext
    Loop
    If cStr2 <> "" Then con.Execute cStr2
    loctable.MoveNext
Loop
Inform " „Ã„Ê⁄«  Œ’„" & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
nRecord = -1
Err.Clear
GoTo lastsub
End Function

