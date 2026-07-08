VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form transDatafrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”Õ» «·»Ì«‰«  „‰ «·—∆Ì”Ì"
   ClientHeight    =   4080
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   " ÕœÌÀ «’‰«› ›ﬁÿ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3195
      TabIndex        =   10
      Top             =   3600
      Width           =   2625
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ISDATE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   4950
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3195
      Width           =   870
   End
   Begin VB.CommandButton cmdsend 
      Caption         =   "‰ﬁ· «·„»Ì⁄«  "
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
      TabIndex        =   8
      Top             =   945
      Width           =   5730
   End
   Begin VB.Frame Frame3 
      Caption         =   "Drive"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1710
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3150
      Width           =   1050
      Begin VB.TextBox xDrive 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   225
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ÿ—Ìﬁ… «·”Õ»"
      Height          =   690
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3105
      Width           =   1590
      Begin VB.CheckBox xAuto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   " ·ﬁ«∆Ì"
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   315
         Width           =   1095
      End
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
      Left            =   90
      TabIndex        =   3
      Top             =   1800
      Width           =   5730
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "”Õ» «·»Ì«‰«  „‰ «·›—⁄ «·—∆Ì”Ì "
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
      Top             =   90
      Width           =   5730
   End
   Begin VB.Frame Frame11 
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2520
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
Attribute VB_Name = "transDatafrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dPostDate As String
Dim con As New ADODB.Connection, cDataFolder As String, cDataFile As String
Public nFlag As Integer
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    Me.Caption = "”Õ» «·»Ì«‰«  „‰ «·›—⁄ «·—∆Ì”Ì"
    If Not getData Then
        MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰ ”Õ» «·»Ì«‰«  „‰ «·›—⁄ «·—∆Ì”Ì"
    Else
        MsgBox " „ﬂ‰ «·‰Ÿ«„ „‰ ”Õ» «·»Ì«‰«  „‰ «·›—⁄ «·—∆Ì”Ì"
    End If
End Sub
Private Sub cmdsend_Click()
    Me.Caption = "‰ﬁ· «·»Ì«‰«  «·Ì «·›—⁄ «·—∆Ì”Ì"
    If Not sendData Then
        MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰ ‰ﬁ· «·»Ì«‰«  ≈·Ï «·›—⁄ «·—∆Ì”Ì"
    Else
        MsgBox " „ﬂ‰ «·‰Ÿ«„ „‰ ‰ﬁ· «·»Ì«‰«  ≈·Ì «·›—⁄ «·—∆Ì”Ì"
        Unload Me
    End If
End Sub

Private Sub Command1_Click()
Dim DDate1 As Date
If InputBox("", "") = "815815" Then
    DDate1 = DateAdd("M", -1, Date)
    con.Execute " UPDATE FILE6_20H SET ISDATE = DATE WHERE DATE <= " & DateSq(DDate1)
End If
End Sub
Private Sub Command2_Click()
    nRecordCount = getItems
    If nRecordCount >= 0 Then
        Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«’‰«›", "»‰Ã«Õ"
    End If
End Sub
Private Sub Form_Load()
'    xDrive.Text = RetSetting("drive", TempSave(Me))
'    xAuto.Value = Val(RetSetting("auto", TempSave(Me)))
    cDataFolder = App.Path & "\mdb"
    cDataFile = "data"
    openCon con
    dPostDate = DateAdd("M", -1, Date)
End Sub
Private Function getData() As Boolean
Dim bExit As Boolean

If validData Then
'    con.BeginTrans
    
    nRecordCount = getCode("FILE3_10")
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·⁄„·«¡  ", "»‰Ã«Õ" Else GoTo myerror
    
''    nRecordCount = getStock
''    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·Ã—œ ", "»‰Ã«Õ" Else GoTo myerror
    
    nRecordCount = GetGroup
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ „Ã„Ê⁄«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = getItems
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = getCode("FILE1_10SC", True)
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«ﬁ”«„", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = Getfact
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„’«‰⁄ ", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = getCode("FILE3_50")
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  „Ã„Ê⁄«  «·⁄„·«¡", "»‰Ã«Õ" Else GoTo myerror
    
    
    nRecordCount = getBox
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·Œ“‰", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = getStores
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„Œ«“‰", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = getCode("FILE6_25")
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·»«∆⁄Ì‰", "»‰Ã«Õ" Else GoTo myerror

'    nRecordCount = getCode("FILE8_51")
'    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «ﬂÊ«œ «·„’«—Ì›", "»‰Ã«Õ" Else GoTo myerror
'
'    nRecordCount = getCode("FILE8_52")
'    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «ﬂÊ«œ „Ã„Ê⁄«  «·„’«—Ì›", "»‰Ã«Õ" Else GoTo myerror
'
'    nRecordCount = getCode("FILE8_61")
'    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «ﬂÊ«œ «·«Ì—«œ« ", "»‰Ã«Õ" Else GoTo myerror
''
'    nRecordCount = getCode("FILE8_62")
'    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «ﬂÊ«œ „Ã„Ê⁄«  «·«Ì—«œ« ", "»‰Ã«Õ" Else GoTo myerror
''
'    nRecordCount = getCode("FILE8_71")
'    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «ﬂÊ«œ «·‘—ﬂ«¡", "»‰Ã«Õ" Else GoTo myerror
    
    nRecordCount = getMOSM
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ «·„Ê«”„ ", "»‰Ã«Õ" Else GoTo myerror
    
    
    nRecordCount = getTrans
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«   ÕÊÌ·«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror
    
    nRecordCount = getPurchase("FILE7_20")
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„‘ —Ì« ", "»‰Ã«Õ" Else GoTo myerror
    
    nRecordCount = getPurchase("FILE7_10")
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  „—œÊœ «·„‘ —Ì« ", "»‰Ã«Õ" Else GoTo myerror
    
    nRecordCount = getCode("FILE4_10")
    If nRecordCount >= 0 Then Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„Ê—œÌ‰ ", "»‰Ã«Õ" Else GoTo myerror
'    con.Execute " UPDATE FILE4_10 SET [GROUP] = '01' "
    
'    con.Execute "  DELETE FROM FILE1_10 WHERE ITEM NOT IN (SELECT ITEM FROM FILE1_11)"
'    con.CommitTrans
    getData = True
Else
    MsgBox "„·› ‰ﬁ· «·»Ì«‰«  €Ì— „ÊÃÊœ"
End If
Exit Function
myerror:
'con.RollbackTrans
End Function
Private Function sendData() As Boolean
If copyData2 Then
    Dim nRecordCount As Long
    Me.Caption = "„»Ì⁄« "
    nRecordCount = SendSales
    If nRecordCount >= 0 Then
        'Inform " „ ‰ﬁ· " & nRecordcount & " ”Ã· „‰ Õ—ﬂ«  «·»Ì⁄", "»‰Ã«Õ"
    Else
        Exit Function
    End If
    
    Me.Caption = " ·Ì›Ê‰"
    SendSalesPhone
    
    
    Me.Caption = "„»Ì⁄«  2"
    nRecordCount = SendSalesHeader
    If nRecordCount >= 0 Then
        Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ „” ‰œ«  «·»Ì⁄", "»‰Ã«Õ"
    Else
        Exit Function
    End If
    
    
    Me.Caption = "„’«—Ì›"
    nRecordCount = sendCharge("FILE8_50")
    If nRecordCount >= 0 Then
        'Inform " „ ‰ﬁ· " & nRecordcount & " ”Ã· „‰ Õ—ﬂ«  «·«Ì—«œ« ", "»‰Ã«Õ"
    Else
        Exit Function
    End If

    Me.Caption = "„’«—Ì›2"
    nRecordCount = sendChargeHeader("FILE8_50H")
    If nRecordCount >= 0 Then
        Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ „” ‰œ«  «·„’«—Ì›", "»‰Ã«Õ"
    Else
        Exit Function
    End If
    
'***

'***
        
    
    MsgBox " „ﬂ‰ «·‰Ÿ«„ „‰ ‰ﬁ· «·„·›«  «·Ì «·›·«‘…"
Else
    MsgBox "„·› ‰ﬁ· «·»Ì«‰«  €Ì— „ÊÃÊœ"
    Exit Function
End If
sendData = True
End Function
Private Function GetGroup() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, cFile As String
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

cFile = "FILE1_50"
con.Execute "Delete  from " & cFile

cString = "SELECT * FROM " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If


Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
'    aInsert = AddFlag(aInsert, "[GROUP]", addvalue(loctable!Group))
    
    con.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    GetGroup = GetGroup + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GetGroup = -1
GoTo lastsub
End Function
Private Function getItems() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, sCaption As String
'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
Dim cFile As String

cFile = "FILE1_10"

'con.Execute " delete from file1_10 "

cString = "SELECT * FROM " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
sCaption = Me.Caption
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    Me.Caption = sCaption & " ”Ã· " & nRecord & " „‰ " & nRecordCount

    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!MODEL))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable!MODELNO))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "FACT", addstring(loctable!Fact))
    aInsert = AddFlag(aInsert, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "MODELFACT", addstring(loctable!modelfact))
    aInsert = AddFlag(aInsert, "MODELFACT0", addstring(loctable!modelfact0))
    aInsert = AddFlag(aInsert, "SUPP", addstring(loctable!SUPP))
    aInsert = AddFlag(aInsert, "FIXPRICE", Val(loctable!fixprice & ""))
    aInsert = AddFlag(aInsert, "OKAZ", Val(loctable!okaz & ""))
    aInsert = AddFlag(aInsert, "RATE", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "REDEM", Val(loctable!Redem & ""))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "[SECTION]", addvalue(loctable!Section))
    aInsert = AddFlag(aInsert, "[ITEM]", addvalue(loctable!Item))
    aInsert = AddFlag(aInsert, "[COST]", Val(loctable!cost & ""))
    aInsert = AddFlag(aInsert, "[COST2]", Val(loctable!cost2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE]", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "[PRICE2]", Val(loctable!PRICE2 & ""))
    aInsert = AddFlag(aInsert, "[SCAL]", addstring(loctable!scal))
    aInsert = AddFlag(aInsert, "[C_SCAL]", addvalue(loctable!C_SCAL))
    aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable!Color))
    aInsert = AddFlag(aInsert, "[C_COLOR]", addstring(loctable!c_Color))
    aInsert = AddFlag(aInsert, "[BARCODE]", addstring(loctable!BARCODE))
    aInsert = AddFlag(aInsert, "[BARCODE2]", addstring(loctable!BARCODE2))
    aInsert = AddFlag(aInsert, "[BARCODE13]", addstring(loctable!BARCODE13))
    aInsert = AddFlag(aInsert, "[ISOKAZITEM]", IIf(loctable!ISOKAZITEM, 1, 0))
    
    con.Execute " delete from file1_10 WHERE ITEM = " & loctable!Item
    con.Execute addInsert(aInsert, "FILE1_10"), nAffect
    loctable.MoveNext
    getItems = getItems + nAffect
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
getItems = -1
GoTo lastsub
End Function
Private Function getCode(cFile As String, Optional isNumber As Boolean = False) As Long
On Error GoTo myerror
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
con.Execute "DELETE FROM " & cFile
cString = "SELECT * FROM " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    

    If isNumber Then
        aInsert = AddFlag(Empty, "CODE", addvalue(loctable!CODE))
    Else
        aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    End If
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    
    con.Execute addInsert(aInsert, cFile), nAffect
    getCode = getCode + nAffect
    loctable.MoveNext
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getCode = -1
GoTo lastsub
End Function
Private Function getStores()
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

cFile = "FILE0_40"
con.Execute "DELETE FROM " & cFile
cString = "SELECT * FROM " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))

    con.Execute addInsert(aInsert, cFile), nAffect
    getStores = getStores + nAffect
    loctable.MoveNext
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getStores = -1
GoTo lastsub
End Function
Private Function getClients() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, cFile As String
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

cFile = "FILE3_10"
'con.Execute "DELETE FROM " & cFile

DeleteValid "FILE3_10", "CODE"

cString = "SELECT * FROM " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "MANAGER", addstring(loctable!Manager))
    aInsert = AddFlag(aInsert, "ADDRESS", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "PHONE1", addstring(loctable!PHONE1))
    aInsert = AddFlag(aInsert, "FAX", addstring(loctable!Fax))
    aInsert = AddFlag(aInsert, "F_BALANCE", Val(loctable!F_Balance & ""))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "[F_DATE]", addDate(Format(loctable!F_DATE, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[EMAIL]", addstring(loctable!email))
    aInsert = AddFlag(aInsert, "[REMARK]", addstring(loctable!remark))
    aInsert = AddFlag(aInsert, "DISC", Val(loctable!disc & ""))
    aInsert = AddFlag(aInsert, "[CASH]", Val(loctable!CASH & ""))
   If GetDesca("SELECT CODE FROM " & cFile & " WHERE CODE = " & MyParn(loctable!CODE), con) <> "" Then
        con.Execute addUpdate(aInsert, cFile, "CODE = " & MyParn(loctable!CODE)), nAffect
    Else
        con.Execute addInsert(aInsert, cFile), nAffect
    End If
    loctable.MoveNext
    getClients = getClients + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getClients = -1
GoTo lastsub
End Function
Private Function getClientsCash() As Long
Dim conmdb As New ADODB.Connection, cFile As String, loctable As New ADODB.Recordset, cWhere As String
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
cFile = "FILE3_20"
cWhere = "FILE3_20.BRANCH <> " & MyParn(sBranchCode)

con.Execute "DELETE FROM " & cFile & turn(cWhere) & cWhere

cString = "SELECT * FROM " & cFile & turn(cWhere) & cWhere
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long

If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If

Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "MANAGER", addstring(loctable!Manager))
    aInsert = AddFlag(aInsert, "ADDRESS", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "PHONE1", addstring(loctable!PHONE1))
    aInsert = AddFlag(aInsert, "FAX", addstring(loctable!Fax))
    aInsert = AddFlag(aInsert, "[F_DATE]", addDate(Format(loctable!F_DATE, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[B_DATE]", addDate(Format(loctable!B_DATE, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[EMAIL]", addstring(loctable!email))
    aInsert = AddFlag(aInsert, "[REMARK]", addstring(loctable!remark))
    aInsert = AddFlag(aInsert, "DISC", Val(loctable!disc & ""))
    aInsert = AddFlag(aInsert, "BRANCH", addstring(loctable!branch))
    
    con.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    getClientsCash = getClientsCash + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getClientsCash = -1
GoTo lastsub
End Function
Private Function getBox() As Long
On Error GoTo myerror
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
Dim cFile As String

cFile = "FILE0_50"
con.Execute "DELETE FROM " & cFile

cString = "SELECT * FROM " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100

    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "F_DATE", addDate(Format(loctable!F_DATE, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "F_BAL", Val(loctable!F_BAL & ""))
              
    con.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    getBox = getBox + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
getBox = -1
GoTo lastsub
End Function
Private Function getTrans() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

cFile = "FILE1_60"
'cString = "Delete from " & cFile
'cString = cString & turn(cString) & " DOC_NO IN (SELECT DOC_NO FROM " & cFile & "H" & " WHERE BRANCH <> " & MyParn(sBranchCode) & ")"
'con.Execute cString

'cString = "Delete from " & cFile & "H"
'cString = cString & turn(cString) & " BRANCH <> " & MyParn(sBranchCode)
'con.Execute cString

Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile & "H"
'cString = cString & turn(cString) & " STORE1 = " & MyParn(cBranchStore) & " OR STORE2 = " & MyParn(cBranchStore)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
Dim aInsert As Variant
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    con.Execute " DELETE FROM FILE1_60  WHERE DOC_NO = " & MyParn(loctable!doc_no)
    con.Execute " DELETE FROM FILE1_60H WHERE DOC_NO = " & MyParn(loctable!doc_no)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!doc_no = "002120" Then
        A = A
    End If
    If loctable!STORE1 = cBranchStore Or loctable!STORE2 = cBranchStore Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
        aInsert = AddFlag(aInsert, "[STORE1]", addstring(loctable!STORE1))
        aInsert = AddFlag(aInsert, "[STORE2]", addstring(loctable!STORE2))
    
        con.Execute addInsert(aInsert, cFile & "H"), nAffect
    End If
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
getTrans = nAffectTotal

Set loctable = New ADODB.Recordset
cString = "Select " & cFile & ".* from " & cFile & " INNER JOIN " & cFile & "H on " & cFile & ".doc_no = " & cFile & "H.DOC_NO"
cString = cString & turn(cString) & cFile & "H.STORE2 = " & MyParn(cBranchStore) & " OR " & cFile & "H.STORE1 = " & MyParn(cBranchStore)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!doc_no = "002120" Then
        A = A
    End If
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[ITEM]", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "[QUANT]", Val(loctable!Quant & ""))
    aInsert = AddFlag(aInsert, "[COST]", Val(loctable!cost & ""))
                    
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
getTrans = -1
GoTo lastsub
End Function
Private Function getPurchase(cFile As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

'cString = "Delete from " & cFile
'con.Execute cString
'
'cString = "Delete from " & cFile & "H"
'con.Execute cString

Set loctable = New ADODB.Recordset
cString = "Select * from " & cFile & "H"
'cString = cString & turn(cString) & " STORE = " & MyParn(cBranchStore)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim aInsert As Variant
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount
'loctable.MoveFirst
Do Until loctable.EOF
    con.Execute " DELETE FROM " & cFile & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    con.Execute " DELETE FROM " & cFile & "H WHERE DOC_NO = " & MyParn(loctable!doc_no)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!STORE = cBranchStore Then
    
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
        aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
        aInsert = AddFlag(aInsert, "[TAX]", Val(loctable!tax & ""))
        aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM))
        aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
        aInsert = AddFlag(aInsert, "[ISCLOSED]", 1)
    
        con.Execute addInsert(aInsert, cFile & "H"), nAffect
    End If
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
getPurchase = nAffectTotal

Set loctable = New ADODB.Recordset
cString = "Select " & cFile & ".* from " & cFile & " INNER JOIN " & cFile & "H on " & cFile & ".doc_no = " & cFile & "H.DOC_NO"
cString = cString & turn(cString) & cFile & "H.STORE = " & MyParn(cBranchStore)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "ITEM", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant & ""))
    aInsert = AddFlag(aInsert, "PRICE", Val(loctable!price & ""))
                        
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
getPurchase = -1
GoTo lastsub
End Function
Private Function CopyData() As Boolean
Dim fs As New FileSystemObject
If fs.FileExists(cDataFolder & "\" & cDataFile & "_" & "blk.mdb") Then
    fs.CopyFile cDataFolder & "\" & cDataFile & "_" & "blk.mdb", cDataFolder & "\" & cDataFile & ".mdb"
End If
CopyData = True
End Function
Private Function CopyToBranch() As Boolean
Dim fs As New FileSystemObject
On Error GoTo myerror
If fs.FileExists(cDataFolder & "\" & cDataFile & ".mdb") Then
    If Trim(xDrive.text) <> "" Then sLastDrive = xDrive.text Else sLastDrive = LastDrive(True)
    If sLastDrive <> "" Then
        MyCreateFolder (sLastDrive & ":\data_pro_trans\mdb")
        fs.CopyFile cDataFolder & "\" & cDataFile & ".mdb", sLastDrive & ":\data_pro_trans\mdb" & "\" & cDataFile & ".mdb"
    End If
End If
CopyToBranch = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function CopyFromMain() As Boolean
Dim fs As New FileSystemObject
On Error GoTo myerror
If Trim(xDrive.text) <> "" Then sLastDrive = xDrive.text Else sLastDrive = LastDrive(True)
If sLastDrive <> "" Then
'    noReadOnly sLastDrive & ":\data_pro_trans\mdb" & "\" & cDataFile & ".mdb"
'    noReadOnly cDataFolder & "\" & cDataFile & ".mdb"
    If fs.FileExists(sLastDrive & ":\data_pro_trans\mdb" & "\" & cDataFile & ".mdb") Then
        fs.CopyFile sLastDrive & ":\data_pro_trans\mdb" & "\" & cDataFile & ".mdb", cDataFolder & "\" & cDataFile & ".mdb"
    Else
        Exit Function
    End If
End If
CopyFromMain = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function CopyToMain() As Boolean
Dim fs As New FileSystemObject
On Error GoTo myerror
If fs.FileExists(cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb") Then
    If Trim(xDrive.text) <> "" Then sLastDrive = xDrive.text Else sLastDrive = LastDrive(True)
    If sLastDrive <> "" Then
        MyCreateFolder (sLastDrive & ":\data_pro_trans\mdb")
        fs.CopyFile cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb", sLastDrive & ":\data_pro_trans\mdb" & "\" & cDataFile & "_" & cBranch & ".mdb", True
    End If
End If
CopyToMain = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function copyData2() As Boolean
Dim fs As New FileSystemObject
If fs.FileExists(cDataFolder & "\" & cDataFile & "_" & "blk.mdb") Then
    fs.CopyFile cDataFolder & "\" & cDataFile & "_" & "blk.mdb", cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"
Else
    Exit Function
End If
copyData2 = True
End Function
Private Function validData() As Boolean
Dim fs As New FileSystemObject
If Not fs.FileExists(cDataFolder & "\" & cDataFile & ".mdb") Then Exit Function
validData = True
End Function
Private Function SendSalesHeader() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cFile = "FILE6_20H"
cString = "Select * from file6_20h where isdate >= " & DateSq(dPostDate)

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
Dim aInsert As Variant
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[DATE]", addstring(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "[NOTES]", addstring(loctable!NOTES))
    aInsert = AddFlag(aInsert, "[RATE]", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[CASH]", Val(loctable!CASH & ""))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[PAY]", Val(loctable!PAY & ""))
    aInsert = AddFlag(aInsert, "[VISA]", Val(loctable!Visa & ""))
    aInsert = AddFlag(aInsert, "[LATE]", Val(loctable!late & ""))
    aInsert = AddFlag(aInsert, "[REST]", Val(loctable!Rest & ""))
    aInsert = AddFlag(aInsert, "[MAN]", addstring(Mid(loctable!MAN, 1, 3)))
    aInsert = AddFlag(aInsert, "[PRINTED]", IIf(loctable!PRINTED, "1", "0"))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(cBranch))
    aInsert = AddFlag(aInsert, "[phone]", addstring(loctable!phone))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
SendSalesHeader = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendSalesHeader = -1
GoTo lastsub
End Function
Private Function SendSales() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cFile = "FILE6_20"
cString = "Select file6_20.* from file6_20 inner join file6_20h on file6_20h.doc_no = file6_20.doc_no where isdate >= " & DateSq(dPostDate)

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim aInsert As Variant
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "ITEM", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "MAN", addstring(loctable!MAN))
    aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant & ""))
    aInsert = AddFlag(aInsert, "PRICE", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "DISCOUNT", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "PRICE_C", Val(loctable!PRICE_C & ""))
    aInsert = AddFlag(aInsert, "ROW", Val(loctable!Row & ""))
    aInsert = AddFlag(aInsert, "COST", Val(loctable!cost & ""))
                    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
SendSales = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendSales = -1
GoTo lastsub
End Function
Private Function sendCashHeader(cFile As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cString = "Select * from " & cFile

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
Dim aInsert As Variant

nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addstring(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(cBranch))
                               
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendCashHeader = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendCashHeader = -1
GoTo lastsub
End Function
Private Function sendCash(cFile As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cString = "Select * from " & cFile

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim aInsert As Variant
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[CODE]", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(loctable!Value & ""))
    aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "[ROW]", Val(loctable!Row & ""))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendCash = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendCash = -1
GoTo lastsub
End Function
Private Function sendChargeHeader(cFile As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cString = "Select * from " & cFile

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(cBranch))
                           
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendChargeHeader = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendChargeHeader = -1
GoTo lastsub
End Function
Private Function sendCharge(cFile As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String

'On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cString = "Select * from " & cFile

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[CHARGE]", addstring(loctable!CHARGE))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(loctable!Value & ""))
'    If Trim(LCase(cFile)) = "file8_50" Then
''        aInsert = AddFlag(aInsert, "[FORM_NO]", addstring(loctable!FORM_NO))
'        aInsert = AddFlag(aInsert, "[BOX2]", addstring(sBranchCode))
'    End If
    aInsert = AddFlag(aInsert, "[ROW]", Val(loctable!Row & ""))
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendCharge = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendCharge = -1
GoTo lastsub
End Function
Private Function sendTransHeader() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cFile = "FILE1_60H"
cString = "Select * from " & cFile
cString = cString & turn(cString) & " FILE1_60H.BRANCH = " & sBranchCode

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
Dim aInsert As Variant
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addstring(Format(loctable!Date, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[STORE1]", addstring(loctable!STORE1))
    aInsert = AddFlag(aInsert, "[STORE2]", addstring(loctable!STORE2))
    aInsert = AddFlag(aInsert, "[RECEIVED]", IIf(loctable!Received, 1, 0))
    aInsert = AddFlag(aInsert, "[RECEIVED_ALL]", IIf(loctable!Received_all, "1", "0"))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranchCode))
                    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendTransHeader = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTransHeader = -1
GoTo lastsub
End Function
Private Function sendTrans() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cFile = "FILE1_60"
cString = "Select FILE1_60.* FROM  file1_60 inner join file1_60h on file1_60.doc_no = FILE1_60H.DOC_NO"
cString = cString & turn(cString) & "FILE1_60H.Branch = " & sBranchCode
Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
Dim aInsert As Variant
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
        
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "ITEM", addvalue(loctable!Item))
    aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant & ""))
    aInsert = AddFlag(aInsert, "COST", Val(loctable!Quant & ""))
    aInsert = AddFlag(aInsert, "[RECEIVED]", IIf(loctable!Received, "1", "0"))
        
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendTans = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTrans = -1
GoTo lastsub
End Function
Private Function sendStockHeader() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cFile = "FILE0_10H"
cString = "Select * from " & cFile
cString = cString & turn(cString) & "FILE0_10H.CLOSED = 1"


Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
Dim aInsert As Variant
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addstring(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranchCode))
                    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendStockHeader = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendStockHeader = -1
GoTo lastsub
End Function
Private Function sendSTOCK() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cFile = "FILE0_10"
cString = "Select FILE0_10.* from FILE0_10 INNER JOIN FILE0_10H ON FILE0_10.DOC_NO = FILE0_10H.DOC_NO"
cString = cString & turn(cString) & "FILE0_10H.CLOSED = 1"

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
Dim aInsert As Variant

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
    conmdb.Execute addInsert(aInsert, cFile & "0"), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendSTOCK = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendSTOCK = -1
GoTo lastsub
End Function
Private Function SendBoxTrans() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cFile = "FILE0_51"
cString = "Select * from " & cFile
cString = cString & turn(cString) & " BRANCH = " & MyParn(sBranchCode)

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "NO1", addstring(loctable!no1))
    aInsert = AddFlag(aInsert, "NO2", addstring(loctable!no2))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(loctable!Value & ""))
    aInsert = AddFlag(aInsert, "BRANCH", addstring(cBeginCol))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
SendBoxTrans = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendBoxTrans = -1
GoTo lastsub
End Function
Private Function sendClientsCash() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cFile = "FILE3_20"
cString = "Select * from " & cFile
cString = cString & turn(cString) & " BRANCH = " & MyParn(sBranchCode)

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText


prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
Dim aInsert As Variant
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "MANAGER", addstring(loctable!Manager))
    aInsert = AddFlag(aInsert, "ADDRESS", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "PHONE1", addstring(loctable!PHONE1))
    aInsert = AddFlag(aInsert, "FAX", addstring(loctable!Fax))
    aInsert = AddFlag(aInsert, "[F_DATE]", addstring(Format(loctable!F_DATE, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[B_DATE]", addstring(Format(loctable!B_DATE, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[EMAIL]", addstring(loctable!email))
    aInsert = AddFlag(aInsert, "[REMARK]", addstring(loctable!remark))
    aInsert = AddFlag(aInsert, "DISC", Val(loctable!disc & ""))
    aInsert = AddFlag(aInsert, "BRANCH", addstring(sBranchCode))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendClientsCash = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendClientsCash = -1
GoTo lastsub
End Function
Private Sub DeleteValid(sFile As String, sField As String, Optional bNum As Boolean)
Dim loctable As New ADODB.Recordset, sCaption As String
loctable.Open "SELECT " & sField & " FROM " & sFile, con, adOpenStatic, adLockReadOnly, adCmdText
On Error Resume Next
prog1.Value = 0
prog1.Visible = True
nRecordCount = loctable.RecordCount
sCaption = Me.Caption
Do Until loctable.EOF
    nRecord = nRecord + 1
    Me.Caption = sCaption & " ”Ã· " & nRecord & " „‰ " & nRecordCount
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    con.Execute "DELETE FROM " & sFile & " WHERE " & sField & " = " & MyParn(loctable(sField & ""))
    If Err.Number <> 0 Then Err.Clear
    loctable.MoveNext
Loop
prog1.Visible = False
Me.Caption = sCaption
End Sub
Private Sub Form_Unload(Cancel As Integer)
'addSetting "drive", xDrive.Text, TempSave(Me)
'addSetting "auto", xAuto.Value, TempSave(Me)
Set transDatafrm = Nothing
End Sub
Private Function sendItems() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

Dim cFile As String
cFile = "FILE1_10"
cString = "SELECT * FROM " & cFile

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
sCaption = Me.Caption
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    Me.Caption = sCaption & " ”Ã· " & nRecord & " „‰ " & nRecordCount
    
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!MODEL))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable!MODELNO))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "FACT", addstring(loctable!Fact))
    aInsert = AddFlag(aInsert, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "MODELFACT", addstring(loctable!modelfact))
    aInsert = AddFlag(aInsert, "MODELFACT0", addstring(loctable!modelfact0))
    aInsert = AddFlag(aInsert, "SUPP", addstring(loctable!SUPP))
    aInsert = AddFlag(aInsert, "FIXPRICE", Val(loctable!fixprice & ""))
    aInsert = AddFlag(aInsert, "OKAZ", Val(loctable!okaz & ""))
    aInsert = AddFlag(aInsert, "RATE", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "REDEM", Val(loctable!Redem & ""))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "[SECTION]", addvalue(loctable!Section))
    aInsert = AddFlag(aInsert, "[ITEM]", addvalue(loctable!Item))
    aInsert = AddFlag(aInsert, "[COST]", Val(loctable!cost & ""))
    aInsert = AddFlag(aInsert, "[COST2]", Val(loctable!cost2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE]", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "[PRICE2]", Val(loctable!PRICE2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE3]", Val(loctable!PRICE3 & ""))
    aInsert = AddFlag(aInsert, "[SCAL]", addstring(loctable!scal))
    aInsert = AddFlag(aInsert, "[C_SCAL]", addvalue(loctable!C_SCAL))
    aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable!Color))
    aInsert = AddFlag(aInsert, "[C_COLOR]", addstring(loctable!c_Color))
    aInsert = AddFlag(aInsert, "[SHOWSALES]", IIf(loctable!ShowSales, 1, 0))
    aInsert = AddFlag(aInsert, "[CODE_SCAL]", Val(loctable!CODE_SCAL & ""))
    aInsert = AddFlag(aInsert, "[ISSTOP]", IIf(loctable!ISSTOP, "1", 0))
    
    conmdb.Execute addInsert(aInsert, "FILE1_10_BRANCH"), nAffect
    
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

Private Function getMOSM() As Long
On Error GoTo myerror
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"
Dim cFile As String

cFile = "MOSM"
con.Execute "DELETE FROM " & cFile

cString = "SELECT * FROM " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100

    aInsert = AddFlag(Empty, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "DATE", addDate(Format(loctable!Date, "DD-MM-YYYY")))
              
    con.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    getMOSM = getMOSM + nAffect
Loop

'If GetDesca("SELECT PRICE FROM FILE0_00", conmdb) = "True" Then
'    con.Execute " UPDATE FILE0_00 SET PRICE = 1 "
'    lokaz = True
'Else
'    con.Execute " UPDATE FILE0_00 SET PRICE = 0 "
'    lokaz = False
'End If

lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
getMOSM = -1
GoTo lastsub
End Function



Private Function Getfact() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset, cFile As String
On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"

cFile = "fact"
con.Execute "Delete  from " & cFile

cString = "SELECT * FROM " & cFile
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long
If Not (loctable.EOF And loctable.BOF) Then
    loctable.MoveLast
    nRecordCount = loctable.RecordCount
    loctable.MoveFirst
End If


Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    
    con.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    Getfact = Getfact + nAffect
Loop
lastsub:
prog1.Visible = False
conmdb.Close
Set conmdb = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
Getfact = -1
GoTo lastsub
End Function
Private Function sendStock0() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cFile = "FILE0_10"
cString = "Select FILE0_10.* from FILE0_10 INNER JOIN FILE0_10H ON FILE0_10.DOC_NO = FILE0_10H.DOC_NO"
cString = cString & turn(cString) & "FILE0_10H.CLOSED = 1"

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
Dim aInsert As Variant

nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[ITEM]", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "[RealBal]", Val(loctable!RealBal & ""))
    aInsert = AddFlag(aInsert, "[ComputerBal]", Val(loctable!ComputerBal & ""))
    aInsert = AddFlag(aInsert, "[Differ]", Val(loctable!Differ & ""))
                        
    conmdb.Execute addInsert(aInsert, "FILE0_100"), nAffect
    conmdb.Execute addInsert(aInsert, "FILE0_10"), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendStock0 = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendStock0 = -1
GoTo lastsub
End Function
Private Function getStock() As Long
Dim conmdb As New ADODB.Connection, loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & ".mdb"


Set loctable = New ADODB.Recordset

cString = "Select * from FILE0_10H"
'cString = cString & turn(cString) & " STORE = " & MyParn(cBranchStore)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
Dim aInsert As Variant
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    con.Execute " DELETE FROM FILE0_10 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    con.Execute " DELETE FROM FILE0_100 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    con.Execute " DELETE FROM FILE0_10H WHERE DOC_NO = " & MyParn(loctable!doc_no)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    If loctable!STORE = cBranchStore Then
    
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
        aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
        aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
        con.Execute addInsert(aInsert, "FILE0_10H"), nAffect
    End If
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
getStock = getStock + nAffectTotal

Set loctable = New ADODB.Recordset


cFile = "FILE0_10"
cString = "Select " & cFile & ".* from " & cFile & " INNER JOIN " & cFile & "H on " & cFile & ".doc_no = " & cFile & "H.DOC_NO"
cString = cString & turn(cString) & cFile & "H.STORE = " & MyParn(cBranchStore)
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[ITEM]", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "[RealBal]", Val(loctable!RealBal & ""))
    aInsert = AddFlag(aInsert, "[ComputerBal]", Val(loctable!ComputerBal & ""))
    aInsert = AddFlag(aInsert, "[Differ]", Val(loctable!Differ & ""))
                    
    con.Execute addInsert(aInsert, cFile), nAffect
    con.Execute addInsert(aInsert, "file0_100"), nAffect
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
Private Function sendpart(cFile As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cString = "Select * from " & cFile

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[code]", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[VALUE_p]", Val(loctable!Value_P & ""))
    aInsert = AddFlag(aInsert, "[VALUE_m]", Val(loctable!Value_M & ""))
    aInsert = AddFlag(aInsert, "[ROW]", Val(loctable!Row & ""))
    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendpart = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendpart = -1
GoTo lastsub
End Function
Private Function sendpartHeader(cFile As String) As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String

On Error GoTo myerror
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cString = "Select * from " & cFile

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(cBranch))
                           
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
sendpartHeader = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendpartHeader = -1
GoTo lastsub
End Function


Private Function SendSalesPhone() As Long
Dim conmdb As New ADODB.Connection, loctable As ADODB.Recordset, cString As String, cFile As String

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source = " & cDataFolder & "\" & cDataFile & "_" & cBranch & ".mdb"

cFile = "SUBCUST"
cString = "Select * FROM SUBCUST "

Set loctable = New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim aInsert As Variant
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "PHONE", addstring(loctable!phone))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "E_MAIL", addstring(loctable!E_MAIL))
    aInsert = AddFlag(aInsert, "branch", addstring(loctable!branch))
                    
    conmdb.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
SendSalesPhone = nAffectTotal

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendSalesPhone = -1
GoTo lastsub
End Function

