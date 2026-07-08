VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form transdata_ip_Sales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " —ÕÌ· «·„»Ì⁄«  «·ÌÊ„Ì… ··”Ì—›—"
   ClientHeight    =   1275
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
   ScaleHeight     =   1275
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Height          =   465
      Left            =   90
      Picture         =   "transdata_ip_Sales.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   270
      Width           =   1005
   End
   Begin VB.Frame Frame11 
      Height          =   600
      Left            =   1125
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   495
      Width           =   5505
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   1
         Top             =   180
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
End
Attribute VB_Name = "transdata_ip_Sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lAddAll As Boolean
Dim con As New ADODB.Connection, pBranch As String
Private Sub Form_Load()
    openCon con
    strConShop = LoadConStringshop
    If openConShop(conShop) = "ok" Then
'       Unload Me
    Else
        MsgBox "ÌÊÃœ „‘ﬂ·… ›Ï ‰ﬁ· «·„»Ì⁄«  «·ÌÊ„Ì… "
        Unload Me
    End If
End Sub

Private Function SendSales() As Long
cFile = "FILE6_20"
Set loctable = New ADODB.Recordset
Dim cStr1 As String, cStr2 As String, cDelStr1 As String, cDelStr2 As String, cStrNew As String, cStrPhone As String
cString = "Select * from " & cFile & "H WHERE PRINTED = 1 and isnew = 1"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long, lSkip As Boolean
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
    If cDelStr1 = "" Then
        cDelStr1 = " delete from file6_20 where doc_no = " & MyParn(loctable!doc_no)
        cDelStr2 = " delete from file6_20H where doc_no = " & MyParn(loctable!doc_no)
    Else
        cDelStr1 = cDelStr1 & " ; delete from file6_20 where doc_no = " & MyParn(loctable!doc_no)
        cDelStr2 = cDelStr2 & " ; delete from file6_20H where doc_no = " & MyParn(loctable!doc_no)
    End If
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!code))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[TIME]", addTime(Format(loctable!Time, "SHORT TIME")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!store))
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
                                                                                
    aInsert = AddFlag(aInsert, "[username_disc]", addstring(loctable!username_disc))
    aInsert = AddFlag(aInsert, "[username_ret]", addstring(loctable!username_RET))
    aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(loctable!SALES_RET))
    aInsert = AddFlag(aInsert, "CODEVISA", addvalue(loctable!CODEVISA))
                                                                                
    If cStr1 = "" Then
        cStr1 = addInsert(aInsert, cFile & "H")
    Else
        cStr1 = cStr1 & " ; " & addInsert(aInsert, cFile & "H")
    End If
    
    lSkip = False
    
    Set loctable2 = New ADODB.Recordset
    cString = "Select FILE6_20.* from FILE6_20 WHERE FILE6_20.doc_no = " & MyParn(loctable!doc_no)
    loctable2.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    Dim sitem As String
    Dim aitem As Variant
    Do Until loctable2.EOF
        sitem = loctable2!Item
        If GetDesca("select item from file1_10 where  item = " & sitem, conShop) <> "" Then
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "MAN", addstring(loctable2!MAN))
            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
            aInsert2 = AddFlag(aInsert2, "DISCOUNT", Val(loctable2!discount & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE_C", Val(loctable2!PRICE_C & ""))
            aInsert2 = AddFlag(aInsert2, "ROW", Val(loctable2!Row & ""))
            aInsert2 = AddFlag(aInsert2, "COST", Val(loctable2!COST & ""))
            aInsert2 = AddFlag(aInsert2, "S_OKAZ", Val(loctable2!S_OKAZ & ""))
            If cStr2 = "" Then
                cStr2 = addInsert(aInsert2, "file6_20")
            Else
                cStr2 = cStr2 & " ; " & addInsert(aInsert2, "file6_20")
            End If
        Else
'            MsgBox " „—«Ã⁄… „»Ì⁄«  »«—ﬂÊœ " & sitem
            lSkip = True
        End If
        loctable2.MoveNext
    Loop
    If Not lSkip Then
        If cStrNew = "" Then
            cStrNew = " update file6_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no)
        Else
            cStrNew = cStrNew & " ;  update file6_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no)
        End If
    End If
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
If cDelStr1 <> "" Then
    conShop.Execute cDelStr1
    conShop.Execute cDelStr2
End If
If cStr1 <> "" Then
    conShop.Execute cStr1
    conShop.Execute cStr2
End If
If cStrNew <> "" Then
    con.Execute cStrNew
End If
Inform " „»Ì⁄«  " & nRecord
SendSales = nAffectTotal


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
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "E_MAIL", addstring(loctable!E_MAIL))
    aInsert = AddFlag(aInsert, "branch", addstring(loctable!branch))
    If cStrPhone = "" Then
        cStrPhone = addInsert(aInsert, "SUBCUST")
    Else
        cStrPhone = cStrPhone & " ; " & addInsert(aInsert, "SUBCUST")
    End If
    loctable.MoveNext
Loop
If cStrPhone <> "" Then
    conShop.Execute cStrPhone
    con.Execute " UPDATE SUBCUST SET isnew = 0 "
End If
Inform " phone : " & nRecord
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


Private Function SendCharge() As Long
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
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranch))
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
        aInsert2 = AddFlag(aInsert2, "[DESCA]", addstring(loctable2!DESCA))
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
SendCharge = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendCharge = -1
GoTo lastsub
End Function

Private Function SendTransBox() As Long
Set loctable = New ADODB.Recordset
Dim cDoc As String

cString = "Select * from FILE0_51 WHERE isnew = 1"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Do Until loctable.EOF
    cDoc = pBranch & loctable!code
    conShop.Execute " delete from file0_52  where CODE = " & MyParn(cDoc)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(cDoc))
    aInsert = AddFlag(aInsert, "NO1", addstring(loctable!no1))
    aInsert = AddFlag(aInsert, "NO2", addstring(loctable!no2))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(loctable!Value))
                                                                                
    conShop.Execute addInsert(aInsert, "FILE0_52"), nAffect
    con.Execute " update FILE0_51 set isnew = 0 where CODE = " & MyParn(loctable!code)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
SendTransBox = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendTransBox = -1
GoTo lastsub
End Function
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
