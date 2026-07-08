VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Begin VB.Form CopyData 
   Caption         =   "äÞá ÇáÈíÇäÇÊ"
   ClientHeight    =   1710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1710
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   240
      Left            =   135
      TabIndex        =   1
      Top             =   1035
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   423
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command1 
      Caption         =   "COPY DATA"
      Height          =   780
      Left            =   90
      TabIndex        =   0
      Top             =   225
      Width           =   4380
   End
End
Attribute VB_Name = "CopyData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub Copyfact()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"



cString = "SELECT * FROM FILE1_70 WHERE FLAG = 3 AND NOT CODE IS NULL "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(2, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = addstring(loctable!CODE)
    
    If Not IsNull(loctable!DESCA) Then
        cDesca = Replace(loctable!DESCA, "'", " ")
    Else
        cDesca = ""
    End If
    
    aInsert(1, 0) = "Desca"
    aInsert(1, 1) = addstring(cDesca)
           
'   aInsert(2, 0) = "[RATE]"
'   aInsert(2, 1) = Val(loctable!Rate & "")
           
    con.Execute CreateInsert(aInsert, "FACT")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub


Private Sub CopyMAN()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

con.Execute "DELETE FROM FILE6_25"

cString = "SELECT * FROM FILE1_70 WHERE FLAG = 9 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = addstring(loctable!CODE)
                  
    aInsert(1, 0) = "Desca"
    aInsert(1, 1) = addstring(loctable!DESCA)
           
           
    con.Execute CreateInsert(aInsert, "FILE6_25")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

Private Sub CopyGroupSupp()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

con.Execute "DELETE FROM FILE4_50"

cString = "SELECT * FROM FILE1_70 WHERE FLAG = 12 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = addstring(loctable!CODE)
                  
    aInsert(1, 0) = "Desca"
    aInsert(1, 1) = addstring(loctable!DESCA)
           
           
    con.Execute CreateInsert(aInsert, "FILE4_50")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub Copysup()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = "SELECT * FROM FILE4_10"
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText


nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE4_22"
con.Execute "DELETE FROM FILE4_10"
Dim aInsert(11, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Code"
    aInsert(0, 1) = addstring(loctable!CODE)
    
    aInsert(1, 0) = "[Desca]"
    aInsert(1, 1) = addstring(loctable!DESCA)
    
    aInsert(2, 0) = "Manager"
    aInsert(2, 1) = addstring(loctable!Manager & "")
    
    aInsert(3, 0) = "Phone1"
    aInsert(3, 1) = addstring(loctable!PHONE1 & "")
    
    aInsert(4, 0) = "Fax"
    aInsert(4, 1) = addstring(loctable!Fax & "")
    
    aInsert(5, 0) = "[Group]"
    aInsert(5, 1) = addstring(loctable!Group)
       
    aInsert(6, 0) = "Address"
    aInsert(6, 1) = addstring(loctable!Address & "")
           
    aInsert(7, 0) = "rate"
    aInsert(7, 1) = addvalue(loctable!Rate & "")
           
    aInsert(8, 0) = "disc"
    aInsert(8, 1) = addvalue(loctable!disc & "")
           
    aInsert(9, 0) = "SUBCODE"
    aInsert(9, 1) = addstring(loctable!SUBCODE)
           
    aInsert(10, 0) = "fact"
    If Not IsNull(loctable!Fact) Then
        aInsert(10, 1) = 1
    Else
        aInsert(10, 1) = 0
    End If
    
    aInsert(11, 0) = "SUPP"
    If IsNull(loctable!Fact) Then
        aInsert(11, 1) = 1
    Else
        aInsert(11, 1) = 0
    End If
           
    con.Execute CreateInsert(aInsert, "FILE4_10")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub Copysup2()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = "SELECT * FROM FILE4_22"
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText


nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE4_22"
Dim aInsert(3, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Code"
    aInsert(0, 1) = addstring(loctable!CODE)
    
    aInsert(1, 0) = "MOSM"
    aInsert(1, 1) = addstring(loctable!MOSM)
    
    aInsert(2, 0) = "F_BAL"
    aInsert(2, 1) = addvalue(loctable!bal & "")
    
           
    aInsert(3, 0) = "DATE"
    aInsert(3, 1) = addDate(loctable!Date & "")
           
    If GetDesca("SELECT CODE FROM FILE4_10 WHERE CODE = " & MyParn(loctable!CODE), con) <> "" Then
        con.Execute CreateInsert(aInsert, "FILE4_22")
    End If
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

Private Sub CopyMosm()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = "SELECT * FROM MOSM WHERE MOSM is not null "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText


nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM MOSM"
Dim aInsert(2, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "MOSM"
    aInsert(0, 1) = addstring(loctable!MOSM)
    
    aInsert(1, 0) = "DATE"
    aInsert(1, 1) = addDate(loctable!Date & "")
    
    aInsert(2, 0) = "DESCA"
    aInsert(2, 1) = addstring(loctable!MOSM)
                   
    con.Execute CreateInsert(aInsert, "MOSM")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

Private Sub CopyCust()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = "SELECT * FROM FILE3_10 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText


Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE3_10"
Dim aInsert(4, 1)
    
Do Until loctable.EOF
    i = i + 1
'   nValue = Round(i / (nRecordCount), 2) * 100
'   prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Code"
    aInsert(0, 1) = addstring(loctable!CODE)
    
    aInsert(1, 0) = "[Desca]"
    aInsert(1, 1) = addstring(loctable!DESCA)
    
    aInsert(2, 0) = "F_balance"
    aInsert(2, 1) = 0
    
    aInsert(3, 0) = "F_DATE"
    aInsert(3, 1) = addDate("1-1-2010")
    
    aInsert(4, 0) = "[GROUP]"
    aInsert(4, 1) = addstring("01")
    
    con.Execute CreateInsert(aInsert, "FILE3_10")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
    
End Sub


Private Sub CopyCust2()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = "SELECT * FROM FILE3_10 WHERE CardCust "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE3_10"
Dim aInsert(4, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Code"
    aInsert(0, 1) = addstring(loctable!CODE)
    
    aInsert(1, 0) = "[Desca]"
    aInsert(1, 1) = addstring(loctable!DESCA)
    
    aInsert(2, 0) = "[PHONE1]"
    aInsert(2, 1) = addstring(loctable!PHONE1)
    
    aInsert(3, 0) = "[F_DATE]"
    aInsert(3, 1) = addDate(loctable!F_DATE)
    
    aInsert(4, 0) = "[ADDRESS]"
    aInsert(4, 1) = addstring(loctable!Address)
    
    
    con.Execute CreateInsert(aInsert, "FILE3_20")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyItems()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"
cString = "SELECT * FROM FILE1_10 WHERE item > 100  "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText
Dim cModel As String
Dim cModelNo As String
conmdb.Execute " UPDATE FILE1_10 SET MOSM = '000' WHERE MOSM = '.000'"
nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double


con.Execute "DELETE FROM FILE1_10"
Dim aInsert(19, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
' **********************************
    cModel = loctable!Fact & loctable!MOSM & RetZero(loctable!modelfact, 4)
    If IsNull(loctable!SUPP) Then
        cModelNo = cModel
    Else
        cModelNo = cModel & loctable!CODE
    End If
    aInsert(0, 0) = "Model"
    cModel = Replace(cModel, "'", "0")
    aInsert(0, 1) = addstring(cModel)
    
    aInsert(1, 0) = "ModelNO"
    cModelNo = Replace(cModelNo, "'", "0")
    aInsert(1, 1) = addstring(cModelNo)
    
    aInsert(2, 0) = "Desca"
    If Trim(loctable!DESCA & "") = "" Then
        cDesca = "----"
    Else
        cDesca = Replace(loctable!DESCA, "'", " ")
    End If
    aInsert(2, 1) = addstring(cDesca)
    
    aInsert(3, 0) = "FACT"
    aInsert(3, 1) = addstring(loctable!Fact)
    
    aInsert(4, 0) = "MOSM"
    If IsNull(loctable!MOSM) Then
        aInsert(4, 1) = addstring("000")
    Else
        aInsert(4, 1) = addstring(loctable!MOSM)
    End If
    
    aInsert(5, 0) = "MODELFACT"
    cModelFact = Replace(loctable!modelfact, "'", "0")
    aInsert(5, 1) = addstring(cModelFact)
    
    aInsert(6, 0) = "MODELFACT0"
    aInsert(6, 1) = addstring(DelZero(cModelFact))
    
    aInsert(7, 0) = "SUPP"
    aInsert(7, 1) = addstring(loctable!SUPP)

    aInsert(8, 0) = "OKAZ"
    aInsert(8, 1) = Val(loctable!okaz & "")
    
    aInsert(9, 0) = "RATE"
    aInsert(9, 1) = Val(loctable!Rate & "")
    
    aInsert(10, 0) = "Code"
    aInsert(10, 1) = addstring(loctable!CODE)
 
    aInsert(11, 0) = "[group]"
    aInsert(11, 1) = addstring(loctable!Group)
    

    aInsert(12, 0) = "item"
    aInsert(12, 1) = addvalue(Trim(loctable!Item))
        
    aInsert(13, 0) = "Cost"
    aInsert(13, 1) = Val(loctable!cost & "")
        
    aInsert(14, 0) = "Price"
    aInsert(14, 1) = Val(loctable!price & "")
        
    aInsert(15, 0) = "Price2"
    aInsert(15, 1) = Val(loctable!PRICE2 & "")

    aInsert(16, 0) = "Scal"
    aInsert(16, 1) = addstring(loctable!scal)

    aInsert(17, 0) = "C_Scal"
    aInsert(17, 1) = addvalue(loctable!C_SCAL)
    
    aInsert(18, 0) = "Color"
    If Trim(loctable!Color & "") = "" Then
        aInsert(18, 1) = addstring("ãæÍÏ")
    Else
        aInsert(18, 1) = addstring(loctable!Color)
    End If

    aInsert(19, 0) = "C_Color"
    aInsert(19, 1) = IIf(IsNull(loctable!c_Color), 1, loctable!c_Color)
    
' *************************
    
    con.Execute CreateInsert(aInsert, "FILE1_10")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
DelItem
FixItem
End Sub
Private Sub CopyItems2()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"


cString = "SELECT * FROM FILE0_70"
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE1_10SC"
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Code"
    aInsert(0, 1) = addstring(loctable!CODE)
    
    aInsert(1, 0) = "desca"
    aInsert(1, 1) = addstring(loctable!DESCA)
           
    con.Execute CreateInsert(aInsert, "FILE1_10SC")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyGroup()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

con.Execute "DELETE FROM FILE1_50"

cString = "SELECT * FROM FILE1_70 WHERE FLAG = 2 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = addstring(loctable!CODE)
                  
    aInsert(1, 0) = "Desca"
    aInsert(1, 1) = addstring(loctable!DESCA)
           
           
    con.Execute CreateInsert(aInsert, "FILE1_50")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopySales1()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
Dim cDoc As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"
conmdb.Execute " UPDATE FILE6_20 SET CODE = '0000' WHERE CODE = '2001' "

cString = " SELECT FILE6_20.DOC_NO,  store2 as STORE , Min(FILE6_20.DATE) AS DATESAL , Min(FILE6_20.CODE) AS CODE  From FILE6_20 GROUP BY FILE6_20.DOC_NO  , STORE2 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE6_20H"
Dim aInsert(5, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    cDoc = RetZero(Day(loctable!DATESAL), 2) & RetZero(Month(loctable!DATESAL), 2) & Mid(Year(loctable!DATESAL), 3, 2) & Mid(loctable!doc_no, 1, 1) & Mid(loctable!doc_no, 5, 3)
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(cDoc)


    aInsert(1, 0) = "code"
    aInsert(1, 1) = addstring(loctable!CODE)
    
    aInsert(2, 0) = "[Date]"
    aInsert(2, 1) = addDate(Format(loctable!DATESAL, "DD-MM-YYYY"))

    aInsert(3, 0) = "store"
    aInsert(3, 1) = addstring(loctable!STORE)

    aInsert(4, 0) = "PRINTED"
    aInsert(4, 1) = 1
           
    aInsert(5, 0) = "BOX"
    aInsert(5, 1) = addstring(loctable!STORE)
           
    
    con.Execute CreateInsert(aInsert, "FILE6_20H")
    
    Err.Clear
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopySalesDisc()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"


cString = " SELECT FILE6_20.DOC_NO, STORE2 , DATE ,ITEM , [TOTAL]*-1 AS DISCOUNT FROM FILE6_20 WHERE STORE = 'ZZ' AND TOTAL <> 0 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(3, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    cDoc = RetZero(Day(loctable!Date), 2) & RetZero(Month(loctable!Date), 2) & Mid(Year(loctable!Date), 3, 2) & Mid(loctable!doc_no, 1, 1) & Mid(loctable!doc_no, 5, 3)
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(cDoc)

    aInsert(1, 0) = "DISCOUNT"
    aInsert(1, 1) = loctable!discount
    
    con.Execute CreateUpdate(aInsert, "FILE6_20H", " where doc_no = " & addstring(cDoc))
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopySales2()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = "SELECT * FROM FILE6_20 WHERE STORE <> 'ZZ' "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE6_20"
Dim aInsert(6, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
'                   2                           2                                       2                                   1                                      3
    cDoc = RetZero(Day(loctable!Date), 2) & RetZero(Month(loctable!Date), 2) & Mid(Year(loctable!Date), 3, 2) & Mid(loctable!doc_no, 1, 1) & Mid(loctable!doc_no, 5, 3)
    
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(cDoc)
    
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addstring(loctable!Item)
    
    aInsert(2, 0) = "MAN"
    aInsert(2, 1) = addstring(loctable!MAN)
    
    aInsert(3, 0) = "quant"
    aInsert(3, 1) = Val(loctable!Quant)

    aInsert(4, 0) = "Price"
    aInsert(4, 1) = Val(loctable!price & "")

    aInsert(5, 0) = "Price_C"
    aInsert(5, 1) = Val(loctable!PRICE_C & "")

    aInsert(6, 0) = "Cost"
    aInsert(6, 1) = Val(loctable!cost & "")
    
    con.Execute CreateInsert(aInsert, "FILE6_20")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyPurchase1()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT FILE7_20.DOC_NO, Min(FILE7_20.STORE) AS STORE, max(FILE7_20.inv_no) AS inv_no , max(FILE7_20.remark ) AS remark , Min(FILE7_20.DATE) AS DATESAL , Min(FILE7_20.mosm) AS mosm, Min(FILE7_20.CODE) AS CODE  From FILE7_20 GROUP BY FILE7_20.DOC_NO"
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE7_20H"
Dim aInsert(6, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)
    
    aInsert(1, 0) = "code"
    aInsert(1, 1) = addstring(loctable!CODE)
    
    aInsert(2, 0) = "[Date]"
    aInsert(2, 1) = addDate(Format(loctable!DATESAL, "dd-mm-yyyy"))
    
    aInsert(3, 0) = "store"
    aInsert(3, 1) = addstring(loctable!STORE)
    
    aInsert(4, 0) = "mosm"
    aInsert(4, 1) = addstring(loctable!MOSM & "")
    
    aInsert(5, 0) = "inv_no"
    aInsert(5, 1) = addstring(loctable!INV_NO & "")
    
    aInsert(6, 0) = "remark"
    aInsert(6, 1) = addstring(loctable!remark & "")
    
    con.Execute CreateInsert(aInsert, "FILE7_20H")
    
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyPurchase2()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = "SELECT * FROM FILE7_20 where store <> 'zz' and store <> 'ss' "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE7_20"
Dim aInsert(3, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(loctable!doc_no & "")
    
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addstring(loctable!Item & "")
    
    aInsert(2, 0) = "quant"
    aInsert(2, 1) = Val(loctable!Quant & "")

    aInsert(3, 0) = "Price"
    aInsert(3, 1) = Val(loctable!price & "")

           
    con.Execute CreateInsert(aInsert, "FILE7_20")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

Private Sub CopyPurchDisc()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT STORE , FILE7_20.DOC_NO, [TOTAL]*-1 AS DISCOUNT , TOTAL AS TAX FROM FILE7_20 WHERE ( FILE7_20.STORE ='ZZ' OR FILE7_20.STORE = 'SS' ) AND TOTAL <> 0 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    If loctable!STORE = "SS" Then
        con.Execute " UPDATE FILE7_20H SET TAX = " & Val(loctable!tax & "") & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    End If
    
    If loctable!STORE = "zz" Then
        con.Execute " UPDATE FILE7_20H SET DISCOUNT = " & loctable!discount & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    End If
    
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

Private Sub CopyPurchase11()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"


cString = " SELECT FILE6_11.DOC_NO, Min(FILE6_11.STORE) AS STORE, max(FILE6_11.inv_no) AS inv_no , max(FILE6_11.remark ) AS remark , Min(FILE6_11.DATE) AS DATESAL , Min(FILE6_11.mosm) AS mosm, Min(FILE6_11.CODE) AS CODE  From FILE6_11 GROUP BY FILE6_11.DOC_NO"
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE7_10H"
Dim aInsert(6, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)
    
    aInsert(1, 0) = "code"
    aInsert(1, 1) = addstring(loctable!CODE)
    
    aInsert(2, 0) = "[Date]"
    aInsert(2, 1) = addDate(Format(loctable!DATESAL, "dd-mm-yyyy"))
    
    aInsert(3, 0) = "store"
    aInsert(3, 1) = addstring(loctable!STORE)
    
    aInsert(4, 0) = "mosm"
    aInsert(4, 1) = addstring(loctable!MOSM & "")
    
    aInsert(5, 0) = "inv_no"
    aInsert(5, 1) = addstring(loctable!INV_NO & "")
    
    aInsert(6, 0) = "remark"
    aInsert(6, 1) = addstring(loctable!remark & "")
    
           
    con.Execute CreateInsert(aInsert, "FILE7_10H")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyPurchase22()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = "SELECT * FROM FILE6_11 where store <> 'zz' and store <> 'ss' "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE7_10"
Dim aInsert(3, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(loctable!doc_no & "")
    
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addstring(loctable!Item & "")
    
    aInsert(2, 0) = "quant"
    aInsert(2, 1) = Val(loctable!Quant & "")

    aInsert(3, 0) = "Price"
    aInsert(3, 1) = Val(loctable!price & "")

    
    con.Execute CreateInsert(aInsert, "FILE7_10")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyPurchDisc2()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT FILE6_11.STORE , FILE6_11.DOC_NO, [TOTAL]*-1 AS DISCOUNT , TOTAL AS TAX FROM FILE6_11 WHERE ( STORE ='ZZ' OR STORE = 'SS' ) AND TOTAL <> 0 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    If loctable!STORE = "SS" Then
        con.Execute " UPDATE FILE7_10H SET TAX = " & loctable!tax & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    End If
    
    If loctable!STORE = "zz" Then
        con.Execute " UPDATE FILE7_10H SET DISCOUNT = " & loctable!discount & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    End If
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

Private Sub CopyPurchDisc1()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT FILE7_20.STORE , FILE7_20.DOC_NO, [TOTAL]*-1 AS DISCOUNT , TOTAL AS TAX FROM FILE7_20 WHERE ( STORE ='ZZ' OR STORE = 'SS' ) AND TOTAL <> 0 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    If loctable!STORE = "SS" Then
        con.Execute " UPDATE FILE7_20H SET TAX = " & loctable!tax & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    End If
    
    If loctable!STORE = "zz" Then
        con.Execute " UPDATE FILE7_20H SET DISCOUNT = " & loctable!discount & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    End If
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub Command2_Click()
    fixcost2
End Sub
Private Sub Command1_Click()
    If InputBox("", "") = "123456" Then
'        If MsgBox("CopySales1", vbYesNo + vbDefaultButton2) = vbYes Then CopySales1
        If MsgBox("CopySalesDisc", vbYesNo + vbDefaultButton2) = vbYes Then CopySalesDisc
'        If MsgBox("CopySales2", vbYesNo + vbDefaultButton2) = vbYes Then CopySales2
'        con.Execute " UPDATE FILE6_20H SET BRANCH = '01' WHERE STORE = '2' "
'        con.Execute " UPDATE FILE6_20H SET BRANCH = '02' WHERE STORE = '5' "
'        con.Execute " UPDATE FILE6_20H SET BRANCH = '00' WHERE BRANCH IS NULL "
        con.Execute " UPDATE    FILE6_20H SET              CASH = T_SALESDOC.t_total - FILE6_20H.discount FROM         FILE6_20H INNER JOIN                       T_SALESDOC ON FILE6_20H.DOC_NO = T_SALESDOC.DOC_NO "
        MsgBox " O K "
    End If
Exit Sub


con.Execute "DELETE FROM fact "
con.Execute "DELETE FROM FILE0_40"
con.Execute "DELETE FROM FILE0_50"

con.Execute "DELETE FROM FILE8_10"
con.Execute "DELETE FROM FILE8_10H"

con.Execute "DELETE FROM FILE8_20"
con.Execute "DELETE FROM FILE8_20H"

con.Execute "DELETE FROM FILE8_30"
con.Execute "DELETE FROM FILE8_30H"

con.Execute "DELETE FROM FILE8_40"
con.Execute "DELETE FROM FILE8_40H"

con.Execute "DELETE FROM FILE8_50"
con.Execute "DELETE FROM FILE8_50H"

con.Execute "DELETE FROM FILE8_00"
con.Execute "DELETE FROM FILE8_00H"

con.Execute "DELETE FROM FILE8_50"
con.Execute "DELETE FROM FILE8_50H"

con.Execute "DELETE FROM FILE8_70"
con.Execute "DELETE FROM FILE8_70H"

con.Execute "DELETE FROM FILE8_71"
con.Execute "DELETE FROM FILE0_60"

con.Execute "DELETE FROM FILE0_10"
con.Execute "DELETE FROM FILE0_100"
con.Execute "DELETE FROM FILE0_10H"

con.Execute "DELETE FROM FILE1_60"
con.Execute "DELETE FROM FILE1_60H"

con.Execute "DELETE FROM FILE6_20"
con.Execute "DELETE FROM FILE6_20H"

con.Execute "DELETE FROM FILE7_20"
con.Execute "DELETE FROM FILE7_20H"

con.Execute "DELETE FROM FILE7_10"
con.Execute "DELETE FROM FILE7_10H"

con.Execute "DELETE FROM FILE4_22"
con.Execute "DELETE FROM FILE3_20"

con.Execute "DELETE FROM FILE0_100"
con.Execute "DELETE FROM FILE0_10"
con.Execute "DELETE FROM FILE0_10H"

con.Execute "DELETE FROM FILE1_11"

con.Execute "DELETE FROM FILE1_10"
con.Execute "DELETE FROM FILE4_10"
con.Execute "DELETE FROM FILE3_10"

If MsgBox("Copyfact", vbYesNo + vbDefaultButton2) = vbYes Then Copyfact
If MsgBox("CopyMOSM", vbYesNo + vbDefaultButton2) = vbYes Then CopyMosm
If MsgBox("CopyMAN", vbYesNo + vbDefaultButton2) = vbYes Then CopyMAN
If MsgBox("CopyGroup", vbYesNo + vbDefaultButton2) = vbYes Then CopyGroup
If MsgBox("CopyGroupSupp", vbYesNo + vbDefaultButton2) = vbYes Then CopyGroupSupp
If MsgBox("Copysup", vbYesNo + vbDefaultButton2) = vbYes Then Copysup
If MsgBox("Copysup2", vbYesNo + vbDefaultButton2) = vbYes Then Copysup2
If MsgBox("CopyCust", vbYesNo + vbDefaultButton2) = vbYes Then CopyCust

If MsgBox("CopyItems", vbYesNo + vbDefaultButton2) = vbYes Then CopyItems

If MsgBox("CopyTrans", vbYesNo + vbDefaultButton2) = vbYes Then CopyTrans
If MsgBox("CopyTrans2", vbYesNo + vbDefaultButton2) = vbYes Then CopyTrans2

If MsgBox("CopyStock", vbYesNo + vbDefaultButton2) = vbYes Then CopyStock
If MsgBox("CopyStock2", vbYesNo + vbDefaultButton2) = vbYes Then CopyStock2

If MsgBox("CopySales1", vbYesNo + vbDefaultButton2) = vbYes Then CopySales1
If MsgBox("CopySalesDisc", vbYesNo + vbDefaultButton2) = vbYes Then CopySalesDisc
If MsgBox("CopySales2", vbYesNo + vbDefaultButton2) = vbYes Then CopySales2

If MsgBox("CopyPurchase1", vbYesNo + vbDefaultButton2) = vbYes Then CopyPurchase1
If MsgBox("CopyPurchase2", vbYesNo + vbDefaultButton2) = vbYes Then CopyPurchase2
If MsgBox("CopyPurchDisc1", vbYesNo + vbDefaultButton2) = vbYes Then CopyPurchDisc1

If MsgBox("CopyPurchase11", vbYesNo + vbDefaultButton2) = vbYes Then CopyPurchase11
If MsgBox("CopyPurchase22", vbYesNo + vbDefaultButton2) = vbYes Then CopyPurchase22
If MsgBox("CopyPurchDisc2", vbYesNo + vbDefaultButton2) = vbYes Then CopyPurchDisc2

If MsgBox("Copystore", vbYesNo + vbDefaultButton2) = vbYes Then CopyStore
If MsgBox("Copybox", vbYesNo + vbDefaultButton2) = vbYes Then Copybox



If MsgBox("CopyCASH_10H", vbYesNo + vbDefaultButton2) = vbYes Then CopyCash_10H
If MsgBox("CopyCASH_10", vbYesNo + vbDefaultButton2) = vbYes Then CopyCash_10


If MsgBox("CopyCASH_20H", vbYesNo + vbDefaultButton2) = vbYes Then CopyCash_20H
If MsgBox("CopyCASH_20", vbYesNo + vbDefaultButton2) = vbYes Then CopyCash_20


If MsgBox("CopyCASH_50H", vbYesNo + vbDefaultButton2) = vbYes Then CopyCash_50H
If MsgBox("CopyCASH_50", vbYesNo + vbDefaultButton2) = vbYes Then CopyCash_50
If MsgBox("CopyCash_501", vbYesNo + vbDefaultButton2) = vbYes Then CopyCash_501



If MsgBox("Copyparth", vbYesNo + vbDefaultButton2) = vbYes Then CopypartH
If MsgBox("Copypart", vbYesNo + vbDefaultButton2) = vbYes Then Copypart
If MsgBox("Copypartcode", vbYesNo + vbDefaultButton2) = vbYes Then Copypartcode
If MsgBox("Copytransbox", vbYesNo + vbDefaultButton2) = vbYes Then Copytransbox

MsgBox " E N D "
End Sub

Private Sub Form_Load()
    openCon con
End Sub
Sub DelItem()
    On Error Resume Next
    con.Execute " DROP TABLE DEL_ITEM"
End Sub
Sub FixItem()
Dim loctable As New ADODB.Recordset
Dim nItem As Double
cStr1 = " SELECT     MODEL, SCAL, C_COLOR, COUNT(ITEM) AS Expr1 Into DEL_ITEM From FILE1_10 GROUP BY MODEL, SCAL, C_COLOR Having (Count(Item) > 1) "
con.Execute cStr1


cString = "SELECT * FROM DEL_ITEM"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    nItem = Val(GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(loctable!MODEL) & " AND SCAL = " & MyParn(loctable!scal) & " AND C_COLOR = " & loctable!c_Color) & "")
    
    con.Execute " UPDATE FILE1_10 SET C_COLOR = 3     WHERE ITEM = " & nItem
    loctable.MoveNext
    Exit Do
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub


Private Sub CopyTrans()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT FILE1_60.DOC_NO, FILE1_60.STORE1 , MAX(FILE1_60.DATE  ) AS DATE2 ,FILE1_60.STORE2 From FILE1_60 GROUP BY FILE1_60.DOC_NO, FILE1_60.STORE1 , FILE1_60.STORE2"
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE1_60H"
Dim aInsert(3, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)


    aInsert(1, 0) = "[Date]"
    aInsert(1, 1) = addDate(Format(loctable!date2, "DD-MM-YYYY"))

    aInsert(2, 0) = "store1"
    aInsert(2, 1) = addstring(loctable!STORE1)

    aInsert(3, 0) = "store2"
    aInsert(3, 1) = addstring(loctable!STORE2)

           
    con.Execute CreateInsert(aInsert, "FILE1_60H")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyTrans2()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = "SELECT FILE1_60.* FROM FILE1_60 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_60.ITEM "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE1_60"
Dim aInsert(2, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(loctable!doc_no & "")
    
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addvalue(loctable!Item & "")
    
    aInsert(2, 0) = "quant"
    aInsert(2, 1) = Val(loctable!Quant & "")


           
    con.Execute CreateInsert(aInsert, "FILE1_60")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyCash_10H()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT FILE8_10.DOC_NO, MAX(FILE8_10.DATE  ) AS DATE2 From FILE8_10 GROUP BY FILE8_10.DOC_NO, FILE8_10.DATE "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE8_10"
con.Execute "DELETE FROM FILE8_10H"
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)


    aInsert(1, 0) = "[Date]"
    aInsert(1, 1) = addDate(Format(loctable!date2, "DD-MM-YYYY"))

           
    con.Execute CreateInsert(aInsert, "FILE8_10H")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyCash_10()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT * FROM FILE8_10  "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE8_10"
Dim aInsert(4, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)


    aInsert(1, 0) = "[CODE]"
    aInsert(1, 1) = addstring(loctable!CODE)

    aInsert(2, 0) = "VALUE"
    aInsert(2, 1) = addvalue(loctable!Value)

           
    aInsert(3, 0) = "DESCA"
    aInsert(3, 1) = addstring(loctable!DESCA)
           
    aInsert(4, 0) = "box"
    aInsert(4, 1) = addstring(loctable!BOX)
           
           
    con.Execute CreateInsert(aInsert, "FILE8_10")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub fixcost2()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"


cString = "SELECT * FROM FILE1_10 WHERE ITEM > 100 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(0, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
' **********************************

        

    aInsert(0, 0) = "cost2"
    aInsert(0, 1) = loctable!fixprice

    
' *************************
    
    cString = CreateUpdate(aInsert, "file1_10", " WHERE item = " & MyParn(loctable!Item))
    con.Execute cString
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
DelItem
FixItem
End Sub




Private Sub CopyItemsSEC()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"
cString = "SELECT * FROM FILE1_10 WHERE ( MID(FILE1_10.MOSM,1,1) = 'S' OR MID(FILE1_10.MOSM,1,1) = 'W' OR MID(FILE1_10.MOSM,1,1) = '0' )  AND MID(FILE1_10.MOSM,2,2) <= '13' AND MOSM IS NOT NULL AND MODELFACT   IS NOT NULL AND FACT IS NOT  NULL  "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText
Dim cModel As String
Dim cModelNo As String

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double


Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
' **********************************
    
    aInsert(0, 0) = "ITEM"
    aInsert(0, 1) = loctable!Item
    
    aInsert(1, 0) = "SEC"
    aInsert(1, 1) = addstring(loctable!SEC)
' *************************
    
    con.Execute CreateUpdate(aInsert, "FILE1_10", " where item = " & Val(loctable!Item))
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopySalesCUST()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT FILE6_22.DOC_NO, FILE6_22.CODE FROM FILE6_22 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)

   If Trim(loctable!doc_no) = "2412120052" Then
       MsgBox "AA"
   End If
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(Trim(loctable!doc_no))

'    If loctable!CODE = "10850" Then
'        MsgBox loctable!doc_no
'    End If
    aInsert(1, 0) = "CODE2"
    aInsert(1, 1) = addstring(Trim(loctable!CODE))

    If Len(loctable!CODE) <= 6 Then
        con.Execute CreateUpdate(aInsert, "FILE6_20H", " where doc_no = " & MyParn(Trim(loctable!doc_no)))
    End If
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyCash_50H()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"




cString = " SELECT FILE8_50.DOC_NO , FILE8_50.DATE  From FILE8_50 WHERE DATE IS NOT NULL GROUP BY DATE , FILE8_50.DOC_NO "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
'con.Execute "DELETE FROM FILE8_50"
'con.Execute "DELETE FROM FILE8_50H"
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)

    aInsert(1, 0) = "[Date]"
    aInsert(1, 1) = addDate(Format(loctable!Date, "DD-MM-YYYY"))
           
    con.Execute CreateInsert(aInsert, "FILE8_50H")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyCash_50()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT * FROM FILE8_50  WHERE DATE IS NOT NULL "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
'con.Execute "DELETE FROM FILE8_50"
Dim aInsert(7, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)

    aInsert(1, 0) = "[CHARGE]"
    aInsert(1, 1) = addstring(loctable!CHARGE)

    aInsert(2, 0) = "VALUE"
    aInsert(2, 1) = addvalue(loctable!Value)

    aInsert(3, 0) = "DESCA"
    aInsert(3, 1) = addstring(loctable!DESCA)
           
    aInsert(4, 0) = "BOX"
    aInsert(4, 1) = addstring(loctable!BOX)
           
    aInsert(5, 0) = "ser_no"
    aInsert(5, 1) = addstring(loctable!Ser_no)
           
    con.Execute CreateInsert(aInsert, "FILE8_50")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyCash_501()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT * FROM FILE8_70  "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE8_51"
Dim aInsert(2, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "code"
    aInsert(0, 1) = addstring(loctable!CODE)

    aInsert(1, 0) = "[desca]"
    aInsert(1, 1) = addstring(loctable!DESCA)

    aInsert(2, 0) = "[MAINGROUP]"
    aInsert(2, 1) = addstring(loctable!MainGroup)

    con.Execute CreateInsert(aInsert, "FILE8_51")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub


Private Sub CopyID_USER()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT * FROM ID_USER"
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM ID_USER "
Dim aInsert(3, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "ID"
    aInsert(0, 1) = addstring(loctable!ID)


    aInsert(1, 0) = "[NAME]"
    aInsert(1, 1) = addstring(loctable!NAMEA)

    aInsert(2, 0) = "[JOB]"
    aInsert(2, 1) = addstring(loctable!JOB)

    aInsert(3, 0) = "[CODE]"
    aInsert(3, 1) = addvalue(loctable!CODE)

    con.Execute CreateInsert(aInsert, "ID_USER")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopySUBRET()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT DOC_NO , MAX(REM1)  AS REM1 , MAX(REM2)  AS REM2 , MAX(REM3)  AS REM3 , MAX(REM4) AS REM4 FROM SUB6_11 GROUP BY DOC_NO "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM SUB_RET"
Dim aInsert(4, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "DOC_NO"
    aInsert(0, 1) = addstring(loctable!doc_no)


    aInsert(1, 0) = "[REM1]"
    aInsert(1, 1) = addstring(Mid(Trim(loctable!REM1), 1, 50))

    aInsert(2, 0) = "[REM2]"
    aInsert(2, 1) = addstring(Mid(Trim(loctable!REM2), 1, 50))

    aInsert(3, 0) = "[REM3]"
    aInsert(3, 1) = addstring(Mid(Trim(loctable!REM3), 1, 50))

    aInsert(4, 0) = "[REM4]"
    aInsert(4, 1) = addstring(Mid(Trim(loctable!REM4), 1, 50))

    con.Execute CreateInsert(aInsert, "SUB_RET")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

Private Sub CopyStore()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"



cString = "SELECT * FROM FILE1_50 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(2, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = addstring(loctable!CODE)
    
    If Not IsNull(loctable!DESCA) Then
        cDesca = Replace(loctable!DESCA, "'", " ")
    Else
        cDesca = ""
    End If
    
    aInsert(1, 0) = "Desca"
    aInsert(1, 1) = addstring(cDesca)
           
    con.Execute CreateInsert(aInsert, "file0_40")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub Copybox()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"



cString = "SELECT * FROM FILE0_50 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(3, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = addstring(loctable!CODE)
    
    If Not IsNull(loctable!DESCA) Then
        cDesca = Replace(loctable!DESCA, "'", " ")
    Else
        cDesca = ""
    End If
    
    aInsert(1, 0) = "Desca"
    aInsert(1, 1) = addstring(cDesca)
           
    aInsert(2, 0) = "f_bal"
    aInsert(2, 1) = loctable!F_BAL
           
    aInsert(3, 0) = "f_date"
    aInsert(3, 1) = addDate(loctable!F_DATE)
           
    con.Execute CreateInsert(aInsert, "file0_50")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub


Private Sub CopyCash_20H()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"


cString = " SELECT FILE8_20.DOC_NO, MAX(FILE8_20.DATE  ) AS DATE2 From FILE8_20 WHERE CASH GROUP BY FILE8_20.DOC_NO, FILE8_20.DATE "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE8_20H"
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    
    
    aInsert(0, 0) = "Doc_No"
    If Len(loctable!doc_no) > 6 Then
        aInsert(0, 1) = addstring("." & Mid(loctable!doc_no, 3))
    Else
        aInsert(0, 1) = addstring(loctable!doc_no)
    End If

    aInsert(1, 0) = "[Date]"
    aInsert(1, 1) = addDate(Format(loctable!date2, "DD-MM-YYYY"))
    con.Execute CreateInsert(aInsert, "FILE8_20H")
    
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyCash_20()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT * FROM FILE8_20  WHERE CASH and doc_no is not null "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE8_20"
Dim aInsert(5, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    
    If Len(loctable!doc_no) > 6 Then
        aInsert(0, 1) = addstring("." & Mid(loctable!doc_no, 3))
    Else
        aInsert(0, 1) = addstring(loctable!doc_no)
    End If
    aInsert(0, 0) = "DOC_NO"

    aInsert(1, 0) = "[CODE]"
    aInsert(1, 1) = addstring(loctable!CODE)

    aInsert(2, 0) = "VALUE"
    aInsert(2, 1) = addvalue(loctable!Value)

    aInsert(3, 0) = "[MOSM]"
    aInsert(3, 1) = addstring(loctable!MOSM)
           
    aInsert(4, 0) = "[box]"
    aInsert(4, 1) = addstring(loctable!BOX)
           
    aInsert(5, 0) = "DESCA"
    aInsert(5, 1) = addstring(loctable!DESCA)
           
        
    aInsert(0, 0) = "DOC_NO"
    con.Execute CreateInsert(aInsert, "FILE8_20")
    
    
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub CopyStock()

Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT FILE0_10.DOC_NO, FILE0_10.STORE  , MAX(FILE0_10.DATE  ) AS DATE2 from FILE0_10 GROUP BY FILE0_10.DOC_NO, FILE0_10.STORE  "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE0_10"
con.Execute "DELETE FROM FILE0_10H"
Dim aInsert(3, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)


    aInsert(1, 0) = "[Date]"
    aInsert(1, 1) = addDate(Format(loctable!date2, "DD-MM-YYYY"))

    aInsert(2, 0) = "store"
    aInsert(2, 1) = addstring(loctable!STORE)

    aInsert(3, 0) = "ISCLOSED"
    aInsert(3, 1) = 1

    con.Execute CreateInsert(aInsert, "FILE0_10H")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

Private Sub CopyStock2()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset

conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = "SELECT FILE0_10.* FROM FILE0_10 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE0_10.ITEM where FILE0_10.ITEM > 100 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE0_100"
Dim aInsert(4, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "doc_no"
    aInsert(0, 1) = addstring(loctable!doc_no & "")
    
    aInsert(1, 0) = "item"
    aInsert(1, 1) = addstring(loctable!Item & "")
    
    aInsert(2, 0) = "computerBal"
    aInsert(2, 1) = Val(loctable!ComputerBal & "")

    aInsert(3, 0) = "DIFFER"
    aInsert(3, 1) = Val(loctable!Differ & "")

    aInsert(4, 0) = "realBal"
    aInsert(4, 1) = Val(loctable!RealBal & "")


    con.Execute CreateInsert(aInsert, "FILE0_10")
    con.Execute CreateInsert(aInsert, "FILE0_100")
    
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub



Private Sub CopypartH()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"




cString = " SELECT FILE8_90.DOC_NO , FILE8_90.DATE  From FILE8_90 WHERE DATE IS NOT NULL GROUP BY DATE , DOC_NO "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)

    aInsert(1, 0) = "[Date]"
    aInsert(1, 1) = addDate(Format(loctable!Date, "DD-MM-YYYY"))
           
    con.Execute CreateInsert(aInsert, "FILE8_70H")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

Private Sub Copypart()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT * FROM FILE8_90  WHERE DATE IS NOT NULL "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double

Dim aInsert(4, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "Doc_No"
    aInsert(0, 1) = addstring(loctable!doc_no)

    aInsert(1, 0) = "[code]"
    aInsert(1, 1) = addstring(loctable!MAN)

    aInsert(2, 0) = "VALUE_M"
    aInsert(2, 1) = addvalue(loctable!Value)

    aInsert(3, 0) = "DESCA"
    aInsert(3, 1) = addstring(loctable!DESCA)
           
    aInsert(4, 0) = "BOX"
    aInsert(4, 1) = addstring(loctable!BOX)
           
           
    con.Execute CreateInsert(aInsert, "FILE8_70")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub


Private Sub Copypartcode()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
Dim dTime As String
Dim cCode As String
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"

cString = " SELECT * FROM FILE1_70  where flag = 5 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
con.Execute "DELETE FROM FILE8_70"
Dim aInsert(1, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    aInsert(0, 0) = "code"
    aInsert(0, 1) = addstring(loctable!CODE)

    aInsert(1, 0) = "[desca]"
    aInsert(1, 1) = addstring(loctable!DESCA)

    con.Execute CreateInsert(aInsert, "FILE8_71")
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub
Private Sub Copytransbox()
Dim conmdb As New ADODB.Connection
Dim loctable As New ADODB.Recordset
conmdb.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\MDF\DATA.mdb"



cString = "SELECT * FROM FILE0_60 "
loctable.Open cString, conmdb, adOpenStatic, adLockReadOnly, adCmdText

nRecordCount = loctable.RecordCount
prog1.Value = 0
prog1.Visible = True
Dim i As Long, nValue As Double
Dim aInsert(5, 1)
Do Until loctable.EOF
    i = i + 1
    nValue = Round(i / (nRecordCount), 2) * 100
    prog1.Value = IIf(nValue > 100, 100, nValue)
    
    
    aInsert(0, 0) = "CODE"
    aInsert(0, 1) = addstring(loctable!CODE)
    
    aInsert(1, 0) = "Desca"
    aInsert(1, 1) = addstring(loctable!DESCA)
           
    aInsert(2, 0) = "VALUE"
    aInsert(2, 1) = loctable!Value
           
    aInsert(3, 0) = "DATE"
    aInsert(3, 1) = addDate(loctable!Date)
           
    
    aInsert(4, 0) = "NO1"
    aInsert(4, 1) = addstring(loctable!no1)
           
    aInsert(5, 0) = "NO2"
    aInsert(5, 1) = addstring(loctable!no2)
           
    con.Execute CreateInsert(aInsert, "file0_51")
    
    loctable.MoveNext
Loop
prog1.Visible = False
loctable.Close
Set loctable = Nothing
End Sub

