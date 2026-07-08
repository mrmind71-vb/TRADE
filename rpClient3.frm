VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpClient3 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   1845
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1845
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   " Ã„Ì⁄ «·›Ê« Ì—"
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
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1545
   End
   Begin VB.CommandButton CmdApply 
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
      Height          =   375
      Left            =   1215
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1395
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
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
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1395
      Width           =   1140
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   810
      Top             =   -45
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "data1"
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
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   6180
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   945
         Width           =   1680
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1680
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   3345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì· :"
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
         Left            =   4785
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   315
         Width           =   600
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   990
         Width           =   465
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   765
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "rpClient3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If TypeOf ActiveControl Is DBCombo And KeyCode = 46 Then ActiveControl.BoundText = ""
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xdate1.text) And Trim(xdate1.text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ”·Ì„"
    Exit Function
End If

If Not IsDate(xDate2.text) And Trim(xDate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim nBalance As Double, nRow As Integer
Dim cWhere As String
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cWhere = "code = " & MyParn(xCode.text)
If IsDate(xdate1.text) Then cWhere = cWhere & turn(cWhere, " and ") & "DATE >= " & DateSq(xdate1.text)
If IsDate(xdate1.text) Then cWhere = cWhere & turn(cWhere, " and ") & "DATE <= " & DateSq(xDate2.text)

cField11 = "(Select Sum(SAL-PAY) from File3_11 where type = '4' " & turn(cWhere, " and ") & cWhere & ")" & " as SalesTotal"
cField12 = "(Select Sum(PAY) from File3_11 where type = '5' " & turn(cWhere, " and ") & cWhere & ")" & " as SalesRetTotal"
cField13 = "(Select Sum(PAY) from File3_11 where (type = '10' or TYPE = '7') " & turn(cWhere, " and ") & cWhere & ")" & " as CashTotal"
cField14 = "(Select Sum(PAY) from File3_11 where (type = 'A') " & turn(cWhere, " and ") & cWhere & ")" & " as ChqTotal"


cWhere = "code1 = " & MyParn(xCode.text)
cField15 = "(Select Sum(VALUE) from File5_20 where (Closed = '0') " & turn(cWhere, " and ") & cWhere & ")" & " as ChqTotal2"

cWhere = ""
If Check1.Value = 1 Then
    If Trim(xCode.text) <> "" Then
        cWhere = cWhere & turn(cWhere, " and ") & " FILE3_11RP.code = " & MyParn(xCode.text)
    End If
              
    If IsDate(xdate1.text) Then
        cWhere = cWhere & turn(cWhere, " and ") & " FILE3_11RP.date < " & DateSq(xdate1.text)
    Else
        cWhere = cWhere & turn(cWhere, " and ") & " 1 = 2"
    End If
    
    cField1 = "(Select Sum(FILE3_11RP.Sal - FILE3_11RP.Pay) " & _
              " from FILE3_11RP " & _
              turnFound(cWhere, " where ") & _
              cWhere & ") as FirstBalance"
    
    cWhere = ""
    cString = "select FILE3_11RP.*,FILE3_12.desca as moveDesca, " & _
              cField1 & "," & _
              cField11 & "," & _
              cField12 & "," & _
              cField13 & "," & _
              cField14 & "," & _
              cField15 & _
              " From FILE3_11RP Left join FILE3_12 on FILE3_11RP.type = FILE3_12.code"
    If Trim(xCode.text) <> "" Then
        cWhere = cWhere & turnFound(cWhere, " and ") & " FILE3_11RP.code = " & MyParn(xCode.text)
        aHeader(0) = "[" & "··⁄„Ì· : " & xCodeDesca.Caption & "]"
    End If
              
    If IsDate(xdate1.text) Then
        cWhere = cWhere & turnFound(cWhere, " and ") & " FILE3_11RP.date >= " & DateSq(xdate1.text)
        aHeader(1) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    End If
    
    If IsDate(xDate2.text) Then
        cWhere = cWhere & turnFound(cWhere, " and ") & " FILE3_11RP.date <= " & DateSq(xDate2.text)
        aHeader(1) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    End If
    
    cString = cString & turnFound(cWhere, " where ") & cWhere
    cString = cString & " Order by FILE3_11RP.date,file3_12.[order],FILE3_11RP.doc_id,pay"
Else
    If Trim(xCode.text) <> "" Then
        cWhere = cWhere & turn(cWhere, " and ") & " FILE3_11.code = " & MyParn(xCode.text)
    End If
              
    If IsDate(xdate1.text) Then
        cWhere = cWhere & turn(cWhere, " and ") & " FILE3_11.date < " & DateSq(xdate1.text)
    Else
        cWhere = cWhere & turn(cWhere, " and ") & " 1 = 2"
    End If
    
    cField1 = "(Select Sum(FILE3_11.Sal - FILE3_11.Pay) " & _
              " from FILE3_11 " & _
              turnFound(cWhere, " where ") & _
              cWhere & ") as FirstBalance"
    
    cWhere = ""
    cString = "select FILE3_11.*,FILE3_12.desca As MoveDesca, " & _
              cField1 & _
              cField11 & "," & _
              cField12 & "," & _
              cField13 & "," & _
              cField14 & "," & _
              cField15 & _
              " From FILE3_11 Left join FILE3_12 on FILE3_11.type = FILE3_12.code"
    If Trim(xCode.text) <> "" Then
        cWhere = cWhere & turnFound(cWhere, " and ") & " FILE3_11.code = " & MyParn(xCode.text)
        aHeader(0) = "[" & "··⁄„Ì· : " & xCodeDesca.Caption & "]"
    End If
              
    If IsDate(xdate1.text) Then
        cWhere = cWhere & turnFound(cWhere, " and ") & " FILE3_11.date >= " & DateSq(xdate1.text)
        aHeader(1) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    End If
    
    If IsDate(xDate2.text) Then
        cWhere = cWhere & turnFound(cWhere, " and ") & " FILE3_11.date <= " & DateSq(xDate2.text)
        aHeader(1) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    End If
    
    cString = cString & turnFound(cWhere, " where ") & cWhere
    cString = cString & " Order by FILE3_11.date,file3_12.[order],FILE3_11.doc_id,pay"
End If

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not sourcetable.EOF Then
    If Val(sourcetable!FirstBalance & "") <> 0 Then
        temptable.AddNew
        nBalance = Val(sourcetable!FirstBalance)
        nRow = 1
        temptable!str2 = "—’Ìœ ”«»ﬁ"
        If Val(sourcetable!FirstBalance & "") > 0 Then
            temptable!VAL1 = Val(sourcetable!FirstBalance & "")
        Else
            temptable!val2 = Abs(Val(sourcetable!FirstBalance & ""))
        End If
        temptable!val3 = nBalance
        temptable!Val6 = nRow
        
        temptable!Val10 = Val(sourcetable!SalesTotal & "")
        temptable!val11 = Val(sourcetable!SalesRetTotal & "")
        temptable!val12 = Val(sourcetable!SalesTotal & "") - Val(sourcetable!SalesRetTotal & "")
        temptable!VAL13 = Val(sourcetable!CashTotal & "")
        temptable!VAL14 = Val(sourcetable!ChqTotal & "")
        temptable!val15 = Val(sourcetable!CashTotal & "") + Val(sourcetable!ChqTotal & "")
        temptable!val16 = Val(sourcetable!ChqTotal2 & "")
        
        temptable!STR21 = TurnValue(retHeader(aHeader, 0, 3))
        temptable!STR20 = Firsttitle
        temptable.Update
    End If
End If

Do Until sourcetable.EOF
    temptable.AddNew
    nBalance = nBalance + Val(sourcetable!SAL & "") - Val(sourcetable!PAY & "")
    nRow = nRow + 1
    temptable!Date1 = sourcetable!Date
    temptable!str1 = sourcetable!doc_id
    temptable!str2 = turn(sourcetable!DESCA & "", "", sourcetable!MoveDesca & "", sourcetable!DESCA & "")
    temptable!VAL1 = sourcetable!SAL
    temptable!val2 = sourcetable!PAY
    temptable!val3 = nBalance
    temptable!Val6 = nRow
    
    temptable!Val10 = Val(sourcetable!SalesTotal & "")
    temptable!val11 = Val(sourcetable!SalesRetTotal & "")
    temptable!val12 = Val(sourcetable!SalesTotal & "") - Val(sourcetable!SalesRetTotal & "")
    temptable!VAL13 = Val(sourcetable!CashTotal & "")
    temptable!VAL14 = Val(sourcetable!ChqTotal & "")
    temptable!val15 = Val(sourcetable!CashTotal & "") + Val(sourcetable!ChqTotal & "")
    temptable!val16 = Val(sourcetable!ChqTotal2 & "")
    
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!STR20 = Firsttitle
    temptable.Update
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\client3.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub

Private Sub Form_Load()
openCon con
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""
If xCode.text = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 6)
xCodeDesca.Caption = GetDesca("select desca from FILE3_10 where code = " & MyParn(xCode.text), con) & ""

End Sub
Sub myProc()
ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Unload Search3
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From FILE3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub


