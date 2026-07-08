VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpcharge_BR2 
   Caption         =   " ﬁ«—Ì— «·‰ﬁœÌ…"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5265
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
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3585
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2670
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   -45
      Width           =   5145
      Begin VB.TextBox xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   315
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1665
         Width           =   3345
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xbox 
         Height          =   315
         Left            =   315
         TabIndex        =   2
         Top             =   945
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xCharge 
         Height          =   315
         Left            =   315
         TabIndex        =   3
         Top             =   1305
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XBRANCH 
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   2070
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSAdodcLib.Adodc DATA3 
         Height          =   330
         Left            =   0
         Top             =   540
         Visible         =   0   'False
         Width           =   2340
         _ExtentX        =   4128
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
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2115
         Width           =   510
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "»Ì«‰ :"
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
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1710
         Width           =   405
      End
      Begin VB.Label Label4 
         Caption         =   "«·„’—Ê› :"
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1395
         Width           =   1005
      End
      Begin VB.Label Label1 
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·Ï  «—ÌŒ :"
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   675
         Width           =   825
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Œ“‰… :"
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
         Left            =   3780
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1035
         Width           =   465
      End
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
      Height          =   420
      Left            =   1395
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3105
      Width           =   1320
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
      Height          =   420
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3105
      Width           =   1320
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3555
      Top             =   2295
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   3330
      Top             =   1800
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
Attribute VB_Name = "rpcharge_BR2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
    doprint1
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub doprint1()
Dim sourcetable As New ADODB.Recordset
Dim temptable As New ADODB.Recordset
Dim aHeader(4)
If Not MYVALID Then Exit Sub

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select acc_charge_BR.Date,acc_charge_BR.DescA,file0_50.desca AS FILE0_50DESCA,FILE8_51.DESCA AS CHARGEDESCA,Value,acc_charge_BR.DOC_NO,BOX , branch.desca  as branch_desca " & _
          " FROM   branch INNER JOIN acc_charge_BR ON branch.code = acc_charge_BR.BRANCH LEFT OUTER JOIN FILE8_51 ON acc_charge_BR.CHARGE = FILE8_51.CODE LEFT OUTER JOIN FILE0_50 ON acc_charge_BR.BOX = FILE0_50.CODE WHERE VALUE <> 0  "

If lIsBranchStore Then cString = cString & " AND FILE0_50.BRANCH = " & MyParn(cBranch)
If xBox.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_charge_BR.BOX = " & MyParn(xBox.BoundText)
     aHeader(1) = "[ Œ“‰… " & xBox.text & "]"
End If
If XBRANCH.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_charge_BR.branch = " & MyParn(XBRANCH.BoundText)
     aHeader(1) = "[ ›—⁄ " & XBRANCH.text & "]"
End If

If IsDate(xDate1.text) Then
    cString = cString & turnFound(cString) & " acc_charge_BR.date >= " & DateSq(xDate1.text)
    aHeader(0) = "[" & BetweenString(xDate1.text, XDATE2.text) & "]"
End If


If IsDate(XDATE2.text) Then
    cString = cString & turnFound(cString) & " acc_charge_BR.date <= " & DateSq(XDATE2.text)
    aHeader(0) = "[" & BetweenString(xDate1.text, XDATE2.text) & "]"
End If
    
    

If xCharge.BoundText <> "" Then
    cString = cString & turnFound(cString) & " acc_charge_BR.charge = " & MyParn(xCharge.BoundText)
End If

If Trim(XDESCA.text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(XDESCA.text, "acc_charge_BR.desca")
    aHeader(2) = "[" & "«·»Ì«‰ : " & XDESCA.text & "]"
End If


sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    If Len(!doc_no) > 10 Then
        temptable!str1 = RetZero(DelZero(Mid(!doc_no, 2)), 6)
    Else
        temptable!str1 = !doc_no
    End If
    temptable!str2 = TurnValue(!FILE0_50DESCA)
    'temptable!Str3 = TurnValue(ArbString(!desca & turn(!chargeDesca & "", turn(!desca, "-") & !chargeDesca) & ""))
    temptable!str3 = !DESCA
    If Not IsNull(!chargeDesca) Then
        temptable!str3 = TurnValue(ArbString(temptable!str3 & turn(temptable!str3 & "", "-") & !chargeDesca & ""))
    End If
    temptable!VAL1 = !Value
    temptable!Date1 = !Date
    temptable!STR21 = "  ›’Ì·Ï „’—Ê› " & xCharge.text
    temptable!str22 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str6 = !branch_desca
    temptable.Update
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
Set temptable = Nothing

contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\charge2.rpt"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
End Sub
Private Sub doprint2()
Dim sourcetable As New ADODB.Recordset
Dim temptable As New ADODB.Recordset
Dim aHeader(3)
If Not MYVALID Then Exit Sub

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select acc_INCOM.Date,acc_INCOM.DescA,file0_50.desca AS FILE0_50DESCA,file8_61.DESCA AS CHARGEDESCA,Value,acc_INCOM.DOC_NO,BOX" & _
          " From acc_INCOM  Left join file8_61 on acc_INCOM.charge = file8_61.code LEFT JOIN FILE0_50 ON acc_INCOM.BOX = FILE0_50.CODE WHERE acc_INCOM.DOC_NO IS NOT NULL "

If lIsBranchStore Then cString = cString & " AND FILE0_50.BRANCH = " & MyParn(cBranch)
If xBox.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_INCOM.BOX = " & MyParn(xBox.BoundText)
     aHeader(1) = "[ Œ“‰… " & xBox.text & "]"
End If
If XBRANCH.BoundText <> "" Then
    cString = cString & turnFound(cString) & "  acc_INCOM.branch = " & MyParn(XBRANCH.BoundText)
     aHeader(1) = "[ ›—⁄ " & XBRANCH.text & "]"
End If

If IsDate(xDate1.text) Then
    cString = cString & turnFound(cString) & " acc_INCOM.date >= " & DateSq(xDate1.text)
    aHeader(0) = "[" & BetweenString(xDate1.text, XDATE2.text) & "]"
End If


If IsDate(XDATE2.text) Then
    cString = cString & turnFound(cString) & " acc_INCOM.date <= " & DateSq(XDATE2.text)
    aHeader(0) = "[" & BetweenString(xDate1.text, XDATE2.text) & "]"
End If
    

If xCharge.BoundText <> "" Then
    cString = cString & turnFound(cString) & " acc_INCOM.charge = " & MyParn(xCharge.BoundText)
End If

If Trim(XDESCA.text) <> "" Then
    cString = cString & turnFound(cString) & MyParnAnd(XDESCA.text, "acc_INCOM.desca")
    aHeader(2) = "[" & "«·»Ì«‰ : " & XDESCA.text & "]"
End If


sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !doc_no
    temptable!str2 = TurnValue(!FILE0_50DESCA)
    temptable!str3 = TurnValue(!DESCA, "", Null)
    If Not IsNull(!chargeDesca) Then
        temptable!str3 = TurnValue(ArbString(temptable!str3 & turn(temptable!str3 & "", "-") & !chargeDesca & ""))
    End If
    temptable!VAL1 = !Value
    temptable!Date1 = !Date
    temptable!STR21 = "  ›’Ì·Ï „’—Ê› " & xCharge.text
    temptable!str22 = TurnValue(retHeader(aHeader, 0, 3))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
Set temptable = Nothing

contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\income2.rpt"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
Set sourcetable = New ADODB.Recordset
Set temptable = New ADODB.Recordset
DATA1.ConnectionString = strCon

DATA1.RecordSource = "Select * FROM " & IIf(publicFlag = 2, "file8_51", "file8_61") & " ORDER BY DESCA"
Set xCharge.RowSource = DATA1
xCharge.BoundColumn = "Code"
xCharge.ListField = "DescA"

DATA2.ConnectionString = strCon
If cBranch = "00" Then
    DATA2.RecordSource = "FILE0_50"
Else
    DATA2.RecordSource = "SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch)
End If
Set xBox.RowSource = DATA2
xBox.BoundColumn = "CODE"
xBox.ListField = "DESCA"


DATA3.ConnectionString = strCon
DATA3.RecordSource = "select * from BRANCH where isbranch2 = 1 "
Set XBRANCH.RowSource = DATA3
XBRANCH.BoundColumn = "CODE"
XBRANCH.ListField = "DESCA"
If cBranch <> "00" Then
   XBRANCH.BoundText = cBranch
   XBRANCH.Enabled = False
End If

End Sub
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID() As Boolean
'If xCharge.BoundText = "" Then
'    MsgBox IIf(publicFlag = 2, "»Ì«‰ «·„’—Ê› „ÿ·Ê»", "»Ì«‰ «·«Ì—«œ „ÿ·Ê»")
'    Exit Function
'End If
If (Not IsDate(xDate1.text)) And Trim(xDate1.text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(XDATE2.text)) And Trim(XDATE2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xbox_Validate(Cancel As Boolean)
If Not xBox.MatchedWithList Then xBox.BoundText = ""
End Sub

Private Sub xCharge_Validate(Cancel As Boolean)
If Not xCharge.MatchedWithList Then xCharge.BoundText = ""
End Sub
