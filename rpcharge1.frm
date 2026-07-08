VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpCharge1 
   Caption         =   " ﬁ«—Ì— «·‰ﬁœÌ…"
   ClientHeight    =   3495
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
   ScaleHeight     =   3495
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2805
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   5145
      Begin VB.CheckBox XISCASH 
         Caption         =   "‰ﬁœÏ ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2340
         Width           =   1500
      End
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1620
         Width           =   4920
         Begin VB.OptionButton XALL 
            Alignment       =   1  'Right Justify
            Caption         =   "«’Ê· ›ﬁÿ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   2
            Left            =   270
            RightToLeft     =   -1  'True
            TabIndex        =   14
            Top             =   270
            Width           =   1185
         End
         Begin VB.OptionButton XALL 
            Alignment       =   1  'Right Justify
            Caption         =   "„’«—Ì› ›ﬁÿ"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   1
            Left            =   1980
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   270
            Width           =   1185
         End
         Begin VB.OptionButton XALL 
            Alignment       =   1  'Right Justify
            Caption         =   "ﬂ· «·„’«—Ì›"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   3375
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   270
            Width           =   1185
         End
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2325
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2325
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xbox 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   945
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XBRANCH 
         Height          =   315
         Left            =   360
         TabIndex        =   9
         Top             =   1305
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSAdodcLib.Adodc DATA2 
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
      Begin VB.Label Label4 
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
         TabIndex        =   10
         Top             =   1350
         Width           =   510
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
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
      TabIndex        =   3
      Top             =   2880
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2880
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
      Top             =   2205
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
Attribute VB_Name = "rpCharge1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
If publicFlag = 1 Then
    doprint1
Else
    doprint2
End If
End Sub
Private Sub doprint1()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(3)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select Charge,Sum(Value) as SumofValue," & _
          "File8_51.MainGroup,File8_51.DescA,File8_52.Desca as MainGroupDesca " & _
          " FROM acc_charge INNER JOIN File8_51 ON acc_charge.CHARGE = File8_51.CODE LEFT JOIN file8_52 ON File8_51.MAINGROUP = file8_52.CODE LEFT JOIN FILE0_50 ON FILE0_50.CODE = acc_charge.BOX WHERE VALUE <> 0  "
If lIsBranchStore Then cString = cString & " AND FILE0_50.BRANCH = " & MyParn(cBranch)
If xbox.BoundText <> "" Then
    cString = cString & " AND BOX = " & MyParn(xbox.BoundText)
    aHeader(1) = "[ Œ“‰… " & xbox.Text & "]"
End If
If XBRANCH.BoundText <> "" Then
    cString = cString & " AND acc_charge.BRANCH = " & MyParn(XBRANCH.BoundText)
    aHeader(2) = "[ ›—⁄ " & XBRANCH.Text & "]"
End If
If IsDate(xDate1.Text) Then
    cString = cString & " AND acc_charge.date  >= " & DateSq(xDate1.Text)
    aHeader(0) = "[" & BetweenString(Format(xDate1.Text, "d-m-yyyy"), Format(xDate2.Text, "d-m-yyyy")) & "]"
End If
If xall(1).Value Then
    cString = cString & " AND File8_51.ISasset = 0 "
    aHeader(3) = " „’«—Ì› ›ﬁÿ "
End If
If xall(2).Value Then
    cString = cString & " AND File8_51.ISasset = 1 "
    aHeader(3) = " «’Ê· ›ﬁÿ "
End If
If XISCASH.Value <> 0 Then
    cString = cString & " AND BOX IS NOT NULL "
    aHeader(3) = " ‰ﬁœÏ ›ﬁÿ ›ﬁÿ "
End If
If IsDate(xDate2.Text) Then
    cString = cString & " AND acc_charge.date <= " & DateSq(xDate2.Text)
     aHeader(0) = "[" & BetweenString(Format(xDate1.Text, "d-m-yyyy"), Format(xDate2.Text, "d-m-yyyy")) & "]"
End If
    
cString = cString & " Group By Charge,File8_51.MainGroup,File8_51.DescA,FILE8_52.DESCA"
          
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !CHARGE
    temptable!str2 = !DESCA
    temptable!str4 = !MainGroup
    temptable!str5 = !MainGroupDesca
    temptable!VAL1 = !Sumofvalue

    temptable!STR7 = TurnValue(retHeader(aHeader, 0, 2))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\charge1.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub
Private Sub doprint2()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(2)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable


cString = "Select Charge,Sum(Value) as SumofValue," & _
          "File8_61.MainGroup,File8_61.DescA,File8_62.Desca as MainGroupDesca " & _
          " FROM acc_INCOM  INNER JOIN File8_61 ON acc_INCOM.CHARGE = File8_61.CODE LEFT JOIN file8_62 ON File8_61.MAINGROUP = file8_62.CODE LEFT JOIN FILE0_50 ON FILE0_50.CODE = acc_INCOM.BOX WHERE acc_INCOM.DOC_NO IS NOT NULL "

If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & " BOX = " & MyParn(xbox.BoundText)
    aHeader(1) = "[ Œ“‰… " & xbox.Text & "]"
End If
If XBRANCH.BoundText <> "" Then
    cString = cString & turnFound(cString) & " acc_INCOM.BRANCH = " & MyParn(XBRANCH.BoundText)
    aHeader(2) = "[ ›—⁄ " & XBRANCH.Text & "]"
End If
If IsDate(xDate1.Text) Then
    cString = cString & turnFound(cString) & " acc_INCOM.date  >= " & DateSq(xDate1.Text)
    aHeader(0) = "[" & BetweenString(Format(xDate1.Text, "d-m-yyyy"), Format(xDate2.Text, "d-m-yyyy")) & "]"
End If
If XISCASH.Value <> 0 Then
    cString = cString & " AND BOX IS NOT NULL "
    aHeader(3) = " ‰ﬁœÏ ›ﬁÿ ›ﬁÿ "
End If

If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & " acc_INCOM.date <= " & DateSq(xDate2.Text)
     aHeader(0) = "[" & BetweenString(Format(xDate1.Text, "d-m-yyyy"), Format(xDate2.Text, "d-m-yyyy")) & "]"
End If
    
cString = cString & " Group By Charge,File8_61.MainGroup,File8_61.DescA,File8_62.DESCA"
          
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !CHARGE
    temptable!str2 = !DESCA
    temptable!str4 = !MainGroup
    temptable!str5 = !MainGroupDesca
    temptable!VAL1 = !Sumofvalue

    temptable!STR7 = TurnValue(retHeader(aHeader, 0, 2))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans

main.Report1.ReportFileName = App.Path & "\Reports\income1.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub

Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
DATA1.ConnectionString = strCon


If cBranch = "00" Or cBranch > "60" Then
    DATA1.RecordSource = "FILE0_50"
Else
    DATA1.RecordSource = "SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch)
End If
Set xbox.RowSource = DATA1
xbox.BoundColumn = "CODE"
xbox.ListField = "DESCA"

Frame2.Visible = (publicFlag = 1)
If cBranch = "00" Then
    data2.ConnectionString = strCon
    data2.RecordSource = "BRANCH"
    Set XBRANCH.RowSource = data2
    XBRANCH.BoundColumn = "CODE"
    XBRANCH.ListField = "DESCA"
    If cBranch <> "00" Then
       XBRANCH.BoundText = cBranch
       XBRANCH.Enabled = False
    End If
End If
End Sub
Private Function RetCharge(pCharge)
ChargeTable.Find "Code = " & MyParn(pCharge), , adSearchForward, adBookmarkFirst
If Not ChargeTable.EOF Then RetCharge = ChargeTable!DESCA
End Function
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID() As Boolean
If (Not IsDate(xDate1.Text)) And Trim(xDate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xDate2.Text)) And Trim(xDate2.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
