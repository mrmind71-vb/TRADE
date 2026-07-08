VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpCharge3 
   Caption         =   " ﬁ«—Ì— «·‰ﬁœÌ…"
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6390
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
   ScaleHeight     =   4020
   ScaleWidth      =   6390
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   5775
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
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2790
         Width           =   1500
      End
      Begin VB.Frame Frame2 
         Height          =   600
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2070
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
            TabIndex        =   16
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
            TabIndex        =   15
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
            TabIndex        =   14
            Top             =   270
            Width           =   1185
         End
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2280
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
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   315
         TabIndex        =   9
         Top             =   1305
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XBRANCH 
         Height          =   315
         Left            =   315
         TabIndex        =   11
         Top             =   1665
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
         Left            =   -45
         Top             =   135
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1710
         Width           =   510
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "»‰œ —∆Ì”ÌÏ "
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1350
         Width           =   930
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
      Top             =   3465
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
      Top             =   3465
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
      Top             =   2160
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
Attribute VB_Name = "rpCharge3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
If publicFlag = 3 Then
    doprint1
Else
    doprint2
End If
End Sub
Private Sub doprint1()
ReDim aHeader(4)
If Not MYVALID Then Exit Sub
Dim sourcetable As ADODB.Recordset
Dim temptable As ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT ACC_CHARGE.DATE, acc_charge.DESCA as DescCharg, acc_charge.VALUE, acc_charge.CHARGE, acc_charge.DOC_NO ,acc_charge.BOX," & _
          "FILE8_51.DESCA AS DescCode,file0_50.DESCA AS BOXDESCA " & _
          " FROM acc_charge  LEFT JOIN FILE8_51 ON acc_charge.CHARGE = FILE8_51.CODE LEFT JOIN FILE0_50 ON acc_charge.BOX = FILE0_50.CODE WHERE VALUE <> 0  "
If lIsBranchStore Then cString = cString & " AND FILE0_50.BRANCH = " & MyParn(cBranch)
If xGroup.BoundText <> "" Then
    cString = cString & turnFound(cString) & " FILE8_51.MAINGROUP = " & MyParn(xGroup.BoundText)
    aHeader(2) = "[" & xGroup.Text & "]"
End If
If IsDate(xDate1.Text) Then
    cString = cString & turnFound(cString) & "acc_charge.Date >= " & DateSq(xDate1.Text)
    aHeader(0) = "[" & BetweenString(xDate1.Text, xdate2.Text) & "]"
End If
If IsDate(xdate2.Text) Then
    cString = cString & turnFound(cString) & "acc_charge.Date <= " & DateSq(xdate2.Text)
    aHeader(0) = "[" & BetweenString(xDate1.Text, xdate2.Text) & "]"
End If

If xall(1).Value Then
    cString = cString & turnFound(cString) & " File8_51.ISasset = 0 "
    aHeader(4) = " „’«—Ì› ›ﬁÿ "
End If
If xall(2).Value Then
    cString = cString & turnFound(cString) & " File8_51.ISasset = 1 "
    aHeader(4) = " «’Ê· ›ﬁÿ "
End If
If XISCASH.Value <> 0 Then
    cString = cString & " AND BOX IS NOT NULL "
    aHeader(3) = " ‰ﬁœÏ ›ﬁÿ ›ﬁÿ "
End If

If xbox.Text <> "" Then
    cString = cString & turnFound(cString) & "acc_charge.BOX = " & MyParn(xbox.BoundText)
    aHeader(1) = "[Œ“‰… " & xbox.Text & "]"
End If
If XBRANCH.BoundText <> "" Then
    cString = cString & turnFound(cString) & " acc_charge.BRANCH = " & MyParn(XBRANCH.BoundText)
    aHeader(3) = "[ ›—⁄ " & XBRANCH.Text & "]"
End If

Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!date3 = !Date
    temptable!str1 = !doc_no
    temptable!str2 = !BOXDESCA
    temptable!str3 = !DescCode
    temptable!str4 = TurnValue(!DescCharg, "", Null)
      
    temptable!VAL1 = !Value
    temptable!str8 = TurnValue(retHeader(aHeader, 0, 3))
    temptable.Update
    sourcetable.MoveNext
Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.Report1.ReportFileName = App.Path & "\Reports\charge3.rpt"
    main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    main.Report1.Action = 1
End If
If temptable.State = adStateOpen Then temptable.Close
If sourcetable.State = adStateOpen Then sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub doprint2()
ReDim aHeader(3)
If Not MYVALID Then Exit Sub
Dim sourcetable As ADODB.Recordset
Dim temptable As ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT ACC_INCOM.DATE, ACC_INCOM.DESCA as DescCharg, ACC_INCOM.VALUE, ACC_INCOM.CHARGE, ACC_INCOM.DOC_NO ,ACC_INCOM.BOX," & _
          "file8_61.DESCA AS DescCode,file0_50.DESCA AS BOXDESCA " & _
          " FROM ACC_INCOM  LEFT JOIN file8_61 ON ACC_INCOM.CHARGE = file8_61.CODE LEFT JOIN FILE0_50 ON ACC_INCOM.BOX = FILE0_50.CODE WHERE ACC_INCOM.DOC_NO IS NOT NULL "

If lIsBranchStore Then cString = cString & " AND FILE0_50.BRANCH = " & MyParn(cBranch)
If xGroup.BoundText <> "" Then
    cString = cString & turnFound(cString) & " file8_61.MAINGROUP = " & MyParn(xGroup.BoundText)
    aHeader(2) = "[" & xGroup.Text & "]"
End If
If IsDate(xDate1.Text) Then
    cString = cString & turnFound(cString) & "ACC_INCOM.Date >= " & DateSq(xDate1.Text)
    aHeader(0) = "[" & BetweenString(xDate1.Text, xdate2.Text) & "]"
End If
If IsDate(xdate2.Text) Then
    cString = cString & turnFound(cString) & "ACC_INCOM.Date <= " & DateSq(xdate2.Text)
    aHeader(0) = "[" & BetweenString(xDate1.Text, xdate2.Text) & "]"
End If

If xbox.Text <> "" Then
    cString = cString & turnFound(cString) & "ACC_INCOM.BOX = " & MyParn(xbox.BoundText)
    aHeader(1) = "[Œ“‰… " & xbox.Text & "]"
End If
If XBRANCH.BoundText <> "" Then
    cString = cString & turnFound(cString) & " ACC_INCOM.BRANCH = " & MyParn(XBRANCH.BoundText)
    aHeader(3) = "[ ›—⁄ " & XBRANCH.Text & "]"
End If
If XISCASH.Value <> 0 Then
    cString = cString & " AND BOX IS NOT NULL "
    aHeader(3) = " ‰ﬁœÏ ›ﬁÿ ›ﬁÿ "
End If

Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!date3 = !Date
    temptable!str1 = !doc_no
    temptable!str2 = !BOXDESCA
    temptable!str3 = !DescCode
    temptable!str4 = TurnValue(!DescCharg, "", Null)
      
    temptable!VAL1 = !Value
    temptable!str8 = TurnValue(retHeader(aHeader, 0, 3))
    temptable.Update
    sourcetable.MoveNext
Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.Report1.ReportFileName = App.Path & "\Reports\income3.rpt"
    main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    main.Report1.Action = 1
End If
If temptable.State = adStateOpen Then temptable.Close
If sourcetable.State = adStateOpen Then sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
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

data1.ConnectionString = strCon
If cBranch = "00" Or cBranch > "60" Then
    data1.RecordSource = "FILE0_50"
Else
    data1.RecordSource = "SELECT * FROM FILE0_50 WHERE BRANCH = " & MyParn(cBranch)
End If
Set xbox.RowSource = data1
xbox.BoundColumn = "CODE"
xbox.ListField = "DESCA"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select * FROM " & IIf(publicFlag = 2, "file8_52", "file8_62") & " ORDER BY DESCA"
Set xGroup.RowSource = DATA2
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

If cBranch = "00" Then
data3.ConnectionString = strCon
data3.RecordSource = "BRANCH"
Set XBRANCH.RowSource = data3
XBRANCH.BoundColumn = "CODE"
XBRANCH.ListField = "DESCA"

If cBranch <> "00" Then
   XBRANCH.BoundText = cBranch
   XBRANCH.Enabled = False
End If
End If
Frame2.Visible = (publicFlag = 3)
End Sub
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID() As Boolean
If (Not IsDate(xDate1.Text)) And Trim(xDate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xdate2.Text)) And Trim(xdate2.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
