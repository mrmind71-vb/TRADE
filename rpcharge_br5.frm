VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpcharge_br5 
   Caption         =   "≈Ã„«·Ï „Êﬁ› Õ”«»  ÊﬂÌ·"
   ClientHeight    =   2400
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
   ScaleHeight     =   2400
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1590
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   5145
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
      Begin MSDataListLib.DataCombo XACC 
         Height          =   315
         Left            =   360
         TabIndex        =   7
         Top             =   1035
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
         Caption         =   "«·Õ”«» : "
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
         TabIndex        =   8
         Top             =   1035
         Width           =   780
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   675
         Width           =   825
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
      TabIndex        =   2
      Top             =   1755
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
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1755
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
      Left            =   2925
      Top             =   2070
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
Attribute VB_Name = "rpcharge_br5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
    doprint1
End Sub
Private Sub doprint1()
Dim temptable As New ADODB.Recordset
Dim temptable2 As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset, nFBal As Double, cAcc As String
ReDim aHeader(3)
contemp.Execute "DELETE * FROM TEMP"
contemp.Execute "DELETE * FROM TEMP2"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
temptable2.Open "temp2", contemp, adOpenStatic, adLockOptimistic, adCmdTable
If xacc.BoundText = "" Then Exit Sub
cAcc = "6" + xacc.BoundText

If IsDate(xDate1.text) Then
    nFBal = Val(GetDesca("SELECT SUM(VAL1-VAL2) FROM ACC_MOVE WHERE ACC = " & MyParn(cAcc) & " AND DATE < " & DateSq(xDate1.text), con) & "")
End If
cString = "Select Sum(VAL1) AS VAL1 , SUM(VAL2) AS VAL2 , flag_desca , SUBGR , charge2 " & _
          " from acc_move  WHERE VAL1 + VAL2 <> 0 "
If xacc.BoundText <> "" Then
    cString = cString & " AND acc_move.ACC = " & MyParn(cAcc)
    aHeader(2) = "[ ≈Ã„«·Ï Õ”«» " & xacc.text & "]"
End If
If IsDate(xDate1.text) Then
    cString = cString & " AND acc_move.date  >= " & DateSq(xDate1.text)
    aHeader(0) = "[" & BetweenString(Format(xDate1.text, "d-m-yyyy"), Format(xdate2.text, "d-m-yyyy")) & "]"
End If
If IsDate(xdate2.text) Then
    cString = cString & " AND acc_move.date <= " & DateSq(xdate2.text)
     aHeader(0) = "[" & BetweenString(Format(xDate1.text, "d-m-yyyy"), Format(xdate2.text, "d-m-yyyy")) & "]"
End If
cString = cString & " Group By flag_desca, SUBGR, charge2"
          
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable2.AddNew
    temptable2!Val10 = 1
    If IsNull(!charge2) Then
        If Trim(!FLAG_DESCA) = "„»Ì⁄«  „’«—Ì›" Then
            temptable2!str2 = "„” ·“„«  »Ì⁄"
            temptable2!str1 = "„’«—Ì›  ÊﬂÌ·"
        Else
            temptable2!str2 = TurnValue(!FLAG_DESCA)
            temptable2!str1 = TurnValue(!SUBGR)
        End If
    Else
        temptable2!str2 = TurnValue(GetDesca("SELECT FILE8_52.DESCA FROM FILE8_51 INNER JOIN FILE8_52 ON FILE8_51.[MAINGROUP] = FILE8_52.CODE  WHERE FILE8_51.CODE = " & MyParn(!charge2), con))
        temptable2!str1 = "„’«—Ì›  ÊﬂÌ·"
    End If
    temptable2!VAL1 = !VAL1
    temptable2!val2 = !val2
    temptable2!str6 = TurnValue(aHeader(2))
    temptable2!STR7 = TurnValue(aHeader(0))
    If temptable2!str1 <> "„’«—Ì›  ÊﬂÌ·" Then
        temptable2!str10 = "0"
    End If
    temptable2.Update
    sourcetable.MoveNext
Loop
End With


    temptable2.AddNew
    temptable2!Val10 = 0
    temptable2!str2 = "—’Ìœ «Ê·"
    If nFBal > 0 Then
        temptable2!VAL1 = nFBal
        temptable2!val2 = 0
    Else
        temptable2!val2 = nFBal * -1
        temptable2!VAL1 = 0
    End If
    temptable2!str6 = TurnValue(aHeader(2))
    temptable2!STR7 = TurnValue(aHeader(0))
    temptable2!str10 = "0"
    temptable2.Update

contemp.Execute " INSERT INTO TEMP ( VAL10, STR1, STR2, VAL1, VAL2, STR5, STR6, STR7, STR10 ) SELECT TEMP2.VAL10, TEMP2.STR1, TEMP2.STR2, SUM(TEMP2.VAL1), SUM(TEMP2.VAL2), TEMP2.STR5, TEMP2.STR6, TEMP2.STR7, TEMP2.STR10 FROM TEMP2 GROUP BY TEMP2.VAL10, TEMP2.STR1, TEMP2.STR2, TEMP2.STR5, TEMP2.STR6 ,TEMP2.STR7 ,TEMP2.STR10  "

contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\charge5.rpt"
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

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM acc0_10 "
Set xacc.RowSource = DATA2
xacc.BoundColumn = "CODE"
xacc.ListField = "DESCA"

End Sub
Private Function RetCharge(pCharge)
ChargeTable.Find "Code = " & MyParn(pCharge), , adSearchForward, adBookmarkFirst
If Not ChargeTable.EOF Then RetCharge = ChargeTable!DESCA
End Function
Private Sub xComp_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then xCOMP.BoundText = ""
End Sub
Function MYVALID() As Boolean
If (Not IsDate(xDate1.text)) And Trim(xDate1.text) <> "" Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xdate2.text)) And Trim(xdate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
