VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form bankStatefrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ﬁ«—Ì— «·»‰Êﬂ"
   ClientHeight    =   1830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
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
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1830
   ScaleWidth      =   5475
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   5325
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   1365
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   900
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xBank 
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   180
         Width           =   4155
         _ExtentX        =   7329
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·»‰ﬂ :"
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Left            =   4380
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Label2 
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
         Left            =   4335
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   990
         Width           =   825
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   4275
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
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
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1350
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
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1350
      Width           =   1230
   End
   Begin MSAdodcLib.Adodc data1 
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
Attribute VB_Name = "bankStatefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim ChargeTable As ADODB.Recordset
Private Sub CmdApply_Click()
doprint1
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
Set ChargeTable = New ADODB.Recordset
ChargeTable.Open "File5_00", con, adOpenStatic, adLockReadOnly, adCmdTable
xDate1.Text = ""
xdate2.Text = ""
data1.ConnectionString = strCon
data1.RecordSource = "FILE5_10"

Set XBANK.RowSource = data1
XBANK.ListField = "Desca"
XBANK.BoundColumn = "code"
End Sub
Function MYVALID() As Boolean
If Not IsDate(xDate1.Text) And (xDate1.Text <> "") Then
    MsgBox "Œÿ√ ›Ï «· «—ÌŒ"
    Exit Function
End If
If Not IsDate(xdate2.Text) And (xdate2.Text <> "") Then
    MsgBox "Œÿ√ ›Ï «· «—ÌŒ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
If Not MYVALID Then Exit Sub
Dim I As Integer, nPrevious As Double
Dim aHeader(1)
Dim sourceTable As ADODB.Recordset
Dim temptable As ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Set sourceTable = New ADODB.Recordset
aHeader(0) = " ﬂ‘› Õ”«» »‰ﬂ " & XBANK.Text
If IsDate(xDate1.Text) Then
    Dim loctable As New ADODB.Recordset
    loctable.Open "select sum([value1]  - [value2] ) as Balance from bankmove where " & _
                  " bank = " & MyParn(XBANK.BoundText) & _
                  " AND TYPE <= 4.5 " & _
                  " and [date] < " & DateSq(xDate1.Text), con, adOpenStatic, adLockReadOnly
    If Not loctable.EOF Then nPrevious = Val(loctable!balance & "")
    If nPrevious <> 0 Then
        temptable.AddNew
        temptable!str1 = "—’Ìœ ”«»ﬁ"
        temptable!val3 = nPrevious
        temptable!val1 = nPrevious
        temptable!str21 = TurnValue(retHeader(aHeader, 0, 1))
        temptable!str22 = TurnValue(retHeader(aHeader, 1, 1))
        temptable!Date1 = DateAdd("d", -1, xDate1.Text)
        temptable!Val10 = 0
        temptable.Update
    End If
End If

cString = "Select * from BankMove Where TYPE <= 4.5 AND  BANK = " & MyParn(XBANK.BoundText)
            
If IsDate(xDate1.Text) Then
    cString = cString & turnFound(cString) & "Date >= " & DateSq(xDate1.Text)
    aHeader(1) = "[" & BetweenString(Format(xDate1.Text, "d-m-yyyy"), Format(xdate2.Text, "d-m-yyyy")) & "]"
End If
          
If IsDate(xdate2.Text) Then
    cString = cString & turnFound(cString) & "Date <= " & DateSq(xdate2.Text)
    aHeader(1) = "[" & BetweenString(Format(xDate1.Text, "d-m-yyyy"), Format(xdate2.Text, "d-m-yyyy")) & "]"
End If
          
cString = cString & " Order by [Date],value1"
If sourceTable.State = adStateOpen Then sourceTable.Close
sourceTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourceTable
    Do Until sourceTable.EOF
        I = I + 1
        temptable.AddNew
        temptable!str1 = !doc_no
        temptable!str2 = !TypeDesca
        temptable!str3 = !DESCA
        temptable!Date1 = !Date
        temptable!val1 = !value1
        temptable!val2 = !Value2
        temptable!val3 = nPrevious + Val(!value1 & "") - Val(!Value2 & "")
        nPrevious = nPrevious + Val(!value1 & "") - Val(!Value2 & "")
        temptable!str21 = TurnValue(retHeader(aHeader, 0, 1))
        temptable!str22 = TurnValue(retHeader(aHeader, 1, 1))
        temptable!Val10 = I
        temptable.Update
        sourceTable.MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.REPORT1.ReportFileName = App.Path & "\Reports\BANK1.rpt"
    main.REPORT1.DataFiles(0) = "c:\tempmrshd\Temp.MDB"
    main.REPORT1.Action = 1
End If
If temptable.State = adStateOpen Then temptable.Close
If sourceTable.State = adStateOpen Then sourceTable.Close
Set temptable = Nothing
Set sourceTable = Nothing
End Sub
Private Sub doprint2()
If Not MYVALID Then Exit Sub
Dim I As Integer, nPrevious As Double

Dim sourceTable As ADODB.Recordset
Dim temptable As ADODB.Recordset

Tempdb.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", Tempdb, adOpenStatic, adLockOptimistic, adCmdTable
Set sourceTable = New ADODB.Recordset
If IsDate(xDate1.Text) Then
    cString = "Select Sum([IN])AS sumOfIN,Sum([OUT]) AS SUMOFOUT FROM FILE5_11 " & _
              " WHERE ( FILE5_11.Date < " & DateSq(xDate1.Text) & " or FILE5_11.TYPE = '1' ) " & _
              " AND BANK = " & MyParn(XBANK.BoundText)
    
    sourceTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    If Not (sourceTable.EOF And sourceTable.BOF) Then
    nPrevious = TurnValue(sourceTable!SUMOFIN, Null, 0) - TurnValue(sourceTable!SUMOFOUT, Null, 0)
    temptable.AddNew
    temptable!str1 = "—’Ìœ ”«»ﬁ"
    'temptable!xdate1 = sourcetable!Date
    temptable!Val6 = nPrevious
    temptable!str7 = " ﬂ‘› Õ”«» «·»‰ﬂ " & XBANK.Text & " „‰  «—ÌŒ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xdate2.Text
    temptable!xDate1 = Format(xDate1.Text, "dd/mm/yyyy")
    temptable!str19 = Firsttitle
    temptable!Val10 = I
  '  temptable!str20 = SecondTitle
    temptable.Update
    End If
End If

cString = "Select * from File5_11 Where BANK = " & MyParn(XBANK.BoundText)

If IsDate(xDate1.Text) Then
    cString = cString & turnFound(cString) & "Date >= " & DateSq(xDate1.Text)
End If
          
If IsDate(xdate2.Text) Then
    cString = cString & turnFound(cString) & "Date <= " & DateSq(xdate2.Text)
End If
          
cString = cString & " Order by [Date],[in]"
If sourceTable.State = adStateOpen Then sourceTable.Close
sourceTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourceTable
    Do Until sourceTable.EOF
        If !Type <> "1" Then
            I = I + 1
            temptable.AddNew
    '        temptable!str1 = !DOC_NO
    '        temptable!str2 = TurnValue(RetFind(ChargeTable, "code", "Desca", !CODE), "", Null)
    '        temptable!str3 = TurnValue(!desca, "", Null)
            temptable!xDate1 = Format(!Date, "dd-mm-yyyy")
            If TurnValue(!In, Null, 0) > 0 And !Type = "2" Then
                temptable!val1 = !In
                temptable!Val6 = nPrevious + TurnValue(!In, Null, 0)
                nPrevious = TurnValue(temptable!Val6, Null, 0)
            End If
            
            If TurnValue(!In, Null, 0) > 0 And !Type = "3" Then
                temptable!val2 = !In
                temptable!Val6 = nPrevious + TurnValue(!In, Null, 0)
                nPrevious = TurnValue(temptable!Val6, Null, 0)
            End If
            
            If TurnValue(!out, Null, 0) > 0 And !Type = "2" And IsNull(!code) Then
                temptable!val3 = !out
                temptable!Val6 = nPrevious - TurnValue(!out, Null, 0)
                    nPrevious = TurnValue(temptable!Val6, Null, 0)
            End If
            
            If TurnValue(!out, Null, 0) > 0 And !Type = "2" And Not IsNull(!code) Then
                temptable!VAL4 = !out
                temptable!Val6 = nPrevious - TurnValue(!out, Null, 0)
                nPrevious = TurnValue(temptable!Val6, Null, 0)
            End If
            
            If TurnValue(!out, Null, 0) > 0 And !Type = "4" Then
                temptable!val5 = !out
                temptable!Val6 = nPrevious - TurnValue(!out, Null, 0)
                nPrevious = TurnValue(temptable!Val6, Null, 0)
            End If
                    
            temptable!str7 = " ﬂ‘› Õ”«» «·»‰ﬂ " & XBANK.Text & " „‰  «—ÌŒ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xdate2.Text
            temptable!str19 = Firsttitle
       '     temptable!str20 = SecondTitle
            temptable!Val10 = I
            temptable.Update
        End If
        sourceTable.MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
    GoTo lastsub
End If
Tempdb.BeginTrans
Tempdb.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Report\BANK3.rpt"
main.REPORT1.DataFiles(0) = "c:\elmorshed\Temp.MDB"
main.REPORT1.Action = 1
lastsub:
If temptable.State = adStateOpen Then temptable.Close
If sourceTable.State = adStateOpen Then sourceTable.Close
Set temptable = Nothing
Set sourceTable = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
