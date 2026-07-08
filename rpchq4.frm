VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpChq4 
   Caption         =   " ﬁ«—Ì— «·‘Ìﬂ« "
   ClientHeight    =   1935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   1935
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   8
      Top             =   1440
      Width           =   1185
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
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   5280
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1680
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   900
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   675
         TabIndex        =   3
         Top             =   180
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
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
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   645
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "Õ Ï :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1020
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄…  :"
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
         Left            =   4080
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   765
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   3330
      Top             =   1845
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
      Left            =   2655
      Top             =   1485
      Visible         =   0   'False
      Width           =   1740
      _ExtentX        =   3069
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
Attribute VB_Name = "rpChq4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sourcetable As ADODB.Recordset
Dim temptable As ADODB.Recordset
Dim grouptable As ADODB.Recordset
Dim ClientTable As ADODB.Recordset
Function MYVALID()
If Not IsDate(xDate1.Text) Then Exit Function
If Not IsDate(xdate2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xGroup1.BoundText = ""
xgroup2.BoundText = ""
xDate1.Text = ""
xdate2.Text = ""
End Sub
Private Sub Form_Load()
Set sourcetable = New ADODB.Recordset
Set temptable = New ADODB.Recordset
Set grouptable = New ADODB.Recordset
Set ClientTable = New ADODB.Recordset

ClientTable.Open IIf(lCust, "file3_10", "file4_10"), CON, adOpenStatic, adLockReadOnly, adCmdTable
grouptable.Open "select * from " & IIf(lCust, "file3_20", "file4_20"), CON, adOpenStatic, adLockReadOnly, adCmdText

Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
data1.ConnectionString = CON.ConnectionString

If lCust Then
    data1.RecordSource = "SELECT * FROM FILE3_20"
Else
    data1.RecordSource = "SELECT * FROM FILE4_20"
End If

xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"
data1.Refresh
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
contemp.Execute "Delete * From Temp"
If temptable.State = adStateOpen Then temptable.Close
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
cHead = " ≈Ã„«·Ï „Êﬁ› «·‘Ìﬂ«  "
If xGroup.BoundText <> "" Then
    ClientTable.Filter = "[GROUP] = " & MyParn(xGroup.BoundText)
End If
         
With ClientTable
If ClientTable.EOF And ClientTable.BOF Then
    Exit Sub
    MsgBox "·«  ÊÃœ ”Ã·«  ·⁄—÷Â«"
End If
Do Until .EOF
    nvalue1 = CalcChq(1, !CODE)
    nvalue2 = CalcChq(2, !CODE)
    nvalue3 = CalcChq(3, !CODE)
    nvalue4 = CalcChq(4, !CODE)
    nvalue5 = CalcChq(5, !CODE)
    nvalue6 = CalcChq(6, !CODE)
    nvalue7 = CalcChq(7, !CODE)
    If nvalue1 + nvalue2 + nvalue3 + nvalue4 + nvalue5 + nvalue6 + nvalue7 + nvalue7 > 0 Then
        temptable.AddNew
        temptable!str1 = !Desca
        temptable!str3 = cHead
        temptable!str4 = " „‰  «—ÌŒ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xdate2.Text
        temptable!Str5 = !Group
        temptable!str6 = RetField(grouptable, !Group, "code", "desca")
        temptable!val1 = nvalue1
        temptable!val2 = nvalue2
        temptable!val3 = nvalue3
        temptable!Val4 = nvalue4
        temptable!Val5 = nvalue5
        temptable!Val6 = nvalue6
        temptable!Val7 = nvalue7
        temptable!str19 = firsttitle
       ' temptable!str20 = SecondTitle
        temptable.Update
    End If
    .MoveNext
Loop
End With

contemp.BeginTrans
contemp.CommitTrans
Report1.ReportFileName = App.Path & "\Reports\chq5.rpt"
Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Report1.Action = 1
End Sub
Function CalcChq(cPass, cCode) As Double
If lCust Then
    Select Case cPass
        Case 1
            cString = "select file5_20.* , file3_10.GROUP" & _
                      " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '0' and date_1 < DateValue(" & MyParn(xDate1.Text) & ")"
        Case 2
            cString = "select file5_20.* , file3_10.GROUP" & _
                      " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '2' and date_1 < DateValue(" & MyParn(xDate1.Text) & ") and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") and date_3 <= DateValue(" & MyParn(xdate2.Text) & ")"
        Case 3
            cString = "select file5_20.* , file3_10.GROUP" & _
                      " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE date_1 >= DateValue(" & MyParn(xDate1.Text) & ") And date_1 <= DateValue(" & MyParn(xdate2.Text) & ")"
        Case 4
            cString = "select file5_20.* , file3_10.GROUP" & _
                      " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '2' and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") And date_3 <= DateValue(" & MyParn(xdate2.Text) & ") And date_1 >= DateValue(" & MyParn(xDate1.Text) & ") and date_1 <= DateValue(" & MyParn(xdate2.Text) & ")"
        Case 5
            cString = "select file5_20.* , file3_10.GROUP" & _
                      " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '1' and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") And date_3 <= DateValue(" & MyParn(xdate2.Text) & ")"

        Case 6
            cString = "select file5_20.* , file3_10.GROUP" & _
                      " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '0' and date_1 <= DateValue(" & MyParn(xdate2.Text) & ")"
        Case 7
            cString = "select file5_20.* , file3_10.GROUP" & _
                      " from file5_20 inner join file3_10 ON FILE3_10.CODE = FILE5_20.CODE WHERE CLOSED = '0' "
    End Select
    cString = cString & " and File3_10.code = " & MyParn(cCode)
Else
    Select Case cPass
        Case 1
            cString = "select file5_21.* , file4_10.Group AS Cl_Group " & _
                      " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE1 WHERE CLOSED = '0' and date_1 < DateValue(" & MyParn(xDate1.Text) & ")" & _
                      " and File4_10.CODE = " & MyParn(cCode)
  
        Case 2
            cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                      " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE1 WHERE CLOSED = '2' and date_1 < DateValue(" & MyParn(xDate1.Text) & ") and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") and date_3 <= DateValue(" & MyParn(xdate2.Text) & ")" & _
                      " and File4_10.CODE = " & MyParn(cCode)
        Case 3
            cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                      " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE1 WHERE date_1 >= DateValue(" & MyParn(xDate1.Text) & ") And date_1 <= DateValue(" & MyParn(xdate2.Text) & ")" & _
                      " and File4_10.CODE = " & MyParn(cCode)
        Case 4
            cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                      " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE1 WHERE CLOSED = '2' and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") And date_3 <= DateValue(" & MyParn(xdate2.Text) & ") And date_1 >= DateValue(" & MyParn(xDate1.Text) & ") and date_1 <= DateValue(" & MyParn(xdate2.Text) & ")" & _
                     " and File4_10.CODE = " & MyParn(cCode)
        Case 5
            cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                      " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE1 WHERE CLOSED = '1' and date_3 >= DateValue(" & MyParn(xDate1.Text) & ") And date_3 <= DateValue(" & MyParn(xdate2.Text) & ")" & _
                      " and File4_10.CODE = " & MyParn(cCode)
        Case 6
            cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                      " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE1 WHERE CLOSED = '0' and date_1 <= DateValue(" & MyParn(xdate2.Text) & ")" & _
                      " and File4_10.CODE = " & MyParn(cCode)
        Case 7
            cString = "select file5_21.* , file4_10.GROUP AS Cl_Group " & _
                      " from file5_21 inner join file4_10 ON FILE4_10.CODE = FILE5_21.CODE1 WHERE CLOSED = '0' " & _
                      " and File4_10.CODE = " & MyParn(cCode)
    End Select
End If
If sourcetable.State = adStateOpen Then sourcetable.Close
sourcetable.Open cString, CON, adOpenStatic, adLockReadOnly, adCmdText
If Not (sourcetable.EOF And sourcetable.BOF) Then
    Do Until sourcetable.EOF
        NRETURN = NRETURN + TurnValue(sourcetable![Value], Null, 0)
        sourcetable.MoveNext
    Loop
End If
CalcChq = NRETURN
End Function
