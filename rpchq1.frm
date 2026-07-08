VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpChq1 
   Caption         =   " ﬁ«—Ì— «Ê—«ﬁ «·ﬁ»÷"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6420
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
   ScaleHeight     =   4920
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox xNoBank 
      Alignment       =   1  'Right Justify
      Caption         =   "€Ì— „Êœ⁄… ›Ï »‰ﬂ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3465
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   4455
      Width           =   2805
   End
   Begin VB.Frame Frame4 
      Height          =   1050
      Left            =   765
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   0
      Width           =   5550
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   540
         Width           =   1725
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   180
         Width           =   1725
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
         Height          =   195
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   675
         Width           =   450
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "„‰ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4725
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   225
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1035
      Width           =   6180
      Begin VB.TextBox XCHARGE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„’—Ê› "
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
         Left            =   4890
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   660
      End
      Begin VB.Label xCodeDesca1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   3345
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2385
      Width           =   6180
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo xGroup2 
         Height          =   315
         Left            =   1290
         TabIndex        =   8
         Top             =   540
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Ã„Ê⁄… :"
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
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label xCodedesca2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   180
         Width           =   3345
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Left            =   4980
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   570
      End
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
      Height          =   465
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   4410
      Width           =   1095
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
      Height          =   465
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   4410
      Width           =   1140
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   2520
      Top             =   4455
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
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   3375
      Width           =   6180
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   1350
         TabIndex        =   13
         Top             =   540
         Width           =   3480
         _ExtentX        =   6138
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xBank 
         Height          =   315
         Left            =   1305
         TabIndex        =   16
         Top             =   180
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
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
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·Œ“‰… :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   585
      End
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
   Begin MSAdodcLib.Adodc data3 
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
   Begin MSAdodcLib.Adodc data4 
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
Attribute VB_Name = "rpChq1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Function MYVALID()
If Not IsDate(xdate1.text) Then Exit Function
If Not IsDate(xDate2.text) Then Exit Function
MYVALID = True
End Function
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xdate1.text = ""
xDate2.text = ""
End Sub
Private Sub CmdApply_Click()
If publicFlag = 1 Then doprint1
If publicFlag = 2 Then doprint2
If publicFlag = 3 Then doprint3
If publicFlag = 4 Then doprint4
If publicFlag = 8 Then doprint8
End Sub
Private Sub doprint1()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim aHeader(7)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "select file5_20.*,FILE3_10.DESCA as Desca1,File4_10.Desca as Desca2,file5_10.desca as bankDesca " & _
        " FROM ((file5_20 LEFT JOIN FILE3_10 ON FILE5_20.CODE1 = FILE3_10.CODE)inner JOIN file4_10 on file5_20.code2 = file4_10.code) left join file5_10 on file5_20.id_bank = file5_10.code " & _
        " WHERE CLOSED = '0' "

If IsDate(xdate1.text) Then
    aHeader(0) = BetweenString(xdate1.text, xDate2.text)
    cString = cString & turnFound(cString) & "date_1 >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    aHeader(0) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    cString = cString & turnFound(cString) & "date_1 <= " & DateSq(xDate2.text)
End If

If xCharge.text <> "" Then
    cString = cString & turnFound(cString) & " CHARGE = " & MyParn(xCharge.text)
    aHeader(1) = "[" & "„’—Ê› :" & xCodedesca1.Caption & "]"
End If


If xCode2.text <> "" Then
    cString = cString & turnFound(cString) & " code2 = " & MyParn(xCode2.text)
    aHeader(3) = "[" & "··„Ê—œ :" & xCodeDesca2.Caption & "]"
End If


If xBank.BoundText <> "" Then
    cString = cString & turnFound(cString) & " ID_BANK = " & MyParn(xBank.BoundText)
    aHeader(5) = "[" & "«·»‰ﬂ : " & xBank.text & "]"
End If

If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & " Box = " & MyParn(xbox.BoundText)
    aHeader(6) = "[" & "«·Œ“‰… :" & xbox.text & "]"
End If

If xNoBank.Value <> 0 Then
    cString = cString & turnFound(cString) & " (ID_BANK is Null)"
    aHeader(7) = "[" & "‘Ìﬂ«  €Ì— „Êœ⁄… ›Ï »‰ﬂ " & "]"
End If


cString = cString & " ORDER BY DATE_1"
    
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !CHK_ID
    If Not IsNull(!desca1) Then
        temptable!str2 = "⁄„Ì·"
        temptable!str3 = !desca1
    ElseIf Not IsNull(!Desca2) Then
         temptable!str2 = "„Ê—œ"
         temptable!str3 = !Desca2
    End If
    If Not IsNull(!DESCA) Then
        temptable!str3 = temptable!str3 & IIf(IsNull(temptable!str3), "", " - ") & !DESCA
    End If
    temptable!str4 = !Bank_rec
'    temptable!str4 = !BankDesca
    temptable!Date1 = !date_1
    'temptable!date2 = !date_R
    temptable!date2 = !DateBank
    temptable!VAL1 = !Value
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 8))
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\CHQ1.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub
Private Sub doprint2()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim aHeader(7)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable


cString = "select file5_20.*,FILE3_10.DESCA as Desca1,File4_10.Desca as Desca2,file5_10.desca as bankDesca " & _
        " FROM ((file5_20 LEFT JOIN FILE3_10 ON FILE5_20.CODE1 = FILE3_10.CODE)inner JOIN file4_10 on file5_20.code2 = file4_10.code) left join file5_10 on file5_20.id_bank = file5_10.code " & _
        " WHERE CLOSED = '2' "

If IsDate(xdate1.text) Then
    aHeader(0) = BetweenString(xdate1.text, xDate2.text)
    cString = cString & turnFound(cString) & " DateValue(date_3) >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    aHeader(0) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    cString = cString & turnFound(cString) & " DateValue(date_3) <= " & DateSq(xDate2.text)
End If

If xCharge.text <> "" Then
    cString = cString & turnFound(cString) & " code1 = " & MyParn(xCharge.text)
    aHeader(2) = "[" & "··⁄„Ì· :" & xCodedesca1.Caption & "]"
End If


If xCode2.text <> "" Then
    cString = cString & turnFound(cString) & " code2 = " & MyParn(xCode2.text)
    aHeader(4) = "[" & "··„Ê—œ :" & xCodeDesca2.Caption & "]"
End If


If xBank.BoundText <> "" Then
    cString = cString & turnFound(cString) & " ID_BANK = " & MyParn(xBank.BoundText)
    aHeader(6) = "[" & "«·»‰ﬂ : " & xBank.text & "]"
End If

If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & " Box = " & MyParn(xbox.BoundText)
    aHeader(7) = "[" & "«·Œ“‰… :" & xbox.text & "]"
End If

cString = cString & " ORDER BY DATE_3"
    
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !CHK_ID
    If Not IsNull(!desca1) Then
        temptable!str2 = "⁄„Ì·"
        temptable!str3 = !desca1
    ElseIf Not IsNull(!Desca2) Then
         temptable!str2 = "„Ê—œ"
         temptable!str3 = !Desca2
    End If
    If Not IsNull(!DESCA) Then
        temptable!str3 = temptable!str3 & IIf(IsNull(temptable!str3), "", " - ") & !DESCA
    End If
    temptable!str4 = !BankDesca
    temptable!Date1 = !date_3
    temptable!date2 = !date_1
    temptable!VAL1 = !Value
    
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 8))
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\CHQ2.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub
Private Sub doprint3()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim aHeader(7)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "select file5_20.*,FILE3_10.DESCA as Desca1,File4_10.Desca as Desca2,File5_10.Desca as BankDesca " & _
        " FROM ((file5_20 LEFT JOIN FILE3_10 ON FILE5_20.CODE1 = FILE3_10.CODE)inner JOIN file4_10 on file5_20.code2 = file4_10.code) left join file5_10 on file5_20.id_bank = file5_10.code " & _
        " WHERE CLOSED = '1' "

If IsDate(xdate1.text) Then
    aHeader(0) = BetweenString(xdate1.text, xDate2.text)
    cString = cString & turnFound(cString) & " date_3 >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    aHeader(0) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    cString = cString & turnFound(cString) & " date_3 <= " & DateSq(xDate2.text)
End If

If xCharge.text <> "" Then
    cString = cString & turnFound(cString) & " code1 = " & MyParn(xCharge.text)
    aHeader(2) = "[" & "··⁄„Ì· :" & xCodedesca1.Caption & "]"
End If


If xCode2.text <> "" Then
    cString = cString & turnFound(cString) & " code2 = " & MyParn(xCode2.text)
    aHeader(4) = "[" & "··„Ê—œ :" & xCodeDesca2.Caption & "]"
End If


If xBank.BoundText <> "" Then
    cString = cString & turnFound(cString) & " ID_BANK = " & MyParn(xBank.BoundText)
    aHeader(6) = "[" & "«·»‰ﬂ : " & xBank.text & "]"
End If

If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & " Box = " & MyParn(xbox.BoundText)
    aHeader(7) = "[" & "«·Œ“‰… :" & xbox.text & "]"
End If

cString = cString & " ORDER BY DATE_3"
    
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !CHK_ID
    If Not IsNull(!desca1) Then
        temptable!str2 = "⁄„Ì·"
        temptable!str3 = !desca1
    ElseIf Not IsNull(!Desca2) Then
         temptable!str2 = "„Ê—œ"
         temptable!str3 = !Desca2
    End If
    If Not IsNull(!DESCA) Then
        temptable!str3 = temptable!str3 & IIf(IsNull(temptable!str3), "", " - ") & !DESCA
    End If
    temptable!Date1 = !date_3
    temptable!date2 = !date_1
    temptable!str4 = !BankDesca
    temptable!VAL1 = !Value
    
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 8))
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\CHQ3.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub
Private Sub doprint4()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim aHeader(7)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
cString = "select file5_20.*,FILE3_10.DESCA as Desca1,File4_10.Desca as Desca2,file5_10.Desca as BankDesca " & _
        " FROM ((file5_20 LEFT JOIN FILE3_10 ON FILE5_20.CODE1 = FILE3_10.CODE)inner JOIN file4_10 on file5_20.code2 = file4_10.code) left join file5_10 on file5_20.id_bank = file5_10.code "


If IsDate(xdate1.text) Then
    aHeader(0) = BetweenString(xdate1.text, xDate2.text)
    cString = cString & turnFound(cString) & " dateBank >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    aHeader(0) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    cString = cString & turnFound(cString) & " dateBank <= " & DateSq(xDate2.text)
End If

If xCharge.text <> "" Then
    cString = cString & turnFound(cString) & " code1 = " & MyParn(xCharge.text)
    aHeader(2) = "[" & "··⁄„Ì· :" & xCodedesca1.Caption & "]"
End If



If xCode2.text <> "" Then
    cString = cString & turnFound(cString) & " code2 = " & MyParn(xCode2.text)
    aHeader(4) = "[" & "··„Ê—œ :" & xCodeDesca2.Caption & "]"
End If

If xBank.BoundText <> "" Then
    cString = cString & turnFound(cString) & " ID_BANK = " & MyParn(xBank.BoundText)
    aHeader(6) = "[" & "«·»‰ﬂ : " & xBank.text & "]"
End If

If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & " Box = " & MyParn(xbox.BoundText)
    aHeader(7) = "[" & "«·Œ“‰… :" & xbox.text & "]"
End If

cString = cString & " ORDER BY DateBank"
    
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !CHK_ID
    If Not IsNull(!desca1) Then
        temptable!str2 = "⁄„Ì·"
        temptable!str3 = !desca1
    ElseIf Not IsNull(!Desca2) Then
         temptable!str2 = "„Ê—œ"
         temptable!str3 = !Desca2
    End If
    If Not IsNull(!DESCA) Then
        temptable!str3 = temptable!str3 & IIf(IsNull(temptable!str3), "", " - ") & !DESCA
    End If
    temptable!str4 = !BankDesca
    temptable!Date1 = !DateBank
    'temptable!date2 = !date_R
    temptable!date2 = !date_1
    temptable!VAL1 = !Value
    
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 8))
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\CHQ4.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub
Private Sub doprint8()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim aHeader(7)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "select file5_20.*,FILE3_10.DESCA as Desca,file5_10.desca as bankDesca " & _
        " FROM (file5_20 INNER JOIN FILE3_10 ON FILE5_20.CODE1 = FILE3_10.CODE) left join file5_10 on file5_20.id_bank = file5_10.code "

If IsDate(xdate1.text) Then
    aHeader(0) = BetweenString(xdate1.text, xDate2.text)
    cString = cString & turnFound(cString) & "date_R >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    aHeader(0) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    cString = cString & turnFound(cString) & "date_R <= " & DateSq(xDate2.text)
End If

If xCharge.text <> "" Then
    cString = cString & turnFound(cString) & " code1 = " & MyParn(xCharge.text)
    aHeader(1) = "[" & "··⁄„Ì· :" & xCodedesca1.Caption & "]"
End If


If xCode2.text <> "" Then
    cString = cString & turnFound(cString) & " code2 = " & MyParn(xCode2.text)
    aHeader(3) = "[" & "··„Ê—œ :" & xCodeDesca2.Caption & "]"
End If


If xBank.BoundText <> "" Then
    cString = cString & turnFound(cString) & " ID_BANK = " & MyParn(xBank.BoundText)
    aHeader(5) = "[" & "«·»‰ﬂ : " & xBank.text & "]"
End If

If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & " Box = " & MyParn(xbox.BoundText)
    aHeader(6) = "[" & "«·Œ“‰… :" & xbox.text & "]"
End If

If xNoBank.Value <> 0 Then
    cString = cString & turnFound(cString) & " (ID_BANK is Null)"
    aHeader(7) = "[" & "‘Ìﬂ«  €Ì— „Êœ⁄… ›Ï »‰ﬂ " & "]"
End If


cString = cString & " ORDER BY DATE_R"
    
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !CHK_ID
    If Not IsNull(!desca1) Then
        temptable!str2 = "⁄„Ì·"
        temptable!str3 = !desca1
    End If
    temptable!str3 = !DESCA
    temptable!str4 = !BankDesca
    'temptable!Date1 = !date_1
    temptable!Date1 = !date_R
    temptable!date2 = !DateBank
    temptable!VAL1 = !Value
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 8))
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\CHQ8.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub
Private Sub Form_Load()
openCon con


DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select * From File4_50 ORDER BY desca"
Set xGroup2.RowSource = DATA2
xGroup2.ListField = "Desca"
xGroup2.BoundColumn = "Code"

data3.ConnectionString = strCon
data3.RecordSource = "Select * From File5_10 ORDER BY desca"
Set xBank.RowSource = data3
xBank.ListField = "Desca"
xBank.BoundColumn = "Code"

data4.ConnectionString = strCon
data4.RecordSource = "Select * From File0_50 ORDER BY desca"
Set xbox.RowSource = data4
xbox.ListField = "Desca"
xbox.BoundColumn = "Code"
Frame2.Visible = publicFlag <> 8
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub XCHARGE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(1, 1)
    
    Set Generalarray(0) = Me
    
    Generalarray(1) = "Select code ,DescA From file3_10 "
    Generalarray(2) = "Order by code"
    Generalarray(3) = 5000
    Generalarray(5) = False
    
    listarray(0, 0) = "«·»Ì«‰"
    listarray(0, 1) = "(%%DESCA%%)"
    
    GrdArray(0, 0) = "«·ﬂÊœ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«·»Ì«‰"
    GrdArray(1, 1) = 6000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "≈” ⁄·«„ "
    Search3.Show 1
End If
End Sub
Private Sub XCHARGE_LostFocus()
xCodedesca1.Caption = ""
xCodedesca1.Caption = GetDesca("Select desca from file3_10 where code = " & MyParn(xCharge.text), con) & ""
End Sub


Private Sub xCODE2_LostFocus()
xCodeDesca2.Caption = ""
xCodeDesca2.Caption = GetDesca("Select desca from file4_10 where code = " & MyParn(xCode2.text), con) & ""
End Sub
Private Sub xCode2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(1, 1)
    
    Set Generalarray(0) = Me
    
    Generalarray(1) = "Select code ,DescA From file4_10 "
    Generalarray(2) = "Order by code"
    Generalarray(3) = 5000
    Generalarray(5) = False
    
    listarray(0, 0) = "«·»Ì«‰"
    listarray(0, 1) = "(%%DESCA%%)"
    
    GrdArray(0, 0) = "«·ﬂÊœ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "«·»Ì«‰"
    GrdArray(1, 1) = 6000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "≈” ⁄·«„ "
    Search3.Show 1
End If
End Sub
Sub myProc()
If ActiveControl.Name = xCharge.Name Then
    xCharge.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    xCodedesca1.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
Else
    xCode2.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    xCodeDesca2.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
End If
Unload Search3
End Sub

