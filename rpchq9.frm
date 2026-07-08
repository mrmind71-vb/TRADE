VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpChq9 
   Caption         =   " ﬁ«—Ì— «Ê—«ﬁ «·ﬁ»÷"
   ClientHeight    =   2895
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
   ScaleHeight     =   2895
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
      TabIndex        =   12
      Top             =   2295
      Width           =   2805
   End
   Begin VB.Frame Frame4 
      Height          =   1050
      Left            =   765
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   5550
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   540
         Width           =   1725
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   8
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
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   225
         Width           =   330
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
      Top             =   2250
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
      Top             =   2250
      Width           =   1140
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   150
      Top             =   450
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
      TabIndex        =   2
      Top             =   1065
      Width           =   6180
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   1350
         TabIndex        =   3
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
Attribute VB_Name = "rpChq9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function MYVALID()
If Not IsDate(xDate1.Text) Then Exit Function
If Not IsDate(xdate2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub CMDEXIT_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xDate1.Text = ""
xdate2.Text = ""
End Sub
Private Sub CmdApply_Click()
doprint9
End Sub
Private Sub doprint9()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim aHeader(7)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "select file5_20.*,file5_10.desca as bankDesca " & _
        " FROM file5_20 left join file5_10 on file5_20.id_bank = file5_10.code " & _
        " WHERE CLOSED = '0' and file5_20.code1 is null "

If IsDate(xDate1.Text) Then
    aHeader(0) = BetweenString(xDate1.Text, xdate2.Text)
    cString = cString & turnFound(cString) & "date_1 >= " & DateSq(xDate1.Text)
End If

If IsDate(xdate2.Text) Then
    aHeader(0) = "[" & BetweenString(xDate1.Text, xdate2.Text) & "]"
    cString = cString & turnFound(cString) & "date_1 <= " & DateSq(xdate2.Text)
End If


If xBank.BoundText <> "" Then
    cString = cString & turnFound(cString) & " ID_BANK = " & MyParn(xBank.BoundText)
    aHeader(5) = "[" & "«·»‰ﬂ : " & xBank.Text & "]"
End If

If XBOX.BoundText <> "" Then
    cString = cString & turnFound(cString) & " Box = " & MyParn(XBOX.BoundText)
    aHeader(6) = "[" & "«·Œ“‰… :" & XBOX.Text & "]"
End If

If xNoBank.Value <> 0 Then
    cString = cString & turnFound(cString) & " isNull(ID_BANK)"
    aHeader(7) = "[" & "‘Ìﬂ«  €Ì— „Êœ⁄… ›Ï »‰ﬂ " & "]"
End If


cString = cString & " ORDER BY DATE_1"
    
sourcetable.Open cString, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
With sourcetable
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !CHK_ID
'    If Not IsNull(!desca1) Then
'        temptable!str2 = "⁄„Ì·"
'        temptable!str3 = !desca1
'    ElseIf Not IsNull(!Desca2) Then
'         temptable!str2 = "„Ê—œ"
'         temptable!str3 = !Desca2
'    End If
    If IsNull(!DESCA) Then
        temptable!str3 = !NAME4
    Else
        temptable!str3 = !DESCA
    End If
    temptable!str4 = !BankDesca
    temptable!Date1 = !date_1
    'temptable!date2 = !date_R
    temptable!date2 = !DateBank
    temptable!val1 = !Value
    temptable!str21 = " ‘Ìﬂ«  ﬂ»«— «·⁄„·«¡ " & TurnValue(retHeader(aHeader, 0, 8))
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
Main.Report1.ReportFileName = App.Path & "\Reports\CHQ1.rpt"
Main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.Report1.Action = 1
End Sub
Private Sub Form_Load()


data3.ConnectionString = CON.ConnectionString
data3.RecordSource = "Select * From File5_10 ORDER BY desca"
Set xBank.RowSource = data3
xBank.ListField = "Desca"
xBank.BoundColumn = "Code"

data4.ConnectionString = CON.ConnectionString
data4.RecordSource = "Select * From File0_50 ORDER BY desca"
Set XBOX.RowSource = data4
XBOX.ListField = "Desca"
XBOX.BoundColumn = "Code"

End Sub
