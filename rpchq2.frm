VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpChq2 
   Caption         =   " ﬁ«—Ì— «Ê—«ﬁ «·ﬁ»÷"
   ClientHeight    =   4830
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
   ScaleHeight     =   4830
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
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
      Left            =   1575
      Picture         =   "rpchq2.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "„”Õ"
      Top             =   4005
      Width           =   1365
   End
   Begin VB.CommandButton CmdExit 
      Height          =   465
      Left            =   135
      MaskColor       =   &H00FFFFFF&
      Picture         =   "rpchq2.frx":23A5
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "Œ—ÊÃ"
      Top             =   4005
      UseMaskColor    =   -1  'True
      Width           =   1410
   End
   Begin VB.Frame Frame4 
      Height          =   1050
      Left            =   765
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   0
      Width           =   5550
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   540
         Width           =   1725
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2835
         RightToLeft     =   -1  'True
         TabIndex        =   17
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
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   225
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   2025
      Width           =   6180
      Begin VB.TextBox XCHARGE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo xbranch 
         Height          =   315
         Left            =   1260
         TabIndex        =   23
         Top             =   585
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·›—⁄ "
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
         TabIndex        =   24
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„’—Ê› "
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
         TabIndex        =   4
         Top             =   225
         Width           =   780
      End
      Begin VB.Label xCodeDesca2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   3345
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1035
      Width           =   6180
      Begin VB.TextBox xCode1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo xGroup1 
         Height          =   315
         Left            =   1290
         TabIndex        =   7
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
         TabIndex        =   10
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label xCodedesca1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   225
         Width           =   570
      End
   End
   Begin VB.CommandButton CmdApply 
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
      Left            =   2970
      Picture         =   "rpchq2.frx":24EF
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "⁄—÷"
      Top             =   4005
      Width           =   1410
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   4995
      Top             =   4140
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
      TabIndex        =   11
      Top             =   2970
      Width           =   6180
      Begin MSDataListLib.DataCombo xBox 
         Height          =   315
         Left            =   1305
         TabIndex        =   12
         Top             =   540
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xBank 
         Height          =   315
         Left            =   1305
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
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
Attribute VB_Name = "rpChq2"
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
If publicFlag = 5 Then doprint1
If publicFlag = 6 Then doprint2
If publicFlag = 7 Then doprint3
End Sub
Private Sub doprint1()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim aHeader(7)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "select FILE5_21.*,File4_10.Desca as Desca1,FILE3_10.DESCA as Desca2,FILE5_10.DESCA as BankDesca" & _
        " FROM ((FILE5_21 inner JOIN file4_10 ON FILE5_21.CODE1 = FILE4_10.CODE)Left join File3_10 on FILE5_21.CHARGE = file3_10.code) LEFT JOIN FILE5_10 ON FILE5_21.ID_BANK = FILE5_10.CODE " & _
        " WHERE CLOSED = '0' "

If IsDate(xdate1.text) Then
    aHeader(0) = BetweenString(xdate1.text, xDate2.text)
    cString = cString & turnFound(cString) & "date_1 >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    aHeader(0) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    cString = cString & turnFound(cString) & "date_1 <= " & DateSq(xDate2.text)
End If

If xCode1.text <> "" Then
    cString = cString & turnFound(cString) & " code1 = " & MyParn(xCode1.text)
    aHeader(1) = "[" & "··„Ê—œ :" & xCodedesca1.Caption & "]"
End If

If xGroup1.BoundText <> "" Then
    cString = cString & turnFound(cString) & " File4_10.[GROUP] = " & MyParn(xGroup1.BoundText)
    aHeader(2) = "[" & "·„Ã„Ê⁄… „Ê—œÌ‰ :" & xGroup1.text & "]"
End If

If xbranch.BoundText <> "" Then
    cString = cString & turnFound(cString) & " File5_21.[branch] = " & MyParn(xbranch.BoundText)
    aHeader(2) = "[" & "··›—⁄ :" & xbranch.text & "]"
End If

If xCharge.text <> "" Then
    cString = cString & turnFound(cString) & " CHARGE = " & MyParn(xCharge.text)
    aHeader(3) = "[" & "··„’—Ê› :" & xCodeDesca2.Caption & "]"
End If


If xBank.BoundText <> "" Then
    cString = cString & turnFound(cString) & " ID_BANK = " & MyParn(xBank.BoundText)
    aHeader(5) = "[" & "«·»‰ﬂ : " & xBank.text & "]"
End If

If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & " Box = " & MyParn(xbox.BoundText)
    aHeader(6) = "[" & "«·Œ“‰… :" & xbox.text & "]"
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
        temptable!str2 = "„Ê—œ"
        temptable!str3 = !desca1
    ElseIf Not IsNull(!Desca2) Then
         temptable!str2 = "⁄„Ì·"
         temptable!str3 = !Desca2
    End If
    If Not IsNull(!DESCA) Then
        temptable!str3 = temptable!str3 & IIf(IsNull(temptable!str3), "", " - ") & !DESCA
    End If
    temptable!Date1 = !date_1
    temptable!date2 = !date_R
'    temptable!date3 = !DateBank
    temptable!str4 = !BankDesca
    temptable!VAL1 = !Value
    
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 8))
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\CHQ5.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub
Private Sub doprint2()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim aHeader(7)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "select FILE5_21.*,File4_10.Desca as Desca1,FILE3_10.DESCA as Desca2,FILE5_10.DESCA as BankDesca" & _
        " FROM ((FILE5_21 inner JOIN file4_10 ON FILE5_21.CODE1 = FILE4_10.CODE)Left join File3_10 on FILE5_21.CHARGE = file3_10.code) LEFT JOIN FILE5_10 ON FILE5_21.ID_BANK = FILE5_10.CODE " & _
        " WHERE CLOSED = '2' "

If IsDate(xdate1.text) Then
    aHeader(0) = BetweenString(xdate1.text, xDate2.text)
    cString = cString & turnFound(cString) & "date_3 >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    aHeader(0) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    cString = cString & turnFound(cString) & "date_3 <= " & DateSq(xDate2.text)
End If

If xCode1.text <> "" Then
    cString = cString & turnFound(cString) & " code1 = " & MyParn(xCode1.text)
    aHeader(1) = "[" & "··„Ê—œ :" & xCodedesca1.Caption & "]"
End If
If xbranch.BoundText <> "" Then
    cString = cString & turnFound(cString) & " File5_21.[branch] = " & MyParn(xbranch.BoundText)
    aHeader(2) = "[" & "··›—⁄ :" & xbranch.text & "]"
End If

If xGroup1.BoundText <> "" Then
    cString = cString & turnFound(cString) & " File4_10.[GROUP] = " & MyParn(xGroup1.BoundText)
    aHeader(2) = "[" & "·„Ã„Ê⁄… „Ê—œÌ‰ :" & xGroup1.text & "]"
End If

If xCharge.text <> "" Then
    cString = cString & turnFound(cString) & " CHARGE = " & MyParn(xCharge.text)
    aHeader(3) = "[" & "··„’—Ê› :" & xCodeDesca2.Caption & "]"
End If

If xBank.BoundText <> "" Then
    cString = cString & turnFound(cString) & " ID_BANK = " & MyParn(xBank.BoundText)
    aHeader(5) = "[" & "«·»‰ﬂ : " & xBank.text & "]"
End If

If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & " Box = " & MyParn(xbox.BoundText)
    aHeader(6) = "[" & "«·Œ“‰… :" & xbox.text & "]"
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
        temptable!str2 = "„Ê—œ"
        temptable!str3 = !desca1
    ElseIf Not IsNull(!Desca2) Then
         temptable!str2 = "⁄„Ì·"
         temptable!str3 = !Desca2
    End If
    If Not IsNull(!DESCA) Then
        temptable!str3 = temptable!str3 & IIf(IsNull(temptable!str3), "", " - ") & !DESCA
    End If
    temptable!Date1 = !date_3
    temptable!date2 = !date_1
    temptable!VAL1 = !Value
    temptable!str4 = !BankDesca
    
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 8))
    temptable.Update
    .MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\CHQ6.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub
Private Sub doprint3()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim aHeader(7)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "select FILE5_21.*,File4_10.Desca as Desca1,FILE3_10.DESCA as Desca2,FILE5_10.DESCA as BankDesca" & _
        " FROM ((FILE5_21 inner JOIN file4_10 ON FILE5_21.CODE1 = FILE4_10.CODE)Left join File3_10 on FILE5_21.CHARGE = file3_10.code) LEFT JOIN FILE5_10 ON FILE5_21.ID_BANK = FILE5_10.CODE " & _
        " WHERE CLOSED = '1' "

If IsDate(xdate1.text) Then
    aHeader(0) = BetweenString(xdate1.text, xDate2.text)
    cString = cString & turnFound(cString) & "date_3 >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    aHeader(0) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    cString = cString & turnFound(cString) & "date_3 <= " & DateSq(xDate2.text)
End If
If xbranch.BoundText <> "" Then
    cString = cString & turnFound(cString) & " File5_21.[branch] = " & MyParn(xbranch.BoundText)
    aHeader(2) = "[" & "··›—⁄ :" & xbranch.text & "]"
End If

If xCode1.text <> "" Then
    cString = cString & turnFound(cString) & " code1 = " & MyParn(xCode1.text)
    aHeader(1) = "[" & "··„Ê—œ :" & xCodedesca1.Caption & "]"
End If

If xGroup1.BoundText <> "" Then
    cString = cString & turnFound(cString) & " File4_10.[GROUP] = " & MyParn(xGroup1.BoundText)
    aHeader(2) = "[" & "·„Ã„Ê⁄… „Ê—œÌ‰ :" & xGroup1.text & "]"
End If

If xCharge.text <> "" Then
    cString = cString & turnFound(cString) & " CHARGE = " & MyParn(xCharge.text)
    aHeader(3) = "[" & "··„’—Ê› :" & xCodeDesca2.Caption & "]"
End If

If xBank.BoundText <> "" Then
    cString = cString & turnFound(cString) & " ID_BANK = " & MyParn(xBank.BoundText)
    aHeader(5) = "[" & "«·»‰ﬂ : " & xBank.text & "]"
End If

If xbox.BoundText <> "" Then
    cString = cString & turnFound(cString) & " Box = " & MyParn(xbox.BoundText)
    aHeader(6) = "[" & "«·Œ“‰… :" & xbox.text & "]"
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
        temptable!str2 = "„Ê—œ"
        temptable!str3 = !desca1
    ElseIf Not IsNull(!Desca2) Then
         temptable!str2 = "⁄„Ì·"
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
main.Report1.ReportFileName = App.Path & "\Reports\CHQ7.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
End Sub
Private Sub Form_Load()
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "Select * From File4_50 ORDER BY desca"
Set xGroup1.RowSource = data1
xGroup1.ListField = "Desca"
xGroup1.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select * From branch ORDER BY desca"
Set xbranch.RowSource = DATA2
xbranch.ListField = "Desca"
xbranch.BoundColumn = "Code"


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
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xCode1_KeyDown(KeyCode As Integer, Shift As Integer)
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
Private Sub xCode1_LostFocus()
xCodedesca1.Caption = ""
xCodedesca1.Caption = GetDesca("Select desca from file4_10 where code = " & MyParn(xCode1.text), con) & ""
End Sub
Private Sub XCHARGE_LostFocus()
    xCodeDesca2.Caption = ""
    xCodeDesca2.Caption = GetDesca("Select desca from file8_51 where code = " & MyParn(xCharge.text), con) & ""
End Sub
Private Sub XCHARGE_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    Dim Generalarray(5)
    Dim listarray(0, 5)
    Dim GrdArray(1, 1)
    
    Set Generalarray(0) = Me
    
    Generalarray(1) = "Select code ,DescA From file8_51 "
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
If ActiveControl.Name = xCode1.Name Then
    xCode1.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    xCodedesca1.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
Else
    xCharge.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    xCodeDesca2.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
End If
Unload Search3
End Sub

