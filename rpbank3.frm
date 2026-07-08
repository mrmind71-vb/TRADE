VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpBank3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ›’Ì·Ì Õ—ﬂ… »‰ﬂÌ…"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5910
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
   ScaleHeight     =   3780
   ScaleWidth      =   5910
   Begin VB.Frame Frame1 
      Height          =   2490
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   5775
      Begin VB.TextBox xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   900
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1260
         Width           =   3345
      End
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
      Begin MSDataListLib.DataCombo XCODE 
         Height          =   315
         Left            =   1530
         TabIndex        =   11
         Top             =   1620
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "«·»‰œ :"
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
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1665
         Width           =   420
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·»Ì«‰ :"
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
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1350
         Width           =   525
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
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   990
         Width           =   825
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   5805
      Top             =   450
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
      Left            =   1305
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3150
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
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3150
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
   Begin VB.Frame Frame2 
      Height          =   555
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2565
      Width           =   5775
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "„”ÕÊ»«  ›ﬁÿ"
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
         Index           =   2
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   1770
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "«Ìœ«⁄«  ›ﬁÿ"
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
         Index           =   1
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "«·ﬂ·"
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
         Index           =   0
         Left            =   4635
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   180
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin MSAdodcLib.Adodc DATA3 
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
Attribute VB_Name = "rpBank3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim ChargeTable As ADODB.Recordset
Private Sub CmdApply_Click()
    doprint1
End Sub
Private Sub CmdExit_Click()
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
xDate1.text = ""
xdate2.text = ""

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE5_10 ORDER BY DESCA"
Set xbank.RowSource = data1
xbank.ListField = "Desca"
xbank.BoundColumn = "code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM FILE5_00 ORDER BY DESCA"

Set xCode.RowSource = DATA2
xCode.ListField = "Desca"
xCode.BoundColumn = "Code"



End Sub
Function MYVALID() As Boolean
If Not IsDate(xDate1.text) And (xDate1.text <> "") Then
    MsgBox "«· «—ÌŒ «·«Ê· €Ì— ’«·Õ"
    Exit Function
End If
If Not IsDate(xdate2.text) And (xdate2.text <> "") Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(5)
If Not MYVALID Then Exit Sub
Dim i As Integer, nPrevious As Double

Dim sourcetable As ADODB.Recordset
Dim temptable As ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
Set sourcetable = New ADODB.Recordset
cString = "Select bankmove.bank , bankmove.* ,FILE5_10.DESCA AS FILE5_10DESCA FROM bankmove INNER join FILE5_10 ON bankmove.BANK = FILE5_10.code WHERE TYPE <= 4.5 "
If Option1(1).Value Then
    cString = cString & turnFound(cString) & " Value1  <> 0"
    aHeader(4) = "«·«Ìœ«⁄«  ›ﬁÿ"
End If

If Option1(2).Value Then
    cString = cString & turnFound(cString) & " Value2 <> 0"
    aHeader(4) = "«·„”ÕÊ»«  ›ﬁÿ"
End If


If Trim(xbank.BoundText) <> "" Then
    cString = cString & turnFound(cString) & " bankmove.BANK = " & MyParn(xbank.BoundText)
    aHeader(0) = "[" & "«·»‰ﬂ : " & xbank.text & "]"
End If

If Trim(xCode.BoundText) <> "" Then
    cString = cString & turnFound(cString) & " bankmove.code = " & MyParn(xCode.BoundText)
    aHeader(1) = "[" & "«·»‰œ : " & xCode.text & "]"
End If


If IsDate(xDate1.text) Then
    cString = cString & turnFound(cString) & "Date >= " & DateSq(xDate1.text)
    aHeader(2) = "[" & BetweenString(xDate1.text, xdate2.text) & "]"
End If
          
If IsDate(xdate2.text) Then
    cString = cString & turnFound(cString) & "Date <= " & DateSq(xdate2.text)
    aHeader(2) = "[" & BetweenString(xDate1.text, xdate2.text) & "]"
End If
                   
If Trim(xDesca.text) <> "" Then
    cString = cString & turnFound(cString) & " bankmove.desca like " & MyParnAll(xDesca.text)
    aHeader(3) = "[" & "«·»Ì«‰ : " & xDesca.text & "]"
End If
                  
cString = cString & " Order by [Date],Doc_no,value2"

sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
    Do Until sourcetable.EOF
        i = i + 1
        temptable.AddNew
        temptable!str1 = !doc_no
        temptable!str2 = !TypeDesca
        temptable!str3 = !Desca
        temptable!str4 = !FILE5_10desca
        temptable!Date1 = !Date
        temptable!VAL1 = Val(!value1 & "")
        temptable!val2 = Val(!Value2 & "")
        temptable!STR21 = TurnValue(retHeader(aHeader, 0, 5))
        temptable!Val10 = i
        temptable.Update
        sourcetable.MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ›Ï «· ﬁ—Ì— ø"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    Main.REPORT1.ReportFileName = App.Path & "\Reports\BANK3.rpt"
    Main.REPORT1.DataFiles(0) = tempFile
    Main.REPORT1.Action = 1
End If
If temptable.State = adStateOpen Then temptable.Close
If sourcetable.State = adStateOpen Then sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

