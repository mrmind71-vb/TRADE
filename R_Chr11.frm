VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form R_Chr11 
   Caption         =   "ÊÞÇÑíÑ ÇáãÕÇÑíÝ"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "ÚÑÖ"
      Height          =   390
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   510
      Width           =   1290
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "ÎÑæÌ"
      Height          =   390
      Left            =   195
      RightToLeft     =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   510
      Width           =   1215
   End
   Begin VB.TextBox date2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3075
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   555
      Width           =   1365
   End
   Begin VB.TextBox Date1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3075
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   150
      Width           =   1365
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Çáì ÊÇÑíÎ :"
      Height          =   195
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   720
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "ãä ÊÇÑíÎ :"
      Height          =   195
      Left            =   4575
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   225
      Width           =   675
   End
End
Attribute VB_Name = "R_Chr11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ChargeTable As Recordset
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
Dim SourceTable As Recordset
Dim TargetTable As Recordset
tempdb.Execute "DELETE * FROM TEMP"
Set TargetTable = tempdb.CreateDynaset("TEMP")
cString = "SELECT SUM(FILE8_90!VALUE) AS T_Value, FILE8_90.MAN ," & _
          "FILE1_70.DESCA AS ManDesca " & _
          " FROM FILE8_90 LEFT JOIN FILE1_70 ON FILE8_90.MAN = FILE1_70.CODE WHERE FILE1_70.FLAG = 9 "
If IsDate(Date1.Text) Then cString = cString & " AND Date >= " & DateSql(Date1.Text)
If IsDate(date2.Text) Then cString = cString & " AND Date <= " & DateSql(date2.Text)
cString = cString & " GROUP BY FILE8_90.MAN , FILE1_70.DESCA "
Set SourceTable = mydb.OpenRecordset(cString, dbOpenSnapshot)
If SourceTable.RecordCount = 0 Then
    MsgBox "áÇ ÊæÌÏ ÈíÇäÇÊ Ýì ÇáÊÞÑíÑ ¿"
    Exit Sub
End If
With SourceTable
Do
    TargetTable.AddNew
    TargetTable.str1 = .MAN
    TargetTable.str2 = .MANDESCA
    TargetTable.VAL1 = .T_value

    TargetTable.STR7 = "ÅÌãÇáì ãÓÍæÈÇÊ ÔÑßÇÁ"
    TargetTable.str8 = "ãä ÊÇÑíÎ " & Date1.Text & " Åáì ÊÇÑíÎ " & date2.Text
    TargetTable.str9 = firsttitle
    TargetTable.str10 = Secondtitle

    TargetTable.Update
    SourceTable.MoveNext
Loop Until SourceTable.EOF
End With
MyWs.BeginTrans
MyWs.CommitTrans

Report1.ReportFileName = App.Path & "\Reports\R_Chrg11.rpt"
Report1.DataFiles(0) = TempPath
Report1.Action = 1

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
    Date1.Text = ""
    date2.Text = ""
End Sub
Function MYVALID()
If Not IsDate(Date1.Text) Then Exit Function
If Not IsDate(date2.Text) Then Exit Function
If DateValue(Date1.Text) > DateValue(date2.Text) Then Exit Function
MYVALID = True
End Function
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
ChargeTable.Close
Set ChargeTable = Nothing
End Sub
