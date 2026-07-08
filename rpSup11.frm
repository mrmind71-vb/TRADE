VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpSup11 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
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
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   420
      Left            =   1395
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2655
      Width           =   1275
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2655
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   2580
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   6180
      Begin VB.CheckBox Check5 
         Alignment       =   1  'Right Justify
         Caption         =   " ”ÊÌ« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2115
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÌÊœ Õ”«»« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2115
         Value           =   1  'Checked
         Width           =   1365
      End
      Begin VB.CheckBox Check3 
         Alignment       =   1  'Right Justify
         Caption         =   "‘Ìﬂ« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2115
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "œ›⁄«  ‰ﬁœÌ…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2115
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   " ﬁ”Ì„ «·œ›⁄«  ··„Ê—œÌ‰"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1755
         Width           =   3210
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1005
         Width           =   1680
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   615
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xmosm 
         Height          =   315
         Left            =   1305
         TabIndex        =   10
         Top             =   1395
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "„Ê”„ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4845
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1395
         Width           =   1005
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   3345
      End
      Begin VB.Label Label1 
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   315
         Width           =   570
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1080
         Width           =   465
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   720
         Width           =   765
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
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
Attribute VB_Name = "rpSup11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If TypeOf ActiveControl Is DBCombo And KeyCode = 46 Then ActiveControl.BoundText = ""
End Sub
Private Sub Form_Load()
openCon con

data4.ConnectionString = strCon
data4.RecordSource = "Select MOSM,DescA From MOSM     ORDER by DATE "
Set xMosm.RowSource = data4
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"
xMosm.BoundText = cPMosm

End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate2.text) And Trim(xDate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(2), cHead2 As String
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable


cHead2 = "  ﬁ—Ì— „Ê—œÌ‰ "
cString1 = " SELECT  * , ' œ›⁄… ‰ﬁœÏ ' as cashtype FROM Q_CASH2  where CODE IS NOT NULL "
If Check2.Value = 0 Then
    cString1 = cString1 & " and 1= 0 "
Else
    cHead2 = cHead2 & " ‰ﬁœÌ…  "
End If
If Trim(xMosm.BoundText) <> "" Then
    cString1 = cString1 & " and mosm = " & MyParn(xMosm.BoundText)
    aHeader(0) = "[" & "„Ê”„ : " & xMosm.BoundText & "]"
End If

If Trim(xCode.text) <> "" Then
    cString1 = cString1 & " and CODE = " & MyParn(xCode.text)
    aHeader(2) = "[" & "··„Ê—œ : " & xCodeDesca.Caption & "]"
End If

          
If IsDate(xdate1.text) Then
    cString1 = cString1 & " and date >= " & DateSq(xdate1.text)
    aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
End If

If IsDate(xDate2.text) Then
    cString1 = cString1 & " and date <= " & DateSq(xDate2.text)
    aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
End If

'''' chq

cString1 = cString1 & " union all  SELECT  * , ' ‘Ìﬂ ' FROM Q_chq where CODE1 IS NOT NULL "
If Check3.Value = 0 Then
    cString1 = cString1 & " and 1= 0 "
Else
    cHead2 = cHead2 & " ‘Ìﬂ«  "
End If
    
If Trim(xMosm.BoundText) <> "" Then
    cString1 = cString1 & " AND mosm = " & MyParn(xMosm.BoundText)
End If

If Trim(xCode.text) <> "" Then
    cString1 = cString1 & " and CODE1 = " & MyParn(xCode.text)
End If

          
If IsDate(xdate1.text) Then
    cString1 = cString1 & " and date_r >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    cString1 = cString1 & " and date_r <= " & DateSq(xDate2.text)
End If

'''' ﬁÌÊœ

cString1 = cString1 & " union all  SELECT   file4_11.code , file4_10.desca , doc_id , date , (pay-sal) , file4_11.desca , mosm , '' , ' ﬁÌœ ' as cashtype FROM file4_11 inner join file4_10 on file4_10.code = file4_11.code where type = 'T' "
If Check4.Value = 0 Then
    cString1 = cString1 & " and 1= 0 "
Else
    cHead2 = cHead2 & " ﬁÌÊœ "
End If
If Trim(xMosm.BoundText) <> "" Then
    cString1 = cString1 & " and mosm = " & MyParn(xMosm.BoundText)
    aHeader(0) = "[" & "„Ê”„ : " & xMosm.BoundText & "]"
End If

If Trim(xCode.text) <> "" Then
    cString1 = cString1 & " and file4_11.CODE = " & MyParn(xCode.text)
    aHeader(2) = "[" & "··„Ê—œ : " & xCodeDesca.Caption & "]"
End If

          
If IsDate(xdate1.text) Then
    cString1 = cString1 & " and date >= " & DateSq(xdate1.text)
    aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
End If

If IsDate(xDate2.text) Then
    cString1 = cString1 & " and date <= " & DateSq(xDate2.text)
    aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
End If


cString1 = cString1 & " union all SELECT  code ,SUPPDESCA , doc_no , date , value , desca , mosm , box  , '  ”ÊÌ…  ' as cashtype FROM Q_CASH0  where CODE IS NOT NULL "
If Check5.Value = 0 Then
    cString1 = cString1 & " and 1= 0 "
Else
    cHead2 = cHead2 & "  ”ÊÌ«  "
End If

If Trim(xMosm.BoundText) <> "" Then
    cString1 = cString1 & " and mosm = " & MyParn(xMosm.BoundText)
    aHeader(0) = "[" & "„Ê”„ : " & xMosm.BoundText & "]"
End If

If Trim(xCode.text) <> "" Then
    cString1 = cString1 & " and CODE = " & MyParn(xCode.text)
    aHeader(2) = "[" & "··„Ê—œ : " & xCodeDesca.Caption & "]"
End If

          
If IsDate(xdate1.text) Then
    cString1 = cString1 & " and date >= " & DateSq(xdate1.text)
    aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
End If

If IsDate(xDate2.text) Then
    cString1 = cString1 & " and date <= " & DateSq(xDate2.text)
    aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
End If

With sourcetable
sourcetable.Open cString1, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str11 = sourcetable!CODE
    temptable!str1 = sourcetable!SUPPDESCA
    temptable!str2 = sourcetable!doc_no
    temptable!str3 = sourcetable!cashtype & "  " & sourcetable!DESCA
    temptable!Date1 = sourcetable!Date

    temptable!VAL1 = !Value
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!str5 = cHead2
   ' If IsDate(sourcetable!Check_Date) Then
'    temptable!str12 = Format(sourcetable!Check_Date, "DD-MM-YYYY")
'    temptable!str11 = TurnValue(sourcetable!Check_No, "", Null)
    
    temptable.Update
    sourcetable.MoveNext
Loop
End With
temptable.Requery
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If Check1.Value Then
    main.Report1.ReportFileName = App.Path & "\Reports\Sup11.rpt"
Else
    main.Report1.ReportFileName = App.Path & "\Reports\Sup11G.rpt"
End If
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    suplookup
End If
End Sub
Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""
If xCode.text = "" Then Exit Sub
xCodeDesca.Caption = GetDesca("select desca from FILE4_10 where code = " & MyParn(xCode.text), con) & ""
End Sub
Sub myProc()
    ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End Sub
Private Sub suplookup()
    Dim Generalarray(5)
    Dim listarray(0, 4)
    Dim GrdArray(1, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "Select Code, DescA From FILE4_10"
    Generalarray(2) = "Order by file4_10.Desca"
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
    listarray(0, 1) = "(%%DESCA%%) "
    
    GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "≈”„ «·„Ê—œ"
    GrdArray(1, 1) = 3000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "«” ⁄·«„"
    Search3.Show 1
End Sub

