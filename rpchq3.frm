VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpChq3 
   Caption         =   " ﬁ«—Ì— «·‘Ìﬂ« "
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6075
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
   ScaleHeight     =   1950
   ScaleWidth      =   6075
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
      TabIndex        =   9
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
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   1365
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   45
      Width           =   5685
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   1320
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   945
         Width           =   1320
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1320
      End
      Begin VB.TextBox xname 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   3435
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   4965
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   600
         Width           =   330
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   4965
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1050
         Width           =   450
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ :"
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
         Left            =   4965
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   360
      End
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   3555
      Top             =   2025
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
End
Attribute VB_Name = "rpChq3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
Dim ClientTable As ADODB.Recordset
Function MYVALID()
If (Not IsDate(xDate1.Text)) And xDate1.Text <> "" Then Exit Function
If (Not IsDate(xdate2.Text)) And xdate2.Text <> "" Then Exit Function
If xCode.Text = "" Then Exit Function
MYVALID = True
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
xDate1.Text = ""
xdate2.Text = ""
End Sub
Private Sub Form_Load()
Set ClientTable = New ADODB.Recordset
Set sourcetable = New ADODB.Recordset
Set temptable = New ADODB.Recordset
ClientTable.Open IIf(lCust, "file3_10", "file4_10"), CON, adOpenStatic, adLockReadOnly, adCmdTable
'If lCust Then
'    xCode.Visible = False
'    Me.Label1.Visible = False
'End If
End Sub
Private Sub CmdApply_Click()
contemp.Execute "DELETE * FROM TEMP"
If temptable.State = adStateOpen Then temptable.Close
If sourcetable.State = adStateOpen Then sourcetable.Close
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cHead = cHead & "Ê –·ﬂ „‰  «—ÌŒ " & xDate1.Text & "≈·Ï  «—ÌŒ " & xdate2.Text
If publicFlag = 1 Then
   If lCust Then
        cString = "select file5_20.*  " & _
                  " from file5_20 WHERE CLOSED = '0' and File5_20.CODE1 = " & MyParn(xCode.Text)
                    
        If IsDate(xDate1.Text) Then
            cString = cString & turnFound(cString) & " date_1 >= " & DateSql(xDate1.Text)
        End If
        
        If IsDate(xDate1.Text) Then
            cString = cString & turnFound(cString) & " date_1 <= " & DateSql(xdate2.Text)
        End If
    Else
        cString = "select file5_21.*  " & _
                  " from file5_21 WHERE CLOSED = '0' and File5_21.CODE1 = " & MyParn(xCode.Text)
        
        If IsDate(xDate1.Text) Then
            cString = cString & turnFound(cString) & " date_1 >= " & DateSql(xDate1.Text)
        End If
        
        If IsDate(xDate1.Text) Then
            cString = cString & turnFound(cString) & " date_1 <= " & DateSql(xdate2.Text)
        End If
    End If
Else
   If lCust Then
        cString = "select file5_20.*  " & _
                  " from file5_20 WHERE CLOSED = '1' and File5_20.CODE1 = " & MyParn(xCode.Text)
                    
        If IsDate(xDate1.Text) Then
            cString = cString & turnFound(cString) & " date_1 >= " & DateSql(xDate1.Text)
        End If
        
        If IsDate(xDate1.Text) Then
            cString = cString & turnFound(cString) & " date_1 <= " & DateSql(xdate2.Text)
        End If
    Else
        cString = "select file5_21.*  " & _
                  " from file5_21 WHERE CLOSED = '1' and File5_21.CODE1 = " & MyParn(xCode.Text)
        
        If IsDate(xDate1.Text) Then
            cString = cString & turnFound(cString) & " date_1 >= " & DateSql(xDate1.Text)
        End If
        
        If IsDate(xDate1.Text) Then
            cString = cString & turnFound(cString) & " date_1 <= " & DateSql(xdate2.Text)
        End If
    End If
End If
sourcetable.Open cString, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
With sourcetable
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If

If lCust Then
    If publicFlag = 1 Then
        cHead = " √Ê—«ﬁ ﬁ»÷ „” Õﬁ…  ··⁄„Ì·  " & xname.Text
    Else
        cHead = " √Ê—«ﬁ ﬁ»÷ „— œ… ··⁄„Ì·  " & xname.Text
    End If
Else
    If publicFlag = 1 Then
        cHead = " √Ê—«ﬁ œ›⁄ „” Õﬁ… ··„Ê—œ  " & xname.Text
    Else
        cHead = " √Ê—«ﬁ œ›⁄ „— œ… ··„Ê—œ  " & xname.Text
    End If
End If
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = !CHK_ID
    temptable!val1 = !Value
    If publicFlag = 1 Then
        temptable!Date1 = !date_1
        temptable!date2 = !date_R
    Else
        temptable!Date1 = !date_3
        temptable!date2 = !date_1
    End If
    temptable!str7 = cHead
    temptable!str8 = " „‰  «—ÌŒ " & xDate1.Text & " ≈·Ï  «—ÌŒ " & xdate2.Text
    'temptable!str19 = FirstTitle
    'temptable!str20 = SecondTitle
    temptable.Update
    .MoveNext
Loop

contemp.BeginTrans
contemp.CommitTrans
Report1.ReportFileName = App.Path & "\Reports\chq4.rpt"
Report1.DataFiles(0) = "c:\tempMrshd\temp.mdb"
Report1.Action = 1
End With
End Sub
Function myDateValue(pDate1)
myDateValue = "DateValue('" & pDate1 & "')"
End Function
Private Sub xCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xCode.Text = ""
    Dim Generalarray(3)
    Dim GrdArray(2)
    If lCust Then
        Set Generalarray(1) = Me
        Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As ≈”„ From File3_10"
        Generalarray(3) = "Where DescA Like '%cFilter%'"
            
        GrdArray(1) = 1000
        GrdArray(2) = 2600
               
        Lookupdata = Array(Generalarray, GrdArray)
        Load Search
        Search.Caption = "«” ⁄·«„ "
        Search.Show 1
    Else
        Set Generalarray(1) = Me
        Generalarray(2) = "Select Code As «·ﬂÊœ,DescA As ≈”„ From File4_10"
        Generalarray(3) = "Where DescA Like '%cFilter%'"
            
        GrdArray(1) = 1000
        GrdArray(2) = 2600
               
        Lookupdata = Array(Generalarray, GrdArray)
        Load Search
        Search.Caption = "«” ⁄·«„ "
        Search.Show 1
    End If
End If
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search.Grid1, 0)
Unload Search
End Sub
Private Sub xCode_LostFocus()
xname.Text = ""
If xCode.Text = "" Then Exit Sub
ClientTable.Find "code = " & xCode.Text, , adSearchForward, adBookmarkFirst
xname.Text = ClientTable!Desca
End Sub
