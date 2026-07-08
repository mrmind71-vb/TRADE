VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpModel3 
   Caption         =   "«—’œ… «·„ÊœÌ·«   "
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
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
   ScaleHeight     =   3150
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1920
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   135
      Width           =   6420
      Begin VB.TextBox xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3150
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1440
         Width           =   1860
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   270
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xfact 
         Height          =   315
         Left            =   1800
         TabIndex        =   9
         Top             =   660
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xcode 
         Height          =   315
         Left            =   1800
         TabIndex        =   12
         Top             =   1050
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "„Ê—œ :"
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
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1155
         Width           =   450
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ï  «—ÌŒ:"
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
         Left            =   5190
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1530
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "„’‰⁄ :"
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
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   716
         Width           =   690
      End
      Begin VB.Label Label5 
         Caption         =   "„Ê”„ :"
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
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   690
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3600
      Top             =   2160
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
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2025
      Width           =   3480
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   420
         Left            =   1170
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "⁄—÷"
         Height          =   420
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   135
         Width           =   1140
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2115
      Top             =   2700
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   3330
      Top             =   2745
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   900
      Top             =   2700
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
   Begin VB.Label Label2 
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
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1215
      Width           =   570
   End
End
Attribute VB_Name = "rpModel3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3

Dim oSearchDoc As New Search3, oSearchItem As New Search3, osearchClient As New Search3
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
    If MYVALID Then doprint
End Sub
Private Sub CmdUndo_Click()
    xMosm.BoundText = ""
    xFact.BoundText = ""
    xCode.BoundText = ""
'    xDate1.Text = ""
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Function MYVALID() As Boolean
    MYVALID = True
End Function
Private Sub Form_Load()
openCon con
data1.ConnectionString = strCon
data1.RecordSource = "Select MOSM ,DescA From MOSM "

Set xMosm.RowSource = data1
xMosm.ListField = "MOSM"
xMosm.BoundColumn = "MOSM"
xMosm.BoundText = cPMosm

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select CODE ,DescA From fact ORDER BY DESCA"
Set xFact.RowSource = DATA2
xFact.ListField = "Desca"
xFact.BoundColumn = "Code"

data3.ConnectionString = strCon
data3.RecordSource = "Select CODE ,DescA From FILE4_10 ORDER BY DESCA"
Set xCode.RowSource = data3
xCode.ListField = "Desca"
xCode.BoundColumn = "Code"

End Sub

Private Sub doprint()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset, nBalance As Single
Dim ModelsTable As New ADODB.Recordset
Dim ScalTable As New ADODB.Recordset
Dim ColorTable As New ADODB.Recordset

Dim aHeader(2)
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
cStr1 = " SELECT  FILE1_10.MODEL, file1_10.fact , file1_10.[GROUP] , file1_10.code , FILE1_10.desca, FILE1_10.mosm , SUM(FILE1_11.[IN] - FILE1_11.OUT ) AS BALMODEL  FROM  FILE1_10  INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM "
If xMosm.BoundText <> "" Then cStr1 = cStr1 & " WHERE FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
If xFact.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[fact] = " & MyParn(xFact.BoundText)
If xCode.BoundText <> "" Then cStr1 = cStr1 & " AND FILE1_10.[code] = " & MyParn(xCode.BoundText)
If IsDate(xdate.text) Then cStr1 = cStr1 & " AND date <= " & DateSq(xdate.text)
cStr1 = cStr1 & " GROUP BY  FILE1_10.MODEL, file1_10.fact , file1_10.[GROUP] , file1_10.code , FILE1_10.desca, FILE1_10.mosm ORDER BY FILE1_10.MODEL "

ModelsTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
Dim cModel As String
Dim TModelDem As Recordset
Dim nTotal As Double
Dim nCol As Double
With temptable
Do While Not ModelsTable.EOF
    If Val(ModelsTable!BALMODEL & "") <> 0 Then
        cModel = ModelsTable!MODEL
        cStr1 = " SELECT FILE1_10.MODEL , FILE1_10.C_SCAL, file1_10.item , FILE1_10.SCAL, SUM(FILE1_11.[IN] - FILE1_11.OUT ) AS T_BAL FROM  FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM  WHERE FILE1_10.MODEL = " & MyParn(cModel)
        If IsDate(xdate.text) Then cStr1 = cStr1 & " AND FILE1_11.[date] <= " & DateSq(xdate.text)
        cStr1 = cStr1 & " GROUP BY FILE1_10.MODEL , FILE1_10.C_SCAL, file1_10.item , FILE1_10.SCAL order by  FILE1_10.C_SCAL "
            
        If ScalTable.State = adStateOpen Then ScalTable.Close
        ScalTable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
            
        Caption = cModel
        .AddNew
        !str17 = cModel
        !str16 = "»«—ﬂÊœ"
        !str18 = ArbString(ModelsTable!DESCA)
        !STR19 = TurnValue(GetDesca("SELECT DESCAA FROM FILE1_50 WHERE CODE = " & MyParn(ModelsTable!Group), con), "", Null)
        !str15 = " ›’Ì·Ï —’Ìœ «·„ÊœÌ·« "
        nTotal = 0
        ScalTable.MoveFirst
        nCol = 0
        Do While Not ScalTable.EOF
            nCol = nCol + 1
            .Fields("STR" & nCol) = ScalTable!scal
            .Fields("VAL" & nCol) = ScalTable!Item
            ScalTable.MoveNext
        Loop
        
        .AddNew
        !str17 = cModel
        !str16 = "«·—’Ìœ"
        !str18 = ArbString(ModelsTable!DESCA)
        !str15 = " ›’Ì·Ï —’Ìœ «·„ÊœÌ·« "
        nTotal = 0
        nCol = 0
        ScalTable.MoveFirst
        Do While Not ScalTable.EOF
            nCol = nCol + 1
            .Fields("STR" & nCol) = ScalTable!scal
            .Fields("VAL" & nCol) = ScalTable!T_BAL
            nTotal = nTotal + Val(ScalTable!T_BAL & "")
            ScalTable.MoveNext
        Loop
        !VAL13 = nTotal
        .Update
    End If
    ModelsTable.MoveNext
Loop
contemp.BeginTrans
contemp.CommitTrans
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    main.Report1.ReportFileName = App.Path & "\Reports\R_TDem.rpt"
    main.Report1.DataFiles(0) = tempFile
    main.Report1.Action = 1
End If
End With
ModelsTable.Close

Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub DelTable()
On Error Resume Next
con.Execute "DROP TABLE BAL_Rep1 "
Err.Clear
End Sub
Private Sub xModel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then ModelLookupAll Me, Search3
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then SuppLookupAll Me, oSearch
End Sub
Sub myProc()
If ActiveControl.Name = xCode.Name Then
    xCode.BoundText = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
End If
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
End Sub
Private Sub xFACT_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Set Generalarray(0) = Me
Generalarray(1) = "Select code ,DescA From FACT"
Generalarray(2) = "Order by DESCA"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1

End If
End Sub
Private Sub xcode_Validate(Cancel As Boolean)
If xCode.BoundText = "" Then Exit Sub
If Not xCode.MatchedWithList Then
    cFact = GetDesca("SELECT CODE FROM FILE4_10 WHERE SUBCODE = " & MyParn(xCode.BoundText), con)
    xCode.BoundText = cFact
    If xCode.MatchedWithList Then Exit Sub
End If
End Sub

Private Sub xFact_Validate(Cancel As Boolean)
If Not xFact.MatchedWithList Then
    cFact = xFact.BoundText
    xFact.BoundText = cFact
    If xFact.MatchedWithList Then Exit Sub
End If

End Sub

