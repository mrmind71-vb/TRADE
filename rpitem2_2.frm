VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpitem2_2 
   Caption         =   "ÿ»«⁄… "
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6495
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
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   6180
      Begin VB.TextBox xDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1620
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   1260
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xCode 
         Height          =   315
         Left            =   855
         TabIndex        =   0
         Top             =   900
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   855
         TabIndex        =   11
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   855
         TabIndex        =   13
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·„’‰⁄ :"
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
         Index           =   3
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   270
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "«·„Ê”„ :"
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   630
         Width           =   1680
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Õ Ï  «—ÌŒ"
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1755
         Width           =   795
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ê—œ  :"
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
         Index           =   0
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label4 
         Caption         =   "„Œ“‰ :"
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
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1350
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   " ›—Ì€"
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
      Left            =   1470
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1140
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "«” Ã«»…"
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
      Left            =   2625
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2400
      Width           =   1185
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
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2400
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   3825
      Top             =   2430
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
      Left            =   180
      Top             =   630
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
      Left            =   2655
      Top             =   2520
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
   Begin VB.Label Label6 
      Height          =   255
      Left            =   4275
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2175
      Width           =   1005
   End
End
Attribute VB_Name = "rpitem2_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdApply_Click()
Dim aHeader(5)
If Not MYVALID Then Exit Sub
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

    cField1 = myiif("type = '2' OR type = '7'   ", "([IN] - [OUT]) ") & " AS Q_IN , "
    cField2 = myiif("type = '2' OR type = '7'   ", "([IN] - [OUT]) * FILE1_11.COST ") & " AS T_IN   , "
    
    cField3 = myiif("type = 'F' OR type = 'T'   ", "([IN] - [OUT]) ") & " AS Q_TRA , "
    cField4 = myiif("type = 'F' OR type = 'T'   ", "([IN] - [OUT]) * FILE1_11.COST ") & " AS T_TRA , "
    
    cField5 = myiif("type = '6'  ", "([OUT]) ") & " AS Q_SAL , "
    cField6 = myiif("type = '6'  ", "([OUT]* FILE1_10.COST ) ") & " AS T_SAL "
    cString = "SELECT Sum((FILE1_11.[IN])-(FILE1_11.[out] )) AS Balance,Sum((FILE1_11.[IN]-FILE1_11.[out] ) * FILE1_10.COST) AS BalanceCOST ,MODELFACT,FILE1_10.FACT , FILE1_10.MODELNO , FILE1_10.MODEL , FILE1_10.DESCA , FILE1_10.MOSM , FACT.DESCA AS FACTDESCA , FILE1_10.CODE , " & _
                cField1 & cField2 & cField3 & cField4 & cField5 & cField6 & _
          " FROM (FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT JOIN FACT ON FILE1_10.[FACT] = FACT.CODE  WHERE FILE1_10.ITEM IS NOT NULL "

If IsDate(xDate.Text) Then
    cString = cString & " AND  date <= " & DateSq(xDate.Text)
    aHeader(0) = "[" & "Õ Ì : " & xDate.Text & "]"
End If

If Trim(xMosm.BoundText) <> "" Then
    cString = cString & " AND File1_10.[MOSM] = " & MyParn(xMosm.BoundText)
    aHeader(1) = "„Ê”„ " & xMosm.Text & "]"
End If

If Trim(xFact.BoundText) <> "" Then
    cString = cString & " AND File1_00.[FACT] = " & MyParn(xFact.Text)
    aHeader(2) = "„’‰⁄ " & xFact.Text & "]"
End If


If Trim(XSTORE.BoundText) <> "" Then
    cString = cString & " AND File1_11.store = " & MyParn(XSTORE.BoundText)
    aHeader(3) = "[" & "«·„Œ“‰ " & XSTORE.Text & "]"
End If

If Trim(xCode.BoundText) <> "" Then
    cString = cString & " AND File1_10.CODE = " & MyParn(xCode.BoundText)
    aHeader(4) = "[" & "«·„Ê—œ " & xCode.Text & "]"
End If

cString = cString & " GROUP BY MODELFACT,FILE1_10.FACT , FILE1_10.MODELNO , FILE1_10.MODEL , FILE1_10.COST , FILE1_10.DESCA , FILE1_10.MOSM , FACT.DESCA , FILE1_10.CODE "
         
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!str11 = !MOSM
        temptable!str15 = !FACTDESCA
        
        temptable!str1 = DelZero(!MODELFACT)
        temptable!str11 = !Fact & !MODELFACT
        temptable!str2 = !DESCA
        temptable!val1 = !Q_IN
        temptable!VAL7 = !T_IN
        
        temptable!val12 = !Q_TRA
        temptable!VAL13 = !T_TRA
        
        
        temptable!val2 = !q_sal
        temptable!val3 = !T_SAL
        temptable!VAL4 = !balance
        If cBranch = "00" Then
            temptable!val5 = !BALANCECOST
        End If
        temptable!str7 = " ›’Ì·Ï „‘ —Ì«  & „»Ì⁄«  & —’Ìœ ·„Ê”„  " & xMosm.Text & " ··„Ê—œ " & xCode.Text
        temptable!str8 = TurnValue(retHeader(aHeader, 0, 5))
        temptable.Update
      .MoveNext
    Loop
End With
    main.Report1.ReportFileName = App.Path & "\Reports\R_BalModel_C.rpt"
    contemp.BeginTrans
    contemp.CommitTrans

    main.Report1.DataFiles(0) = tempFile
    main.Report1.Action = 1

    temptable.Close
    sourcetable.Close
    Set temptable = Nothing
    Set sourcetable = Nothing
End Sub
Private Sub CmdClear_Click()
xGroup.BoundText = ""
XSTORE.BoundText = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "Select Code,DescA From FACT order by Desca"
Set xFact.RowSource = data1
xFact.ListField = "Desca"
xFact.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select Code,DescA From FILE4_10 ORDER by Desca"
Set xCode.RowSource = DATA2
xCode.ListField = "Desca"
xCode.BoundColumn = "Code"

data3.ConnectionString = strCon
data3.RecordSource = "Select MOSM,DescA From MOSM     ORDER by DATE "
Set xMosm.RowSource = data3
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"
xMosm.BoundText = cPMosm
data4.ConnectionString = strCon
data4.RecordSource = "Select Code,DescA From File0_40"
Set XSTORE.RowSource = data4
XSTORE.ListField = "Desca"
XSTORE.BoundColumn = "Code"

End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate.Text) And Trim(xDate.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If
If xCode.BoundText = "" Then
    MsgBox " ÕœÌœ «·„Ê—œ "
    Exit Function
End If
If xMosm.BoundText = "" Then
    MsgBox " ÕœÌœ «·„Ê”„ "
    Exit Function
End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xGroupMain_Validate(Cancel As Boolean)
If Not xGroupMain.MatchedWithList Then xGroupMain.BoundText = ""
data3.RecordSource = "Select Code,DescA From File1_50 " & IIf(xGroupMain.BoundText <> "", " where file1_50.[GROUP] = " & xGroupMain.BoundText, "")
data3.Refresh
End Sub
