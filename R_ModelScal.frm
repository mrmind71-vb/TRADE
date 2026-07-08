VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form R_ModelScal 
   Caption         =   "ÿ»«⁄… "
   ClientHeight    =   3150
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   2325
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   90
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
         TabIndex        =   5
         Top             =   1845
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
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   1800
         TabIndex        =   7
         Top             =   1050
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   1800
         TabIndex        =   8
         Top             =   660
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "„Ê”„"
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
         TabIndex        =   12
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label2 
         Caption         =   "„Ã„Ê⁄…"
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
         TabIndex        =   11
         Top             =   716
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1155
         Width           =   540
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
         TabIndex        =   9
         Top             =   1935
         Width           =   840
      End
   End
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
      Left            =   1518
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2580
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
      Left            =   2676
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   2580
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
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2580
      Width           =   1320
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
   Begin MSAdodcLib.Adodc DATA2 
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   1170
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
      TabIndex        =   3
      Top             =   2175
      Width           =   1005
   End
End
Attribute VB_Name = "R_ModelScal"
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

cString = "SELECT Sum((FILE1_11.[IN])-(FILE1_11.[out] )) AS Balance,FILE1_10.MODEL, FILE1_10.SCAL ,FILE1_10.C_SCAL, FILE1_10.COLOR, FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP] AS FILE1_50GROUP,FILE1_10.COST, file1_10.price , FILE1_50.DESCA AS FILE1_50DESCA,FILE1_50G.DESCA AS FILE1_50GDESCA" & _
          " FROM ((FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_50G ON FILE1_50.[GROUP] = FILE1_50G.CODE WHERE FILE1_10.MODEL IS NOT NULL "

If IsDate(xdate.Text) Then
    cString = cString & turnFound(cString) & " date <= " & DateSq(xdate.Text)
    aHeader(0) = "[" & "Õ Ì : " & xdate.Text & "]"
End If

If Trim(xGroup.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File1_10.[group] = " & MyParn(xGroup.BoundText)
    aHeader(1) = "„Ã„Ê⁄… " & xGroup.Text & "]"
End If

If Trim(xMosm.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File1_10.[MOSM] = " & MyParn(xMosm.BoundText)
    aHeader(1) = "„Ê”„ " & xMosm.Text & "]"
End If

If Trim(xStore.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File1_11.store = " & MyParn(xStore.BoundText)
    aHeader(2) = "[" & "«·„Œ“‰ " & xStore.Text & "]"
End If

cString = cString & " GROUP BY FILE1_10.SCAL ,FILE1_10.C_SCAL, FILE1_10.COLOR,FILE1_10.MODEL, FILE1_10.ITEM,file1_10.price , FILE1_10.DESCA, FILE1_10.[GROUP],FILE1_50.[GROUP],FILE1_10.COST, FILE1_50.DESCA,FILE1_50G.DESCA"
Set sourcetable = New ADODB.Recordset
con.CommandTimeout = 200
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        If Round(Val(!balance & ""), 2) <> 0 Then
            temptable.AddNew
            temptable!STR11 = !FILE1_50GROUP
            temptable!STR7 = ![file1_50GDESCA]
            temptable!str12 = ![Group]
            temptable!str8 = ![file1_50desca]
            temptable!str1 = !MODEL
            temptable!str4 = !Color
            temptable!str5 = !SCAL
            temptable!VAL5 = !C_SCAL
            
            temptable!str2 = ![DESCA]
            temptable!val2 = !balance
            
            temptable!VAL3 = !cost
            
            temptable!STR21 = TurnValue(retHeader(aHeader, 0, 6))
            temptable.Update
        End If
      .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    main.Report1.ReportFileName = App.Path & "\Reports\Item_Model.rpt"
    contemp.BeginTrans
    contemp.CommitTrans
    
    main.Report1.DataFiles(0) = tempFile
    main.Report1.Action = 1
End If
temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub cmdClear_Click()
xGroup.BoundText = ""
xStore.BoundText = ""
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con

data1.ConnectionString = strCon
data1.RecordSource = "Select MOSM ,DescA From MOSM "

Set xMosm.RowSource = data1
xMosm.ListField = "MOSM"
xMosm.BoundColumn = "MOSM"
xMosm.BoundText = cPMosm

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select CODE ,DescA From FILE0_40 ORDER BY DESCA"
Set xStore.RowSource = DATA2
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

DATA3.ConnectionString = strCon
DATA3.RecordSource = "SELECT * FROM file1_50 order by desca  "
Set xGroup.RowSource = DATA3
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

End Sub
Private Function MYVALID() As Boolean
'If Not IsDate(xDate1.Text) And Trim(xDate1.Text) <> "" Then
'    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
'    Exit Function
'End If
MYVALID = True
End Function

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xGroupMain_Validate(Cancel As Boolean)
If Not xGroupMain.MatchedWithList Then xGroupMain.BoundText = ""
DATA3.RecordSource = "Select Code,DescA From File1_50 " & turn(xGroupMain.BoundText <> "", " where file1_50.[GROUP] = " & MyParn(xGroupMain.BoundText))
DATA3.Refresh
End Sub
