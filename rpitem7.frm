VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpitem7 
   Caption         =   "≈Ã„«·Ì „‘ —Ì«  Œ·«· › —…"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6015
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
   ScaleHeight     =   2265
   ScaleWidth      =   6015
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
      Height          =   1710
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   5685
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1260
         Width           =   1680
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   855
         TabIndex        =   0
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
         Index           =   2
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1350
         Width           =   465
      End
      Begin VB.Label Label2 
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
         Index           =   1
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   990
         Width           =   765
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄…:"
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
         Top             =   225
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
         Top             =   540
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
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1710
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
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1710
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
      Left            =   2160
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1710
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   0
      Top             =   1800
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
      Left            =   45
      Top             =   810
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
      Top             =   1590
      Width           =   1005
   End
End
Attribute VB_Name = "rpitem7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
Dim aHeader(5)
If Not MYVALID Then Exit Sub
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT Sum(Val(Flag & '') *  Val(FILE7_20.Quant & '') ) as SumOfQuant, Sum(Val(Flag & '') * Val(PURCHASENET.QUANT & '')* val(PURCHASENET.PRICE & '')) as SumOfCost ,FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.PRICE, FILE1_50.DESCA " & _
          "FROM (FILE1_10 INNER JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) INNER JOIN PURCHASENET ON FILE1_10.ITEM = PURCHASENET.ITEM"

If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & " date >= " & DateSq(xdate1.Text)
    aHeader(0) = BetweenString(xdate1.Text, xDate2.Text)
End If

If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & " date <= " & DateSq(xDate2.Text)
    aHeader(0) = BetweenString(xdate1.Text, xDate2.Text)
End If

If Trim(xGroup.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "File1_10.GROUP = " & MyParn(xGroup.BoundText)
    aHeader(1) = "„Ã„Ê⁄… " & xGroup.Text & "]"
End If

If Trim(xStore.BoundText) <> "" Then
    cString = cString & turnFound(cString) & "PURCHASENET.store = " & MyParn(xStore.BoundText)
    aHeader(4) = "[" & "«·„Œ“‰ " & xStore.Text & "]"
End If

cString = cString & " GROUP BY FILE1_10.ITEM,FILE1_10.DESCA, FILE1_10.GROUP, FILE1_50.DESCA,FILE1_10.PRICE"
          
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, CON, adOpenForwardOnly, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!str5 = !Group
        temptable!STR6 = ![file1_50.desca]
        temptable!str1 = !Item
        temptable!str2 = ![FILE1_10.DESCA]
        temptable!val1 = !sumOfQuant
        temptable!val2 = !SumOfCOST
        temptable!val3 = !sumOfQuant * !price
        temptable!val4 = Val(!SumOfCOST & "") - Val(!sumOfQuant & "") * Val(!price & "")
        temptable!str7 = TurnValue(retHeader(aHeader, 0, 6))
        temptable.Update
        .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    REPORT1.ReportFileName = App.Path & "\Reports\Item5.rpt"
    contemp.BeginTrans
    contemp.CommitTrans
    
    REPORT1.DataFiles(0) = "c:\elmorshed\temp.mdb"
    REPORT1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub CmdClear_Click()
xGroup.BoundText = ""
xStore.BoundText = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
Data1.ConnectionString = CON.ConnectionString
Data1.RecordSource = "Select Code,DescA From File0_40"
Set xStore.RowSource = Data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

DATA2.ConnectionString = CON.ConnectionString
DATA2.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
Set xGroup.RowSource = DATA2
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xdate1.Text) And Trim(xdate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If
MYVALID = True
End Function

