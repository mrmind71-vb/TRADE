VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpitem4 
   Caption         =   "ÿ»«⁄… "
   ClientHeight    =   3870
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
   ScaleHeight     =   3870
   ScaleWidth      =   6495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1230
      Left            =   3735
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2610
      Width           =   2610
      Begin VB.CheckBox xbal 
         Alignment       =   1  'Right Justify
         Caption         =   "«’‰«› ·Â« —’Ìœ ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   180
         Value           =   1  'Checked
         Width           =   1920
      End
      Begin VB.CheckBox xnobal 
         Alignment       =   1  'Right Justify
         Caption         =   "«—’œ… ”«·» ›ﬁÿ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   495
         Width           =   1920
      End
      Begin VB.CheckBox xcost 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ «·—’Ìœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   810
         Value           =   1  'Checked
         Width           =   1920
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
   Begin VB.Frame Frame1 
      Height          =   2580
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   6180
      Begin VB.TextBox XMOSM2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   855
         MaxLength       =   1
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   540
         Width           =   510
      End
      Begin VB.CheckBox xshowcost 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷  ﬁÌ„ «·—’œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1665
         Width           =   1920
      End
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
         Left            =   1530
         TabIndex        =   11
         Top             =   540
         Width           =   2760
         _ExtentX        =   4868
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
      Begin MSDataListLib.DataCombo xgroup 
         Height          =   315
         Left            =   855
         TabIndex        =   18
         Top             =   2025
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ã„Ê⁄… :"
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
         Index           =   2
         Left            =   4410
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2115
         Width           =   1230
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
      Top             =   2715
      Width           =   1005
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
      Left            =   2610
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2700
      Width           =   1050
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
      Top             =   2715
      Width           =   1185
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   0
      Top             =   1845
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
      Top             =   2835
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   0
      Top             =   0
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
   Begin MSAdodcLib.Adodc Adodc1 
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
Attribute VB_Name = "rpitem4"
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
    cString = "SELECT Sum((FILE1_11.[IN])-(FILE1_11.[out] )) AS Balance,Sum((FILE1_11.[IN]-FILE1_11.[out] ) * FILE1_10.COST) AS BalanceCOST ,file1_10.cost , MODELFACT,FILE1_10.ITEM , FILE1_10.FACT , FILE1_10.MODELNO , FILE1_10.MODEL , FILE1_10.PRICE , FILE1_10.COST , FILE1_10.DESCA , FILE1_10.SCAL , FILE1_10.COLOR , FILE1_10.C_SCAL , FILE1_10.C_COLOR , FILE1_10.MOSM , FACT.DESCA AS FACTDESCA , FILE1_10.CODE  " & _
        " FROM (FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM) LEFT JOIN FACT ON FILE1_10.[FACT] = FACT.CODE  WHERE FILE1_10.ITEM IS NOT NULL "
If IsDate(xDate.text) Then
    cString = cString & " AND date <= " & DateSq(xDate.text)
    aHeader(0) = "[" & "Õ Ì : " & xDate.text & "]"
End If

If Trim(xMosm.BoundText) <> "" Then
    cString = cString & " AND File1_10.[MOSM] = " & MyParn(xMosm.BoundText)
    aHeader(1) = "„Ê”„ " & xMosm.text & "]"
End If

If Trim(XMOSM2.text) <> "" Then
    cString = cString & " AND SUBSTRING(File1_10.[MOSM],1,1)  = " & MyParn(XMOSM2.text)
    aHeader(1) = "„Ê”„ " & XMOSM2.text & "]"
End If

If Trim(xFact.BoundText) <> "" Then
    cString = cString & " and File1_10.[FACT] = " & MyParn(xFact.BoundText)
    aHeader(2) = "„’‰⁄ " & xFact.text & "]"
End If

If Trim(xGroup.BoundText) <> "" Then
    cString = cString & " and File1_10.[group] = " & MyParn(xGroup.BoundText)
    aHeader(5) = "„’‰⁄ " & xFact.text & "]"
End If

If Trim(XSTORE.BoundText) <> "" Then
    cString = cString & " AND File1_11.store = " & MyParn(XSTORE.BoundText)
    aHeader(3) = "[" & "«·„Œ“‰ " & XSTORE.text & "]"
End If

If Trim(xCode.BoundText) <> "" Then
    cString = cString & " AND File1_10.CODE = " & MyParn(xCode.BoundText)
    aHeader(4) = "[" & "«·„Ê—œ " & xCode.text & "]"
End If
cString = cString & " GROUP BY file1_10.cost ,MODELFACT,FILE1_10.ITEM , FILE1_10.FACT , FILE1_10.MODELNO , FILE1_10.MODEL , FILE1_10.PRICE , FILE1_10.COST , FILE1_10.DESCA , FILE1_10.SCAL , FILE1_10.COLOR , FILE1_10.C_SCAL , FILE1_10.C_COLOR , FILE1_10.MOSM , FACT.DESCA , FILE1_10.CODE"
If xnobal.Value <> 0 Then
    cString = cString & " having Sum((FILE1_11.[IN])-(FILE1_11.[out] )) < 0 "
Else
    If xBal.Value <> 0 Then cString = cString & " having Sum((FILE1_11.[IN])-(FILE1_11.[out] )) <> 0 "
End If
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!str5 = ![Item]
        temptable!str11 = !MOSM
        temptable!str15 = !FACTDESCA
        temptable!str16 = TurnValue(!CODE, "", Null)
        temptable!str17 = !MODEL
        temptable!str18 = !Fact
        
        temptable!str1 = !modelfact
        temptable!str11 = !Fact & !modelfact
        temptable!str2 = !DESCA
        temptable!str3 = !Color
        temptable!str4 = TurnValue(DelZero(!scal), "", Null)
        temptable!val11 = !c_Color
        temptable!val12 = !C_SCAL
        
        If XCOST.Value <> 0 Then temptable!VAL4 = !balance
        temptable!VAL6 = !price
        If cBranch = "00" Then
            temptable!VAL7 = !cost
            temptable!Val8 = !balance * !cost
        End If
        temptable!STR7 = " ›’Ì·Ï √—’œ… «·„ÊœÌ·«  (  ›’Ì·Ï „ﬁ«”«   Ê √·Ê«‰ )·„Ê”„  " & xMosm.text & " ··„Ê—œ " & xCode.text
        temptable!str8 = TurnValue(retHeader(aHeader, 0, 5))
        temptable.Update
      .MoveNext
    Loop
End With
    If xshowcost.Value = 0 Then
        If xnobal.Value = 0 Then
            Main.REPORT1.ReportFileName = App.Path & "\Reports\rp_item4_0.rpt"
        Else
            Main.REPORT1.ReportFileName = App.Path & "\Reports\rp_item4_0.rpt"
        End If
    Else
        Main.REPORT1.ReportFileName = App.Path & "\Reports\rp_item4cost.rpt"
    End If
    contemp.BeginTrans
    contemp.CommitTrans

    Main.REPORT1.DataFiles(0) = tempFile
    Main.REPORT1.Action = 1

    temptable.Close
    sourcetable.Close
    Set temptable = Nothing
    Set sourcetable = Nothing
End Sub
Private Sub cmdClear_Click()
xGroup.BoundText = ""
XSTORE.BoundText = ""
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con
xshowcost.Visible = (cBranch = "00")
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

If cBranch <> "00" Then
    XSTORE.BoundText = cBranchStore
    XSTORE.Enabled = False
End If

data5.ConnectionString = strCon
data5.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA "
Set xGroup.RowSource = data5
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate.text) And Trim(xDate.text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If
'If xCode.BoundText = "" Then
'    MsgBox " ÕœÌœ «·„Ê—œ "
'    Exit Function
'End If
'If xMosm.BoundText = "" Then
'    MsgBox " ÕœÌœ «·„Ê”„ "
'    Exit Function
'End If
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
