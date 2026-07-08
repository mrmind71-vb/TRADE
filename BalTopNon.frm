VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form BalTopNon 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«—’œ… « Ê«» »«·”«·»"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin Threed.SSCommand cmd_print2 
      Height          =   465
      Left            =   135
      TabIndex        =   19
      Top             =   45
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   820
      _Version        =   196610
      ForeColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ÿ»«⁄…  ›’Ì·Ï «—’œ… «·Œ«„«  »«· Ê»"
      ButtonStyle     =   1
      PictureAlignment=   9
      BevelWidth      =   10
      ShapeSize       =   1
   End
   Begin VB.Frame Frame4 
      Height          =   1365
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   540
      Width           =   3795
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2475
         Picture         =   "BalTopNon.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "BalTopNon.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "BalTopNon.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   13
         Top             =   630
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   820
         _Version        =   196610
         BackColor       =   13822956
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "BalTopNon.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   45
         TabIndex        =   18
         Top             =   1125
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1890
      Left            =   3900
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11235
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   6000
         TabIndex        =   3
         Top             =   1320
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XCODE 
         Height          =   315
         Left            =   1635
         TabIndex        =   4
         Top             =   960
         Width           =   7800
         _ExtentX        =   13758
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XGROUP 
         Height          =   315
         Left            =   6000
         TabIndex        =   5
         Top             =   600
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XGROUP2 
         Height          =   315
         Left            =   1620
         TabIndex        =   14
         Top             =   1395
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XMOSM 
         Height          =   315
         Left            =   1620
         TabIndex        =   16
         Top             =   540
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„ "
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
         Left            =   4245
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   540
         Width           =   660
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„’‰⁄ «·Œ«„…"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label2 
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
         Height          =   240
         Index           =   1
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1410
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "«·Œ«„… : "
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "‰Ê⁄ «·Œ«„… : "
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
         Left            =   9555
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   615
         Width           =   1050
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   15240
      _ExtentX        =   26882
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -135
      Top             =   300
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "BalTopNon.frx":9583
      Height          =   7335
      Left            =   90
      TabIndex        =   2
      Top             =   1935
      Width           =   15045
      _cx             =   26538
      _cy             =   12938
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   0
      Top             =   450
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
      Top             =   600
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
   Begin MSAdodcLib.Adodc DATA5 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2475
      _ExtentX        =   4366
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
   Begin MSAdodcLib.Adodc DATA6 
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
Attribute VB_Name = "BalTopNon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim osearchDoc As New Search3, oSearchitem As New Search3, osearchClient As New Search3
Dim con As New ADODB.Connection
Dim cString As String
Dim cStr1 As String, cStr2 As String
Private Sub cmd_excel_Click()
        ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub Cmd_Print_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = Me.Caption & "  " & XCODE.Text
    cHead2 = " „Ã„Ê⁄… " & XGROUP.Text & " „Œ“‰ " & xStore.Text
    Set printGrd3.grid1 = grid1
    Set printGrd3.myform = Me
    printGrd3.doprint 1.2, 0, cHead1, cHead2, "", False, True, 11
    printGrd3.Show 1
End Sub
Private Sub cmd_print2_Click()
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim SubTable As New ADODB.Recordset
Dim nCountSub As Double, nCount As Double, nItemCount As Double

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT Q_ITEM.ITEM, Q_ITEM.[DESCA], SUM(FILE1_11.[IN]-FILE1_11.[OUT] ) AS BAL , Q_ITEM.FACTNAME ,Q_ITEM.MOSM ,Q_ITEM.COLOR , SUM(IN_N - OUT_N ) AS BAL2 ,  Q_ITEM.unit, Q_ITEM.ITEM_NO , Q_ITEM.SUB_COLOR , Q_ITEM.[GROUP] " & _
          " FROM Q_ITEM inner JOIN FILE1_11 ON FILE1_11.ITEM = Q_ITEM.ITEM WHERE Q_ITEM.ITEM IN (SELECT ITEM  FROM Q_BAL_TOB WHERE BAL < 0 )"

If XGROUP.BoundText <> "" Then cString = cString & " AND [Q_ITEM].[M_GROUP] = " & MyParn(XGROUP.BoundText)
If XCODE.BoundText <> "" Then cString = cString & " AND [Q_ITEM].[GROUP] = " & MyParn(XCODE.BoundText)
If XGROUP2.BoundText <> "" Then cString = cString & " AND [FACT] = " & MyParn(XGROUP2.BoundText)
If XMOSM.BoundText <> "" Then cString = cString & " AND MOSM = " & MyParn(XMOSM.BoundText)
cString = cString & " GROUP BY Q_ITEM.ITEM, Q_ITEM.COLOR , Q_ITEM.[DESCA], Q_ITEM.FACTNAME ,Q_ITEM.MOSM,  Q_ITEM.unit, Q_ITEM.ITEM_NO , Q_ITEM.SUB_COLOR , Q_ITEM.[GROUP]"
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do While Not sourcetable.EOF
    If SubTable.State = adStateOpen Then SubTable.Close
    SubTable.Open " SELECT * FROM Q_BAL_TOB WHERE ITEM = " & MyParn(sourcetable!Item) & " AND BAL < 0 ORDER BY ID ", con, adOpenStatic, adLockReadOnly, adCmdText
    nCountSub = Val(GetDesca("SELECT COUNT(*) FROM Q_BAL_TOB WHERE BAL < 0 AND ITEM = " & MyParn(sourcetable!Item)) & "")
    If nCountSub > 0 Then
    nCount = Round(nCountSub, 0) / 5
    nItemCount = 0
    If nCount <> Int(nCount) Then nCount = Int(nCount) + 1
    SubTable.MoveFirst
    For I = 1 To nCount
        temptable.AddNew
        temptable!STR21 = " ›’·Ï √—’œ… « Ê«» »«·”«·» "
        temptable!str1 = TurnValue(XGROUP.Text)
        temptable!str2 = TurnValue(XMOSM.Text)
        temptable!str3 = TurnValue(XGROUP2.Text)
        temptable!STR11 = sourcetable!Item
        temptable!STR7 = sourcetable!FACTNAME
        temptable!str12 = sourcetable!DESCA
        temptable!str13 = sourcetable!Color
        temptable!str14 = sourcetable!MOSM
        temptable!STR6 = sourcetable!UNIT
        temptable!str8 = sourcetable!Group
        temptable!str4 = sourcetable!ITEM_NO
        temptable!STR5 = sourcetable!SUB_COLOR
        
        If I = 1 Then
            temptable!val9 = sourcetable!BAL2
            temptable!VAL1 = sourcetable!BAL
        End If
  
        If nItemCount < nCountSub Then
            temptable!VAL12 = Val(SubTable!BAL & "")
            temptable!str15 = SubTable!ID
        End If
        nItemCount = nItemCount + 1
        
        If nItemCount < nCountSub Then
            SubTable.MoveNext
            temptable!VAL13 = Val(SubTable!BAL & "")
            temptable!str16 = SubTable!ID
            nItemCount = nItemCount + 1
        End If
        
        If nItemCount < nCountSub Then
            SubTable.MoveNext
            temptable!VAL14 = Val(SubTable!BAL & "")
            temptable!str17 = SubTable!ID
            nItemCount = nItemCount + 1
        End If
        
        If nItemCount < nCountSub Then
            SubTable.MoveNext
            temptable!val15 = Val(SubTable!BAL & "")
            temptable!str18 = SubTable!ID
            nItemCount = nItemCount + 1
        End If
        
        If nItemCount < nCountSub Then
            SubTable.MoveNext
            temptable!Val16 = Val(SubTable!BAL & "")
            temptable!STR19 = SubTable!ID
            nItemCount = nItemCount + 1
        End If
        temptable.Update
        If nItemCount < nCountSub Then
            SubTable.MoveNext
        End If
        If nItemCount > nCountSub Then Exit For
    Next I
    End If
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\BALITEM_SUB2.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing

End Sub

Private Sub cmdExit_Click()
    Unload Me
    Set TSalItem = Nothing
End Sub
Private Sub cmdundo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    MYLOAD
End Sub
Private Sub Form_Load()
    openCon con
    grid1.FixedRows = 1
    grid1.Rows = 2
    Set grid1.DataSource = data4
    data4.ConnectionString = strCon
    
    
    data1.ConnectionString = strCon
    data1.RecordSource = "Select CODE ,DescA From FILE1_50G "
    Set XGROUP.RowSource = data1
    XGROUP.ListField = "DESCA"
    XGROUP.BoundColumn = "CODE"
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select CODE ,Desc2 From FILE1_50 WHERE ISRAW = 1 ORDER BY DESC2"
    Set XCODE.RowSource = data3
    XCODE.ListField = "Desc2"
    XCODE.BoundColumn = "Code"
    
    DATA5.ConnectionString = strCon
    DATA5.RecordSource = "SELECT * FROM file1_53 "
    Set XGROUP2.RowSource = DATA5
    XGROUP2.ListField = "Desca"
    XGROUP2.BoundColumn = "Code"
    
    DATA6.ConnectionString = strCon
    DATA6.RecordSource = "SELECT * FROM MOSM     "
    Set XMOSM.RowSource = DATA6
    XMOSM.ListField = "DESCA"
    XMOSM.BoundColumn = "MOSM"
    
      
    FixGrid
End Sub
Private Sub MYLOAD()
With grid1
    CSTRALL = "  SELECT GROUPNAME ,Q_ITEM.[GROUP], Q_ITEM.DESCA  , FACTNAME , MOSM , ITEM_NO , SUB_COLOR , COLOR , UNIT ,Q_ITEM.ITEM , Q_FILE1_11_SUB.ID  , SUM(Q_IN-Q_OUT) " & _
              " FROM Q_ITEM INNER JOIN Q_FILE1_11_SUB ON Q_ITEM.ITEM = Q_FILE1_11_SUB.ITEM  WHERE Q_ITEM.ITEM IS NOT NULL "
    If XGROUP.BoundText <> "" Then CSTRALL = CSTRALL & " AND [Q_ITEM].[M_GROUP] = " & MyParn(XGROUP.BoundText)
    If XCODE.BoundText <> "" Then CSTRALL = CSTRALL & " AND [Q_ITEM].[GROUP] = " & MyParn(XCODE.BoundText)
    If XGROUP2.BoundText <> "" Then CSTRALL = CSTRALL & " AND [FACT] = " & MyParn(XGROUP2.BoundText)
    If XMOSM.BoundText <> "" Then CSTRALL = CSTRALL & " AND MOSM = " & MyParn(XMOSM.BoundText)
    CSTRALL = CSTRALL & " GROUP BY GROUPNAME ,Q_ITEM.[GROUP] ,Q_ITEM.COST,Q_ITEM.[ITEM], Q_ITEM.DESCA , COLOR , FACTNAME , MOSM , ITEM_NO , SUB_COLOR , UNIT  , Q_FILE1_11_SUB.ID HAVING SUM(Q_IN-Q_OUT) < 0 "
    CSTRALL = CSTRALL & " ORDER BY GROUPNAME ,Q_ITEM.DESCA , FACTNAME , MOSM , ITEM_NO , SUB_COLOR , COLOR , UNIT "
    data4.RecordSource = CSTRALL
    data4.Refresh
End With
FixGrid
End Sub
Sub FixGrid()
    With grid1
    .Cols = 12
    .RowHeight(0) = 1000
    .WordWrap = True
    .FixedRows = 1
    .ColHidden(1) = True
    .TextMatrix(0, 0) = "‰Ê⁄"
    .TextMatrix(0, 2) = "«·Œ«„…"
    .TextMatrix(0, 3) = "«·„Ê—œ"
    .TextMatrix(0, 4) = "«·„Ê”„"
    .TextMatrix(0, 5) = "—ﬁ„"
    .TextMatrix(0, 6) = "«·—”„…"
    .TextMatrix(0, 7) = "«··Ê‰"
    .TextMatrix(0, 8) = "«·ÊÕœ…"
    .TextMatrix(0, 9) = "ITEM"
    .TextMatrix(0, 10) = "—ﬁ„  Ê»"
    .TextMatrix(0, 11) = "«·—’Ìœ"
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    
    .ColHidden(1) = True
    
    
    .ColWidth(0) = 1500
    .ColWidth(2) = 2000
    .ColWidth(3) = 800
    .ColWidth(4) = 800
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 2000
    .ColWidth(8) = 800
    .ColWidth(9) = 800
    .ColWidth(10) = 500
    .ColWidth(11) = 800
    
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .SubtotalPosition = flexSTBelow
    
    .MergeCells = flexMergeFree
    .MergeCol(0) = True
    .MergeCol(1) = True
    .MergeCol(2) = True
    .MergeCol(3) = True
    .MergeCol(4) = True
    .MergeCol(5) = True
    .MergeCol(6) = True
    .MergeCol(7) = True
    .MergeCol(8) = True
    .MergeCol(9) = True
    
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_DblClick()
    StoreMove_sub.xItem.Text = grid1.TextMatrix(grid1.Row, 9)
    StoreMove_sub.xtob_no.Text = grid1.TextMatrix(grid1.Row, 10)
    StoreMove_sub.Show
End Sub
Private Sub grid1_EnterCell()
With grid1
    If .Col = 14 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub XCODE_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me

Generalarray(1) = "SELECT FILE1_50.Code, File1_50.Desca , FILE1_53.DESCA , MOSM , ITEM_NO , SUB_COLOR " & _
                  " FROM FILE1_50 INNER JOIN FILE1_53 ON FILE1_53.CODE = FILE1_50.GROUP2 where israw = 1"
Generalarray(2) = " Order by FILE1_50.Code"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·’‰›"
listarray(0, 1) = "%%FILE1_50.DESC2%% "

GrdArray(0, 0) = "ﬂÊœ «·„Ã„Ê⁄…"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·„Ã„Ê⁄…"
GrdArray(1, 1) = 3000

GrdArray(2, 0) = "„Ê—œ"
GrdArray(2, 1) = 1000

GrdArray(3, 0) = "„Ê”„"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "—ﬁ„"
GrdArray(4, 1) = 2000

GrdArray(5, 0) = "—”„…"
GrdArray(5, 1) = 2000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "„Ã„Ê⁄«  «·Œ«„« "
Search3.Show 1

End If
End Sub
Sub myProc()
    If ActiveControl.Name = XCODE.Name Then
        XCODE.BoundText = (Search3.grid1.TextMatrix(Search3.grid1.Row, 0))
        Unload Search3
    End If
End Sub

