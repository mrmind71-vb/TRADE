VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpClient10 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   4995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5730
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
   ScaleHeight     =   4995
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1410
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   5460
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1680
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   945
         Width           =   1680
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì· :"
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
         TabIndex        =   12
         Top             =   315
         Width           =   600
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   2445
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
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì  «—ÌŒ :"
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
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1035
         Width           =   900
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   6615
      Top             =   2475
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame1 
      Height          =   3135
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1350
      Width           =   5505
      Begin VB.CheckBox xTable 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ ÃœÊ·Ì"
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2790
         Width           =   1680
      End
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   360
         TabIndex        =   13
         Top             =   2070
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grid1 
         Height          =   1440
         Left            =   0
         TabIndex        =   14
         Top             =   555
         Width           =   3765
         _cx             =   6641
         _cy             =   2540
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
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
         BackColorBkg    =   -2147483633
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
         Rows            =   10
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"rpclient10.frx":0000
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
         ShowComboButton =   -1  'True
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   315
         TabIndex        =   15
         Top             =   180
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XUNIT 
         Height          =   315
         Left            =   360
         TabIndex        =   16
         Top             =   2430
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "„Œ“‰"
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
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2115
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "„‰ „Ã„Ê⁄… "
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
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… —∆Ì”Ì…"
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
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   255
         Width           =   1230
      End
      Begin VB.Label Label3 
         Caption         =   "«·ÊÕœ…"
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
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2475
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
      Left            =   3285
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4545
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
      Left            =   4455
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   4545
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
      Left            =   1935
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4545
      Width           =   1320
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   45
      Top             =   1125
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   150
      Top             =   1725
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
      Left            =   -1620
      Top             =   -90
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
      TabIndex        =   7
      Top             =   2175
      Width           =   1005
   End
End
Attribute VB_Name = "rpClient10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
If xTable.Value <> 0 Then
    doprint2
Else
    doprint1
End If
End Sub

Private Sub CmdClear_Click()
xGroup.BoundText = ""
xstore.BoundText = ""
XUNIT.BoundText = ""
grid1.Rows = 0
grid1.Rows = 10
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
grdMake "Select Code,DescA From File1_50", "code", "desca", CON, grid1
data1.ConnectionString = CON.ConnectionString
data1.RecordSource = "Select Code,DescA From File1_51 ORDER BY CODE"
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"

DATA2.ConnectionString = CON.ConnectionString
DATA2.RecordSource = "Select Code,DescA From File0_40"
Set xstore.RowSource = DATA2
xstore.ListField = "Desca"
xstore.BoundColumn = "Code"

data3.ConnectionString = CON.ConnectionString
data3.RecordSource = "Select Code,DescA From File1_13 order by code"
Set XUNIT.RowSource = data3
XUNIT.ListField = "Desca"
XUNIT.BoundColumn = "Code"
End Sub
Private Sub xGroup_Click(Area As Integer)
If Area = 2 Then grdMake "Select * From File1_50" & IIf(xGroup.BoundText <> "", " where [GROUP] = " & MyParn(xGroup.BoundText), ""), "code", "desca", CON, grid1
End Sub
Private Function itemWidth(pItem) As String
itemWidth = retitem(pItem, "width1") & ""
If Not IsNull(retitem(pItem, "width2")) Then itemWidth = itemWidth & IIf(itemWidth = "", "", " x ") & retitem(pItem, "width2")
If Not IsNull(retitem(pItem, "length")) Then itemWidth = itemWidth & IIf(itemWidth = "", "", " x ") & retitem(pItem, "length")
End Function
Private Sub doprint1()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
Dim aHeader(4)
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
cString = "SELECT FILE1_10.ITEM,FILE1_10.WIDTH1,FILE1_10.DESCA AS ITEMDESCA, FILE1_10.PRICE AS ItemPRICE,FILE1_13.DESCA as UnitDesca," & _
          " FILE1_10.COST AS ItemCOST, FILE1_10.DESCA, SUM(VAL(ITEMSALES.QUANT & '')) AS SUMOFQUANT, FILE1_10.GROUP AS GroupCode, FILE1_50.DESCA AS GroupDesca,   FILE1_50.GROUP AS MainGroupCode, FILE1_51.DESCA as  MainGroupDesca FROM (((FILE1_10 INNER JOIN ITEMSALES ON FILE1_10.ITEM = ITEMSALES.ITEM) LEFT  JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_51 ON  FILE1_50.GROUP = FILE1_51.CODE)LEFT JOIN FILE1_13 ON FILE1_10.UNIT = FILE1_13.CODE "

If Trim(xCode.Text) <> "" Then
    aHeader(0) = "[" & " «·⁄„Ì· : " & xCodeDesca.Caption & "]"
    cString = cString & turnFound(cString) & "  ITEMSALES.CODE = " & MyParn(xCode.Text)
End If

If GrdQry(grid1, "file1_50.code", True) <> "" Then
    cString = cString & turnFound(cString) & GrdQry(grid1, "File1_50.code", True)
    aHeader(1) = "[" & "«·’‰› : " & GrdTitle(grid1) & "]"
ElseIf xGroup.BoundText <> "" Then
    aHeader(1) = "[" & " „Ã„Ê⁄… : " & xGroup.Text & "]"
    cString = cString & turnFound(cString) & "  File1_50.GROUP = " & MyParn(xGroup.BoundText)
End If

If xstore.BoundText <> "" Then
    cString = cString & turnFound(cString) & "File1_11.store = " & MyParn(xstore.BoundText)
    aHeader(2) = "[" & "„Œ“‰ : " & xstore.Text & "]"
End If
If XUNIT.BoundText <> "" Then
    cString = cString & turnFound(cString) & "File1_10.UNIT = " & MyParn(XUNIT.BoundText)
    aHeader(2) = "[" & "«·ÊÕœ… : " & XUNIT.Text & "]"
End If

If xdate1.Text <> "" Then
    cString = cString & turnFound(cString) & "ITEMSALES.date >= " & DateSql(xdate1.Text)
    aHeader(3) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If xDate2.Text <> "" Then
    cString = cString & turnFound(cString) & "ITEMSALES.date <= " & DateSql(xDate2.Text)
    aHeader(3) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If


cString = cString & " GROUP BY FILE1_10.ITEM,FILE1_10.WIDTH1,FILE1_10.DESCA,FILE1_10.PRICE,FILE1_13.DESCA, " & _
          " FILE1_10.COST, FILE1_10.DESCA,FILE1_10.GROUP,FILE1_50.DESCA,FILE1_50.GROUP, FILE1_51.DESCA"
                    
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, CON, adOpenForwardOnly, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
        temptable.AddNew
        temptable!str6 = !MainGroupDesca
        temptable!str5 = !MAINGROUPCODE
        temptable!str1 = !Item
        temptable!str2 = itemWidth(sourcetable!Item)
        temptable!str3 = !GroupCode
        temptable!str4 = !GroupDesca
        temptable!str8 = !unitDesca
        temptable!VAL1 = !sumOfQuant
        temptable!VAL20 = !width1
        temptable!STR20 = !width1
        temptable!str17 = TurnValue(retHeader(aHeader, 0, 5))
        temptable.Update
      .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    REPORT1.ReportFileName = App.Path & "\Reports\CLIENT10.rpt"
    contemp.BeginTrans
    contemp.CommitTrans
    REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    REPORT1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub doprint2()
Dim bMainGroupHeader As Boolean, bGroupHeader As Boolean
bGroupHeader = True: bMainGroupHeader = True
If Not MYVALID Then Exit Sub
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset, cHeader As String
Dim aHeader(4)
contemp.Execute "delete * from temp"
Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT FILE1_10.ITEM, FILE1_10.WIDTH1, FILE1_50.DESCA, SUM(VAL(ITEMSALES.QUANT & '')) AS SumofQuant, FILE1_10.WIDTH2, FILE1_10.GROUP AS GroupCode, FILE1_50.group AS MainGroupCode, FILE1_51.DESCA AS MainGroupDesca " & _
           " FROM ((FILE1_10 INNER JOIN ITEMSALES ON FILE1_10.ITEM = ITEMSALES.ITEM) LEFT JOIN FILE1_50 ON FILE1_10.GROUP = FILE1_50.CODE) LEFT JOIN FILE1_51 ON FILE1_50.group = FILE1_51.CODE  "

If Trim(xCode.Text) <> "" Then
    aHeader(0) = "[" & " «·⁄„Ì· : " & xCodeDesca.Caption & "]"
    cString = cString & turnFound(cString) & "  ITEMSALES.CODE = " & MyParn(xCode.Text)
End If

If GrdQry(grid1, "file1_50.code", True) <> "" Then
    cString = cString & turnFound(cString) & GrdQry(grid1, "File1_50.code", True)
    aHeader(1) = "[" & "«·’‰› : " & GrdTitle(grid1) & "]"
ElseIf xGroup.BoundText <> "" Then
    aHeader(1) = "[" & " „Ã„Ê⁄… : " & xGroup.Text & "]"
    cString = cString & turnFound(cString) & "  File1_50.GROUP = " & MyParn(xGroup.BoundText)
End If

If xstore.BoundText <> "" Then
    cString = cString & turnFound(cString) & "File1_11.store = " & MyParn(xstore.BoundText)
    aHeader(2) = "[" & "„Œ“‰ : " & xstore.Text & "]"
End If
If XUNIT.BoundText <> "" Then
    cString = cString & turnFound(cString) & "File1_10.UNIT = " & MyParn(XUNIT.BoundText)
    aHeader(2) = "[" & "«·ÊÕœ… : " & XUNIT.Text & "]"
End If

If xdate1.Text <> "" Then
    cString = cString & turnFound(cString) & "ITEMSALES.date >= " & DateSql(xdate1.Text)
    aHeader(3) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If xDate2.Text <> "" Then
    cString = cString & turnFound(cString) & "ITEMSALES.date <= " & DateSql(xDate2.Text)
    aHeader(3) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

cString = cString & " GROUP BY FILE1_10.ITEM, FILE1_10.WIDTH1,FILE1_10.WIDTH2, FILE1_50.DESCA,FILE1_10.GROUP, FILE1_50.group, FILE1_51.DESCA"
contemp.Execute "delete * from print1"
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, CON, adOpenForwardOnly, adLockReadOnly, adCmdText
With sourcetable
    Do Until .EOF
        If Round(Val(!sumOfQuant & ""), 4) <> 0 Then
            contemp.Execute "INSERT INTO PRINT1(ITEM,QUANT,ITEMDESCA,[GROUP],GROUPDESCA,WIDTH1,WIDTH2,MAINGROUP)" & _
                    "VALUES(" & _
                    addstring(!Item) & "," & _
                    addvalue(!sumOfQuant) & "," & _
                    addstring(!desca) & "," & _
                    addstring(!GroupCode) & "," & _
                    addstring(!desca) & "," & _
                    addstring(!width1) & "," & _
                    addstring(!width2) & "," & _
                    addstring(!MAINGROUPCODE) & ")"
       End If
      .MoveNext
    Loop
End With
printTablefrm.doprint retHeader(aHeader, 0, 3), retHeader(aHeader, 3, 2), , bGroupHeader, bMainGroupHeader
printTablefrm.Show 1
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xdate1.Text) And Trim(xdate1.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’ÕÌÕ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then cardlookup
End Sub
Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 5)
xCodeDesca.Caption = GetDesca("select desca from FILE3_10 where code = " & MyParn(xCode.Text)) & ""
End Sub
Sub myProc()
ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Unload Search3
End Sub
Private Sub cardlookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From FILE3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = True

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub

