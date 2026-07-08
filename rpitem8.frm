VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpitem8 
   Caption         =   " ﬁ«—Ì— "
   ClientHeight    =   4080
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
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   4080
   ScaleWidth      =   5730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1320
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   0
      Width           =   5460
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1305
         Visible         =   0   'False
         Width           =   1320
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1320
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo xSupGroup 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   3780
         _ExtentX        =   6668
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
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
         Left            =   4050
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1395
         Visible         =   0   'False
         Width           =   570
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
         TabIndex        =   19
         Top             =   1305
         Visible         =   0   'False
         Width           =   2445
      End
      Begin VB.Label Label8 
         Caption         =   "„Ã„Ê⁄… ⁄„·«¡ :"
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
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   225
         Width           =   1365
      End
      Begin VB.Label Label7 
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
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   630
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "«·Ï :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   4005
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   990
         Width           =   465
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
      Height          =   2025
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1305
      Width           =   5415
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   405
         TabIndex        =   4
         Top             =   1260
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VSFlex7LCtl.VSFlexGrid grid1 
         Height          =   1035
         Left            =   45
         TabIndex        =   3
         Top             =   135
         Width           =   3765
         _cx             =   6641
         _cy             =   1826
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
         FormatString    =   $"rpitem8.frx":0000
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
      Begin MSDataListLib.DataCombo XUNIT 
         Height          =   315
         Left            =   405
         TabIndex        =   5
         Top             =   1620
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1620
         Width           =   1005
      End
      Begin VB.Label Label2 
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
         Height          =   240
         Index           =   0
         Left            =   3915
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label4 
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
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1305
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
      Left            =   3150
      RightToLeft     =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3375
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
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   3375
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
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3375
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
      Left            =   3285
      Top             =   3375
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
      TabIndex        =   12
      Top             =   2445
      Width           =   1005
   End
End
Attribute VB_Name = "rpitem8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdApply_Click()
Dim temptable As ADODB.Recordset
Dim sourcetable As ADODB.Recordset
Dim aHeader(6)
contemp.Execute "delete * from temp"

Set temptable = New ADODB.Recordset
temptable.Open "temp", contemp, adOpenKeyset, adLockOptimistic, adCmdTable

cString = "SELECT FILE3_10.CODE, FILE3_10.DESCA, FILE3_20.CODE, FILE3_20.DESCA, Sum(ITEMSALES.Quant) AS SumOFQuant, Sum(ITEMSALES.total) AS SumOFtotal " & _
          " FROM ((((ITEMSALES INNER JOIN FILE3_10 ON ITEMSALES.CODE = FILE3_10.CODE) INNER JOIN FILE3_20 ON FILE3_10.[GROUP] = FILE3_20.CODE) INNER JOIN FILE1_10 ON ITEMSALES.ITEM = FILE1_10.ITEM) INNER JOIN FILE1_50 ON FILE1_10.[GROUP] = FILE1_50.CODE) LEFT JOIN FILE1_51 ON FILE1_50.[GROUP] = FILE1_51.CODE" & _
          " WHERE (ITEMSALES.TYPE = 1 or type = 2) "

If GrdQry(grid1, "file1_51.code", True) <> "" Then
    cString = cString & turnFound(cString) & GrdQry(grid1, "File1_51.code", True)
    aHeader(0) = "[" & "„Ã„Ê⁄… «’‰«› : " & GrdTitle(grid1) & "]"
End If

If xstore.BoundText <> "" Then
    cString = cString & turnFound(cString) & "ITEMSALES.STORE = " & MyParn(xstore.BoundText)
    aHeader(1) = "[" & "„Œ“‰ : " & xstore.Text & "]"
End If


If xUnit.BoundText <> "" Then
    cString = cString & turnFound(cString) & "File1_10.UNIT = " & MyParn(xUnit.BoundText)
    aHeader(2) = "[" & "«·ÊÕœ… : " & xUnit.Text & "]"
End If

If IsDate(xdate1.Text) Then
    cString = cString & turnFound(cString) & "ITEMSALES.date >= " & DateSq(xdate1.Text)
    aHeader(3) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If IsDate(xDate2.Text) Then
    cString = cString & turnFound(cString) & "ITEMSALES.date <= " & DateSq(xDate2.Text)
    aHeader(3) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

If xcode.Text <> "" Then
    cString = cString & turnFound(cString) & "ITEMSALES.CODE = " & MyParn(xcode.Text)
    aHeader(5) = "[" & "«·„Ê—œ : " & xCodeDesca.Caption & "]"
End If

If xSupGroup.BoundText <> "" Then
    cString = cString & turnFound(cString) & "FILE3_10.[GROUP] = " & MyParn(xSupGroup.BoundText)
    aHeader(6) = "[" & " „Ã„Ê⁄… „Ê—œÌ‰ : " & xSupGroup.Text & "]"
End If

cString = cString & " GROUP BY FILE3_10.CODE, FILE3_10.DESCA, FILE3_20.CODE, FILE3_20.DESCA"
          
Set sourcetable = New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

With sourcetable
    Do Until .EOF
            temptable.AddNew
            temptable!STR6 = ![FILE3_20.DESCA]
            temptable!str5 = ![FILE3_20.CODE]
            temptable!str1 = ![FILE3_10.CODE]
            temptable!str2 = ![FILE3_10.DESCA]
            temptable!val1 = !sumOfQuant
            temptable!val2 = !sumofTotal
            temptable!str17 = TurnValue(retHeader(aHeader, 0, 3))
            temptable!str18 = TurnValue(retHeader(aHeader, 4, 6))
            temptable.Update
      .MoveNext
    Loop
End With

If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  ·ÿ»«⁄ Â«"
Else
    main.Report1.ReportFileName = App.Path & "\Reports\item6.rpt"
    contemp.BeginTrans
    contemp.CommitTrans
    main.Report1.DataFiles(0) = "c:\elmorshed\temp.mdb"
    main.Report1.Action = 1
End If

temptable.Close
sourcetable.Close
Set temptable = Nothing
Set sourcetable = Nothing
End Sub
Private Sub CmdClear_Click()
xstore.BoundText = ""
xUnit.BoundText = ""
grid1.Rows = 0
grid1.Rows = 10
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
'grdMake "Select Code,DescA From File1_51", "code", "desca", con, grid1

data1.ConnectionString = strCon
data1.RecordSource = "Select Code,DescA From FILE3_20"
Set xSupGroup.RowSource = data1
xSupGroup.ListField = "Desca"
xSupGroup.BoundColumn = "Code"

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select Code,DescA From File0_40"
Set xstore.RowSource = DATA2
xstore.ListField = "Desca"
xstore.BoundColumn = "Code"

data3.ConnectionString = strCon
data3.RecordSource = "Select Code,DescA From File1_13 order by code"
Set xUnit.RowSource = data3
xUnit.ListField = "Desca"
xUnit.BoundColumn = "Code"
End Sub
Private Sub xGroup_Click(Area As Integer)
'If Area = 2 Then grdMake "Select * From File1_50" & IIf(xGroup.BoundText <> "", " where [GROUP] = " & MyParn(xGroup.BoundText), ""), "code", "desca", con, grid1
End Sub

Private Sub Form_Unload(Cancel As Integer)
'closeCon con
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    xcode.Text = ""
    CardLookup
End If
End Sub
Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""
If xcode.Text = "" Then Exit Sub
xCodeDesca.Caption = GetDesca("select desca from FILE3_10 where code = " & MyParn(xcode.Text)) & ""
End Sub
Sub myProc()
If ActiveControl.Name = xcode.Name Then
    xcode.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    xCodeDesca.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
    Unload Search
End If
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From FILE3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub
