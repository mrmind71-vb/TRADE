VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsModelNoLook 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ›’Ì·Ï —’Ìœ „ﬁ«”«  „ÊœÌ·"
   ClientHeight    =   5655
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   12705
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5655
   ScaleWidth      =   12705
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   3555
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   9060
      Begin VB.TextBox xDescA 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   2655
         Locked          =   -1  'True
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   585
         Width           =   4785
      End
      Begin VB.TextBox xModelFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   135
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox xmodelno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   135
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   540
         Visible         =   0   'False
         Width           =   1545
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   5085
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   180
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XMOSM 
         Height          =   315
         Left            =   2835
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈”„ «·„ÊœÌ·  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’‰⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   7515
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   675
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   195
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdExit 
      Height          =   510
      Left            =   45
      Picture         =   "VsModelNoLook.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      Width           =   1635
   End
   Begin VB.Frame Frame4 
      Caption         =   "√—’œ… „ﬁ«”«  Ê«·Ê«‰ «·„ÊœÌ·"
      Height          =   3930
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1260
      Width           =   12480
      Begin VSFlex7Ctl.VSFlexGrid Grid1 
         Height          =   3495
         Left            =   90
         TabIndex        =   2
         Top             =   270
         Width           =   12300
         _cx             =   21696
         _cy             =   6165
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   3
         FixedCols       =   2
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
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   5325
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   582
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   17639
            MinWidth        =   17639
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1260
      Top             =   585
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
      Left            =   45
      Top             =   585
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
   Begin VSFlex7Ctl.VSFlexGrid Grid2 
      Height          =   360
      Left            =   2805
      TabIndex        =   3
      Top             =   585
      Visible         =   0   'False
      Width           =   285
      _cx             =   503
      _cy             =   635
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
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
      Rows            =   1
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   1
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
   Begin Threed.SSCommand cmdgo 
      Height          =   510
      Left            =   1800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   45
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   900
      _Version        =   196610
      ForeColor       =   0
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
      Picture         =   "VsModelNoLook.frx":246C
      Caption         =   "«·—’Ìœ"
      Alignment       =   4
      ButtonStyle     =   1
      PictureAlignment=   1
      BevelWidth      =   10
      ShapeSize       =   1
   End
End
Attribute VB_Name = "VsModelNoLook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sModelNo As String
Public SModelFact As String
Public SFact As String
Public sdesca As String
Public SMosm As String



Public cMyModel As String
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset
Public CardTable As New ADODB.Recordset
Private Sub cmd_look_Click()
    ModelLookupAll Me, Search3
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
con.CursorLocation = adUseClient
con.Open strCon
CardTable.Open "select file1_10h.*,file4_10.desca as supDesca from file1_10h INNER join file4_10 on file1_10h.code = file4_10.code   order by model", con, adOpenStatic, adLockReadOnly, acmdtext

data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FACT ORDER BY DESCA "
Set xFact.RowSource = data1
xFact.ListField = "DESCA"
xFact.BoundColumn = "CODE"


DATA2.ConnectionString = strCon
DATA2.RecordSource = "SELECT * FROM MOSM ORDER BY DATE DESC "
Set xMosm.RowSource = DATA2
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"
If sModelNo <> "" Then
    xFact.Text = SFact
    xModelFact.Text = SModelFact
    xmodelno.Text = sModelNo
    xDesca.Text = sdesca
    xMosm.Text = SMosm
    myloadGrdModel
    DataGrdModel
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub myDefineGrd()

Grid2.Rows = 3
Grid2.Cols = 2

grid1.Rows = 3
grid1.Cols = 2

grid1.MergeCells = flexMergeRestrictRows
grid1.TextMatrix(0, 1) = "«·„ﬁ«”"
grid1.TextMatrix(1, 1) = "”⁄— „’‰⁄"
grid1.TextMatrix(2, 1) = "”⁄— „” Â·ﬂ"
'grid1.RowHidden(1) = True

End Sub
Private Sub FillItem(cFieldas As String, cField As String)
Dim GRDTABLE As New ADODB.Recordset
' „·∆ «·ÃœÊ·
cString = "Select c_color as [—ﬁ„ «··Ê‰] , Color  as [«··Ê‰]" & turn(cFieldas, ",") & cFieldas & _
          " From " & _
          " (Select c_color,Color,scal,item from file1_10 WHERE MODELno = " & MyParn(xmodelno.Text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(item)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  " & _
          " order by pvt.c_color"

GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grid1.Cols = GRDTABLE.Fields.Count: Grid2.Cols = GRDTABLE.Fields.Count

For nCol = 2 To GRDTABLE.Fields.Count - 1
    grid1.TextMatrix(0, nCol) = GRDTABLE.Fields(nCol).Name
Next

Do Until GRDTABLE.EOF
    Grid2.AddItem ""
    grid1.AddItem ""
    For nCol = 0 To GRDTABLE.Fields.Count - 1
        If nCol <= 1 Then
            grid1.TextMatrix(Grid2.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
        Else
            Grid2.TextMatrix(Grid2.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
            grid1.TextMatrix(grid1.Rows - 1, nCol) = ""
        End If
    Next
    GRDTABLE.MoveNext
Loop
GRDTABLE.Close
Set GRDTABLE = Nothing
End Sub
Private Sub FixCost(cFieldas, cField)
' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,Price from file1_10 WHERE MODELno = " & MyParn(xmodelno.Text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(price)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grid1.Cols - 1
        grid1.TextMatrix(2, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If
End Sub
Public Sub myloadGrdModel()
Dim loctable As New ADODB.Connection
Dim cFieldas As String, cField As String
cString = cString

myDefineGrd

aret = retFields(cFieldas, cField)
If cField = "" And cFieldas = "" Then Exit Sub

FillItem cFieldas, cField
FixCost cFieldas, cField

Fixgrd2

End Sub


Private Sub Fixgrd2()
With grid1
.ColHidden(0) = True
.RowHidden(1) = True

.TextMatrix(0, 1) = "«··Ê‰/«·„ﬁ«”"
.ColWidth(0) = 400
.ColWidth(1) = 1300
nColWidth = (.Width - 200 - .ColWidth(0) - .ColWidth(1)) / .Cols
If nColWidth < 500 Then nColWidth = 500
If nColWidth > 1200 Then nColWidth = 1200
For nCol = 2 To .Cols - 1
    .ColWidth(nCol) = nColWidth
    .ColAlignment(nCol) = flexAlignCenterCenter
Next
End With
End Sub

Public Sub DataGrdModel()
Dim loctable As New ADODB.Recordset, cString As String, cWhere As String, cOr As String, nTotal As Long, nSaveRow As Long, nSaveCol As Long
nSaveRow = grid1.Row: nSaveCol = grid1.Col

ClearGrdModel

cString = "Select FILE1_11.ITEM,SUM([IN] - [OUT]) as Balance FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM"
cString = cString & turn(cString) & " FILE1_10.MODELno = " & MyParn(xmodelno.Text)
cString = cString & " GROUP BY FILE1_11.ITEM"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until loctable.EOF
    For nCol = 2 To Grid2.Cols - 1
        nFound = Grid2.FindRow(loctable!Item, , nCol, False)
        If nFound <> -1 Then
            grid1.TextMatrix(nFound, nCol) = loctable!balance
        End If
    Next
    loctable.MoveNext
Loop
Set loctable = Nothing
If grid1.Rows > 3 Then
    grid1.AddItem ""
    grid1.Cell(flexcpBackColor, grid1.Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = &HE0E0E0
    grid1.Cols = grid1.Cols + 1
    grid1.Cell(flexcpBackColor, 0, grid1.Cols - 1, grid1.Rows - 1, grid1.Cols - 1) = &HE0E0E0
    grid1.TextMatrix(0, grid1.Cols - 1) = "≈Ã„«·Ì ·Ê‰"
    grid1.TextMatrix(grid1.Rows - 1, 1) = "≈Ã„«·Ì „ﬁ«”"
    
    For nCol = 2 To grid1.Cols - 1
        For nRow = 3 To grid1.Rows - 2
            nTotal = nTotal + Val(grid1.TextMatrix(nRow, nCol))
        Next
        grid1.TextMatrix(grid1.Rows - 1, nCol) = Myvalue(nTotal)
        nTotal = 0
    Next
    For nRow = 3 To grid1.Rows - 1
        For nCol = 2 To grid1.Cols - 2
            nTotal = nTotal + Val(grid1.TextMatrix(nRow, nCol))
        Next
        grid1.TextMatrix(nRow, grid1.Cols - 1) = Myvalue(nTotal)
        nTotal = 0
    Next
    grid1.ColAlignment(grid1.Cols - 1) = flexAlignRightCenter
End If
If nSaveRow <= grid1.Rows - 1 Then grid1.Select nSaveRow, nSaveCol
End Sub
Private Sub ClearGrdModel()
With grid1
.Rows = Grid2.Rows
.Cols = Grid2.Cols
For nRow = 3 To .Rows - 1
    For nCol = 2 To .Cols - 1
        .TextMatrix(nRow, nCol) = ""
    Next
Next
End With
End Sub
Private Function retFields(ByRef cFieldas As String, ByRef cField As String)
Dim aret(1) As String
Dim FieldTable As New ADODB.Recordset
'  ⁄—Ì› «·«⁄„œ…
FieldTable.Open "Select SCAL from file1_10 where modelno = " & MyParn(xmodelno.Text) & " GROUP BY SCAL,C_SCAL order by c_scal", con, adOpenStatic, adLockReadOnly
Do Until FieldTable.EOF
    If Not IsNull(FieldTable!scal) Then
        cFieldas = cFieldas & turn(cField, ",") & "[" & FieldTable!scal & "]" & " as " & "[" & FieldTable!scal & "]"
        cField = cField & turn(cField, ",") & "[" & FieldTable!scal & "]"
    End If
    FieldTable.MoveNext
Loop

' ⁄œ„ ÊÃÊœ «⁄„œ…
FieldTable.Close
Set FieldTable = Nothing
End Function
