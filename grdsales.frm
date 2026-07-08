VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form grdsales 
   ClientHeight    =   10365
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   3555
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   45
      Width           =   4470
      Begin VB.TextBox XPHONE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   90
         TabIndex        =   15
         Top             =   630
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄ :"
         Height          =   240
         Index           =   5
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   675
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «· ·Ì›Ê‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   1020
      End
   End
   Begin VB.CheckBox xall 
      Alignment       =   1  'Right Justify
      Caption         =   "⁄—÷  ›’Ì·Ï «·»Ê‰"
      Height          =   285
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   45
      Width           =   2355
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   405
      Width           =   2445
      Begin VB.CommandButton cmdGo 
         Height          =   495
         Left            =   1215
         MaskColor       =   &H00E0E0E0&
         Picture         =   "grdsales.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   45
         Picture         =   "grdsales.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1140
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "«·› —…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   8055
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   7125
      Begin VB.TextBox xitem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   315
         Width           =   1815
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   675
         Width           =   1770
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   675
         Width           =   1815
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   315
         Width           =   4470
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬂÊœ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   0
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   315
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   675
         Width           =   330
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2250
      Top             =   -90
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
      Top             =   -630
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
      Left            =   45
      Top             =   -1035
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
   Begin MSAdodcLib.Adodc DATABOX 
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   9195
      Left            =   45
      TabIndex        =   10
      Top             =   1125
      Width           =   15135
      _cx             =   26696
      _cy             =   16219
      _ConvInfo       =   1
      Appearance      =   0
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc Adodc1 
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
End
Attribute VB_Name = "grdsales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim oSearchItem As New Search3
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    If xitem.text = "" And xphone.text = "" Then Exit Sub
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    Set DATA2.Recordset = myRecordSet("Select Code,DescA From FILE0_40 ORDER BY code ", con)
    Set XSTORE.RowSource = DATA2
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    If cBranch <> "00" Then
        XSTORE.BoundText = cBranchStore
        XSTORE.Enabled = False
    End If
    Fixgrd
End Sub
Private Sub Fixgrd()
With grid1
.ExplorerBar = flexExSortShow
.Cols = 13
.TextMatrix(0, 0) = "—ﬁ„ „” ‰œ"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "«·Êﬁ "
.TextMatrix(0, 3) = "«·»«∆⁄"
.TextMatrix(0, 4) = "«·„ÊœÌ·"
.TextMatrix(0, 5) = "«·„’‰⁄"
.TextMatrix(0, 6) = "ﬂÊœ"
.TextMatrix(0, 7) = "«·’‰›"
.TextMatrix(0, 8) = "„»Ì⁄« "
.TextMatrix(0, 9) = "«·”⁄—"
.TextMatrix(0, 10) = "«·≈Ã„«·Ï"

.TextMatrix(0, 11) = "«·›—⁄"
.TextMatrix(0, 12) = "—ﬁ„  ·Ì›Ê‰"


.MergeCells = flexMergeFree
.MergeCol(0) = True
.ColDataType(1) = flexDTDate
.ColFormat(1) = "dd-mm-yyyy"
.ColWidth(0) = 1300
.ColWidth(1) = 1300
.ColWidth(2) = 800
.ColWidth(3) = 1200
.ColWidth(4) = 1500
.ColWidth(5) = 1000
.ColWidth(6) = 700
.ColWidth(7) = 2200
.ColWidth(8) = 600
.ColWidth(9) = 700
.ColWidth(10) = 800

.ColWidth(11) = 1000
.ColWidth(12) = 1500

.ColComboList(11) = StrList("SELECT CODE , DESCA FROM FILE0_40", con)
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
End With
End Sub
Private Sub xDate1_DblClick()
Set datefrm.oDate = xdate1
datefrm.Show 1
End Sub
Private Sub xdate2_DblClick()
Set datefrm.oDate = xDate2
datefrm.Show 1
End Sub

Private Sub xITEM_KeyPress(KeyAscii As Integer)
On Error GoTo myerror
If KeyAscii = 13 Then
    If Len(xitem.text) >= 8 Then
        sBarCode = GetDesca("select item from file1_10 where barcode = " & MyParn(xitem.text), con)
        If sBarCode <> "" Then
            xitem.text = sBarCode
            xDesca.Caption = ItemField(xitem.text, "desca", con)
            myload
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xitem_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then ItemsLookupAll Me, oSearchItem
End Sub
Sub myProc()
xitem.text = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
xDesca.Caption = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 5)
End Sub
Private Sub xitem_LostFocus()
xDesca.Caption = ""
If Not ValidInt(xitem.text) Then Exit Sub
xDesca.Caption = ItemField(xitem.text, "desca", con)
End Sub
Private Sub myload()
Dim cWhere As String, cString As String
If IsDate(xdate1.text) Then cWhere = turn(cWhere, " and ") & " SALESDTL.DATE >= " & DateSq(xdate1.text)
If IsDate(xDate2.text) Then cWhere = cWhere & turn(cWhere, " and ") & " SALESDTL.DATE <= " & DateSq(xDate2.text)
If ValidInt(xitem.text) Then cWhere = cWhere & turn(cWhere, " AND ") & " SALESDTL.DOC_NO IN (SELECT DOC_NO FROM FILE6_20  WHERE FILE6_20.ITEM = " & xitem.text & ")"
If xphone.text <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & " FILE6_20H.PHONE = " & MyParn(xphone.text)
If XSTORE.BoundText <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & " SALESDTL.STORE = " & MyParn(XSTORE.BoundText)

If xall.Value = 0 And xitem.text <> "" Then cWhere = cWhere & turn(cWhere, " AND ") & " SALESDTL.item = " & xitem.text
cString = "SELECT SALESDTL.* , file1_10.modelfact , FILE6_20H.PHONE " & _
          " FROM SALESDTL INNER JOIN FILE6_20H ON FILE6_20H.DOC_NO = SALESDTL.DOC_NO LEFT   join file1_10 on file1_10.item = SALESDTL.item "
If cWhere <> "" Then cString = cString & turn(cWhere, " WHERE ") & cWhere

cString = cString & " ORDER BY SALESDTL.date desc , SALESDTL.DOC_NO,FLAG"

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
With grid1
.Rows = 1
Do Until loctable.EOF
    If loctable!Flag = 0 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "dd-mm-yyyy")
        .TextMatrix(.Rows - 1, 0) = loctable!doc_no
        .TextMatrix(.Rows - 1, 2) = Format(loctable!Time, "HH:NN")
        .TextMatrix(.Rows - 1, 3) = loctable!MANDESCA & ""
        .TextMatrix(.Rows - 1, 4) = DelZero(loctable!modelfact & "")
        .TextMatrix(.Rows - 1, 5) = loctable!fact_Desca & ""
        .TextMatrix(.Rows - 1, 6) = loctable!Item & ""
        .TextMatrix(.Rows - 1, 7) = loctable!DESCA & ""
        .TextMatrix(.Rows - 1, 8) = loctable!Quant
        .TextMatrix(.Rows - 1, 9) = loctable!price & ""
        .TextMatrix(.Rows - 1, 10) = Format(Val(loctable!TOTAL & ""), "Fixed")
        
        .TextMatrix(.Rows - 1, 11) = loctable!STORE
        .TextMatrix(.Rows - 1, 12) = loctable!phone & " "
    
    ElseIf loctable!Flag = 1 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "dd-mm-yyyy")
        .TextMatrix(.Rows - 1, 0) = loctable!doc_no
        .TextMatrix(.Rows - 1, 11) = loctable!STORE
        .TextMatrix(.Rows - 1, 12) = loctable!phone & " "
        For i = 0 To 9
            .TextMatrix(.Rows - 1, i) = "«·Œ’„"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 9 + 1) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
    ElseIf loctable!Flag = 3 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "dd-mm-yyyy")
        .TextMatrix(.Rows - 1, 0) = loctable!doc_no
        .TextMatrix(.Rows - 1, 11) = loctable!STORE
        .TextMatrix(.Rows - 1, 12) = loctable!phone & " "
        For i = 0 To 9
            .TextMatrix(.Rows - 1, i) = "«·«Ã„«·Ì"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 9 + 1) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
    ElseIf loctable!Flag = 5 Then
        .AddItem ""
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "dd-mm-yyyy")
        .TextMatrix(.Rows - 1, 0) = loctable!doc_no
        .TextMatrix(.Rows - 1, 11) = loctable!STORE
        .TextMatrix(.Rows - 1, 12) = loctable!phone & " "
        For i = 0 To 9
            .TextMatrix(.Rows - 1, i) = "”œ«œ ›Ì“«"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 9 + 1) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &H8080FF
    End If
    loctable.MoveNext
Loop
End With
Fixgrd
End Sub


