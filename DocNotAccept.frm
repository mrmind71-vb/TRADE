VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form DocNotAccept 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ »«⁄… «·„” ‰œ«  «· Ï ·„ Ì „ «” ·«„Â« „‰ «·›—⁄"
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20400
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   20400
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   8100
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   45
      Width           =   12075
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2520
         Picture         =   "DocNotAccept.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   90
         Picture         =   "DocNotAccept.frx":24F2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1305
         Picture         =   "DocNotAccept.frx":495E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   7605
         TabIndex        =   4
         Top             =   360
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   90
         TabIndex        =   9
         Top             =   630
         Width           =   2370
         _ExtentX        =   4180
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
         Picture         =   "DocNotAccept.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   420
         Left            =   2520
         TabIndex        =   10
         Top             =   675
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   741
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label2 
         Caption         =   "«·›—⁄ :"
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
         Left            =   11115
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   405
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Height          =   9465
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1215
      Width           =   20040
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   9060
         Left            =   135
         TabIndex        =   2
         Top             =   180
         Width           =   19815
         _cx             =   34951
         _cy             =   15981
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
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
      Top             =   10710
      Width           =   20400
      _ExtentX        =   35983
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
   Begin MSAdodcLib.Adodc DATA1 
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
   Begin MSAdodcLib.Adodc DATA2 
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
Attribute VB_Name = "DocNotAccept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    cHead1 = Me.Caption & "  " & XSTORE.Text
    PrintGrdNew.doprint grid1, 1, , cHead1, , , , False, 10
    PrintGrdNew.Show 1
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub
Private Sub Form_Load()
    openCon con
    
    Set DATA2.Recordset = myRecordSet("Select Code,DescA From File0_40 WHERE ISSTOP = 0 ORDER BY CODE ", con)
    Set XSTORE.RowSource = DATA2
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    
    Set grid1.DataSource = data1
    grid1.Rows = 1
    Fixgrd
End Sub
Private Sub myload()
    Dim cWhere As String, cString As String
    cString = " SELECT flag , Q_DATA_SEND.desca, DOC_NO, DATE, [name], QUANT, USERSEND, TIMESEND , FILE0_40.DESCA , Notes_Doc  FROM  Q_DATA_SEND  INNER JOIN FILE0_40 ON FILE0_40.CODE = Q_DATA_SEND.STORE where DOC_NO is not null "
    If XSTORE.BoundText <> "" Then cString = cString & "  and store = " & MyParn(XSTORE.BoundText)
    If Not bOpt5 And cBranch = "00" Then cString = cString & "  and store in (select store from USERSHOP where code =  " & nusercode & ")"
    cString = cString & " order by DATE , flag , doc_no "
    Set data1.Recordset = myRecordSet(cString, con)
    Fixgrd
End Sub
Sub Fixgrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .Cols = 11
    .RowHeight(0) = 1000
    .TextMatrix(0, 0) = ""
    .TextMatrix(0, 1) = "‰Ê⁄ «·„” ‰œ"
    .TextMatrix(0, 2) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 3) = " «—ÌŒ"
    .TextMatrix(0, 4) = "«·≈”„"
    .TextMatrix(0, 5) = "⁄œœ «·ﬁÿ⁄"
    .TextMatrix(0, 6) = "≈”„ «·—«”·"
    .TextMatrix(0, 7) = " «·Êﬁ "
    .TextMatrix(0, 8) = " «·›—⁄ "
    .TextMatrix(0, 9) = " „·«ÕŸ« "
    .TextMatrix(0, 10) = " › Õ «·„” ‰œ"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 1300
    .ColWidth(2) = 1000
    .ColWidth(3) = 1300
    .ColWidth(4) = 2500
    .ColWidth(5) = 800
    
    .ColWidth(6) = 1800
    .ColWidth(7) = 1300
    .ColWidth(8) = 1500
    .ColWidth(9) = 6000
    .ColWidth(10) = 800
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTDate
    .ColDataType(4) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColComboList(10) = "..."
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With grid1
    Select Case .TextMatrix(Row, 0)
        Case 1
            Purchasefrm.myPublic = 0
            Purchasefrm.sDoc_no = .TextMatrix(Row, 2)
            Purchasefrm.Show
        Case 2
            Purchasefrm.myPublic = 1
            Purchasefrm.sDoc_no = .TextMatrix(Row, 2)
            Purchasefrm.Show
        Case 3
            Transfrm.sDoc_no = .TextMatrix(Row, 2)
            Transfrm.Show
    End Select
End With
myload
End Sub
Private Sub grid1_EnterCell()
With grid1
If .Col = 10 Then
    .Editable = flexEDKbdMouse
Else
    .Editable = flexEDNone
End If
End With
End Sub
