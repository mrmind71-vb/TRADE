VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form NewDocFromFact 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "⁄—÷ «·„” ‰œ«  «·ÃœÌœ… ··›—⁄"
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15120
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
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2925
      Top             =   90
   End
   Begin VB.Frame Frame3 
      Height          =   10500
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   135
      Width           =   19950
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   10050
         Left            =   135
         TabIndex        =   2
         Top             =   180
         Width           =   14730
         _cx             =   25982
         _cy             =   17727
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
      Width           =   15120
      _ExtentX        =   26670
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
End
Attribute VB_Name = "NewDocFromFact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub Form_Load()
    On Error GoTo myerror
    openCon con
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    grid1.Rows = 1
    Fixgrd
    myload
    Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub myload()
    Dim cWhere As String
    If lIsBranchStore Then
        cString = " SELECT     flag, desca, DOC_NO, DATE, [name], QUANT, USERSEND, TIMESEND, ' ' , issend FROM Q_DATA_SEND WHERE STORE = " & MyParn(cBranchStore) & " order by DATE , flag , doc_no "
    Else
        cString = " SELECT     flag, desca, DOC_NO, DATE, [name], QUANT, USERSEND, TIMESEND , ' ' ,  0   FROM Q_DATA_SEND  order by DATE , flag , doc_no "
    End If
    data1.RecordSource = cString
    data1.Refresh
    Fixgrd
End Sub
Sub Fixgrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .WordWrap = True
    .Cols = 10
    .RowHeight(0) = 1000
    .TextMatrix(0, 0) = ""
    .TextMatrix(0, 1) = "‰Ê⁄ «·„” ‰œ"
    .TextMatrix(0, 2) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 3) = " «—ÌŒ"
    .TextMatrix(0, 4) = "«·≈”„"
    .TextMatrix(0, 5) = "⁄œœ «·ﬁÿ⁄"
    .TextMatrix(0, 6) = "≈”„ «·—«”·"
    .TextMatrix(0, 7) = " «·Êﬁ "
    .TextMatrix(0, 8) = " › Õ «·„” ‰œ"
    
    .TextMatrix(0, 9) = " «ŸÂ«— ··›—⁄"
    .ColHidden(9) = Not lIsBranchStore
    .ColWidth(0) = 0
    .ColWidth(1) = 2000
    .ColWidth(2) = 1200
    .ColWidth(3) = 1500
    .ColWidth(4) = 3000
    .ColWidth(5) = 1000
    
    .ColWidth(6) = 2500
    .ColWidth(7) = 1500
    .ColWidth(8) = 1000
    
    
    .ColWidth(9) = 700
    .ColDataType(9) = flexDTBoolean
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTDate
    .ColDataType(4) = flexDTDouble
    .ColDataType(7) = flexDTDouble
    .ColComboList(8) = "..."
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
With grid1
    If lIsBranchStore And .TextMatrix(Row, 9) = 0 Then Exit Sub
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
If .Col = 8 Then
    .Editable = flexEDKbdMouse
Else
    .Editable = flexEDNone
End If
End With
End Sub
Private Sub Timer1_Timer()
    myload
End Sub
