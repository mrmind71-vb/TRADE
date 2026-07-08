VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdcust_balancefrm 
   BackColor       =   &H00FFFFFF&
   Caption         =   "√—’œ… Ê  ⁄«„·«  «·⁄„·«¡"
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20370
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
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   8820
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   585
      Width           =   5145
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cust_balance.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1305
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cust_balance.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "cust_balance.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   3870
         TabIndex        =   14
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cust_balance.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2565
         TabIndex        =   15
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "cust_balance.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1365
      Left            =   13995
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   -45
      Width           =   5415
      Begin VB.TextBox Xcode 
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
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   585
         Width           =   1185
      End
      Begin VB.TextBox xDate1 
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
         Left            =   1950
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1815
      End
      Begin VB.TextBox xDate2 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   90
         TabIndex        =   4
         Top             =   945
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label xCustName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   585
         Width           =   2445
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ ⁄„Ì·"
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
         Index           =   1
         Left            =   3885
         TabIndex        =   7
         Top             =   630
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„Ã„Ê⁄… ⁄„·«¡"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   990
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   3885
         TabIndex        =   3
         Top             =   270
         Width           =   660
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2475
      Top             =   75
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
      Left            =   150
      Top             =   75
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
      Height          =   8115
      Left            =   90
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1350
      Width           =   19320
      _cx             =   34078
      _cy             =   14314
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Rows            =   1
      Cols            =   15
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   16
      Top             =   9900
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   17
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   4095
         TabIndex        =   18
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   8100
         TabIndex        =   19
         Top             =   45
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   3
         Left            =   12150
         TabIndex        =   20
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   4
         Left            =   16155
         TabIndex        =   21
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   10
      Top             =   9705
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "grdcust_balancefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aHeader(3)
Dim oSearch As New Search_abd
Dim con As New ADODB.Connection
Private Sub cmdExcel_Click()
Me.MousePointer = 11
If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 19)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2))
Me.MousePointer = 0
End Sub

Private Sub cmdExit_Click()
Unload Me

End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub

Private Sub cmdPdf_Click()

End Sub

Private Sub cmdPrint_Click()
myPrint
End Sub

Private Sub Form_Load()
    openCon con
    
    xdate1.text = myFormat_p("1-1-" & Year(Date))
    xdate2.text = myFormat_p(Date)
    

    Set Data1.Recordset = myCmd("SELECT * FROM FILE3_50 ", con)
    Set xGroup.RowSource = Data1
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    Set grid1.DataSource = Data2
    
    fixGrd
End Sub
Private Sub myload()
Dim aPrm As Variant
ReDim aHeder(3)
aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xdate1.text))
aPrm = AddFlag(aPrm, "date2", myFormat_sp(xdate2.text))
aHeader(0) = BetweenString(xdate1.text, xdate2.text)
aPrm = AddFlag(aPrm, "code", TurnValue(xcode.text))
If Trim(xcode.text) <> "" Then
    aHeader(1) = "«·⁄„Ì· : " & xCustName.Caption
End If
    
aPrm = AddFlag(aPrm, "group", TurnValue(xGroup.BoundText))
If xGroup.MatchedWithList Then
    aHeader(1) = "«·„Ã„Ê⁄… : " & xGroup.text
End If
Set Data2.Recordset = myCmd("dbo.sp_client_balance", con, adStoredProc, aPrm)
fixGrd
End Sub
Sub fixGrd()

    With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·⁄„Ì·"
    
    .TextMatrix(0, 2) = "«·—’Ìœ " & myFormat_p(xdate1.text)
    .TextMatrix(0, 3) = "Ã. „»Ì⁄« "
    .TextMatrix(0, 4) = "Ã. „— Ã⁄«  "
    .TextMatrix(0, 5) = "’«›Ï ﬁÌ„… „»Ì⁄« "
    .TextMatrix(0, 6) = " ﬂ·›… «·„»Ì⁄«  "
    .TextMatrix(0, 7) = "—»Õ «·„»Ì⁄« "
    .TextMatrix(0, 8) = " ”œ«œ ‰ﬁœÏ"
    .TextMatrix(0, 9) = "”œ«œ √.ﬁ "
    .TextMatrix(0, 10) = " ”ÊÌ« "
    .TextMatrix(0, 11) = " ”ÊÌ«  Õ”«»« "
    .TextMatrix(0, 12) = "—’Ìœ " & xdate2.text
    .TextMatrix(0, 13) = "‘Ìﬂ«  €Ì— „Õ’·…"
    .TextMatrix(0, 14) = " «—ÌŒ √Œ— „»Ì⁄« "
        
    .ColWidth(0) = 800
    .ColWidth(1) = 2000
    .ColWidth(2) = 1400
    .ColWidth(3) = 1200
    .ColWidth(4) = 1200
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1200
    .ColWidth(12) = 1400
    .ColWidth(13) = 1000
    .ColWidth(14) = 1300
    
    For nCol = 2 To .Cols - 2
        .ColDataType(nCol) = flexDTDouble
    Next nCol
    
    .ColDataType(13 + 1) = flexDTDate
    .ColFormat(13 + 1) = "yyyy/m/d"
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
    panel1(0).Caption = IIf(grid1.Rows = 1, "·«  ÊÃœ ”Ã·«  ", "⁄œœ «·”Ã·«  : " & grid1.Rows - 2)
End Sub

Private Sub Form_Resize()
grid1.Height = IIf(Me.Height - grid1.Top - 1000 < 3000, 3000, Me.Height - grid1.Top - 1000)
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set grdcust_balancefrm = Nothing
End Sub
Private Sub grid1_DblClick()
    If grid1.Row > 1 Then
        Load ClientMoveFrm
        ClientMoveFrm.xcode.text = grid1.TextMatrix(grid1.Row, 0)
        ClientMoveFrm.xdesca.Caption = grid1.TextMatrix(grid1.Row, 1)
        ClientMoveFrm.xdate1.text = xdate1.text
        ClientMoveFrm.xdate2.text = xdate2.text
        ClientMoveFrm.Fillgrd
        ClientMoveFrm.Show
    End If
End Sub
Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then CLIENTLOOKUP Me, oSearch
End Sub
Private Sub xcode_LostFocus()
myLostFocus xcode
xCustName.Caption = ""
If xcode.text = "" Then Exit Sub
xcode.text = RetZero(xcode.text, 4)
xCustName.Caption = GetDesca("select desca from FILE3_10 where code = " & MyParn(xcode.text), con) & ""
End Sub
Sub myProc()
ActiveControl.text = oSearch.grid1.TextMatrix(Search.grid1.Row, 0)
oSearch.Hide
End Sub

Private Sub xcode_GotFocus()
myGotFocus xcode
End Sub
Private Sub xdate1_GotFocus()
myGotFocus xdate1
End Sub
Private Sub xdate1_LostFocus()
myLostFocus xdate1
myValidDate xdate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xdate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xdate2
myValidDate xdate2
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
If grid1.Rows < 3 Then Exit Sub
Dim aRow As Variant, aSub As Variant
aSub = AddFlag(Empty, "row", 1)
aSub = AddFlag(aSub, "col", 1)
aSub = AddFlag(aSub, "cols", 2)
aSub = AddFlag(aSub, "text", "≈Ã„«·Ì")
aRow = AddFlag(aRow, aSub)

PrintGrdNew.DOPRINT Me.grid1, 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption, retHeader(aHeader, 0, 2), , False, True, 10, , aRow, Array(1)
If Not bIgPreview Then
    PrintGrdNew.Show 1
Else
    Unload PrintGrdNew
End If
'PrintGrdNew.Show 1
'Unload PrintGrdNew
End Sub

