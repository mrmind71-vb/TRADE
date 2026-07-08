VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VsTCashDay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈Ã„«·Ï «·‰ﬁœÌ… ··›—Ê⁄ ·› —…"
   ClientHeight    =   10365
   ClientLeft      =   0
   ClientTop       =   375
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
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   1365
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   -45
      Width           =   3795
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "VsTCashDay.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "VsTCashDay.frx":242A
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2475
         Picture         =   "VsTCashDay.frx":4896
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "⁄—÷"
         Top             =   135
         Width           =   1275
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   8
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
         Picture         =   "VsTCashDay.frx":6D88
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   45
         TabIndex        =   14
         Top             =   1125
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1080
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   11100
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄—÷ «—’œ… ‰ﬁœÌ… «·›—Ê⁄"
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   2400
      End
      Begin MSComCtl2.DTPicker xdate1 
         Height          =   375
         Left            =   8370
         TabIndex        =   9
         Top             =   225
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   441778177
         CurrentDate     =   43098.0832986111
      End
      Begin MSComCtl2.DTPicker XDATE2 
         Height          =   375
         Left            =   6390
         TabIndex        =   10
         Top             =   225
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   441778177
         CurrentDate     =   43098.0832986111
      End
      Begin MSDataListLib.DataCombo xbranch 
         Height          =   315
         Left            =   225
         TabIndex        =   11
         Top             =   225
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777088
         Text            =   ""
         RightToLeft     =   -1  'True
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
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —… :"
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
         Left            =   10350
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   540
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   10035
      Width           =   15120
      _ExtentX        =   26670
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "VsTCashDay.frx":9583
      Height          =   8325
      Left            =   90
      TabIndex        =   1
      Top             =   1125
      Width           =   14955
      _cx             =   26379
      _cy             =   14684
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
      BackColorSel    =   16776960
      ForeColorSel    =   64
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   2
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   180
      Top             =   630
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   180
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
   Begin MSAdodcLib.Adodc data2 
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
Attribute VB_Name = "VsTCashDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conShop As New ADODB.Connection
Dim con As New ADODB.Connection

Dim nRowEnd As Double, nRowTop As Double
Dim nRowBranch As Double
Private Sub cmd_excel_Click()
'    ToFileExel2 grid1, , , , , 1.1, , , , , , Me ' , , , , , 1.1, , , , , , Me
End Sub
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    Dim cHead3 As String
    cHead1 = "≈Ã„«·Ï ‰ﬁœÌ… «·›—Ê⁄ "
    cHead2 = " „‰  «—ÌŒ " & Format(xDate1.Value, "DD-MM-YYYY") & " Õ Ï  «—ÌŒ " & Format(xDate2.Value, "DD-MM-YYYY")
    
    
    Load PrintGrd
'   PrintGrd.doprint grid1, 0.85, , cHead1, cHead2, , , True, 8, , , , True
    PrintGrd.Show 1
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub CmdGo_Click()
    If IsDate(xDate1.Value) And IsDate(xDate2.Value) Then
        myload
    Else
        MsgBox "—«Ã⁄ «· «—ÌŒ"
    End If
End Sub
Private Sub Form_Load()
    openCon con
    Check1.Visible = (cBranch < "60")
    grid1.Rows = 1
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    
    Label2(1).Visible = (cBranch = "00" And cBranch < "60")
    XBRANCH.Visible = (cBranch = "00" And cBranch < "60")
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From branch   ORDER BY code "
    Set XBRANCH.RowSource = DATA2
    XBRANCH.ListField = "Desca"
    XBRANCH.BoundColumn = "Code"
    If cBranch <> "00" Then
        XBRANCH.BoundText = cBranchStore
        XBRANCH.Enabled = True
    End If
    xDate1.Value = DateAdd("d", -1, Date)
    xDate2.Value = DateAdd("d", -1, Date)
    FIXGRID
End Sub
Private Sub myload()
With grid1

'
'    .TextMatrix(0, 8) = "”œ«œ „Ê—œÌ‰"
'    .TextMatrix(0, 9) = "„’«—Ì›"
'    .TextMatrix(0, 10) = " ÕÊÌ· ’«œ—"
'    .TextMatrix(0, 11) = "„”ÕÊ»«  ‘—ﬂ«¡"
'    .TextMatrix(0, 12) = "≈Ìœ«⁄«  »‰ﬂÌ…"

Dim sourcetable As New ADODB.Recordset, nBalance As Double

'--------------  Ê«—œ
cWhere = " AND DATE < " & DateSq(xDate1.Value)
cField1 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove WHERE BOX = FILE0_50.CODE " & _
           cWhere & _
           ") as First_Balance"

'    .TextMatrix(0, 3) = "„»Ì⁄«  ‰ﬁœÏ"

cWhere = " AND DATE >= " & DateSq(xDate1.Value) & " and DATE <= " & DateSq(xDate2.Value)
 cField2 = myiif( _
        " (FLAG = 9 )" & cWhere, "PLUS - MINUS") & _
        " As Sales"

'    .TextMatrix(0, 4) = "”œ«œ ⁄„·«¡"
cField3 = myiif( _
        " (FLAG = 1 OR FLAG = 3 )" & cWhere, "PLUS - MINUS") & _
        " As Cash_In"

'    .TextMatrix(0, 5) = " ÕÊÌ·«  Ê«—œ"
cField4 = myiif( _
        " (FLAG = 8)" & cWhere, "PLUS") & _
        " As Trans_In"

'    .TextMatrix(0, 6) = "≈Ì—«œ« "
cField5 = myiif( _
        " (FLAG =6 OR FLAG =20 )" & cWhere, "PLUS - MINUS") & _
        " As INCOMP "


'    .TextMatrix(0, 8) = "”œ«œ „Ê—œÌ‰"
' ----------- ’«œ—
cField6 = myiif( _
        " (FLAG = 2 OR FLAG = 4 )" & cWhere, "MINUS - PLUS") & _
        " As Cash_out"

'    .TextMatrix(0, 9) = "„’«—Ì›"
cField7 = myiif( _
        " (FLAG = 5)" & cWhere, "MINUS") & _
        " As CHARGES"

'    .TextMatrix(0, 10) = " ÕÊÌ· ’«œ—"

cField8 = myiif( _
        " (FLAG = 7)" & cWhere, "MINUS") & _
        " As TRANS_OUT"
        
'    .TextMatrix(0, 11) = "„”ÕÊ»«  ‘—ﬂ«¡"
cField9 = myiif( _
        " (FLAG = 17)" & cWhere, "MINUS-PLUS") & _
        " As PART"
        
'    .TextMatrix(0, 12) = "≈Ìœ«⁄«  »‰ﬂÌ…"
cField10 = myiif( _
        " (FLAG = 14 OR FLAG = 15 OR FLAG = 16 OR FLAG = 13)" & cWhere, "MINUS-PLUS") & _
        " As PART"

'    .TextMatrix(0, 10) = " ÕÊÌ· Ê«—œ"


cWhere = " AND DATE <= " & DateSq(xDate2.Value)
cField11 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove WHERE BOX = FILE0_50.CODE " & _
           cWhere & _
           ") as Last_Balance"

cString = "Select FILE0_50.CODE , FILE0_50.DESCA , " & cField1 & "," & cField2 & "," & cField3 & "," & cField4 & "," & cField5 & ", ' ' AS N6 , " & _
           cField6 & "," & cField7 & "," & cField8 & "," & cField9 & "," & cField10 & ", ' '  AS N12 , " & cField11 & _
           " From BOXMOVE INNER JOIN FILE0_50 ON FILE0_50.CODE = BOXMOVE.BOX WHERE FILE0_50.ISSTOP = 0 "
If XBRANCH.BoundText <> "" Then cString = cString & " and file0_50.branch = " & MyParn(XBRANCH.BoundText)
cString = cString & "group by FILE0_50.CODE , file0_50.DESCA ORDER BY FILE0_50.CODE  "
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

data1.RecordSource = cString
data1.Refresh
End With

nRowTop = 1
nRowEnd = grid1.Rows - 1
nRowBranch = nRowEnd + 2

FIXGRID


If Check1.Value <> 0 Then
    If cBranch = "00" Then
        FixGridBranch
    Else
        FixGridBranch_2
    End If


    With grid1
        For nRow = 1 To nRowEnd
            cStore = .TextMatrix(nRow, 0)
            For NR = nRowBranch To .Rows - 1
                If .TextMatrix(NR, 0) = cStore Then
                    For nCol = 2 To .Cols - 1
                        If Val(.TextMatrix(nRow, nCol)) <> Val(.TextMatrix(NR, nCol)) Then
                            .Cell(flexcpBackColor, nRow, nCol) = vbRed
                            .Cell(flexcpBackColor, NR, nCol) = vbRed
                        End If
                    Next nCol
                End If
            Next NR
        Next nRow
    End With

End If
With grid1
    If .Rows > 1 Then
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
End With
End Sub
Sub FIXGRID()
With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
    .FrozenCols = 2
    .Cols = 15
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·Œ“‰…"
    .TextMatrix(0, 2) = "—’Ìœ «Ê·"
    
    .TextMatrix(0, 3) = "„»Ì⁄«  ‰ﬁœÏ"
    .TextMatrix(0, 4) = "”œ«œ ⁄„·«¡"
    .TextMatrix(0, 5) = " ÕÊÌ·«  Ê«—œ"
    .TextMatrix(0, 6) = "≈Ì—«œ« "
    .TextMatrix(0, 7) = "≈Ã„«·Ï Ê«—œ"
    
    .TextMatrix(0, 8) = "”œ«œ „Ê—œÌ‰"
    .TextMatrix(0, 9) = "„’«—Ì›"
    .TextMatrix(0, 10) = " ÕÊÌ· ’«œ—"
    .TextMatrix(0, 11) = "„”ÕÊ»«  ‘—ﬂ«¡"
    .TextMatrix(0, 12) = "≈Ìœ«⁄«  »‰ﬂÌ…"
    
    .TextMatrix(0, 13) = "≈Ã„«·Ï Ê«—œ"
    
    .TextMatrix(0, 14) = "«·—’Ìœ"
    
    .ColWidth(0) = 500
    .ColWidth(1) = 2000
    .SubtotalPosition = flexSTBelow
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 7) = Val(.TextMatrix(nRow, 3)) + Val(.TextMatrix(nRow, 4)) + Val(.TextMatrix(nRow, 5)) + Val(.TextMatrix(nRow, 6))
        .TextMatrix(nRow, 13) = Val(.TextMatrix(nRow, 8)) + Val(.TextMatrix(nRow, 9)) + Val(.TextMatrix(nRow, 10)) + Val(.TextMatrix(nRow, 11)) + Val(.TextMatrix(nRow, 12))
    Next nRow
    For nCol = 2 To .Cols - 1
        .ColWidth(nCol) = 1200
        .ColDataType(nCol) = flexDTDouble
        .ColFormat(nCol) = "#0.00"
        .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, vbRed, True, " "
    Next nCol
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    closeCon con
End Sub
Sub FixGridBranch()
    Dim BranchTable As New ADODB.Recordset
    If XBRANCH.BoundText = "" Then
        BranchTable.Open "SELECT * FROM BRANCH WHERE ipsql IS NOT NULL AND CODE <> '00' ORDER BY CODE  ", con, adOpenStatic, adLockReadOnly, adCmdText
    Else
        BranchTable.Open "SELECT * FROM BRANCH WHERE ipsql IS NOT NULL AND CODE = " & MyParn(XBRANCH.BoundText), con, adOpenStatic, adLockReadOnly, adCmdText
    End If
    With BranchTable
    Do While Not .EOF
        cDataShop = !data
        strConShop = LoadConStringshop(!IPsql, !data)
        If openConShop(conShop, strConShop) = "ok" Then
            grid1.AddItem ""
            LoadBalBranch !CODE, conShop
        End If
        .MoveNext
    Loop
    End With
    If grid1.Rows > 1 Then grid1.Cell(flexcpAlignment, 1, 0, grid1.Rows - 1, grid1.Cols - 1) = 4
End Sub

Sub FixGridBranch_2()
    Dim BranchTable As New ADODB.Recordset
    BranchTable.Open "SELECT * FROM BRANCH ", con, adOpenStatic, adLockReadOnly, adCmdText
    With BranchTable
    Do While Not .EOF
        cDataShop = !dataSERVER
        strConShop = LoadConStringshop(!ipserver, !dataSERVER)
        If openConShop(conShop, strConShop) = "ok" Then
            grid1.AddItem ""
            grid1.AddItem ""
            grid1.TextMatrix(grid1.Rows - 1, 0) = !STORE
            grid1.TextMatrix(grid1.Rows - 1, 1) = !DESCA
            LoadBalBranch !STORE, conShop
        End If
        .MoveNext
    Loop
    End With
    If grid1.Rows > 1 Then grid1.Cell(flexcpAlignment, 1, 0, grid1.Rows - 1, grid1.Cols - 1) = 4
End Sub

Private Function LoadConStringshop(pServer, pData)
Dim cServerName As String, cUserId As String, cPassword As String
cUserId = decrypt(RetSetting("userId", App.Path & "\conf.txt"), "dr")
cPassword = decrypt(RetSetting("Password", App.Path & "\conf.txt"), "dr")
LoadConStringshop = "provider=SQLOLEDB;data source=" & pServer & ";initial " _
            & "catalog=" & pData & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Function openConShop(ByRef pCon As ADODB.Connection, Optional ByVal pString As String = "") As String
On Error GoTo myerror
Dim cString As String
If pString = "" Then cString = strConShop Else cString = pString
If pCon.State = adStateOpen Then pCon.Close
pCon.CursorLocation = adUseClient
pCon.Open cString
openConShop = "ok"
Exit Function
myerror:
openConShop = Err.Description
Err.Clear
End Function
Sub LoadBalBranch(pstore, pConShop)
Dim sourcetable As New ADODB.Recordset, nBalance As Double

'--------------  Ê«—œ
cWhere = " AND DATE < " & DateSq(xDate1.Value)
cField1 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove WHERE BOX = FILE0_50.CODE " & _
           cWhere & _
           ") as First_Balance"

cWhere = " AND DATE >= " & DateSq(xDate1.Value) & " and DATE <= " & DateSq(xDate2.Value)
 cField2 = myiif( _
        " (FLAG = 9 )" & cWhere, "PLUS - MINUS") & _
        " As Sales"

cField3 = myiif( _
        " (FLAG = 1 )" & cWhere, "PLUS - MINUS") & _
        " As Cash_In"

cField4 = myiif( _
        " (FLAG = 8)" & cWhere, "PLUS") & _
        " As Trans_In"

cField5 = myiif( _
        " (FLAG =17)" & cWhere, "PLUS") & _
        " as t17 "

' ----------- ’«œ—
cField7 = myiif( _
        " (FLAG = 2 )" & cWhere, "MINUS - PLUS") & _
        " As Cash_out"

cField8 = myiif( _
        " (FLAG = 5)" & cWhere, "MINUS") & _
        " As CHARGES"

cField9 = myiif( _
        " (FLAG = 16)" & cWhere, "MINUS") & _
        " As CHQ_OUT"

cField10 = myiif( _
        " (FLAG = 7)" & cWhere, "MINUS") & _
        " As TRANS_OUT"
        
cField11 = myiif( _
        " (FLAG = 17)" & cWhere, "MINUS") & _
        " As PART"
        
cWhere = " AND DATE <= " & DateSq(xDate2.Value)
cField13 = "(" & _
           "Select Sum(PLUS - MINUS) From BoxMove WHERE BOX = FILE0_50.CODE " & _
           cWhere & _
           ") as Last_Balance"

cString = "Select FILE0_50.CODE , FILE0_50.DESCA , " & cField1 & "," & cField2 & "," & cField3 & "," & cField4 & "," & cField5 & ", ' ' AS N6 , " & _
           cField7 & "," & cField8 & "," & cField9 & "," & cField10 & "," & cField11 & ", ' '  AS N12 , " & cField13 & _
           " From BOXMOVE INNER JOIN FILE0_50 ON FILE0_50.CODE = BOXMOVE.BOX "

cString = cString & "group by FILE0_50.CODE , file0_50.DESCA ORDER BY FILE0_50.CODE  "
sourcetable.Open cString, pConShop, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.RecordCount > 0 Then
    With grid1
        Do While Not sourcetable.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = sourcetable!CODE
            .TextMatrix(.Rows - 1, 1) = sourcetable!DESCA
            .TextMatrix(.Rows - 1, 2) = sourcetable!First_Balance & ""
            .TextMatrix(.Rows - 1, 3) = sourcetable!sales & ""
            .TextMatrix(.Rows - 1, 4) = sourcetable!Cash_In & ""
            .TextMatrix(.Rows - 1, 5) = sourcetable!Trans_In & ""
            .TextMatrix(.Rows - 1, 6) = sourcetable!t17 & ""
            
            .TextMatrix(.Rows - 1, 8) = sourcetable!Cash_out & ""
            .TextMatrix(.Rows - 1, 9) = sourcetable!CHARGES & ""
            .TextMatrix(.Rows - 1, 10) = sourcetable!CHQ_OUT & ""
            .TextMatrix(.Rows - 1, 11) = sourcetable!TRANS_OUT & ""
            .TextMatrix(.Rows - 1, 12) = sourcetable!PART & ""
            .TextMatrix(.Rows - 1, 14) = sourcetable!Last_Balance & ""
            nRow = .Rows - 1
            .TextMatrix(nRow, 7) = Val(.TextMatrix(nRow, 3)) + Val(.TextMatrix(nRow, 4)) + Val(.TextMatrix(nRow, 5)) + Val(.TextMatrix(nRow, 6))
            .TextMatrix(nRow, 13) = Val(.TextMatrix(nRow, 8)) + Val(.TextMatrix(nRow, 9)) + Val(.TextMatrix(nRow, 10)) + Val(.TextMatrix(nRow, 11)) + Val(.TextMatrix(nRow, 12))
            sourcetable.MoveNext
        Loop
    End With
End If
End Sub



