VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsTBalAcc 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇÑÕÏÉ ÇáÍÓÇÈÇÊ ÇáÃÎÑì"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   1560
      Left            =   5820
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   90
      Width           =   9345
      Begin VB.Frame Frame4 
         Height          =   1455
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   0
         Width           =   3795
         Begin VB.CommandButton cmdGo 
            Height          =   465
            Left            =   2475
            Picture         =   "VsTBalAcc.frx":0000
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "ÚÑÖ"
            Top             =   135
            Width           =   1275
         End
         Begin VB.CommandButton cmdExit 
            Height          =   465
            Left            =   45
            Picture         =   "VsTBalAcc.frx":24F2
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   135
            Width           =   1185
         End
         Begin VB.CommandButton CMD_PRINT 
            Height          =   465
            Left            =   1260
            Picture         =   "VsTBalAcc.frx":495E
            RightToLeft     =   -1  'True
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   135
            Width           =   1185
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
            Picture         =   "VsTBalAcc.frx":6D88
            Caption         =   "ÊÍæíá ÇßÓíá "
            Alignment       =   4
            PictureAlignment=   1
         End
         Begin ComctlLib.ProgressBar prog1 
            Height          =   240
            Left            =   45
            TabIndex        =   9
            Top             =   1170
            Width           =   3705
            _ExtentX        =   6535
            _ExtentY        =   423
            _Version        =   327682
            Appearance      =   1
         End
      End
      Begin VB.TextBox date2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6675
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   315
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "ÍÊì :"
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
         Left            =   8550
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   345
         Width           =   450
      End
   End
   Begin MSAdodcLib.Adodc data1 
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
   Begin MSAdodcLib.Adodc data10 
      Height          =   330
      Left            =   135
      Top             =   675
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
      Height          =   8430
      Left            =   90
      TabIndex        =   3
      Top             =   1710
      Width           =   15045
      _cx             =   26538
      _cy             =   14870
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   8712429
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
      Rows            =   50
      Cols            =   4
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
Attribute VB_Name = "VsTBalAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim i%, sp As String, mCol%, t As Double
Sub Fillgrd()
Dim lOk As Boolean
Dim cF1 As String, cF2 As String, cF3 As String, cF4 As String, cF5 As String, cF6 As String

DoTable

cStr1 = "SELECT  ACC1_10.CODE, ACC1_10.DESCA , CASE WHEN SUM(VAL1 - VAL2) > 0 THEN  SUM(VAL1 - VAL2) ELSE 0  END , CASE WHEN SUM(VAL1 - VAL2) < 0 THEN  SUM(VAL2 - VAL1) ELSE 0 END " & _
        " FROM    ACC1_10 INNER JOIN QT_ACC_MOVE ON ACC1_10.CODE = QT_ACC_MOVE.acc " & _
        " WHERE ACC1_10.FLAG = 6 AND DATE <= " & DateSq(date2.text) & _
        " GROUP BY ACC1_10.CODE, ACC1_10.DESCA "
data10.RecordSource = cStr1
data10.Refresh
End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me

End Sub

Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    cHead1 = "ØÈÇÚÉ ÇáÃÑÕÏÉ " & date2.text
    Load PrintGrd
    PrintGrd.doprint grid1, 1, , cHead1, , , , False, 10
    PrintGrd.Show 1
End Sub
Private Sub cmdGo_Click()
If IsDate(date2.text) Then
    Fillgrd
    FIXGRID
Else
    MsgBox "ÊÓÌíá ÊÇÑíÎ"
End If
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
con.CursorLocation = adUseClient
con.Open strCon
Set AccTable = New ADODB.Recordset

'Date1.Text = "1-1-" & Year(Date)
date2.text = Format(Date, "DD-MM-YYYY")

Set grid1.DataSource = data10
data10.ConnectionString = strCon
grid1.Rows = 1
FIXGRID
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Sub FIXGRID()
With grid1
    .Cols = 4
    .FixedRows = 1
    
    .Cell(flexcpText, 0, 0) = "ßæÏ ÇáÍÓÇÈ"
    .Cell(flexcpText, 0, 1) = "ÇáÍÓÇÈ"
    .Cell(flexcpText, 0, 2) = "ÑÕíÏ ãÏíä"
    .Cell(flexcpText, 0, 3) = "ÑÕíÏ ÏÇÆä"
    
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 2) = Round(Val(.TextMatrix(nRow, 2)), 2)
        .TextMatrix(nRow, 3) = Round(Val(.TextMatrix(nRow, 3)), 2)
    Next nRow
    .ColWidth(0) = 1000
    .ColWidth(1) = 6000
    .ColWidth(2) = 1500
    .ColWidth(3) = 1500
    .ColFormat(2) = "#,###,##0.00"
    .ColFormat(3) = "#,###,##0.00"
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 2, "#0.00", vbYellow, , True, " "
    .Subtotal flexSTSum, -1, 3, "#0.00", vbYellow, , True, " "
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    
End With
End Sub
Sub DoTable()
    On Error Resume Next
    con.Execute " DROP TABLE QT_ACC_MOVE "
    con.Execute " SELECT * INTO QT_ACC_MOVE  FROM ACC_MOVE  "
    Err.Clear
End Sub
