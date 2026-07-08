VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form ModelNotDem2 
   Caption         =   "„ÊœÌ·«  ·Ì” ·Â« ÿ·»Ì… ›Ï «·›« Ê—…"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12495
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   12495
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   870
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   5010
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   45
         Picture         =   "ModelNotDem2.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton CMD_PRINT 
         Height          =   465
         Left            =   1260
         Picture         =   "ModelNotDem2.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   135
         Width           =   1185
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   2475
         TabIndex        =   4
         Top             =   135
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
         Picture         =   "ModelNotDem2.frx":4896
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   45
         TabIndex        =   5
         Top             =   630
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "ModelNotDem2.frx":7091
      Height          =   6570
      Left            =   90
      TabIndex        =   0
      Top             =   945
      Width           =   12300
      _cx             =   21696
      _cy             =   11589
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
   Begin VB.Label xhead 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   5220
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   135
      Width           =   7080
   End
End
Attribute VB_Name = "ModelNotDem2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Public pDoc_No As String

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me

End Sub

Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 1, -2, Me.Caption, ArbString(xhead.Caption), , False, False, 8
    PrintGrd.Show 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub Form_Load()
'                               0               1               2           3               4                   5               6                   7            8              9
    xhead.Caption = ArbString(" ›« Ê—… —ﬁ„  " & Purchasefrm.xDoc_No.Text & "  » «—ÌŒ  " & Purchasefrm.xDate.Text & " ›—⁄ " & Purchasefrm.XSTORE.Text & " ··„Ê—œ " & Purchasefrm.xCodeDesca.Caption)
    grid1.FixedRows = 1
    grid1.Rows = 1
    
    cString = " SELECT  SUB7_20.DOC_NO, FILE0_40.DESCA , FACT.DESCA, FILE1_10.MOSM, FILE1_10.MODELFACT0, FILE1_10.DESCA ,FILE1_10.SCAL, FILE1_10.COLOR, SUB7_20.QUANT,SUB7_20.Q_DEM " & _
              " FROM    SUB7_20 INNER JOIN FILE1_10 ON SUB7_20.ITEM = FILE1_10.ITEM INNER JOIN FACT ON FILE1_10.code = FACT.CODE LEFT OUTER JOIN FILE0_40 ON SUB7_20.store = FILE0_40.CODE WHERE SUB7_20.DOC_NO = " & MyParn(pDoc_No)
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    data1.RecordSource = cString
    data1.Refresh
    FIXGRID
End Sub
Sub FIXGRID()
    With grid1
        .Cols = 10
        .ColHidden(0) = True
        .ColWidth(0) = 0
        .ColWidth(1) = 1000
        .ColWidth(2) = 1500
        .ColWidth(3) = 700
        .ColWidth(4) = 1200
        .ColWidth(5) = 3000
        .ColWidth(6) = 800
        .ColWidth(7) = 1200
        .ColWidth(8) = 700
        .ColWidth(9) = 700
        
        .TextMatrix(0, 1) = "›—⁄"
        .TextMatrix(0, 2) = "„’‰⁄"
        .TextMatrix(0, 3) = "„Ê”„"
        .TextMatrix(0, 4) = "—ﬁ„"
        .TextMatrix(0, 5) = "„ÊœÌ·"
        .TextMatrix(0, 6) = "„ﬁ«”"
        .TextMatrix(0, 7) = "·Ê‰"
        .TextMatrix(0, 8) = "„‘ —Ì« "
        .TextMatrix(0, 9) = "ÿ·»Ì…"
        .MergeCells = flexMergeFree
        
        .MergeCol(0) = True
        .MergeCol(1) = True
        .MergeCol(2) = True
        .MergeCol(3) = True
        .MergeCol(4) = True
        .MergeCol(5) = True
    
    End With
End Sub
