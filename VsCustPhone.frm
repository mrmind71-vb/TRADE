VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form VsCustPhone 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "«” ⁄·«„ ›Ê« Ì— "
   ClientHeight    =   9420
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15150
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
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9420
   ScaleWidth      =   15150
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Height          =   600
      Left            =   90
      Picture         =   "VsCustPhone.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1635
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   3780
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   11265
      Begin VB.CommandButton Command1 
         Caption         =   "Ÿ»ÿ «Œ—  «—ÌŒ «Œ— „»Ì⁄« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   765
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   540
         Width           =   3300
      End
      Begin VB.TextBox XPHONE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   990
         Width           =   3420
      End
      Begin VB.TextBox XDATE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   765
         MaxLength       =   15
         TabIndex        =   10
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox XDATE1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2520
         MaxLength       =   15
         TabIndex        =   8
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   3420
      End
      Begin MSDataListLib.DataCombo XSTORE 
         Height          =   315
         Left            =   5985
         TabIndex        =   11
         Top             =   180
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label xcount 
         Alignment       =   2  'Center
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   945
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1125
         Width           =   3120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰   ·Ì›Ê‰ :"
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
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1035
         Width           =   1350
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·› —… „‰ "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4170
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "«·›—⁄ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "»ÕÀ ⁄‰  ⁄„Ì· :"
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
         Index           =   4
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   622
         Width           =   1245
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9045
      Width           =   15150
      _ExtentX        =   26723
      _ExtentY        =   661
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1485
      Top             =   45
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   1980
      Top             =   -45
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
      Height          =   2835
      Left            =   180
      TabIndex        =   5
      Top             =   7560
      Visible         =   0   'False
      Width           =   195
      _cx             =   344
      _cy             =   5001
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
      Height          =   600
      Left            =   2160
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   1058
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
      Picture         =   "VsCustPhone.frx":246C
      Caption         =   "⁄—÷ "
      Alignment       =   4
      ButtonStyle     =   1
      PictureAlignment=   1
      BevelWidth      =   10
      ShapeSize       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6900
      Left            =   45
      TabIndex        =   12
      Top             =   1530
      Width           =   15045
      _cx             =   26538
      _cy             =   12171
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   50
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
      ScrollBars      =   2
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
   Begin Threed.SSCommand cmd_excel 
      Height          =   555
      Left            =   810
      TabIndex        =   16
      Top             =   630
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   979
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
      Picture         =   "VsCustPhone.frx":498E
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin ComctlLib.ProgressBar prog1 
      Height          =   330
      Left            =   135
      TabIndex        =   17
      Top             =   1215
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   582
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "VsCustPhone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CMD_EXEL_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
End Sub

Private Sub Command1_Click()
    If cBranch = "00" Then
        con.Execute " UPDATE SUBCUST SET F_DATE = (SELECT MAX(DATE) FROM QFILE6_20H WHERE PHONE = SUBCUST.PHONE) ", nRec
    Else
        con.Execute " UPDATE SUBCUST SET F_DATE = (SELECT MAX(DATE) FROM FILE6_20H WHERE PHONE = SUBCUST.PHONE) ", nRec
    End If
    Inform "  „ Ÿ»ÿ ⁄œœ  " & nRec
End Sub
Private Sub Form_Load()
    openCon con
    
    If cBranch = "00" Then
        Set data1.Recordset = myRecordSet("Select Code,DescA From BRANCH union all Select Code,DescA From BRANCH_FR ", con)
    Else
        Set data1.Recordset = myRecordSet("Select branch as code ,DescA From file0_40", con)
    End If
    Set XSTORE.RowSource = data1
    XSTORE.ListField = "Desca"
    XSTORE.BoundColumn = "Code"
    XCOUNT.Caption = ""
    If cBranch <> "00" Then
        XSTORE.Visible = False
    End If
'    XSTORE.BoundText = cManBox
    
    xdate1.text = Format(Date, "dd-mm-yyyy")
    xDate2.text = Format(Date, "dd-mm-yyyy")
    
    Set grid1.DataSource = DATA2
    
    grid1.Rows = 1
    FIXGRID
End Sub
Private Sub myload()
    On Error GoTo myerror
    Dim cStr1 As String

    cStr1 = "SELECT  max(SUBCUST.DESCA) AS NAME , SUBCUST.PHONE, max(SUBCUST.branch) , max(f_date) , max(id) FROM SUBCUST where phone is not null "
    If XSTORE.BoundText <> "" Then cStr1 = cStr1 & " and branch = " & MyParn(XSTORE.BoundText)
    If IsDate(xdate1.text) Then cStr1 = cStr1 & " and f_date >= " & DateSq(xdate1.text)
    If IsDate(xDate2.text) Then cStr1 = cStr1 & " and f_date <= " & DateSq(xDate2.text)
    cStr1 = cStr1 & " group by SUBCUST.PHONE  ORDER BY max(f_date) "
    Set DATA2.Recordset = myRecordSet(cStr1, con)
    FIXGRID
    Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub FIXGRID()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .RowHeight(0) = 1000
    .WordWrap = True
    
    .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
    .Cols = 5
    .TextMatrix(0, 0) = "«·⁄„Ì·"
    .TextMatrix(0, 1) = "«· ·Ì›Ê‰"
    .TextMatrix(0, 2) = "«·›—⁄"
    .TextMatrix(0, 3) = " «—ÌŒ  ”ÃÌ·"
    
    .ColWidth(0) = 3000
    .ColWidth(1) = 3000
    .ColWidth(2) = 2000
    .ColWidth(3) = 1500
    .ColHidden(4) = True
    .ColComboList(2) = StrList("SELECT CODE,DESCA FROM BRANCH", con)
    .ColDataType(3) = flexDTDate
    XCOUNT.Caption = "⁄œœ «·⁄„·«¡ : " & grid1.Rows - 1
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
Dim cPhone  As String
With grid1
    If .Col = 1 Then
        cPhone = GetDesca("SELECT PHONE FROM SUBCUST WHERE ID = " & Val(.TextMatrix(.Row, 4)), con)
        If lIsBranchStore Then
            con.Execute " update SUBCUST set phone = " & addstring(.TextMatrix(.Row, 1)) & " where id = " & Val(.TextMatrix(.Row, 4))
        Else
            con.Execute " update SUBCUST set isnew = 1 , phone = " & addstring(.TextMatrix(.Row, 1)) & " where id = " & Val(.TextMatrix(.Row, 4))
        End If
        con.Execute " update FILE6_20H set PHONE = " & addstring(.TextMatrix(.Row, 1)) & " where PHONE = " & MyParn(cPhone)
        Inform "  „ «· ⁄œÌ· "
    End If

    If .Col = 0 Then
        If lIsBranchStore Then
            con.Execute " update SUBCUST set DESCA = " & addstring(.TextMatrix(.Row, 0)) & " where id = " & Val(.TextMatrix(.Row, 4))
        Else
            con.Execute " update SUBCUST set isnew = 1 , DESCA = " & addstring(.TextMatrix(.Row, 0)) & " where id = " & Val(.TextMatrix(.Row, 4))
        End If
        Inform "  „ «· ⁄œÌ· "
    End If
End With
End Sub
Private Sub grid1_EnterCell()
    With grid1
        If .Col = 0 Or .Col = 1 And bopt2 = True Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
        
    End With
End Sub
