VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form SalesDoc 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8640
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15405
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
   ScaleHeight     =   8640
   ScaleWidth      =   15405
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   6180
      Left            =   45
      TabIndex        =   12
      Top             =   2385
      Width           =   15315
      _ExtentX        =   27014
      _ExtentY        =   10901
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   " ›’Ì·Ï »Ê‰ Ê «·„— Ã⁄"
      TabPicture(0)   =   "SalesDoc.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "CMD_SHOW"
      Tab(0).Control(1)=   "GRID11"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " ›’Ì·Ï »Ê‰ «·„»Ì⁄« "
      TabPicture(1)   =   "SalesDoc.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.CommandButton CMD_SHOW 
         Caption         =   "⁄—÷  ›’Ì·Ï „»Ì⁄«  Ê „— Ã⁄ «·»Ê‰"
         Height          =   330
         Left            =   -74865
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   5715
         Width           =   3435
      End
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   4965
         Left            =   90
         TabIndex        =   13
         Top             =   360
         Width           =   15135
         _cx             =   26696
         _cy             =   8758
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
      Begin VSFlex7Ctl.VSFlexGrid GRID11 
         Height          =   5280
         Left            =   -74910
         TabIndex        =   14
         Top             =   405
         Width           =   15135
         _cx             =   26696
         _cy             =   9313
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
   End
   Begin VB.CommandButton cmd_retall2 
      Caption         =   "„— Ã⁄ «·⁄—÷ »«·ﬂ«„·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   810
      Width           =   2400
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2445
      Left            =   11880
      Picture         =   "SalesDoc.frx":0038
      ScaleHeight     =   2445
      ScaleWidth      =   3210
      TabIndex        =   2
      Top             =   45
      Width           =   3210
      Begin ImgXCtrl6.ImgXCtrl imgx1 
         Height          =   2175
         Left            =   45
         TabIndex        =   3
         Top             =   45
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   3836
         BackColor       =   16777215
         BorderStyle     =   0
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "amr123"
         LicenseRegCode  =   "íß“µª∫≥ºΩ∫æ“±æß´¥Ω∏•OANH-FAZOHIZB-EFTP6gI"
      End
   End
   Begin VB.CommandButton CMD_EXIT 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1890
      Width           =   2400
   End
   Begin VB.CommandButton cmd_retall 
      Caption         =   "„— Ã⁄ ··»Ê‰ »«·ﬂ«„·"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   135
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1350
      Width           =   2400
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin VB.Label xNOTE_DISC2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5310
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   270
      Width           =   5145
   End
   Begin VB.Label xretdisc1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   6795
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   810
      Width           =   3660
   End
   Begin VB.Label xretdisc2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2655
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   810
      Width           =   4065
   End
   Begin VB.Label xNOTE_DISC 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5310
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1350
      Width           =   5145
   End
   Begin VB.Label xDOC_DISC 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2655
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1350
      Width           =   2535
   End
   Begin VB.Label xdate 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1890
      Width           =   2535
   End
   Begin VB.Label xbranch 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5310
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1890
      Width           =   2535
   End
   Begin VB.Label xdoc_no 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2655
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1890
      Width           =   2535
   End
End
Attribute VB_Name = "SalesDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cInv_No As String, nDiscInvoice As Double, cNote_Disc As String, cMan As String, lConServerRet As Boolean, lRetOnePecs As Boolean, lRetOnePecs2 As Boolean
Dim lOtherBranch As Boolean, lSalesDisc2 As Boolean
Dim lConServer As Boolean
Public sDoc_no As String
Dim con As New ADODB.Connection
Private Sub CMD_EXIT_Click()
    If lConServerRet Then conShop.Execute " UPDATE FILE6_20H SET ISRET = 0 WHERE DOC_NO = " & MyParn(sDoc_no)
    Unload Me
End Sub
Private Sub cmd_retall_Click()
With grid1
    SalesFrm.XSALES_RET.Caption = sDoc_no
    SalesFrm.xMan.BoundText = cMan
    If SalesFrm.grid1.Rows > 2 Then
        MsgBox "ÌÊÃœ «’‰«› „”Ã·… ⁄·Ï «·»Ê‰"
        Exit Sub
    End If
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 6)) <> 0 And (.ValueMatrix(nRow, 8) - .ValueMatrix(nRow, 9)) > 0 Then
            If lConServerRet And .TextMatrix(nRow, 2) <> cBranchStore Then
                If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & Val(.TextMatrix(nRow, 6)), con) = "" Then
                    sendItems Val(.TextMatrix(nRow, 6))
                End If
            End If
            If IsNumeric(SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 1)) Then
                SalesFrm.grid1.AddItem ""
            End If
            .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = vbRed
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 10) = (.ValueMatrix(nRow, 8) - .ValueMatrix(nRow, 9)) * -1
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 1) = .TextMatrix(nRow, 6)
            SalesFrm.GrdDesc .TextMatrix(nRow, 6), SalesFrm.grid1.Rows - 1
            SalesFrm.xMan.BoundText = .TextMatrix(nRow, 3)
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 16) = .TextMatrix(nRow, 3)
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 15) = .TextMatrix(nRow, 3)
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 11) = Val(.TextMatrix(nRow, 10))
            SalesFrm.grid1_AfterEdit SalesFrm.grid1.Rows - 1, 1
            If Not lRetOnePecs2 And Not lRetOnePecs Then SalesFrm.XISNODEL.Value = 1
        End If
    Next nRow
    If lConServerRet Then conShop.Execute " UPDATE FILE6_20H SET ISRET = 0 WHERE DOC_NO = " & MyParn(sDoc_no)
    Unload Me
End With
End Sub
Private Sub cmd_retall2_Click()
With grid1
    SalesFrm.XSALES_RET.Caption = sDoc_no
    SalesFrm.xMan.BoundText = cMan
    If SalesFrm.grid1.Rows > 2 Then
        MsgBox "ÌÊÃœ «’‰«› „”Ã·… ⁄·Ï «·»Ê‰"
        Exit Sub
    End If
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 6)) <> 0 And (.ValueMatrix(nRow, 8) - .ValueMatrix(nRow, 9)) > 0 And .TextMatrix(nRow, 12) <> "" And Val(.TextMatrix(nRow, 13)) > 0 Then
            If lConServerRet And .TextMatrix(nRow, 2) <> cBranchStore Then
                If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & Val(.TextMatrix(nRow, 6)), con) = "" Then
                    sendItems Val(.TextMatrix(nRow, 6))
                End If
            End If
            If IsNumeric(SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 1)) Then
                SalesFrm.grid1.AddItem ""
            End If
            .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = vbRed
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 10) = (.ValueMatrix(nRow, 8) - .ValueMatrix(nRow, 9)) * -1
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 1) = .TextMatrix(nRow, 6)
            SalesFrm.GrdDesc .TextMatrix(nRow, 6), SalesFrm.grid1.Rows - 1
            SalesFrm.xMan.BoundText = .TextMatrix(nRow, 3)
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 16) = .TextMatrix(nRow, 3)
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 15) = .TextMatrix(nRow, 3)
            SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 11) = Val(.TextMatrix(nRow, 10))
            SalesFrm.grid1_AfterEdit SalesFrm.grid1.Rows - 1, 1
            If Not lRetOnePecs2 And Not lRetOnePecs Then SalesFrm.XISNODEL.Value = 1
        End If
    Next nRow
    If lConServerRet Then conShop.Execute " UPDATE FILE6_20H SET ISRET = 0 WHERE DOC_NO = " & MyParn(sDoc_no)
    Unload Me
End With
End Sub
Private Sub CMD_SHOW_Click()
Dim cString As String
    cString = " SELECT   FILE6_20H.DOC_NO2, FILE6_20H.date, FILE6_20H.store, FILE6_20H.BOX, FILE6_20H.CASH, FILE6_20H.VISA, FILE1_10.ITEM, FACT.DESCA, FILE1_10.desca ,FILE1_10.SCAL,FILE1_10.COLOR, FILE6_20.QUANT as sales , 0  as ret , FILE6_20.PRICE, FILE6_20.TOTAL , FILE6_20H.DOC_NO " & _
          " FROM     FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE " & _
          " where FILE6_20H.DOC_NO = " & MyParn(xDoc_No.Caption) & _
          " Union All " & _
          " SELECT   FILE6_20H.DOC_NO2, FILE6_20H.date, FILE6_20H.store, FILE6_20H.BOX, FILE6_20H.CASH, FILE6_20H.VISA, FILE1_10.ITEM, FACT.DESCA, FILE1_10.desca ,FILE1_10.SCAL,FILE1_10.COLOR, 0 , FILE6_20.QUANT*-1 , FILE6_20.PRICE, FILE6_20.TOTAL , file6_20h.sales_ret " & _
          " FROM     FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE " & _
          " where FILE6_20H.DOC_NO in(select doc_no from file6_20h as file6_20h_2 where FILE6_20.QUANT < 0 AND file6_20h.sales_ret  = " & MyParn(xDoc_No.Caption) & ")"
    If lConServerRet Then          ' «·„— Ã⁄ „‰ «·”Ì—›—
        Set DATA11.Recordset = myRecordSet(cString, conShop)
    Else
        Set DATA11.Recordset = myRecordSet(cString, con)
    End If
    Fixgrd11
End Sub
Private Sub Fixgrd11()
With GRID11
'                       0               1               2               3               4               5               6               7       8                   9           10                  11              12              13              14              15
'    " SELECT   FILE6_20H.DOC_NO, FILE6_20H.date, FILE6_20H.store, FILE6_20H.BOX, FILE6_20H.CASH, FILE6_20H.VISA, FILE1_10.ITEM, FACT.DESCA, FILE1_10.desca ,FILE1_10.SCAL,FILE1_10.COLOR, FILE6_20.QUANT as sales , 0  as ret , FILE6_20.PRICE, FILE6_20.TOTAL , FILE6_20H.DOC_NO " & _

.Cols = 16

.TextMatrix(0, 0) = "—ﬁ„ „” ‰œ"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "«·›—⁄"
.TextMatrix(0, 3) = "«·Œ“‰…"
.TextMatrix(0, 4) = "‰ﬁœÏ"
.TextMatrix(0, 5) = "›Ì“«"
.TextMatrix(0, 6) = "»«—ﬂÊœ"
.TextMatrix(0, 7) = "«·„’‰⁄"
.TextMatrix(0, 8) = "«·’‰›"
.TextMatrix(0, 9) = "„ﬁ«”"
.TextMatrix(0, 10) = "«··Ê‰"

.TextMatrix(0, 11) = "„»Ì⁄« "
.TextMatrix(0, 12) = "„— Ã⁄"
.TextMatrix(0, 13) = "”⁄— »Ì⁄"
.TextMatrix(0, 14) = "«·≈Ã„«·Ï"
.TextMatrix(0, 15) = "„” ‰œ „»Ì⁄« "

.MergeCells = flexMergeFree
.ColWidth(0) = 1300

.ColWidth(1) = 1300
.ColWidth(2) = 600
.ColWidth(3) = 600
.ColWidth(4) = 900
.ColWidth(5) = 900
.ColWidth(6) = 900
.ColWidth(7) = 1500
.ColWidth(8) = 2200
.ColWidth(9) = 800
.ColWidth(10) = 1000
.ColWidth(11) = 600
.ColWidth(12) = 600
.ColWidth(13) = 700
.ColWidth(14) = 700
.ColWidth(15) = 0
.RowHeight(0) = 700
.WordWrap = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
For nRow = 1 To .Rows - 1
    If .TextMatrix(nRow, 12) > 0 Then .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbYellow
Next nRow
.SubtotalPosition = flexSTBelow
.Subtotal flexSTSum, -1, 11, "#0", vbGreen, , True, " "
.Subtotal flexSTSum, -1, 12, "#0", vbGreen, , True, " "
End With
End Sub

Private Sub Form_Load()
' 250922181005
    openCon con
    
    Set GRID11.DataSource = DATA11
    
    If sDoc_no = "" Then Exit Sub
    
    If Not lIsBranchStore Then
        strConShop = LoadConStringshop
        If openConShop(conShop) <> "ok" Then
            MsgBox "·« ÌÊÃœ « ’«· »«·”Ì—›—"
            lConServerRet = False
        Else
            lConServerRet = True
            lConServer = True
            Inform "  „ «·≈ ’«· »«·”Ì—›— "
        End If
    Else
        lConServerRet = False
    End If
    If lConServerRet Then
        If retDOC_NO(sDoc_no, conShop) = "" Then lConServerRet = False
    End If
    If lConServerRet Then          ' «·„— Ã⁄ „‰ «·”Ì—›—
        conShop.Execute " DELETE FROM Q_RETBAL_ITEM WHERE DOC_NO = " & MyParn(sDoc_no)
        conShop.Execute " INSERT INTO Q_RETBAL_ITEM ( DOC_NO    , store, date, ITEM, QUANT, QUANT_RET     , PRICE, TOTAL, rate, price_sales , MAN, DOC_DISC,[S_OKAZ])  SELECT DOC_NO                     , store, date   , ITEM, QUANT   , 0             , PRICE , TOTAL , rate  , PRICE * ((100 - rate) / 100) AS PRICE_SALES, MAN  , DOC_DISC  , [S_OKAZ] FROM  QFILE6_20 WHERE (QUANT > 0) and   DOC_NO = " & MyParn(sDoc_no)
        conShop.Execute " INSERT INTO Q_RETBAL_ITEM ( DOC_NO    , store, date, ITEM, QUANT, QUANT_RET     , PRICE, TOTAL, rate, price_sales , MAN, DOC_DISC)  SELECT " & addstring(sDoc_no) & " , NULL , NULL   , ITEM, 0       , QUANT *-1     , 0     , 0     , 0     , 0                                          , NULL , NULL      FROM  QFILE6_20 WHERE QUANT  < 0  AND  DOC_NO IN (SELECT DOC_NO FROM  QFILE6_20H WHERE SALES_RET = " & MyParn(sDoc_no) & ")"
    Else
        con.Execute " DELETE FROM Q_RETBAL_ITEM WHERE DOC_NO = " & MyParn(sDoc_no)
        con.Execute " INSERT INTO Q_RETBAL_ITEM ( DOC_NO    , store, date, ITEM, QUANT, QUANT_RET     , PRICE, TOTAL, rate, price_sales , MAN, DOC_DISC,[S_OKAZ])  SELECT DOC_NO                     , store, date   , ITEM, QUANT   , 0             , PRICE , TOTAL , rate  , PRICE * ((100 - rate) / 100) AS PRICE_SALES, MAN  , DOC_DISC  ,[S_OKAZ] FROM  QFILE6_20 WHERE (QUANT > 0) and   DOC_NO = " & MyParn(sDoc_no)
        con.Execute " INSERT INTO Q_RETBAL_ITEM ( DOC_NO    , store, date, ITEM, QUANT, QUANT_RET     , PRICE, TOTAL, rate, price_sales , MAN, DOC_DISC)  SELECT " & addstring(sDoc_no) & " , NULL , NULL   , ITEM, 0       , QUANT *-1     , 0     , 0     , 0     , 0                                          , NULL , NULL      FROM  QFILE6_20 WHERE QUANT  < 0  AND  DOC_NO IN (SELECT DOC_NO FROM  QFILE6_20H WHERE SALES_RET = " & MyParn(sDoc_no) & ")"
    End If
    If myload Then
        Fixgrd
    End If
End Sub
Private Sub Fixgrd()
With grid1
.Cols = 14
.TextMatrix(0, 0) = "—ﬁ„ „” ‰œ"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "«·›—⁄"
.TextMatrix(0, 3) = "«·»«∆⁄"
.TextMatrix(0, 4) = "«·„ÊœÌ·"
.TextMatrix(0, 5) = "«·„’‰⁄"
.TextMatrix(0, 6) = "»«—ﬂÊœ"
.TextMatrix(0, 7) = "«·’‰›"
.TextMatrix(0, 8) = "„»Ì⁄« "
.TextMatrix(0, 9) = "„— Ã⁄"
.TextMatrix(0, 10) = "”⁄— »Ì⁄"
.TextMatrix(0, 11) = "«·”⁄—"
.TextMatrix(0, 12) = "„” ‰œ Œ’„"
.TextMatrix(0, 13) = "‰”»… Œ’„"
.MergeCells = flexMergeFree
.ColHidden(2) = True
.MergeCol(1) = True
.ColComboList(3) = StrList("SELECT CODE , DESCA FROM FILE6_25", con)
.ColWidth(0) = 0

.ColWidth(1) = 1300
.ColWidth(2) = 1500
.ColWidth(3) = 1000
.ColWidth(4) = 1300
.ColWidth(5) = 2000
.ColWidth(6) = 1000
.ColWidth(7) = 2800
.ColWidth(8) = 800
.ColWidth(9) = 800
.ColWidth(10) = 900
.ColWidth(11) = 900
.ColWidth(12) = 800
.ColWidth(13) = 600
.RowHeight(0) = 800
.WordWrap = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

End With
End Sub
Private Function myload() As Boolean
Dim cWhere As String, cString As String, aReT2 As Variant
Dim nRateDisc As Double
Dim loctable As New ADODB.Recordset
Dim loctable2 As ADODB.Recordset
If lConServerRet Then
    Set loctable2 = SALES_DOC_FIND(sDoc_no, conShop)
Else
    Set loctable2 = SALES_DOC_FIND(sDoc_no, con)
End If
If Not loctable2.EOF Then
    XDOC_DISC.Caption = loctable2!DOC_DISC & ""
    
    xDoc_No.Caption = loctable2!doc_no & ""
    xbranch.Caption = GetDesca("SELECT DESCA FROM STORE_BR WHERE CODE = " & MyParn(loctable2!STORE), con)
    xDate.Caption = Format(loctable2!Date, "DD-MM-YYYY")
    xNOTE_DISC.Caption = loctable2!NOTE_DISC & ""
    xNOTE_DISC2.Caption = GetDesca("select desca from file0_90h where doc_no = " & MyParn(XDOC_DISC.Caption), con)
'''''''    If xNOTE_DISC.Caption <> xNOTE_DISC2.Caption Then
'''''''        MsgBox " Ÿ»ÿ „” ‰œ «·⁄—÷ "
'''''''        If lConServerRet Then
'''''''            conShop.Execute " UPDATE FILE6_20H SET FILE6_20H.DOC_DISC = FILE0_90H.DOC_NO FROM FILE0_90H INNER JOIN FILE6_20H ON FILE6_20H.NOTE_DISC = FILE0_90H.DESCA WHERE FILE0_90H.DESCA = " & MyParn(xNOTE_DISC.Caption) & " AND FILE6_20H.DOC_NO = " & MyParn(SalesFrm.xdoc_ret.text)
'''''''        Else
'''''''            con.Execute " UPDATE FILE6_20H SET FILE6_20H.DOC_DISC = FILE0_90H.DOC_NO FROM FILE0_90H INNER JOIN FILE6_20H ON FILE6_20H.NOTE_DISC = FILE0_90H.DESCA WHERE FILE0_90H.DESCA = " & MyParn(xNOTE_DISC.Caption) & " AND FILE6_20H.DOC_NO = " & MyParn(SalesFrm.xdoc_ret.text)
'''''''        End If
'''''''        Unload Me
'''''''    End If
    lRetOnePecs2 = False
    If XDOC_DISC.Caption <> "" Then
        aRet = aGetDesca("select qty1 , qty2 FROM FILE0_90H WHERE DOC_NO = " & MyParn(XDOC_DISC.Caption), con)
        If UBound(aRet) > 0 Then
            If Val(aRet(1) & "") = 1 And Val(aRet(2) & "") = 0 Then
                lRetOnePecs = True
            Else
                lRetOnePecs = False
            End If
            If Val(aRet(1) & "") >= 2 And Val(aRet(2) & "") = 0 Then lRetOnePecs2 = True
        Else
            lRetOnePecs = True
        End If
    Else
        lRetOnePecs = True
    End If
    If DateDiff("D", DateValue(xDate.Caption), Date) > 30 And Not loctable2!ISRET Then
        MsgBox "  ›« Ê—… „»Ì⁄«  «ﬂ»— „‰ 30 ÌÊ„ - ·«Ì„”Õ »⁄„· „— Ã⁄  "
        Exit Function
    End If
End If
If lRetOnePecs Then xretdisc1.Caption = "«·”„«Õ »„— Ã⁄ ﬁÿ⁄… "
If lRetOnePecs2 Then xretdisc2.Caption = "«·”„«Õ »„— Ã⁄ ﬁÿ⁄… Œ«—Ã «·⁄—÷"

cmd_retall2.Visible = Not lRetOnePecs
SalesFrm.lDiscModelRet = Not lRetOnePecs
If Not lRetOnePecs Then
    MsgBox " ›« Ê—… «·„»Ì⁄«  ·Â« Œ’„ ⁄—÷ = ÌÃ» ⁄„· „— Ã⁄ ··⁄—÷ »«·ﬂ«„· "
End If
cStr1 = " SELECT  Q_RETBAL_ITEM.DOC_NO, MAX(Q_RETBAL_ITEM.date) AS DATE, MAX(STORE) AS STORE, Q_RETBAL_ITEM.ITEM,FACT.DESCA AS fact_Desca , FILE1_10.desca,MODELFACT, SUM(QUANT ) AS QUANT ,  SUM(QUANT_RET) AS QUANT_RET ,  MAX(Q_RETBAL_ITEM.PRICE) AS PRICE, MAX(Q_RETBAL_ITEM.PRICE_SALES) AS PRICE_SALES, MAX(Q_RETBAL_ITEM.MAN) AS MAN, MAX(Q_RETBAL_ITEM.DOC_DISC) AS DOC_DISC , MAX(S_OKAZ) AS S_OKAZ  FROM  Q_RETBAL_ITEM INNER JOIN FILE1_10 ON Q_RETBAL_ITEM.ITEM = FILE1_10.ITEM INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE WHERE Q_RETBAL_ITEM.DOC_NO = " & MyParn(sDoc_no) & _
            " GROUP BY Q_RETBAL_ITEM.DOC_NO,MODELFACT, Q_RETBAL_ITEM.ITEM, FILE1_10.desca,FACT.DESCA  "
If lConServerRet Then
    loctable.Open cStr1, conShop, adOpenStatic, adLockUnspecified, adCmdText
Else
    loctable.Open cStr1, con, adOpenStatic, adLockUnspecified, adCmdText
End If
With grid1
.Rows = 1
.Cols = 14
.ColComboList(3) = StrList("SELECT CODE , DESCA FROM STORE_BR", con)
If loctable.RecordCount <> 0 Then loctable.MoveFirst
Do Until loctable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = Mid(loctable!doc_no, 7, 4)
    .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "dd-mm-yyyy")
    .TextMatrix(.Rows - 1, 2) = loctable!STORE
    If loctable!STORE <> cBranchStore Then
        .TextMatrix(.Rows - 1, 3) = "000"
    Else
        .TextMatrix(.Rows - 1, 3) = loctable!MAN
    End If
    .TextMatrix(.Rows - 1, 4) = DelZero(loctable!modelfact & "")
    .TextMatrix(.Rows - 1, 5) = loctable!fact_Desca & ""
    .TextMatrix(.Rows - 1, 6) = loctable!Item & ""
    .TextMatrix(.Rows - 1, 7) = loctable!DESCA & ""
    .TextMatrix(.Rows - 1, 8) = loctable!Quant
    .TextMatrix(.Rows - 1, 9) = loctable!Quant_RET
    .TextMatrix(.Rows - 1, 10) = loctable!PRICE_SALES
    .TextMatrix(.Rows - 1, 11) = loctable!price & ""
    .TextMatrix(.Rows - 1, 12) = loctable!DOC_DISC & ""
    .TextMatrix(.Rows - 1, 13) = loctable!S_OKAZ & ""
    
    If .TextMatrix(.Rows - 1, 9) = .TextMatrix(.Rows - 1, 8) Then .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
    loctable.MoveNext
Loop
End With
Fixgrd
myload = True
End Function
Private Sub grid1_DblClick()
If xDoc_No.Caption = "" Then Exit Sub
If Not lRetOnePecs And Not lRetOnePecs2 Then Exit Sub
Dim nQtyRet As Double, nQTY As Double
With grid1
    If .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = vbRed Then Exit Sub
    If lRetOnePecs2 And Val(.TextMatrix(.Row, 13)) > 0 Then Exit Sub
    nQTY = (.ValueMatrix(.Row, 8) - .ValueMatrix(.Row, 9))
    If nQTY > 0 Then
        nQtyRet = nQTY
        If nQTY > 1 Then nQtyRet = Val(InputBox(" ⁄œœ „— Ã⁄ ", , nQtyRet) & "")
        If nQtyRet > nQTY Then
            MsgBox "  ⁄œœ „— Ã⁄  €Ì— ’ÕÌÕ"
            Exit Sub
        End If
        If lConServerRet And .TextMatrix(.Row, 2) <> cBranchStore Then
            If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & Val(.TextMatrix(.Row, 6)), con) = "" Then
                sendItems Val(.TextMatrix(.Row, 6))
            End If
        End If
        
        SalesFrm.grid1.AddItem ""
        SalesFrm.XSALES_RET.Caption = sDoc_no
        SalesFrm.xMan.BoundText = .TextMatrix(.Row, 3)
        .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = vbRed
        
        SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 10) = nQtyRet * -1
        SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 1) = .TextMatrix(.Row, 6)
        SalesFrm.GrdDesc .TextMatrix(.Row, 6), SalesFrm.grid1.Rows - 1
        SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 16) = .TextMatrix(.Row, 3)
        SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 15) = .TextMatrix(.Row, 3)
        SalesFrm.grid1.TextMatrix(SalesFrm.grid1.Rows - 1, 11) = Val(.TextMatrix(.Row, 10))
        SalesFrm.XSALES_RET.Caption = sDoc_no
        SalesFrm.grid1_AfterEdit SalesFrm.grid1.Rows - 1, 1
    End If
End With
End Sub
Private Function sendItems(pItem) As Long
Dim cFile As String
cFile = "FILE1_10"
cString = "SELECT * FROM FILE1_10 WHERE  ITEM  = " & Val(pItem)
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    con.Execute " DELETE FROM FILE1_10 where item = " & loctable!Item
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!MODEL))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable!MODELNO))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "FACT", addstring(loctable!Fact))
    aInsert = AddFlag(aInsert, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "MODELFACT", addstring(loctable!modelfact))
    aInsert = AddFlag(aInsert, "MODELFACT0", addstring(loctable!modelfact0))
    aInsert = AddFlag(aInsert, "SUPP", addstring(loctable!SUPP))
    aInsert = AddFlag(aInsert, "OKAZ", Val(loctable!okaz & ""))
    aInsert = AddFlag(aInsert, "OKAZ_2", Val(loctable!okaz_2 & ""))
    aInsert = AddFlag(aInsert, "RATE", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "[SECTION]", addvalue(loctable!Section))
    aInsert = AddFlag(aInsert, "[ITEM]", addvalue(loctable!Item))
    aInsert = AddFlag(aInsert, "[COST]", Val(loctable!cost & ""))
    If cBranch > "60" Then
        aInsert = AddFlag(aInsert, "[COSTITEM]", Val(loctable!costITEM_FR & ""))
    Else
        aInsert = AddFlag(aInsert, "[COSTITEM]", Val(loctable!costITEM & ""))
    End If
    aInsert = AddFlag(aInsert, "[COST2]", Val(loctable!cost2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE]", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "[PRICE2]", Val(loctable!PRICE2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE_2]", Val(loctable!PRICE_2 & ""))
    aInsert = AddFlag(aInsert, "[SCAL]", addstring(loctable!scal))
    aInsert = AddFlag(aInsert, "[C_SCAL]", addvalue(loctable!C_SCAL))
    aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable!color))
    aInsert = AddFlag(aInsert, "[C_COLOR]", addvalue(loctable!c_Color))
    aInsert = AddFlag(aInsert, "[BARCODE]", addstring(loctable!BARCODE))
    aInsert = AddFlag(aInsert, "[BARCODE2]", addstring(loctable!BARCODE2))
    aInsert = AddFlag(aInsert, "[BARCODE13]", addstring(loctable!BARCODE13))
    aInsert = AddFlag(aInsert, "[ISOKAZITEM]", IIf(loctable!ISOKAZITEM, 1, 0))
    aInsert = AddFlag(aInsert, "[MOSM2]", addstring(loctable!MOSM2))
    aInsert = AddFlag(aInsert, "[ISNOITEM]", IIf(loctable!ISNOITEM, 1, 0))
    aInsert = AddFlag(aInsert, "[ISNODEM]", IIf(loctable!ISNODEM, 1, 0))
    con.Execute addInsert(aInsert, "FILE1_10"), nAffect
    Inform "  „ ≈÷«›… »«—ﬂÊœ " & loctable!Item & "  " & loctable!DESCA
    loctable.MoveNext
Loop
End Function
Private Sub myLoadPicture(pModel As String)
On Error Resume Next
    If Not lServerPict Then Exit Sub
    ImgX1.Images.Clear
    If pModel = "" Then Exit Sub
    If conPict Is Nothing Then Exit Sub
    Dim loctable As Recordset
    Dim bytes() As Byte, nSize As Long
    ImgX1.Images.Clear
    Set loctable = myRecordSet("select PICT,size from PICT where MODELNO = " & MyParn(pModel), conPict)
    If loctable.EOF Then Exit Sub
    nSize = CLng(loctable("size").Value)
    bytes = loctable("PICT").GetChunk(nSize)
    ImgX1.Import.FromMemoryFile bytes, ixmfJPG
    Err.Clear
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
    If OldRow <> NewRow Then
        myLoadPicture GetDesca("SELECT MODELNO FROM FILE1_10 WHERE ITEM = " & Val(.TextMatrix(NewRow, 6)), con)
    End If
End With
End Sub


