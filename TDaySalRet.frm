VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form TDaySalRet 
   ClientHeight    =   8520
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
   ScaleHeight     =   8520
   ScaleWidth      =   15240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "ÇáÝÊÑÉ"
      Height          =   600
      Left            =   6885
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   45
      Width           =   8250
      Begin VB.CommandButton cmdGo 
         Caption         =   "O  K"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   135
         Width           =   1200
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   5355
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   180
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ãä ÊÇÑíÎ :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   7290
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   225
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ÍÊì ÊÇÑíÎ :"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   3495
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Width           =   885
      End
   End
   Begin MSComCtl2.MonthView xDate 
      Height          =   2370
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   6470346
      Appearance      =   1
      StartOfWeek     =   142278663
      TitleBackColor  =   3778241
      TrailingForeColor=   3778241
      CurrentDate     =   40954
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   810
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
      Top             =   405
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
   Begin VSFlex7Ctl.VSFlexGrid gridbox 
      Bindings        =   "TDaySalRet.frx":0000
      Height          =   5880
      Left            =   90
      TabIndex        =   7
      Top             =   2475
      Width           =   15045
      _cx             =   26538
      _cy             =   10372
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   6470346
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12056827
      ForeColorSel    =   4194368
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
   Begin Threed.SSCommand PRINT 
      CausesValidation=   0   'False
      Height          =   555
      Left            =   4950
      TabIndex        =   8
      Top             =   90
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   979
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "TDaySalRet.frx":0014
      Caption         =   "ØÈÇÚÉ "
      Alignment       =   1
      PictureAlignment=   3
   End
End
Attribute VB_Name = "TDaySalRet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub CmdGo_Click()
    myload3
End Sub
Private Sub Form_Load()
    openCon con
    xDate.Visible = True
    
    Set GridBox.DataSource = data3
    data3.ConnectionString = strCon
    
    xDate.Visible = bopt3
    Frame1.Visible = bopt3
    
    GridBox.Rows = 1
    Fixgrd3
    xDate.Value = Date
    xDate1.Text = xDate.Value
    xdate2.Text = xDate.Value
    
    myload3

End Sub
Private Sub Fixgrd3()
With GridBox
.FixedCols = 0
.FixedRows = 1
.Cols = 14

.TextMatrix(0, 0) = "ÑÞã ÇáÈæä"
.TextMatrix(0, 1) = "ÇáÊÇÑíÎ"
.TextMatrix(0, 2) = "ÇáæÞÊ"
.TextMatrix(0, 3) = "ÇáÈÇÆÚ"
.TextMatrix(0, 4) = "ÇáÎÒäÉ"

.TextMatrix(0, 5) = "ÈÇÑßæÏ"
.TextMatrix(0, 6) = "ÇáãÕäÚ"
.TextMatrix(0, 7) = "ÇáãæÓã"
.TextMatrix(0, 8) = "ÑÞã ÇáãæÏíá"
.TextMatrix(0, 9) = "ÇáãæÏíá"
.TextMatrix(0, 10) = "ãÞÇÓ"
.TextMatrix(0, 11) = "áæä"

.TextMatrix(0, 12) = "ÚÏÏ"
.TextMatrix(0, 13) = "ÞíãÉ"

.MergeCells = flexMergeFree
.MergeCol(0) = True
.MergeCol(1) = True
.MergeCol(2) = True
.MergeCol(3) = True
.MergeCol(4) = True

.RowHeight(0) = 800
.WordWrap = True
.ColFormat(1) = "DD-MM-YYYY"
.ColFormat(2) = "SHORT TIME"
.ColWidth(0) = 1000
.ColWidth(1) = 1200
.ColWidth(2) = 800
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1500
.ColWidth(7) = 800
.ColWidth(8) = 800
.ColWidth(9) = 2200
.ColWidth(10) = 900
.ColWidth(11) = 900
.ColWidth(12) = 800

.ColWidth(13) = 1000


For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 1) = Format(.TextMatrix(nRow, 1), "DD-MM-YYYY")
    .TextMatrix(nRow, 2) = Format(.TextMatrix(nRow, 2), "SHORT TIME")
    .TextMatrix(nRow, 8) = DelZero(.TextMatrix(nRow, 8))
Next nRow
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTBelow

.Subtotal flexSTSum, 0, 12, "#0", vbYellow, vbRed, True, " "
.Subtotal flexSTSum, 0, 13, "#0.00", vbYellow, vbRed, True, " "
End With
End Sub
Private Sub PRINT_Click()
Load PrintGrd
PrintGrd.doprint GridBox, 1, , "ãÊÇÈÚÉ ÈæäÇÊ ÈåÇ ãÑÊÌÚÇÊ ááÝÊÑÉ ãä " & xDate1.Text & " ÍÊì " & xdate2.Text, , , , True, 8
PrintGrd.Show 1
End Sub
Private Sub xdate_DateClick(ByVal DateClicked As Date)
    xDate1.Text = Format(xDate.Value, "DD-MM-YYYY")
    xdate2.Text = Format(xDate.Value, "DD-MM-YYYY")
    myload3
End Sub
Sub myload3()
    cStr1 = "SELECT  DOC_NO,date, [TIME],  MANDESCA, BOXDESCA , ITEM, FACTDESCA, MOSM, MODELFACT, desca, COLOR , SCAL , QUANT, TOTAL FROM SALES_MODEL WHERE DATE >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSq(xdate2.Text)
    If cManBox <> "" Then cStr1 = cStr1 & " and box = " & MyParn(cManBox)
    cStr1 = cStr1 & " AND DOC_NO IN ( SELECT FILE6_20H.DOC_NO FROM FILE6_20 INNER JOIN FILE6_20H ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO WHERE QUANT < 0 AND  DATE >= " & DateSq(xDate1.Text) & " AND DATE <= " & DateSq(xdate2.Text) & " ) "
    cStr1 = cStr1 & " ORDER BY DOC_NO , QUANT DESC "
    data3.RecordSource = cStr1
    data3.Refresh
    Fixgrd3
End Sub


