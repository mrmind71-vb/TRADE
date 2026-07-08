VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form salesdtlfrm 
   Caption         =   "»Ì«‰«  ›« Ê—…"
   ClientHeight    =   6045
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   13785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   13785
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "«·»«∆⁄"
      Height          =   735
      Left            =   990
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   45
      Width           =   12705
      Begin VB.Label xMan 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   12435
      End
   End
   Begin VB.CommandButton CmdExit 
      CausesValidation=   0   'False
      Height          =   510
      Left            =   90
      MaskColor       =   &H00FFFFFF&
      Picture         =   "salesDtl.frx":0000
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5490
      UseMaskColor    =   -1  'True
      Width           =   1545
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4605
      Left            =   90
      TabIndex        =   0
      Top             =   810
      Width           =   13605
      _cx             =   23998
      _cy             =   8123
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
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1665
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   5400
      Width           =   12030
      Begin VB.Label Label5 
         Caption         =   "«·⁄œœ :"
         Height          =   240
         Left            =   11250
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   690
      End
      Begin VB.Label xTotalquant 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   9630
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label xtotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label xDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label xTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label3 
         Caption         =   "’«›Ì «·›« Ê—… :"
         Height          =   240
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "≈Ã„«·Ì «·Œ’„ :"
         Height          =   240
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "≈Ã„«·Ì «·›« Ê—… :"
         Height          =   240
         Left            =   7785
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   1365
      End
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
Attribute VB_Name = "salesdtlfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cDoc_no
Dim con As New adodb.Connection
Dim cList As String

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
openCon con
Set grid1.DataSource = DATA3
DATA3.ConnectionString = strCon
cList = StrList("SELECT CODE, DESCA FROM FACT")
myload
End Sub
Private Sub myload()
With grid1
    cField1 = "case when file6_20.Discount = 0 then Null else file6_20.Discount end "
'                           0               1               2           3           4               5                       6               7               8           9               10              11              12                  13              14
    cString = "SELECT FILE6_20.ROW, FILE6_20.ITEM, FILE1_10.MOSM, FILE1_10.FACT, FILE1_10.SUPP, FILE1_10.MODELFACT0, FILE1_10.DESCA, FILE1_10.SCAL , FILE1_10.COLOR , FILE1_10.PRICE, file6_20.Quant, file6_20.Price, file6_20.DISCOUNT,FILE6_20.TOTAL" & _
          " FROM FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM"
    If cDoc_no <> "" Then
        cString = cString & turn(cString) & " FILE6_20.DOC_NO = " & MyParn(cDoc_no)
    End If
    cString = cString & " order by  FILE6_20.ROW"
    DATA3.RecordSource = cString
    DATA3.Refresh
End With
calcTotals
fixgrd
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For I = 1 To grid1.Rows - 1
    grid1.TextMatrix(I, 0) = I
Next
End Sub
Private Sub fixgrd()
With grid1
'       0        1           2       3           4           5           6           7       8               9           10          11           12        13              14
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "”⁄— „” Â·ﬂ|" & "⁄œœ|" & "”⁄— «·»Ì⁄|" & "‰”»… Œ’„|" & "«·≈Ã„«·Ï|" & "«·»«∆⁄"
.RowHeight(0) = 1000
.WordWrap = True
.ColWidth(0) = 400
.ColWidth(1) = 1000
.ColWidth(2) = 0
.ColWidth(3) = 1500
.ColWidth(4) = 0
.ColWidth(5) = 1000
.ColWidth(6) = 3000
.ColWidth(7) = 700
.ColWidth(8) = 800
.ColWidth(9) = 800
.ColWidth(10) = 800
.ColWidth(11) = 800
.ColWidth(12) = 800
.ColWidth(13) = 800
.ColFormat(12) = "#.##%"
.ColHidden(.Cols - 1) = True

.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
.ShowCell .Rows - 1, 0
.ColComboList(3) = cList
End With
End Sub
Private Sub calcTotals()
Dim nTotalItem As Single, nTotalQuant As Single
With grid1
For I = 1 To grid1.Rows - 1
    nTotalQuant = nTotalQuant + Val(.TextMatrix(I, 10))
    nTotalItem = nTotalItem + (Val(.TextMatrix(I, 10)) * Val(.TextMatrix(I, 9)))
Next

cString = "Select discount,file6_25.DESCA from file6_20h inner join file6_25 on file6_20h.man = file6_25.code "
If cDoc_no <> "" Then
    cString = cString & turn(cString) & " doc_no = " & MyParn(cDoc_no)
End If
aret = aGetDesca(cString)
xTotalItem.Caption = Format(nTotalItem, "Fixed")
xTotalquant.Caption = nTotalQuant
xtotal.Caption = Format(nTotalItem - Val(xDiscount.Caption), "Fixed")
If UBound(aret) <> 0 Then
    xDiscount.Caption = aret(1) & ""
    xMan.Caption = aret(2) & ""
End If
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set salesdtlfrm = Nothing
End Sub

