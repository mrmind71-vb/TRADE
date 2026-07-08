VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form BarCodeFRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»«— ﬂÊœ"
   ClientHeight    =   10155
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   15270
   BeginProperty Font 
      Name            =   "Simplified Arabic"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picPrint 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   90
      ScaleHeight     =   945
      ScaleWidth      =   1125
      TabIndex        =   32
      Top             =   2070
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      Caption         =   "»œÊ‰ ”⁄— „” Â·ﬂ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4050
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   9450
      Width           =   1500
   End
   Begin VB.CheckBox XPRICE_P 
      Alignment       =   1  'Right Justify
      Caption         =   "”⁄— ÿ»«⁄… 2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5715
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   9450
      Width           =   1500
   End
   Begin VB.Frame Frame5 
      Height          =   915
      Left            =   7245
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   9135
      Width           =   7935
      Begin VB.OptionButton BARCODE 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ«—ÃÌ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   360
         Width           =   915
      End
      Begin VB.OptionButton BARCODE 
         Alignment       =   1  'Right Justify
         Caption         =   "BARCODE GS1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   1995
      End
      Begin VB.OptionButton BARCODE 
         Alignment       =   1  'Right Justify
         Caption         =   "BARCODE 13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   4020
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   360
         Width           =   1815
      End
      Begin VB.OptionButton BARCODE 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂÊœ «·’‰›"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   6300
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   360
         Value           =   -1  'True
         Width           =   1410
      End
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Caption         =   "≈Œ›«¡ «·„’‰⁄"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   -135
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Height          =   870
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   3435
      Begin VB.CheckBox XALL 
         Alignment       =   1  'Right Justify
         Caption         =   "≈Œ Ì«— «·ﬂ·"
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   1140
      End
      Begin VB.TextBox xmodelfact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1305
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   1230
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   " ÕœÌœ „ÊœÌ·"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   405
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Height          =   870
      Left            =   3645
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   6495
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H008080FF&
         Caption         =   "”⁄— «Êﬂ«“ÊÌ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   225
         Width           =   960
      End
      Begin Threed.SSCommand p_sal 
         Height          =   285
         Left            =   5535
         TabIndex        =   22
         Top             =   225
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "SALE"
      End
      Begin VB.TextBox XRATE 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFF80&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   4995
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   270
         Width           =   510
      End
      Begin VB.CommandButton cmd_barmodel 
         Caption         =   "—ﬁ„ «·„ÊœÌ·« "
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   1005
      End
      Begin VB.CommandButton cmdPrint96 
         Caption         =   "»«—ﬂÊœ 96"
         CausesValidation=   0   'False
         Height          =   585
         Left            =   4770
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   1125
         UseMaskColor    =   -1  'True
         Width           =   1500
      End
      Begin VB.CommandButton cmdPrint2 
         Caption         =   "Barcode ◊ 1"
         CausesValidation=   0   'False
         Height          =   585
         Left            =   2205
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton cmdPrint1 
         Caption         =   "Barcode ◊ 2"
         CausesValidation=   0   'False
         Height          =   585
         Left            =   3645
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin Threed.SSCommand P_DISCOUNT 
         Height          =   285
         Left            =   5535
         TabIndex        =   23
         Top             =   540
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "DISCOUNT"
      End
   End
   Begin VB.Frame Frame3 
      Height          =   870
      Left            =   10125
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   540
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelAll 
         CausesValidation=   0   'False
         Height          =   540
         Left            =   1305
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE.frx":246C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   540
         Left            =   2520
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE.frx":4D06
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3735
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BARCODE.frx":727F
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Õ›Ÿ"
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8295
      Left            =   180
      TabIndex        =   10
      Top             =   900
      Width           =   15045
      _cx             =   26538
      _cy             =   14631
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
      Rows            =   50
      Cols            =   12
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
      WordWrap        =   -1  'True
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   9270
      Width           =   3855
      Begin VB.TextBox xCol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   915
      End
      Begin VB.TextBox xRow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   270
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "«·⁄„Êœ :"
         Height          =   300
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "«·’›:"
         Height          =   390
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   270
         Width           =   540
      End
   End
   Begin MSAdodcLib.Adodc data1 
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
Attribute VB_Name = "BarCodeFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearchItem As New Search3
Dim con As New ADODB.Connection
Dim temptable As New ADODB.Recordset
Dim nCols As Double
Dim NROWS As Double
Private Sub CmdDelPrinted_Click()
If MsgBox("Õ–› „«  „  ÿ»«⁄ Â", vbOKCancel + vbDefaultButton2 + vbCritical, "Õ–› „«  „  ÿ»«⁄ Â") <> vbOK Then Exit Sub
Dim cError As String, cFilter As String
cFilter = "ADDPRINT.ISPRINT = 1"
cError = myDelete(cFilter)
If cError = "ok" Then Inform " „ Õ–› ﬂ· «·”Ã·«  «·„ÿ»Ê⁄… »‰Ã«Õ" Else MsgBox cError
End Sub

Private Sub cmd_barmodel_Click()
    aPrinterBar = AddFlag(Empty, "name", RetPrinterByType("2"))
    Doprint_model
    Set myform = Me
    CardPrintNew_2.Show 1
End Sub
Private Sub CmdDelAll_Click()
If MsgBox("Õ–› «·ﬂ· ", vbOKCancel + vbDefaultButton2 + vbCritical, "Õ–› ﬂ· «·”Ã·« ") <> vbOK Then Exit Sub
Dim cError As String
cError = myDelete
If cError = "ok" Then Inform " „ Õ–› ﬂ· «·”Ã·«  »‰Ã«Õ" Else MsgBox cError
End Sub
Private Sub cmdExit_Click()
    Unload Me
    Set BarCodeFRM = Nothing
End Sub
Private Sub cmdPrint1_Click()
    aPrinterBar = AddFlag(Empty, "name", RetPrinterByType("2"))
    Doprint_Zebra_1 2
    Set myform = Me
    CardPrintNew_2.Show 1
End Sub
Private Sub CmdPrint2_Click()
    aPrinterBar = AddFlag(Empty, "name", RetPrinterByType("2"))
    Doprint_Zebra_1
    Set myform = Me
    CardPrintNew_2.Show 1
End Sub
Private Sub cmdSave_Click()
    If myreplaceGrd Then Inform " „ «·Õ›Ÿ"
    myLoadGrd
End Sub
Private Sub cmduno_Click()
myLoadGrd
End Sub
Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Command3_Click()
'Grid1.LoadGrid App.Path & "\addPrint.grd", flexFileData
frmReturn.Show
End Sub
Private Sub Form_Load()
    openCon con
    Set grid1.DataSource = data1
    myLoadGrd
End Sub
Sub myLoadGrd()
Dim cString As String
With grid1
'                   0                      1           2              3               4               5              6          7              8                 9               10              11
cString = "Select ADDPRINT.ITEM," & _
          "FILE1_10.MODElFACT0," & _
          "fact.desca," & _
          "file1_10.desca," & _
          "FILE1_10.COLOR," & _
          "FILE1_10.SCAL," & _
          "ADDPRINT.QUANT," & _
          "FILE1_10.PRICE," & _
          "FILE1_10.PRICE_2," & _
          "FILE1_10.PRICE_P," & _
          "ADDPRINT.ISPRINT," & _
          "ADDPRINT.ID " & _
          " FROM FILE1_10 " & _
          " INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT" & _
          " INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM" & _
          " WHERE  BRANCH = " & MyParn(cBranch)

Set data1.Recordset = myRecordSet(cString, con)
myAddItem
fixGrd
End With
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal col As Long)
On Error GoTo myerror
Dim bNew As Boolean
If col = 0 Then GrdDesc Row
If Not validRow(Row) Then Exit Sub
With grid1
If Row = .Rows - 1 Then
    myAddItem
End If
CalcTotals

If myreplaceGrd(Row) Then
   bNew = grid1.TextMatrix(Row, .Cols - 1) = ""
End If
myLoadGrd
If bNew Then
    'grid1.Row = grid1.Rows - 1
    grid1.ShowCell grid1.Rows - 1, 1
End If
End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_DblClick()
With grid1
    If .col = 1 Then
        xmodelfact.text = .TextMatrix(.Row, 1)
        xmodelfact_KeyPress (13)
    End If
End With
End Sub

Private Sub grid1_EnterCell()
If grid1.col = 0 Or grid1.col = 6 Or grid1.col = 10 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub

Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, col
End If
End Sub

Private Sub grid1_Validate(Cancel As Boolean)
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
End If
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then .RemoveItem OldRow
End If
End With
End Sub
Private Function validRow(Row As Long, Optional bIgMsg As Boolean) As Boolean
With grid1
If (.TextMatrix(Row, 0)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub delcheck()
For i = 1 To grid1.Rows - 2
   If Val(grid1.TextMatrix(i, 8)) <> 0 Then
        grid1.RowHidden(i) = True
   End If
Next
myreplaceGrd
End Sub
Sub myProc()
Dim bNew As Boolean
bNew = grid1.Row = grid1.Rows - 1
grid1.TextMatrix(grid1.Row, 0) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
Grid1_AfterEdit grid1.Row, grid1.col
If Not bNew Then
    Unload oSearchItem
    CellPos 13, grid1.Row, grid1.col
Else
    grid1.Row = grid1.Rows - 1
    grid1.col = 0
End If
End Sub
Private Function myvalidRowCol() As Boolean
'If Val(xRow.Text) > SettingArray(cRows) Then
If Val(xRow.text) > NROWS Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Function
End If

If Val(xCol.text) > nCols Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Function
End If
myvalidRowCol = True
End Function
Private Sub fixGrd()
With grid1
'                   0                      1           2              3               4               5              6          7              8                 9               10              11
'cString = "Select ADDPRINT.ITEM,FILE1_10.MODElFACT,FILE1_10.MODEL,file1_10.desca,FILE1_10.COLOR,FILE1_10.SCAL,ADDPRINT.QUANT,FILE1_10.PRICE,FILE1_10.PRICE2,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,ADDPRINT.ID "
    .FormatString = "«·ﬂÊœ|" & "„ÊœÌ· «·„’‰⁄|" & "«·„’‰⁄|" & "≈”„ «·„ÊœÌ·|" & "«··Ê‰|" & "«·„ﬁ«”|" & "«·ﬂ„Ì…|" & "«·”⁄—|" & "”⁄— «·«Êﬂ«“ÌÊ‰|" & "”⁄— ÿ»«⁄…2|" & "«·ÿ»«⁄…|"
    .FixedCols = 0
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 1500
    .ColWidth(3) = 3800
    .ColWidth(5) = 1000
    .ColWidth(4) = 1300
    .ColWidth(6) = 800
    .ColWidth(7) = 800
    .ColWidth(8) = 900
    .ColWidth(9) = 1000
    .ColWidth(10) = 1000
    .ColDataType(10) = flexDTBoolean
    .ColHidden(11) = True
    .ColHidden(.Cols - 1) = True
    .RowHeight(0) = 700
'    For i = 0 To grid1.Cols - 1
'        If i <> 3 Then .ColAlignment(i) = flexAlignRightCenter
'    Next
    .Cell(flexcpFontSize, 1, 3, .Rows - 1, 3) = 10
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 3) = DelZero(.TextMatrix(nRow, 3))
    Next nRow
End With
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
'For i = 1 To grid1.Rows - 1
'    grid1.TextMatrix(i, 0) = i
'Next
End Sub
Private Sub myEdit(Row As Long, col As Long)
Dim bNew As Boolean
With grid1
If Not validRow(Row) Then Exit Sub

If Row = .Rows - 1 Then
    myAddItem
End If

If myreplaceGrd Then
    bNew = grid1.TextMatrix(Row, .Cols - 1) = ""
End If

If bNew Then
    myLoadGrd
    grid1.ShowCell grid1.Rows - 1, 2
    grid1.Select .Rows - 1, 2
End If
End With
End Sub
Private Function myreplaceGrd(Optional nRow As Long = -1) As Boolean
Dim aInsert(4, 1)
With grid1
For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
     aInsert(0, 0) = "doc_no"
     aInsert(0, 1) = addstring(grid1.TextMatrix(i, 9))
     
     aInsert(1, 0) = "item"
     aInsert(1, 1) = addstring(grid1.TextMatrix(i, 0))
             
     aInsert(2, 0) = "quant"
     aInsert(2, 1) = Val(.TextMatrix(i, 6))
    
     aInsert(3, 0) = "isprint"
     aInsert(3, 1) = IIf(Abs(Val(.TextMatrix(i, 10))) <> 0, 1, 0)
     
     aInsert(4, 0) = "BRANCH"
     aInsert(4, 1) = addstring(cBranch)
     
     If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
         con.Execute CreateInsert(aInsert, "ADDPRINT")
     Else
         con.Execute CreateUpdate(aInsert, "ADDPRINT", " where ID = " & grid1.TextMatrix(i, 11), Array(-1))
     End If
Next
End With
myreplaceGrd = True
End Function
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem (grid1.Row)
        CalcTotals
        MakeSerial grid1.Row
    End If
ElseIf KeyCode = 112 And grid1.col = 0 Then
    ItemsLookupAll Me, oSearchItem
ElseIf KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.col
End If
End Sub
Private Function RemoveItem(Row As Long) As Boolean
con.BeginTrans
On Error GoTo myerror
If grid1.TextMatrix(Row, grid1.Cols - 1) <> "" Then
    con.Execute "Delete  From ADDPRINT where id = " & grid1.TextMatrix(Row, grid1.Cols - 1)
End If
grid1.RemoveItem Row
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function myDelete(Optional cFilter As String) As String
Dim cString As String
cString = "DELETE FROM ADDPRINT WHERE BRANCH = " & MyParn(cBranch)
If cFilter <> "" Then cString = cString & " AND " & cFilter
con.BeginTrans
On Error GoTo myerror
con.Execute cString
con.CommitTrans
myDelete = "ok"
myLoadGrd
Exit Function
myerror:
    myDelete = Err.Description
    con.RollbackTrans
    Err.Clear
End Function
Private Sub CalcTotals()

End Sub

Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal col As Long, Cancel As Boolean)
With grid1
    If col = 0 Then
        If Not IsNumeric(.EditText) Then
            Cancel = True
            Exit Sub
        End If
        If Not validItem(.EditText, con) Then
            MsgBox ("ﬂÊœ «·„ÊœÌ· €Ì— ”·Ì„")
            Cancel = True
        End If
    End If
End With
End Sub

Private Sub P_DISCOUNT_Click()
    aPrinterBar = AddFlag(Empty, "name", RetPrinterByType("2"))
    Doprint_SAL "DISCOUNT"
    Set myform = Me
    CardPrintNew_2.Show 1

End Sub

Private Sub p_sal_Click()
    aPrinterBar = AddFlag(Empty, "name", RetPrinterByType("2"))
    Doprint_SAL "SALE"
    Set myform = Me
    CardPrintNew_2.Show 1
End Sub

Private Sub XALL_Click()
With grid1
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 10) = XALL.Value
    Next nRow
    cmdSave_Click
End With
End Sub

Private Sub xmodelfact_KeyPress(KeyAscii As Integer)
Dim nRow As Long
If KeyAscii = 13 Then
    With grid1
        If xmodelfact.text = "" Then
            For nRow = 1 To .Rows - 1
                .TextMatrix(nRow, 10) = 1
                myreplaceGrd nRow
            Next nRow
        Else
            For nRow = 1 To .Rows - 1
                If .TextMatrix(nRow, 1) = xmodelfact.text Then
                    .TextMatrix(nRow, 10) = 1
                End If
                myreplaceGrd nRow
            Next nRow
        End If
    End With
End If
End Sub
Private Sub GrdDesc(Row)
With grid1
grid1.TextMatrix(Row, 1) = ""
grid1.TextMatrix(Row, 2) = ""
grid1.TextMatrix(Row, 5) = ""
grid1.TextMatrix(Row, 3) = ""
grid1.TextMatrix(Row, 4) = ""
grid1.TextMatrix(Row, 7) = ""
grid1.TextMatrix(Row, 6) = ""
grid1.TextMatrix(Row, 9) = ""
grid1.TextMatrix(Row, 8) = ""
If Not IsNumeric(grid1.TextMatrix(Row, 0)) Then Exit Sub
Dim ItemTable As New ADODB.Recordset
Set ItemTable = ItemFind(grid1.TextMatrix(Row, 0), con)
If Not ItemTable.EOF Then
    grid1.TextMatrix(Row, 1) = ItemTable!MOSM
    grid1.TextMatrix(Row, 2) = ItemTable!Fact
    grid1.TextMatrix(Row, 5) = ItemTable!SUPP & ""
    grid1.TextMatrix(Row, 3) = ItemTable!modelfact0
    grid1.TextMatrix(Row, 4) = ItemTable!desca
    grid1.TextMatrix(Row, 7) = ItemTable!SCAL
    grid1.TextMatrix(Row, 6) = ItemTable!color
    grid1.TextMatrix(Row, 9) = Format(ItemTable!price, "#0.00")
    grid1.TextMatrix(Row, 8) = 1
End If
End With
End Sub
Private Sub cmdPrintBar_Click()
If Not myreplaceGrd Then
    MsgBox "Œÿ√ «À‰«¡«·Õ›Ÿ"
    Exit Sub
End If
If Val(xRow.text) > 3 Then
    MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
    Exit Sub
End If

If Val(xCol.text) > 6 Then
    MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
    Exit Sub
End If
DOPRINT
Set myform = Me
CardPrintNew.Show 1
End Sub
Private Function DOPRINT() As Boolean
nLastMargin = 0.1
SettingArray(cUpMargin) = MyMeasure(0.3)
SettingArray(cLeftMargin) = MyMeasure(0)
SettingArray(cCardWidth) = MyMeasure(3.4)
SettingArray(cCardHeight) = MyMeasure(6.79)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 3
SettingArray(cCols) = 6
SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select FILE1_10.MODELFACT,ADDPRINT.ITEM,FILE1_10.MODEL,FILE1_10.[SEC]" & _
          ",dbo.f_min_scal(FILE1_10.MODEL) as minScal,dbo.f_max_scal(FILE1_10.MODEL) as maxScal" & _
          ",FILE1_10.DESCA,FILE1_10.PRICE,FILE1_10.PRICE2,FILE1_10.COLOR,ADDPRINT.QUANT,FILE1_10.SCAL" & _
          " FROM FILE1_10 INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM  WHERE BRANCH = " & MyParn(cBranch)
cString = cString & " AND ADDPRINT.ISPRINT = 1"
cString = cString & "order by file1_10.model, file1_10.c_scal,file1_10.color "

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

Dim nCard As Long, nRow As Long, nCol As Long, nCols As Long, NROWS As Long
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    BLASTROW = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next

Do Until sourcetable.EOF
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        BLASTROW = (nRow = NROWS)
        
        temptable.AddNew
        temptable!Left = MyMeasure(3.5)
        temptable!Top = MyMeasure(0.4)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 9
        temptable!text = IIf(nCol = nCols, Null, "|")
        temptable!CardNo = nCard
        temptable.Update
        
        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(0.3)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 9
        temptable!text = "TRES CHIC"
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(0.7)
        temptable!Width = MyMeasure(2.3)
        temptable!Height = MyMeasure(0.6)
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 9
        temptable!text = sourcetable!Item
        temptable!isBarcode = True
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

'/*****
        temptable.AddNew
        temptable!Left = MyMeasure(3)
        temptable!Top = MyMeasure(0.85)
        temptable!Width = MyMeasure(0)
        temptable!Height = MyMeasure(0.6)
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 9
        temptable!text = sourcetable!SEC
        temptable!isBarcode = False
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(1.3)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!FontBold = True
        temptable!FontUnderline = True
        temptable!FontSize = 9
        temptable!text = sourcetable!MODELFACT
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(1.9)
        temptable!Top = MyMeasure(1.3)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!FontBold = True
        temptable!FontSize = 8
        temptable!text = sourcetable!Item
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(1.8)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 9
        temptable!text = sourcetable!MODEL
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(2.3)
        temptable!Width = MyMeasure(2.5)
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 6
        temptable!text = sourcetable!desca
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(2.7)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontUnderline = True
        temptable!FontSize = 10
        temptable!ForeColor = vbBlack
        temptable!text = sourcetable!SCAL
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(1.25)
        temptable!Top = MyMeasure(2.7)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontUnderline = False
        temptable!FontSize = 12
        temptable!text = "L.E." & Format(sourcetable!price, "Fixed")
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
    
       
        temptable.AddNew
        temptable!Left = MyMeasure(2.5)
        temptable!Top = MyMeasure(0.4)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!FontSize = 5
        temptable!text = sourcetable!MinScal & "-" & sourcetable!MaxScal
        temptable!isBarcode = False
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
       
       
       ' ----------
        temptable.AddNew
        temptable!Left = MyMeasure(3.5)
        temptable!Top = MyMeasure(3.7) - IIf(BLASTROW, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 9
        temptable!text = IIf(nCol = nCols, Null, "|")
        temptable!CardNo = nCard
        temptable.Update
        

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(3.7) - IIf(BLASTROW, MyMeasure(nLastMargin), 0)
        temptable!Width = MyMeasure(2.3)
        temptable!Height = MyMeasure(0.6)
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 9
        temptable!text = sourcetable!Item
        temptable!isBarcode = True
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(3)
        temptable!Top = MyMeasure(3.7) - IIf(BLASTROW, MyMeasure(nLastMargin), 0)
        temptable!Width = MyMeasure(0)
        temptable!Height = MyMeasure(0.6)
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 9
        temptable!text = sourcetable!SEC
        temptable!isBarcode = False
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(4.3) - IIf(BLASTROW, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!FontBold = True
        temptable!FontUnderline = True
        temptable!FontSize = 9
        temptable!text = sourcetable!MODELFACT
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(1.9)
        temptable!Top = MyMeasure(4.3) - IIf(BLASTROW, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!FontBold = True
        temptable!FontSize = 8
        temptable!text = sourcetable!Item
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(4.7) - IIf(BLASTROW, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 9
        temptable!text = sourcetable!MODEL
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(5.07) - IIf(BLASTROW, MyMeasure(nLastMargin), 0)
        temptable!Width = MyMeasure(2.5)
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 6
        temptable!text = sourcetable!desca
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(0.6)
        temptable!Top = MyMeasure(5.35) - IIf(BLASTROW, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontUnderline = True
        temptable!FontSize = 9
        temptable!ForeColor = vbBlack
        temptable!text = sourcetable!SCAL
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(1.25)
        temptable!Top = MyMeasure(5.35) - IIf(BLASTROW, MyMeasure(nLastMargin), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontUnderline = False
        temptable!FontSize = 12
        temptable!text = "L.E." & Format(sourcetable!price, "Fixed")
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
         
        temptable.AddNew
        temptable!Left = MyMeasure(1)
        temptable!Top = MyMeasure(6)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "times new Roman"
        temptable!FontBold = False
        temptable!FontUnderline = False
        temptable!FontSize = 11
        temptable!text = IIf(BLASTROW, Null, "------------------")
        temptable!CardNo = nCard
        temptable.Update
                     
        temptable.AddNew
        temptable!Left = MyMeasure(3.5)
        temptable!Top = MyMeasure(6)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!ForeColor = vbBlack
        temptable!FontSize = 10
        temptable!text = IIf(BLASTROW Or nCol = nCols, Null, "|")
        temptable!CardNo = nCard
        temptable.Update
    
    ' «··Ê‰
        temptable.AddNew
        temptable!Left = MyMeasure(2.9)
        temptable!Top = MyMeasure(2.2) - IIf(BLASTROW, MyMeasure(0.2), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!TextAngle = 900
        temptable!FontBold = True
        temptable!FontUnderline = False
        temptable!FontSize = 8
        If Not IsNull(sourcetable!color) Then
            temptable!text = IIf(sourcetable!color = "„ÊÕœ" Or sourcetable!color = "√·Ê«‰" Or sourcetable!color = "---", Null, sourcetable!color)
        End If
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
    
    ' «··Ê‰
        temptable.AddNew
        temptable!Left = MyMeasure(2.9)
        temptable!Top = MyMeasure(5.1) - IIf(BLASTROW, MyMeasure(0.1), 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!TextAngle = 900
        temptable!FontBold = True
        temptable!FontUnderline = False
        temptable!FontSize = 8
        If Not IsNull(sourcetable!color) Then
            temptable!text = IIf(sourcetable!color = "„ÊÕœ" Or sourcetable!color = "√·Ê«‰" Or sourcetable!color = "---", Null, sourcetable!color)
        End If
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
    Next
    sourcetable.MoveNext
Loop
Set sourcetable = Nothing
Set temptable = Nothing
DOPRINT = True
End Function
Private Sub myAddItem()
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 10) = -1
MakeSerial
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal col As Long)
KeyCode = 0
If col < grid1.Cols - 4 Then
    If col < 6 Then
        grid1.col = 6
    Else
        grid1.col = col + 1
    End If
ElseIf Row < grid1.Rows - 1 Then
    grid1.Row = grid1.Row + 1
    grid1.col = 0
End If
End Sub
Private Function doprint1() As Boolean
nLastMargin = 0.1
SettingArray(cUpMargin) = MyMeasure(0)
SettingArray(cLeftMargin) = MyMeasure(0.1)
SettingArray(cCardWidth) = MyMeasure(3.5)
SettingArray(cCardHeight) = MyMeasure(2.5)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 1
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select ADDPRINT.ITEM,FILE1_10.MOSM,fact.desca as descafact ,FILE1_10.MODElfaCT,file1_10.desca , FILE1_10.SUPP,FILE1_10.COLOR,FILE1_10.SCAL,ADDPRINT.QUANT,FILE1_10.PRICE,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,ADDPRINT.ID " & _
          " FROM (FILE1_10 inner join fact on fact.code = file1_10.fact ) INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM  WHERE BRANCH = " & MyParn(cBranch)
cString = cString & " AND addPrint.isPrint = 1 order by file1_10.mosm , file1_10.fact , file1_10.modelfact , file1_10.color , file1_10.c_scal "

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    BLASTROW = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next
'«‰ Â«¡

nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)
nFixLeft = MyMeasure(0.1)
nFixlast = MyMeasure(0.18)
nFixlastbar = MyMeasure(0.1)

With sourcetable
Do Until .EOF
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        
        nHeight = 0
        For nCount = 1 To 1
            BLASTROW = (nRow = NROWS And nCount = 2)
            
            temptable.AddNew
            temptable!Top = MyMeasure(0.1) + nHeight
            temptable!Left = MyMeasure(0.2) - nFixLeft
            temptable!Width = MyMeasure(2.5)
            temptable!Height = 0
            temptable!TextAlign = taCenterTop
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontSize = 12
            temptable!text = "ﬂÌœ“"
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(2.3) - nFixLeft
            temptable!Top = MyMeasure(0.1) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 12
            temptable!TextAlign = taRightTop
            temptable!text = sourcetable!Item
            temptable!CardNo = nCard
            temptable.Update
            
            If Check1.Value = 0 Then
                temptable.AddNew
                temptable!Left = MyMeasure(0.1) - nFixLeft
                temptable!Top = MyMeasure(2.5) + nHeight
                temptable!Width = MyMeasure(2)
                temptable!Height = 0
                temptable!FontName = "Arial"
                temptable!TextAngle = 900
                temptable!FontBold = True
                temptable!TextAlign = taCenterTop
                temptable!FontSize = 12
                temptable!text = sourcetable!descafact
                temptable!ForeColor = vbBlack
                temptable!CardNo = nCard
                temptable.Update
            End If
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.6) - nFixLeft
            temptable!Top = MyMeasure(0.67) + nHeight
            temptable!Width = MyMeasure(2)
            temptable!Height = MyMeasure(0.6)
            temptable!TextAlign = taRightTop
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 8
            temptable!isBarcode = True
            temptable!text = sourcetable!Item
            temptable!CardNo = nCard
            temptable.Update
           
            temptable.AddNew
            temptable!Left = MyMeasure(2.5) - nFixLeft
            temptable!Top = MyMeasure(0.68) + nHeight
            temptable!Width = MyMeasure(0.8)
            temptable!Height = MyMeasure(0.5)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            temptable!TextAlign = taRightTop
            temptable!text = sourcetable!MOSM
            temptable!CardNo = nCard
            temptable.Update
           
            temptable.AddNew
            temptable!Left = MyMeasure(2.5) - nFixLeft
            temptable!Top = MyMeasure(1.1) + nHeight
            temptable!Width = MyMeasure(0.8)
            temptable!Height = MyMeasure(0.5)
            temptable!TextAlign = taRightTop
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            temptable!text = sourcetable!SUPP
            temptable!CardNo = nCard
            temptable.Update
           
           
            temptable.AddNew
            temptable!Left = MyMeasure(0.7) - nFixLeft
            temptable!Top = MyMeasure(1.25) + nHeight
            temptable!Width = MyMeasure(2.5)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!ForeColor = vbBlack
            temptable!FontSize = 9
            temptable!text = sourcetable!desca
            temptable!TextAlign = taCenterTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(0.7) - nFixLeft
            temptable!Top = MyMeasure(1.55) + nHeight
            temptable!Width = MyMeasure(1.5)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!ForeColor = vbBlack
            temptable!FontSize = 8
            If sourcetable!color <> "„ÊÕœ" Then temptable!text = sourcetable!color
            temptable!TextAlign = taLeftTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(2) - nFixLeft
            temptable!Top = MyMeasure(1.55) + nHeight
            temptable!Width = MyMeasure(1.3)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!ForeColor = vbBlack
            temptable!FontSize = 8
            temptable!text = DelZero(sourcetable!MODELFACT)
            temptable!TextAlign = taRightTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(0.6) - nFixLeft
            temptable!Top = MyMeasure(1.8) + nHeight
            temptable!Width = MyMeasure(2)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontUnderline = True
            temptable!FontSize = 12
            temptable!TextAlign = taLeftTop
            temptable!text = "L.E." & Format(sourcetable!price, "#0.00")
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(2.3) - nFixLeft
            temptable!Top = MyMeasure(1.95) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!TextAlign = taRightTop
            temptable!FontSize = 8
            temptable!ForeColor = vbBlack
            
            temptable!text = sourcetable!SCAL
            temptable!CardNo = nCard
            temptable.Update
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
doprint1 = True
End Function
Private Function doprint2() As Boolean
nLastMargin = 0.1
SettingArray(cUpMargin) = MyMeasure(-0.3)
SettingArray(cLeftMargin) = MyMeasure(0.1)
SettingArray(cCardWidth) = MyMeasure(3.5)
SettingArray(cCardHeight) = MyMeasure(1.25)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 2
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select ADDPRINT.ITEM,FILE1_10.MOSM,fact.desca as descafact ,FILE1_10.MODElfaCT,file1_10.desca , FILE1_10.SUPP,FILE1_10.COLOR,FILE1_10.SCAL,ADDPRINT.QUANT,FILE1_10.PRICE,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,ADDPRINT.ID " & _
          " FROM (FILE1_10 inner join fact on fact.code = file1_10.fact ) INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM  WHERE BRANCH = " & MyParn(cBranch)
cString = cString & " AND addPrint.isPrint = 1 order by file1_10.mosm , file1_10.fact , file1_10.modelfact , file1_10.color , file1_10.c_scal "

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    BLASTROW = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next
'«‰ Â«¡

nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)
nFixLeft = MyMeasure(0.1)
nFixlast = MyMeasure(0.18)
nFixlastbar = MyMeasure(0.1)

With sourcetable
Do Until .EOF
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        
        nHeight = 0
        For nCount = 1 To 1
            BLASTROW = (nRow = NROWS And nCount = 2)
            If Check1.Value = 0 Then
                temptable.AddNew
                temptable!Top = MyMeasure(0.25) + nHeight
                temptable!Left = MyMeasure(0.2) - nFixLeft
                temptable!Width = MyMeasure(2)
                temptable!Height = 0
                temptable!FontName = "Arial"
                temptable!FontBold = False
                temptable!FontSize = 8
                temptable!TextAlign = taCenterTop
                temptable!text = sourcetable!descafact
                temptable!CardNo = nCard
                temptable.Update
            End If
            temptable.AddNew
            temptable!Left = MyMeasure(2.2) - nFixLeft
            temptable!Top = MyMeasure(0.3) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 8
            temptable!TextAlign = taRightTop
            temptable!text = sourcetable!Item
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.2) - nFixLeft
            temptable!Top = MyMeasure(0.57) + nHeight
            temptable!Width = MyMeasure(2)
            temptable!Height = MyMeasure(0.3)
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontSize = 8
            temptable!isBarcode = True
            temptable!TextAlign = taRightTop
            temptable!text = sourcetable!Item
            temptable!CardNo = nCard
            temptable.Update
           
            temptable.AddNew
            temptable!Left = MyMeasure(2.2) - nFixLeft
            temptable!Top = MyMeasure(0.6) + nHeight
            temptable!Width = MyMeasure(0.5)
            temptable!Height = MyMeasure(0)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 6
            temptable!TextAlign = taRightTop
            temptable!text = sourcetable!MOSM
            temptable!CardNo = nCard
            temptable.Update
           
            temptable.AddNew
            temptable!Left = MyMeasure(2.75) - nFixLeft
            temptable!Top = MyMeasure(0.6) + nHeight
            temptable!Width = MyMeasure(0.4)
            temptable!Height = MyMeasure(0)
            temptable!TextAlign = taRightTop
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 6
            temptable!text = sourcetable!SUPP
            temptable!CardNo = nCard
            temptable.Update
           
           
            temptable.AddNew
            temptable!Left = MyMeasure(0.2) - nFixLeft
            temptable!Top = MyMeasure(0.85) + nHeight
            temptable!Width = MyMeasure(2.3)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!ForeColor = False
            temptable!FontSize = 7
            temptable!text = sourcetable!desca
            temptable!TextAlign = taCenterTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(2.2) - nFixLeft
            temptable!Top = MyMeasure(0.85) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!ForeColor = vbBlack
            temptable!FontSize = 8
            temptable!text = DelZero(sourcetable!MODELFACT)
            temptable!TextAlign = taRightTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
    
            temptable.AddNew
            temptable!Left = MyMeasure(0.2) - nFixLeft
            temptable!Top = MyMeasure(1.1) + nHeight
            temptable!Width = MyMeasure(2)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!FontUnderline = True
            temptable!FontSize = 8
            temptable!TextAlign = taLeftTop
            temptable!text = "L.E." & Format(sourcetable!price, "#0.00")
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(2.2) - nFixLeft
            temptable!Top = MyMeasure(1.12) + nHeight
            temptable!Width = MyMeasure(1)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontUnderline = False
            temptable!TextAlign = taRightTop
            temptable!FontSize = 7
            temptable!ForeColor = vbBlack
            
            temptable!text = sourcetable!SCAL
            temptable!CardNo = nCard
            temptable.Update
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
doprint2 = True
End Function
Private Function DoprintOkazStr() As Boolean
nLastMargin = 0.1
SettingArray(cUpMargin) = MyMeasure(0.1)
SettingArray(cLeftMargin) = MyMeasure(2.8)
SettingArray(cCardWidth) = MyMeasure(2.6)
SettingArray(cCardHeight) = MyMeasure(3.35)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 4
SettingArray(cCols) = 6
SettingArray(cPageWidth) = MyMeasure(16.8)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select FILE1_10.MODELFACT,ADDPRINT.ITEM,FILE1_10.MODEL,FILE1_10.[SEC]" & _
          ",dbo.f_min_scal(FILE1_10.MODEL) as minScal,dbo.f_max_scal(FILE1_10.MODEL) as maxScal" & _
          ",FILE1_10.DESCA,FILE1_10.PRICE,FILE1_10.PRICE2,FILE1_10.COLOR,ADDPRINT.QUANT,FILE1_10.SCAL" & _
          " FROM FILE1_10 INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM  WHERE BRANCH = " & MyParn(cBranch)
cString = cString & " AND ADDPRINT.ISPRINT = 1"
cString = cString & "order by file1_10.model, file1_10.c_scal,file1_10.color "

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

Dim nCard As Long, nRow As Long, nCol As Long, nCols As Long, NROWS As Long
nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ
nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    BLASTROW = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next

Do Until sourcetable.EOF
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        BLASTROW = (nRow = NROWS)
        blastcol = (nCol = nCols)
        
        nHeight = 0
        nLast = 0
        nLastCol = 0
        
        For nCount = 1 To 2
        temptable.AddNew
        temptable!Left = MyMeasure(0.15)
        temptable!Top = MyMeasure(0.3) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Book Antiqua"
        temptable!FontBold = False
        temptable!FontItalic = False
        temptable!FontSize = 8
        temptable!text = sourcetable!MODEL
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

      
        temptable.AddNew
        temptable!Left = MyMeasure(1)
        temptable!Top = MyMeasure(0.65) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontSize = 8
        temptable!text = sourcetable!Item
        temptable!CardNo = nCard
        temptable.Update

    
        temptable.AddNew
        temptable!Left = MyMeasure(0.2)
        temptable!Top = MyMeasure(0.65) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "Arial"
        temptable!FontBold = False
        temptable!FontSize = 8
        temptable!FontUnderline = True
        temptable!text = sourcetable!MODELFACT
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update

        temptable.AddNew
        temptable!Left = MyMeasure(1)
        temptable!Top = MyMeasure(1) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = True
        temptable!FontUnderline = False
        temptable!FontSize = 11
        temptable!text = Format(sourcetable!PRICE2, "Fixed")
        temptable!ForeColor = vbBlack
        temptable!CardNo = nCard
        temptable.Update
        
        temptable.AddNew
        temptable!Left = MyMeasure(0.2)
        temptable!Top = MyMeasure(1) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0)
        temptable!Width = 0
        temptable!Height = 0
        temptable!FontName = "arial"
        temptable!FontBold = False
        temptable!FontUnderline = True
        temptable!FontSize = 9
        temptable!ForeColor = vbBlack
        temptable!text = sourcetable!SCAL
        temptable!CardNo = nCard
        temptable.Update
        nHeight = SettingArray(cCardHeight) / 2
        Next nCount
    Next i
    sourcetable.MoveNext
Loop
Set sourcetable = Nothing
Set temptable = Nothing
DoprintOkazStr = True
End Function


Private Function Doprint_Zebra() As Boolean
nLastMargin = 0

SettingArray(cUpMargin) = MyMeasure(0.24)
SettingArray(cLeftMargin) = MyMeasure(0)
SettingArray(cCardWidth) = MyMeasure(4.2)
SettingArray(cCardHeight) = MyMeasure(3)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1

SettingArray(cRows) = 1
SettingArray(cCols) = 2


SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable


cString = " Select ADDPRINT.ITEM,FILE1_10.MOSM,FILE1_10.SUPP ,fact.desca as descafact ,FILE1_10.MODElfaCT,file1_10.desca , FILE1_10.code,FILE1_10.COLOR,FILE1_10.SCAL,ADDPRINT.QUANT,FILE1_10.PRICE2,FILE1_10.PRICE_2,FILE1_10.PRICE,ADDPRINT.DOC_NO,ADDPRINT.ISPRINT,ADDPRINT.ID , FILE1_10.BARCODE13 , FILE1_10.BARCODE_GS1 , PRICE_P " & _
          " FROM (FILE1_10 left join fact on fact.code = file1_10.fact ) INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM  WHERE BRANCH = " & MyParn(cBranch)
cString = cString & " AND addPrint.isPrint = 1 order by file1_10.mosm , file1_10.fact , file1_10.modelfact , file1_10.color , file1_10.c_scal "

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    BLASTROW = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next
'«‰ Â«¡

nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)
nFixLeft = MyMeasure(0.1)
nFixlast = MyMeasure(0.18)
nFixlastbar = MyMeasure(0.1)

With sourcetable
Do Until .EOF
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        
        nHeight = 0
        For nCount = 1 To 1
            BLASTROW = (nRow = NROWS And nCount = 2)
            
            temptable.AddNew
            temptable!Left = MyMeasure(2)
            temptable!Top = MyMeasure(0#) - IIf(BLASTROW, MyMeasure(0.2), 0)
            temptable!Width = MyMeasure(1.5)
            temptable!Height = MyMeasure(0.7)
            temptable!FontBold = True
            temptable!TextAlign = taCenterTop
            temptable!text = (App.Path & "\junior.jpg")
            temptable!isPhoto = True
            temptable!CardNo = nCard
            temptable.Update
            
            
            temptable.AddNew
            temptable!Top = MyMeasure(0.05) + nHeight
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(1.5)
            temptable!TextAlign = taCenterTop
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!FontSize = 12
            If BARCODE(0).Value <> 0 Then temptable!text = sourcetable!Item
            temptable!CardNo = nCard
            temptable.Update
            
            
            temptable.AddNew
            temptable!Top = MyMeasure(0.5) + nHeight
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(1.5)
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = False
            temptable!FontUnderline = False
            temptable!FontSize = 8
            temptable!TextAlign = taCenterTop
            temptable!text = DelZero(sourcetable!MODELFACT)
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.5) - IIf(blastcol, nLastCol, 0)
            temptable!Top = MyMeasure(0.9) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Width = MyMeasure(3)
            temptable!Height = MyMeasure(0.6)
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            temptable!TextAlign = taRightTop
            
            If BARCODE(0).Value <> 0 Then temptable!text = sourcetable!Item
            If BARCODE(1).Value <> 0 And Not IsNull(sourcetable!BARCODE13) Then temptable!text = sourcetable!BARCODE13
            If BARCODE(2).Value <> 0 And Not IsNull(sourcetable!BARCODE_GS1) Then temptable!text = sourcetable!BARCODE_GS1
            
            temptable!isBarcode = True
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update

            If BARCODE(0).Value <> 0 Or (BARCODE(1).Value <> 0 And Not IsNull(sourcetable!BARCODE13)) Or (BARCODE(2).Value <> 0 And Not IsNull(sourcetable!BARCODE_GS1)) Then
                temptable.AddNew
                temptable!Top = MyMeasure(1.4) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
                temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
                temptable!Width = MyMeasure(3.5)
                temptable!Height = 0
                temptable!FontName = "Simplified arabic"
                temptable!TextAlign = taCenterTop
                temptable!FontBold = False
                temptable!ForeColor = vbBlack
                temptable!FontSize = 8
                If BARCODE(1).Value <> 0 And Not IsNull(sourcetable!BARCODE13) Then temptable!text = sourcetable!BARCODE13
                If BARCODE(2).Value <> 0 And Not IsNull(sourcetable!BARCODE_GS1) Then temptable!text = sourcetable!BARCODE_GS1
                temptable!CardNo = nCard
                temptable.Update
            End If
            
            temptable.AddNew
            temptable!Top = MyMeasure(1.6) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(3.5)
            temptable!Height = 0
            temptable!FontName = "Simplified arabic"
            temptable!TextAlign = taCenterTop
            temptable!FontBold = False
            temptable!ForeColor = vbBlack
            temptable!FontSize = 8
            temptable!text = ArbString(sourcetable!desca)
            temptable!CardNo = nCard
            temptable.Update

            
            temptable.AddNew
            temptable!Top = MyMeasure(1.55) + nHeight
            temptable!Left = MyMeasure(3.5) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(0)
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = False
            temptable!FontUnderline = False
            temptable!FontSize = 8
            temptable!ForeColor = vbBlack
            temptable!text = sourcetable!MOSM
            temptable!CardNo = nCard
            temptable!TextAngle = 900
            temptable!TextAlign = taCenterTop
            temptable.Update
            
            temptable.AddNew
            temptable!Top = MyMeasure(1.95) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(2)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 7
            temptable!text = sourcetable!descafact
            temptable!TextAlign = taLeftTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
        
            temptable.AddNew
            temptable!Top = MyMeasure(1.95) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Left = MyMeasure(2.3) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(1.5)
            temptable!Height = 0
            temptable!TextAlign = taRightTop
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            If Not (sourcetable!color = "„ÊÕœ" Or sourcetable!color = "---" Or sourcetable!color = "√·Ê«‰") Then
                temptable!text = sourcetable!color
            End If
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
        

            temptable.AddNew
            temptable!Top = MyMeasure(2.2) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Left = MyMeasure(2.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(1.5)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!TextAlign = taCenterTop
            temptable!FontUnderline = False
            temptable!FontSize = 8
            temptable!ForeColor = vbBlack
            temptable!text = "Size : " & sourcetable!SCAL
            temptable!CardNo = nCard
            temptable.Update

'''''''''''''''''''
            If Check3.Value = 0 Then
            If Check2.Value = 0 Or sourcetable!PRICE_2 = 0 Or sourcetable!PRICE_2 > sourcetable!price Then
                temptable.AddNew
                temptable!Top = MyMeasure(2.2) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
                temptable!Left = MyMeasure(0.1) - IIf(blastcol, nLastCol, 0)
                temptable!Width = MyMeasure(2.5)
                temptable!Height = 0
                temptable!FontName = "arial"
                temptable!FontBold = True
                temptable!FontSize = 10
                temptable!TextAlign = taLeftTop
                If XPRICE_P.Value = 0 Or Val(sourcetable!PRICE_P & "") = 0 Then
                    temptable!text = "L.E." & Format(sourcetable!price, "Fixed")
                    temptable!FontUnderline = False
                Else
                    temptable!text = "L.E." & Format(sourcetable!PRICE_P, "Fixed")
                    temptable!FontUnderline = True
                End If
                temptable!ForeColor = vbBlack
                temptable!CardNo = nCard
                temptable.Update
            Else
                temptable.AddNew
                temptable!Top = MyMeasure(2.15) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
                temptable!Left = MyMeasure(0.1) - IIf(blastcol, nLastCol, 0)
                temptable!Width = MyMeasure(2.5)
                temptable!Height = 0
                temptable!FontName = "arial"
                temptable!FontBold = True
                temptable!FontUnderline = False
                temptable!FontSize = 10
                temptable!TextAlign = taLeftTop
                If XPRICE_P.Value = 0 Or Val(sourcetable!PRICE_P & "") = 0 Then
                    temptable!text = "L.E." & Format(sourcetable!price, "Fixed")
                Else
                    temptable!text = "L.E." & Format(sourcetable!PRICE_P, "Fixed")
                    temptable!FontUnderline = True
                End If
                temptable!ForeColor = vbBlack
                temptable!CardNo = nCard
                temptable.Update
            
                temptable.AddNew
                temptable!Top = MyMeasure(2.4) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
                temptable!Left = MyMeasure(0.1) - IIf(blastcol, nLastCol, 0)
                temptable!Width = MyMeasure(2.5)
                temptable!Height = 0
                temptable!FontName = "arial"
                temptable!FontBold = False
                temptable!FontUnderline = False
                temptable!Strikethrough = True
                temptable!FontSize = 8
                temptable!TextAlign = taLeftTop
                temptable!text = "L.E." & Format(sourcetable!price, "Fixed")
                temptable!ForeColor = vbBlack
                temptable!CardNo = nCard
                temptable.Update
            
            End If
            End If
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
End Function
Private Function Doprint_Zebra_1(Optional nCols As Long = 1) As Boolean
nLastMargin = 0

SettingArray(cUpMargin) = MyMeasure(0.2)
SettingArray(cLeftMargin) = MyMeasure(0)
SettingArray(cCardWidth) = MyMeasure(4)
SettingArray(cCardHeight) = MyMeasure(3)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = IIf(Val(xCol.text) = 0, 1, Val(xCol.text))

SettingArray(cRows) = 1
SettingArray(cCols) = nCols

SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable


cString = "Select ADDPRINT.ITEM," & _
          "FILE1_10.MOSM," & _
          "FILE1_10.SUPP," & _
          "fact.desca as descafact," & _
          "FILE1_10.MODElfaCT," & _
          "file1_10.desca," & _
          "FILE1_10.code," & _
          "FILE1_10.COLOR," & _
          "FILE1_10.SCAL," & _
          "ADDPRINT.QUANT," & _
          "FILE1_10.PRICE2," & _
          "FILE1_10.PRICE_2," & _
          "FILE1_10.PRICE," & _
          "ADDPRINT.DOC_NO," & _
          "ADDPRINT.ISPRINT," & _
          "ADDPRINT.ID," & _
          " FILE1_10.BARCODE," & _
          " FILE1_10.BARCODE13," & _
          " FILE1_10.BARCODE_GS1," & _
          " FILE1_10.PRICE_P " & _
          " FROM FILE1_10 " & _
          " left join fact on fact.code = file1_10.fact" & _
          " INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM" & _
          " WHERE BRANCH = " & MyParn(cBranch)

cString = cString & " AND addPrint.isPrint = 1" & _
                    " order by file1_10.mosm," & _
                    " file1_10.fact," & _
                    "file1_10.modelfact," & _
                    "file1_10.color," & _
                    "file1_10.c_scal "

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

'nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))

'«‰ Â«¡

nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)
nFixLeft = MyMeasure(0.1)
nFixlast = MyMeasure(0.18)
nFixlastbar = MyMeasure(0.1)

For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    BLASTROW = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next

Set picPrint.Picture = LoadPicture(App.Path & "\junior.jpg")
Dim nH As Long
GetNewSize picPrint.Width, picPrint.Height, MyMeasure(2), nH

With sourcetable
Do Until .EOF
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        
        nHeight = 0
        For nCount = 1 To 1
            BLASTROW = (nRow = NROWS And nCount = 2)
            
            temptable.AddNew
            temptable!Left = MyMeasure(1)
            temptable!Top = MyMeasure(0)
            temptable!Width = MyMeasure(2)
            temptable!Height = nH
            temptable!FontBold = True
            temptable!TextAlign = taCenterTop
            temptable!text = (App.Path & "\junior.jpg")
            temptable!isPhoto = True
            temptable!CardNo = nCard
            temptable.Update
            
            
            temptable.AddNew
            temptable!Top = MyMeasure(0.93) + nHeight
            temptable!Left = MyMeasure(0.5) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(3)
            temptable!TextAlign = taCenterTop
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!FontSize = 9
            'If BARCODE(0).Value Then temptable!text = "ITEM"
            'temptable!text = !Item
            If BARCODE(0).Value Then
                temptable!text = sourcetable!Item
            ElseIf BARCODE(1).Value And Not IsNull(sourcetable!BARCODE13) Then
                temptable!text = sourcetable!BARCODE13
            ElseIf BARCODE(2).Value And Not IsNull(sourcetable!BARCODE_GS1) Then
                temptable!text = sourcetable!BARCODE_GS1
            ElseIf BARCODE(3).Value And Not IsNull(sourcetable!BARCODE) Then
                temptable!text = sourcetable!BARCODE
            End If
            temptable!CardNo = nCard
            temptable.Update
            
            
            temptable.AddNew
            temptable!Top = MyMeasure(1.24) + nHeight
            temptable!Left = MyMeasure(0.5) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(3)
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = False
            temptable!FontUnderline = False
            temptable!FontSize = 8
            temptable!TextAlign = taCenterTop
            temptable!text = DelZero(sourcetable!MODELFACT)
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
            
            If BARCODE(0).Value Or (BARCODE(1).Value And Not IsNull(sourcetable!BARCODE13)) Or (BARCODE(2).Value And Not IsNull(sourcetable!BARCODE_GS1)) Or (BARCODE(3).Value And Not IsNull(sourcetable!BARCODE)) Then
                temptable.AddNew
                temptable!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
                temptable!Top = MyMeasure(0.45) + nHeight
                temptable!Width = MyMeasure(3.5)
                temptable!Height = MyMeasure(0.48)
                temptable!FontName = "arial"
                temptable!FontBold = False
                temptable!FontSize = 8
                temptable!TextAlign = taRightTop
                
                If BARCODE(0).Value Then
                    temptable!text = sourcetable!Item
                ElseIf BARCODE(1).Value And Not IsNull(sourcetable!BARCODE13) Then
                    temptable!text = sourcetable!BARCODE13
                ElseIf BARCODE(2).Value And Not IsNull(sourcetable!BARCODE_GS1) Then
                    temptable!text = sourcetable!BARCODE_GS1
                ElseIf BARCODE(3).Value And Not IsNull(sourcetable!BARCODE) Then
                    temptable!text = sourcetable!BARCODE
                End If
                temptable!isBarcode = True
                temptable!ForeColor = vbBlack
                temptable!CardNo = nCard
                temptable.Update
            End If
                    
            temptable.AddNew
            temptable!Top = MyMeasure(1.4) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(3.5)
            temptable!Height = 0
            temptable!FontName = "Simplified arabic"
            temptable!TextAlign = taCenterTop
            temptable!FontBold = False
            temptable!ForeColor = vbBlack
            temptable!FontSize = 8
            If BARCODE(1).Value <> 0 And Not IsNull(sourcetable!BARCODE13) Then temptable!text = sourcetable!BARCODE13
            If BARCODE(2).Value <> 0 And Not IsNull(sourcetable!BARCODE_GS1) Then temptable!text = sourcetable!BARCODE_GS1
            temptable!CardNo = nCard
            temptable.Update


            temptable.AddNew
            temptable!Top = MyMeasure(1.46) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0)
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(3.5)
            temptable!Height = 0
            temptable!FontName = "Simplified arabic"
            temptable!TextAlign = taCenterTop
            temptable!FontBold = False
            temptable!ForeColor = vbBlack
            temptable!FontSize = 8
            temptable!text = ArbString(sourcetable!desca)
            temptable!CardNo = nCard
            temptable.Update

            
            temptable.AddNew
            temptable!Top = MyMeasure(1.55) + nHeight
            temptable!Left = MyMeasure(3.5) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(0)
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = False
            temptable!FontUnderline = False
            temptable!FontSize = 8
            temptable!ForeColor = vbBlack
            'temptable!text = sourcetable!MOSM
            temptable!CardNo = nCard
            temptable!TextAngle = 900
            temptable!TextAlign = taCenterTop
            temptable.Update
            
            temptable.AddNew
            temptable!Top = MyMeasure(1.95) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(2)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 7
            'temptable!text = sourcetable!descafact
            temptable!TextAlign = taLeftTop
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
        
            temptable.AddNew
            temptable!Top = MyMeasure(1.95) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
            temptable!Left = MyMeasure(2)
            temptable!Width = MyMeasure(1.5)
            temptable!Height = 0
            temptable!TextAlign = taRightTop
            temptable!FontName = "arial"
            temptable!FontBold = False
            temptable!FontSize = 8
            If Not (sourcetable!color = "„ÊÕœ" Or sourcetable!color = "---" Or sourcetable!color = "√·Ê«‰") Then
                temptable!text = sourcetable!color
            End If
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
        

            temptable.AddNew
            temptable!Top = MyMeasure(2.2) + nHeight
            temptable!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(2.5)
            temptable!Height = 0
            temptable!FontName = "arial"
            temptable!FontBold = True
            temptable!TextAlign = taRightTop
            temptable!FontUnderline = False
            temptable!FontSize = 9
            temptable!ForeColor = vbBlack
            temptable!text = "Size : " & sourcetable!SCAL
            temptable!CardNo = nCard
            temptable.Update

'''''''''''''''''''
            If Check3.Value = 0 Then
            If Check2.Value = 0 Or sourcetable!PRICE_2 = 0 Or sourcetable!PRICE_2 > sourcetable!price Then
                temptable.AddNew
                temptable!Top = MyMeasure(2.2) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0)
                temptable!Left = MyMeasure(0.1) - IIf(blastcol, nLastCol, 0)
                temptable!Width = MyMeasure(2.5)
                temptable!Height = 0
                temptable!FontName = "arial"
                temptable!FontBold = True
                temptable!FontSize = 10
                temptable!TextAlign = taLeftTop
                If XPRICE_P.Value = 0 Or Val(sourcetable!PRICE_P & "") = 0 Then
                    temptable!text = "L.E." & Format(sourcetable!price, "Fixed")
                    temptable!FontUnderline = False
                Else
                    temptable!text = "L.E." & Format(sourcetable!PRICE_P, "Fixed")
                    temptable!FontUnderline = True
                End If
                temptable!ForeColor = vbBlack
                temptable!CardNo = nCard
                temptable.Update
            Else
                temptable.AddNew
                temptable!Top = MyMeasure(2.2) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0)
                temptable!Left = MyMeasure(0.1) - IIf(blastcol, nLastCol, 0)
                temptable!Width = MyMeasure(2.5)
                temptable!Height = 0
                temptable!FontName = "arial"
                temptable!FontBold = True
                temptable!FontUnderline = False
                temptable!FontSize = 10
                temptable!TextAlign = taLeftTop
                temptable!text = "L.E." & Format(sourcetable!PRICE_2, "Fixed")
                temptable!ForeColor = vbBlack
                temptable!CardNo = nCard
                temptable.Update
            
                temptable.AddNew
                temptable!Top = MyMeasure(2.4) + nHeight - IIf(BLASTROW And nCount = 2, nLast, 0) - MyMeasure(0.1)
                temptable!Left = MyMeasure(0.1) - IIf(blastcol, nLastCol, 0)
                temptable!Width = MyMeasure(2.5)
                temptable!Height = 0
                temptable!FontName = "arial"
                temptable!FontBold = False
                temptable!Strikethrough = True
                temptable!FontSize = 8
                temptable!TextAlign = taLeftTop
                If XPRICE_P.Value = 0 Or Val(sourcetable!PRICE_P & "") = 0 Then
                    temptable!text = "L.E." & Format(sourcetable!price, "Fixed")
                    temptable!FontUnderline = False
                Else
                    temptable!text = "L.E." & Format(sourcetable!PRICE_P, "Fixed")
                    temptable!FontUnderline = True
                End If
                temptable!ForeColor = vbBlack
                temptable!CardNo = nCard
                temptable.Update
            End If
            End If
            
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
End Function
Private Function Doprint_model() As Boolean
nLastMargin = 0

SettingArray(cUpMargin) = MyMeasure(0.24)
SettingArray(cLeftMargin) = MyMeasure(0)
SettingArray(cCardWidth) = MyMeasure(4)
SettingArray(cCardHeight) = MyMeasure(3)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1

SettingArray(cRows) = 1
SettingArray(cCols) = 1

SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable


cString = "Select FILE1_10.MODEL,FILE1_10.MOSM,fact.desca as descafact ,FILE1_10.MODElfaCT,file1_10.desca , FILE1_10.COLOR,ADDPRINT.QUANT,ADDPRINT.ISPRINT,ADDPRINT.ID " & _
          " FROM (FILE1_10 inner join fact on fact.code = file1_10.fact ) INNER JOIN ADDPRINT ON FILE1_10.ITEM = ADDPRINT.ITEM  WHERE BRANCH = " & MyParn(cBranch)
cString = cString & " AND addPrint.isPrint = 1 order by file1_10.mosm , file1_10.fact , file1_10.modelfact , file1_10.color , file1_10.c_scal "

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

Dim sourcetable As New ADODB.Recordset
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Function
End If

nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    BLASTROW = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next
'«‰ Â«¡

nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)
nFixLeft = MyMeasure(0.1)
nFixlast = MyMeasure(0.18)
nFixlastbar = MyMeasure(0.1)

With sourcetable
Do Until .EOF
'************
    For i = 1 To sourcetable!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        
        nHeight = 0
        For nCount = 1 To 1
            BLASTROW = (nRow = NROWS And nCount = 2)
            
'           modelfact
            temptable.AddNew
            temptable!Top = MyMeasure(0#) + nHeight
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(3.5)
            temptable!TextAlign = taCenterTop
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!FontSize = 14
            temptable!text = DelZero(sourcetable!MODELFACT)
            temptable!CardNo = nCard
            temptable.Update
            
'           factname
            temptable.AddNew
            temptable!Top = MyMeasure(0.6) + nHeight
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(3.5)
            temptable!TextAlign = taLeftTop
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!FontSize = 12
            temptable!text = sourcetable!descafact
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Top = MyMeasure(0.6) + nHeight
            temptable!Left = MyMeasure(3) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(0.8)
            temptable!TextAlign = taRightTop
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!FontSize = 12
            temptable!text = sourcetable!MOSM
            temptable!CardNo = nCard
            temptable.Update
            
'           desca
            temptable.AddNew
            temptable!Top = MyMeasure(1.1) + nHeight
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(3.5)
            temptable!TextAlign = taCenterTop
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!FontSize = 12
            temptable!text = sourcetable!desca
            temptable!CardNo = nCard
            temptable.Update
            
'           desca
            temptable.AddNew
            temptable!Top = MyMeasure(1.6) + nHeight
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(3.5)
            temptable!TextAlign = taCenterTop
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!FontSize = 12
            temptable!text = "Size range :" & DelZero(GetDesca("SELECT SCAL FROM FILE1_10 WHERE MODEL = " & MyParn(sourcetable!MODEL) & " ORDER BY C_SCAL ", con)) & " - " & DelZero(GetDesca("SELECT SCAL FROM FILE1_10 WHERE MODEL = " & MyParn(sourcetable!MODEL) & " ORDER BY C_SCAL DESC ", con))
            temptable!CardNo = nCard
            temptable.Update
            
'           color
            temptable.AddNew
            temptable!Top = MyMeasure(2.2) + nHeight
            temptable!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            temptable!Width = MyMeasure(3.5)
            temptable!TextAlign = taCenterTop
            temptable!Height = 0
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!FontSize = 12
            If Not (sourcetable!color = "„ÊÕœ" Or sourcetable!color = "---" Or sourcetable!color = "√·Ê«‰") Then
                temptable!text = sourcetable!color
            End If
            temptable!CardNo = nCard
            temptable.Update
            temptable!ForeColor = vbBlack
            temptable!CardNo = nCard
            temptable.Update
        

            
            
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
    sourcetable.MoveNext
Loop
End With
Set sourcetable = Nothing
End Function
Private Function Doprint_SAL(pText As String) As Boolean
nLastMargin = 0

SettingArray(cUpMargin) = MyMeasure(0.24)
SettingArray(cLeftMargin) = MyMeasure(0)
SettingArray(cCardWidth) = MyMeasure(4)
SettingArray(cCardHeight) = MyMeasure(3)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1

SettingArray(cRows) = 1
SettingArray(cCols) = 1

SettingArray(cPageWidth) = MyMeasure(21)

Dim temptable As New ADODB.Recordset
contemp.Execute "DELETE * FROM card"
temptable.Open "card", contemp, adOpenStatic, adLockOptimistic, adCmdTable



' ·«Œ Ì«— «·’› Ê«·⁄„Êœ


nBegin = ((IIf(Val(xRow.text) <= 0, 1, Val(xRow.text)) - 1) * nCols) + IIf(Val(xCol.text) <= 0, 1, Val(xCol.text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = nCols, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    BLASTROW = (nRow = NROWS)
    temptable.AddNew
    temptable!CardNo = nCard
    temptable.Update
Next
'«‰ Â«¡

nCard = 0
nRow = 0
nCard = 0
nCol = 0
nCols = SettingArray(cCols)
NROWS = SettingArray(cRows)
nFixLeft = MyMeasure(0.1)
nFixlast = MyMeasure(0.18)
nFixlastbar = MyMeasure(0.1)

'************
    For i = 1 To 1
        nCard = nCard + 1
        nCol = IIf(nCol = nCols, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        
        nHeight = 0
        For nCount = 1 To 1
            BLASTROW = (nRow = NROWS And nCount = 2)
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.5)
            temptable!Top = MyMeasure(0.5) - IIf(BLASTROW, MyMeasure(0.2), 0)
            temptable!Width = MyMeasure(3)
            temptable!Height = MyMeasure(0)
            temptable!FontBold = True
            temptable!TextAlign = taCenterTop
            temptable!text = pText
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            If pText = "SALE" Then
                temptable!FontSize = 26
            Else
                temptable!FontSize = 16
            End If
            temptable!CardNo = nCard
            temptable.Update
            
            temptable.AddNew
            temptable!Left = MyMeasure(0.5)
            temptable!Top = MyMeasure(1.5) - IIf(BLASTROW, MyMeasure(0.2), 0)
            temptable!Width = MyMeasure(3)
            temptable!Height = MyMeasure(0)
            temptable!FontBold = True
            temptable!TextAlign = taCenterTop
            temptable!text = Val(XRATE.text) & " % "
            temptable!FontName = "Arial"
            temptable!FontBold = True
            temptable!FontUnderline = False
            temptable!FontSize = 26
            temptable!CardNo = nCard
            temptable.Update
            
            nHeight = SettingArray(cCardHeight) / 2
        Next
' ----------------
    Next
Set sourcetable = Nothing
End Function


