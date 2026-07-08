VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdMonthProfit 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ì «—»«Õ ‘Â—Ì… ··›—Ê⁄"
   ClientHeight    =   10290
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
   ScaleHeight     =   10290
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   360
      Width           =   5640
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " ÊﬂÌ·«  ›ﬁÿ"
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
         Index           =   2
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   315
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "›—Ê⁄ ›ﬁÿ"
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
         Index           =   1
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   315
         Width           =   1140
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
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
         Index           =   0
         Left            =   4545
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   315
         Value           =   -1  'True
         Width           =   780
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   9990
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   5685
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdMonthProfit.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1170
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdMonthProfit.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdMonthProfit.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   4545
         TabIndex        =   3
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdMonthProfit.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2295
         TabIndex        =   11
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdMonthProfit.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPdf 
         Height          =   555
         Left            =   3420
         TabIndex        =   12
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
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
         Picture         =   "grdMonthProfit.frx":BBFA
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1050
      Left            =   15705
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   4560
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
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
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1935
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   330
         Left            =   90
         TabIndex        =   2
         Top             =   585
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ"
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
         Left            =   3870
         TabIndex        =   7
         Top             =   270
         Width           =   510
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·›—⁄"
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
         Left            =   3870
         TabIndex        =   6
         Top             =   630
         Width           =   405
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   90
      Top             =   270
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   3285
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   7935
      Left            =   90
      TabIndex        =   4
      Top             =   1125
      Width           =   20220
      _cx             =   35666
      _cy             =   13996
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
      Cols            =   13
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   1080
      Top             =   -90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Width           =   1335
      _ExtentX        =   2355
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   13
      Top             =   10095
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "grdMonthProfit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim FieldTable As ADODB.Recordset
Dim cString As String, aHeader()
Private Sub cmdExcel_Click()
Me.MousePointer = 11
Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)

For I = 1 To grid1.Rows - 2
    If grid1.TextMatrix(I, 2) = "" Then
        aSub = AddFlag(Empty, "row", I)
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "word_wrap", False)
        aSub = AddFlag(aSub, "back_color", 19)
        aRow = AddFlag(aRow, aSub)
    End If
Next

If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", grid1.Rows - 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 40)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2), retHeader(aHeader, 4, 2))
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
myPrint "1", True
End Sub
Private Sub cmdPrint_Click()
myPrint
End Sub
Private Sub Form_Load()
Dim cWhere As String
openCon con

If Not (lSupperVisor Or bOpt5 Or cBranch <> "00") Then
    cWhere = " Where STORE_BAR.CODE IN (SELECT STORE FROM USERSHOP WHERE CODE = " & MyParn(nusercode) & " ) " & "  ORDER BY CODE "
End If

Set data1.Recordset = myCmd("SELECT * FROM STORE_BR " & cWhere, con)
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

Set grid1.DataSource = DATA11

fixGrd
End Sub
Private Sub myload()
Dim cString As String, cWhere As String
With grid1

ReDim aHeader(3)

If xStore.MatchedWithList Then
    aPrm = AddFlag(aPrm, "STORE", xStore.BoundText)
    aHeader(0) = "„Œ“‰ : " & xStore.text
End If

If IsDate(xDate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", myFormat_sp(xDate1.text))
    aHeader(1) = BetweenString(xDate1.text, xDate2.text)
End If

If IsDate(xDate2.text) Then
    aPrm = AddFlag(aPrm, "DATE2", myFormat_sp(xDate2.text))
    aHeader(1) = BetweenString(xDate1.text, xDate2.text)
End If

Set DATA11.Recordset = myCmd("dbo.sp_branch_sales_profit", con, adStoredProc, aPrm, 400)

End With
fixGrd

End Sub
Sub fixGrd()
Dim nTotal_Sales As Double, nTotal_in As Double
    With grid1
    .RowHeight(0) = 1000
    .WordWrap = True
    .TextMatrix(0, 0) = "«·›—⁄"
    .TextMatrix(0, 1) = "«·›—⁄"
    .TextMatrix(0, 2) = "«· «—ÌŒ"
    .TextMatrix(0, 3) = "„»Ì⁄« "
    .TextMatrix(0, 4) = " ﬂ·›…"
    .TextMatrix(0, 5) = "«—»«Õ „»Ì⁄« "
    .TextMatrix(0, 6) = "„’«—Ì›"
    .TextMatrix(0, 7) = " ”ÊÌ«  „Ê—œÌ‰"
    .TextMatrix(0, 8) = "’«›Ì «—»«Õ"
    .TextMatrix(0, 9) = "‰”»… —»Õ ··„»Ì⁄« "
    .TextMatrix(0, 10) = "«·‰”»… „»Ì⁄«  ··≈Ã„«·Ï"
    .TextMatrix(0, 11) = "‰”»… „’«—Ì› ··„»Ì⁄« "
    .TextMatrix(0, 12) = "‰”»… ’«›Ï «·—»Õ ··„»Ì⁄« "
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 1500
    .ColWidth(2) = 1000
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    .ColWidth(5) = 1400
    .ColWidth(6) = 1400
    .ColWidth(7) = 1400
    .ColWidth(8) = 1200
    .ColWidth(9) = 1200
    .ColWidth(10) = 1200
    .ColWidth(11) = 1000
    .ColWidth(12) = 1000
    
    .ColHidden(0) = True
    
    .MergeCells = flexMergeFree
    .MergeCol(1) = True
    
    
    For I = 2 To .Cols - 1
        .ColDataType(I) = flexDTDouble
    Next
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    
    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 3, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 5, "#0", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 6, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 7, "#0.00", vbRed, vbYellow, True, "  "
    .Subtotal flexSTSum, -1, 8, "#0.00", vbRed, vbYellow, True, "  "
    If .Rows > 1 Then
        .TextMatrix(1, 1) = "«·≈Ã„«·Ï"
                
        If Val(.TextMatrix(1, 2)) <> 0 Then .TextMatrix(1, 8) = Round(Val(.TextMatrix(1, 4)) / Val(.TextMatrix(1, 2)) * 100, 2)
        If Val(.TextMatrix(1, 2)) <> 0 Then .TextMatrix(1, 10) = Round(Val(.TextMatrix(1, 5)) / Val(.TextMatrix(1, 2)) * 100, 2)
        If Val(.TextMatrix(1, 2)) <> 0 Then .TextMatrix(1, 11) = Round(Val(.TextMatrix(1, 7)) / Val(.TextMatrix(1, 2)) * 100, 2)
        .Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
    End If
    .Subtotal flexSTSum, -1, 9, "#0.00", vbRed, vbYellow, True, "  "
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdChargefrm1 = Nothing
End Sub
Private Function MYVALID() As Boolean
MYVALID = True
End Function

Private Sub grid1_DblClick()
'If grid1.Row > 1 And (grid1.Col = 2 Or grid1.Col = 3 Or grid1.Col = 4) Then
'    Dim aLocal As Variant
'    aLocal = AddFlag(aLocal, "DATE1", xDate1.text)
'    aLocal = AddFlag(aLocal, "DATE2", xDate2.text)
'    aLocal = AddFlag(aLocal, "MAN", grid1.TextMatrix(grid1.Row, 0))
'    aLocal = AddFlag(aLocal, "STORE", xstore.BoundText)
'    If grid1.Col = 3 Then
'        aLocal = AddFlag(aLocal, "WHERE", "INV_TOTAL.DISCOUNT <> 0")
'        aLocal = AddFlag(aLocal, "CAPTION", " ›’Ì·Ì „»Ì⁄«  ·Â« Œ’„")
'    End If
'    showSales_inv.aData = aLocal
'    showSales_inv.Show 1
'ElseIf grid1.Col > 4 Then
'    aLocal = AddFlag(aLocal, "TYPE", 0)
'    aLocal = AddFlag(aLocal, "STORE", xstore.BoundText)
'    aLocal = AddFlag(aLocal, "DATE1", xDate1.text)
'    aLocal = AddFlag(aLocal, "DATE2", xDate2.text)
'    aLocal = AddFlag(aLocal, "MAN", grid1.TextMatrix(grid1.Row, 0))
'    aLocal = AddFlag(aLocal, "GROUP", retFlag(aGroup(grid1.Col - 5), "CODE"))
'    showSales.aData = aLocal
'    showSales.Show 1
'End If
End Sub
Private Sub Option1_Click(Index As Integer)
myload
End Sub
Private Sub CreateFieldTable(cWhere As String)
Dim cString As String
cString = "SELECT vw_item_exp.item ,vw_item_exp.item_desca as desca" & _
          " FROM vw_item_exp"
If cWhere <> "" Then cString = cString & " WHERE " & cWhere
cString = cString & " GROUP BY  vw_item_exp.item,vw_item_exp.item_desca"
cString = cString & " ORDER BY vw_item_exp.item"
Set FieldTable = Nothing
Set FieldTable = myCmd(cString, con)
End Sub
Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
If grid1.Rows < 3 Then Exit Sub
Dim aRow As Variant, aSub As Variant

For I = 1 To grid1.Rows - 2
    If grid1.TextMatrix(I, 2) = "" Then
        aSub = AddFlag(Empty, "row", I)
        aSub = AddFlag(aSub, "col", 1)
        aSub = AddFlag(aSub, "cols", 2)
        aSub = AddFlag(aSub, "text", grid1.TextMatrix(I, 1))
        aRow = AddFlag(aRow, aSub)
    End If
Next

aSub = AddFlag(Empty, "row", grid1.Rows - 1)
aSub = AddFlag(aSub, "col", 1)
aSub = AddFlag(aSub, "cols", 2)
aSub = AddFlag(aSub, "text", "≈Ã„«·Ì")
aRow = AddFlag(aRow, aSub)

'PrintGrdNew.sDeviceType = pDevice
'PrintGrdNew.Vp.Preview = Not bIgPreview

'PrintGrdNew.bIgPreview = bIgPreview
PrintGrdNew.DOPRINT Me.grid1, 0.95, 0, "›ÌÊ ‘—", Me.Caption, retHeader(aHeader, 0, 2), , False, IIf(Option1(0).Value, True, False), 10, , aRow, Array(1)
If Not bIgPreview Then
    PrintGrdNew.Show 1
Else
    Unload PrintGrdNew
    'PrintGrdNew.Show 1
    'Unload PrintGrdNew
End If
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xProject_GotFocus()
myGotFocus xProject
End Sub
Private Sub xProject_LostFocus()
myLostFocus xProject
If Not xProject.MatchedWithList Then xProject.BoundText = ""
End Sub
Private Sub xitem_GotFocus()
myGotFocus xitem
End Sub
Private Sub xitem_LostFocus()
myLostFocus xitem
If Not xitem.MatchedWithList Then xitem.BoundText = ""
End Sub
