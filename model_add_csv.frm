VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form model_add_csv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "”Õ» „ÊœÌ·«  „‰ „·›"
   ClientHeight    =   9615
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   20370
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9615
   ScaleMode       =   0  'User
   ScaleWidth      =   20370
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   20265
      _ExtentX        =   35745
      _ExtentY        =   15266
      _Version        =   393216
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "»Ì«‰«  €Ì— ’«·Õ…"
      TabPicture(0)   =   "model_add_csv.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "grid3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "„ÊœÌ·«  „ÊÃÊœ… »«·›⁄·"
      TabPicture(1)   =   "model_add_csv.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grid2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "„ÊœÌ·«  ÃœÌœ…"
      TabPicture(2)   =   "model_add_csv.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grid1"
      Tab(2).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   8205
         Left            =   -74910
         TabIndex        =   9
         Top             =   360
         Width           =   20070
         _cx             =   35401
         _cy             =   14473
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   14737632
         ForeColorFixed  =   0
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   -2147483633
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   8205
         Left            =   -74955
         TabIndex        =   10
         Top             =   360
         Width           =   20085
         _cx             =   35428
         _cy             =   14473
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
         BackColorFixed  =   12648384
         ForeColorFixed  =   0
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
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
      Begin VSFlex7Ctl.VSFlexGrid grid3 
         Height          =   8205
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   20085
         _cx             =   35428
         _cy             =   14473
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
         BackColorFixed  =   6249205
         ForeColorFixed  =   0
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   -2147483633
         SheetBorder     =   -2147483642
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
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
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   690
      ScaleWidth      =   20370
      TabIndex        =   0
      Top             =   8925
      Width           =   20370
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   -45
         Width           =   8925
         Begin Threed.SSCommand cmdExcel 
            Height          =   510
            Left            =   3015
            TabIndex        =   2
            Top             =   135
            Width           =   1950
            _ExtentX        =   3440
            _ExtentY        =   900
            _Version        =   196610
            CaptionStyle    =   1
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
            Picture         =   "model_add_csv.frx":0054
            Caption         =   " ’œÌ— Excel"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdPrint 
            Height          =   510
            Left            =   1440
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   135
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   900
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
            Picture         =   "model_add_csv.frx":24BE
            Caption         =   " ÿ»«⁄…"
            ButtonStyle     =   3
            PictureAlignment=   9
            BevelWidth      =   0
            PictureDisabledFrames=   1
            PictureDisabled =   "model_add_csv.frx":4834
         End
         Begin Threed.SSCommand cmdExit 
            Height          =   510
            Left            =   45
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   135
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   900
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
            Picture         =   "model_add_csv.frx":69B7
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            ShapeSize       =   1
         End
         Begin Threed.SSCommand cmdCSV 
            Height          =   510
            Left            =   6750
            TabIndex        =   5
            Top             =   135
            Width           =   2130
            _ExtentX        =   3757
            _ExtentY        =   900
            _Version        =   196610
            CaptionStyle    =   1
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
            Picture         =   "model_add_csv.frx":8CDA
            Caption         =   "”Õ» „‰ „·› CSV"
            TagVariant      =   "«Œ «— «·„ﬁ«Ê·"
            ButtonStyle     =   3
            PictureAlignment=   9
         End
         Begin Threed.SSCommand cmdSave 
            Height          =   510
            Left            =   4995
            TabIndex        =   7
            Top             =   135
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   900
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
            Picture         =   "model_add_csv.frx":B2AD
            Alignment       =   8
            ButtonStyle     =   3
            PictureAlignment=   11
            BevelWidth      =   0
            PictureDisabledFrames=   1
            ShapeSize       =   1
            PictureDisabled =   "model_add_csv.frx":DBD2
         End
      End
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   8730
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "model_add_csv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sDoc_no As String
Public myForm As Form
Private Sub cmdCSV_Click()
getData
End Sub

Private Sub cmdExcel_Click()
If SSTab1.Tab = 1 Then
    ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, TabStrip1.Tabs(1).Caption)
ElseIf SSTab1.Tab = 1 = 2 Then
    ToFileExelNew grid2, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, TabStrip1.Tabs(2).Caption)
ElseIf SSTab1.Tab = 1 = 3 Then
    ToFileExelNew grid3, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, TabStrip1.Tabs(3).Caption)
End If
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdPrint_Click()
Set PrintGrdNew.myForm = Me
If TabStrip1.SelectedItem.index = 1 Then
    PrintGrdNew.doprint grid1, 0.85, 0, "ÃÊ‰ÌÊ—", Me.Caption, ArbString(retHeader(aHeader, 0, 5)), TabStrip1.Tabs(1).Caption, False, True, 10
ElseIf TabStrip1.SelectedItem.index = 2 Then
    PrintGrdNew.doprint grid2, 0.85, 0, "ÃÊ‰ÌÊ—", Me.Caption, ArbString(retHeader(aHeader, 0, 5)), TabStrip1.Tabs(2).Caption, False, True, 10
ElseIf TabStrip1.SelectedItem.index = 3 Then
    PrintGrdNew.doprint grid3, 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption, ArbString(retHeader(aHeader, 0, 5)), TabStrip1.Tabs(2).Caption, False, False, 10
End If
PrintGrdNew.Show 1
End Sub
Private Function myreplace() As Boolean
With grid1
Dim sCaption As String
Dim i As Long
prog1.Visible = True
prog1.Max = grid1.Rows
Dim db As New clsDb
Dim sb As New ChilkatStringBuilder
Dim aInsert As Variant
Dim nCount As Long
For i = 1 To .Rows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & grid1.Rows - 1
    prog1.Value = i
    
    aInsert = AddFlag(Empty, "MODEL", addstring(grid1.TextMatrix(i, 1)))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(grid1.TextMatrix(i, 1)))
    aInsert = AddFlag(aInsert, "DESCA", addstring(grid1.TextMatrix(i, 2)))
    aInsert = AddFlag(aInsert, "FACT", addstring(grid1.TextMatrix(i, 3)))
    aInsert = AddFlag(aInsert, "MOSM", addstring(grid1.TextMatrix(i, 4)))
    aInsert = AddFlag(aInsert, "MODELFACT", addstring(grid1.TextMatrix(i, 5)))
    aInsert = AddFlag(aInsert, "MODELFACT0", addstring(DelZero(grid1.TextMatrix(i, 5))))
    aInsert = AddFlag(aInsert, "CODE", addstring(grid1.TextMatrix(i, 6)))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(grid1.TextMatrix(i, 7)))
    aInsert = AddFlag(aInsert, "[SECTION]", addstring(grid1.TextMatrix(i, 8)))
    aInsert = AddFlag(aInsert, "[MODELSEX]", addstring(grid1.TextMatrix(i, 9)))
    aInsert = AddFlag(aInsert, "[MODELAGE]", addstring(grid1.TextMatrix(i, 10)))
    aInsert = AddFlag(aInsert, "[SCAL]", addstring(grid1.TextMatrix(i, 11)))
    aInsert = AddFlag(aInsert, "[COLOR]", addstring(grid1.TextMatrix(i, 12)))
    aInsert = AddFlag(aInsert, "[COST]", grid1.ValueMatrix(i, 13))
    aInsert = AddFlag(aInsert, "[PRICE]", grid1.ValueMatrix(i, 14))
    aInsert = AddFlag(aInsert, "[COST2]", grid1.ValueMatrix(i, 15))
    aInsert = AddFlag(aInsert, "[BARCODE]", addstring(grid1.TextMatrix(i, 16)))
    aInsert = AddFlag(aInsert, "[C_COLOR]", db.MyFuncValue("dbo.fn_last_color", MyParn(sModel)))
    aInsert = AddFlag(aInsert, "[C_SCAL]", db.MyFuncValue("dbo.fn_last_scal", MyParn(sModel)))
    aInsert = AddFlag(aInsert, "ITEM", db.MyFuncValue("dbo.fn_last_item"))
    If Not db.Execute(addInsert(aInsert, "FILE1_10")) Then GoTo cleanUP
Next
End With
prog1.Visible = False
myInform " „  ÕœÌÀ «·„” ‰œ«  »‰Ã«Õ"
bSave = True
myreplace = True
cleanUP:
prog1.Visible = True
Set db = Nothing
End Function

Private Sub cmdSave_Click()
If grid1.Rows <= 1 Then MsgBox "·«  ÊÃœ ”Ã·«  ··Õ›Ÿ"
If myreplace Then
    myForm.myProc
    Unload Me
End If
End Sub

Private Sub Form_Load()
Fixgrd1
Fixgrd2
Fixgrd3
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set dbm = Nothing
Set online_ship_checkfrm = Nothing
End Sub
Private Sub GRID3_DblClick()
If grid3.Col = 1 And grid3.Row > 0 Then
    Clipboard.Clear
    Clipboard.SetText grid3.TextMatrix(grid3.Row, grid3.Col)
    Inform " „ «·‰”Œ"
End If
End Sub
Private Function getData()
Dim cFileName As String
Dim sb As New ChilkatStringBuilder

Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName

If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
Dim i As Long

'MsgBox GetTextFileEncoding(cFileName)

nAccess = sb.LoadFile(cFileName, "auto")

If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


cSv.HasColumnNames = 1
nAccess = cSv.LoadFromString(sb.GetAsString)

If cSv.NumRows < 1 Then Exit Function

myloadNew cSv
End Function
Private Function myloadNew(ByRef cSv As ChilkatCsv) As Boolean
Dim i As Long
Dim sCaption As String
Dim ship_no As String
Dim date_Pick As String
Dim cod As String
Dim Tb As New ChilkatStringBuilder


sCaption = Me.Caption
prog1.Visible = True

grid1.Rows = 1
grid1.Redraw = flexRDNone

grid2.Rows = 1
grid2.Redraw = flexRDNone

grid3.Rows = 1
grid3.Redraw = flexRDNone


Dim strSql As String
strSql = "SELECT * " & _
          " FROM file1_10"
                       
Dim db As New clsData

On Error GoTo myError
Dim locTable As ADODB.RecordSet
Set locTable = db.rs(strSql)
prog1.Max = cSv.NumRows
For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = i + 1
        
    mCode = Trim(cSv.GetCellByName(i, "«·„Ê—œ"))
    mDesca = Trim(cSv.GetCellByName(i, "«·’‰›"))
    mFact = Trim(cSv.GetCellByName(i, "«·„’‰⁄"))
    mseason = Trim(cSv.GetCellByName(i, "«·„Ê”„"))
    mModelFact = Trim(cSv.GetCellByName(i, "„ÊœÌ·"))
    mModelSex = Trim(cSv.GetCellByName(i, "«·‰Ê⁄"))
    mModelAge = Trim(cSv.GetCellByName(i, "«·›∆Â «·⁄„—ÌÂ"))
    mScal = Trim(cSv.GetCellByName(i, "„ﬁ«”"))
    mColor = Trim(cSv.GetCellByName(i, "«··Ê‰"))
    mCost = Trim(cSv.GetCellByName(i, "”⁄— «·„’‰⁄"))
    mPrice = Trim(cSv.GetCellByName(i, "”⁄— «·„” Â·ﬂ"))
    mCost2 = Trim(cSv.GetCellByName(i, "”⁄— «·Ã„·Â"))
    mBarCode = Trim(cSv.GetCellByName(i, "»«—ﬂÊœ Œ«—ÃÏ"))
    mGroup = Trim(cSv.GetCellByName(i, "«·„Ã„Ê⁄Â"))
    
    mSection = Trim(cSv.GetCellByName(i, "«·ﬁ”„"))
    
    sError = ""
    
    If Not (mCode = "" And mFact = "" And mseason = "" And mModelFact = "" And mModelSex = "" And mScal = "" And mColor = "" And mPrice = "" And mCost2 = "" And mBarCode = "") Then
         If mCode = "" Then
            sError = "ﬂÊœ «·„Ê—œ €Ì— „”Ã·"
         ElseIf mCode <> "" Then
            If IsEmpty(db.rsField("select code from file4_10 where code = " & MyParn(mCode))) Then
                sError = "ﬂÊœ «·„Ê—œ €Ì— ’ÕÌÕ"
            End If
         End If
         
         If mFact = "" Then
            sError = sError & Tr(sError, "-") & "ﬂÊœ «·„’‰⁄ €Ì— „”Ã·"
         Else
            If IsEmpty(db.rsField("select code from fact where code = " & MyParn(mFact))) Then
                sError = sError & Tr(sError, "-") & "ﬂÊœ ««·„’‰⁄ €Ì— ’ÕÌÕ"
            End If
         End If
         
         If mseason = "" Then
            sError = sError & Tr(sError, "-") & "ﬂÊœ «·„Ê”„ €Ì— „”Ã·"
         Else
            If IsEmpty(db.rsField("select mosm from mosm where mosm = " & MyParn(mseason))) Then
                sError = sError & Tr(sError, "-") & "ﬂÊœ «·„Ê”„ €Ì— ’ÕÌÕ"
            End If
         End If

         If mModelSex = "" Then
            sError = sError & Tr(sError, "-") & "‰Ê⁄ «·„ÊœÌ· €Ì— „”Ã·"
         Else
            If IsEmpty(db.rsField("select code from file0_81 where code = " & MyParn(mModelSex))) Then
                sError = sError & Tr(sError, "-") & "‰Ê⁄ «·„ÊœÌ· €Ì— ’ÕÌÕ"
            End If
         End If
         
         If mModelAge = "" Then
            sError = sError & Tr(sError, "-") & "«·›∆… «·⁄„—Ì… €Ì— „”Ã·…"
         Else
            If IsEmpty(db.rsField("select code from file0_82 where code = " & MyParn(mModelAge))) Then
                sError = sError & Tr(sError, "-") & "«·›∆… «·⁄„—Ì… €Ì— ’ÕÌÕ…"
            End If
         End If
         
         If mGroup = "" Then
            sError = sError & Tr(sError, "-") & "«·„Ã„Ê⁄… €Ì— „”Ã·…"
         Else
            If IsEmpty(db.rsField("select code from file1_50 where code = " & MyParn(mGroup))) Then
                sError = sError & Tr(sError, "-") & "«·„Ã„Ê⁄… €Ì— ’ÕÌÕ…"
            End If
         End If
         
         If mSection = "" Then
            sError = sError & Tr(sError, "-") & "«·ﬁ”„ €Ì— „”Ã·"
         Else
            If IsEmpty(db.rsField("select code from file1_10SC where code = " & MyParn(mSection))) Then
                sError = sError & Tr(sError, "-") & "«·ﬁ”„ €Ì— ’ÕÌÕ"
            End If
         End If
         
         If mScal = "" Then
            sError = sError & Tr(sError, "-") & "«·„ﬁ«” €Ì— „”Ã·"
         End If
         
         If mColor = "" Then
            sError = sError & Tr(sError, "-") & "«··Ê‰ €Ì— „”Ã·"
         End If
         
         sModel = mFact & mseason & RetZero(mModelFact, 10)
        
            
         If sError <> "" Then
            grid3.AddItem ""
            grid3.TextMatrix(grid3.Rows - 1, 0) = grid3.Rows - 1
            grid3.TextMatrix(grid3.Rows - 1, 1) = sError
         Else
            locTable.Find "MODEL = " & MyParn(sModel), , adSearchForward, adBookmarkFirst
            If locTable.EOF Then
                grid1.AddItem ""
                grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
                grid1.TextMatrix(grid1.Rows - 1, 1) = sModel
                grid1.TextMatrix(grid1.Rows - 1, 2) = mDesca
                grid1.TextMatrix(grid1.Rows - 1, 3) = mFact
                grid1.TextMatrix(grid1.Rows - 1, 4) = mseason
                grid1.TextMatrix(grid1.Rows - 1, 5) = mModelFact
                grid1.TextMatrix(grid1.Rows - 1, 6) = mCode
                grid1.TextMatrix(grid1.Rows - 1, 7) = mGroup
                grid1.TextMatrix(grid1.Rows - 1, 8) = mSection
                
                grid1.TextMatrix(grid1.Rows - 1, 9) = mModelSex
                grid1.TextMatrix(grid1.Rows - 1, 10) = mModelAge
                grid1.TextMatrix(grid1.Rows - 1, 11) = mScal
                grid1.TextMatrix(grid1.Rows - 1, 12) = mColor
                                                           
                grid1.TextMatrix(grid1.Rows - 1, 13) = mCost
                grid1.TextMatrix(grid1.Rows - 1, 14) = mPrice
                grid1.TextMatrix(grid1.Rows - 1, 15) = mCost2
                grid1.TextMatrix(grid1.Rows - 1, 16) = mBarCode
            Else
                grid2.AddItem ""
                grid2.TextMatrix(grid2.Rows - 1, 0) = grid2.Rows - 1
                grid2.TextMatrix(grid2.Rows - 1, 1) = sModel
                grid2.TextMatrix(grid2.Rows - 1, 2) = mDesca
                grid2.TextMatrix(grid2.Rows - 1, 3) = mFact
                grid2.TextMatrix(grid2.Rows - 1, 4) = mseason
                grid2.TextMatrix(grid2.Rows - 1, 5) = mModelFact
                grid2.TextMatrix(grid2.Rows - 1, 6) = mCode
                grid2.TextMatrix(grid2.Rows - 1, 7) = mGroup
                grid2.TextMatrix(grid2.Rows - 1, 8) = mSection
                
                grid2.TextMatrix(grid2.Rows - 1, 9) = mModelSex
                grid2.TextMatrix(grid2.Rows - 1, 10) = mModelAge
                grid2.TextMatrix(grid2.Rows - 1, 11) = mScal
                grid2.TextMatrix(grid2.Rows - 1, 12) = mColor
                                                           
                grid2.TextMatrix(grid2.Rows - 1, 13) = mCost
                grid2.TextMatrix(grid2.Rows - 1, 14) = mPrice
                grid2.TextMatrix(grid2.Rows - 1, 15) = mCost2
                grid2.TextMatrix(grid2.Rows - 1, 16) = mBarCode
            End If
        End If
    End If
Next
cleanUP:
grid1.Redraw = flexRDBuffered
grid2.Redraw = flexRDBuffered
grid3.Redraw = flexRDBuffered
Fixgrd1
Fixgrd2
prog1.Visible = False
Me.Caption = sCaption
Set db = Nothing
Exit Function
myError:
MsgBox Err.Description
Err.Clear
End Function
Private Sub Fixgrd1()
With grid1

.Redraw = flexRDNone
.RowHeight(0) = 600

.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "—ﬁ„ «·„ÊœÌ·"
.TextMatrix(0, 2) = "«·’‰›"
.TextMatrix(0, 3) = "«·„’‰⁄"
.TextMatrix(0, 4) = "«·„Ê”„"
.TextMatrix(0, 5) = "„ÊœÌ· «·„’‰⁄"
.TextMatrix(0, 6) = "«·„Ã„Ê⁄…"
.TextMatrix(0, 6 + 1) = "«·„Ê—œ"
.TextMatrix(0, 7 + 1) = "«·ﬁ”„"
.TextMatrix(0, 8 + 1) = "«·‰Ê⁄"
.TextMatrix(0, 9 + 1) = "«·›∆… «·⁄„—Ì…"
.TextMatrix(0, 10 + 1) = "«·„ﬁ«”"
.TextMatrix(0, 11 + 1) = "«··Ê‰"
.TextMatrix(0, 12 + 1) = "”⁄— «·„’‰⁄"
.TextMatrix(0, 13 + 1) = "”⁄— «·„” Â·ﬂ"
.TextMatrix(0, 14 + 1) = "”⁄— «·Ã„·…"
.TextMatrix(0, 15 + 1) = "»«—ﬂÊœ Œ«—ÃÌ"

.ColWidth(0) = 600
.ColWidth(1) = 2400
.ColWidth(2) = 2900
.ColWidth(3) = 800
.ColWidth(4) = 800
.ColWidth(5) = 1500
.ColWidth(6) = 820
.ColWidth(6 + 1) = 800
.ColWidth(7 + 1) = 800
.ColWidth(8 + 1) = 800
.ColWidth(9 + 1) = 800
.ColWidth(10 + 1) = 1400
.ColWidth(11 + 1) = 1400
.ColWidth(12 + 1) = 800
.ColWidth(13 + 1) = 800
.ColWidth(14 + 1) = 800
.ColWidth(15 + 1) = 1400

Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

For i = 1 To grid1.Rows - 1
    .TextMatrix(i, 0) = i
Next
.Redraw = flexRDBuffered
End With
End Sub
Private Sub Fixgrd2()
With grid2
.Redraw = flexRDNone

.RowHeight(0) = 600

.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "—ﬁ„ «·„ÊœÌ·"
.TextMatrix(0, 2) = "«·’‰›"
.TextMatrix(0, 3) = "«·„’‰⁄"
.TextMatrix(0, 4) = "«·„Ê”„"
.TextMatrix(0, 5) = "„ÊœÌ· «·„’‰⁄"
.TextMatrix(0, 6) = "«·„Ã„Ê⁄…"
.TextMatrix(0, 6 + 1) = "«·„Ê—œ"
.TextMatrix(0, 7 + 1) = "«·ﬁ”„"
.TextMatrix(0, 8 + 1) = "«·‰Ê⁄"
.TextMatrix(0, 9 + 1) = "«·›∆… «·⁄„—Ì…"
.TextMatrix(0, 10 + 1) = "«·„ﬁ«”"
.TextMatrix(0, 11 + 1) = "«··Ê‰"
.TextMatrix(0, 12 + 1) = "”⁄— «·„’‰⁄"
.TextMatrix(0, 13 + 1) = "”⁄— «·„” Â·ﬂ"
.TextMatrix(0, 14 + 1) = "”⁄— «·Ã„·…"
.TextMatrix(0, 15 + 1) = "»«—ﬂÊœ Œ«—ÃÌ"

.ColWidth(0) = 600
.ColWidth(1) = 2400
.ColWidth(2) = 2900
.ColWidth(3) = 800
.ColWidth(4) = 800
.ColWidth(5) = 1500
.ColWidth(6) = 820
.ColWidth(6 + 1) = 800
.ColWidth(7 + 1) = 800
.ColWidth(8 + 1) = 800
.ColWidth(9 + 1) = 800
.ColWidth(10 + 1) = 1400
.ColWidth(11 + 1) = 1400
.ColWidth(12 + 1) = 800
.ColWidth(13 + 1) = 800
.ColWidth(14 + 1) = 800
.ColWidth(15 + 1) = 1400

Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
Next
.Redraw = flexRDBuffered
End With
End Sub
Private Sub Fixgrd3()
With grid3
'.RowHeight(0) = 400
.TextMatrix(0, 0) = "„"
.TextMatrix(0, 1) = "«·„ÊœÌ·"
.TextMatrix(0, 2) = "«·’‰›"
.TextMatrix(0, 3) = "«·Œÿ√"


.ColWidth(0) = 600
.ColWidth(1) = 2400
.ColWidth(2) = 2900
.ColWidth(3) = 13500
Dim i As Long
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next

For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
Next
End With
End Sub
