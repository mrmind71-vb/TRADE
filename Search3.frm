VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Search3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "ÅÓÊÚáÇã"
   ClientHeight    =   6825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
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
   ScaleHeight     =   6825
   ScaleWidth      =   11880
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Factory"
   Begin VSFlex7Ctl.VSFlexGrid Grid1 
      Bindings        =   "Search3.frx":0000
      Height          =   5790
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   11490
      _cx             =   20267
      _cy             =   10213
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   14737632
      ForeColorSel    =   0
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
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
      AutoResize      =   -1  'True
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
      TabBehavior     =   0
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
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   6510
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   75
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   5925
      Width           =   11415
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   150
      Top             =   6000
      Visible         =   0   'False
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
         Name            =   "MS Sans Serif"
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
Attribute VB_Name = "Search3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const pGeneral = 0, pGrid = 1
Dim lApply As Boolean, nwidth As Integer, nHeight
Sub MyApply()
Dim myBound As Integer
myBound = UBound(Lookupdata(pGeneral))
If myBound >= 2 Then cWhere = Lookupdata(pGeneral)(2)
cLookupstr = Lookupdata(pGeneral)(1)
If myBound >= 3 Then cLookupstr = cLookupstr & " " & Lookupdata(pGeneral)(3)
data1.RecordSource = cLookupstr
data1.Refresh
For i = 0 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next
For i = 0 To UBound(Lookupdata(pGrid))
   grid1.ColWidth(i) = Lookupdata(pGrid)(i)
Next
grid1.ColWidth(grid1.Cols - 1) = 500
'If Data1.Recordset.RecordCount > 0 Then Data1.Recordset.MoveLast
StatusBar1.Panels(1).Text = IIf(data1.Recordset.RecordCount > 0, " ÚÏÏ ÇáÓÌáÇÊ ÇáãØÇÈÞÉ " & data1.Recordset.RecordCount, " áÇ ÊæÌÏ ÓÌáÇÊ ãØÇÈÞÉ ")
End Sub
Function MyFormatString(pArray)
MyString = ""
For i = 0 To UBound(pArray) - 1
   MyString = MyString & IIf(MyString = "", "", "|") & pArray(i, 2)
Next
MyFormatString = MyString
End Function
Private Sub CmdAll_Click()
cLookupstr = Lookupdata(pGeneral)(1)
cmdgo.Enabled = False
End Sub
Private Sub CmdExit_Click()
ReDim Lookupdata(0)
Unload Search
End Sub
Private Sub Form_Activate()
If Text1.Visible Then Text1.SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
grid1.ExplorerBar = flexExSort
data1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & PublicPath & "\DATA\Data.mdb"
data1.CommandType = adCmdText
grid1.Cols = UBound(Lookupdata(pGrid)) + 1
MyApply
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Search = Nothing
End Sub
Private Sub Grid1_DblClick()
If grid1.Row > 0 Then
    Lookupdata(pGeneral)(0).myProc
    Unload Me
End If
End Sub
Private Sub LstKind_Click()
If cmbLookup.Visible Then
    cmdgo.Enabled = cmbLookup.Text <> ""
Else
    cmdgo.Enabled = txtlookup.Text <> ""
End If
End Sub
'Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then Grid1_DblClick
'End Sub
Private Sub Text1_Change()
lApply = True
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If lApply Then MyApply
End If
End Sub
Private Sub MYAPPLY2()
grid1.SetFocus
nRow = grid1.FindRow(Text1.Text, , 6, False, False)
If nRow > 0 Then
    grid1.ShowCell nRow, 0
    grid1.Select nRow, 0
    Exit Sub
End If

nRow = grid1.FindRow(Text1.Text, , 4, False, False)
If nRow > 0 Then
    grid1.ShowCell nRow, 0
    grid1.Select nRow, 0
    Exit Sub
End If
End Sub
