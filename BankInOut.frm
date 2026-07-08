VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form BankInOutfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ· Õ—ﬂ… «·»‰ﬂ"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14865
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   14865
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   555
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   8055
      Width           =   3975
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   2985
         Picture         =   "BankInOut.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   2070
         Picture         =   "BankInOut.frx":25C0
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   915
         Picture         =   "BankInOut.frx":4B93
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   45
         Picture         =   "BankInOut.frx":726D
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   135
         Width           =   870
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   5535
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   630
      Width           =   1500
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BankInOut.frx":994C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BankInOut.frx":9D8E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   9225
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   5505
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BankInOut.frx":C307
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1410
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BankInOut.frx":C451
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   2730
         MaskColor       =   &H00FFFFFF&
         Picture         =   "BankInOut.frx":ECEB
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4050
         Picture         =   "BankInOut.frx":11297
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   7065
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   630
      Width           =   7665
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5175
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1290
      End
      Begin MSDataListLib.DataCombo XBANK 
         Height          =   315
         Left            =   2790
         TabIndex        =   4
         Top             =   630
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·»‰ﬂ :"
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
         Left            =   6570
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   675
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6555
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   930
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
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   8640
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "09:58 „"
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc data10 
      Height          =   465
      Left            =   2025
      Top             =   1170
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   820
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
      Height          =   6270
      Left            =   180
      TabIndex        =   6
      Top             =   1755
      Width           =   14550
      _cx             =   25665
      _cy             =   11060
      _ConvInfo       =   1
      Appearance      =   0
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
   Begin Threed.SSCommand CMD_PRINT 
      Height          =   465
      Left            =   225
      TabIndex        =   20
      Top             =   1215
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   820
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
      Picture         =   "BankInOut.frx":13A6A
      Caption         =   "ÿ»«⁄… «·„” ‰œ  "
      Alignment       =   1
      PictureAlignment=   3
   End
End
Attribute VB_Name = "BankInOutfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bedit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset
Dim cStrBox As String, SearchItems As New Search3
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(1, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xdoc_no.text)

aInsert(1, 0) = "Bank"
aInsert(1, 1) = addstring(xBank.BoundText)

On Error GoTo myerror
con.BeginTrans
If xdoc_no.Enabled Then
    
    xdoc_no.text = RetZero(Val(Newflag("FILE5_30H", "doc_no", con)))
    aInsert(0, 1) = addstring(xdoc_no.text)
    con.Execute CreateInsert(aInsert, "FILE5_30H")
Else
    con.Execute CreateUpdate(aInsert, "FILE5_30H", " where doc_no = " & addstring(xdoc_no.text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd()
Dim aInsert(7, 1)
With grid1
    For i = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xdoc_no.text)
                
        aInsert(1, 0) = "code"
        aInsert(1, 1) = addvalue(grid1.TextMatrix(i, 0))
        
        aInsert(2, 0) = "date"
        aInsert(2, 1) = addDate(.TextMatrix(i, 2))
                
        aInsert(3, 0) = "Box"
        aInsert(3, 1) = addstring(.TextMatrix(i, 3))
                        
        aInsert(4, 0) = "Desca"
        aInsert(4, 1) = addstring(grid1.TextMatrix(i, 4))
        
        aInsert(5, 0) = "[value1]"
        aInsert(5, 1) = Val(grid1.TextMatrix(i, 5))

        aInsert(6, 0) = "[value2]"
        aInsert(6, 1) = Val(grid1.TextMatrix(i, 6))

        aInsert(7, 0) = "row"
        aInsert(7, 1) = i
        
        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "FILE5_30")
        Else
            con.Execute CreateUpdate(aInsert, "FILE5_30", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 0 Then
        grid1.TextMatrix(grid1.Row, 0) = SearchItems.grid1.TextMatrix(SearchItems.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 1) = SearchItems.grid1.TextMatrix(SearchItems.grid1.Row, 1)
        If grid1.Row = grid1.Rows - 1 And validRow(grid1.Row) Then
            grid1.AddItem ""
            grid1.Select grid1.Row + 1, 0
        End If
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    myload
    Unload Search3
End If
End Sub


Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete From FILE5_30 where Doc_No = " & MyParn(xdoc_no.text)
    con.Execute "Delete From FILE5_30H where Doc_No = " & MyParn(xdoc_no.text)
    con.CommitTrans
    CardTable.Requery
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xdoc_no.text), , adSearchBackward, adBookmarkLast
        If CardTable.EOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT FILE5_30H.Doc_No, CONVERT(VARCHAR(10),MIN(FILE5_30.[Date]),111),File5_10.Desca " & _
                  " FROM (FILE5_30H INNER JOIN FILE5_30 ON FILE5_30H.DOC_NO = FILE5_30.DOC_NO) INNER JOIN FILE5_10 ON FILE5_30H.BANK = FILE5_10.CODE"
Generalarray(2) = " group by FILE5_30H.Doc_No,FILE5_30.Date,File5_10.Desca order by FILE5_30H.Doc_No"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = " «—ÌŒ-»‰ﬂ-»‰œ «·Õ—ﬂ…"
listarray(0, 1) = "##FILE5_30.Date## " & _
                  " or ( %%file5_10.Desca%% )" & _
                  " Or (FILE5_30.code in (Select FILE5_30.code From File5_30 inner join file5_00 on file5_30.code = file5_00.code where %%FILE5_00.desca%%))"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1200

GrdArray(2, 0) = "≈”„ «·»‰ﬂ"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "Customers Query"
Search3.Show 1
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
myload
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    myload
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub
Private Sub CmdNewInv_Click()
mydefine
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
CardTable.Requery
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
If xdoc_no.Enabled Then
    CmdNewInv_Click
Else
    CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
    myload
End If
End Sub
Private Sub CmdUndo_Click()
If CardTable.RecordCount = 0 Then
    mydefine
    Exit Sub
End If
CardTable.Find "Doc_No = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then
    CardTable.MoveLast
    myload
Else
    myload
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
cStrBox = StrBox
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT DOC_NO,BANK FROM FILE5_30H ORDER BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText

data1.ConnectionString = strCon
data1.RecordSource = "FILE5_10"

Set grid1.DataSource = DATA10
DATA10.ConnectionString = strCon

Set xBank.RowSource = data1
xBank.ListField = "Desca"
xBank.BoundColumn = "CODE"

If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    Fixgrd
    mydefine
End If
End Sub
Sub dispProc()
formMode = dispMode
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
On Error Resume Next
Unload SearchItems
Set SearchItems = Nothing
Err.Clear
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 1 Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And Col = 0 Then BankItemsLookup
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from FILE5_30 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Col = 0 Then BankItemsLookup
End Sub
Private Function MYVALID() As Boolean
CardTable.Find "Doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF And xdoc_no.Enabled Then
    MsgBox "›« Ê—… »‰›” «·—ﬁ„ „‰ ﬁ»·"
    Exit Function
End If

If xdoc_no.text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If xBank.BoundText = "" Then
    MsgBox "«·»‰ﬂ €Ì— „”Ã·"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "»‰Êœ €Ì— „”Ã·…"
    Exit Function
End If


With grid1
For i = 1 To .Rows - 2
'    If Trim(.TextMatrix(I, 2)) = "" Then
'        MsgBox "«·Œ“«‰… €Ì— „ÊÃÊœ…"
'        Exit Function
'    End If

'    If .TextMatrix(I, 0) = "" Then
'        .Select I, 0, I, grid1.Cols - 1
'        MsgBox "ﬂÊœ «·»‰œ €Ì— „ÊÃÊœ"
'        Exit Function
'    Else
If .TextMatrix(i, 0) <> "" Then
    If GetDesca("SELECT CODE FROM FILE5_00 WHERE CODE = " & MyParn(.TextMatrix(i, 0)), con) = "" Then
        .Select i, 0, i, 2
        MsgBox "ﬂÊœ «·»‰œ €Ì— ”·Ì„"
        Exit Function
    End If
End If
    If Val(.TextMatrix(i, 5)) = 0 And Val(.TextMatrix(i, 6)) = 0 Then
        MsgBox "ﬁÌ„… «·»‰œ €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
xdoc_no.text = CardTable!doc_no
xBank.BoundText = CardTable!BANK
Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT FILE5_30.Code,file5_00.Desca,convert(varchar(10),DATE,111),BOX,FILE5_30.DESCA,VALUE1,VALUE2,ID " & _
               " FROM FILE5_30 LEFT JOIN FILE5_00 ON FILE5_30.CODE = FILE5_00.CODE " & _
               " where Doc_no = " & MyParn(xdoc_no.text) & " Order by Row"

    DATA10.RecordSource = cString
    DATA10.Refresh
    grid1.AddItem ""
End With
Fixgrd
End Sub
Private Sub mydefine()

xdoc_no.text = RetZero(Val(Newflag("FILE5_30H", "doc_no", con)))
xBank.BoundText = ""
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode And bedit)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xdoc_no.Enabled = (nMode = DefineMode)
cmdSave.Enabled = bedit
CmdDelInv.Enabled = bedit
End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 5 And Val(.TextMatrix(Row, 6)) <> 0 And Val(.EditText) <> 0 Then
    .TextMatrix(Row, 6) = ""
ElseIf Col = 6 And Val(.TextMatrix(Row, 5)) <> 0 And Val(.EditText) <> 0 Then
    .TextMatrix(Row, 5) = ""
End If
End With
End Sub

Private Sub xDoc_No_LostFocus()
If xdoc_no.text = "" Then Exit Sub
xdoc_no.text = RetZero(xdoc_no.text)
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xdoc_no.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function StrBox()
Dim boxtable As ADODB.Recordset
Set boxtable = New ADODB.Recordset
boxtable.Open "SELECT * FROM file0_50 ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
If Not (boxtable.EOF And boxtable.BOF) Then
    StrBox = "#  " & ";       "
    Do Until boxtable.EOF
        StrBox = StrBox & "|#" & boxtable!CODE & ";" & boxtable!DESCA
        boxtable.MoveNext
    Loop
End If
End Function
Private Sub BankItemsLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE5_00 where isstop = 0 "
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = True

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
SearchItems.Caption = "≈” ⁄·«„ "
SearchItems.Show 1
End Sub
Private Sub Fixgrd()
With grid1
    .MergeCells = flexMergeFree
    .MergeRow(0) = True
     .FormatString = "«·Õ—ﬂ…|" & "«·Õ—ﬂ…|" & "«· «—ÌŒ|" & "Œ“‰…|" & "«·»Ì«‰|" & "«Ìœ«⁄« |" & "„”ÕÊ»« |"
    .ColWidth(0) = 500
    .ColWidth(1) = 2000
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 6000
    .ColWidth(5) = 1200
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    .ColHidden(.Cols - 1) = True
    For i = 1 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColComboList(3) = cStrBox
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
        'CalcTotals
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 Then
    .RemoveItem .Row
    'CalcTotals
End If
End With
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Not (IsNumeric(.TextMatrix(nRow, 0)) And IsDate(.TextMatrix(nRow, 2)) And (Val(.TextMatrix(nRow, 5)) <> 0 Or Val(.TextMatrix(nRow, 6)) <> 0)) Then Exit Function
'If Not (IsNumeric(.TextMatrix(nRow, 0))) Then Exit Function
validRow = True
End With
validRow = True
End Function
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Col = 0 Then GrdDesc Row
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
    .AddItem ""
    'CalcTotals
End If
End With
End Sub
Private Sub GrdDesc(nRow)
grid1.TextMatrix(nRow, 1) = ""
If IsNumeric(grid1.TextMatrix(nRow, 0)) Then grid1.TextMatrix(nRow, 1) = GetDesca("select Desca from file5_00 where code = " & MyParn(grid1.TextMatrix(nRow, 0)), con)
End Sub


Private Sub CMD_PRINT_Click()
Dim aHeader(2)
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
'On Error GoTo myerror
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
For i = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!STR21 = "„” ‰œ ”Õ» / ≈Ìœ«⁄ »‰ﬂÏ " & xdoc_no.text
    temptable!STR7 = xdoc_no.text
    temptable!str1 = cComp_Name
    temptable!VAL1 = grid1.TextMatrix(i, 5)
    temptable!val2 = grid1.TextMatrix(i, 6)
    temptable!STR7 = TurnValue(grid1.TextMatrix(i, 2), "", Null)
    temptable!str3 = TurnValue(grid1.TextMatrix(i, 4), "", Null)
    temptable!str8 = TurnValue(GetDesca("select desca from file0_50 where code = " & MyParn(grid1.TextMatrix(1, 3)), con), "", Null)
    temptable!str5 = TurnValue(xBank.text, "", Null)
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans

main.Report1.ReportFileName = App.Path & "\Reports\P_CASH_B.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\Temp.MDB"
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear

End Sub

