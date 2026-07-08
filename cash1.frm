VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Cashfrm1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁœÌ…"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14955
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
   ScaleHeight     =   8085
   ScaleWidth      =   14955
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   8550
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   585
      Width           =   1500
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cash1.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cash1.frx":0442
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   510
      ScaleWidth      =   14955
      TabIndex        =   10
      Top             =   7275
      Width           =   14955
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   -45
         Width           =   3975
         Begin VB.CommandButton cmdFirst 
            Height          =   375
            Left            =   45
            Picture         =   "cash1.frx":29BB
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "√Ê·"
            Top             =   135
            Width           =   870
         End
         Begin VB.CommandButton cmdLast 
            Height          =   375
            Left            =   915
            Picture         =   "cash1.frx":509A
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "√ŒÌ—"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdPrevious 
            Height          =   375
            Left            =   2070
            Picture         =   "cash1.frx":7774
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "«·”«»ﬁ"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdNext 
            Height          =   375
            Left            =   2985
            Picture         =   "cash1.frx":9D47
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "«· «·Ì"
            Top             =   135
            Width           =   915
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   9
      Top             =   7785
      Width           =   14955
      _ExtentX        =   26379
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
            TextSave        =   "10:54 ’"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Height          =   615
      Left            =   2745
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   45
      Visible         =   0   'False
      Width           =   3630
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   3465
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   5505
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4095
         Picture         =   "cash1.frx":C307
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   2790
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cash1.frx":EADA
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cash1.frx":11086
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "cash1.frx":13920
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   585
      Width           =   4740
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
         Height          =   315
         Left            =   2295
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1290
      End
      Begin VB.TextBox xDate 
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
         Height          =   315
         Left            =   2295
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   585
         Width           =   1290
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   600
         Width           =   600
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
         Left            =   3675
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   255
         Width           =   930
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5625
      Left            =   90
      TabIndex        =   2
      Top             =   1665
      Width           =   14775
      _cx             =   26061
      _cy             =   9922
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   495
      Top             =   495
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
End
Attribute VB_Name = "Cashfrm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte, bEdit As Boolean
Dim con As New ADODB.Connection
Dim CardTable As ADODB.Recordset, cStrBox As String
Dim cFile As String, cFileHeader As String, sName As String
Dim DocTitle As String
Dim DocClient As String
Dim dLastdate As String, defBox As String
Dim formMode
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(2, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.Text)

aInsert(2, 0) = "userName"
aInsert(2, 1) = addstring(sUserName)

On Error GoTo myerror
con.BeginTrans
If xDoc_No.Enabled Then
    xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
    con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xDoc_No.Text))
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
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 1 Then
        grid1.TextMatrix(grid1.Row, 1) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
        GrdDesc grid1.Row
        If grid1.Row = grid1.Rows - 1 And validRow(grid1.Row) Then
            grid1.AddItem ""
            grid1.TextMatrix(grid1.Rows - 1, 0) = defBox
        End If
        Unload Search3
    ElseIf grid1.Col = 5 Then
        grid1.TextMatrix(grid1.Row, 5) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 6) = Search3.grid1.TextMatrix(Search3.grid1.Row, 3)
        grid1.TextMatrix(grid1.Row, 7) = Search3.grid1.TextMatrix(Search3.grid1.Row, 4)
        Unload Search3
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
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    CardTable.Requery
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT " & cFileHeader & ".Doc_No, Convert(VarChar(10)," & cFileHeader & ".Date,111),Min(" & DocClient & ".Desca)" & _
          " FROM (" & cFileHeader & " inner join " & cFile & " on " & cFileHeader & ".doc_no = " & cFile & ".Doc_NO) Inner Join " & DocClient & " on " & cFile & ".Code = " & DocClient & ".Code"
          
Generalarray(1) = cString
Generalarray(2) = " group by " & cFileHeader & ".Doc_No," & cFileHeader & ".Date order by " & cFileHeader & ".Doc_No," & cFileHeader & ".Date"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·«”„- «—ÌŒ «·„” ‰œ"
listarray(0, 1) = "(%%" & DocClient & ".Desca%% or " & _
                  " ##" & cFileHeader & ".Date##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·≈”„"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub
Private Sub CmdInform_Click()
CardLookup
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
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    mydefine
Else
    If xDoc_No.Enabled Then CardTable.MoveLast Else CardTable.Find "Doc_No = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
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
defBox = RetDefBox
sName = "«·⁄„Ì·"
cFile = "File8_10"
cFileHeader = "FILE8_10H"
DocTitle = "„ﬁ»Ê÷«  „‰ ⁄„·«¡"
DocClient = "File3_10"
DocField = "pay"
Me.Caption = DocTitle
cStrBox = StrBox
Set grid1.DataSource = data1
data1.ConnectionString = strCon
Set CardTable = New ADODB.Recordset
CardTable.Open "SELECT * FROM FILE8_10H ORDER BY DOC_NO", con, adOpenStatic, adLockOptimistic, adCmdText
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    mydefine
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub Grid1_EnterCell()
If grid1.Col = 2 Or grid1.Col = 6 Or grid1.Col = 7 Then grid1.Editable = flexEDNone Else grid1.Editable = flexEDKbdMouse
End Sub

Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.SetFocus
    grid1.Select 1, 0
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 1 Then CLIENTLOOKUP
If KeyCode = 112 And grid1.Col = 5 And Trim(grid1.TextMatrix(Row, 1)) <> "" Then InvLookup
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 And bEdit Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from file8_10 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        grid1.RemoveItem grid1.Row
        CalcTotals
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Col = 2 Then CLIENTLOOKUP
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case grid1.Col
    Case 0
        If KeyCode = 27 Then
            Exit Sub
        End If
        If KeyCode = 112 Then
            CLIENTLOOKUP
        End If
End Select
End Sub
Private Function MYVALID() As Boolean
If Trim(xDoc_No.Text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If IsDate(dLastdate) Then
    If DateValue(xDate.Text) <= DateValue(dLastdate) Then
        MsgBox "«· «—ÌŒ «ﬁ· „‰ «Œ—  «—ÌŒ «€·«ﬁ"
        Exit Function
    End If
End If
If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ »‰Êœ  „  ”ÃÌ·Â«"
    Exit Function
End If



With grid1
For I = 1 To .Rows - 2
    If .TextMatrix(I, 1) = "" Then
        .Select I, 0, I, grid1.Cols - 1
        MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
        Exit Function
    End If
    If Val(.TextMatrix(I, 4)) = 0 Then
        MsgBox "ﬁÌ„… «·»‰œ €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
Dim GRDTABLE As New ADODB.Recordset
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xusername.Text = TurnValue(CardTable!UserName, Null, "")
Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub mydefine()
xDoc_No.Text = RetZero(Val(Newflag(cFileHeader, "doc_no")))
xDate.Text = Format(Date, "dd-mm-yyyy")
xusername.Text = ""
grid1.Rows = 1
grid1.AddItem ""
Fixgrd
grid1.TextMatrix(grid1.Rows - 1, 0) = defBox
Handlecontrols DefineMode
CalcTotals
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode And bEdit)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
CmdDelInv.Enabled = (nMode = LoadMode) And bEdit
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
CmdSave.Enabled = bEdit
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
    If Col = 1 And Trim(.EditText) <> "" And IsNumeric(.EditText) Then .EditText = RetZero(.EditText, 6)
End With
End Sub

Private Sub xDoc_No_LostFocus()
If Trim(xDoc_No.Text) = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function StrBox()
Dim boxtable As New ADODB.Recordset
boxtable.Open "SELECT * FROM file0_50 ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
StrBox = "#  " & ";       "
Do Until boxtable.EOF
    StrBox = StrBox & "|#" & boxtable!Code & ";" & boxtable!Desca
    boxtable.MoveNext
Loop
End Function
Private Sub CLIENTLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From " & DocClient
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Private Function CalcTotals()
Dim nTotal As Double
With grid1
For I = 1 To grid1.Rows - 2
    nTotal = nTotal + Round(Val(grid1.TextMatrix(I, 4)), 2)
Next
StatusBar1.Panels(1).Text = "«·«Ã„«·Ì : " & Format(nTotal, "Fixed")
End With
End Function
Private Sub GrdDesc(nRow)
grid1.TextMatrix(nRow, 2) = GetDesca("Select Desca From FILE3_10 Where code = " & MyParn(grid1.TextMatrix(nRow, 1))) & ""
End Sub
Private Function RetDefBox() As String
Dim loctable As New ADODB.Recordset
loctable.Open "file0_50", con, adOpenStatic, adLockReadOnly, adCmdTable
If loctable.EOF And loctable.BOF Then Exit Function
loctable.MoveLast
If loctable.RecordCount = 1 Then
    loctable.MoveFirst
    RetDefBox = Trim(loctable!Code & "")
End If
End Function

Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.Text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With grid1
    .Editable = flexEDKbdMouse
    .Cols = 6
    .FormatString = "Œ“‰…|" & "ﬂÊœ|" & "«·⁄„Ì·|" & "«·»Ì«‰|" & "«·ﬁÌ„…|"
    .ColWidth(0) = 1300
    .ColWidth(1) = 1000
    .ColWidth(2) = 2900
    .ColWidth(3) = 4000
    .ColWidth(4) = 1000

    .ColHidden(grid1.Cols - 1) = True
    
    For I = 1 To grid1.Cols - 1
        .ColAlignment(I) = flexAlignRightCenter
    Next
    .ColComboList(0) = cStrBox
End With
End Sub
Private Sub InvLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select doc_no,Convert(VarChar(10),[date],111),total,pay,total - pay from invpaid where code = " & MyParn(grid1.TextMatrix(grid1.Row, 1)) & " and total - pay > 0"
Generalarray(2) = "Order by date"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = " «—ÌŒ «·›« Ê—…-—ﬁ„ «·›« Ê—…"
listarray(0, 1) = "(%%date%%) or (**total**)"

GrdArray(0, 0) = "≈Ã„«·Ì «·›« Ê—…"
GrdArray(0, 1) = 1500

GrdArray(1, 0) = " «—ÌŒ «·›« Ê—…"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "≈Ã„«·Ì «·›« Ê—…"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·„”œœ"
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "«·„ »ﬁÌ"
GrdArray(4, 1) = 1500

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Private Sub myreplaceGrd()
Dim aInsert(6, 1)
With grid1
    For I = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
                
        aInsert(1, 0) = "Box"
        aInsert(1, 1) = addstring(.TextMatrix(I, 0))
        
        aInsert(2, 0) = "code"
        aInsert(2, 1) = addstring(grid1.TextMatrix(I, 1))
        
        aInsert(3, 0) = "Desca"
        aInsert(3, 1) = addstring(grid1.TextMatrix(I, 3))
        
        aInsert(4, 0) = "[value]"
        aInsert(4, 1) = Val(grid1.TextMatrix(I, 4))

        aInsert(5, 0) = "[inv_no]"
        aInsert(5, 1) = addstring(grid1.TextMatrix(I, 5))

        aInsert(6, 0) = "row"
        aInsert(6, 1) = I
        
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, cFile)
        Else
            con.Execute CreateUpdate(aInsert, cFile, " where ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub myloadgrd()
With grid1
'    cString = "SELECT FILE8_10.[BOX], FILE8_10.CODE, FILE3_10.DESCA, FILE8_10.desca, FILE8_10.VALUE, FILE8_10.INV_NO," & _
'              " INVPAID.PAY,INVPAID.TOTAL - INVPAID.PAY,FILE8_10.ID " & _
'              " FROM (FILE8_10 LEFT JOIN FILE3_10 ON FILE8_10.CODE = FILE3_10.CODE) LEFT JOIN INVPAID ON FILE8_10.INV_NO = INVPAID.DOC_NO" & _
'              " where file8_10.doc_no = " & MyParn(xDoc_No.Text) & " Order By Row"

    cString = "SELECT FILE8_10.[BOX], FILE8_10.CODE, FILE3_10.DESCA, FILE8_10.desca, FILE8_10.VALUE," & _
              "FILE8_10.ID " & _
              " FROM FILE8_10 LEFT JOIN FILE3_10 ON FILE8_10.CODE = FILE3_10.CODE" & _
              " where file8_10.doc_no = " & MyParn(xDoc_No.Text) & " Order By Row"

    data1.RecordSource = cString
    data1.Refresh
    grid1.AddItem ""
    .TextMatrix(.Rows - 1, 0) = defBox
End With
CalcTotals
Fixgrd
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
        CalcTotals
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 Then
    .RemoveItem .Row
    CalcTotals
End If
End With
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Not (IsNumeric(.TextMatrix(nRow, 1)) And Val(.TextMatrix(nRow, 4)) <> 0) Then Exit Function
End With
validRow = True
End Function
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Col = 1 Then GrdDesc Row
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
    .AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = defBox
    CalcTotals
End If
End With
End Sub

