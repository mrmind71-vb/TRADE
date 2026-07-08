VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form chargefrm3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁœÌ…"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13875
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   8160
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   7155
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   90
         TabIndex        =   18
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
         _Version        =   196610
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
         Picture         =   "charge2.frx":0000
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge2.frx":21D0
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   870
         TabIndex        =   19
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
         _Version        =   196610
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
         Picture         =   "charge2.frx":4318
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge2.frx":64E0
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1650
         TabIndex        =   20
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
         _Version        =   196610
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
         Picture         =   "charge2.frx":862F
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge2.frx":A80F
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   2430
         TabIndex        =   21
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   820
         _Version        =   196610
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
         Picture         =   "charge2.frx":C96A
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "charge2.frx":EB26
      End
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   8325
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   5460
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4095
         Picture         =   "charge2.frx":10C75
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   2790
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge2.frx":13448
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge2.frx":159F4
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge2.frx":1828E
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   7515
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   1500
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge2.frx":183D8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "charge2.frx":1A951
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   7
      Top             =   7860
      Width           =   13875
      _ExtentX        =   24474
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
            TextSave        =   "10:56 ’"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame6 
      Height          =   540
      Left            =   2565
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   3630
      Begin VB.TextBox xusername 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   3465
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   9045
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   720
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
         Left            =   2340
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
         Left            =   2340
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   4
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   255
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   465
      Left            =   0
      Top             =   0
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
      Height          =   5325
      Left            =   90
      TabIndex        =   16
      Top             =   1800
      Width           =   15135
      _cx             =   26696
      _cy             =   9393
      _ConvInfo       =   1
      Appearance      =   0
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
      BackColorFixed  =   14737632
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
      Cols            =   10
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
End
Attribute VB_Name = "chargefrm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte, bEdit As Boolean
Dim cStrBox As String
Dim CardTable As ADODB.Recordset, GRDTABLE As New ADODB.Recordset
Dim cFile As String, cFileHeader As String, sName As String
Dim docMoveType As String
Dim DocTitle As String
Dim DocClient As String, CGROUP As String
Dim dLastdate As String, defBox As String
Dim DocField As String, dDateLast As String
Dim formMode
Dim con As New ADODB.Connection
Dim lCellButton As Boolean
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(2, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.Text)

aInsert(2, 0) = "userName"
aInsert(2, 1) = addstring(sUserName)


Dim cString As String
cString = "SELECT DOC_NO FROM " & cFileHeader
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.Text)
xDoc_No.Enabled = GetDesca(cString) = ""
On Error GoTo myError
con.BeginTrans
If xDoc_No.Enabled Then
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, cFileHeader)
Else
    con.Execute CreateUpdate(aInsert, cFileHeader, " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd()
Dim aInsert(5, 1)
With grid1
    For i = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
               
        aInsert(1, 0) = "Box"
        aInsert(1, 1) = addstring(.TextMatrix(i, 0))
        
        aInsert(2, 0) = "Charge"
        aInsert(2, 1) = addstring(grid1.TextMatrix(i, 1))
        
        aInsert(3, 0) = "Desca"
        aInsert(3, 1) = addstring(grid1.TextMatrix(i, 3))
        
        aInsert(4, 0) = "[value]"
        aInsert(4, 1) = Val(grid1.TextMatrix(i, 4))

        aInsert(5, 0) = "[Row]"
        aInsert(5, 1) = i

        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, cFile)
        Else
            con.Execute CreateUpdate(aInsert, cFile, " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 1 Then
        grid1.TextMatrix(grid1.Row, 1) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
        GrdDesc grid1.Row
        If grid1.Row = grid1.Rows - 1 And validRow(grid1.Row) Then
            myAddItem
        End If
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
    On Error GoTo myError
    con.BeginTrans
    con.Execute "Delete  From " & cFile & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From " & cFileHeader & " where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    openCardTable
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myError:
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
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT " & cFileHeader & ".Doc_No, Convert(Varchar," & cFileHeader & ".Date,111),Min(" & DocClient & ".Desca)" & _
          " FROM (" & cFileHeader & " inner join " & cFile & " on " & cFileHeader & ".doc_no = " & cFile & ".Doc_NO) Inner Join " & DocClient & " on " & cFile & ".Charge = " & DocClient & ".Code"
          
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
Load Search3
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
myUndo
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
openCardTable
myUndo
End Sub
Private Sub CmdUndo_Click()
openCardTable
myUndo
End Sub
Private Sub Form_Activate()
On Error Resume Next
grid1.SetFocus
Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
Select Case myPublic
    Case 1 '„’«—Ì›
        sName = "«·„’—Ê›"
        cFile = "File8_50"
        cFileHeader = "FILE8_50H"
        DocClient = "FILE8_51"
    Case 2 '«·«Ì—«œ
        sName = "«·«Ì—«œ"
        cFile = "File8_60"
        cFileHeader = "FILE8_60H"
        DocClient = "FILE8_61"
End Select
Me.Caption = DocTitle

cStrBox = StrBox
Set grid1.DataSource = DATA1
DATA1.ConnectionString = strCon
Fixgrd
openCardTable
xDoc_No.Text = Format(Date, "DDMMYY")
myUndo
If xDoc_No.Text <> Format(Date, "DDMMYY") Then mydefine
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
GRDTABLE.Close
Set CardTable = Nothing
Set GRDTABLE = Nothing
closeCon con
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 2 Or grid1.Col = 7 Or (cManBox <> "" And grid1.Col = 0) Then grid1.Editable = flexEDNone Else grid1.Editable = flexEDKbdMouse
End Sub

Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.SetFocus
    grid1.Select 1, 0
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 1 Then grdLookup
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 And bEdit Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myError
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from " & cFile & " where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        grid1.RemoveItem grid1.Row
        grid1_EnterCell
    End If
End If
Exit Sub
myError:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Col = 2 Then grdLookup
If KeyCode = 46 And Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        grid1.RemoveItem Row
        grid1_EnterCell
    End If
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case grid1.Col
    Case 0
        If KeyCode = 112 Then grdLookup
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

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If

With grid1
For i = 1 To .Rows - 2
    If .TextMatrix(i, 1) = "" Then
        .Select i, 0, i, grid1.Cols - 1
        MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
        Exit Function
    End If
    If Val(.TextMatrix(i, 4)) = 0 Then
        MsgBox "«·ﬁÌ„… €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xusername.Text = CardTable!UserName & ""
Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT " & cFile & ".[BOX], " & cFile & ".Charge," & DocClient & ".DESCA," & cFile & ".desca,[VALUE],[ID] " & _
               " FROM " & cFile & " LEFT JOIN " & DocClient & " ON " & cFile & ".Charge = " & DocClient & ".CODE "
    cString = cString & turn(cString) & cFile & ".Doc_no = " & MyParn(xDoc_No.Text)
    If cManBox <> "" Then
        cString = cString & turn(cString) & " BOX = " & MyParn(cManBox)
    End If
    DATA1.RecordSource = cString
    DATA1.Refresh
    myAddItem
End With
CalcTotals
Fixgrd
End Sub
Private Sub mydefine()
xDoc_No.Text = Format(Date, "DDMMYY")
xDate.Text = Format(Date, "DD-MM-YYYY")
grid1.Rows = 1
myAddItem
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
    If Col = 1 And Trim(.EditText) <> "" And IsNumeric(.EditText) Then .EditText = RetZero(.EditText, 3)
End With
End Sub

Private Sub xDoc_No_LostFocus()
If Trim(xDoc_No.Text) = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
CardTable.Find "Doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function StrBox()
Dim boxtable As ADODB.Recordset
Set boxtable = New ADODB.Recordset
boxtable.Open "SELECT * FROM file0_50 ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
If Not (boxtable.EOF And boxtable.BOF) Then
    boxtable.MoveFirst
    StrBox = "#  " & ";       "
    Do Until boxtable.EOF
        StrBox = StrBox & "|#" & boxtable!CODE & ";" & boxtable!DESCA
        boxtable.MoveNext
    Loop
End If
End Function
Private Sub grdLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From " & DocClient
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Private Function CalcTotals()
Dim nTotal As Double
With grid1
For i = 1 To grid1.Rows - 2
    nTotal = nTotal + Round(Val(grid1.TextMatrix(i, 4)), 2)
Next
StatusBar1.Panels(1).Text = "«·«Ã„«·Ì : " & Format(nTotal, "Fixed")
End With
End Function
Private Sub GrdDesc(nRow)
grid1.TextMatrix(nRow, 2) = GetDesca("Select Desca From " & DocClient & " Where code = " & MyParn(grid1.TextMatrix(nRow, 1))) & ""
End Sub
Private Function RetDefBox() As String
Dim loctable As New ADODB.Recordset
loctable.Open "file0_50", con, adOpenStatic, adLockReadOnly, adCmdTable
If loctable.EOF And loctable.BOF Then Exit Function
loctable.MoveLast
If loctable.RecordCount = 1 Then
    loctable.MoveFirst
    RetDefBox = Trim(loctable!CODE & "")
End If
End Function

Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.Text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With grid1
    If myPublic = 1 Or myPublic = 3 Then
        .FormatString = "Œ“‰…|" & "«·„’—Ê›|" & "Ê’› «·„’—Ê›|" & "«·»Ì«‰|" & "«·ﬁÌ„…|"
    Else
        .FormatString = "Œ“‰…|" & "«·«Ì—«œ|" & "Ê’› «·«Ì—«œ|" & "«·»Ì«‰|" & "«·ﬁÌ„…|"
    End If
    .ColWidth(0) = 1300
    .ColWidth(1) = 1000
    .ColWidth(2) = 2900
    .ColWidth(3) = 4000
    .ColWidth(4) = 1000
    .ColHidden(.Cols - 1) = True
    For i = 1 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColComboList(0) = cStrBox
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, .Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
        CalcTotals
    End If
End If
End With
End Sub
Private Sub grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 And .TextMatrix(.Row, .Cols - 1) = "" Then
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
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Col = 1 Then GrdDesc Row
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
    myAddItem
    CalcTotals
End If
End With
End Sub
Private Sub myAddItem()
With grid1
.AddItem ""
.TextMatrix(.Rows - 1, 0) = cManBox
End With
End Sub
Private Sub myUndo()
On Error GoTo myError
If CardTable.BOF And CardTable.EOF Then
    mydefine
Else
    If xDoc_No.Text <> "" Then
        CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
        If CardTable.EOF Then CardTable.MoveLast
    Else
        CardTable.MoveLast
    End If
    myload
End If
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cString = "SELECT * FROM " & cFileHeader
If cManBox <> "" Then cString = cString & turn(cString) & "DOC_NO = " & MyParn(Format(Date, "DDMMYY"))
'If sDoc_no <> "" Then cString = cString & turn(cString) & "DOC_NO = " & MyParn(sDoc_no)
cString = cString & " ORDER BY DOC_NO"
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub


