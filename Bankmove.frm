VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form BankMovefrm 
   Caption         =   "»Ì«‰ Õ—ﬂ… «·»‰Êﬂ"
   ClientHeight    =   10455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13260
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
   ScaleHeight     =   10455
   ScaleWidth      =   13260
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD_PRINT 
      Caption         =   "ÿ»«⁄…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5760
      TabIndex        =   15
      Top             =   1170
      Width           =   1590
   End
   Begin VB.CommandButton Cmdgo 
      Caption         =   "«” Ã«»…"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   1170
      Width           =   1500
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   13260
      TabIndex        =   9
      Top             =   0
      Width           =   13260
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00000000&
         Caption         =   "Œ—ÊÃ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   75
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   45
         Width           =   1440
      End
   End
   Begin VB.Frame Frame3 
      ForeColor       =   &H8000000D&
      Height          =   1050
      Left            =   9450
      TabIndex        =   4
      Top             =   540
      Width           =   5700
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   1365
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo xBank 
         Height          =   315
         Left            =   90
         TabIndex        =   0
         Top             =   225
         Width           =   4740
         _ExtentX        =   8361
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰ :"
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
         Left            =   4965
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   705
         Width           =   330
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ï :"
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
         Left            =   1545
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   705
         Width           =   450
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Left            =   4950
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   270
         Width           =   480
      End
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1920
      Width           =   405
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Õ—ﬂ… ÃœÌœ…"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1575
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   75
      Width           =   1215
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   8775
      Left            =   45
      TabIndex        =   8
      Top             =   1620
      Width           =   15120
      _cx             =   26670
      _cy             =   15478
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
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
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
   Begin Threed.SSCommand cmd_excel 
      Height          =   465
      Left            =   90
      TabIndex        =   13
      Top             =   990
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   13822956
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Bankmove.frx":0000
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin ComctlLib.ProgressBar prog1 
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   630
      Width           =   3750
      _ExtentX        =   6615
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "BankMovefrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim DocTable As ADODB.Recordset
Dim formMode
Sub Fillgrd()
Dim nPrevious As Double
Grid1.Rows = 1
If DocTable.State = adStateOpen Then DocTable.Close
If IsDate(xDate1.Text) Then
    Dim loctable As New ADODB.Recordset
    loctable.Open "select sum(value1 - value2) as Balance from bankmove where " & _
                  " TYPE <= 4.5 " & _
                  " AND bank = " & MyParn(xBank.BoundText) & _
                  " and [date] < " & DateSq(xDate1.Text), con, adOpenStatic, adLockReadOnly
    If Not loctable.EOF Then nPrevious = Val(loctable!balance & "")
    If nPrevious <> 0 Then
        Grid1.AddItem ""
        Grid1.TextMatrix(Grid1.Rows - 1, 0) = Format(DateAdd("d", -1, xDate1.Text), "dd-mm-yyyy")
        Grid1.TextMatrix(Grid1.Rows - 1, 2) = "—’Ìœ ”«»ﬁ"
        Grid1.TextMatrix(Grid1.Rows - 1, 6) = Format(loctable!balance, "#0.00")
    End If
    loctable.Close
    Set loctable = Nothing
End If

cString = "Select * from BankMove Where BANK = " & MyParn(xBank.BoundText) & _
          " and [type] <= 4.5 "
If IsDate(xDate1.Text) Then cString = cString & " and date >= " & DateSq(xDate1.Text)
If IsDate(xDate2.Text) Then cString = cString & " and date <= " & DateSq(xDate2.Text)

cString = cString & " Order by [Date],BRANCH, VALUE1,VALUE2"
DocTable.Open cString, con, adOpenStatic, , adCmdText

If DocTable.EOF And DocTable.BOF Then Exit Sub
With Grid1
Do
   Grid1.AddItem ""
   .TextMatrix(.Rows - 1, 0) = Format(DocTable![Date], "dd-mm-yyyy")
   .TextMatrix(.Rows - 1, 7) = DocTable!doc_no & ""
   .TextMatrix(.Rows - 1, 2) = DocTable!TypeDesca & ""
   .TextMatrix(.Rows - 1, 3) = DocTable!DESCA & ""
   .TextMatrix(.Rows - 1, 4) = Format(TurnValue(DocTable!value1 & "", "0", ""), "#0.00")
   .TextMatrix(.Rows - 1, 5) = Format(TurnValue(DocTable!Value2 & "", "0", ""), "#0.00")
   .TextMatrix(.Rows - 1, 6) = Format(nPrevious + Val(DocTable!value1 & "") - Val(DocTable!Value2 & ""), "#0.00")
   .TextMatrix(.Rows - 1, 1) = DocTable!branch & ""
   
   nPrevious = Format(nPrevious + Val(DocTable!value1 & "") - Val(DocTable!Value2 & ""), "Fixed")
   DocTable.MoveNext
   i = i + 1
Loop Until DocTable.EOF
.SubtotalPosition = flexSTBelow
.Subtotal flexSTNone, 0, 1, " ", &H8000000F
.Subtotal flexSTSum, 1, 4, "#0.00", &H80000010, , True, " "
.Subtotal flexSTSum, 1, 5, "#0.00", &H80000010, , True, " "

End With
End Sub
Sub myProc()
ActiveControl.Text = GrdText(Search3.Grid1, 0)
Unload Search
End Sub
Function MYVALID()
If xBank.Text = "" Then Exit Function
MYVALID = True
End Function
Private Sub cmd_excel_Click()
    ToFileExel2 Grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub CMD_PRINT_Click()
    Set printGrd3.Grid1 = Grid1
    Set printGrd3.myForm = Me
    printGrd3.doprint 1, 0, " ›’Ì·Ï Õ—ﬂ… »‰ﬂ " & xBank.Text, " „‰  «—ÌŒ " & xDate1.Text & " Õ Ï " & xDate2.Text, , False, True, 10, 1
    printGrd3.Show 1
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    If Not MYVALID Then Exit Sub
    Fillgrd
End Sub
Private Sub cmdNew_Click()
Grid1.Rows = 1
cmdNew.Enabled = False
Cmdgo.Enabled = True
xBank.Enabled = True
xBankName.Caption = ""
xBank.SetFocus
End Sub
Private Sub Form_Load()
openCon con
Set DocTable = New ADODB.Recordset
data1.ConnectionString = strCon
data1.RecordSource = "FILE5_10"

Set xBank.RowSource = data1
xBank.ListField = "Desca"
xBank.BoundColumn = "code"

'                       0           1           2           3       4               5           6       7
Grid1.FormatString = " «—ÌŒ |" & "›—⁄|" & "‰Ê⁄ |" & "»Ì«‰ |" & "«Ìœ«⁄« |" & "„”ÕÊ»« |" & "—’Ìœ|" & "—ﬁ„ «·„” ‰œ|"
Grid1.Cols = 8
Grid1.ColWidth(0) = 1300
Grid1.ColWidth(7) = 1500
Grid1.ColWidth(2) = 3500
Grid1.ColWidth(3) = 4000
Grid1.ColWidth(4) = 1400
Grid1.ColWidth(5) = 1400
Grid1.ColWidth(6) = 1400
Grid1.ColHidden(1) = True

For i = 0 To Grid1.Cols - 1
    Grid1.ColAlignment(i) = flexAlignRightCenter
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
