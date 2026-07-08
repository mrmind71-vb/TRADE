VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpSup32 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
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
   ScaleHeight     =   4200
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3435
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   5775
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Caption         =   " ﬁ”Ì„ «·›Ê« Ì— Õ”» «·„Ê—œ"
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
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   3060
         Width           =   3210
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1890
         Width           =   1680
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   2250
         Width           =   1680
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin VSFlex7LCtl.VSFlexGrid grid1 
         Height          =   450
         Left            =   990
         TabIndex        =   1
         Top             =   630
         Width           =   3765
         _cx             =   6641
         _cy             =   794
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   0
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
         BackColorBkg    =   -2147483633
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
         Rows            =   10
         Cols            =   1
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"rpSup32.frx":0000
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
         Editable        =   2
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
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   1305
         TabIndex        =   12
         Top             =   2655
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   1035
         TabIndex        =   15
         Top             =   1215
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "«·›—⁄ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1305
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "«·„Ê”„ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2745
         Width           =   1680
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1980
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2340
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
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
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   315
         Width           =   570
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   3345
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·„Ã„Ê⁄… :"
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
         Left            =   4830
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   630
         Width           =   840
      End
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   375
      Left            =   1215
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3510
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3510
      Width           =   1140
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data3 
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
   Begin MSAdodcLib.Adodc DATA2 
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
End
Attribute VB_Name = "rpSup32"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If TypeOf ActiveControl Is DBCombo And KeyCode = 46 Then ActiveControl.BoundText = ""
End Sub
Private Sub Form_Load()
openCon con
grdMake "Select Code,DescA From FILE4_50", "code", "desca", con, grid1

DATA2.ConnectionString = strCon
DATA2.RecordSource = "Select MOSM,DescA From MOSM     ORDER by DATE "
Set xMosm.RowSource = DATA2
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"


data3.ConnectionString = strCon
data3.RecordSource = "Select CODE ,DescA From BRANCH_FR "
Set XSTORE.RowSource = data3
XSTORE.ListField = "Desca"
XSTORE.BoundColumn = "CODE"


End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
'If Not IsDate(xdate1.Text) Then
'    MsgBox "«· «—ÌŒ «·«Ê· ÷—Ê—Ì"
'    Exit Function
'End If
If Not IsDate(xDate2.text) And Trim(xDate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(3)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable


cString = "SELECT FR7_10H.DOC_NO,FR7_10H.MOSM ,FR7_10H.INV_NO, FR7_10H.DATE, SUM(FR7_10.TOTAL) as sumOfTotal , SUM(FR7_10.QUANT * FR7_10.PRICE ) as TITEM , FR7_10H.DISCOUNT, FR7_10H.TAX,FR7_10H.CODE,FILE4_10.DESCA,FILE4_10.[GROUP],FILE4_50.DESCA AS GROUPDESCA,BRANCH_FR.DESCA AS STORENAME " & _
         " FROM ((FR7_10H INNER JOIN FR7_10 ON FR7_10H.doc_no =   FR7_10.doc_no) inner join file4_10 on FR7_10H.CODE = FILE4_10.CODE) LEFT JOIN FILE4_50 ON FILE4_10.[GROUP] = FILE4_50.CODE INNER JOIN BRANCH_FR ON BRANCH_FR.CODE = FR7_10H.STORE "

If Trim(xCode.text) <> "" Then
    cWhere = cWhere & turn(cWhere, " AND ") & "FR7_10H.CODE = " & MyParn(xCode.text)
    aHeader(0) = "[" & "«·„Ê—œ : " & xCodeDesca.Caption & "]"
End If
If Trim(xMosm.BoundText) <> "" Then
    cWhere = cWhere & turn(cWhere, " AND ") & "FR7_10H.mosm = " & MyParn(xMosm.BoundText)
    aHeader(1) = "[" & "„Ê”„ : " & xMosm.text & "]"
End If
If Trim(XSTORE.BoundText) <> "" Then
    cWhere = cWhere & turn(cWhere, " AND ") & "FR7_10H.STORE = " & MyParn(XSTORE.BoundText)
    aHeader(2) = "[" & "›—⁄ : " & XSTORE.text & "]"
End If
If GrdQry(grid1, "file4_10.[GROUP]", True) <> "" Then
    cWhere = cWhere & turn(cWhere, " and ") & GrdQry(grid1, "File4_10.[GROUP]", True)
    aHeader(2) = "[" & "„Ã„Ê⁄… „Ê—œÌ‰ : " & GrdTitle(grid1) & "]"
End If

If IsDate(xdate1.text) Then
    cWhere = cWhere & turnFound(cWhere, " and ") & " FR7_10H.date >= " & DateSq(xdate1.text)
    aHeader(3) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

If IsDate(xDate2.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " FR7_10H.date <= " & DateSq(xDate2.text)
    aHeader(3) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

cString = cString & turnFound(cWhere) & cWhere
cString = cString & " GROUP BY BRANCH_FR.DESCA ,  FR7_10H.DOC_NO,FR7_10H.MOSM ,FR7_10H.INV_NO, FR7_10H.DATE, FR7_10H.DISCOUNT ,FR7_10H.TAX,FR7_10H.CODE,FILE4_10.DESCA,FILE4_10.[GROUP],FILE4_50.DESCA"
                        
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str5 = sourcetable!doc_no
    temptable!Date1 = sourcetable!Date
    temptable!str4 = sourcetable!INV_NO
    temptable!str3 = sourcetable!MOSM
    temptable!str1 = sourcetable!CODE
    temptable!str2 = sourcetable!DESCA
    temptable!VAL1 = sourcetable!sumofTotal + sourcetable!tax - (sourcetable!discount)
    temptable!val2 = sourcetable!TITEM
    temptable!val3 = sourcetable!TITEM - (sourcetable!sumofTotal + sourcetable!tax - (sourcetable!discount))
    temptable!STR21 = "≈Ã„«·Ì ›Ê« Ì— „‘ —Ì«  „Ê—œÌ‰ Œ·«· › —…"
    temptable!str22 = TurnValue(retHeader(aHeader, 0, 2))
    temptable!str12 = sourcetable!STORENAME

    temptable.Update
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & IIf(Check1.Value = 0, "\Reports\sup4.rpt", "\Reports\sup4G.rpt")
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
suplookup
End If
End Sub
Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""
If xCode.text = "" Then Exit Sub
xCode.text = RetZero(xCode.text, 3)
xCodeDesca.Caption = GetDesca("select desca from file4_10 where code = " & MyParn(xCode.text), con)
End Sub
Sub myProc()
    ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End Sub
Private Sub suplookup()
    Dim Generalarray(5)
    Dim listarray(0, 4)
    Dim GrdArray(1, 1)
    
    Set Generalarray(0) = Me
    Generalarray(1) = "Select Code, DescA From FILE4_10"
    Generalarray(2) = "Order by file4_10.Desca"
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
    listarray(0, 1) = "(%%DESCA%%) "
    
    GrdArray(0, 0) = "ﬂÊœ «·„Ê—œ"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "≈”„ «·„Ê—œ"
    GrdArray(1, 1) = 3000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "«” ⁄·«„"
    Search3.Show 1
End Sub
