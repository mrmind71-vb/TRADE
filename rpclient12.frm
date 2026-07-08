VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpclient12 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   3885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6300
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
   ScaleHeight     =   3885
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1350
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   3225
      Width           =   1275
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3225
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   3195
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   6180
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1230
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
      Begin VSFlex7LCtl.VSFlexGrid grid1 
         Height          =   1215
         Left            =   990
         TabIndex        =   1
         Top             =   630
         Width           =   3765
         _cx             =   6641
         _cy             =   2143
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
         FormatString    =   $"rpclient12.frx":0000
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
      Begin MSDataListLib.DataCombo xstore 
         Height          =   315
         Left            =   1275
         TabIndex        =   12
         Top             =   2700
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "„Œ“‰ :"
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
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2790
         Width           =   1005
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
         TabIndex        =   11
         Top             =   630
         Width           =   840
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
         TabIndex        =   10
         Top             =   225
         Width           =   3345
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì· :"
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
         Left            =   4905
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   315
         Width           =   600
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
         TabIndex        =   8
         Top             =   2340
         Width           =   465
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
         TabIndex        =   7
         Top             =   1980
         Width           =   765
      End
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
   Begin MSAdodcLib.Adodc data4 
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
Attribute VB_Name = "rpclient12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub chkpaid_Click(Index As Integer)
'If chkPaid(Index).Value = 0 Then Exit Sub
For i = 0 To chkPaid.Count - 1
    If i <> Index Then chkPaid(i).Value = 0
Next
End Sub

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
grdMake "Select Code,DescA From FILE3_50", "code", "desca", con, grid1
data4.ConnectionString = strCon
data4.RecordSource = "Select Code,DescA From File0_40"
Set XSTORE.RowSource = data4
XSTORE.ListField = "Desca"
XSTORE.BoundColumn = "Code"

End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
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

If Trim(xCode.text) <> "" Then
    cWhere = cWhere & turnFound(cWhere) & " FILE6_10h.code = " & MyParn(xCode.text)
    aHeader(2) = "[" & "··⁄„Ì· : " & xCodeDesca.Caption & "]"
End If

If GrdQry(grid1, "FILE3_10.[GROUP]", True) <> "" Then
    cWhere = cWhere & turnFound(cWhere, " and ") & GrdQry(grid1, "FILE3_10.[GROUP]", True)
    aHeader(1) = "[" & "„Ã„Ê⁄… ⁄„·«¡ : " & GrdTitle(grid1) & "]"
End If
          
If IsDate(xdate1.text) Then
    cWhere = cWhere & turnFound(cWhere, " and ") & " FILE6_10H.date >= " & DateSq(xdate1.text)
    aHeader(2) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

If IsDate(xDate2.text) Then
    cWhere = cWhere & turnFound(cWhere, " and ") & " FILE6_10H.date <= " & DateSq(xDate2.text)
    aHeader(2) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

If Trim(XSTORE.BoundText) <> "" Then
    cWhere = cWhere & turnFound(cWhere, " and ") & " FILE6_10h.STORE = " & MyParn(XSTORE.BoundText)
    aHeader(3) = "[" & "·„Œ“‰ : " & XSTORE.text & "]"
End If

cField1 = "(Select Sum(FILE6_10.Total)  " & _
          " from (FILE6_10 inner join FILE6_10h on FILE6_10.doc_no = FILE6_10h.doc_no) left join FILE3_10 on FILE6_10h.code = FILE3_10.code " & _
          turnFound(cWhere) & _
          cWhere & ") as sumoftotal"
          
cField2 = "(Select Sum(Discount) " & _
          " from FILE6_10h  left join FILE3_10 on FILE6_10h.code = FILE3_10.code " & _
          turnFound(cWhere) & _
          cWhere & ") as Sumofdiscount"
         
cField3 = "(Select Sum(tax) " & _
          " from FILE6_10h  left join FILE3_10 on FILE6_10h.code = FILE3_10.code " & _
          turnFound(cWhere) & _
          cWhere & ") as SumOftax"
                               
cString = "SELECT FILE6_10H.DOC_NO, FILE6_10H.DATE, FILE6_10H.CODE, FILE0_40.DESCA AS FILE0_40DESCA, FILE6_10.ITEM, FILE1_10.DESCA, FILE6_10.PRICE, FILE6_10.QUANT, FILE6_10H.tax,FILE6_10.DISCOUNT AS FILE6_10DISCOUNT, FILE6_10H.discount, FILE3_10.DESCA AS FILE3_10DESCA,FILE6_10.TOTAL, " & _
          cField1 & "," & cField2 & "," & cField3 & _
          " FROM (((FILE6_10 INNER JOIN FILE6_10H ON FILE6_10.DOC_NO = FILE6_10H.DOC_NO) LEFT JOIN FILE0_40 ON FILE6_10H.store = FILE0_40.CODE) LEFT JOIN FILE1_10 ON FILE6_10.ITEM = FILE1_10.ITEM) LEFT JOIN FILE3_10 ON FILE6_10H.code = FILE3_10.CODE"
cString = cString & turn(cWhere, " where ") & cWhere

                                                      
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = sourcetable!doc_no
    temptable!str2 = sourcetable![FILE3_10DESCA]
    temptable!str3 = sourcetable![FILE0_40DESCA]
    temptable!Date1 = sourcetable!Date
    temptable!str4 = sourcetable!Item
    temptable!str5 = sourcetable![DESCA]
    temptable!VAL1 = Val(sourcetable!Quant & "")
    temptable!val2 = Val(sourcetable!price & "")
    temptable!val3 = Val(sourcetable!TOTAL & "")
    temptable!VAL22 = Val(sourcetable![FILE6_10discount] & "")
    temptable!val5 = Val(sourcetable![discount] & "")
    temptable!VAL7 = Val(sourcetable!tax & "")
        
    temptable!VAL13 = Val(sourcetable!sumofTotal & "")
    temptable!VAL14 = Val(sourcetable!sumOfDiscount & "")
    temptable!val15 = Val(sourcetable!sumOftax & "")
    temptable!val16 = Val(sourcetable!sumofTotal & "") + Val(sourcetable!sumOftax & "") - Val(sourcetable!sumOfDiscount & "")
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 4))
    temptable.Update
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\client12.rpt"
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
xCodeDesca.Caption = GetDesca("select desca from FILE3_10 where code = " & MyParn(xCode.text), con) & ""
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
    Generalarray(1) = "Select Code, DescA From FILE3_10"
    Generalarray(2) = "Order by FILE3_10.Desca"
    Generalarray(3) = 4200
    Generalarray(5) = False
    
    listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
    listarray(0, 1) = "(%%DESCA%%) "
    
    GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
    GrdArray(0, 1) = 1000
    
    GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
    GrdArray(1, 1) = 3000
    
    searchArray = Array(Generalarray, listarray, GrdArray)
    Load Search3
    Search3.Caption = "«” ⁄·«„"
    Search3.Show 1
End Sub

Private Sub xPost_Click()
End Sub

