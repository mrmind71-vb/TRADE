VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpSup9 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   3630
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
   ScaleHeight     =   3630
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   45
      Width           =   6180
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1890
         Width           =   1680
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   2250
         Width           =   1680
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   225
         Width           =   1230
      End
      Begin VSFlex7LCtl.VSFlexGrid grid1 
         Height          =   1215
         Left            =   990
         TabIndex        =   6
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
         FormatString    =   $"rpSup9.frx":0000
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
      Begin MSDataListLib.DataCombo xmosm 
         Height          =   315
         Left            =   1305
         TabIndex        =   7
         Top             =   2625
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   630
         Width           =   840
      End
      Begin VB.Label Label5 
         Caption         =   "„Ê”„ "
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
         TabIndex        =   8
         Top             =   2715
         Width           =   1005
      End
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1275
   End
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   420
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3150
      Width           =   1275
   End
   Begin Crystal.CrystalReport REPORT1 
      Left            =   45
      Top             =   45
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
      Left            =   45
      Top             =   45
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "data1"
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
Attribute VB_Name = "rpSup9"
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

data4.ConnectionString = strCon
data4.RecordSource = "Select MOSM,DescA From MOSM     ORDER by DATE "
Set xMosm.RowSource = data4
xMosm.ListField = "Desca"
xMosm.BoundColumn = "MOSM"
xMosm.BoundText = cPMosm

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
Private Sub OLDdoprint1()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

If Trim(xCode.text) <> "" Then
    cWhere = cWhere & turnFound(cWhere) & " file7_10h.code = " & MyParn(xCode.text)
    aHeader(2) = "[" & "··„Ê—œ : " & xCodeDesca.Caption & "]"
End If
If GrdQry(grid1, "file4_10.[GROUP]", True) <> "" Then
    cWhere = cWhere & turnFound(cWhere, " and ") & GrdQry(grid1, "File4_10.[GROUP]", True)
    aHeader(1) = "[" & "„Ã„Ê⁄… „Ê—œÌ‰ : " & GrdTitle(grid1) & "]"
End If
          
If IsDate(xdate1.text) Then
    cWhere = cWhere & turnFound(cWhere, " and ") & " FILE7_10H.date >= " & DateSq(xdate1.text)
    aHeader(2) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

If IsDate(xDate2.text) Then
    cWhere = cWhere & turnFound(cWhere, " and ") & " FILE7_10H.date <= " & DateSq(xDate2.text)
    aHeader(2) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

cField1 = "(Select Sum(Total) " & _
          " from (file7_10 inner join file7_10h on file7_10.doc_no = file7_10h.doc_no) inner join file4_10 on file7_10h.code = file4_10.code " & _
          turnFound(cWhere) & _
          cWhere & ") as sumoftotal"
          
cField2 = "(Select Sum(Discount) " & _
          " from file7_10h  inner JOIN file4_10 on file7_10h.code = file4_10.code " & _
          turnFound(cWhere) & _
          cWhere & ") as Sumofdiscount"
          
cField3 = "(Select Sum(tax) " & _
          " from file7_10h  inner JOIN file4_10 on file7_10h.code = file4_10.code " & _
          turnFound(cWhere) & _
          cWhere & ") as SumOftax"
                               
cString = "SELECT FILE7_10H.DOC_NO, FILE7_10H.DATE, FILE7_10H.CODE, FILE0_40.DESCA AS FILE0_40DESCA, FILE7_10.ITEM, FILE1_10.DESCA, FILE7_10.PRICE, FILE7_10.QUANT, FILE7_10.CHARGES, FILE7_10.CHARGE2, FILE7_10H.tax, FILE7_10H.discount, file4_10.DESCA AS FILE4_10DESCA ,FILE7_10.CHARGESDESCA, " & _
          cField1 & "," & cField2 & "," & cField3 & _
          " FROM (((FILE7_10 INNER JOIN FILE7_10H ON FILE7_10.DOC_NO = FILE7_10H.DOC_NO) LEFT JOIN FILE0_40 ON FILE7_10H.store = FILE0_40.CODE) LEFT JOIN FILE1_10 ON FILE7_10.ITEM = FILE1_10.ITEM) inner JOIN file4_10 ON FILE7_10H.code = file4_10.CODE"
cString = cString & turnFound(cWhere) & cWhere
                                                      
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
        temptable.AddNew
        temptable!str1 = sourcetable!doc_no
        temptable!str2 = sourcetable!FILE4_10DESCA
        temptable!str3 = sourcetable!FILE0_40DESCA
        temptable!Date1 = sourcetable!Date
        temptable!str4 = sourcetable!Item
        temptable!str5 = sourcetable!DESCA
        temptable!str6 = sourcetable!Chargesdesca
        temptable!VAL1 = Val(sourcetable!Quant & "")
        temptable!val2 = Val(sourcetable!price & "")
        temptable!val3 = Val(sourcetable!Quant & "") * Val(sourcetable!price & "")
        temptable!VAL4 = (Val(sourcetable!Quant & "") * Val(sourcetable!price & "")) + Val(sourcetable!CHARGES & "") + Val(sourcetable!charge2 & "")
        temptable!val5 = Val(sourcetable!discount & "")
        temptable!VAL7 = Val(sourcetable!tax & "")
        temptable!Val8 = Val(sourcetable!CHARGES & "")
        temptable!val9 = Val(sourcetable!charge2 & "")
        temptable!Val10 = Val(sourcetable!sumofTotal & "") + Val(sourcetable!sumOftax & "") - Val(sourcetable!sumOfDiscount & "")
        temptable!STR21 = TurnValue(retHeader(aHeader, 0, 3))
        temptable.Update
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = IIf(chkCharge.Value = 1, App.Path & "\Reports\Sup8c.rpt", App.Path & "\Reports\Sup8.rpt")
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
xCodeDesca.Caption = GetDesca("select desca from FILE4_10 where code = " & MyParn(xCode.text), con) & ""
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
Private Sub doprint1()
Dim aHeader(3)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

If Trim(xCode.text) <> "" Then
    cWhere = cWhere & " AND FILE7_10h.code = " & MyParn(xCode.text)
    aHeader(2) = "[" & "··„Ê—œ : " & xCodeDesca.Caption & "]"
End If


If GrdQry(grid1, "FILE4_10.[GROUP]", True) <> "" Then
    cWhere = cWhere & " AND " & GrdQry(grid1, "FILE4_10.[GROUP]", True)
    aHeader(1) = "[" & "„Ã„Ê⁄… „Ê—œÌ‰ : " & GrdTitle(grid1) & "]"
End If
          
If IsDate(xdate1.text) Then
    cWhere = cWhere & " AND FILE7_10H.date >= " & DateSq(xdate1.text)
    aHeader(2) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

If IsDate(xDate2.text) Then
    cWhere = cWhere & " AND FILE7_10H.date <= " & DateSq(xDate2.text)
    aHeader(2) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

If Trim(xMosm.BoundText) <> "" Then
    cWhere = cWhere & " AND FILE7_10h.mosm  = " & MyParn(xMosm.BoundText)
    aHeader(3) = "[" & "„Ê”„ : " & xMosm.text & "]"
End If

cString = "SELECT FILE7_10H.DOC_NO, FILE7_10H.DATE, FILE7_10H.CODE, FILE1_10.MODELFACT, FILE1_10.desca AS ITEMDESCA, SUM(FILE7_10.QUANT) AS T_QUANT , SUM(FILE7_10.TOTAL) AS T_TOT , FILE7_10H.TAX, FILE7_10H.DISCOUNT, FILE4_10.DESCA AS SUPPDESCA, FACT.DESCA AS FACTDESCA " & _
          " FROM FILE7_10 INNER JOIN FILE7_10H ON FILE7_10.DOC_NO = FILE7_10H.DOC_NO INNER JOIN FILE1_10 ON FILE7_10.ITEM = FILE1_10.ITEM INNER JOIN FILE4_10 ON FILE7_10H.CODE = FILE4_10.CODE INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE  WHERE FILE7_10H.DOC_NO IS NOT NULL "
cString = cString & cWhere & " GROUP BY FILE7_10H.DOC_NO, FILE7_10H.DATE, FILE7_10H.CODE, FILE1_10.MODELFACT, FILE1_10.desca, FILE7_10H.TAX, FILE7_10H.DISCOUNT, FILE4_10.DESCA, FACT.DESCA "
                                                     
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = sourcetable!doc_no
    
    temptable!str2 = sourcetable!SUPPDESCA
    temptable!Date1 = sourcetable!Date
    
    temptable!str6 = TurnValue(DelZero(sourcetable!modelfact), "", Null)
    temptable!str4 = TurnValue(sourcetable!FACTDESCA, "", Null)
    temptable!str5 = TurnValue(sourcetable!ITEMDESCA, "", Null)
    temptable!VAL1 = Val(sourcetable!T_QUANT & "")
    temptable!val2 = Val(sourcetable!T_TOT & "") / Val(sourcetable!T_QUANT & "")
    temptable!val3 = Val(sourcetable!T_TOT & "")
    temptable!val5 = Val(sourcetable!discount & "")
    temptable!VAL7 = Val(sourcetable!tax & "")
    
        
    temptable!str22 = " ›’Ì·Ï „‘ —Ì«  „‰ «·„Ê—œÌ‰ "
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
main.Report1.ReportFileName = App.Path & "\Reports\sup8.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub




