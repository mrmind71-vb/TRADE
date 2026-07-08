VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpClient2 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   2820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
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
   ScaleHeight     =   2820
   ScaleWidth      =   5100
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   375
      Left            =   1305
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2385
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2385
      Width           =   1140
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   5250
      Top             =   1725
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
   Begin VB.Frame Frame1 
      Height          =   2280
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   45
      Width           =   4980
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1845
         Width           =   1650
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1485
         Width           =   1650
      End
      Begin VSFlex7LCtl.VSFlexGrid grid1 
         Height          =   1215
         Left            =   45
         TabIndex        =   0
         Top             =   180
         Width           =   3720
         _cx             =   6562
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
         FormatString    =   $"rpClient2.frx":0000
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Ã„Ê⁄…"
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
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   915
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
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1920
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
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1530
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
End
Attribute VB_Name = "rpClient2"
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
grid1.ColAlignment(0) = flexAlignRightCenter
grdMake "Select Code,DescA From FILE3_50", "code", "desca", con, grid1
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xdate1.text) Or Not IsDate(xDate2.text) Then
    MsgBox " ÕœÌœ «· «—ÌŒ "
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()

Dim aHeader(1)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

Dim FBalanceTable As New ADODB.Recordset
'FBalanceTable.Open "Select code,Sum(Format(SAL - PAY,'Fixed')) as firstBalance from File3_11 where date < " & DateSq(xDate1.Text) & " group by Code", CON, adOpenStatic, adLockReadOnly

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

If IsDate(xdate1.text) Then
    cWhere = "DATE < " & DateSq(xdate1.text)
    cField1 = myiif(cWhere, _
             " SAL - PAY ") & " as firstBalance"
Else
    cField1 = "0 as firstBalance"
End If


cWhere = ""
If IsDate(xdate1.text) Then
    cWhere = cWhere & " date >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " date <= " & DateSq(xDate2.text)
End If

cField2 = myiif(cWhere & turnFound(cWhere, " and ") & _
        "[TYPE] = '4'", _
        " Sal  ") & _
        " as Sales"

cField3 = myiif(cWhere & turnFound(cWhere, " and ") & _
        "[TYPE] = '5'", _
        " PAY  ") & _
        " as RET"

cField4 = myiif(cWhere & turnFound(cWhere, " and ") & _
        "( [TYPE] <> '4' AND  [TYPE] <> '5' AND [TYPE] <> 'A' ) ", _
        " Pay  - Sal  ") & _
        " as Paid"

cField5 = myiif(cWhere & turnFound(cWhere, " and ") & _
        "( [TYPE] = 'A'  ) ", _
        " Pay ") & _
        " as TCHQ"


cWhere = "DATE <= " & DateSq(xDate2.text)
cField7 = myiif(cWhere, _
         " SAL  - PAY ") & " as LASTBalance"

cString = "SELECT File3_11.CODE, File3_10.DESCA, File3_10.[group], FILE3_50.DESCA AS GroupDesca ," & _
          cField1 & "," & _
          cField2 & "," & _
          cField3 & "," & _
          cField4 & "," & _
          cField5 & "," & _
          cField7 & _
         " FROM (File3_11 INNER JOIN File3_10 ON File3_11.CODE = File3_10.CODE) LEFT JOIN FILE3_50 ON File3_10.[group] = FILE3_50.CODE"

If GrdQry(grid1, "file3_10.[group]", True) <> "" Then
    cString = cString & turnFound(cString) & GrdQry(grid1, "File3_10.[group]", True)
    aHeader(0) = "[" & "„Ã„Ê⁄… ⁄„·«¡ : " & GrdTitle(grid1) & "]"
End If
          
If IsDate(xdate1.text) Then
    aHeader(1) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

If IsDate(xDate2.text) Then
    aHeader(1) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

cString = cString & " GROUP BY File3_11.CODE, File3_10.DESCA, File3_10.[group],FILE3_50.DESCA"
                         
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText


Do Until sourcetable.EOF
   If Val(sourcetable!sales & "") <> 0 Or Val(sourcetable!RET & "") <> 0 Or Val(sourcetable!PAID & "") <> 0 Or Val(sourcetable!TCHQ & "") <> 0 Then
        temptable.AddNew
        temptable!str1 = sourcetable!CODE
        temptable!str2 = sourcetable![DESCA]
        temptable!str3 = sourcetable![Group]
        temptable!str4 = sourcetable!GroupDesca
        temptable!VAL1 = sourcetable!FirstBalance
        temptable!val2 = sourcetable!sales
        temptable!val3 = sourcetable!RET
        temptable!VAL4 = sourcetable!PAID
        temptable!val5 = GetDesca("SELECT COUNT(CODE1) AS TCOUNT FROM FILE5_20 WHERE CODE1 = " & MyParn(sourcetable!CODE) & " AND DATE_R >= " & DateSq(xdate1.text) & " AND DATE_R <= " & DateSq(xDate2.text), con)
        
        temptable!Val6 = sourcetable!TCHQ
        temptable!VAL7 = sourcetable!LastBalance
        temptable!STR21 = TurnValue(retHeader(aHeader, 0, 2))
        temptable.Update
    End If
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\CLIENT2.rpt"
main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
main.Report1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
