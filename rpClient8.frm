VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpClient8 
   Caption         =   " ﬁ«—Ì— «·⁄„·«¡"
   ClientHeight    =   3195
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
   ScaleHeight     =   3195
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   375
      Left            =   1260
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2700
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      Height          =   2670
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   6180
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1230
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   2205
         Width           =   1680
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1845
         Width           =   1680
      End
      Begin VSFlex7LCtl.VSFlexGrid grid1 
         Height          =   1215
         Left            =   990
         TabIndex        =   1
         Top             =   585
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
         FormatString    =   $"rpClient8.frx":0000
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
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   630
         Width           =   915
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
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
         Left            =   4785
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
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2295
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
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1935
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
Attribute VB_Name = "rpClient8"
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
If publicFlag = 8 Then doprint1 Else doprint2
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xdate1.text) And Trim(xdate1.text) <> "" Then
    MsgBox "«· «—ÌŒ «·√Ê· €Ì— ”·Ì„"
    Exit Function
End If
If Not IsDate(xDate2.text) And Trim(xDate2.text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString1 = "Select FILE6_20H.Doc_No,FILE6_20H.[Date],Sum(FILE6_20.TOTAL) + FILE6_20H.TAX  - FILE6_20H.DISCOUNT AS AMOUNT,'„œ›Ê⁄«  ‰ﬁœÌ…' as Desca,FILE3_10.DESCA AS SUPDESCA " & _
           " From (FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO) INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE " & _
           " WHERE (NOT ( FILE6_20H.BOX IS NULL))  "

If Trim(xCode.text) <> "" Then
    cString1 = cString1 & turnFound(cString1) & " FILE6_20H.CODE = " & MyParn(xCode.text)
    aHeader(2) = "[" & "··⁄„Ì· : " & xCodeDesca.Caption & "]"
End If

If GrdQry(grid1, "FILE3_10.[GROUP]", True) <> "" Then
    cString1 = cString1 & turnFound(cString1) & GrdQry(grid1, "FILE3_10.[GROUP]", True)
    aHeader(1) = "[ " & "„Ã„Ê⁄… ⁄„·«¡ : " & GrdTitle(grid1) & " ]"
End If
          
If IsDate(xdate1.text) Then
    cString1 = cString1 & turnFound(cString1) & " FILE6_20H.date >= " & DateSq(xdate1.text)
    aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
End If

If IsDate(xDate2.text) Then
    cString1 = cString1 & turnFound(cString1) & " FILE6_20H.date <= " & DateSq(xDate2.text)
    aHeader(2) = "[ " & BetweenString(xdate1.text, xDate2.text) & " ]"
End If
cString1 = cString1 & " Group by FILE6_20H.Doc_No,FILE6_20H.Code,FILE6_20H.Date,FILE3_10.DESCA,FILE6_20H.TAX,FILE6_20H.DISCOUNT"

cString2 = "SELECT FILE8_10H.Doc_No,FILE8_10H.[Date],[VALUE], FILE8_10.DESCA,FILE3_10.DESCA  From (FILE8_10H INNER JOIN FILE8_10 ON FILE8_10H.DOC_NO = FILE8_10.DOC_NO) INNER JOIN FILE3_10 ON FILE8_10.CODE = FILE3_10.CODE "
If Trim(xCode.text) <> "" Then
    cString2 = cString2 & turnFound(cString2) & " FILE8_10.CODE = " & MyParn(xCode.text)
End If

If GrdQry(grid1, "FILE3_10.[GROUP]", True) <> "" Then
    cString2 = cString2 & turnFound(cString2) & GrdQry(grid1, "FILE3_10.[GROUP]", True)
End If
          
If IsDate(xdate1.text) Then
    cString2 = cString2 & turnFound(cString2) & " FILE8_10H.date >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    cString2 = cString2 & turnFound(cString2) & " FILE8_10H.date <= " & DateSq(xDate2.text)
End If

cString3 = "SELECT FILE8_30H.Doc_No,FILE8_30H.[Date],-1 * [VALUE], FILE8_30.DESCA,FILE3_10.DESCA From (FILE8_30H INNER JOIN FILE8_30 ON FILE8_30H.DOC_NO = FILE8_30.DOC_NO) INNER JOIN FILE3_10 ON FILE8_30.CODE = FILE3_10.CODE "
If Trim(xCode.text) <> "" Then
    cString3 = cString3 & turnFound(cString3) & " FILE8_30.CODE = " & MyParn(xCode.text)
End If

If GrdQry(grid1, "FILE3_10.[GROUP]", True) <> "" Then
    cString3 = cString3 & turnFound(cString3) & GrdQry(grid1, "FILE3_10.[GROUP]", True)
End If
          
If IsDate(xdate1.text) Then
    cString3 = cString3 & turnFound(cString3) & " FILE8_30H.date >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    cString3 = cString3 & turnFound(cString3) & " FILE8_30H.date <= " & DateSq(xDate2.text)
End If

cString = cString1 & _
          " union all " & _
          cString2 & _
          " union all " & _
          cString3

With sourcetable
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str1 = sourcetable!SUPDESCA
    temptable!str2 = sourcetable!doc_no
    temptable!str3 = sourcetable!DESCA
    temptable!Date1 = sourcetable!Date
    temptable!VAL1 = Val(!AMOUNT & "")
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 3))
    temptable.Update
    sourcetable.MoveNext
Loop
End With
temptable.Requery
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\CLIENT8.rpt"
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
    CardLookup
End If
End Sub
Private Sub doprint2()
Dim aHeader(2), cString As String
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "Select File3_10.Desca,FILE8_30.CODE,FILE8_30.DOC_NO,FILE8_30.DATE,FILE8_30.DESCA,FILE8_30.VALUE from " & _
          " FILE8_30  INNER JOIN FILE3_10 ON FILE8_30.CODE = FILE3_10.CODE "

If Trim(xCode.text) <> "" Then
    cString = cString & turnFound(cString) & " File3_10.code = " & MyParn(xCode.text)
    aHeader(0) = "[" & "«·⁄„Ì· : " & xCodeDesca.Caption & "]"
End If

If GrdQry(grid1, "file3_10.[GROUP]", True) <> "" Then
    cString = cString & turnFound(cString) & GrdQry(grid1, "File3_10.[GROUP]", True)
    aHeader(0) = "[" & "„Ã„Ê⁄… ⁄„·«¡ : " & GrdTitle(grid1) & "]"
End If

If IsDate(xdate1.text) Then
    cString = cString & turnFound(cString) & " date >= " & DateSq(xdate1.text)
    aHeader(2) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If

If IsDate(xDate2.text) Then
    cString = cString & turnFound(cString) & " date <= " & DateSq(xDate2.text)
    aHeader(2) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
End If
         
If Check1.Value = 0 Then cString = cString & turnFound(cString) & " (not isnull(Box)) "
cString = cString & " Order by Date,File3_10.Code"
                                            
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!Date1 = sourcetable!Date
    temptable!str1 = sourcetable!doc_no
    temptable!str2 = sourcetable![FILE3_10.DESCA]
    temptable!str3 = sourcetable![FILE8_30.desca]
    temptable!VAL1 = sourcetable!Value
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 3))
    If Check1.Value <> 0 Then temptable!str22 = "«ŸÂ«— «· ”ÊÌ« "
    temptable.Update
    sourcetable.MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.Report1.ReportFileName = App.Path & "\Reports\CLIENT9.rpt"
    main.Report1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    main.Report1.Action = 1
End If
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub
Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""
If xCode.text = "" Then Exit Sub
xCodeDesca.Caption = GetDesca("select desca from file3_10 where code = " & MyParn(xCode.text), con) & ""
End Sub
Sub myProc()
xCode.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
xCodeDesca.Caption = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
Unload Search3
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From FILE3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub

