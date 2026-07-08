VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpSup6 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   2805
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
   ScaleHeight     =   2805
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkCharge 
      Alignment       =   1  'Right Justify
      Caption         =   "≈ŸÂ«— «·„’«—Ì›"
      Height          =   240
      Left            =   2565
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2385
      Value           =   1  'Checked
      Width           =   3210
   End
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   5775
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   3060
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1530
         Width           =   1680
      End
      Begin VB.TextBox xDate2 
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
         Height          =   1230
         Left            =   585
         TabIndex        =   0
         Top             =   225
         Width           =   4155
         _cx             =   7329
         _cy             =   2170
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
         FormatString    =   $"rpSup6.frx":0000
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
         TabIndex        =   8
         Top             =   1575
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
         TabIndex        =   7
         Top             =   1935
         Width           =   465
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
         Left            =   4860
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   840
      End
   End
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
      Height          =   375
      Left            =   1215
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2340
      Width           =   1140
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
      Height          =   375
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2340
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
End
Attribute VB_Name = "rpSup6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Private Sub CmdExit_Click()
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
If Not IsDate(XDATE2.Text) And Trim(XDATE2.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(1)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT FILE7_20H.DOC_NO,FILE7_20H.INV_NO, FILE7_20H.DATE, SUM(FILE7_20.TOTAL)AS SUMOFTOTAL, FILE7_20H.DISCOUNT2,FILE7_20H.DISCOUNT, FILE7_20H.TAX,FILE7_20H.CODE,FILE4_10.DESCA,FILE4_10.[GROUP],FILE4_50.DESCA AS GROUPDESCA" & _
         " FROM ((FILE7_20H INNER JOIN FILE7_20 ON FILE7_20H.doc_no =   FILE7_20.doc_no) inner join file4_10 on file7_20H.CODE = FILE4_10.CODE) LEFT JOIN FILE4_50 ON FILE4_10.[GROUP] = FILE4_50.CODE "

If GrdQry(grid1, "file4_10.[GROUP]", True) <> "" Then
    cWhere = cWhere & turnFound(cWhere, " and ") & GrdQry(grid1, "File4_10.[GROUP]", True)
    aHeader(0) = "[" & "„Ã„Ê⁄… „Ê—œÌ‰ : " & GrdTitle(grid1) & "]"
End If

If IsDate(XDATE2.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " FILE7_20H.date >= " & DateSq(xDate1.Text)
    aHeader(1) = "[" & BetweenString(xDate1.Text, XDATE2.Text) & "]"
End If

If IsDate(XDATE2.Text) Then
    cWhere = cWhere & turn(cWhere, " and ") & " FILE7_20H.date <= " & DateSq(XDATE2.Text)
    aHeader(1) = "[" & BetweenString(xDate1.Text, XDATE2.Text) & "]"
End If

cString = cString & turnFound(cWhere) & cWhere
cString = cString & " GROUP BY FILE7_20H.DOC_NO,FILE7_20H.INV_NO, FILE7_20H.DATE, FILE7_20H.DISCOUNT,FILE7_20H.DISCOUNT2 ,FILE7_20H.TAX,FILE7_20H.CODE,FILE4_10.DESCA,FILE4_10.[GROUP],FILE4_50.DESCA"
                         
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
'    If Abs(sourcetable!Balance) >= 0 Then
        temptable.AddNew
        temptable!str1 = sourcetable!doc_no
        temptable!str2 = sourcetable!INV_NO
        temptable!str12 = sourcetable!CODE
        temptable!str11 = sourcetable!desca
        temptable!Date1 = sourcetable!Date
        temptable!val1 = sourcetable!sumofTotal - sourcetable!Discount2
        temptable!val2 = sourcetable!tax
        temptable!val3 = sourcetable!sumofTotal + sourcetable!tax - sourcetable!Discount2
        temptable!VAL4 = sourcetable!discount
        temptable!val5 = sourcetable!sumofTotal + sourcetable!tax - (sourcetable!discount + sourcetable!Discount2)
        temptable!str21 = "≈Ã„«·Ì ›Ê« Ì— „‘ —Ì«  „Ê—œÌ‰ Œ·«· › —…"
        temptable!str22 = TurnValue(retHeader(aHeader, 0, 2))
        temptable.Update
'    End If
    sourcetable.MoveNext
Loop
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.REPORT1.ReportFileName = App.Path & "\Reports\sup6" & IIf(chkCharge.Value = 1, "C", "") & ".rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
sourcetable.Close
temptable.Close
Set sourcetable = Nothing
Set temptable = Nothing
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
suplookup
End If
End Sub
Private Sub xCode_LostFocus()
xCodeDesca.Caption = ""
If xCode.Text = "" Then Exit Sub
xCode.Text = RetZero(xCode.Text, 6)
xCodeDesca.Caption = GetDesca("select desca from file4_10 where code = " & MyParn(xCode.Text))
End Sub
Sub myProc()
    ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
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

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub
