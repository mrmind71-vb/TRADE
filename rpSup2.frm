VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpSup2 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   3210
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
   ScaleHeight     =   3210
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2670
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   5775
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
         FormatString    =   $"rpSup2.frx":0000
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
      Top             =   2745
      Width           =   1140
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   375
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2745
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
Attribute VB_Name = "rpSup2"
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
If Not IsDate(xDate2.Text) And Trim(xDate2.Text) <> "" Then
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

If IsDate(xdate1.Text) Then
    cwhere = "Date < " & DateSq(xdate1.Text)
     cField1 = myiif(cwhere, _
        " SAL - PAY ") & _
        " as FirstBalance"
Else
    cField1 = "0 as FirstBalance"
End If

    
cwhere = ""
If IsDate(xdate1.Text) Then cwhere = "Date >= " & DateSq(xdate1.Text)
    
cField2 = myiif(cwhere & turn(cwhere, " and ") & _
        "[TYPE] = '4'", _
        " SAL - PAY ") & _
        " as Purchase"
        

cField3 = myiif(cwhere & turn(cwhere, " and ") & _
        "[TYPE] = '5'", _
        " PAY - SAL ") & _
        " as Purchase_ret"

cField4 = myiif(cwhere & turn(cwhere, " and ") & _
        "[TYPE] = '16' or [TYPE] = '17' ", _
        " PAY - SAL ") & _
        " as Tax"

cField5 = myiif(cwhere & turn(cwhere, " and ") & _
        " [TYPE] = '7' or [TYPE] = '8' OR [TYPE] = '10' or [TYPE] = '11' ", _
        " PAY - SAL ") & _
        " as Cash"

cField6 = myiif(cwhere & turn(cwhere, " and ") & _
        "[TYPE] = 'A' or [TYPE] = 'C' ", _
        " PAY - Sal") & _
        " as Chq"

cField7 = myiif("", _
        " SAL - PAY ") & _
        " as Balance"

cField8 = myiif("", _
        " PAY ") & _
        " as SumOfPay"

cField9 = myiif("", _
        " Sal ") & _
        " as SumOfSal"


cString = "SELECT FILE4_11.CODE, FILE4_10.DESCA, FILE4_10.[GROUP], FILE4_50.DESCA AS GroupDesca ," & _
          cField1 & "," & _
          cField2 & "," & _
          cField3 & "," & _
          cField4 & "," & _
          cField5 & "," & _
          cField6 & "," & _
          cField7 & "," & _
          cField8 & "," & _
          cField9 & _
         " FROM (FILE4_11 INNER JOIN FILE4_10 ON FILE4_11.CODE = FILE4_10.CODE) LEFT JOIN FILE4_50 ON FILE4_10.[GROUP] = FILE4_50.CODE"


cwhere = ""
If GrdQry(grid1, "file4_10.[GROUP]", True) <> "" Then
    cwhere = cwhere & turn(cwhere, " and ") & GrdQry(grid1, "File4_10.[GROUP]", True)
    aHeader(1) = "[" & "„Ã„Ê⁄… „Ê—œÌ‰ : " & GrdTitle(grid1) & "]"
End If

If IsDate(xDate2.Text) Then
    cwhere = cwhere & turn(cwhere, " and ") & " date <= " & DateSq(xDate2.Text)
    aHeader(1) = "[" & BetweenString(xdate1.Text, xDate2.Text) & "]"
End If

cString = cString & cwhere
cString = cString & " GROUP BY FILE4_11.CODE, FILE4_10.DESCA, FILE4_10.[GROUP],FILE4_50.DESCA"
                         
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Do Until sourcetable.EOF
    If Abs(sourcetable!balance) >= 0 Then
        temptable.AddNew
        temptable!str1 = sourcetable!CODE
        temptable!str2 = sourcetable![DESCA]
        temptable!str3 = sourcetable![Group]
        temptable!str4 = sourcetable!GroupDesca
        temptable!val1 = sourcetable!FirstBalance
        temptable!val2 = sourcetable!purchase
        temptable!val3 = sourcetable!Purchase_Ret
        temptable!VAL4 = sourcetable!tax
        temptable!val5 = sourcetable!CASH
        temptable!Val6 = sourcetable!Chq
        temptable!VAL7 = sourcetable!balance
        temptable!Val8 = sourcetable!SumofPay
        temptable!val9 = sourcetable!SumofSal
        temptable!str21 = TurnValue(retHeader(aHeader, 0, 2))
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
main.REPORT1.ReportFileName = App.Path & "\Reports\sup2.rpt"
main.REPORT1.DataFiles(0) = tempFile
main.REPORT1.Action = 1
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

