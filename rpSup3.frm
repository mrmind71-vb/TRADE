VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form rpSup3 
   Caption         =   " ﬁ«—Ì— «·„Ê—œÌ‰"
   ClientHeight    =   3210
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
   ScaleHeight     =   3210
   ScaleWidth      =   6300
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdApply 
      Caption         =   "⁄—÷"
      Height          =   420
      Left            =   1395
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2700
      Width           =   1275
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "Œ—ÊÃ"
      Height          =   420
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2700
      Width           =   1275
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
         Left            =   3510
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
         FormatString    =   $"rpSup3.frx":0000
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
End
Attribute VB_Name = "rpSup3"
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
End Sub
Private Sub CmdApply_Click()
If Not MYVALID Then Exit Sub
doprint1
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate1.Text) Then
    MsgBox "«· «—ÌŒ «·«Ê· ÷—Ê—Ì"
    Exit Function
End If
If Not IsDate(xdate2.Text) And Trim(xdate2.Text) <> "" Then
    MsgBox "«· «—ÌŒ «·À«‰Ì €Ì— ”·Ì„"
    Exit Function
End If
MYVALID = True
End Function
Private Sub doprint1()
Dim aHeader(1), cString, cString2, cCode
If Not MYVALID Then Exit Sub
Dim TempTable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset

Dim FBalanceTable As New ADODB.Recordset
cString = "Select FILE4_10.code,Sum(Total) as FirstBalance from SupList Inner Join FILE4_10 on SupList.code = FILE4_10.Code where date < " & DateSq(xDate1.Text)
cString = cString & " group by FILE4_10.Code"
FBalanceTable.Open cString, con, adOpenStatic, adLockReadOnly

contemp.Execute "DELETE * FROM TEMP"
TempTable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT SupList.*,FILE4_10.Desca as [clientDesca]" & _
         " FROM SupList INNER JOIN FILE4_10 ON SupList.CODE = FILE4_10.CODE"

If Trim(xcode.Text) <> "" Then
    cString = cString & turnFound(cString) & " FILE4_10.code = " & MyParn(xcode.Text)
End If

If GrdQry(grid1, "file4_10.[GROUP]", True) <> "" Then
    cString = cString & turnFound(cString) & GrdQry(grid1, "File4_10.[GROUP]", True)
    aHeader(1) = "[" & "„Ã„Ê⁄… „Ê—œÌ‰ : " & GrdTitle(grid1) & "]"
End If
          
If IsDate(xDate1.Text) Then
    cString = cString & turnFound(cString) & " date >= " & DateSq(xDate1.Text)
    aHeader(1) = "[" & BetweenString(xDate1.Text, xdate2.Text) & "]"
End If

If IsDate(xdate2.Text) Then
    cString = cString & turnFound(cString) & " date <= " & DateSq(xdate2.Text)
    aHeader(1) = "[" & BetweenString(xDate1.Text, xdate2.Text) & "]"
End If

cString = cString & " " & _
                    " Order By SupList.Code,SupList.Date,iif(Val(supList.type) > 1 and Val(supList.Type) < 10,'0' & Val(supList.type),supList.Type) ,SupList.Doc_No"
                                             
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With sourcetable
Do Until sourcetable.EOF
    If cCode <> sourcetable!CODE Then
        nBalance = 0
        cCode = sourcetable!CODE
        FBalanceTable.Find "code = " & MyParn(sourcetable!CODE), , adSearchForward, adBookmarkFirst
        If Not FBalanceTable.EOF Then
            If Val(FBalanceTable!FirstBalance & "") <> 0 Then
                nCount = nCount + 1
                TempTable.AddNew
                TempTable!str1 = "—’Ìœ ”«»ﬁ"
                TempTable!val5 = Val(FBalanceTable!FirstBalance & "")
                TempTable!Val10 = nCount
                TempTable!str3 = sourcetable!CODE
                nBalance = TempTable!val5
                TempTable!str7 = sourcetable!ClientDesca
                TempTable.Update
            End If
        End If
    End If
    If sourcetable!Type = 1 And sourcetable!total = 0 Then GoTo LastLoop
    TempTable.AddNew
    TempTable!str3 = sourcetable!CODE
    TempTable!Date1 = sourcetable!Date
    TempTable!str1 = !TypeDesca & " " & sourcetable!doc_no
    TempTable!str4 = sourcetable!DESCA
    TempTable!str8 = sourcetable!ItemDescA
    TempTable!str2 = sourcetable!ItemDescA
    
    If Not (!Type = 2 Or !Type = 3 Or !Type = 20 Or !Type = 21) Then
        TempTable!str5 = RetZero(sourcetable!Type, 2) & RetZero(sourcetable!doc_no & "", 10) & nCount
    Else
        TempTable!str5 = RetZero(sourcetable!Type, 2) & RetZero(sourcetable!doc_no & "", 10)
    End If
    
    TempTable!VAL1 = sourcetable!Quant
    TempTable!val2 = sourcetable!price
    If TempTable!str5 <> cString2 Then
        cString2 = TempTable!str5
        If (!Type = 2 Or !Type = 3 Or !Type = 20 Or !Type = 21) Then
            If !Type = 2 Or !Type = 21 Then
                TempTable!val3 = Val(GetDesca("Select Sum(Total) - VAL(DISCOUNT & '') From supList where Type = " & !Type & " and Doc_no = " & MyParn(!doc_no) & " Group by Discount") & "")
            Else
                TempTable!val4 = -1 * Val(GetDesca("Select Sum(Total) - VAL(DISCOUNT & '')  From supList where Type = " & !Type & " and Doc_no = " & MyParn(!doc_no) & " Group by Discount") & "")
            End If
        ElseIf !Type <> 1 Then
            TempTable!val4 = -1 * Val(sourcetable!total)
        End If
        TempTable!VAL22 = 0
    Else
        TempTable!VAL1 = sourcetable!Quant
        TempTable!val2 = sourcetable!price
        If !Type <> "2" And !Type <> "20" Then
            TempTable!Val6 = sourcetable!total
        End If
        TempTable!VAL22 = 1
    End If
    If !Type = 1 Then
        nBalance = nBalance + Val(sourcetable!total)
    Else
        nBalance = nBalance + Val(TempTable!val3 & "") - Val(TempTable!val4 & "")
    End If
    TempTable!val5 = nBalance
    TempTable!Val10 = nCount
    TempTable!str7 = sourcetable!ClientDesca
    nCount = nCount + 1
    TempTable.Update
LastLoop:
    sourcetable.MoveNext
Loop
End With
If TempTable.EOF And TempTable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
Else
    contemp.Execute "update temp set val21 = 1 where str5 in (select str5 from temp group by str5 having sum(1) > 1)"
    contemp.BeginTrans
    contemp.CommitTrans
    main.Report1.ReportFileName = App.Path & "\Reports\SUP3.rpt"
    main.Report1.DataFiles(0) = tempFile
    main.Report1.Action = 1
End If
sourcetable.Close
TempTable.Close
Set sourcetable = Nothing
Set TempTable = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    suplookup
End If
End Sub

Private Sub xcode_LostFocus()
xCodeDesca.Caption = ""
If xcode.Text = "" Then Exit Sub
xCodeDesca.Caption = GetDesca("select desca from FILE4_10 where code = " & MyParn(xcode.Text)) & ""
End Sub
Sub myproc()
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

