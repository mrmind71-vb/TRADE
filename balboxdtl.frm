VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form balBoxDtlfrm 
   Caption         =   " ›«’Ì· —’Ìœ «·Œ“«‰…"
   ClientHeight    =   7590
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   10905
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1005
      Left            =   2700
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   8070
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   1320
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5715
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   225
         Width           =   1320
      End
      Begin VB.Label lblDate2 
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
         Left            =   7170
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lblDate1 
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
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   330
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   5880
      Left            =   45
      TabIndex        =   0
      Top             =   1035
      Width           =   10785
      _cx             =   19024
      _cy             =   10372
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
      Rows            =   50
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
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame2 
      Height          =   690
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   6885
      Width           =   3435
      Begin VB.CommandButton cmdPrint 
         Caption         =   "ÿ»«⁄…"
         Height          =   420
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   1725
      End
      Begin VB.CommandButton Command1 
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   1500
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTop       =   0
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      BoundReportHeading=   "dddd"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "balBoxDtlfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdPrint_Click()
Dim temptable As New ADODB.Recordset, sourceTable As New ADODB.Recordset, loctable As New ADODB.Recordset
Dim nBalance As Double, aHeader(0)

contemp.Execute "Delete * From Temp"
temptable.Open "TEMP", contemp, adOpenKeyset, adLockOptimistic, adCmdTable


If IsDate(xDate1.Text) Then
    cString = "Select SUM(VALUE) as SumofValue,Sum(1) as SumOfOne   FROM boxmove WHERE  DateValue(Date) < " & DateSq(xDate1.Text) & " and box = " & MyParn(BalBox.XBOX.BoundText)
    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    If Not (loctable.EOF And loctable.BOF) Then
        nBalance = Val(loctable!Sumofvalue & "")
    End If
End If

cString = "Select BOXMOVE.* " & _
           " From boxmove where box = " & MyParn(BalBox.XBOX.BoundText)

If IsDate(xDate1.Text) Then
    cString = cString & turnFound2(cString) & " date >= " & DateSq(xDate1.Text)
    aHeader(0) = BetweenString(xDate1.Text, xdate2.Text)
End If

If IsDate(xdate2.Text) Then
    cString = cString & turnFound2(cString) & " date <= " & DateSq(xdate2.Text)
    aHeader(0) = BetweenString(xDate1.Text, xdate2.Text)
End If

cString = cString & " order by Date,Flag"
sourceTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not (loctable.EOF And loctable.EOF) Then
    If nBalance > 0 Then
        temptable.AddNew
        temptable!str10 = " ›’Ì·Ì Õ—ﬂ… " & BalBox.XBOX.Text
        temptable!str2 = "—’Ìœ ”«»ﬁ"
        temptable!val1 = nBalance
        temptable!val3 = nBalance
        temptable!str21 = retHeader(aHeader, 0, 1)
        temptable.Update
    End If
End If
With sourceTable
Do Until sourceTable.EOF
    temptable.AddNew
    temptable!Date1 = !Date
    temptable!str1 = !doc_no
    temptable!str2 = !Desca
    temptable!str3 = !CodeDesca
    If !Flag = 2 Or !Flag = 3 Or !Flag = 5 Or !Flag = 6 Or !Flag = 11 Or !Flag = 10 Then
        temptable!val2 = -1 * Val(!Value & "")
    Else
        temptable!val1 = !Value
    End If
    nBalance = nBalance + Val(!Value & "")
    temptable!val3 = nBalance
    temptable!str10 = " ›’Ì·Ì Õ—ﬂ… " & BalBox.XBOX.Text
    temptable!str21 = retHeader(aHeader, 0, 1)
    temptable.Update
    sourceTable.MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
Else
    contemp.BeginTrans
    contemp.CommitTrans
    main.REPORT1.ReportFileName = App.Path & "\Reports\Box2.rpt"
    main.REPORT1.DataFiles(0) = tempFile
    main.REPORT1.Action = 1
End If
sourceTable.Close: Set sourceTable = Nothing
temptable.Close: Set temptable = Nothing
loctable.Close: Set loctable = Nothing
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim loctable As New ADODB.Recordset
With GRID1
    .Cols = 7
    .Rows = 1
    .FormatString = "«· «—ÌŒ|" & "‰Ê⁄ «·Õ—ﬂ…|" & "—ﬁ„ «·„” ‰œ|" & "«·»Ì«‰|" & "«Ìœ«⁄« |" & "„”ÕÊ»« |" & "—’Ìœ"
    .ColDataType(0) = flexDTDate
    .ColWidth(0) = 1000
    .ColWidth(1) = 2000
    .ColWidth(2) = 900
    .ColWidth(3) = 2500
    .ColWidth(4) = 1100
    .ColWidth(5) = 1100
    .ColWidth(6) = 1200
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
End With

With BalBox
    If .vsBox.Row = 1 And Not IsDate(.vsBox.TextMatrix(.vsBox.Row, 0)) Then
        loctable.Open "Select * From BoxMove " & _
                      " where datevalue(date) < " & DateSq(.xDate1.Text) & _
                      " and box = " & MyParn(.XBOX.BoundText) & _
                      " Order by Date,Flag,Doc_no", con, adOpenStatic, adLockReadOnly, adCmdText
    Else
        loctable.Open "Select * From BoxMove " & _
                  " where datevalue(date) = " & DateSq(.vsBox.TextMatrix(.vsBox.Row, 0)) & _
                  " and box = " & MyParn(.XBOX.BoundText) & _
                  " Order by Date,Flag,Doc_no", con, adOpenStatic, adLockReadOnly, adCmdText
        GRID1.ColHidden(0) = True
        lblDate2.Visible = False
        xdate2.Visible = False
        Frame1.Height = Frame1.Height - 300
        GRID1.Top = GRID1.Top - 300
        Me.Height = Me.Height - 300
        Frame2.Top = Frame2.Top - 300
        xDate1.Text = Format(.vsBox.TextMatrix(.vsBox.Row, 0), "dd-mm-yyyy")
        xdate2.Text = Format(.vsBox.TextMatrix(.vsBox.Row, 0), "dd-mm-yyyy")
        lblDate1.Caption = "«·ÌÊ„ :"
        If .vsBox.Row > 1 Then
            If Val(.vsBox.TextMatrix(.vsBox.Row - 1, 2)) <> 0 Then
                GRID1.AddItem ""
                GRID1.TextMatrix(GRID1.Rows - 1, 1) = "—’Ìœ ”«»ﬁ"
                GRID1.TextMatrix(GRID1.Rows - 1, 4) = .vsBox.TextMatrix(.vsBox.Row - 1, 2)
                GRID1.TextMatrix(GRID1.Rows - 1, 6) = .vsBox.TextMatrix(.vsBox.Row - 1, 2)
                nBalance = Val(.vsBox.TextMatrix(.vsBox.Row - 1, 2))
            End If
        End If
    End If
    Do Until loctable.EOF
        GRID1.AddItem ""
        GRID1.TextMatrix(GRID1.Rows - 1, 0) = Format(loctable!Date, "dd-mm-yyyy")
        GRID1.TextMatrix(GRID1.Rows - 1, 1) = loctable!Desca & ""
        GRID1.TextMatrix(GRID1.Rows - 1, 2) = loctable!doc_no & ""
        GRID1.TextMatrix(GRID1.Rows - 1, 3) = loctable!CodeDesca & ""
        If loctable!Flag = 2 Or loctable!Flag = 3 Or loctable!Flag = 5 Or loctable!Flag = 6 Or loctable!Flag = 11 Or loctable!Flag = 10 Then
            GRID1.TextMatrix(GRID1.Rows - 1, 5) = Round(Val(loctable!Value & "") * -1)
        Else
            GRID1.TextMatrix(GRID1.Rows - 1, 4) = Round(Val(loctable!Value & ""), 2)
        End If
        nBalance = nBalance + Val(GRID1.TextMatrix(GRID1.Rows - 1, 4)) - Val(GRID1.TextMatrix(GRID1.Rows - 1, 5))
        GRID1.TextMatrix(GRID1.Rows - 1, 6) = Round(nBalance, 2)
        loctable.MoveNext
    Loop
    If .vsBox.Row = 1 And Not IsDate(.vsBox.TextMatrix(.vsBox.Row, 0)) Then
        xDate1.Text = GRID1.TextMatrix(1, 0)
        xdate2.Text = GRID1.TextMatrix(GRID1.Rows - 1, 0)
    End If
End With
GRID1.Select 0, 0, 0, 0
GRID1.Sort = flexSortGenericAscending
End Sub
