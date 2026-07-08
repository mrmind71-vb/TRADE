VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form AccMove_Tot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈Ã„«·Ï Õ—ﬂ… Õ”«»"
   ClientHeight    =   11040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.TextBox xlevel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3150
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   765
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.CommandButton Cmd_print 
      Caption         =   "ÿ»«⁄… «” «– «·Õ”«»"
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   375
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1710
      Width           =   2100
   End
   Begin VB.CommandButton Cmdgo 
      BackColor       =   &H00C0FFFF&
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
      Height          =   390
      Left            =   2925
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1710
      Width           =   1725
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   615
      ScaleWidth      =   15270
      TabIndex        =   6
      Top             =   0
      Width           =   15270
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00C0FFFF&
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
         Height          =   465
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   75
         Width           =   1365
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid Grid1 
      Height          =   8010
      Left            =   360
      TabIndex        =   5
      Top             =   2205
      Width           =   14715
      _cx             =   25956
      _cy             =   14129
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
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
   Begin VB.TextBox xType 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3525
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   750
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   5040
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   675
      Width           =   10020
      Begin VB.TextBox date2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3300
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   1365
      End
      Begin VB.TextBox Date1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7050
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   607
         Width           =   1815
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   7050
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xCent 
         Height          =   315
         Left            =   5430
         TabIndex        =   16
         Top             =   990
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSAdodcLib.Adodc data1 
         Height          =   330
         Left            =   3330
         Top             =   990
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label2 
         Caption         =   "«·⁄„·… :"
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
         Index           =   1
         Left            =   8955
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1035
         Width           =   960
      End
      Begin VB.Label xBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   1965
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   4755
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   675
         Width           =   450
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ"
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
         Left            =   8970
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   645
         Width           =   675
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3300
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   3615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·Õ”«»"
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
         Left            =   8925
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   300
         Width           =   975
      End
   End
   Begin Threed.SSCheck XPRINTSE 
      Height          =   315
      Left            =   2925
      TabIndex        =   15
      Top             =   1350
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   556
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ÿ»«⁄… «·„Õœœ ›ﬁÿ"
      Alignment       =   1
   End
End
Attribute VB_Name = "AccMove_Tot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim AccTable As ADODB.Recordset, nFirst As Double
Sub Fillgrd()
nFirst = 0
Dim GrdTable As New ADODB.Recordset
If IsDate(Date1.Text) Then
    cString = "Select Sum([VAL1])AS TVAL1 ,Sum([VAL2]) AS TVAL2 FROM ACC_MOVE LEFT JOIN ACC1_10 ON ACC_MOVE.ACC = ACC1_10.CODE " & _
              " WHERE  Date < " & DateSq(Date1.Text)
    If xCode.Text <> "" Then
        cString = cString & " AND ACC = " & MyParn(xCode.Text)
    End If
    GrdTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
    If Not (GrdTable.BOF And GrdTable.EOF) Then nFirst = TurnValue(GrdTable!TVAL1, Null, 0) - TurnValue(GrdTable!TVAL2, Null, 0)
    GrdTable.Close
End If

cString = "Select ACC_MOVE.*     FROM ACC_MOVE Where ACC IS NOT NULL "
cString = cString & " AND ACC = " & MyParn(xCode.Text)

If IsDate(Date1.Text) Then cString = cString & " and [date] >= " & DateSq(Date1.Text)
If IsDate(date2.Text) Then cString = cString & " and [date] <= " & DateSq(date2.Text)
cString = cString & " Order by DATE , VAL1 "

GrdTable.Open cString, con, adOpenForwardOnly, adLockReadOnly, adCmdText

Dim nPrevious As Double
Dim nValue As Double
nPrevious = 0
Grid1.Rows = 1
With Grid1
If nFirst <> 0 Then
    Grid1.AddItem ""
    Grid1.TextMatrix(.Rows - 1, 0) = "—’Ìœ «Ê· «·„œ…"
    Grid1.TextMatrix(.Rows - 1, 1) = Date1.Text
    Grid1.TextMatrix(.Rows - 1, IIf(nFirst > 0, 3, 4)) = Format(nFirst, "Fixed")
    Grid1.TextMatrix(.Rows - 1, 5) = Format(nFirst, "Fixed")
    nPrevious = Format(nFirst, "Fixed")
End If

If (GrdTable.EOF And GrdTable.BOF) Then Exit Sub
Do
    Grid1.AddItem ""
    Grid1.TextMatrix(.Rows - 1, 0) = IIf(IsDate(GrdTable!Date), Format(GrdTable!Date, "d-m-yyyy"), "")
    Grid1.TextMatrix(.Rows - 1, 1) = TurnValue(GrdTable!DESCA, Null, "")
    Grid1.TextMatrix(.Rows - 1, 2) = TurnValue(GrdTable!DOC_ID, Null, "")
    Grid1.TextMatrix(.Rows - 1, 3) = Format(GrdTable!VAL1, "Fixed")
    Grid1.TextMatrix(.Rows - 1, 4) = Format(GrdTable!val2, "Fixed")
    nValue = Format(nPrevious + Val(Grid1.TextMatrix(.Rows - 1, 3)) - Val(Grid1.TextMatrix(.Rows - 1, 4)), "Fixed")
         
    
    Grid1.TextMatrix(.Rows - 1, 5) = Format(nValue, "Fixed")
    nCol3 = nCol3 + Val(.TextMatrix(.Rows - 1, 3))
    nCol4 = nCol4 + Val(.TextMatrix(.Rows - 1, 4))
    nPrevious = nValue
    GrdTable.MoveNext
    i = i + 1
Loop Until GrdTable.EOF
End With
xBal.Caption = Format(nValue, "Fixed")
End Sub
Sub myProc()
    xCode.Text = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 0)
    xDesca.Caption = Search3.Grid1.TextMatrix(Search3.Grid1.Row, 1)
    Unload Search3
End Sub
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdGo_Click()
    Fillgrd
End Sub

Private Sub Form_Load()
con.CursorLocation = adUseClient
con.Open strCon
Set AccTable = New ADODB.Recordset
AccTable.Open "ACC1_10", con, adOpenDynamic, adLockOptimistic, adCmdTableDirect

'data1.ConnectionString = strCon
'data1.RecordSource = "Select Code,DescA From File1_10SC order by Desca"
'Set xCent.RowSource = data1
'xCent.ListField = "Desca"
'xCent.BoundColumn = "Code"

Grid1.FormatString = " «—ÌŒ |" & "«·»Ì«‰ |" & "—ﬁ„ «·„” ‰œ |" & "„œÌ‰|" & "œ«∆‰|" & "—’Ìœ|" & "ÿ»«⁄…|"
Grid1.Cols = 7
Grid1.ColWidth(0) = 1400
Grid1.ColWidth(1) = 5500
Grid1.ColWidth(2) = 1600
Grid1.ColWidth(3) = 1600
Grid1.ColWidth(4) = 1600
Grid1.ColWidth(5) = 1600
Grid1.ColWidth(6) = 800
For i = 0 To Grid1.Cols - 1
    Grid1.ColAlignment(i) = flexAlignRightCenter
Next
Grid1.Rows = 1
Grid1.FixedCols = 0
Grid1.FixedRows = 1
Grid1.ColDataType(6) = flexDTBoolean
Date1.Text = "1-1-" & Year(Date)
date2.Text = Format(Date, "DD-MM-YYYY")
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub grid1_EnterCell()
    With Grid1
        If .Col = 6 Then
            .Editable = flexEDKbdMouse
        Else
            .Editable = flexEDNone
        End If
    End With
End Sub
Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        AcntLookup_Sub Me, Search3
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub xCode_LostFocus()
If xCode.Text = "" Then Exit Sub
AccTable.Find " code = " & MyParn(xCode.Text), , adSearchForward, adBookmarkFirst
If Not AccTable.EOF Then
    xDesca.Caption = AccTable!DESCA & ""
    Cmdgo.Enabled = Not AccTable.EOF
End If
End Sub
Private Sub CMD_PRINT_Click()
    Dim temptable As New ADODB.Recordset
    Dim sourcetable As New ADODB.Recordset
    contemp.Execute "DELETE * FROM TEMP"
    temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
                        
    With Grid1
    For i = 1 To .Rows - 1
        If XPRINTSE.Value <> -1 Or TurnValue(.TextMatrix(i, 6), "", False) Then
            temptable.AddNew
            temptable!str19 = cComp
            
            temptable!str8 = "ÿ»«⁄… «” «– ⁄«„ ··Õ”«» —ﬁ„ " & xCode.Text & "   " & xDesca.Caption
            temptable!STR7 = " „‰  «—ÌŒ " & Date1.Text & "  ≈·Ï  «—ÌŒ " & date2.Text
            
            temptable!str1 = TurnValue(.TextMatrix(i, 2), "", Null)
            temptable!str3 = TurnValue(.TextMatrix(i, 1), "", Null)
            temptable!Date1 = IIf(IsDate(.TextMatrix(i, 0)), .TextMatrix(i, 0), Null)
            temptable!VAL1 = Val(.TextMatrix(i, 3))
            temptable!val2 = Val(.TextMatrix(i, 4))
            temptable!val3 = Val(.TextMatrix(i, 5))
        temptable!STR30 = cusername
            
            temptable.Update
        End If
    Next i
    End With

    contemp.BeginTrans
    contemp.CommitTrans
    If XPRINTSE.Value Then
        main.REPORT1.ReportFileName = App.Path & "\Reports\R_ACC30.rpt"
    Else
        main.REPORT1.ReportFileName = App.Path & "\Reports\R_ACC3.rpt"
    End If
    main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
    main.REPORT1.Action = 1
End Sub

