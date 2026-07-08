VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Supler_Move 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… «·„Ê—œÌ‰"
   ClientHeight    =   9675
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   2520
      Visible         =   0   'False
      Width           =   2310
      _ExtentX        =   4075
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
      Height          =   1005
      Left            =   8460
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   45
      Width           =   6675
      Begin VB.TextBox XDATE2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   270
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox xdate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox XCODE 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Õ Ì  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5445
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "ﬂÊœ «·„Ê—œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5415
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   825
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   180
         Width           =   3300
      End
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1920
      Width           =   405
   End
   Begin VSFlex7Ctl.VSFlexGrid invGrid 
      Height          =   5955
      Left            =   270
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   14865
      _cx             =   26220
      _cy             =   10504
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   11.25
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
      SelectionMode   =   0
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
      AutoResize      =   0   'False
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
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   270
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   405
      Width           =   4650
      Begin VB.CommandButton CmdGo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "≈ŸÂ«— «·Õ—ﬂ…"
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
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   1515
      End
      Begin VB.CommandButton CMD_PRINT 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ÿ»«⁄… "
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
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   1515
      End
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00E0E0E0&
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
         Height          =   390
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   180
         Width           =   1365
      End
   End
   Begin VB.Frame Frame3 
      Height          =   555
      Left            =   5850
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   495
      Width           =   2580
      Begin VB.Label xBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   135
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   180
         Width           =   525
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid VSTOT 
      Height          =   2490
      Left            =   270
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   7065
      Width           =   9150
      _cx             =   16140
      _cy             =   4392
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12648447
      ForeColor       =   -2147483640
      BackColorFixed  =   15523803
      ForeColorFixed  =   -2147483630
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648447
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
      Rows            =   5
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   450
      RowHeightMax    =   450
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
End
Attribute VB_Name = "Supler_Move"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim ClientTable As New ADODB.Recordset
Sub Fillgrd()
Dim loctable As New ADODB.Recordset
cString = "select FILE4_11_2.*,FILE4_12.desca as MoveDesca  " & _
          " From FILE4_11_2 Left join FILE4_12 on FILE4_11_2.[type] = FILE4_12.code"
cString = cString & turnFound(cString) & " FILE4_11_2.code = " & MyParn(XCODE.text)

If IsDate(xdate1.text) Then
    cString = cString & turnFound(cString) & " FILE4_11_2.date >= " & DateSq(xdate1.text)
End If

If IsDate(XDATE2.text) Then
    cString = cString & turnFound(cString) & " FILE4_11_2.date <= " & DateSq(XDATE2.text)
End If

cString = cString & " order by Date,FILE4_12.[ORDER],FILE4_11_2.SAL,FILE4_11_2.doc_id"

With invGrid
    .Rows = 1
    If IsDate(xdate1.text) Then
       cString2 = "Select sum([SAL] - PAY) as Balance from FILE4_11_2 where FILE4_11_2.CODE = " & MyParn(XCODE.text) & _
                  " and FILE4_11_2.date < " & DateSq(xdate1.text)
       nPrevious = Val(GetDesca(cString2, con))
       If nPrevious <> 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate1.text
            If nPrevious >= 0 Then
                .TextMatrix(.Rows - 1, 4) = Format(nPrevious, "#0.00")
            Else
                .TextMatrix(.Rows - 1, 3) = Format(nPrevious * -1, "#0.00")
            End If
              .TextMatrix(.Rows - 1, 5) = Format(nPrevious, "#0.00")
       End If
    End If

    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext
    Do Until loctable.EOF
         invGrid.AddItem ""
         nPrevious = nPrevious + Val(loctable!SAL & "") - Val(loctable!PAY & "")
        If IsNull(Trim(loctable!Desca)) Then
            .TextMatrix(.Rows - 1, 0) = loctable!MoveDesca & ""
        Else
            .TextMatrix(.Rows - 1, 0) = Trim(loctable!Desca)
        End If
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "yyyy/mm/dd")
        .TextMatrix(.Rows - 1, 2) = loctable!DOC_ID & ""
        .TextMatrix(.Rows - 1, 3) = Format(TurnValue(Val(loctable!PAY & ""), 0, ""), "#0.00")
        .TextMatrix(.Rows - 1, 4) = Format(TurnValue(Val(loctable!SAL & ""), 0, ""), "#0.00")
        .TextMatrix(.Rows - 1, 5) = Format(nPrevious, "#0.00")
        .TextMatrix(.Rows - 1, 6) = loctable!Type & ""
        .TextMatrix(.Rows - 1, 7) = loctable!STORE & ""
        
'        If loctable!Type = "4" Then .TextMatrix(.Rows - 1, 7) = GetDesca("SELECT STORE FROM FILE7_20H WHERE DOC_NO = " & MyParn(loctable!doc_id), con)
'        If loctable!Type = "5" Then .TextMatrix(.Rows - 1, 7) = GetDesca("SELECT STORE FROM FILE7_30H WHERE DOC_NO = " & MyParn(loctable!doc_id), con)
        
        loctable.MoveNext
    Loop
    If .Rows > 1 Then
        .SubtotalPosition = flexSTBelow
        .Subtotal flexSTSum, -1, 3, "#0.00", vbYellow, vbRed, True, "  "
        .Subtotal flexSTSum, -1, 4, "#0.00", vbYellow, vbRed, True, "  "
        .TextMatrix(.Rows - 1, 0) = "«·«Ã„«·Ì"
        .TextMatrix(.Rows - 1, 5) = Format(Round(nPrevious, 2), "#0.00")
    End If
End With
xBal.Caption = Format(nPrevious, "#0.00")
'invGrid.ColComboList(7) = StrList("SELECT CODE , DESCA FROM FILE0_40", con)

End Sub
Sub myProc()
ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Search3.Hide
End Sub
Function MYVALID() As Boolean
If XCODE.text = "" Then
    MsgBox "ﬂÊœ €Ì— „”Ã·"
    Exit Function
End If
ClientTable.Find " code = " & MyParn(XCODE.text), , adSearchForward, adBookmarkFirst
If ClientTable.EOF Then
    MsgBox "ﬂÊœ €Ì— ’ÕÌÕ"
    Exit Function
End If
If (Not IsDate(xdate1.text)) And Trim(xdate1.text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(XDATE2.text)) And Trim(XDATE2.text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub cmdcorect_Click()

End Sub

Private Sub CMD_PRINT_Click()
    Load rpSup12
    rpSup12.XCODE.text = XCODE.text
    rpSup12.xdate1.text = xdate1.text
    rpSup12.XDATE2.text = XDATE2.text
    rpSup12.Show 1
End Sub

Private Sub cmdGo_Click()
    If Not MYVALID Then Exit Sub
    Fillgrd
    FixTot
End Sub
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub Form_Load()
openCon con

XDATE2.text = Format(Date, "DD-MM-YYYY")
ClientTable.Open "FILE4_10", con, adOpenStatic, adLockReadOnly, adCmdTable

With invGrid
invGrid.Cols = 8
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = " «—ÌŒ"
.TextMatrix(0, 2) = "„” ‰œ"
.TextMatrix(0, 3) = "„œÌ‰"
.TextMatrix(0, 4) = "œ«∆‰"
.TextMatrix(0, 5) = "—’Ìœ"

.TextMatrix(0, 7) = "«·›—⁄"

invGrid.ColWidth(0) = 5000
invGrid.ColWidth(1) = 1500
invGrid.ColWidth(2) = 1500
invGrid.ColWidth(3) = 1500
invGrid.ColWidth(4) = 1500
invGrid.ColWidth(5) = 1500
invGrid.ColWidth(6) = 0
invGrid.ColWidth(7) = 1500

invGrid.ColComboList(7) = StrList("SELECT CODE , DESCA FROM BRANCH ", con)

xdate1.text = "1-1-" & Year(Date)
End With
For i = 0 To invGrid.Cols - 1
    invGrid.ColAlignment(i) = flexAlignRightCenter
Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Unload Search3
Err.Clear
End Sub
Private Sub invGrid_dblClick()
    If invGrid.TextMatrix(invGrid.Row, 6) = "4" Or invGrid.TextMatrix(invGrid.Row, 6) = "5" Then
        cDoc_No = invGrid.TextMatrix(invGrid.Row, 2)
        Purchasefrm.myPublic = IIf(invGrid.TextMatrix(invGrid.Row, 6) = "4", 0, 1)
        Purchasefrm.sDoc_no = cDoc_No
        Purchasefrm.Show
    End If
End Sub

Private Sub XDATE1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdGo_Click
End Sub

Private Sub xCode_Change()
invGrid.Rows = 1
CmdGo.Enabled = Trim(XCODE.text) <> ""
End Sub

Private Sub xcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdGo_Click
End Sub

Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then SendKeys "{Tab}"
If KeyCode = 112 Then CardLookup
End Sub
Private Sub xCode_LostFocus()
xDesca.Caption = ""
If Trim(XCODE.text) = "" Then Exit Sub
XCODE.text = RetZero(XCODE.text, 3)
ClientTable.Find "code = " & MyParn(XCODE.text), , adSearchForward, adBookmarkFirst
If Not ClientTable.EOF Then xDesca.Caption = ClientTable!Desca & ""
End Sub
Private Sub xStore_Click(Area As Integer)
If Not CmdGo.Enabled Then CmdGo.Enabled = True
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE4_10"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·»Ì«‰"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·»Ì«‰"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Sub FixTot()
    Dim n11 As Double
    Dim n12 As Double
    Dim n13 As Double
    Dim n14 As Double
    Dim n15 As Double
    Dim n16 As Double
    Dim n17 As Double
    Dim nFBal As Double
    Dim nCBal As Double
    cStrW = " and date >= " & DateSq(xdate1.text)
    If IsDate(XDATE2.text) Then cStrW = cStrW & " and date <= " & DateSq(XDATE2.text)
    n11 = Val(GetDesca("SELECT SUM(SAL)  FROM FILE4_11_2 WHERE [TYPE] = '4' AND CODE = " & MyParn(XCODE.text) & cStrW, con) & "")
    n12 = Val(GetDesca("SELECT SUM(PAY)  FROM FILE4_11_2 WHERE [TYPE] = '5' AND CODE = " & MyParn(XCODE.text) & cStrW, con) & "")
    n13 = n11 - n12
    n14 = Val(GetDesca("SELECT SUM(PAY - SAL )  FROM FILE4_11_2 WHERE ( [TYPE] = '7' OR [TYPE] = '8') AND CODE = " & MyParn(XCODE.text) & cStrW, con) & "")
    n15 = Val(GetDesca("SELECT SUM(PAY - SAL )  FROM FILE4_11_2 WHERE ([TYPE] = 'A' OR [TYPE] = 'C' ) AND CODE = " & MyParn(XCODE.text) & cStrW, con) & "")
    
    n13 = n11 - n12
    n16 = n14 + n15
    n17 = Val(GetDesca("SELECT SUM(VALUE)  FROM FILE5_21 WHERE [CLOSED] = '0' AND CODE1 = " & MyParn(XCODE.text), con) & "")
    
    
    nFBal = Val(GetDesca("SELECT SUM(SAL - PAY )  FROM FILE4_11_2 WHERE CODE = " & MyParn(XCODE.text) & " AND DATE < " & DateSq(xdate1.text), con) & "")
    nCBal = Val(GetDesca("SELECT SUM(SAL - PAY )  FROM FILE4_11_2 WHERE CODE = " & MyParn(XCODE.text) & " AND DATE <= " & DateSq(XDATE2.text), con) & "")
    
    With VSTOT
        .Rows = 5
        .Cols = 6
        .FixedCols = 0
        .FixedRows = 0
        .TextMatrix(0, 0) = "—’Ìœ √Ê·"
        .TextMatrix(1, 0) = "Ã. „‘ —Ì« "
        .TextMatrix(1, 2) = "Ã. „— Ã⁄« "
        .TextMatrix(1, 4) = "’«›Ï „‘ —Ì« "
    
        .TextMatrix(2, 0) = "œ›⁄«  ‰ﬁœÏ"
        .TextMatrix(2, 2) = "œ›⁄«  ‘Ìﬂ« "
        .TextMatrix(2, 4) = "≈Ã„«·Ï œ›⁄« "
        .TextMatrix(3, 4) = "√Ê—«ﬁ œ›⁄ €Ì— „”œœ…"
        .TextMatrix(4, 4) = "—’Ìœ Õ«·Ï"
        
        .TextMatrix(0, 1) = Round(nFBal, 2)
        .TextMatrix(1, 1) = Round(n11, 2)
        .TextMatrix(1, 3) = Round(n12, 2)
        .TextMatrix(1, 5) = Round(n13, 2)
    
        .TextMatrix(2, 1) = Round(n14, 2)
        .TextMatrix(2, 3) = Round(n15, 2)
        .TextMatrix(2, 5) = Round(n16, 2)
        
        .TextMatrix(3, 5) = Round(n17, 2)
        .TextMatrix(4, 5) = Round(nCBal, 2)
        For i = 0 To 5
            .ColWidth(i) = 1500
        Next i
    End With
End Sub

