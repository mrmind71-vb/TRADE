VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ClientMoveFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… «·⁄„·«¡"
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   180
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   360
      Width           =   2850
      Begin VB.CommandButton Command1 
         Caption         =   " ÿ»«⁄…"
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
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   135
         Width           =   1320
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
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         Width           =   1365
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid VSTOT 
      Height          =   2085
      Left            =   270
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8505
      Width           =   9150
      _cx             =   16140
      _cy             =   3678
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
      Rows            =   50
      Cols            =   6
      FixedRows       =   0
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   180
      Top             =   45
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Left            =   8370
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
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
         Caption         =   "ﬂÊœ «·⁄„Ì·"
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
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   270
         Width           =   855
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
      Height          =   7215
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   14865
      _cx             =   26220
      _cy             =   12726
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
      BackColorSel    =   16777152
      ForeColorSel    =   128
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
      Left            =   6660
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   1680
      Begin VB.CommandButton CmdGo 
         BackColor       =   &H00ECDFDB&
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      Height          =   645
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   360
      Width           =   2670
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
         Top             =   180
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
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   525
      End
   End
End
Attribute VB_Name = "ClientMoveFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim ClientTable As New ADODB.Recordset
Dim nPrev As Double
Public Sub Fillgrd()
Dim cStrW As String
Dim loctable As New ADODB.Recordset
Dim datatable As New ADODB.Recordset
Dim n11 As Double
Dim n12 As Double
Dim n13 As Double
Dim n14 As Double
Dim n15 As Double
Dim n16 As Double
Dim n17 As Double
Dim n18 As Double
cString = "select FILE3_11.*,file3_12.desca as moveDesca  " & _
          " From FILE3_11 Left join file3_12 on FILE3_11.type = file3_12.code"

cString = cString & turnFound(cString) & " FILE3_11.code = " & MyParn(xCode.text)

If IsDate(xdate1.text) Then
    cString = cString & turnFound(cString) & " FILE3_11.date >= " & DateSq(xdate1.text)
End If

If IsDate(xDate2.text) Then
    cString = cString & turnFound(cString) & " FILE3_11.date <= " & DateSq(xDate2.text)
End If
cString = cString & " Order by FILE3_11.date,file3_12.[order],FILE3_11.doc_id,pay"
With invGrid
    .Rows = 1
    If IsDate(xdate1.text) Then
        cString2 = "Select sum([SAL] - PAY)  as Balance from FILE3_11 where FILE3_11.CODE = " & MyParn(xCode.text) & _
                  " and FILE3_11.date < " & DateSq(xdate1.text)
        nPrevious = Round(Val(GetDesca(cString2, con)))
        nPrev = nPrevious
        If nPrevious <> 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate1.text
            .TextMatrix(.Rows - 1, 3) = Format(nPrevious, "Fixed")
       End If
    End If

    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext

    Do Until loctable.EOF
         invGrid.AddItem ""
         nPrevious = nPrevious + Round(Val(loctable!SAL & ""), 2) - Round(Val(loctable!PAY & ""), 2)
         
        .TextMatrix(.Rows - 1, 0) = Trim(loctable!DESCA & "")
        
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Date, "yyyy/mm/dd")
        If Len(loctable!doc_id) > 8 Then
            .TextMatrix(.Rows - 1, 2) = Mid(loctable!doc_id & "", 6, 20)
        Else
            .TextMatrix(.Rows - 1, 2) = loctable!doc_id & ""
        End If
        .TextMatrix(.Rows - 1, 3) = TurnValue(Round(Val(loctable!SAL & ""), 2), 0, "")
        .TextMatrix(.Rows - 1, 4) = TurnValue(Round(Val(loctable!PAY & ""), 2), 0, "")
        .TextMatrix(.Rows - 1, 5) = Round(nPrevious, 2)
        .TextMatrix(.Rows - 1, 6) = loctable!Type
        loctable.MoveNext
    Loop
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 3, "#0.00", vbYellow, vbRed, True, "  "
    .Subtotal flexSTSum, -1, 4, "#0.00", vbYellow, vbRed, True, "  "
    .TextMatrix(.Rows - 1, 0) = "«·«Ã„«·Ì"
    .TextMatrix(.Rows - 1, 5) = Round(nPrevious, 2)
End With
xBal.Caption = Round(nPrevious, 2)

If IsDate(xdate1.text) Then
    cStrW = cStrW & " AND DATE >= " & DateSq(xdate1.text)
End If
If IsDate(xDate2.text) Then
    cStrW = cStrW & " AND DATE <= " & DateSq(xDate2.text)
End If


cStr1 = " SELECT SUM(SAL + PAY )  AS VALMOVE , [TYPE] FROM FILE3_11 WHERE CODE = " & MyParn(xCode.text) & cStrW & " group by [type] "
datatable.Open cStr1, con, adOpenKeyset, adLockOptimistic, adCmdText

n11 = Val(GetDesca("SELECT SUM(SAL)  FROM FILE3_11 WHERE [TYPE] = '4' AND CODE = " & MyParn(xCode.text) & cStrW, con) & "")
n12 = Val(GetDesca("SELECT SUM(PAY)  FROM FILE3_11 WHERE [TYPE] = '5' AND CODE = " & MyParn(xCode.text) & cStrW, con) & "")
n13 = n11 - n12
n14 = Val(GetDesca("SELECT SUM(PAY - SAL )  FROM FILE3_11 WHERE ([TYPE] = '10' OR [TYPE] = '11' OR [TYPE] = '7' OR [TYPE] = '8') AND CODE = " & MyParn(xCode.text) & cStrW, con) & "")
n18 = Val(GetDesca("SELECT SUM(PAY)  FROM FILE3_11 WHERE [TYPE] = '9' AND CODE = " & MyParn(xCode.text) & cStrW, con) & "")



n13 = n11 - n12
n16 = n14 + n15
n17 = Val(GetDesca("SELECT SUM(VALUE)  FROM FILE5_20 WHERE [CLOSED] = '0' AND CODE1 = " & MyParn(xCode.text), con) & "")
With VsTot
    .Rows = 3
    .Cols = 6
    .FixedCols = 0
    .FixedRows = 0
    .TextMatrix(0, 0) = "Ã. „»Ì⁄« "
    .TextMatrix(0, 2) = "Ã. „— Ã⁄« "
    .TextMatrix(0, 4) = "’«›Ï „»Ì⁄«  "

    .TextMatrix(1, 0) = "œ›⁄«  ‰ﬁœÏ"
    .TextMatrix(1, 2) = "œ›⁄«  ‘Ìﬂ« "
    .TextMatrix(1, 4) = "≈Ã„«·Ï œ›⁄« "
    .TextMatrix(2, 4) = "√Ê—«ﬁ ﬁ»÷ €Ì— „Õ’·…"
    
    .TextMatrix(2, 0) = " ”ÊÌ« "
    .TextMatrix(2, 1) = n18
    
    .TextMatrix(0, 1) = n11
    .TextMatrix(0, 3) = n12
    .TextMatrix(0, 5) = n13

    .TextMatrix(1, 1) = n14
    .TextMatrix(1, 3) = n15
    .TextMatrix(1, 5) = n16
    
    .TextMatrix(2, 5) = n17
    For i = 0 To 5
        .ColWidth(i) = 1500
    Next i
End With
End Sub
Sub myProc()
ActiveControl.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Search3.Hide
End Sub
Function MYVALID() As Boolean
If xCode.text = "" Then
    MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
    Exit Function
End If

If GetDesca("select Desca from file3_10 where code = " & MyParn(xCode.text), con) = "" Then
    MsgBox "ﬂÊœ «·⁄„Ì· €Ì— ’ÕÌÕ"
    Exit Function
End If

If (Not IsDate(xdate1.text)) And Trim(xdate1.text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
If (Not IsDate(xDate2.text)) And Trim(xDate2.text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub cmdcorect_Click()

End Sub

Private Sub Check1_Click()
If Not MYVALID Then Exit Sub
Fillgrd
End Sub

Private Sub cmdGo_Click()
If Not MYVALID Then Exit Sub
Fillgrd
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub Command1_Click()
    doprint
End Sub


Private Sub Form_Load()
Dim cDocNo As String
With invGrid
invGrid.Cols = 7
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = " «—ÌŒ"
.TextMatrix(0, 2) = "„” ‰œ"
.TextMatrix(0, 3) = "„œÌ‰"
.TextMatrix(0, 4) = "œ«∆‰"
.TextMatrix(0, 5) = "—’Ìœ"

invGrid.ColWidth(0) = 5000
invGrid.ColWidth(1) = 1500
invGrid.ColWidth(2) = 1500
invGrid.ColWidth(3) = 1500
invGrid.ColWidth(4) = 1500
invGrid.ColWidth(5) = 1500
invGrid.ColWidth(6) = 0
End With
For i = 0 To invGrid.Cols - 1
    invGrid.ColAlignment(i) = flexAlignRightCenter
Next
openCon con
End Sub

Private Sub Form_Unload(Cancel As Integer)
closeCon con
SalesFrm.sDoc_no = cDoc_No
On Error Resume Next
Unload Search3
Err.Clear
End Sub

Private Sub invGrid_dblClick()
    Dim cDoc_No As String
    Select Case invGrid.TextMatrix(invGrid.Row, 6)
        Case "4"
            If cBranch = "00" Then
                cDoc_No = invGrid.TextMatrix(invGrid.Row, 2)
                SalesFrm.sDoc_no = cDoc_No
                SalesFrm.Show
            Else
                cDoc_No = invGrid.TextMatrix(invGrid.Row, 2)
                SalesFrm.myPublic = IIf(invGrid.TextMatrix(invGrid.Row, 6) = "4", 0, 1)
                SalesFrm.sDoc_no = cDoc_No
                SalesFrm.Show
            End If
    End Select
End Sub
Private Sub XDATE1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdGo_Click
End Sub
Private Sub xCode_Change()
invGrid.Rows = 1
cmdGo.Enabled = Trim(xCode.text) <> ""
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
If Trim(xCode.text) = "" Then Exit Sub

xDesca.Caption = GetDesca("select Desca from file3_10 where code = " & MyParn(xCode.text), con)
End Sub
Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From file3_10"
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
Private Sub doprint()
Dim nBalance As Double, nRow As Integer
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim n11 As Double, n12 As Double, n13 As Double, n14 As Double, n15 As Double, n16 As Double, n17 As Double
Dim cStrW As String

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
If Trim(xCode.text) <> "" Then
    aHeader(0) = "[" & "··⁄„Ì· : " & xDesca.Caption & "]"
End If
If IsDate(xdate1.text) Then
    aHeader(1) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    cStrW = cStrW & " AND DATE >= " & DateSq(xdate1.text)
End If
If IsDate(xDate2.text) Then
    aHeader(1) = "[" & BetweenString(xdate1.text, xDate2.text) & "]"
    cStrW = cStrW & " AND DATE <= " & DateSq(xDate2.text)
End If

n11 = Val(GetDesca("SELECT SUM(SAL)  FROM FILE3_11 WHERE [TYPE] = '4' AND CODE = " & MyParn(xCode.text) & cStrW, con) & "")
n12 = Val(GetDesca("SELECT SUM(PAY)  FROM FILE3_11 WHERE [TYPE] = '5' AND CODE = " & MyParn(xCode.text) & cStrW, con) & "")
n13 = n11 - n12
n14 = Val(GetDesca("SELECT SUM(PAY - SAL )  FROM FILE3_11 WHERE ([TYPE] = '10' OR [TYPE] = '11' OR [TYPE] = '7' OR [TYPE] = '8') AND CODE = " & MyParn(xCode.text) & cStrW, con) & "")
n18 = Val(GetDesca("SELECT SUM(PAY)  FROM FILE3_11 WHERE [TYPE] = '9' AND CODE = " & MyParn(xCode.text) & cStrW, con) & "")
'n14 = Val(GetDesca("SELECT SUM(PAY - SAL )  FROM FILE3_11 WHERE ([TYPE] = '10' OR [TYPE] = '11' OR [TYPE] = '7' OR [TYPE] = '8') AND CODE = " & MyParn(XCODE.Text) & cStrW) & "")

n15 = Val(GetDesca("SELECT SUM(PAY - SAL )  FROM FILE3_11 WHERE ([TYPE] = 'A' OR [TYPE] = 'C' ) AND CODE = " & MyParn(xCode.text) & cStrW, con) & "")
n16 = n14 + n15
n17 = Val(GetDesca("SELECT SUM(VALUE)  FROM FILE5_20 WHERE [CLOSED] = '0' AND CODE1 = " & MyParn(xCode.text), con) & "")

With invGrid
For i = 1 To .Rows - 2
    temptable.AddNew
    temptable!Date1 = TurnValue(RealDate(.TextMatrix(i, 1)))
    temptable!str1 = RetZero(DelZero(TurnValue(.TextMatrix(i, 2))), 5)
    temptable!str2 = TurnValue(.TextMatrix(i, 0))
    temptable!VAL1 = Val(.TextMatrix(i, 3))
    temptable!val2 = Val(.TextMatrix(i, 4))
    temptable!val3 = Val(.TextMatrix(i, 5))
    temptable!Val6 = i
    temptable!STR21 = TurnValue(retHeader(aHeader, 0, 3))
    temptable!STR20 = Firsttitle
    
    temptable!Val8 = nPrev
    temptable!Val10 = n11
    temptable!val11 = n12
    temptable!val12 = n13
    temptable!VAL13 = n14
    temptable!VAL14 = n15
    temptable!val15 = n16
    temptable!val16 = n17
    temptable!val17 = n18
    temptable.Update
Next
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\client3.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing

End Sub
Function RealDate(pDate, Optional cFormat As String = "") As String
If Not IsDate(pDate) Then Exit Function
RealDate = DateValue(Format(pDate, "dd/mm/yyyy"))
If cFormat <> "" Then RealDate = Format(RealDate, cFormat)
End Function
