VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form onlineCSVfrm 
   BackColor       =   &H80000005&
   Caption         =   "”Õ» ÿ·»Ì«  „‰ „·› CSV"
   ClientHeight    =   11055
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   11055
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   4200
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5220
      Width           =   20220
      _cx             =   35666
      _cy             =   7408
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
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
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5010
      Left            =   90
      TabIndex        =   9
      Top             =   90
      Width           =   20220
      _cx             =   35666
      _cy             =   8837
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   0
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   27
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -765
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   0
      Top             =   -765
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
   Begin MSComDlg.CommonDialog Common1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6300
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   9360
      Width           =   14010
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "«·ﬂ·"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   12960
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   270
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ÿ·»Ì… ÃœÌœ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   225
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   270
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ﬂÊœ ’‰› Œÿ√"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   11340
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   270
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ÿ·»Ì… „ÊÃÊœ…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   1410
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ÿ·»Ì«  »›« Ê—…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   9090
         RightToLeft     =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   270
         Width           =   1635
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ÿ·»Ì«  »ﬂÊœ Œÿ√ «Ê »›« Ê—…"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   6165
         RightToLeft     =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   270
         Width           =   2445
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ÿ·»Ì«  ﬁ«»·… ··Õ›Ÿ"
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3870
         RightToLeft     =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   270
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   9360
      Width           =   5505
      Begin Threed.SSCommand CMD_SEND 
         Height          =   510
         Left            =   1755
         TabIndex        =   11
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_csv.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_csv.frx":2925
      End
      Begin Threed.SSCommand cmdScv 
         Height          =   510
         Left            =   3465
         TabIndex        =   12
         Top             =   180
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_csv.frx":5179
         Caption         =   "› Õ „” ‰œ CSV"
         ButtonStyle     =   3
         PictureAlignment=   10
         BevelWidth      =   0
         PictureDisabledFrames=   1
         ShapeSize       =   1
         PictureDisabled =   "online_csv.frx":774C
      End
      Begin Threed.SSCommand cmdExit 
         Height          =   510
         Left            =   45
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   180
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   900
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "online_csv.frx":9A42
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   14
      Top             =   10815
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "onlineCSVfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myForm As Form
Public con As New ADODB.Connection
Public bNoItems As Boolean
Private Sub CMD_SEND_Click()
If Not myValid Then
    Exit Sub
End If

If MsgBox(" ÕÊÌ· «·Ì «·ÿ·»Ì« ", vbOKCancel + vbDefaultButton2) <> vbOK Then
    Exit Sub
End If

Me.MousePointer = vbHourglass
CMD_SEND.Enabled = False
Dim nAdd As Long, nEdit As Long, sMsg As String
If myreplace(nAdd, nEdit) Then
    If nAdd > 0 Then sMsg = "«÷«›… ⁄œœ " & nAdd
    If nEdit > 0 Then sMsg = sMsg & Tr(sMsg, " Ê ") & " ⁄œÌ· " & nEdit
    If sMsg = "" Then
        MsgBox "·„ Ì „  ⁄œÌ· «Ê «÷«›… «Ì ”Ã·« "
    Else
        MsgBox " „ " & sMsg & " ”Ã·"
    End If
    
    grid1.Rows = 1
    GRID2.Rows = 1
    myForm.myload
End If
CMD_SEND.Enabled = True
Me.MousePointer = vbNormal
End Sub
Private Function myValid() As Boolean
Dim i As Long
With grid1
For i = 1 To .Rows - 1
    prog1.Value = Round(i / (.Rows - 1), 2) * 100
    Caption = sCaption & " - " & i & " „‰ " & .Rows - 1
    If .TextMatrix(i, 19) = "" And .ValueMatrix(i, .Cols - 1) = 0 Then
        myValid = True
        Exit Function
    End If
Next
End With
MsgBox "·«  ÊÃœ ”Ã·«  ’«·Õ… ·· ÕÊÌ·"
End Function
Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()

End Sub

Private Sub cmdScv_Click()
cmdScv.Enabled = False
Me.MousePointer = vbHourglass
getData2
Fixgrd2
Fixgrd
Me.MousePointer = vbNormal
cmdScv.Enabled = True
End Sub
Private Sub Form_Load()
If openCon(con) <> "ok" Then Exit Sub
Set grid1.DataSource = data1
Set GRID2.DataSource = DATA2
Fixgrd
Fixgrd2
'sStoreOnline = myField("SELECT CODE , DESCA FROM FILE0_40 WHERE ONLINE = 1 ORDER BY CODE ", con) & ""
'CellPos 13, 0, grid1.Cols - 1
End Sub
Sub Fixgrd2()
With GRID2
    .RowHeight(0) = 600
    .TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
    .TextMatrix(0, 1) = "SKU"
    .TextMatrix(0, 2) = "ﬂÊœ"
    .TextMatrix(0, 3) = "«·’‰›"
    .TextMatrix(0, 4) = "«·”⁄—"
    .TextMatrix(0, 5) = "«·ﬂ„Ì…"
    .TextMatrix(0, 6) = "›« Ê—… „»Ì⁄« "
    .TextMatrix(0, 7) = "„ÊÃÊœ"
    .ColDataType(7) = flexDTBoolean
    
    .ColWidth(0) = 1200
    .ColWidth(1) = 2000
    .ColWidth(2) = 1500
    .ColWidth(3) = 8000
    .ColWidth(4) = 1200
    .ColWidth(5) = 1000
    .ColWidth(6) = 1200
    .ColWidth(7) = 1200
    
    .MergeCol(0) = True
    .MergeCells = flexMergeRestrictColumns
    
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 2) = "" Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H8080FF
            nFound = grid1.FindRow(.TextMatrix(i, 0), , 1)
            If nFound <> -1 Then
                If grid1.ValueMatrix(nFound, grid1.Cols - 1) = 0 Then
                    grid1.TextMatrix(nFound, grid1.Cols - 1) = -1
                End If
          End If
        ElseIf .TextMatrix(i, 6) <> "" Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0C0FF
        ElseIf .ValueMatrix(i, 7) <> 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0FFC0
        End If
        
        If Option1(0).Value Then
            .RowHidden(i) = False
        ElseIf Option1(1).Value Then
            .RowHidden(i) = .ValueMatrix(i, 7) <> 0
        ElseIf Option1(2).Value Then
            .RowHidden(i) = .TextMatrix(i, 2) <> ""
        ElseIf Option1(3).Value Then
            .RowHidden(i) = .ValueMatrix(i, 7) = 0
        ElseIf Option1(4).Value Then
            .RowHidden(i) = .TextMatrix(i, 6) = ""
        ElseIf Option1(5).Value Then
            .RowHidden(i) = Not (.TextMatrix(i, 2) = "" Or .TextMatrix(i, 6) <> "")
        ElseIf Option1(6).Value Then
            .RowHidden(i) = .TextMatrix(i, 2) = "" Or .TextMatrix(i, 6) <> ""
        End If
    Next
End With
End Sub
Sub Fixgrd()
With grid1
    .RowHeight(0) = 600
    .TextMatrix(0, 0) = "„"
    .TextMatrix(0, 1) = "—ﬁ„ «·ÿ·»"
    .TextMatrix(0, 2) = "«· «—ÌŒ"
    .TextMatrix(0, 3) = "«·«”„"
    .TextMatrix(0, 4) = "«· ·Ì›Ê‰"
    .TextMatrix(0, 5) = "»—Ìœ «·Ìﬂ —Ê‰Ì"
    .TextMatrix(0, 6) = "›Ì„… «·«’‰«›"
    .TextMatrix(0, 7) = "‘Õ‰"
    .TextMatrix(0, 8) = "≈Ã„«·Ì"
    .TextMatrix(0, 9) = "Œ’„"
    .TextMatrix(0, 10) = "ﬂÊœ «·Œ’„"
    .TextMatrix(0, 11) = "⁄‰Ê«‰1"
    .TextMatrix(0, 12) = "⁄‰Ê«‰2"
    .TextMatrix(0, 13) = "«·‘«—⁄"
    .TextMatrix(0, 14) = "«·„œÌ‰…"
    .TextMatrix(0, 15) = "ÿ—Ìﬁ… «·”œ«œ"
    .TextMatrix(0, 16) = "„œÌ‰… «·‘Õ‰"
    .TextMatrix(0, 17) = "„·«ÕŸ« "
    .TextMatrix(0, 18) = "ID"
    
    .TextMatrix(0, 19) = "„” ‰œ «·»Ì⁄"
    .TextMatrix(0, 20) = "„ÊÃÊœ"
    .TextMatrix(0, 21) = "Tags"
    
    .TextMatrix(0, 22) = "«·„—Õ·…"
    .TextMatrix(0, 23) = "«·„—Õ·…"
    .TextMatrix(0, 24) = "«·„—Õ·… «·Õ«·Ì…"
    .TextMatrix(0, 25) = "«·„—Õ·… «·Õ«·Ì…"
    .TextMatrix(0, .Cols - 1) = "ﬂÊœ Œÿ√"
        
    
    .ColDataType(20) = flexDTBoolean
    .ColDataType(.Cols - 1) = flexDTBoolean
    
    .ColFormat(2) = "YYYY/M/D"
    .ColWidth(0) = 500
    .ColWidth(1) = 1000
    .ColWidth(2) = 1300
    .ColWidth(3) = 3000
    .ColWidth(4) = 2000
    .ColWidth(5) = 2000
    .ColWidth(6) = 1100
    .ColWidth(7) = 900
    .ColWidth(8) = 1150
    .ColWidth(9) = 1000
    .ColWidth(10) = 1150
    .ColWidth(11) = 5300
    .ColWidth(12) = 0
    .ColWidth(13) = 0
    .ColWidth(14) = 2000
    .ColWidth(15) = 1000
    .ColWidth(19) = 1000
    .ColWidth(20) = 800
    .ColWidth(21) = 2500
    .ColWidth(.Cols - 1) = 800
    
    .ColHidden(4) = True
    .ColHidden(5) = True
    .ColHidden(10) = True
    .ColHidden(11) = True
    .ColHidden(14) = True
    '.ColHidden(15) = True
    .ColHidden(16) = True
    .ColHidden(17) = True
    .ColHidden(18) = True
    .ColHidden(22) = True
    .ColHidden(24) = True
    
    Dim i As Long
    For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    
    .ColAlignment(15) = flexAlignLeftCenter
    .ColAlignment(21) = flexAlignLeftCenter
    If .Rows > 1 Then
        .Cell(flexcpFontSize, 1, 15, .Rows - 1, 15) = 10
        .Cell(flexcpFontSize, 1, 21, .Rows - 1, 21) = 10
    End If
    For i = 1 To grid1.Rows - 1
        .TextMatrix(i, 0) = i
        If .TextMatrix(i, 19) <> "" Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0C0FF
        ElseIf .ValueMatrix(i, 20) <> 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0FFC0
        ElseIf .ValueMatrix(i, .Cols - 1) <> 0 Then
            .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &H8080FF
        End If
        If Option1(0).Value Then
            .RowHidden(i) = False
        ElseIf Option1(1).Value Then
            .RowHidden(i) = .ValueMatrix(i, 20) <> 0
        ElseIf Option1(2).Value Then
            .RowHidden(i) = .ValueMatrix(i, .Cols - 1) = 0
        ElseIf Option1(3).Value Then
            .RowHidden(i) = .ValueMatrix(i, 20) = 0
        ElseIf Option1(4).Value Then
            .RowHidden(i) = .TextMatrix(i, 19) = ""
        ElseIf Option1(5).Value Then
            .RowHidden(i) = Not (.TextMatrix(i, 19) <> "" Or .ValueMatrix(i, .Cols - 1) <> 0)
        ElseIf Option1(6).Value Then
            .RowHidden(i) = (.TextMatrix(i, 19) <> "" Or .ValueMatrix(i, .Cols - 1) <> 0)
        End If
    Next
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set onlineCSVfrm = Nothing
End Sub
Private Function getData()
Dim cFileName As String
Dim sb As New ChilkatStringBuilder

Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName

If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
Dim i As Long
SetKbLayout Lang_AR

cSv.HasColumnNames = 1
'nAccess = cSv.LoadFile(cFileName)
nAccess = sb.LoadFile(cFileName, "UTF-8")
If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


nAccess = cSv.LoadFromString(sb.GetAsString)

If cSv.NumRows < 1 Then Exit Function

Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder
Dim locTable As New ADODB.RecordSet

prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption
For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    If Trim(cSv.GetCellByName(i, "Name")) <> "" Then
        If Trim(Replace(cSv.GetCellByName(i, "Name"), "#", "")) <> sDoc_no Then
            sDoc_no = Trim(Replace(cSv.GetCellByName(i, "Name"), "#", ""))
            Tb.Append "SELECT " & addstring(sDoc_no) & " AS DOC_NO,"
            Tb.Append addDate(Mid(cSv.GetCellByName(i, "Created at"), 1, 10)) & " AS [date],"
            Tb.Append addstring(cSv.GetCellByName(i, "Shipping Name")) & " AS NAME,"
            Tb.Append addstring(cSv.GetCellByName(i, "Shipping Phone")) & " AS PHONE,"
            Tb.Append addstring(cSv.GetCellByName(i, "Email")) & " AS E_MAIL,"
            Tb.Append Val(cSv.GetCellByName(i, "Subtotal")) & " AS Subtotal,"
            Tb.Append addstring(cSv.GetCellByName(i, "Shipping")) & " AS Shipping,"
            Tb.Append Val(cSv.GetCellByName(i, "Total")) & " AS Total,"
            Tb.Append Val(cSv.GetCellByName(i, "Discount Amount")) & " AS Discount,"
            Tb.Append addstring(cSv.GetCellByName(i, "Discount Code")) & " AS Discount_Code,"
            Tb.Append addstring(Replace(cSv.GetCellByName(i, "Shipping Street"), "?", "")) & " AS STREET,"
            Tb.Append addstring(Replace(cSv.GetCellByName(i, "Shipping Address1"), "?", "")) & " AS ADDRESS1,"
            Tb.Append addstring(Replace(cSv.GetCellByName(i, "Shipping Address2"), "?", "")) & " AS ADDRESS2,"
            Tb.Append addstring(Replace(cSv.GetCellByName(i, "Shipping City"), "?", "")) & " AS CITY,"
            Tb.Append addstring(cSv.GetCellByName(i, "Payment Method")) & " AS Payment_Method,"
            Tb.Append addstring(cSv.GetCellByName(i, "Shipping Province Name")) & " AS Shipping_City,"
            Tb.Append addstring(cSv.GetCellByName(i, "Notes")) & " AS Notes_Order,"
            Tb.Append addstring(cSv.GetCellByName(i, "Payment ID")) & " AS Payment_ID,"
            
            Set locTable = mycmd("select sales_doc,doc_no from file6_90h where doc_no = " & MyParn(sDoc_no), con)
            
            If Not locTable.EOF Then
                Tb.Append addstring(locTable!sales_Doc) & " AS SALES_DOC,"
                Tb.Append "1 AS FOUND,"
            Else
                Tb.Append "NULL AS SALES_DOC,"
                Tb.Append "0  AS FOUND,"
            End If
            Tb.Append "0 AS NO_ITEM"
            Tb.Append " UNION ALL "
        End If
    End If
Next
Tb.Shorten 11
prog1.Visible = False
Me.Caption = sCaption

grid1.Rows = 1
GRID2.Rows = 1

If Tb.length = 0 Then Exit Function
Set data1.RecordSet = mycmd(Tb.GetAsString, con)

prog1.Visible = True
Dim SKU As String
Dim sitem As String

Tb.Clear
sDoc_no = ""
For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    SKU = cSv.GetCellByName(i, "Lineitem sku")
    
    If Trim(Replace(cSv.GetCellByName(i, "Name"), "#", "")) <> "" Then
        sDoc_no = Trim(Replace(cSv.GetCellByName(i, "Name"), "#", ""))
        sitem = retItemCode(SKU, con) & ""
        sitem = IIf(IsNumeric(sitem), sitem, "")
        
        Tb.Append "SELECT " & addstring(sDoc_no) & " AS DOC_NO,"
        Tb.Append addstring(SKU) & " AS SKU,"
        Tb.Append addvalue(sitem) & " AS ITEM,"
        Tb.Append addstring(cSv.GetCellByName(i, "Lineitem name")) & " AS ITEM_NAME,"
        Tb.Append Val(cSv.GetCellByName(i, "Lineitem price")) & " AS PRICE,"
        Tb.Append Val(cSv.GetCellByName(i, "Lineitem quantity")) & " AS QUANT,"
                
        Set locTable = mycmd("select sales_doc,doc_no from file6_90h where doc_no = " & MyParn(sDoc_no), con)
        If Not locTable.EOF Then
            Tb.Append addstring(locTable!sales_Doc) & " AS SALES_DOC,"
            Tb.Append "1 AS FOUND"
        Else
            Tb.Append "NULL AS SALES_DOC,"
            Tb.Append "0  AS FOUND"
        End If
        Tb.Append " UNION ALL "
    End If
Next
Tb.Shorten 11
If Tb.length = 0 Then Exit Function
Set DATA2.RecordSet = mycmd(Tb.GetAsString, con)

Me.Caption = sCaption
prog1.Visible = False
getData = True
Finaly:
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GoTo Finaly
End Function

Private Sub Option1_Click(index As Integer)
Fixgrd2
Fixgrd
End Sub
Private Function myreplace(ByRef nAdd As Long, ByRef nEdit As Long) As Boolean
Dim i As Long, sCaption As String
Dim aInsert As Variant

prog1.Value = 0
prog1.Visible = True
sCaption = Me.Caption

con.BeginTrans
On Error GoTo myerror
For i = 1 To grid1.Rows - 1
    prog1.Value = Round(i / (grid1.Rows - 1), 2) * 100
    Me.Caption = sCaption & " - " & i & " „‰ " & grid1.Rows - 1
    If grid1.TextMatrix(i, 19) = "" And grid1.ValueMatrix(i, grid1.Cols - 1) = 0 Then
        aInsert = AddFlag(Empty, "[DOC_NO]", addstring(grid1.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(grid1.TextMatrix(i, 2)))
        aInsert = AddFlag(aInsert, "[NAME]", addstring(grid1.TextMatrix(i, 3)))
        aInsert = AddFlag(aInsert, "[PHONE]", addstring(grid1.TextMatrix(i, 4)))
        aInsert = AddFlag(aInsert, "[E_MAIL]", addstring(Mid(Trim(grid1.TextMatrix(i, 5)), 1, 50)))
        aInsert = AddFlag(aInsert, "[SUBTOTAL]", grid1.ValueMatrix(i, 6))
        aInsert = AddFlag(aInsert, "[SHIPPING]", grid1.ValueMatrix(i, 7))
        aInsert = AddFlag(aInsert, "[TOTAL]", grid1.ValueMatrix(i, 8))
        aInsert = AddFlag(aInsert, "[DISCOUNT]", grid1.ValueMatrix(i, 9))
        aInsert = AddFlag(aInsert, "[DISCOUNT_CODE]", addstring(grid1.TextMatrix(i, 10)))
        aInsert = AddFlag(aInsert, "[STREET]", addstring(grid1.TextMatrix(i, 11)))
        aInsert = AddFlag(aInsert, "[ADDRESS1]", addstring(grid1.TextMatrix(i, 12)))
        aInsert = AddFlag(aInsert, "[ADDRESS2]", addstring(grid1.TextMatrix(i, 13)))
        aInsert = AddFlag(aInsert, "[CITY]", addstring(grid1.TextMatrix(i, 14)))
        aInsert = AddFlag(aInsert, "[Payment_Method]", addstring(Mid(grid1.TextMatrix(i, 15), 1, 50)))
        aInsert = AddFlag(aInsert, "[Shipping_City]", addstring(grid1.TextMatrix(i, 16)))
        aInsert = AddFlag(aInsert, "[Notes_Order]", addstring(grid1.TextMatrix(i, 17)))
        aInsert = AddFlag(aInsert, "[Payment_ID]", addstring(grid1.TextMatrix(i, 18)))
        aInsert = AddFlag(aInsert, "[tags]", addstring(grid1.TextMatrix(i, 21)))
        If grid1.ValueMatrix(i, 24) = 0 Then
            aInsert = AddFlag(aInsert, "[stage]", grid1.ValueMatrix(i, 22))
        End If
        aInsert = AddFlag(aInsert, "[STORE]", addstring(sStoreOnline))
        aInsert = AddFlag(aInsert, "[SEND_USER]", addstring(cusername))
        aInsert = AddFlag(aInsert, "[SEND_TIME]", "GETDATE()")
        
        aInsert = AddFlag(aInsert, "[NEW]", "1")
        con.Execute addUpdate(aInsert, "FILE6_90H", "DOC_NO = " & MyParn(grid1.TextMatrix(i, 1))), nAffect
        If nAffect = 0 Then
            con.Execute addInsert(aInsert, "FILE6_90H"), nAffect
            nAdd = nAdd + 1
        Else
            nEdit = nEdit + 1
        End If
        If Not bNoItems Then
            con.Execute "Delete from file6_90 where doc_no = " & MyParn(grid1.TextMatrix(i, 1))
        End If
    End If
Next


prog1.Value = 0
prog1.Visible = True

Dim nFound As Long
For i = 1 To GRID2.Rows - 1
    prog1.Value = Round(i / (GRID2.Rows - 1), 2) * 100
    Me.Caption = sCaption & " - " & i & " „‰ " & GRID2.Rows - 1
    
    nFound = grid1.FindRow(GRID2.TextMatrix(i, 0), , 1)
    
    If grid1.TextMatrix(nFound, 19) = "" And grid1.ValueMatrix(nFound, grid1.Cols - 1) = 0 Then
        If GRID2.TextMatrix(i, 2) <> "" And GRID2.TextMatrix(i, 6) = "" And ((Not bNoItems) Or grid1.ValueMatrix(nFound, 20) = 0) Then
            aInsert = AddFlag(Empty, "[DOC_NO]", addstring(GRID2.TextMatrix(i, 0)))
            aInsert = AddFlag(aInsert, "[SKU]", addstring(GRID2.TextMatrix(i, 1)))
            aInsert = AddFlag(aInsert, "[ITEM]", addvalue(GRID2.TextMatrix(i, 2)))
            aInsert = AddFlag(aInsert, "[ITEM_NAME]", addstring(GRID2.TextMatrix(i, 3)))
            aInsert = AddFlag(aInsert, "[PRICE]", GRID2.ValueMatrix(i, 4))
            aInsert = AddFlag(aInsert, "[QUANT]", GRID2.ValueMatrix(i, 5))
            aInsert = AddFlag(aInsert, "[NEW]", "1")
            con.Execute addInsert(aInsert, "FILE6_90") & ";"
        End If
    End If
Next
con.CommitTrans
myreplace = True
Finally:
Me.Caption = sCaption
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
myreplace = -1
GoTo Finally
End Function
Private Function getData2()
Dim cFileName As String
Dim sb As New ChilkatStringBuilder

Common1.InitDir = ""
Common1.FileName = ""
Common1.Filter = "CSV (*.CSV*)|*.CSV*"
Common1.ShowOpen

cFileName = Common1.FileName

If cFileName = "" Then Exit Function

Dim cSv As New ChilkatCsv
Dim i As Long
SetKbLayout Lang_AR

cSv.HasColumnNames = 1
'nAccess = cSv.LoadFile(cFileName)
nAccess = sb.LoadFile(cFileName, "UTF-8")
If nAccess = 0 Then
    MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰  Õ„Ì· «·„·›"
    Exit Function
End If


nAccess = cSv.LoadFromString(sb.GetAsString)

If cSv.NumRows < 1 Then Exit Function

Dim Tb As New ChilkatStringBuilder
Dim cString As New ChilkatStringBuilder
Dim locTable As ADODB.RecordSet

prog1.Visible = True
Dim sCaption As String
sCaption = Me.Caption

Dim db As New clsDb

Dim stageTable As ADODB.RecordSet
Set stageTable = db.myRs("select * from STAGES_CODES")
Dim nStage As Variant


grid1.Rows = 1
GRID2.Rows = 1
For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    
    If Trim(cSv.GetCellByName(i, "Name")) <> "" Then
        If Trim(Replace(cSv.GetCellByName(i, "Name"), "#", "")) <> sDoc_no Then
            grid1.AddItem ""
            sDoc_no = Trim(Replace(cSv.GetCellByName(i, "Name"), "#", ""))
            grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows
            grid1.TextMatrix(grid1.Rows - 1, 1) = sDoc_no
            grid1.TextMatrix(grid1.Rows - 1, 2) = myFormat_p(Mid(cSv.GetCellByName(i, "Created at"), 1, 10))
            grid1.TextMatrix(grid1.Rows - 1, 3) = cSv.GetCellByName(i, "Shipping Name")
            grid1.TextMatrix(grid1.Rows - 1, 4) = cSv.GetCellByName(i, "Shipping Phone")
            grid1.TextMatrix(grid1.Rows - 1, 5) = cSv.GetCellByName(i, "Email")
            grid1.TextMatrix(grid1.Rows - 1, 6) = Val(cSv.GetCellByName(i, "Subtotal"))
            grid1.TextMatrix(grid1.Rows - 1, 7) = cSv.GetCellByName(i, "Shipping")
            grid1.TextMatrix(grid1.Rows - 1, 8) = Val(cSv.GetCellByName(i, "Total"))
            grid1.TextMatrix(grid1.Rows - 1, 9) = Val(cSv.GetCellByName(i, "Discount Amount"))
            grid1.TextMatrix(grid1.Rows - 1, 10) = cSv.GetCellByName(i, "Discount Code")
            grid1.TextMatrix(grid1.Rows - 1, 11) = Replace(cSv.GetCellByName(i, "Shipping Street"), "?", "")
            grid1.TextMatrix(grid1.Rows - 1, 12) = Replace(cSv.GetCellByName(i, "Shipping Address1"), "?", "")
            grid1.TextMatrix(grid1.Rows - 1, 13) = Replace(cSv.GetCellByName(i, "Shipping Address2"), "?", "")
            grid1.TextMatrix(grid1.Rows - 1, 14) = Replace(cSv.GetCellByName(i, "Shipping City"), "?", "")
            grid1.TextMatrix(grid1.Rows - 1, 15) = cSv.GetCellByName(i, "Payment Method")
            grid1.TextMatrix(grid1.Rows - 1, 16) = cSv.GetCellByName(i, "Shipping Province Name")
            grid1.TextMatrix(grid1.Rows - 1, 17) = cSv.GetCellByName(i, "Notes")
            grid1.TextMatrix(grid1.Rows - 1, 18) = cSv.GetCellByName(i, "Payment ID")
            
            Set locTable = db.myRs("select sales_doc,doc_no,stage from file6_90h where doc_no = " & MyParn(sDoc_no))
            If locTable Is Nothing Then GoTo CleanUp
            If Not locTable.EOF Then
                grid1.TextMatrix(grid1.Rows - 1, 19) = locTable!sales_Doc & ""
                grid1.TextMatrix(grid1.Rows - 1, 20) = "1"
                grid1.TextMatrix(grid1.Rows - 1, 24) = locTable!Stage
                If grid1.TextMatrix(grid1.Rows - 1, 24) <> "" Then
                    stageTable.Find "CODE = " & grid1.TextMatrix(grid1.Rows - 1, 24), , adSearchForward, adBookmarkFirst
                    If Not stageTable.EOF Then
                        grid1.TextMatrix(grid1.Rows - 1, 25) = stageTable!desca & ""
                    End If
                End If
            End If
            
            grid1.TextMatrix(grid1.Rows - 1, 21) = cSv.GetCellByName(i, "Tags")
            If Not db.rsValue("select  [dbo].[fn_stage](" & _
                                    addstring(grid1.TextMatrix(grid1.Rows - 1, 15)) & "," & _
                                    addstring(grid1.TextMatrix(grid1.Rows - 1, 21)) & ")", _
                                    nStage) Then
                GoTo CleanUp
            End If
            
            grid1.TextMatrix(grid1.Rows - 1, 22) = nStage
            stageTable.Find "CODE = " & Val(nStage), , adSearchForward, adBookmarkFirst
            If Not stageTable.EOF Then
                grid1.TextMatrix(grid1.Rows - 1, 23) = stageTable!desca & ""
            End If
            
            grid1.TextMatrix(grid1.Rows - 1, grid1.Cols - 1) = "0"
            grid1.ShowCell grid1.Rows - 1, 0
            DoEvents
        End If
    End If
Next


Me.Caption = sCaption

Dim SKU As String
Dim sitem As String

Tb.Clear
sDoc_no = ""
For i = 0 To cSv.NumRows - 1
    Me.Caption = sCaption & " - " & "”Ã· " & (i + 1) & " „‰ " & cSv.NumRows
    prog1.Value = Round(i / (cSv.NumRows), 2) * 100
    SKU = cSv.GetCellByName(i, "Lineitem sku")
    
    If Trim(Replace(cSv.GetCellByName(i, "Name"), "#", "")) <> "" Then
        sDoc_no = Trim(Replace(cSv.GetCellByName(i, "Name"), "#", ""))
        sitem = retItemCode(SKU, con) & ""
        sitem = IIf(IsNumeric(sitem), sitem, "")
        
        GRID2.AddItem ""
        GRID2.TextMatrix(GRID2.Rows - 1, 0) = sDoc_no
        GRID2.TextMatrix(GRID2.Rows - 1, 1) = SKU
        GRID2.TextMatrix(GRID2.Rows - 1, 2) = sitem
        GRID2.TextMatrix(GRID2.Rows - 1, 3) = cSv.GetCellByName(i, "Lineitem name")
        GRID2.TextMatrix(GRID2.Rows - 1, 4) = Val(cSv.GetCellByName(i, "Lineitem price"))
        GRID2.TextMatrix(GRID2.Rows - 1, 5) = cSv.GetCellByName(i, "Lineitem quantity")
                         
        Set locTable = db.myRs("select sales_doc,doc_no from file6_90h where doc_no = " & MyParn(sDoc_no))
        
        If locTable Is Nothing Then GoTo CleanUp
        
        If Not locTable.EOF Then
            GRID2.TextMatrix(GRID2.Rows - 1, 6) = locTable!sales_Doc & ""
            GRID2.TextMatrix(GRID2.Rows - 1, 7) = "1"
        End If
        GRID2.ShowCell GRID2.Rows - 1, 0
        DoEvents
    End If
Next
CleanUp:
Set db = Nothing
Me.Caption = sCaption
prog1.Visible = False
getData2 = True
Set db = Nothing
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
Resume CleanUp
End Function
