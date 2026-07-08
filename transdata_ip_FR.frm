VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form transdata_ip_FR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕœÌÀ Ê ”Õ» »Ì«‰«  «· ÊﬂÌ·« "
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "transdata_ip_FR.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   8235
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox FIXDATA 
      Alignment       =   1  'Right Justify
      Caption         =   "FIX DATA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6255
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   6840
      Width           =   1050
   End
   Begin VB.CheckBox xhead 
      Alignment       =   1  'Right Justify
      Caption         =   "head"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   5715
      Width           =   915
   End
   Begin VB.CheckBox XDISC 
      Alignment       =   1  'Right Justify
      Caption         =   "XDISC"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   7335
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   6840
      Width           =   915
   End
   Begin VB.CheckBox XGRDISC 
      Alignment       =   1  'Right Justify
      Caption         =   "„Ã„Ê⁄«  «·Œ’„"
      Height          =   330
      Left            =   450
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6075
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   6345
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   4590
      Width           =   2175
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   "ISSTOP"
         Height          =   270
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1845
         Width           =   1365
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ· «·«’‰«›"
         Height          =   285
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   945
         Width           =   1770
      End
      Begin VB.CommandButton CMD_SENPRICE 
         BackColor       =   &H00FDD0CE&
         Caption         =   " ÕœÌÀ «”⁄«— & BARCODE13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
         Width           =   2040
      End
      Begin VB.CheckBox Check4 
         Alignment       =   1  'Right Justify
         Caption         =   "FIX  COST SALES"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1215
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox XBAR13 
         Alignment       =   1  'Right Justify
         Caption         =   "  ALL BARCODE13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1530
         Width           =   1815
      End
   End
   Begin VB.TextBox xdatesales 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   315
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5310
      Width           =   1770
   End
   Begin VB.Frame Frame2 
      Caption         =   "FIX  IP"
      Height          =   870
      Left            =   585
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   7155
      Width           =   7710
      Begin VB.TextBox XIP 
         Alignment       =   2  'Center
         Height          =   465
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   270
         Width           =   4020
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SND IP"
         Height          =   495
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   1635
      End
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   " ÕœÌÀ «· ﬂÊÌœ"
      Height          =   465
      Left            =   630
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   4680
      Width           =   1410
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   "”Õ» «·»Ì«‰«  „‰ «· ÊﬂÌ·« "
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2205
      TabIndex        =   5
      Top             =   5265
      Width           =   4110
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2205
      TabIndex        =   3
      Top             =   5895
      Width           =   4110
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "÷»ÿ „·› »Ì«‰«  ·· ÊﬂÌ·« "
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2205
      TabIndex        =   0
      Top             =   4635
      Width           =   4110
   End
   Begin VB.Frame Frame11 
      Height          =   600
      Left            =   585
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   6480
      Width           =   5730
      Begin MSComctlLib.ProgressBar prog1 
         Height          =   375
         Left            =   45
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   4515
      Left            =   90
      TabIndex        =   4
      Top             =   45
      Width           =   6765
      _cx             =   11933
      _cy             =   7964
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
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
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
   Begin VB.Label xcost 
      Alignment       =   1  'Right Justify
      Height          =   4425
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   90
      Width           =   1545
   End
End
Attribute VB_Name = "transdata_ip_FR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lAddAll As Boolean
Dim con As New ADODB.Connection, pBranch As String
Dim conShop2 As New ADODB.Connection
Dim cFilesave As String
Public strConShop2 As String
Dim cDataFolder As String, cDataFile As String
Dim dPostDate As String
Dim nBranchPost As Double
Private Function ToBranch() As Boolean
'On Error GoTo myerror
Dim lMyPrice As Boolean, lMyPriceItem As Boolean

FixItemDel

'conShop2.Execute " delete from model_dem where year(date) = 2018 "
'conShop2.Execute " delete from model_dem where month(date) = 1 "
'conShop2.Execute " delete from model_dem where month(date) = 2 "

Me.Caption = " «·«’‰«› "

Me.Caption = " «·«’‰«› send1 "
nRecordCount = sendItems
Me.Caption = " «·«’‰«› send2 "

If xhead.Value <> 0 Then
'    cId_No1 = GetDesca("SELECT ID_NO1 FROM ADDRESS", con)
'    cId_No2 = GetDesca("SELECT ID_NO2 FROM ADDRESS", con)
    cHead1 = GetDesca("SELECT HEAD1 FROM ADDRESS", con)
    conShop2.Execute " UPDATE ADDRESS SET HEAD1 = " & addstring(cHead1)
    ToBranch = True
    Exit Function
End If


'Inform " «·«’‰«› "

If XGRDISC.Value <> 0 Then nRecordCount = SendGrDiscout

lMyPrice = (GetDesca("SELECT PRICE FROM FILE0_00 ", con) = "True")
lMyPriceItem = (GetDesca("SELECT OKAZ_ITEM2 FROM FILE0_00 ", con) = "True")

If lMyPrice Then
    conShop2.Execute " update FILE0_00 SET PRICE = 1 "
Else
    conShop2.Execute " update FILE0_00 SET PRICE = 0 "
End If

If lMyPriceItem Then
    conShop2.Execute " update FILE0_00 SET OKAZ_ITEM2 = 1 "
Else
    conShop2.Execute " update FILE0_00 SET OKAZ_ITEM2 = 0 "
End If



Me.Caption = " „‘ —Ì«  1"
If lAddAll Then
    nRecordCount = SendPurchaseHeader("SELECT FR7_20H.* FROM FR7_20H WHERE STORE = " & MyParn(pBranch), "FR7_20h", "FR7_20", "FILE7_20H", "FILE7_20")
Else
    nRecordCount = SendPurchaseHeader("SELECT * from FR7_20H WHERE ISNEW1 = 1 AND STORE =  " & MyParn(pBranch), "FR7_20h", "FR7_20", "FILE7_20H", "FILE7_20")
End If
Me.Caption = " „‘ —Ì«  2"

Me.Caption = " „—œÊœ „‘ —Ì«  1"
If lAddAll Then
    nRecordCount = SendPurchaseHeader("SELECT FR7_10H.* FROM FR7_10H WHERE STORE = " & MyParn(pBranch), "FR7_10h", "FR7_10", "FILE7_10H", "FILE7_10")
Else
    nRecordCount = SendPurchaseHeader("SELECT FR7_10H.* FROM FR7_10H WHERE ISNEW1 = 1 AND STORE =  " & MyParn(pBranch), "FR7_10h", "FR7_10", "FILE7_10H", "FILE7_10")
End If
Me.Caption = " „—œÊœ „‘ —Ì«  2"

    Me.Caption = " ÕÊÌ·« 1"
    nRecordCount = sendTransHeader
    
    Me.Caption = "«· ÕÊÌ·«  «·Ê«—œ…"
    nRecordCount = sendTransHeader2

nRecordCount = SendDiscount
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ «·Œ’Ê„«  ", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = SendSection
sendSUB_USER

nRecordCount = NewPrice
senTarget
If Check3.Value <> 0 Then
    nRecordCount = SendGroup
    nRecordCount = Sendfact
    nRecordCount = sendCode("FILE4_50")
    nRecordCount = sendCode("FILE4_10")
    conShop2.Execute " update file4_10 set [group] = '01'"
    nRecordCount = sendmosm
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„Ê«”„ ", "»‰Ã«Õ" Else GoTo myerror
    nRecordCount = sendCode("FILE0_81", True)
    nRecordCount = sendCode("FILE0_82", True)
    nRecordCount = sendCode("SHIP")
    If pBranch <> "65" Then
        conShop2.CommandTimeout = 1000
        conShop2.Execute " DELETE FROM MOSM      WHERE MOSM NOT IN (SELECT MOSM FROM FILE1_10   )"
'       conShop2.Execute " DELETE FROM FACT      WHERE CODE NOT IN (SELECT FACT FROM FILE1_10   )"
'       conShop2.Execute " DELETE FROM FILE4_10  WHERE CODE NOT IN (SELECT CODE FROM file7_20h  )"
    End If
End If

' nRecordCount = sendItems2

ToBranch = True
'MsgBox " „ ‰ﬁ· «·»Ì«‰« "
ToBranch = True
Exit Function
ToBranch = False
myerror:
'MsgBox "·„ Ì „ ‰ﬁ· ﬂ«›… «·»Ì«‰«  »‰Ã«Õ"
Err.Clear
End Function
Private Sub CMD_SENPRICE_Click()
    Me.Caption = " ÕœÌÀ  ·ﬁ«∆Ï ··»Ì«‰«  „‰ Ê ≈·Ï «·›—Ê⁄"
    With grid1
    For nRow = 1 To .Rows - 1
        If .TextMatrix(nRow, 2) <> "" Then
            If .TextMatrix(nRow, 2) <> "" And TurnValue(.TextMatrix(nRow, 4), "True", True) Then
                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbYellow
                IpShop = .TextMatrix(nRow, 2)
                cDataShop = .TextMatrix(nRow, 3)
                lAddAll = TurnValue(.TextMatrix(nRow, 5), "True", True)
                If lAddAll Then
                    If MsgBox("Â–« «·«Œ Ì«— ”Ê› Ì „ ‰ﬁ· ﬂ· «·»Ì«‰«  ≈·Ï «· ÊﬂÌ·", vbYesNo + vbDefaultButton2) = vbNo Then
                        lAddAll = False
                    End If
                End If
                
                strConShop2 = LoadConStringshop2(IpShop, cDataShop)
                nBranchPost = Val(.TextMatrix(nRow, 0))
                pBranch = .TextMatrix(nRow, 0)
                
                If openConShop(conShop2, strConShop2) = "ok" Then
                    Inform "  „ «·≈ ’«· »«· ÊﬂÌ·  "
                    SubFixData
                    If Not ToBranch_COST() Then
                        .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
                    Else
                        .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbGreen
                    End If
                Else
                    .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
                End If
            End If
            If conShop2.State = adStateOpen Then closeCon conShop
        End If
    Next nRow
    End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
'   On Error GoTo myerror
    Me.Caption = " ÕœÌÀ  ·ﬁ«∆Ï ··»Ì«‰«  „‰ Ê ≈·Ï «·›—Ê⁄"
    With grid1
    For nRow = 1 To .Rows - 1
        If .TextMatrix(nRow, 2) <> "" Then
            If .TextMatrix(nRow, 2) <> "" And TurnValue(.TextMatrix(nRow, 4), "True", True) Then
                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbYellow
                IpShop = .TextMatrix(nRow, 2)
                cDataShop = .TextMatrix(nRow, 3)
                lAddAll = TurnValue(.TextMatrix(nRow, 5), "True", True)
                If lAddAll Then
                    If MsgBox("Â–« «·«Œ Ì«— ”Ê› Ì „ ‰ﬁ· ﬂ· «·»Ì«‰«  ≈·Ï «· ÊﬂÌ·", vbYesNo + vbDefaultButton2) = vbNo Then
                        lAddAll = False
                    End If
                End If
                strConShop2 = LoadConStringshop2(IpShop, cDataShop)
                pBranch = .TextMatrix(nRow, 0)
                nBranchPost = Val(pBranch)
                Inform " branch " & Val(pBranch)
                If openConShop(conShop2, strConShop2) = "ok" Then
                    SubFixData
                    If FIXDATA.Value = 0 Then
                        Inform "  „ «·≈ ’«· »«· ÊﬂÌ·  "
                        If Not ToBranch() Then
                            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
                        Else
                            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbGreen
                        End If
                    End If
                Else
                    .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
                End If
            End If
            If conShop2.State = adStateOpen Then closeCon conShop
        End If
    Next nRow
    End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Command1_Click()
    ToMain
End Sub
Private Sub Form_Load()
    cDataFolder = App.Path & "\mdb"
    cDataFile = "data"
    XDISC.Visible = bSupermode
    
    CMD_SENPRICE.Visible = bopt2
    openCon con
    myload
    dPostDate = DateAdd("M", -1, Date)
End Sub
Private Function SendGroup() As Long
'On Error GoTo myerror

Dim cFile As String, nRecord As Double
Dim loctable As New ADODB.Recordset
Dim cStr1  As String, cStr2   As String
cFile = "FILE1_50"
cString = "SELECT * FROM " & cFile
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
            
    cStr1 = cStr1 & addInsert(aInsert, cFile) & " ; "
    loctable.MoveNext
    SendGroup = SendGroup + nAffect
Loop
conShop2.Execute " DELETE FROM FILE1_50 "
conShop2.Execute cStr1
Inform " „Ã„Ê⁄«  " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
SendGroup = -1
Err.Clear
GoTo lastsub
End Function
Private Function sendItems() As Long
'On Error GoTo myerror
Dim cFile As String, nCount As Double
Dim cStrPostAll As String, cStrPostAll_DEL As String, cStrDel1 As String
cFile = "FILE1_10"

If Check2.Value <> 0 Then
    cString = "SELECT * FROM FILE1_10 WHERE  ITEM IN (SELECT ITEM FROM FR1_11 WHERE  STORE = " & MyParn(pBranch) & ")  "
Else
    cString = "SELECT * FROM FILE1_10 WHERE  FILE1_10.ITEM IN (SELECT ITEMNEW FROM BARCODE_EDIT) OR ITEM IN (SELECT ITEM FROM FR1_11 WHERE STORE = " & MyParn(pBranch) & ") AND ITEM IN ( SELECT ITEM FROM NEW_ITEM_FR WHERE BRANCH = " & MyParn(pBranch) & " ) "
End If

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
sCaption = Me.Caption
Do Until loctable.EOF
    nCount = nCount + 1
    Me.Caption = loctable!Item
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!MODEL))
    aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable!MODELNO))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "FACT", addstring(loctable!Fact))
    aInsert = AddFlag(aInsert, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "MODELFACT", addstring(loctable!modelfact))
    aInsert = AddFlag(aInsert, "MODELFACT0", addstring(loctable!modelfact0))
    aInsert = AddFlag(aInsert, "SUPP", addstring(loctable!SUPP))
    aInsert = AddFlag(aInsert, "OKAZ", Val(loctable!okaz & ""))
    aInsert = AddFlag(aInsert, "OKAZ_2", Val(loctable!okaz_2 & ""))
    aInsert = AddFlag(aInsert, "RATE", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "[SECTION]", addvalue(loctable!Section))
    aInsert = AddFlag(aInsert, "MODELSEX", addvalue(loctable!MODELSEX))
    aInsert = AddFlag(aInsert, "MODELAGE", addvalue(loctable!MODELAGE))
    aInsert = AddFlag(aInsert, "[ITEM]", addvalue(loctable!Item))
    aInsert = AddFlag(aInsert, "[COSTITEM]", Val(loctable!costITEM & ""))
    aInsert = AddFlag(aInsert, "[COST2]", Val(loctable!cost2 & ""))
    aInsert = AddFlag(aInsert, "[PRICE]", Val(loctable!price & ""))
    aInsert = AddFlag(aInsert, "[PRICE_2]", Val(loctable!PRICE_2 & ""))
    aInsert = AddFlag(aInsert, "[SCAL]", addstring(loctable!scal))
    aInsert = AddFlag(aInsert, "[C_SCAL]", addvalue(loctable!C_SCAL))
    aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable!color))
    aInsert = AddFlag(aInsert, "[C_COLOR]", addvalue(loctable!c_Color))
    aInsert = AddFlag(aInsert, "[BARCODE]", addstring(loctable!BARCODE))
    aInsert = AddFlag(aInsert, "[BARCODE2]", addstring(loctable!BARCODE2))
    aInsert = AddFlag(aInsert, "[BARCODE13]", addstring(loctable!BARCODE13))
    aInsert = AddFlag(aInsert, "[BARCODE_GS1]", addstring(loctable!BARCODE_GS1))
    aInsert = AddFlag(aInsert, "[ISOKAZITEM]", IIf(loctable!ISOKAZITEM, 1, 0))
    aInsert = AddFlag(aInsert, "[ISNOITEM]", IIf(loctable!ISNOITEM, 1, 0))
    aInsert = AddFlag(aInsert, "[ISNOITEM2]", IIf(loctable!ISNOITEM2, 1, 0))
    aInsert = AddFlag(aInsert, "[ISNODEM]", IIf(loctable!ISNODEM, 1, 0))
    aInsert = AddFlag(aInsert, "[MOSM2]", addstring(loctable!MOSM2))
    aInsert = AddFlag(aInsert, "[PRICE2]", Val(loctable!PRICE2 & ""))
    aInsert = AddFlag(aInsert, "[COST]", Val(loctable!cost & ""))
    aInsert = AddFlag(aInsert, "[ISSTOP]", IIf(loctable!ISSTOP, 1, 0))
    
    cStrPostAll = cStrPostAll & " DELETE FROM FILE1_10 where item = " & loctable!Item & " ; "
    cStrDel1 = cStrDel1 & " DELETE FROM NEW_ITEM_FR WHERE ITEM = " & loctable!Item & " AND BRANCH = " & MyParn(pBranch) & " ; "
    If loctable!Item = 11759 Then
        A = A
    End If
    
    cStrPostAll = cStrPostAll & addInsert(aInsert, "FILE1_10") & " ; "
''''''''''''''''''''
''''''''''''''''''''
''''''''''''''''''''
''''''''''''''''''''
    If nCount = 100 Then
'    If nCount = 1 Then
        If cStrPostAll <> "" Then
            conShop2.Execute cStrPostAll, nRec
con.Execute cStrDel1
            cStrDel1 = ""
            cStrPostAll = ""
        End If
        nCount = 0
    End If
    Me.Caption = loctable!Item
    loctable.MoveNext
    sendItems = sendItems + 1
Loop


If cStrPostAll <> "" Then
    conShop2.Execute cStrPostAll, nRec
    con.Execute cStrDel1
End If
Inform "  „ ‰ﬁ· «·«’‰«› " & nRecord
lastsub:
Me.Caption = sCaption
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendItems = -1
GoTo lastsub
End Function
Private Function sendCode(cFile, Optional isNumber As Boolean = False) As Long
On Error GoTo myerror
Dim cStrPostAll As String
cString = "SELECT * FROM " & cFile
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100

    If isNumber Then
        aInsert = AddFlag(Empty, "CODE", addvalue(loctable!CODE))
    Else
        aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    End If
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    If cStrPostAll = "" Then
        cStrPostAll = addInsert(aInsert, cFile)
    Else
        cStrPostAll = cStrPostAll & " ; " & addInsert(aInsert, cFile)
    End If
    loctable.MoveNext
    sendCode = sendCode + 1
Loop
conShop2.Execute " delete from " & cFile
conShop2.Execute cStrPostAll
Inform cFile & " : " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendCode = -1
GoTo lastsub
End Function
Private Function SendPurchaseHeader(cString, cFile, cFile2, cFile_FR, cFile2_FR) As Long
'On Error GoTo myerror
Dim loctable As New ADODB.Recordset, cDoc_No As String
Dim loctable2 As New ADODB.Recordset, cStr1 As String, cStr2 As String, cStrDel1 As String, cStrDel2 As String, cNew1 As String, cNew2 As String
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    cDoc_No = (loctable!doc_no)
    
    cStrDel2 = " delete from " & cFile2_FR & " where doc_no = " & MyParn(cDoc_No)
    cStrDel1 = " delete from " & cFile_FR & " where doc_no = " & MyParn(cDoc_No)
    conShop2.Execute cStrDel2
    conShop2.Execute cStrDel1
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!STORE = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(cDoc_No))
        aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
        aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
        aInsert = AddFlag(aInsert, "[TAX]", Val(loctable!tax & ""))
        aInsert = AddFlag(aInsert, "[r_tax]", Val(loctable!R_TAX & ""))
        aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM))
        aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))

 '      cStr1 = addInsert(aInsert, cFile)
        conShop2.Execute addInsert(aInsert, cFile_FR), nAffect
        
        cStr2 = ""
        Dim aInsert2 As Variant
        If loctable2.State = adStateOpen Then loctable2.Close
        loctable2.Open "SELECT * FROM " & cFile2 & " WHERE DOC_NO = " & MyParn(loctable!doc_no), con, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not loctable2.EOF
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(cDoc_No))
            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
            aInsert2 = AddFlag(aInsert2, "DISCOUNT", Val(loctable2!discount & ""))
'           aInsert2 = AddFlag(aInsert2, "cost_inv", Val(loctable2!cost_inv & ""))
            aInsert2 = AddFlag(aInsert2, "r_tax", Val(loctable2!R_TAX & ""))
            cStr2 = cStr2 & addInsert(aInsert2, cFile2_FR) & " ; "
            loctable2.MoveNext
        Loop
        If cStr2 <> "" Then conShop2.Execute cStr2
    End If
    cNew1 = cNew1 & " UPDATE " & cFile & "  SET ISNEW1 = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no) & " ; "
    loctable.MoveNext
    SendPurchaseHeader = SendPurchaseHeader + nAffect
Loop
If cNew1 <> "" Then con.Execute cNew1
Inform cFile2_FR & " : " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendPurchaseHeader = -1
GoTo lastsub
End Function
Private Function sendmosm() As Long
'On Error GoTo myerror
cFile = "mosm"
cString = "SELECT * FROM MOSM "

Dim loctable As New ADODB.Recordset, cStr1 As String
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop2.Execute " delete from mosm "
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "MOSM", addstring(loctable!MOSM))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
    cStr1 = cStr1 & addInsert(aInsert, cFile) & " ; "
    loctable.MoveNext
    sendmosm = sendmosm + nAffect
Loop
conShop2.Execute " delete from mosm "
conShop2.Execute cStr1
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendmosm = -1
GoTo lastsub
End Function
Private Function Sendfact() As Long
On Error GoTo myerror
Dim cFile As String
cFile = "fact"
cString = "SELECT FACT.CODE , FACT.DESCA FROM FACT "
Dim loctable As New ADODB.Recordset, cStr1 As String
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    cStr1 = cStr1 & addInsert(aInsert, cFile) & " ; "
    loctable.MoveNext
Loop
conShop2.Execute " delete from fact "
conShop2.Execute cStr1
Inform " „’«‰⁄ " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Sendfact = -1
Err.Clear
GoTo lastsub
End Function
Private Sub myload()
openCon con
Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT * FROM branch_FR  WHERE ISSTOP =0 AND ipsql IS NOT NULL "
cString = cString & " ORDER BY CODE"
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With grid1
Do Until loctable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = loctable(0) & ""
    .TextMatrix(.Rows - 1, 1) = loctable!DESCA & ""
    If (loctable!IPsql & "") = "" Then
        .TextMatrix(.Rows - 1, 4) = False
    Else
        IpShop = loctable!IPsql & ""
        cDataShop = loctable!data & ""
        strConShop = LoadConStringshop
        .TextMatrix(.Rows - 1, 2) = loctable!IPsql & ""
        .TextMatrix(.Rows - 1, 3) = loctable!data & ""
        .TextMatrix(.Rows - 1, 4) = False
        .TextMatrix(.Rows - 1, 5) = False
    End If
    loctable.MoveNext
Loop
Fixgrd
End With
End Sub
Private Function ToMain() As Boolean
'On Error GoTo myerror
With grid1
For nRow = 1 To .Rows - 1
    If .TextMatrix(nRow, 2) <> "" And TurnValue(.TextMatrix(nRow, 4), "True", True) Then
        IpShop = .TextMatrix(nRow, 2)
        cDataShop = .TextMatrix(nRow, 3)
        strConShop2 = LoadConStringshop2(IpShop, cDataShop)
        openConShop conShop2, strConShop2
        .Cell(flexcpBackColor, nRow, 1, nRow, .Cols - 1) = &HE0E0E0
        XCOST.Caption = ""

        If Not getData(.TextMatrix(nRow, 0)) Then
            Inform "„‘ﬂ·… «À‰«¡ ”Õ» »Ì«‰«   ÊﬂÌ· : " & .TextMatrix(nRow, 1)
            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
'            Exit Function
        Else
            Inform " „ ”Õ» »Ì«‰«   ÊﬂÌ· : " & .TextMatrix(nRow, 1), " »‰Ã«Õ", 1200
            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbGreen
        End If
    End If
Next
End With
ToMain = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function getData(sBranch As String) As Boolean
Dim nRecordCount As Long
    
    If XDISC.Value <> 0 Then
        FixDiscDox (sBranch)
        getData = True
        Exit Function
    End If
    
    
    Me.Caption = "„»Ì⁄« "
    nRecordCount = GetSales(sBranch)
    If nRecordCount >= 0 Then
        Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ Õ—ﬂ«  «·»Ì⁄", "»‰Ã«Õ"
    Else
        GoTo myerror
    End If

    Me.Caption = "«·Ã—œ"
    nRecordCount = getStock(sBranch)
    If nRecordCount >= 0 Then
        Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ „” ‰œ«  «·Ã—œ ", "»‰Ã«Õ"
    Else
        GoTo myerror
    End If


getData = True
Exit Function
myerror:
End Function
Private Function GetSales(sBranch As String) As Long
cFile = "FILE6_20"
Set loctable = New ADODB.Recordset
Dim cStr1 As String, cStr2 As String, cDelStr1 As String, cDelStr2 As String, cStrNew As String, cStrPhone As String

If IsDate(xdatesales.text) Then
    conShop2.Execute " UPDATE FILE6_20H SET ISNEW = 1 WHERE DATE >= " & DateSq(xdatesales.text)
    Exit Function
End If

cString = "Select * from " & cFile & "H WHERE (PRINTED = 1 and isnew = 1 )  "
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long, lSkip As Boolean
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Dim aInsert2 As Variant
Inform " „»Ì⁄«  "
Do Until loctable.EOF
    con.Execute " delete from FR6_20 where doc_no = " & MyParn(loctable!doc_no)
    con.Execute " delete from FR6_20H where doc_no = " & MyParn(loctable!doc_no)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[TIME]", addTime(Format(loctable!Time, "SHORT TIME")))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "[NOTES]", addstring(loctable!NOTES))
    aInsert = AddFlag(aInsert, "[RATE]", Val(loctable!Rate & ""))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[CASH]", Val(loctable!CASH & ""))
    aInsert = AddFlag(aInsert, "[BOX]", addstring(loctable!BOX))
    aInsert = AddFlag(aInsert, "[PAY]", Val(loctable!PAY & ""))
    aInsert = AddFlag(aInsert, "[VISA]", Val(loctable!Visa & ""))
    aInsert = AddFlag(aInsert, "[LATE]", Val(loctable!late & ""))
    aInsert = AddFlag(aInsert, "[REST]", Val(loctable!rest & ""))
    aInsert = AddFlag(aInsert, "[MAN]", addstring(loctable!MAN))
    aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
    aInsert = AddFlag(aInsert, "[PRINTED]", IIf(loctable!PRINTED, "1", "0"))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "[phone]", addstring(loctable!phone))
                                                                                
    aInsert = AddFlag(aInsert, "[ISRET]", IIf(loctable!ISRET, 1, 0))
                                                                                
    aInsert = AddFlag(aInsert, "[username_disc]", addstring(loctable!username_disc))
    aInsert = AddFlag(aInsert, "[username_ret]", addstring(loctable!username_ret))
    aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(loctable!SALES_RET))
    aInsert = AddFlag(aInsert, "[NOTE_DISC]", addstring(loctable!NOTE_DISC))
    aInsert = AddFlag(aInsert, "[DOC_DISC]", addstring(loctable!DOC_DISC))
                                                                                
    aInsert = AddFlag(aInsert, "CHARGE1", Val(loctable!CHARGE1 & ""))
    aInsert = AddFlag(aInsert, "CHARGE2", Val(loctable!charge2 & ""))
    aInsert = AddFlag(aInsert, "[NAME]", addstring(loctable!Name))
    aInsert = AddFlag(aInsert, "[SHIP]", addstring(loctable!SHIP))
    aInsert = AddFlag(aInsert, "[ADDRESS]", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "[SHIP_NO]", addstring(loctable!SHIP_NO))
    aInsert = AddFlag(aInsert, "[DATE_PAY]", addDate(loctable!DATE_PAY))
    aInsert = AddFlag(aInsert, "[BANK]", addstring(loctable!BANK))
    aInsert = AddFlag(aInsert, "[ONLINE]", addvalue(loctable!online))
    aInsert = AddFlag(aInsert, "[note_ret]", addstring(loctable!note_ret))
    aInsert = AddFlag(aInsert, "[card_disc]", addstring(loctable!BANK))
    aInsert = AddFlag(aInsert, "[card_desca]", addstring(loctable!card_desca))
    aInsert = AddFlag(aInsert, "[card_VALUE]", addvalue(loctable!card_VALUE))
    aInsert = AddFlag(aInsert, "[card_DOC]", addstring(loctable!card_DOC))
    
    aInsert = AddFlag(aInsert, "[VISA3]", Val(loctable!VISA3 & ""))
    aInsert = AddFlag(aInsert, "[CODEVISA3]", addstring(loctable!CODEVISA3))
    aInsert = AddFlag(aInsert, "[VISADOC3]", addstring(loctable!VISADOC3))
    aInsert = AddFlag(aInsert, "[VISA3_2]", Val(loctable!VISA3_2))
    aInsert = AddFlag(aInsert, "[TYPEVISA3 ]", addstring(loctable!TYPEVISA3))

    
    con.Execute addInsert(aInsert, "FR6_20H")
    
    If Not IsNull(loctable!card_DOC) Then
        con.Execute " UPDATE CARD_DISC SET SALES_DOC_NO = " & addstring(loctable!doc_no) & " WHERE BARCODE_DISC = " & MyParn(loctable!card_DOC)
    End If
    
    lSkip = False
    
    Set loctable2 = New ADODB.Recordset
    cString = "Select FILE6_20.* from FILE6_20 WHERE FILE6_20.doc_no = " & MyParn(loctable!doc_no)
    loctable2.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText
    Dim sitem As String
    Dim aitem As Variant
    cStr2 = ""
    Do Until loctable2.EOF
        sitem = loctable2!Item
        If GetDesca("select item from file1_10 where  item = " & sitem, con) <> "" Then
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "MAN", addstring(loctable2!MAN))
            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
            aInsert2 = AddFlag(aInsert2, "DISCOUNT", Val(loctable2!discount & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE_C", Val(loctable2!PRICE_C & ""))
            aInsert2 = AddFlag(aInsert2, "ROW", Val(loctable2!Row & ""))
            aInsert2 = AddFlag(aInsert2, "COST", Val(loctable2!cost & ""))
            aInsert2 = AddFlag(aInsert2, "S_OKAZ", Val(loctable2!S_OKAZ & ""))
            aInsert2 = AddFlag(aInsert2, "ISDISC2_1", IIf(loctable2!ISDISC2_1, 1, 0))
            cStr2 = cStr2 & addInsert(aInsert2, "FR6_20") & " ; "
        Else
            Inform " „—«Ã⁄… „»Ì⁄«  »«—ﬂÊœ " & sitem
            XCOST.Caption = XCOST.Caption & Chr(13) & sBranch & " - " & sitem
            lSkip = True
        End If
        loctable2.MoveNext
    Loop
    If cStr2 <> "" Then con.Execute cStr2
    If Not lSkip Then
        conShop2.Execute "  update file6_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no)
    End If
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop

Inform " „»Ì⁄«  " & nRecord
GetSales = nAffectTotal

Inform "”Õ» «· Ì·›Ê‰« "
Set loctable = New ADODB.Recordset
cString = "Select * from SUBCUST WHERE isnew = 1 "
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

nRecordCount = loctable.RecordCount
nRecord = 0

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "PHONE", addstring(loctable!phone))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "E_MAIL", addstring(loctable!E_MAIL))
    aInsert = AddFlag(aInsert, "branch", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "F_DATE", addDate(loctable!F_DATE))
    cStrPhone = cStrPhone & addInsert(aInsert, "SUBCUST") & " ; "
    loctable.MoveNext
Loop
If cStrPhone <> "" Then
    con.Execute cStrPhone
    conShop2.Execute " UPDATE SUBCUST SET isnew = 0 "
End If
Inform " phone : " & nRecord
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GetSales = -1
GoTo lastsub
End Function
Private Sub Fixgrd()
With grid1
.Cols = 6
.ColWidth(0) = 1000
.ColWidth(1) = 1500
.ColWidth(2) = 2000
.ColWidth(3) = 0
.ColWidth(4) = 600
.ColWidth(5) = 600
.TextMatrix(0, 0) = "ﬂÊœ «· ÊﬂÌ·"
.TextMatrix(0, 1) = "≈”„ «· ÊﬂÌ·"
.TextMatrix(0, 2) = "IP SQL"
.TextMatrix(0, 4) = "„ ’·"
.TextMatrix(0, 5) = "«·ﬂ·"

.ColDataType(4) = flexDTBoolean
.ColDataType(5) = flexDTBoolean

.ColAlignment(0) = flexAlignRightCenter
.ColAlignment(1) = flexAlignRightCenter

End With
End Sub
Private Function validData(sBranch As String) As Boolean
Dim fs As New FileSystemObject
If Not fs.FileExists(cDataFolder & "\" & cDataFile & "_" & sBranch & ".mdb") Then Exit Function
validData = True
End Function

Private Sub grid1_EnterCell()
With grid1
    If .Col >= 4 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Function getStock(pBranch) As Long
cFile = "FILE0_10H"
Dim loctable2 As New ADODB.Recordset, nCountTot As Double
Dim cStr1 As String, cStr2 As String, cStr3 As String, cStrDel1 As String, cStrDel2 As String, cStrDel3 As String, cNew As String
cString = "SELECT  FILE0_10H.* FROM  file0_10h where CLOSED =  1  AND isnew  = 1 AND BRANCH = " & MyParn(pBranch)
'cString = "SELECT  FILE0_10H.* FROM  file0_10h "

Dim loctable As New ADODB.Recordset, nCountS As String
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
XCOST.Caption = "Ã—œ" & Chr(13)
Do Until loctable.EOF
    cStrDel1 = ""
    cStrDel2 = ""
    cStrDel3 = ""
    cStr1 = ""
    cStr2 = ""
    cStr3 = ""
    
    cStrDel1 = " delete from FR0_10 where doc_no = " & MyParn(loctable!doc_no)
    cStrDel2 = " delete from FR0_100 where doc_no = " & MyParn(loctable!doc_no)
    cStrDel3 = " delete from FR0_10h where doc_no = " & MyParn(loctable!doc_no)
        
    con.Execute cStrDel1
    con.Execute cStrDel2
    con.Execute cStrDel3
    
    nRecord = nRecord + 1
    
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
    
    
    con.Execute addInsert(aInsert, "FR0_10H")
    nCountS = 0
    cStr2 = ""
    cStr3 = ""
    
    Set loctable2 = Nothing
    Set loctable2 = New ADODB.Recordset
    'If loctable2.State = adStateOpen Then loctable2.Close
    Set loctable2 = myCmd("SELECT * FROM FILE0_10 WHERE DOC_NO = " & MyParn(loctable!doc_no), conShop2)
    'loctable2.Open "SELECT * FROM FILE0_10 WHERE DOC_NO = " & MyParn(loctable!doc_no), conShop2, adOpenStatic, adLockReadOnly, adCmdText
    Do Until loctable2.EOF
        nCountTot = nCountTot + 1
        Me.Caption = loctable2!Item
        If GetDesca("select item from file1_10 where item = " & loctable2!Item, con) <> "" Then
            prog1.Value = 0
            prog1.Max = 1000
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "[ITEM]", addstring(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "[RealBal]", Val(loctable2!RealBal & ""))
            aInsert2 = AddFlag(aInsert2, "[ComputerBal]", Val(loctable2!ComputerBal & ""))
            aInsert2 = AddFlag(aInsert2, "[Differ]", Val(loctable2!Differ & ""))
            aInsert2 = AddFlag(aInsert2, "[cost]", Val(loctable2!cost & ""))
            
            cStr2 = cStr2 & addInsert(aInsert2, "FR0_10") & " ; "
            cStr3 = cStr3 & addInsert(aInsert2, "FR0_100") & " ; "
            Me.Caption = loctable2!doc_no & " " & loctable2!Item
            nCountS = nCountS + 1
            prog1.Value = nCountS
            If nCountS = 500 Then
                Inform " add to fr0_10"
                nCountS = 0
                If cStr2 <> "" Then con.Execute cStr2
                If cStr3 <> "" Then con.Execute cStr3
                cStr2 = ""
                cStr3 = ""
            End If
        Else
            Inform loctable2!doc_no & " " & loctable2!Item
            XCOST.Caption = XCOST.Caption & loctable2!Item & Chr(13) & " T Count : " & nCountTot
        End If
        loctable2.MoveNext
    Loop
    
    If cStr2 <> "" Then con.Execute cStr2
    If cStr3 <> "" Then con.Execute cStr3
    
'    con.Execute " UPDATE FR0_10H SET CLOSED = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    If loctable!CLOSED Then
        Inform " Post Stock " & loctable!doc_no
        con.Execute " UPDATE FR0_10H SET CLOSED = 1 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    Else
        con.Execute " UPDATE FR0_10H SET CLOSED = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    End If
    conShop2.Execute " UPDATE FILE0_10H  SET ISNEW = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    getStock = getStock + nAffect
Loop
Inform " «·Ã—œ " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getStock = -1
GoTo lastsub
End Function
Private Function SendDiscount() As Long
'On Error GoTo myerror
cString = "SELECT * FROM DISCOUNT "
Me.Caption = "Œ’„"
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop2.Execute " delete from DISCOUNT "
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "ROW", addstring(loctable!Row))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[DATE1]", addDate(Format(loctable!Date1, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[DATE2]", addDate(Format(loctable!date2, "DD-MM-YYYY")))
    aInsert = AddFlag(aInsert, "[DISC]", Val(loctable!disc & ""))
    aInsert = AddFlag(aInsert, "[MAXDISC]", Val(loctable!MaxDisc & ""))
    aInsert = AddFlag(aInsert, "[ISPRICE_2]", IIf(loctable!ISPRICE_2, 1, 0))
    aInsert = AddFlag(aInsert, "[ISDISC2]", IIf(loctable!ISdisc2, 1, 0))
    
    conShop2.Execute addInsert(aInsert, "DISCOUNT"), nAffect
    loctable.MoveNext
    SendDiscount = SendDiscount + nAffect
Loop
lastsub:
prog1.Visible = False

Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendDiscount = -1
GoTo lastsub
End Function
Function LoadConStringshop2(pServerIp, pServerData)
    Dim cServerName As String, cUserId As String, cPassword As String
    cUserId = "DATA_PRO" 'decrypt(RetSetting("userId", App.Path & "\conf.txt"), "dr")
'    cPassword = decrypt(RetSetting("Password", App.Path & "\conf.txt"), "dr")
    cPassword = "2015"
    LoadConStringshop2 = "provider=SQLOLEDB;data source=" & pServerIp & ";initial " _
                & "catalog=" & pServerData & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Private Function ToBranch_COST() As Boolean
'On Error GoTo myerror
Dim lMyPrice As Boolean, lMyPriceItem As Boolean
Inform " aa " & grid1.TextMatrix(grid1.Row, 1)
Inform "Send All Price "

nRecordCount = sendItems_COST
If nRecordCount >= 0 Then Inform " „  ⁄œÌ· «”⁄«— & »«—ﬂÊœ " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror

ToBranch_COST = True
Inform " „ ‰ﬁ· «·»Ì«‰« "
ToBranch_COST = True
Exit Function
ToBranch_COST = False
myerror:
MsgBox "·„ Ì „ ‰ﬁ· ﬂ«›… «·»Ì«‰«  »‰Ã«Õ"
Err.Clear
End Function
Private Function sendItems_COST() As Long
'On Error GoTo myerror
Dim cFile As String, cString As String, cStrall As String, nCount As Double, cStrDel1 As String
cFile = "FILE1_10"
Inform "Select Item "
If Check6.Value <> 0 Then
    cString = "SELECT * FROM FILE1_10 WHERE   ISSTOP = 1  "
Else

    If XBAR13.Value <> 0 Then
        cString = "SELECT * FROM FILE1_10 WHERE  ( BARCODE13 IS NOT NULL  OR BARCODE_GS1 IS NOT NULL  ) AND ITEM IN (SELECT ITEM FROM FR1_11 WHERE FR1_11.STORE = " & MyParn(pBranch) & ")  "
    Else
        If Check2.Value <> 0 Then
            cString = "SELECT * FROM FILE1_10 WHERE  ITEM IN (SELECT ITEM FROM FR1_11 WHERE FR1_11.STORE = " & MyParn(pBranch) & ")  "
        Else
            cString = "SELECT * FROM FILE1_10 WHERE  ITEM IN (SELECT ITEM FROM FR1_11 WHERE  FR1_11.STORE = " & MyParn(pBranch) & ")   AND ITEM IN ( SELECT ITEM FROM NEW_ITEM_FR WHERE BRANCH = " & MyParn(pBranch) & " ) "
        End If
    End If
End If
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
'Inform cString
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
If Check6.Value <> 0 Then
    conShop2.Execute " UPDATE FILE1_10 SET ISSTOP = 0 WHERE ISSTOP = 1 "
End If

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
sCaption = Me.Caption
Do Until loctable.EOF
    nCount = nCount + 1
    nRecord = nRecord + 1
'   Me.Caption = loctable!Item
    prog1.Visible = True
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If Check6.Value <> 0 Then
        cStrall = cStrall & " UPDATE FILE1_10 SET ISSTOP = 1  WHERE ITEM = " & Val(loctable!Item & "") & " ; "
    Else
        cStrall = cStrall & " UPDATE FILE1_10 SET COSTITEM = " & Val(loctable!costITEM_FR & "") & " , cost = " & Val(loctable!cost & "") & " , cost2 = " & Val(loctable!cost2 & "") & " , price = " & Val(loctable!price & "") & " , price2 = " & Val(loctable!PRICE2 & "") & " , price_2 = " & Val(loctable!PRICE_2 & "") & " , OKAZ = " & Val(loctable!okaz & "") & " , OKAZ_2 = " & Val(loctable!okaz_2 & "") & " , BARCODE13 = " & addstring(loctable!BARCODE13) & " , BARCODE_GS1 = " & addstring(loctable!BARCODE_GS1) & " WHERE ITEM = " & Val(loctable!Item & "") & " ; "
    End If
    If Check2.Value = 0 And Check6.Value = 0 Then
        cStrDel1 = cStrDel1 & " DELETE FROM NEW_ITEM_FR WHERE ITEM = " & loctable!Item & " AND BRANCH = " & MyParn(pBranch) & " ; "
    End If
    
    If nCount = 100 Then
        conShop2.Execute cStrall
        If cStrDel1 <> "" Then con.Execute cStrDel1
        nCount = 0
        cStrDel1 = ""
        cStrall = ""
    End If
    Me.Caption = loctable!Item
    loctable.MoveNext
Loop
If cStrall <> "" Then conShop2.Execute cStrall
If cStrDel1 <> "" Then con.Execute cStrDel1
lastsub:
Me.Caption = sCaption
prog1.Visible = False
sendItems_COST = nRecord
'If Check4.Value <> 0 Then
'    cStr1 = " UPDATE    FR6_20 SET  cost = (SELECT TOP 1 (COST ) FROM Q_ITEMCOST_FR WHERE FR6_20.ITEM = Q_ITEMCOST_FR.ITEM AND Q_ITEMCOST_FR.DATE <= FR6_20H.DATE ORDER BY DATE DESC ) " & _
'            " FROM      FR6_20 INNER JOIN FR6_20H ON FR6_20.DOC_NO = FR6_20H.DOC_NO where year(FR6_20h.date) >= " & Year(Date)
'    conShop2.CommandTimeout = 1000
'    conShop2.Execute cStr1, nRec
'    Inform "  „ Ÿ»ÿ  ﬂ·›… „»Ì⁄«  " & grid1.TextMatrix(grid1.Row, 1) & "  " & nRec
'End If
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendItems_COST = -1
GoTo lastsub
End Function
Private Function NewPrice() As Long
'On Error GoTo myerror
cString = "SELECT * FROM NEWPRICEITEM WHERE ISNEW" & nBranchPost & "  = 1  "
Dim loctable As New ADODB.Recordset, cStr1 As String, cStr2 As String, cStr3 As String
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!MODEL))
    If cStr1 = "" Then
        cStr1 = addInsert(aInsert, "NEWPRICEITEM")
        cStr2 = " delete from NEWPRICEITEM where model = " & MyParn(loctable!MODEL)
        cStr3 = " update    NEWPRICEITEM set    ISNEW" & nBranchPost & "  = 0 where model = " & MyParn(loctable!MODEL)
    Else
        cStr1 = cStr1 & " ; " & addInsert(aInsert, "NEWPRICEITEM")
        cStr2 = cStr2 & " ; " & " delete from NEWPRICEITEM where model = " & MyParn(loctable!MODEL)
        cStr3 = cStr3 & " ; " & " update    NEWPRICEITEM set    ISNEW" & nBranchPost & "  = 0 where model = " & MyParn(loctable!MODEL)
    End If
    loctable.MoveNext
    NewPrice = NewPrice + nAffect
Loop
If Check1.Value <> 0 Then conShop2.Execute " delete from NEWPRICEITEM "

If cStr2 <> "" Then conShop2.Execute cStr2
If cStr1 <> "" Then conShop2.Execute cStr1
If cStr3 <> "" Then con.Execute cStr3

Inform "  ⁄œÌ· «”⁄«— " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
NewPrice = -1
GoTo lastsub
End Function
Private Function sendTransHeader2() As Long

cString = "SELECT * FROM FR1_61h CODE WHERE ISNEW = 1 AND STORE1 = " & MyParn(pBranch)

Dim loctable As New ADODB.Recordset
Dim cStr1 As String, cStr2 As String
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
FixSubShop1
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[TYPEDOC]", addstring(loctable!typedoc))
    aInsert = AddFlag(aInsert, "[DOC_TRANS]", addstring(loctable!DOC_TRANS))
    aInsert = AddFlag(aInsert, "[NOTES2]", addstring(loctable!NOTES2))
    cStr1 = cStr1 & addUpdate(aInsert, "FILE1_61H", " DOC_NO = " & MyParn(loctable!doc_no))
    cStr2 = cStr2 & "UPDATE FR1_61H SET ISNEW = 0 " & " ; "
    loctable.MoveNext
    sendTransHeader2 = sendTransHeader2 + nAffect
Loop
If cStr1 <> "" Then conShop2.Execute cStr1
If cStr2 <> "" Then con.Execute cStr2
Inform "  ÕÊÌ·«  Ê«—œ… " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTransHeader2 = -1
GoTo lastsub
End Function
Sub FixSubShop1()
Dim cString As String
On Error Resume Next
cString = " alter table [dbo].[FILE1_61H] add   [Notes2] [nvarchar](200) NULL"
createCommand cString, conShop2
Err.Clear
End Sub
Sub SubFixData()
    On Error Resume Next
    Dim cString As String
    
    
'    cString = " ALTER TABLE [dbo].[FILE0_90] ADD [DISC2] [int] NOT NULL CONSTRAINT [DF_FILE0_90_DISC2]  DEFAULT ((0)) "
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE1_10SC] ADD [ISDISC] [bit] NULL"
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[DISCOUNT] ADD [ISDISC2] [bit] NULL"
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE6_20H] ADD [NOTE_DISC] [nvarchar](50) NULL"
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE6_20] ADD [ISDISC2_1] [bit]  NOT NULL CONSTRAINT [DF_FILE6_2_ISDISC2_1]  DEFAULT ((0)) "
'    createCommand cString, conShop2
'
'    cString = " alter table [dbo].[DISCOUNT] add   [MAXDISC] [decimal](18, 2) NULL"
'    createCommand cString, conShop2
'
'
'    cString = " create VIEW [dbo].[FILE1_11_ALL] AS SELECT  * FROM FILE1_11"
'    createCommand cString, conShop2
'
'    createCommand cString, conShop2
'    cString = " create VIEW [dbo].[FILE1_11_ALL] AS SELECT FILE1_11.* FROM FILE1_11 inner join file0_40 on file0_40.code = file1_11.store where file0_40.ISSTOP = 0   UNION ALL SELECT * FROM FR1_11"
'
'    cString = "ALTER TABLE [dbo].[FILE4_10] ADD [disc2] [decimal](18, 2) NULL"
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE1_10] ADD    [MODELSEX] [varchar](3) NULL , [MODELAGE] [varchar](3) NULL"
'    createCommand cString, conShop2
'
'    cString = "ALTER VIEW [dbo].[FILE1_10H] AS SELECT        MODEL, MAX(desca) AS DESCA, MAX([GROUP]) AS [GROUP], MAX(RATE) AS RATE, MAX(OKAZ) AS OKAZ, MAX(FACT) AS FACT, MAX(MOSM) AS MOSM,  MAX(modelno) AS MODELNO, MAX(MODELFACT) AS MODELFACT, MAX(code) AS code, MAX(SUPP) AS SUPP, MAX(MODELFACT0) AS MODELFACT0, REDEM,  FIXPRICE, SHOWSALES, ISOKAZITEM, SECTION, ISSTOP, OKAZ_2, MOSM2, ISNODEM, ISNOITEM, MODELSEX, MODELAGE FROM            dbo.FILE1_10 GROUP BY MODEL, REDEM, FIXPRICE, SHOWSALES, ISOKAZITEM, SECTION, ISSTOP, OKAZ_2, MOSM2, ISNODEM, ISNOITEM, MODELSEX, MODELAGE"
'    createCommand cString, conShop2
'
'    cString = "CREATE TABLE [dbo].[FILE0_81](  [CODE] [int] NOT NULL, [DESCA] [nvarchar](100) NULL, [ISDISC] [bit] NULL, CONSTRAINT [PK_FILE0_81] PRIMARY KEY CLUSTERED  ( [CODE] ASC )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ) ON [PRIMARY]"
'    createCommand cString, conShop2
'
'    cString = "CREATE TABLE [dbo].[FILE0_82](  [CODE] [int] NOT NULL, [DESCA] [nvarchar](100) NULL, [ISDISC] [bit] NULL, CONSTRAINT [PK_FILE0_82] PRIMARY KEY CLUSTERED  ( [CODE] ASC )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ) ON [PRIMARY]"
'
'    createCommand cString, conShop2
'
'    cString = " ALTER VIEW [dbo].[QBRANCH_ALL] AS SELECT        code, desca, STORE FROM            dbo.branch    "
'    createCommand cString, conShop2
'
'    cString = " CREATE VIEW [dbo].[Q_ITEMCOST] AS SELECT        FILE7_20.ITEM, FILE7_20H.DATE, FILE7_20.TOTAL / FILE7_20.QUANT AS cost FROM            FILE7_20 INNER JOIN                          FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO"
'    createCommand cString, conShop2
'
'    cString = "ALTER  TABLE [dbo].[FILE1_10]  ADD [BARCODE13] [nvarchar](50) NULL"
'    createCommand cString, conShop2
'
'    cString = "CREATE TABLE [dbo].[MODEL_STOP](    [MODEL] [nvarchar](20) COLLATE Arabic_CI_AS NULL,   [BRANCH] [nvarchar](3) COLLATE Arabic_CI_AS NULL,   [ISSTOP] [bit] NOT NULL CONSTRAINT [DF_MODEL_STOP_ISSTOP]  DEFAULT ((1)),   [ID] [int] IDENTITY(1,1) NOT NULL) ON [PRIMARY]"
'    createCommand cString, conShop2
'
'    cString = " CREATE NONCLUSTERED INDEX [IX_MODEL_STOP] ON [dbo].[MODEL_STOP](    [MODEL] ASC,    [BRANCH] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, SORT_IN_TEMPDB = OFF, DROP_EXISTING = OFF, ONLINE = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]    "
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[VISA] ADD [RATE6] [decimal](18, 4) NULL, [RATE18] [decimal](18, 4) NULL"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE6_20H] ADD [TYPEVISA] [int] NULL   , [VISA2] [decimal](18, 2) NOT NULL CONSTRAINT [DF_file6_20h_VISA2]  DEFAULT ((0))"
'    createCommand cString, conShop2
'
'    cString = " CREATE TABLE [dbo].[SUB_USER](  [CODE] [nvarchar](4) COLLATE Arabic_CI_AS NULL, [DESCA] [nvarchar](50) COLLATE Arabic_CI_AS NULL,[PASSWORD] [nvarchar](50) COLLATE Arabic_CI_AS NULL,   [BRANCH] [nvarchar](3) COLLATE Arabic_CI_AS NULL) ON [PRIMARY]"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[VISA] ADD [ISSTOP] [bit] NOT NULL"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[file1_10] ADD [ISSTOP] [bit] NOT NULL CONSTRAINT [DF_file1_10_ISSTOP]      DEFAULT ((0))"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE6_20H] ADD   [SHIP] [nvarchar](5) NULL,  [NAME] [nvarchar](50) NULL, [CHARGE1] [numeric](18, 2) NULL,    [CHARGE2] [numeric](18, 2) NULL,    [SHIP_NO] [nvarchar](50) NULL,  [DATE_PAY] [datetime] NULL, [ADDRESS] [nvarchar](1000) NULL,    [ONLINE] [int] NULL,    [BANK] [nvarchar](6) NULL"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE6_20H] ADD  CONSTRAINT [DF_FILE6_20H_ONLINE]  DEFAULT ((0)) FOR [ONLINE]"
'    createCommand cString, conShop2
'
'    cString = " CREATE TABLE SHIP ( [CODE] [nvarchar](3) NULL,[DESCA] [nvarchar](100) NULL) ON [PRIMARY]    "
'    createCommand cString, conShop2
'
'    cString = " alter TABLE [dbo].[VISA] add    [ISSTOP] [bit] NOT NULL"
'    createCommand cString, conShop2
'    cString = " ALTER TABLE [dbo].[VISA] ADD  CONSTRAINT [DF_VISA_ISSTOP]  DEFAULT ((0)) FOR [ISSTOP]"
'    createCommand cString, conShop2
'
'    conShop2.Execute " UPDATE FILE6_20H SET SHIP = '0' + FILE6_20H.SHIP  WHERE LEN(SHIP) = 1 ", nRec
'
'    cString = " CREATE TABLE [dbo].[FILE0_90H]( [DOC_NO] [varchar](6) NOT NULL, [DATE] [datetime] NULL, [DESCA] [nvarchar](250) NULL,   [QTY1] [int] NULL,  [QTY2] [int] NULL,  [DISC] [numeric](18, 2) NULL,   [ISSTOP] [bit] NOT NULL) ON [PRIMARY]"
'    createCommand cString, conShop2
'
'    If Val(GetDesca("SELECT COUNT(*) FROM FILE0_90", conShop2) & "") = 0 Then
'        cString = " CREATE TABLE [dbo].[FILE0_90](  [DOC_NO] [varchar](6) NULL,     [MODELNO] [nvarchar](50) NULL,  [id] [int] IDENTITY(1,1) NOT NULL,      [ISSELECT] [bit] NOT NULL ) ON [PRIMARY]"
'        createCommand cString, conShop2
'    End If
'
'    cString = " CREATE TABLE [dbo].[FILE6_80]( [DOC_NO] [nvarchar](6) NULL, [DATE1] [smalldatetime] NULL, [DATE2] [smalldatetime] NULL, [BRANCH] [nvarchar](3) NULL, [VALUE_TARGET] [numeric](18, 2) NULL, [ID] [int] IDENTITY(1,1) NOT NULL, CONSTRAINT [PK_FILE6_80] PRIMARY KEY CLUSTERED  ( [ID] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE1_10] ADD [barcode_gs1] [nvarchar](50) NULL  "
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE1_10] ADD [PRICE_P] [decimal](18, 2) NULL"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE0_90] ADD    [DISC2] [int] NOT NULL CONSTRAINT [DF_FILE0_90_DISC2]  DEFAULT ((0)) "
'    createCommand cString, conShop2
    
'    cString = " ALTER TABLE [dbo].[FILE6_20H] ADD   [DOC_DISC] [nvarchar](6) NULL "
'    createCommand cString, conShop2
'
'    cString = " CREATE PROCEDURE [dbo].[Item_BARCODE]    @BARCODE nVarchar(20),  @RET int  OUTPUT AS BEGIN DECLARE @nitem int  SELECT  @nITEM  = item From FILE1_10 WHERE ITEM = @BARCODE OR BARCODE = @BARCODE OR BARCODE2 = @BARCODE OR BARCODE13 = @BARCODE OR BARCODE_GS1 = @BARCODE  set @ret = @nitem END"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE6_20H] ADD   [DOC_DISC] [nvarchar](6) NULL "
'    createCommand cString, conShop2
'
'    cString = " CREATE PROCEDURE [dbo].[ITEMFIND_BARCODE]   @BARCODE NVARCHAR(30) AS BEGIN      SELECT  FILE1_10.* From FILE1_10 WHERE  BARCODE = @BARCODE OR BARCODE2 = @BARCODE OR BARCODE13 = @BARCODE OR BARCODE_GS1 = @BARCODE END"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[ADDRESS] ALTER COLUMN [head1] [nvarchar] (1000) NULL     "
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[SUB_USER] ADD    [PASSWORD2] [nvarchar](50) NULL,    [STOP1] [bit] NULL,     [STOP2] [bit] NULL"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE6_20] ADD    [S_OKAZ_2] [decimal](18, 2) NULL "
'    createCommand cString, con
'
'
'    cString = " ALTER VIEW [dbo].[Q_DOC_SALES_DISC] AS SELECT        dbo.FILE6_20.ITEM, dbo.FILE6_20.QUANT, dbo.FILE6_20.PRICE, dbo.FILE0_90H.DOC_NO, dbo.FILE0_90H.QTY1, dbo.FILE0_90H.DISC, dbo.FILE6_20.DOC_NO AS DOC_SALES " & _
'                " FROM            dbo.FILE0_90H INNER JOIN dbo.FILE0_90 ON dbo.FILE0_90H.DOC_NO = dbo.FILE0_90.DOC_NO INNER JOIN dbo.FILE6_20 INNER JOIN dbo.FILE1_10 ON dbo.FILE6_20.ITEM = dbo.FILE1_10.ITEM ON dbo.FILE0_90.MODELNO = dbo.FILE1_10.modelno " & _
'                " WHERE        (dbo.FILE0_90H.ISSTOP = 0) AND (dbo.FILE0_90H.DATE1 <= CAST(GETDATE() AS DATE)) AND (dbo.FILE0_90H.DATE2 >= CAST(GETDATE() AS DATE)) "
'    createCommand cString, conShop2
'
'    cString = " ALTER VIEW [dbo].[Q_ITEMCOST] AS SELECT        FILE7_20.ITEM, FILE7_20H.DATE, FILE7_20.TOTAL / FILE7_20.QUANT AS cost FROM            FILE7_20 INNER JOIN                          FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO  where quant <> 0     "
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE6_20H] ADD   [DOC_AGE]  AS (datediff(day,[DATE],getdate())) "
'    createCommand cString, conShop2
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [note_ret] [nvarchar](100) NULL     "
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[ADDRESS] ADD     [DATE1_DISC] [nvarchar](50) NULL,   [DATE2_DISC] [nvarchar](50) NULL,   [DISC] [numeric](18, 2) NULL     "
'    createCommand cString, con
'''''
'
'    cString = " create VIEW [dbo].[QFILE6_20] AS SELECT        FILE6_20H.DOC_NO, FILE6_20H.store, FILE6_20H.date, FILE6_20H.CASH, FILE6_20H.VISA, FILE6_20H.BOX, FILE6_20H.PAY, FILE6_20H.branch,                           FILE6_20H.username_disc, FILE6_20.ITEM, FILE6_20.QUANT, FILE6_20.PRICE, FILE6_20.TOTAL, FILE6_20H.discount, FILE6_20H.rate, FILE6_20.cost,                          FILE6_20.PRICE_C, FILE6_20.MAN, FILE6_20.ISDISC2_1, FILE6_20.S_OKAZ FROM            FILE6_20H INNER JOIN                          FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO    "
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE0_50]  ADD  [ISBOXONLINE] [bit] NOT NULL CONSTRAINT [DF_FILE0_50_ISBOXONLINE] DEFAULT ((0)) "
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE0_40]  ADD  [ONLINE] [bit] NOT NULL CONSTRAINT [DF_FILE0_40_ONLINE] DEFAULT ((0)) "
'    createCommand cString, conShop2
'
'    cString = "  ALTER VIEW [dbo].[QFILE6_20] AS   SELECT        FILE6_20H.DOC_NO, FILE6_20H.store, FILE6_20H.date, FILE6_20H.CASH, FILE6_20H.VISA, FILE6_20H.BOX, FILE6_20H.PAY, FILE6_20H.branch, FILE6_20H.username_disc, FILE6_20.ITEM, FILE6_20.QUANT,  (FILE6_20.PRICE  ) AS PRICE, (FILE6_20.TOTAL ) AS TOTAL, FILE6_20H.discount, FILE6_20H.rate, FILE6_20.cost, FILE6_20.PRICE_C, FILE6_20.MAN, FILE6_20.ISDISC2_1,  FILE6_20.S_OKAZ, FILE1_10.MODEL, FILE1_10.[GROUP], FILE1_10.code, FILE1_10.MOSM, FILE1_10.SECTION, 1 AS iSBRANCH, FILE6_20H.SALES_RET, FILE6_20H.username_ret, FILE6_20H.PHONE,  0 as Outlet,  DOC_DISC, FILE6_20H.CODE AS CUST , ISRET FROM  FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM    "
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE1_10]  ADD  [ISNOITEM2] [bit] NOT NULL CONSTRAINT [DF_FILE1_10_ISNOITEM2] DEFAULT ((0)) "
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[file6_20h] ADD [NORET_CHARGE1] [bit] NOT NULL CONSTRAINT [DF_file6_20H_NORET_CHARGE1] DEFAULT ((0))"
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[SUB_USER]       ADD  [ISRETURN] [bit] NOT NULL CONSTRAINT [DF_SUB_USER_ISRETURN] DEFAULT ((0)) "
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE0_40] ADD   [ISNOGRDISC] [bit] NOT NULL CONSTRAINT [DF_FILE0_40_ISNOGRDISC]  DEFAULT ((0))     "
'    createCommand cString, conShop2
'
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [CARD_DISC] [nvarchar](10) NULL     "
'    createCommand cString, conShop2
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [CARD_DOC] [nvarchar](6) NULL     "
'    createCommand cString, conShop2
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [CARD_DESCA] [nvarchar](100) NULL     "
'    createCommand cString, conShop2
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [CARD_VALUE] [numeric](18, 2) NULL   "
'    createCommand cString, conShop2
'
'    ctrsing = " alter TABLE [dbo].[FILE6_20H]  add  [visa3] [decimal](18, 2) NOT NULL CONSTRAINT [DF_FILE6_20H_visa3]  DEFAULT ((0)),  [CODEVISA3] [nvarchar](6) NULL,     [TYPEVISA3] [nvarchar](3) NULL,     [VISADOC3] [nvarchar](50) NULL, [visa3_2] [decimal](18, 2) NOT NULL CONSTRAINT [DF_FILE6_20H_visa3_2]  DEFAULT ((0))  "
'    createCommand cString, conShop2
'
'
'    ctrsing = ReadFile(App.Path & "\TXT\QQFILE6_20.TXT")
'    createCommand cString, conShop2
'
'    ctrsing = ReadFile(App.Path & "\TXT\QQFILE6_20H.TXT")
'    createCommand cString, conShop2
'
'    ctrsing = ReadFile(App.Path & "\TXT\QSALES_MODEL.TXT")
'    createCommand cString, conShop2
'
'    ctrsing = ReadFile(App.Path & "\TXT\QSALESDTL.TXT")
'    createCommand cString, conShop2
'
'    ctrsing = ReadFile(App.Path & "\TXT\QT_SALESDOC.TXT")
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE6_20H] ADD   [VISA3] [numeric](18, 2) NOT NULL CONSTRAINT [DF_FILE6_20H_VISA3]  DEFAULT ((0)),  [CODEVISA3] [nvarchar](6) NULL, [TYPEVISA3] [nvarchar](3) NULL, [VISADOC3] [nvarchar](50) NULL, [visa3_2] [numeric](18, 2) NOT NULL CONSTRAINT [DF_FILE6_20H_visa3_2]  DEFAULT ((0)),    [DATE_REC2] [datetime] NULL,    [ISRETONLINE] [bit] NOT NULL CONSTRAINT [DF_FILE6_20H_ISRETONLINE]  DEFAULT ((0))"
'    createCommand cString, conShop2
'
'    cString = ReadFile(App.Path & "\TXT\QFILE6_20.TXT")
'    createCommand cString, conShop2
'
'    cString = ReadFile(App.Path & "\TXT\SALESDTL.TXT")
'    createCommand cString, conShop2
'
'    cString = ReadFile(App.Path & "\TXT\T_SALESDOC.TXT")
'    createCommand cString, conShop2
'
'    cString = " ALTER VIEW [dbo].[QFILE6_20H] AS SELECT        DOC_NO, code, store, date, isclosed, discount, username, rate, tax, DESCA, MAN, CASH, VISA+VISA3 AS VISA, BOX, NOTES, PAY, PRINTED, TIME, closed, rest, late, code2, PHONE, username_disc, SALES_RET, username_ret, POINT, INV_NO, NOTE_DISC, DOC_DISC , ISRET  FROM            FILE6_20H"
'    createCommand cString, conShop2
'
    
    
    Err.Clear
End Sub

Private Function SendSection() As Long
'On Error GoTo myerror

Dim cFile As String, nRecord As Double
Dim loctable As New ADODB.Recordset
Dim cStr1  As String, cStr2   As String
cFile = "FILE1_10SC"
cString = "SELECT * FROM " & cFile
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "ISDISC", IIf(loctable!ISDISC, 1, 0))
    cStr1 = cStr1 & addInsert(aInsert, cFile) & " ; "
    loctable.MoveNext
    SendSection = SendSection + nAffect
Loop
conShop2.Execute " DELETE FROM FILE1_10SC "
conShop2.Execute cStr1
Inform " «·√ﬁ”«„ " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
SendSection = -1
Err.Clear
GoTo lastsub
End Function
Private Function FixItemDel() As Long

'On Error GoTo myerror
cString = "SELECT * FROM BARCODE_EDIT "

Me.Caption = " ⁄œÌ· »«—ﬂÊœ"
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount


'    conShop2.Execute " UPDATE FILE6_20  SET ITEM = 94685 WHERE PRICE < 100 AND ITEM =94684 "
'    conShop2.Execute " UPDATE FILE7_10  SET ITEM = 94685 WHERE PRICE < 100 AND ITEM =94684 "
'    conShop2.Execute " UPDATE FILE7_20  SET ITEM = 94685 WHERE PRICE < 100 AND ITEM =94684 "
'    conShop2.Execute " UPDATE FILE1_60  SET ITEM = 94685 WHERE COST < 100 AND ITEM =94684 "
'    conShop2.Execute " UPDATE FILE1_61  SET ITEM = 94685 WHERE COST < 100 AND ITEM =94684 "
'
'    conShop2.Execute " UPDATE FILE0_10  SET ITEM = 94685 WHERE COST < 100 AND ITEM =94684 "
'    conShop2.Execute " UPDATE FILE0_100 SET ITEM = 94685 WHERE COST < 100 AND ITEM =94684 "
'
'    If pBranch = "04" Or pBranch = "03" Then
'        conShop2.Execute " UPDATE FILE1_11_ALL   SET ITEM = 94685 WHERE COST  < 100 AND ITEM = 94684 "
'    End If



Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If GetDesca("select item from file1_11 where item   = " & loctable!ITEMOLD, conShop2) <> "" Then
        conShop2.Execute " UPDATE FILE6_20 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD), nRec
        Inform "REC SALES " & nRec
        conShop2.Execute " UPDATE FILE7_10 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        conShop2.Execute " UPDATE FILE7_20 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        conShop2.Execute " UPDATE FILE1_60 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        conShop2.Execute " UPDATE FILE1_61 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        
        conShop2.Execute " UPDATE FILE0_10  SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        conShop2.Execute " UPDATE FILE0_100 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
    End If
    conShop2.Execute " DELETE FROM FILE1_10 WHERE ITEM = " & (loctable!ITEMOLD)
    loctable.MoveNext
    FixItemDel = FixItemDel + nAffect
Loop
lastsub:
prog1.Visible = False

Exit Function
myerror:
MsgBox Err.Description
Err.Clear
FixItemDel = -1
GoTo lastsub
End Function

Private Function sendItems2() As Long
On Error GoTo myerror
Dim cFile As String, nCount As Double
Dim cStrPostAll As String, cStrPostAll_DEL As String, cStrDel1 As String
cFile = "FILE1_10"

cString = "SELECT FILE1_11.ITEM FROM FILE1_11 LEFT JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE  FILE1_10.ITEM IS NULL "
Dim loctable As New ADODB.Recordset
Dim loctable2 As New ADODB.Recordset
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
sCaption = Me.Caption
Do Until loctable.EOF
    nCount = nCount + 1
    Me.Caption = loctable!Item
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open "SELECT * FROM FILE1_10 WHERE ITEM = " & loctable!Item, con, adOpenStatic, adCmdText
    If Not loctable2.EOF And Not loctable2.BOF Then
        aInsert = AddFlag(Empty, "MODEL", addstring(loctable2!MODEL))
        aInsert = AddFlag(aInsert, "MODELNO", addstring(loctable2!MODELNO))
        aInsert = AddFlag(aInsert, "DESCA", addstring(loctable2!DESCA))
        aInsert = AddFlag(aInsert, "FACT", addstring(loctable2!Fact))
        aInsert = AddFlag(aInsert, "MOSM", addstring(loctable2!MOSM))
        aInsert = AddFlag(aInsert, "MODELFACT", addstring(loctable2!modelfact))
        aInsert = AddFlag(aInsert, "MODELFACT0", addstring(loctable2!modelfact0))
        aInsert = AddFlag(aInsert, "SUPP", addstring(loctable2!SUPP))
        aInsert = AddFlag(aInsert, "OKAZ", Val(loctable2!okaz & ""))
        aInsert = AddFlag(aInsert, "OKAZ_2", Val(loctable2!okaz_2 & ""))
        aInsert = AddFlag(aInsert, "RATE", Val(loctable2!Rate & ""))
        aInsert = AddFlag(aInsert, "CODE", addstring(loctable2!CODE))
        aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable2!Group))
        aInsert = AddFlag(aInsert, "[SECTION]", addvalue(loctable2!Section))
        aInsert = AddFlag(aInsert, "[ITEM]", addvalue(loctable2!Item))
        aInsert = AddFlag(aInsert, "[COSTITEM]", Val(loctable2!costITEM_FR & ""))
        aInsert = AddFlag(aInsert, "[COST2]", Val(loctable2!cost2 & ""))
        aInsert = AddFlag(aInsert, "MODELSEX", addvalue(loctable2!MODELSEX))
        aInsert = AddFlag(aInsert, "MODELAGE", addvalue(loctable2!MODELAGE))
    
        aInsert = AddFlag(aInsert, "[PRICE]", Val(loctable2!price & ""))
        aInsert = AddFlag(aInsert, "[PRICE_2]", Val(loctable2!PRICE_2 & ""))
        aInsert = AddFlag(aInsert, "[SCAL]", addstring(loctable2!scal))
        aInsert = AddFlag(aInsert, "[C_SCAL]", addvalue(loctable2!C_SCAL))
        aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable2!color))
        aInsert = AddFlag(aInsert, "[C_COLOR]", addvalue(loctable2!c_Color))
        aInsert = AddFlag(aInsert, "[BARCODE]", addstring(loctable2!BARCODE))
        aInsert = AddFlag(aInsert, "[BARCODE2]", addstring(loctable2!BARCODE2))
        aInsert = AddFlag(aInsert, "[BARCODE13]", addstring(loctable2!BARCODE13))
        aInsert = AddFlag(aInsert, "[BARCODE_GS1]", addstring(loctable2!BARCODE_GS1))
        aInsert = AddFlag(aInsert, "[ISOKAZITEM]", IIf(loctable2!ISOKAZITEM, 1, 0))
        aInsert = AddFlag(aInsert, "[ISNOITEM]", IIf(loctable2!ISNOITEM, 1, 0))
        aInsert = AddFlag(aInsert, "[ISNODEM]", IIf(loctable2!ISNODEM, 1, 0))
        aInsert = AddFlag(aInsert, "[MOSM2]", addstring(loctable2!MOSM2))
        aInsert = AddFlag(aInsert, "[PRICE2]", Val(loctable2!PRICE2 & ""))
        aInsert = AddFlag(aInsert, "[COST]", Val(loctable2!cost & ""))
    
        cStrPostAll = cStrPostAll & " DELETE FROM FILE1_10 where item = " & loctable!Item & " ; "
        cStrPostAll = cStrPostAll & addInsert(aInsert, "FILE1_10") & " ; "
        If nCount = 100 Then
            If cStrPostAll <> "" Then
                conShop2.Execute cStrPostAll, nRec
                If cStrDel1 <> "" Then con.Execute cStrDel1
                cStrDel1 = ""
                cStrPostAll = ""
            End If
            nCount = 0
        End If
    Else
'        conShop2.Execute " UPDATE FILE6_20 SET ITEM = 101  WHERE ITEM = " & loctable!Item
'        conShop2.Execute " UPDATE FILE0_10 SET ITEM = 101  WHERE ITEM = " & loctable!Item
'        conShop2.Execute " UPDATE FILE7_20 SET ITEM = 101  WHERE ITEM = " & loctable!Item
'        conShop2.Execute " UPDATE FILE7_10 SET ITEM = 101  WHERE ITEM = " & loctable!Item
    End If
    Me.Caption = loctable!Item
    loctable.MoveNext
    sendItems2 = sendItems2 + 1
Loop
If cStrPostAll <> "" Then
    conShop2.Execute cStrPostAll, nRec
    If cStrDel1 <> "" Then con.Execute cStrDel1
End If
Inform "  „ ‰ﬁ· «·«’‰«›2 " & nRecord
lastsub:
Me.Caption = sCaption
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendItems2 = -1
GoTo lastsub
End Function

Private Sub Command3_Click()
    Me.Caption = " ⁄œÌ· IP  «·„’‰⁄ "
    With grid1
    For nRow = 1 To .Rows - 1
        If .TextMatrix(nRow, 2) <> "" Then
            If .TextMatrix(nRow, 2) <> "" And TurnValue(.TextMatrix(nRow, 4), "True", True) Then
                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbYellow
                IpShop = .TextMatrix(nRow, 2)
                cDataShop = .TextMatrix(nRow, 3)
                lAddAll = TurnValue(.TextMatrix(nRow, 5), "True", True)
                strConShop2 = LoadConStringshop2(IpShop, cDataShop)
                nBranchPost = Val(.TextMatrix(nRow, 0))
                pBranch = .TextMatrix(nRow, 0)
                
                If openConShop(conShop2, strConShop2) = "ok" Then
                    Inform "  „ «·≈ ’«· »«·›—⁄  "
                    If XIP.text <> "" Then
                        conShop2.Execute " UPDATE BRANCH SET IPSERVER = " & addstring(XIP.text)
                    Else
                        conShop2.Execute " UPDATE BRANCH SET DATASERVER = 'JUNIOR_SHOP' "
                    End If
                Else
                    .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
                End If
            End If
            If conShop2.State = adStateOpen Then closeCon conShop
        End If
    Next nRow
    End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear

End Sub
Private Function sendSUB_USER()
'On Error GoTo myerror
cString = "SELECT * FROM SUB_USER "

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop2.Execute " delete from SUB_USER"
Dim aInsert As Variant
cStr1 = ""
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[PassWord]", IIf(loctable!Stop1, addstring(Null), addstring(loctable!PassWord)))
    aInsert = AddFlag(aInsert, "[PassWord2]", IIf(loctable!Stop2, addstring(Null), addstring(loctable!PassWord2)))
    aInsert = AddFlag(aInsert, "[STOP1]", IIf(loctable!Stop1, 1, 0))
    aInsert = AddFlag(aInsert, "[STOP2]", IIf(loctable!Stop2, 1, 0))
    aInsert = AddFlag(aInsert, "[ISRETURN]", IIf(loctable!IsReturn, 1, 0))
    cStr1 = cStr1 & addInsert(aInsert, "SUB_USER") & " ; "
    loctable.MoveNext
Loop
If cStr1 <> "" Then conShop2.Execute cStr1
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendSUB_USER = -1
GoTo lastsub
End Function
Private Function sendTransHeader() As Long
cFile = "FR1_60H"
    
cString = "SELECT * FROM FR1_60H WHERE ( ISNEW1 = 1 AND STORE1 = " & MyParn(pBranch) & " ) OR ( STORE2 = " & MyParn(pBranch) & " AND ISNEW2 = 1 )"

Dim loctable As New ADODB.Recordset
Dim loctable2 As New ADODB.Recordset
Dim cStr1 As String, cStr2 As String, cDelStr1 As String, cDelStr2 As String, c
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
   
    cDelStr1 = " delete from file1_60  where doc_no = " & MyParn(loctable!doc_no)
    cDelStr2 = " delete from file1_60h where doc_no = " & MyParn(loctable!doc_no)
    
    conShop2.Execute cDelStr1
    conShop2.Execute cDelStr2
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
        aInsert = AddFlag(aInsert, "[STORE1]", addstring(loctable!STORE1))
        aInsert = AddFlag(aInsert, "[STORE2]", addstring(loctable!STORE2))
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
        cStr1 = addInsert(aInsert, "file1_60h")
        conShop2.Execute cStr1, nAffect
        
        cStr2 = ""
        Dim aInsert2 As Variant
        If loctable2.State = adStateOpen Then loctable2.Close
        loctable2.Open "select * FROM FR1_60 WHERE DOC_NO = " & MyParn(loctable!doc_no), con, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not loctable2.EOF
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
            aInsert2 = AddFlag(aInsert2, "COST", Val(loctable2!cost & ""))
            cStr2 = cStr2 & addInsert(aInsert2, "FILE1_60") & " ; "
            loctable2.MoveNext
        Loop
        If cStr2 <> "" Then conShop2.Execute cStr2
        If loctable!STORE1 = pBranch Then con.Execute " UPDATE FR1_60H SET ISNEW1 = 0 WHERE DOC_NO  = " & MyParn(loctable!doc_no)
        If loctable!STORE2 = pBranch Then con.Execute " UPDATE FR1_60H SET ISNEW2 = 0 WHERE DOC_NO  = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    sendTransHeader = sendTransHeader + nAffect
Loop
Inform "  ÕÊÌ·«   " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTransHeader = -1
GoTo lastsub
End Function
Private Function SendGrDiscout() As Long
'On Error GoTo myerror
Dim cFile As String
cString = "SELECT * FROM FILE0_90H "
Dim loctable As New ADODB.Recordset, cStr1 As String
Dim loctable2 As New ADODB.Recordset, nCountR As Double
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "DATE1", addDate(loctable!Date1))
    aInsert = AddFlag(aInsert, "DATE2", addDate(loctable!date2))
    aInsert = AddFlag(aInsert, "DATE", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "QTY1", Val(loctable!QTY1))
    aInsert = AddFlag(aInsert, "QTY2", Val(loctable!QTY2))
    aInsert = AddFlag(aInsert, "DISC", Val(loctable!disc))
    aInsert = AddFlag(aInsert, "ISSTOP", IIf(loctable!ISSTOP, 1, 0))
    
    conShop2.Execute " DELETE FROM  FILE0_90  WHERE DOC_NO = " & MyParn(loctable!doc_no)
    conShop2.Execute " DELETE FROM  FILE0_90H WHERE DOC_NO = " & MyParn(loctable!doc_no)
    conShop2.Execute addInsert(aInsert, "FILE0_90H")
    
    
    cStr2 = ""
    Dim aInsert2 As Variant
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open "SELECT * FROM FILE0_90 WHERE DOC_NO = " & MyParn(loctable!doc_no), con, adOpenStatic, adLockReadOnly, adCmdText
    nCountR = 0
    Do While Not loctable2.EOF
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
        aInsert2 = AddFlag(aInsert2, "MODELNO", addstring(loctable2!MODELNO))
        aInsert2 = AddFlag(aInsert2, "DISC2", Val(loctable2!disc2))
        cStr2 = cStr2 & addInsert(aInsert2, "FILE0_90") & " ; "
        nCountR = nCountR + 1
        If nCountR = 100 Then
            If cStr2 <> "" Then conShop2.Execute cStr2
            cStr2 = ""
            nCountR = 0
        End If
        loctable2.MoveNext
    Loop
    If cStr2 <> "" Then conShop2.Execute cStr2
    loctable.MoveNext
Loop
Inform " „Ã„Ê⁄«  Œ’„" & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
nRecord = -1
Err.Clear
GoTo lastsub
End Function


Private Function senTarget() As Long
'On Error GoTo myerror
Dim loctable As New ADODB.Recordset, cStr1 As String
loctable.Open "SELECT * FROM Q_FILE6_80  WHERE BRANCH = " & MyParn(pBranch), con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
cStr1 = ""
cStr2 = ""
conShop2.Execute " DELETE FROM FILE6_80"
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "BRANCH", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "DATE1", addDate(loctable!Date1))
    aInsert = AddFlag(aInsert, "DATE2", addDate(loctable!date2))
    aInsert = AddFlag(aInsert, "VALUE_TARGET", addvalue(loctable!VALUE_TARGET))
    cStr2 = cStr2 & addInsert(aInsert, "FILE6_80") & " ; "
    loctable.MoveNext
    senTarget = senTarget + nAffect
Loop
If cStr2 <> "" Then conShop2.Execute cStr2
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
senTarget = -1
GoTo lastsub
End Function
Private Function FixDiscDox(sBranch As String) As Long
Dim loctable As New ADODB.Recordset
Dim cStr1 As String, cStr2 As String, cDelStr1 As String, cDelStr2 As String, cStrNew As String, cStrPhone As String

conShop2.Execute " UPDATE FILE6_20H SET DOC_DISC = INV_NO WHERE DOC_DISC <> '' AND DOC_DISC IS NOT NULL AND INV_NO IS NOT NULL "
cString = "Select * from file6_20h  WHERE DOC_DISC IS NOT NULL AND YEAR(DATE) = 2022 "
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long, lSkip As Boolean
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    con.Execute " UPDATE FILE6_20H SET DOC_DISC = " & addstring(loctable!DOC_DISC) & " , NOTE_DISC = " & addstring(loctable!NOTE_DISC) & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + 1
Loop
Inform "  ŒœÌÀ —ﬁ„ „” ‰œ «·Œ’„ " & nAffectTotal
FixDiscDox = nAffectTotal
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
FixDiscDox = -1
End Function


