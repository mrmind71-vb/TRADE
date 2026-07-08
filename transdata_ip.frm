VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form transdata_ip 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕœÌÀ Ê ”Õ» »Ì«‰«  «·›—Ê⁄"
   ClientHeight    =   11025
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   9300
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
   Picture         =   "transdata_ip.frx":0000
   RightToLeft     =   -1  'True
   ScaleHeight     =   11025
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
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
      Left            =   6075
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   9315
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
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   8460
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
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   8145
      Width           =   915
   End
   Begin VB.CheckBox XREC 
      Alignment       =   1  'Right Justify
      Caption         =   "XREC"
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
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   7830
      Width           =   825
   End
   Begin VB.CheckBox XGRDISC 
      Alignment       =   1  'Right Justify
      Caption         =   "„Ã„Ê⁄«  «·Œ’„"
      Height          =   330
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   8640
      Width           =   1770
   End
   Begin VB.CommandButton CMD_FIXPHONE 
      Caption         =   "FIX PHONE"
      Height          =   420
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   9180
      Width           =   1770
   End
   Begin VB.TextBox xdatesales 
      Alignment       =   1  'Right Justify
      Height          =   390
      Left            =   540
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9675
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.CheckBox Check5 
      Alignment       =   1  'Right Justify
      Caption         =   " ÕœÌÀ  »«—ﬂÊœ„⁄œ·"
      Height          =   285
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   8235
      Width           =   1770
   End
   Begin VB.Frame Frame2 
      Caption         =   "FIX  IP"
      Height          =   870
      Left            =   540
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   10125
      Width           =   7710
      Begin VB.CommandButton getcash 
         Caption         =   "getcash"
         Height          =   375
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   315
         Width           =   1275
      End
      Begin VB.CommandButton Command3 
         Caption         =   "SND IP"
         Height          =   495
         Left            =   4365
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   270
         Width           =   1635
      End
      Begin VB.TextBox XIP 
         Alignment       =   2  'Center
         Height          =   465
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   270
         Width           =   4020
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   6750
      Width           =   2175
      Begin VB.CheckBox Check6 
         Alignment       =   1  'Right Justify
         Caption         =   "ISSTOP"
         Height          =   270
         Left            =   540
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1890
         Width           =   1365
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
         TabIndex        =   17
         Top             =   1530
         Width           =   1815
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
         TabIndex        =   11
         Top             =   1215
         Visible         =   0   'False
         Width           =   1815
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
         TabIndex        =   10
         Top             =   180
         Width           =   2040
      End
      Begin VB.CheckBox Check2 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ· «·«’‰«›"
         Height          =   285
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   945
         Width           =   1770
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "FIX BARCODE "
      Height          =   510
      Left            =   7155
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   9135
      Width           =   1770
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   " ÕœÌÀ «· ﬂÊÌœ"
      Height          =   330
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   7785
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Check1"
      Height          =   600
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ÕœÌÀ „»Ì⁄«  „‰ «·›—Ê⁄"
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
      Left            =   2160
      TabIndex        =   3
      Top             =   8460
      Width           =   3840
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
      Left            =   2160
      TabIndex        =   1
      Top             =   9090
      Width           =   3840
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "÷»ÿ „·› »Ì«‰«  ··›—Ê⁄"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   7785
      Width           =   3840
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7485
      Left            =   90
      TabIndex        =   2
      Top             =   225
      Width           =   6990
      _cx             =   12330
      _cy             =   13203
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
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   375
      Left            =   2430
      TabIndex        =   22
      Top             =   9675
      Visible         =   0   'False
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label xcost 
      Alignment       =   1  'Right Justify
      Height          =   4425
      Left            =   7245
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   315
      Width           =   1905
   End
End
Attribute VB_Name = "transdata_ip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lAddAll As Boolean
Dim con As New ADODB.Connection, pBranch As String
Dim conShop2 As New ADODB.Connection
Dim cFilesave As String, cStoreSend As String
Public strConShop2 As String
Dim cDataFolder As String, cDataFile As String
Dim dPostDate As String
Dim nBranchPost As Double
Private Function ToBranch() As Boolean
On Error GoTo myerror
Dim lMyPrice As Boolean, lMyPriceItem As Boolean
Dim cId_No1 As String, cId_No2 As String, cHead1 As String
'Inform "1"
'Fix_Ifs

If Check5.Value <> 0 Then
    FixItemDel
End If
If XREC.Value <> 0 Then
    FixRecDoc
    Exit Function
End If

If xhead.Value <> 0 Then
    cId_No1 = GetDesca("SELECT ID_NO1 FROM ADDRESS", con)
    cId_No2 = GetDesca("SELECT ID_NO2 FROM ADDRESS", con)
    cHead1 = GetDesca("SELECT HEAD1 FROM ADDRESS", con)
    conShop2.Execute " UPDATE ADDRESS SET ID_NO1 = " & addstring(cId_No1) & " , ID_NO2 = " & addstring(cId_No2) & " , HEAD1 = " & addstring(cHead1)
    ToBranch = True
    Exit Function
End If

lMyPrice = (GetDesca("SELECT ISOKAZ FROM BRANCH WHERE CODE = " & MyParn(pBranch), con) = "True")
lMyPriceItem = (GetDesca("SELECT ISOKAZMODEL FROM BRANCH WHERE CODE = " & MyParn(pBranch), con) = "True")

If lMyPrice Then
    conShop2.Execute " update FILE0_00 SET PRICE = 1 "
Else
    conShop2.Execute " update FILE0_00 SET PRICE = 0 "
End If

'conShop2.Execute " delete from model_dem where year(date) = 2018 "
'conShop2.Execute " delete from model_dem where month(date) = 1 "
'conShop2.Execute " delete from model_dem where month(date) = 2 "

If lMyPriceItem Then
    conShop2.Execute " update FILE0_00 SET OKAZ_ITEM2 = 1 "
Else
    conShop2.Execute " update FILE0_00 SET OKAZ_ITEM2 = 0 "
End If

Inform " «·«’‰«› "

'If pBranch = "03" Then
'    nRecordCount = sendf_item_2016
'End If

Me.Caption = " ITEM "
nRecordCount = sendItems

If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror

If lAddAll Then
    nRecordCount = SendPurchaseHeader("SELECT FILE7_20H.* , FILE0_40.BRANCH FROM FILE7_20H LEFT OUTER JOIN FILE0_40 ON FILE7_20H.STORE = FILE0_40.CODE WHERE STORE = " & MyParn(cBranchStore), "file7_20h", "file7_20")
Else
    nRecordCount = SendPurchaseHeader("SELECT FILE7_20H.* , FILE0_40.BRANCH FROM FILE7_20H LEFT OUTER JOIN FILE0_40 ON FILE7_20H.STORE = FILE0_40.CODE  WHERE DOC_NO  IN (SELECT DOC_NO FROM NEW7_20H WHERE BRANCH = " & MyParn(pBranch) & " ) ", "file7_20h", "file7_20")
End If

If lAddAll Then
    nRecordCount = SendPurchaseHeader("SELECT     FILE7_10H.* , FILE0_40.BRANCH FROM         FILE7_10H LEFT OUTER JOIN                       FILE0_40 ON FILE7_10H.STORE = FILE0_40.CODE  WHERE STORE = " & MyParn(cBranchStore), "file7_10h", "file7_10")
'    nRecordCount = SendPurchase("SELECT     FILE7_10.* , FILE0_40.BRANCH FROM  FILE7_10H INNER JOIN FILE7_10 ON FILE7_10H.DOC_NO = FILE7_10.DOC_NO LEFT OUTER JOIN FILE0_40 ON FILE7_10H.STORE = FILE0_40.CODE   ", "file7_10")
'    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  „—œÊœ «·„‘ —Ì« ", "»‰Ã«Õ" Else GoTo myerror
Else
    nRecordCount = SendPurchaseHeader("SELECT FILE7_10H.* , FILE0_40.BRANCH FROM FILE7_10H LEFT OUTER JOIN FILE0_40 ON FILE7_10H.STORE = FILE0_40.CODE  where DOC_NO  IN (SELECT DOC_NO FROM NEW7_10H WHERE BRANCH = " & MyParn(pBranch) & " ) ", "file7_10h", "file7_10")
'    nRecordCount = SendPurchase("SELECT     FILE7_10.* , FILE0_40.BRANCH FROM  FILE7_10H INNER JOIN FILE7_10 ON FILE7_10H.DOC_NO = FILE7_10.DOC_NO LEFT OUTER JOIN FILE0_40 ON FILE7_10H.STORE = FILE0_40.CODE   where ISNEW" & nBranchPost & "  = 1 ", "file7_10")
'    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  „—œÊœ «·„‘ —Ì« ", "»‰Ã«Õ" Else GoTo myerror
End If

SendSection

sendSUB_USER

Me.Caption = " ÕÊÌ·« 1"
nRecordCount = sendTransHeader

Me.Caption = "«· ÕÊÌ·«  «·Ê«—œ…"
nRecordCount = sendTransHeader2

nRecordCount = SendDiscount
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ «·Œ’Ê„«  ", "»‰Ã«Õ" Else GoTo myerror

nRecordCount = senSubDem
If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  Êﬁ› ÿ·»Ì«  ", "»‰Ã«Õ" Else GoTo myerror

If XGRDISC.Value <> 0 Then nRecordCount = SendGrDiscout
nRecordCount = senTarget
nRecordCount = NewPrice
If Check3.Value <> 0 Then
    nRecordCount = sendCode("FILE8_61", False)
    nRecordCount = Sendfact
    nRecordCount = SendGroup
    nRecordCount = SendMAN
    nRecordCount = sendCode("FILE0_81", True)
    nRecordCount = sendCode("FILE0_81", True)
    nRecordCount = sendCode("FILE3_50")
    nRecordCount = sendCode("FILE4_10")
    nRecordCount = sendCode("FILE4_50")
    nRecordCount = sendCode("FILE4_10")
    nRecordCount = sendCode("SHIP")
    conShop2.Execute " update file4_10 set [group] = '01'"
    conShop2.Execute " update branch set visa = null "

    nRecordCount = sendbox
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·Œ“‰", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = sendstores
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„Œ«“‰", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = sendvisa
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·›Ì“«", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = sendBANK
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  »‰Êﬂ ", "»‰Ã«Õ" Else GoTo myerror

    nRecordCount = sendmosm
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„Ê«”„ ", "»‰Ã«Õ" Else GoTo myerror
    
    nRecordCount = sendUSER
    If nRecordCount >= 0 Then Inform " „ ‰ﬁ· " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·„” Œœ„Ì‰ ", "»‰Ã«Õ" Else GoTo myerror
    
End If
conShop2.Execute " delete from model_dem where ITEM IN (SELECT ITEM FROM FILE1_11 WHERE [IN] > 0 AND (DATE >= CONVERT(DATETIME, '2019-03-01 00:00:00', 102))) "


ToBranch = True
'MsgBox " „ ‰ﬁ· «·»Ì«‰« "
ToBranch = True
Exit Function
ToBranch = False
myerror:
'MsgBox "·„ Ì „ ‰ﬁ· ﬂ«›… «·»Ì«‰«  »‰Ã«Õ"
Err.Clear
End Function

Private Sub cmd_allitem_Click()
If MsgBox(" „  ÕœÌÀ ﬂ· «·«’‰«› »«·ﬂ«„· ··›—Ê⁄", vbYesNo + vbDefaultButton2) = vbYes Then
    con.Execute " update file1_10 set isnew1 = 0 "
End If
End Sub

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
                    If MsgBox("Â–« «·«Œ Ì«— ”Ê› Ì „ ‰ﬁ· ﬂ· «·»Ì«‰«  ≈·Ï «·›—⁄", vbYesNo + vbDefaultButton2) = vbNo Then
                        lAddAll = False
                    End If
                End If
                
                strConShop2 = LoadConStringshop2(IpShop, cDataShop)
                nBranchPost = Val(.TextMatrix(nRow, 0))
                pBranch = .TextMatrix(nRow, 0)
                
                If openConShop(conShop2, strConShop2) = "ok" Then
                    Inform "  „ «·≈ ’«· »«·›—⁄  "
                    SubFixData
                    If Not ToBranch_COST() Then
'                        MsgBox "·„ Ì „ﬂ‰ «·‰Ÿ«„ „‰ ‰ﬁ· «·»Ì«‰«  ≈·Ï " & .TextMatrix(nRow, 1)
                        .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
                    Else
'                        MsgBox " „ﬂ‰ «·‰Ÿ«„ „‰ ‰ﬁ· «·»Ì«‰«  ≈·Ï " & .TextMatrix(nRow, 1)
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
    
'    If FIXDATA.Value = 0 Then
'        con.Execute " delete from NEW_ITEM  where  branch in ( SELECT CODE FROM BRANCH WHERE IPSQL IS NULL )"
'        con.Execute " delete from NEW1_60H  where  branch in ( SELECT CODE FROM BRANCH WHERE IPSQL IS NULL )"
'        con.Execute " delete from NEW7_20H  where  branch in ( SELECT CODE FROM BRANCH WHERE IPSQL IS NULL )"
'        con.Execute " delete from NEW7_10H  where  branch in ( SELECT CODE FROM BRANCH WHERE IPSQL IS NULL )"
'    End If
    
    For nRow = 1 To .Rows - 1
        If .TextMatrix(nRow, 2) <> "" Then
            If .TextMatrix(nRow, 2) <> "" And TurnValue(.TextMatrix(nRow, 4), "True", True) Then
                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbYellow
                IpShop = .TextMatrix(nRow, 2)
                cDataShop = .TextMatrix(nRow, 3)
                lAddAll = TurnValue(.TextMatrix(nRow, 5), "True", True)
                cStoreSend = GetDesca("SELECT CODE FROM FILE0_40 WHERE BRANCH = " & MyParn(.TextMatrix(nRow, 0)), con)
                
                If lAddAll Then
                    If MsgBox("Â–« «·«Œ Ì«— ”Ê› Ì „ ‰ﬁ· ﬂ· «·»Ì«‰«  ≈·Ï «·›—⁄", vbYesNo + vbDefaultButton2) = vbNo Then
                        lAddAll = False
                    End If
                End If
                
                strConShop2 = LoadConStringshop2(IpShop, cDataShop)
                pBranch = .TextMatrix(nRow, 0)
                nBranchPost = Val(pBranch)
                Inform " branch " & Val(pBranch)
                If openConShop(conShop2, strConShop2) = "ok" Then
                    Inform "  „ «·≈ ’«· »«·›—⁄  "
                    SubFixData
                    If FIXDATA.Value = 0 Then
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
Private Sub Command2_Click()
    Me.Caption = "„—«Ã⁄… «·»«—ﬂÊœ ›Ï «·›—Ê⁄ "
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
                    If Not FixBarcode_Barch() Then
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


Private Sub Form_Load()
    cDataFolder = App.Path & "\mdb"
    cDataFile = "data"
    XDISC.Visible = bSupermode
    CMD_SENPRICE.Visible = bopt2
    openCon con
    myload
    dPostDate = DateAdd("M", -1, Date)
    XREC.Visible = bSupermode
End Sub
Private Function CopyData() As Boolean
Dim fs As New FileSystemObject
On Error GoTo myerror
If fs.FileExists(cDataFolder & "\" & cDataFile & ".mdb") Then
    fs.DeleteFile cDataFolder & "\" & cDataFile & ".mdb"
End If
If fs.FileExists(cDataFolder & "\" & cDataFile & "_" & "blk.mdb") Then
    fs.CopyFile cDataFolder & "\" & cDataFile & "_" & "blk.mdb", cDataFolder & "\" & cDataFile & ".mdb"
End If
CopyData = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function
Private Function SendGroup() As Long
On Error GoTo myerror

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
            
    If cStr1 = "" Then
        cStr1 = addInsert(aInsert, cFile)
    Else
        cStr1 = cStr1 & "  ; " & addInsert(aInsert, cFile)
    End If
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
On Error GoTo myerror
Dim cFile As String, nCount As Double, cSubStore As String
Dim cStrPostAll As String, cStrPostAll_DEL As String, cStrDel1 As String
cFile = "FILE1_10"
cSubStore = GetDesca("SELECT CODE FROM FILE0_40 WHERE BRANCH = " & MyParn(pBranch), con)
'If lAddAll Then
'    cString = " SELECT * FROM FILE1_10 AS FILE1_10_2 WHERE  FILE1_10_2.ITEM IN (  SELECT FILE1_10.ITEM  FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE FILE1_11.STORE = " & MyParn(cSubStore) & " )  or isnoitem =  1 "
    cString = "  SELECT * FROM FILE1_10 WHERE item in ( select item from NEW_ITEM where branch = " & MyParn(pBranch) & ")"

'Else
 '   cString = " SELECT * FROM FILE1_10 AS FILE1_10_2 WHERE FILE1_10_2.ITEM IN (SELECT ITEMNEW FROM BARCODE_EDIT) OR FILE1_10_2.ITEM IN ( SELECT        FILE1_10.ITEM  FROM   FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE FILE1_11.STORE = " & MyParn(cSubStore) & "  and file1_10.item in ( select item from NEW_ITEM where branch = " & MyParn(pBranch) & "  ) GROUP BY FILE1_10.ITEM  ) or isnoitem =  1 "
'End If

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
     
    If IsNull(loctable!color) Then
        aInsert = AddFlag(aInsert, "[COLOR]", addstring("-"))
    Else
        aInsert = AddFlag(aInsert, "[COLOR]", addstring(loctable!color))
    End If
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
    
    cStrPostAll_DEL = cStrPostAll_DEL & "  DELETE FROM FILE1_10 where item = " & loctable!Item & " ; "
    cStrDel1 = cStrDel1 & "  DELETE FROM NEW_ITEM WHERE ITEM = " & loctable!Item & " AND BRANCH = " & MyParn(pBranch) & " ; "
    cStrPostAll = cStrPostAll & addInsert(aInsert, "FILE1_10") & " ; "
    If nCount = 100 Then
        If cStrPostAll_DEL <> "" Then conShop2.Execute cStrPostAll_DEL
        If cStrPostAll <> "" Then conShop2.Execute cStrPostAll
        If cStrDel1 <> "" Then con.Execute cStrDel1
        cStrDel1 = ""
        cStrPostAll = ""
        cStrPostAll_DEL = ""
        nCount = 0
    End If
    Me.Caption = loctable!Item
    loctable.MoveNext
    sendItems = sendItems + 1
Loop
    Inform " »œ¡ ‰ﬁ· «·«’‰«› "
    If cStrPostAll_DEL <> "" Then conShop2.Execute cStrPostAll_DEL
    If cStrPostAll <> "" Then conShop2.Execute cStrPostAll
    If cStrDel1 <> "" Then con.Execute cStrDel1
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
If cStrPostAll <> "" Then conShop2.Execute cStrPostAll
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
Private Function sendstores()
On Error GoTo myerror
cFile = "FILE0_40"
cString = "SELECT * FROM FILE0_40 "

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
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "branch", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "SERVER", addstring(loctable!SERVER))
    aInsert = AddFlag(aInsert, "USERNAME", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "PASSWORD", addstring(loctable!PassWord))
    aInsert = AddFlag(aInsert, "ONLINE", IIf(loctable!online, 1, 0))
    If cStr1 = "" Then
        cStr1 = addInsert(aInsert, cFile)
    Else
        cStr1 = cStr1 & " ; " & addInsert(aInsert, cFile)
    End If
    loctable.MoveNext
    sendstores = sendstores + nAffect
Loop
conShop2.Execute " delete from file0_40"
conShop2.Execute cStr1
Inform " Rec No : " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendstores = -1
GoTo lastsub
End Function
Private Function sendvisa()
On Error GoTo myerror
cFile = "VISA"
cString = "SELECT * FROM VISA WHERE BRANCH  = '00' OR BRANCH = " & MyParn(pBranch)
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop2.Execute " delete from visa "
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "BRANCH", addstring(pBranch))
    
    aInsert = AddFlag(aInsert, "ISSTOP", IIf(loctable!ISSTOP, 1, 0))
    conShop2.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendvisa = sendvisa + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendvisa = -1
GoTo lastsub
End Function

Private Function sendClients() As Long
On Error GoTo myerror
Dim cFile As String
cFile = "FILE3_10"
cString = "SELECT * FROM " & cFile
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop2.Execute " delete from file3_10"
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "MANAGER", addstring(loctable!Manager))
    aInsert = AddFlag(aInsert, "ADDRESS", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "PHONE1", addstring(loctable!PHONE1))
    aInsert = AddFlag(aInsert, "FAX", addstring(loctable!Fax))
    aInsert = AddFlag(aInsert, "F_BALANCE", Val(loctable!F_Balance & ""))
    aInsert = AddFlag(aInsert, "[GROUP]", addstring(loctable!Group))
    aInsert = AddFlag(aInsert, "[F_DATE]", addstring(Format(loctable!F_DATE, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[EMAIL]", addstring(loctable!email))
    aInsert = AddFlag(aInsert, "[REMARK]", addstring(loctable!remark))
    aInsert = AddFlag(aInsert, "DISC", Val(loctable!disc & ""))
    aInsert = AddFlag(aInsert, "[CASH]", IIf(loctable!CASH, 1, 0))
           
    conShop2.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendClients = sendClients + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendClients = -1
GoTo lastsub
End Function
Private Function SendPurchaseHeader(cString, cFile, cFile2) As Long
On Error GoTo myerror
Dim loctable As New ADODB.Recordset
Dim loctable2 As New ADODB.Recordset, cStr1 As String, cStr2 As String, cStrDel1 As String, cStrDel2 As String, cNew1 As String, cNew2 As String
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    cStrDel1 = " delete from " & cFile & " where doc_no = " & MyParn(loctable!doc_no)
    cStrDel2 = " delete from " & cFile2 & " where doc_no = " & MyParn(loctable!doc_no)
    conShop2.Execute cStrDel1
    conShop2.Execute cStrDel2
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "CODE", addstring(loctable!CODE))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
        aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
        aInsert = AddFlag(aInsert, "[TAX]", Val(loctable!tax & ""))
        aInsert = AddFlag(aInsert, "[MOSM]", addstring(loctable!MOSM))
        aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))

        aInsert = AddFlag(aInsert, "[USERACCEPT]", addstring(loctable!UserAccept))
        aInsert = AddFlag(aInsert, "[TIMEACCEPT]", addDate(loctable!TimeAccept))
        aInsert = AddFlag(aInsert, "[Received ]", IIf(loctable!Received, 1, 0))

'       cStr1 = addInsert(aInsert, cFile)
        conShop2.Execute addInsert(aInsert, cFile), nAffect
        
        cStr2 = ""
        Dim aInsert2 As Variant
        If loctable2.State = adStateOpen Then loctable2.Close
        loctable2.Open "SELECT " & cFile2 & ".* FROM " & cFile2 & " WHERE DOC_NO = " & MyParn(loctable!doc_no), con, adOpenStatic, adLockReadOnly, adCmdText
        Do While Not loctable2.EOF
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
            aInsert2 = AddFlag(aInsert2, "PRICE", Val(loctable2!price & ""))
'            conShop2.Execute addInsert(aInsert2, cFile2), nAffect
            cStr2 = cStr2 & addInsert(aInsert2, cFile2) & " ; "
            loctable2.MoveNext
        Loop
        If cStr2 <> "" Then conShop2.Execute cStr2
    End If
    If UCase(cFile) = "FILE7_20H" Then
        cNew1 = " DELETE FROM NEW7_20H WHERE DOC_NO = " & MyParn(loctable!doc_no) & " AND BRANCH = " & MyParn(pBranch)
    Else
        cNew1 = " DELETE FROM NEW7_10H WHERE DOC_NO = " & MyParn(loctable!doc_no) & " AND BRANCH = " & MyParn(pBranch)
    End If
    con.Execute cNew1
    loctable.MoveNext
    SendPurchaseHeader = SendPurchaseHeader + nAffect
Loop

'conShop2.CommandTimeout = 600

If UCase(cFile) = "FILE7_20H" Then
    con.Execute " DELETE FROM NEW7_20H WHERE BRANCH = " & MyParn(pBranch)
Else
    con.Execute " DELETE FROM NEW7_10H WHERE BRANCH = " & MyParn(pBranch)
End If
Inform cFile2 & " : " & nRecord

lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendPurchaseHeader = -1
GoTo lastsub
End Function
Private Function SendPurchase(cString, cFile) As Long
On Error GoTo myerror
Dim loctable As New ADODB.Recordset

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100

    If loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "ITEM", addvalue(loctable!Item))
        aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant & ""))
        aInsert = AddFlag(aInsert, "PRICE", Val(loctable!price & ""))
        conShop2.Execute addInsert(aInsert, cFile), nAffect
    End If
    loctable.MoveNext
    SendPurchase = SendPurchase + nAffect
Loop
'con.Execute " update " & cFile & "H  SET ISNEW" & nBranchPost & "  = 0 "
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SendPurchase = -1
GoTo lastsub
End Function
Private Function sendTransHeader() As Long
cFile = "FILE1_60H"

If lAddAll Then
    cString = "SELECT  FILE0_40.branch AS branch_1 ,FILE0_40_1.branch AS BRANCH , file1_60h.* FROM         file1_60h INNER JOIN   FILE0_40 ON file1_60h.store1 = FILE0_40.CODE INNER JOIN FILE0_40 AS FILE0_40_1 ON file1_60h.store2 = FILE0_40_1.CODE "
Else
    cString = "SELECT  FILE0_40.branch AS branch_1 ,FILE0_40_1.branch AS BRANCH , file1_60h.* FROM         file1_60h INNER JOIN   FILE0_40 ON file1_60h.store1 = FILE0_40.CODE INNER JOIN FILE0_40 AS FILE0_40_1 ON file1_60h.store2 = FILE0_40_1.CODE WHERE DOC_NO IN (SELECT DOC_NO FROM NEW1_60H WHERE BRANCH = " & MyParn(pBranch) & " ) "
End If

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
    If loctable!BRANCH_1 = pBranch Or loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "DD-MM-YYYY")))
        aInsert = AddFlag(aInsert, "[STORE1]", addstring(loctable!STORE1))
        aInsert = AddFlag(aInsert, "[STORE2]", addstring(loctable!STORE2))
        aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
        aInsert = AddFlag(aInsert, "[USERACCEPT]", addstring(loctable!UserAccept))
        aInsert = AddFlag(aInsert, "[TIMEACCEPT]", addDate(loctable!TimeAccept))
        aInsert = AddFlag(aInsert, "[Received ]", IIf(loctable!Received, 1, 0))
        
        cStr1 = addInsert(aInsert, cFile)
        conShop2.Execute cStr1, nAffect
        
        cStr2 = ""
        Dim aInsert2 As Variant
        If loctable2.State = adStateOpen Then loctable2.Close
        loctable2.Open "select FILE1_60.* from file1_60 INNER JOIN FILE1_10 ON FILE1_60.ITEM = FILE1_10.ITEM where doc_no = " & MyParn(loctable!doc_no), con, adOpenStatic, adLockReadOnly, adCmdText
        If loctable2.RecordCount > 0 Then loctable2.MoveFirst
        Do While Not loctable2.EOF
            aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
            aInsert2 = AddFlag(aInsert2, "ITEM", addvalue(loctable2!Item))
            aInsert2 = AddFlag(aInsert2, "QUANT", Val(loctable2!Quant & ""))
            aInsert2 = AddFlag(aInsert2, "COST", Val(loctable2!cost & ""))
            cStr2 = cStr2 & addInsert(aInsert2, "FILE1_60") & " ; "
            loctable2.MoveNext
        Loop
        If cStr2 <> "" Then conShop2.Execute cStr2
    End If
    cNew1 = " DELETE FROM NEW1_60H WHERE DOC_NO = " & MyParn(loctable!doc_no) & " AND BRANCH = " & MyParn(pBranch)
    con.Execute cNew1
    loctable.MoveNext
    sendTransHeader = sendTransHeader + nAffect
Loop
con.Execute " DELETE FROM NEW1_60H WHERE BRANCH = " & MyParn(pBranch)
Inform "  ÕÊÌ·«   " & nAffect
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTransHeader = -1
GoTo lastsub
End Function
Private Function sendTrans() As Long
On Error GoTo myerror
cFile = "FILE1_60"
If lAddAll Then
    cString = "SELECT  FILE1_60.* , FILE0_40.branch AS BRANCH_1 , FILE0_40_1.branch  AS BRANCH FROM  FILE1_60 INNER JOIN file1_60h ON FILE1_60.DOC_NO = file1_60h.doc_no INNER JOIN  FILE0_40 ON file1_60h.store1 = FILE0_40.CODE INNER JOIN FILE0_40 AS FILE0_40_1 ON file1_60h.store2 = FILE0_40_1.CODE   "
Else
    cString = "SELECT  FILE1_60.* , FILE0_40.branch AS BRANCH_1 , FILE0_40_1.branch  AS BRANCH FROM  FILE1_60 INNER JOIN file1_60h ON FILE1_60.DOC_NO = file1_60h.doc_no INNER JOIN  FILE0_40 ON file1_60h.store1 = FILE0_40.CODE INNER JOIN FILE0_40 AS FILE0_40_1 ON file1_60h.store2 = FILE0_40_1.CODE   where ISNEW" & nBranchPost & "  = 1 "
End If

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    Me.Caption = " ÕÊÌ·« "
    If loctable!BRANCH_1 = pBranch Or loctable!branch = pBranch Then
        Me.Caption = loctable!doc_no
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "ITEM", addvalue(loctable!Item))
        aInsert = AddFlag(aInsert, "QUANT", Val(loctable!Quant & ""))
        aInsert = AddFlag(aInsert, "COST", Val(loctable!cost & ""))
        conShop2.Execute addInsert(aInsert, cFile), nAffect
    End If
    con.Execute " UPDATE FILE1_60H SET ISNEW" & nBranchPost & "  = 0 where doc_no = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    sendTrans = sendTrans + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendTrans = -1
GoTo lastsub
End Function
Private Function sendSTOCK() As Long
cFile = "FILE0_10"
If lAddAll Then
    cString = "SELECT   FILE0_10.* , FILE0_40.BRANCH  FROM  FILE0_10 INNER JOIN file0_10h ON FILE0_10.DOC_NO = file0_10h.doc_no INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE0_10H.STORE WHERE FILE0_10H.branch IS NULL "
Else
    cString = "SELECT   FILE0_10.* , FILE0_40.BRANCH  FROM  FILE0_10 INNER JOIN file0_10h ON FILE0_10.DOC_NO = file0_10h.doc_no INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE0_10H.STORE where FILE0_10H.branch IS NULL  AND isnew" & nBranchPost & " = 1  "
End If
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    If loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "[ITEM]", addstring(loctable!Item))
        aInsert = AddFlag(aInsert, "[RealBal]", Val(loctable!RealBal & ""))
        aInsert = AddFlag(aInsert, "[ComputerBal]", Val(loctable!ComputerBal & ""))
        aInsert = AddFlag(aInsert, "[Differ]", Val(loctable!Differ & ""))
        conShop2.Execute addInsert(aInsert, cFile), nAffect
        conShop2.Execute addInsert(aInsert, "FILE0_100"), nAffect
    End If
    loctable.MoveNext
    sendSTOCK = sendSTOCK + nAffect
Loop
con.Execute " UPDATE FILE0_10H set isnew" & nBranchPost & "  = 0 "
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendSTOCK = -1
GoTo lastsub
End Function

Private Function sendmosm() As Long
On Error GoTo myerror
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
    If cStr1 = "" Then
        cStr1 = addInsert(aInsert, cFile)
    Else
        cStr1 = cStr1 & " ; " & addInsert(aInsert, cFile)
    End If
    loctable.MoveNext
    sendmosm = sendmosm + nAffect
Loop

conShop2.Execute " delete from mosm "
conShop2.Execute cStr1

lastsub:
prog1.Visible = False

'If lokaz Then
'    conShop2.Execute " update file0_00 set price = 1 "
'Else
'    conShop2.Execute " update file0_00 set price = 0 "
'End If

Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendmosm = -1
GoTo lastsub
End Function
Private Function sendbox()
On Error GoTo myerror
cFile = "FILE0_50"
cString = "SELECT * FROM FILE0_50 where branch = " & MyParn(pBranch)

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop2.Execute " delete from file0_50"
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[F_DATE]", addDate(loctable!F_DATE))
    aInsert = AddFlag(aInsert, "[ISVISA]", IIf(loctable!ISVISA, 1, 0))
    aInsert = AddFlag(aInsert, "[BRANCH]", addDate(loctable!branch))
    aInsert = AddFlag(aInsert, "[type]", addstring(loctable!Type))
    aInsert = AddFlag(aInsert, "[ISBOXONLINE]", IIf(loctable!ISBOXONLINE, 1, 0))
    conShop2.Execute addInsert(aInsert, cFile), nAffect
    loctable.MoveNext
    sendbox = sendbox + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendbox = -1
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
            
    If cStr1 = "" Then
        cStr1 = addInsert(aInsert, cFile)
    Else
        cStr1 = cStr1 & " ; " & addInsert(aInsert, cFile)
    End If
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
Private Function SendSUPL() As Long
On Error GoTo myerror
Dim cFile As String
cString = "SELECT CODE , DESCA , DISC2 FROM FILE4_10 "
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
            
    If cStr1 = "" Then
        cStr1 = addInsert(aInsert, cFile)
    Else
        cStr1 = cStr1 & " ; " & addInsert(aInsert, cFile)
    End If
    loctable.MoveNext
Loop
conShop2.Execute " delete from FILE4_10 "
conShop2.Execute cStr1
Inform " „Ê—œÌ‰ " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
SendSUPL = -1
Err.Clear
GoTo lastsub
End Function

Private Sub myload()
openCon con
Dim loctable As New ADODB.Recordset, cString As String
cString = "SELECT * FROM branch  WHERE ipsql IS NOT NULL AND (CODE <> '00')  ORDER BY CODE"

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
'        If openConShop(conShop) = "ok" Then
            .TextMatrix(.Rows - 1, 2) = loctable!IPsql & ""
            .TextMatrix(.Rows - 1, 3) = loctable!data & ""
            .TextMatrix(.Rows - 1, 4) = False
'        Else
'            .TextMatrix(.Rows - 1, 4) = False
'        End If
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
        If openConShop(conShop2, strConShop2) = "ok" Then
            Inform " „ «·≈ ’«· ›—⁄ : " & .TextMatrix(nRow, 1)
            .Cell(flexcpBackColor, nRow, 1, nRow, .Cols - 1) = &HE0E0E0
            XCOST.Caption = ""
            lAddAll = TurnValue(.TextMatrix(nRow, 5), "True", True)
            If Not getData(.TextMatrix(nRow, 0)) Then
                Inform "„‘ﬂ·… «À‰«¡ ”Õ» »Ì«‰«  ›—⁄ : " & .TextMatrix(nRow, 1)
                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
            Else
                Inform " „ ”Õ» »Ì«‰«  ›—⁄ : " & .TextMatrix(nRow, 1), " »‰Ã«Õ", 1200
                .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbGreen
            End If
        Else
            .Cell(flexcpBackColor, nRow, 1, nRow, .Cols - 1) = vbRed
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
'If MsgBox(" ÕœÌÀ «·„»Ì⁄« ", vbYesNo) = vbYes Then
    nRecordCount = GetSales(sBranch)
    If nRecordCount >= 0 Then
        Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ Õ—ﬂ«  «·»Ì⁄", "»‰Ã«Õ"
    Else
        GoTo myerror
    End If
'End If

Me.Caption = "«·Ã—œ"
'If MsgBox(" ÕœÌÀ «·Ã—œ ", vbYesNo) = vbYes Then
    nRecordCount = getStock(sBranch)
    If nRecordCount >= 0 Then
        Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ „” ‰œ«  «·Ã—œ ", "»‰Ã«Õ"
    Else
        GoTo myerror
    End If
'End If
Me.Caption = "≈” ·«„ „” ‰œ« "

''''''''''''''''''''''''
'   Õ–› »⁄œ «” ·«„ «·„ÊœÌ·«  »«·‰Ÿ«„ «·”«»ﬁ
''''''''''''''''''''''''
'nRecordCount = Get_Purch_R("FILE7_20H")
'nRecordCount = Get_Purch_R("FILE7_10H")
'nRecordCount = Get_Purch_R("FILE1_60H")
''''''''''''''''''''''''

Me.Caption = "≈Ì—«œ"
nRecordCount = getCharge(sBranch)
If nRecordCount >= 0 Then
    Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ Õ—ﬂ«  ≈Ì—«œ ", "»‰Ã«Õ"
Else
    GoTo myerror
End If


Me.Caption = " ÕÊÌ·«  ‰ﬁœÌ…"
'If sBranch = "03" Then
    nRecordCount = GetTransBox(sBranch)
'End If
Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰  ÕÊÌ·«  ‰ﬁœÌ… ", "»‰Ã«Õ"

getData = True
Exit Function
myerror:
End Function
Private Function GetSales(sBranch As String) As Long
cFile = "FILE6_20"
Set loctable = New ADODB.Recordset
Dim cStr1 As String, cStr2 As String, cDelStr1 As String, cDelStr2 As String, cStrNew As String, cStrPhone As String


If IsDate(xdatesales.text) Then
    conShop2.Execute " UPDATE FILE6_20H SET ISNEW = 1 WHERE DATE = " & MyParn(xdatesales.text)
End If
'conShop2.Execute " update file6_20h set codevisa = '14' , isnew = 1 where point > 0 and codevisa is null "
cString = "Select * from file6_20h  WHERE (PRINTED = 1 and isnew = 1)  "
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long, lSkip As Boolean
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
    cStr2 = ""
    Me.Caption = " „»Ì⁄«  " & loctable!doc_no & "  " & loctable!Date
    cDelStr1 = " delete from file6_20 where doc_no = " & MyParn(loctable!doc_no)
    cDelStr2 = " delete from file6_20H where doc_no = " & MyParn(loctable!doc_no)
    
    con.Execute cDelStr1
    con.Execute cDelStr2
    
'    If cDelStr1 = "" Then
'    Else
'        cDelStr1 = cDelStr1 & " ; delete from file6_20 where doc_no = " & MyParn(loctable!doc_no)
'        cDelStr2 = cDelStr2 & " ; delete from file6_20H where doc_no = " & MyParn(loctable!doc_no)
'    End If
    
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
    aInsert = AddFlag(aInsert, "[BOXVISA]", addstring(loctable!BOXVISA))
    aInsert = AddFlag(aInsert, "[PAY]", Val(loctable!PAY & ""))
    aInsert = AddFlag(aInsert, "[VISA]", Val(loctable!Visa & ""))
    aInsert = AddFlag(aInsert, "[LATE]", Val(loctable!late & ""))
    aInsert = AddFlag(aInsert, "[REST]", Val(loctable!rest & ""))
    aInsert = AddFlag(aInsert, "[MAN]", addstring(loctable!MAN))
    aInsert = AddFlag(aInsert, "[PRINTED]", IIf(loctable!PRINTED, 1, 0))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
    aInsert = AddFlag(aInsert, "[phone]", addstring(loctable!phone))
                                                                                
    aInsert = AddFlag(aInsert, "[username_disc]", addstring(loctable!username_disc))
    aInsert = AddFlag(aInsert, "[username_ret]", addstring(loctable!username_ret))
    aInsert = AddFlag(aInsert, "[SALES_RET]", addstring(loctable!SALES_RET))
    aInsert = AddFlag(aInsert, "CODEVISA", addvalue(loctable!CODEVISA))
    aInsert = AddFlag(aInsert, "[NOTE_DISC]", addstring(loctable!NOTE_DISC))
    aInsert = AddFlag(aInsert, "[DOC_DISC]", addstring(loctable!DOC_DISC))
    aInsert = AddFlag(aInsert, "[VISADOC]", addstring(loctable!VISADOC))
    aInsert = AddFlag(aInsert, "[isbankahly]", IIf(loctable!isbankahly, 1, 0))
    
    aInsert = AddFlag(aInsert, "VISA2", addvalue(loctable!VISA2))
    aInsert = AddFlag(aInsert, "point", Val(loctable!Point & ""))
    aInsert = AddFlag(aInsert, "TYPEVISA", addvalue(loctable!TYPEVISA))
    
    aInsert = AddFlag(aInsert, "CHARGE1", Val(loctable!CHARGE1 & ""))
    aInsert = AddFlag(aInsert, "CHARGE2", Val(loctable!charge2 & ""))
    aInsert = AddFlag(aInsert, "[NAME]", addstring(loctable!Name))
    aInsert = AddFlag(aInsert, "[SHIP]", addstring(loctable!SHIP))
    aInsert = AddFlag(aInsert, "[ADDRESS]", addstring(loctable!Address))
    aInsert = AddFlag(aInsert, "[SHIP_NO]", addstring(loctable!SHIP_NO))
    aInsert = AddFlag(aInsert, "[DATE_PAY]", addDate(loctable!DATE_PAY))
    aInsert = AddFlag(aInsert, "[BANK]", addstring(loctable!BANK))
    aInsert = AddFlag(aInsert, "[INV_NO]", addstring(loctable!INV_NO))
    aInsert = AddFlag(aInsert, "[ONLINE]", addvalue(loctable!online))
    aInsert = AddFlag(aInsert, "[note_ret]", addstring(loctable!note_ret))
    aInsert = AddFlag(aInsert, "[ONLINE_DOC]", addstring(loctable!ONLINE_DOC))
    aInsert = AddFlag(aInsert, "[NORET_CHARGE1]", IIf(loctable!NORET_CHARGE1, 1, 0))
    
    aInsert = AddFlag(aInsert, "[card_disc]", addstring(loctable!BANK))
    aInsert = AddFlag(aInsert, "[card_desca]", addstring(loctable!card_desca))
    aInsert = AddFlag(aInsert, "[card_VALUE]", addvalue(loctable!card_VALUE))
    aInsert = AddFlag(aInsert, "[card_DOC]", addstring(loctable!card_DOC))
    
    
    aInsert = AddFlag(aInsert, "[VISA3]", Val(loctable!VISA3 & ""))
    aInsert = AddFlag(aInsert, "[CODEVISA3]", addstring(loctable!CODEVISA3))
    aInsert = AddFlag(aInsert, "[VISADOC3]", addstring(loctable!VISADOC3))
    aInsert = AddFlag(aInsert, "[VISA3_2]", Val(loctable!VISA3_2 & ""))
    aInsert = AddFlag(aInsert, "[TYPEVISA3 ]", addstring(loctable!TYPEVISA3))
    
    
    con.Execute addInsert(aInsert, "file6_20h")
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
            If cStr2 = "" Then
                cStr2 = addInsert(aInsert2, "file6_20")
            Else
                cStr2 = cStr2 & " ; " & addInsert(aInsert2, "file6_20")
            End If
        Else
            Inform " „—«Ã⁄… „»Ì⁄«  »«—ﬂÊœ " & sitem
            XCOST.Caption = XCOST.Caption & Chr(13) & sBranch & " - " & sitem
            lSkip = True
        End If
        loctable2.MoveNext
    Loop
    If cStr2 <> "" Then con.Execute cStr2
    If Not lSkip Then
        conShop2.Execute " update file6_20h set isnew = 0 where doc_no = " & MyParn(loctable!doc_no)
    End If
    cStr2 = ""
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop

'If cDelStr1 <> "" Then
'    CON.Execute cDelStr1
'    CON.Execute cDelStr2
'End If
'If cStr1 <> "" Then
'    CON.Execute cStr1
'    CON.Execute cStr2
'End If
'If cStrNew <> "" Then
'    conShop2.Execute cStrNew
'End If

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
    If cStrPhone = "" Then
        cStrPhone = addInsert(aInsert, "SUBCUST")
    Else
        cStrPhone = cStrPhone & " ; " & addInsert(aInsert, "SUBCUST")
    End If
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
.TextMatrix(0, 0) = "ﬂÊœ «·›—⁄"
.TextMatrix(0, 1) = "≈”„ «·›—⁄"
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
Private Function sendStockHeader() As Long
cFile = "FILE0_10H"

If lAddAll Then
    cString = "SELECT  FILE0_10H.* , FILE0_40.BRANCH FROM  file0_10h INNER JOIN FILE0_40 ON FILE0_10H.STORE = FILE0_40.CODE WHERE FILE0_10H.branch IS NULL "
Else
    cString = "SELECT  FILE0_10H.* , FILE0_40.BRANCH FROM  file0_10h INNER JOIN FILE0_40 ON FILE0_10H.STORE = FILE0_40.CODE where FILE0_10H.branch IS NULL AND isnew" & nBranchPost & "  = 1 "
End If
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    
    conShop2.Execute " delete from file0_10 where doc_no = " & MyParn(loctable!doc_no)
    conShop2.Execute " delete from file0_100 where doc_no = " & MyParn(loctable!doc_no)
    conShop2.Execute " delete from file0_10h where doc_no = " & MyParn(loctable!doc_no)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!branch = pBranch Then
'       Me.Caption = loctable!doc_no & ""
        aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
        aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
        aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
        aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
        aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
'       aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
        conShop2.Execute addInsert(aInsert, cFile), nAffect
    End If
    loctable.MoveNext
    sendStockHeader = sendStockHeader + nAffect

Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendStockHeader = -1
GoTo lastsub
End Function

Private Sub getcash_Click()
    If InputBox("", "") = "112233" Then
        ToMain_cash
    End If
End Sub

Private Sub grid1_EnterCell()
With grid1
    If .Col >= 4 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub

Private Function sendBALstores()
On Error GoTo myerror
cString = "SELECT * FROM Q_BAL1"

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop2.Execute " delete from BAL_1"
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "ITEM", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "BAL1", Val(loctable!BAL1))
    conShop2.Execute addInsert(aInsert, "BAL_1"), nAffect
    loctable.MoveNext
    sendBALstores = sendBALstores + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendBALstores = -1
GoTo lastsub
End Function
Private Function Get_Purch_R(pFile) As Long
Set loctable = New ADODB.Recordset

cString = "Select * from " & pFile & " WHERE Received = 1 "
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Dim aInsert2 As Variant

Do Until loctable.EOF
    con.Execute " UPDATE " & pFile & " SET Received = 1 , TIMEACCEPT = " & addDate(loctable!TimeAccept) & " , USERACCEPT = " & addstring(loctable!UserAccept) & " where doc_no = " & MyParn(loctable!doc_no)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
Get_Purch_R = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
Get_Purch_R = -1
GoTo lastsub
End Function
Private Function GetTransBox(pBranch) As Long
Set loctable = New ADODB.Recordset
Dim cDoc As String

If IsDate(xdatesales.text) Then
    conShop2.Execute " UPDATE FILE0_51 SET ISNEW = 1 WHERE DATE >= " & DateSq(xdatesales.text)
    Exit Function
End If
cString = "Select * from FILE0_51 WHERE isnew = 1"

loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount

Dim aInsert As Variant
Do Until loctable.EOF
    cDoc = pBranch & loctable!CODE
    con.Execute " delete from file0_52  where CODE = " & MyParn(cDoc)
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "CODE", addstring(cDoc))
    aInsert = AddFlag(aInsert, "NO1", addstring(loctable!no1))
    aInsert = AddFlag(aInsert, "NO2", addstring(loctable!no2))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[VALUE]", Val(loctable!Value))
    aInsert = AddFlag(aInsert, "BRANCH", addstring(pBranch))
                                                                                
    con.Execute addInsert(aInsert, "FILE0_52"), nAffect
    conShop2.Execute " update FILE0_51 set isnew = 0 where CODE = " & MyParn(loctable!CODE)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
GetTransBox = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GetTransBox = -1
GoTo lastsub
End Function
Private Function getStock(pBranch) As Long
cFile = "FILE0_10H"
Dim loctable2 As New ADODB.Recordset
Dim cStr1 As String, cStr2 As String, cStr3 As String, cStrDel1 As String, cStrDel2 As String, cStrDel3 As String, cNew As String
cString = "SELECT  FILE0_10H.* FROM  file0_10h where isnew  = 1 AND BRANCH = " & MyParn(pBranch)
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText
Dim nRecCount As Double
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

Do Until loctable.EOF
    cStrDel1 = ""
    cStrDel2 = ""
    cStrDel3 = ""
    cStr1 = ""
    cStr2 = ""
    cStr3 = ""
'    If cStrDel1 = "" Then
'        cStrDel1 = " delete from file0_10 where doc_no = " & MyParn(loctable!doc_no)
'        cStrDel2 = " delete from file0_100 where doc_no = " & MyParn(loctable!doc_no)
'        cStrDel3 = " delete from file0_10h where doc_no = " & MyParn(loctable!doc_no)
'    Else
'        cStrDel1 = cStrDel1 & "  ; delete from file0_10 where doc_no = " & MyParn(loctable!doc_no)
'        cStrDel2 = cStrDel2 & "  ; delete from file0_100 where doc_no = " & MyParn(loctable!doc_no)
'        cStrDel3 = cStrDel3 & "  ; delete from file0_10h where doc_no = " & MyParn(loctable!doc_no)
'    End If
    
    con.Execute " delete from file0_10 where doc_no = " & MyParn(loctable!doc_no)
    con.Execute " delete from file0_100 where doc_no = " & MyParn(loctable!doc_no)
    con.Execute " delete from file0_10h where doc_no = " & MyParn(loctable!doc_no)
    
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(loctable!Date))
    aInsert = AddFlag(aInsert, "[STORE]", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "[DESCA]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[CLOSED]", IIf(loctable!CLOSED, 1, 0))
'   aInsert = AddFlag(aInsert, "[USERSEND]", addstring(loctable!USERSEND))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(loctable!branch))
    
    con.Execute addInsert(aInsert, "FILE0_10H")
    
    Set loctable2 = New ADODB.Recordset
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open "SELECT * FROM FILE0_10 WHERE DOC_NO = " & MyParn(loctable!doc_no), conShop2, adOpenStatic, adLockReadOnly, adCmdText
    nRecCount = 0
    cStr2 = ""
    Do Until loctable2.EOF
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
        aInsert2 = AddFlag(aInsert2, "[ITEM]", addstring(loctable2!Item))
        aInsert2 = AddFlag(aInsert2, "[RealBal]", Val(loctable2!RealBal & ""))
        aInsert2 = AddFlag(aInsert2, "[ComputerBal]", Val(loctable2!ComputerBal & ""))
        aInsert2 = AddFlag(aInsert2, "[Differ]", Val(loctable2!Differ & ""))

        cStr2 = cStr2 & addInsert(aInsert2, "file0_10") & " ; "
        If nRecCount = 200 Then
            con.Execute cStr2
            nRecCount = 0
            cStr2 = ""
        End If
        Me.Caption = loctable2!doc_no & " " & loctable2!Item
        loctable2.MoveNext
    Loop
    If cStr2 <> "" Then con.Execute cStr2
    
    Set loctable2 = New ADODB.Recordset
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open "SELECT * FROM FILE0_100 WHERE DOC_NO = " & MyParn(loctable!doc_no), conShop2, adOpenStatic, adLockReadOnly, adCmdText
    cStr3 = ""
    nRecCount = 0
    Do Until loctable2.EOF
        Me.Caption = loctable2!Item
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
        aInsert2 = AddFlag(aInsert2, "[ITEM]", addstring(loctable2!Item))
        aInsert2 = AddFlag(aInsert2, "[RealBal]", Val(loctable2!RealBal & ""))
        aInsert2 = AddFlag(aInsert2, "[ComputerBal]", Val(loctable2!ComputerBal & ""))
        aInsert2 = AddFlag(aInsert2, "[Differ]", Val(loctable2!Differ & ""))
        cStr3 = cStr3 & addInsert(aInsert2, "file0_100") & " ; "
        If nRecCount = 200 Then
            con.Execute cStr3
            nRecCount = 0
            cStr3 = ""
        End If
        
        Me.Caption = loctable2!doc_no & " " & loctable2!Item
        loctable2.MoveNext
    Loop
    If cStr3 <> "" Then con.Execute cStr3
        
    con.Execute " UPDATE FILE0_10H SET ISNEW = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    Inform " „” ‰œ Ã—œ" & loctable!doc_no
    
    conShop2.Execute " UPDATE FILE0_10H SET ISNEW = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    con.Execute " UPDATE FILE0_10H SET CLOSED = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    If loctable!CLOSED Then
        con.Execute " UPDATE FILE0_10H SET CLOSED = 1 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    End If
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
Private Function SendMAN() As Long
On Error GoTo myerror

Dim cFile As String, cStr1 As String, cStr2 As String
Dim loctable As New ADODB.Recordset, nRecord As Double

cFile = "FILE6_25"
cString = "SELECT * FROM FILE6_25 "
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If IsNull(loctable!branch) Or loctable!branch = pBranch Then
        aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
        aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
        aInsert = AddFlag(aInsert, "BRANCH", addstring(loctable!branch))
        aInsert = AddFlag(aInsert, "ID_EMP", addstring(loctable!ID_EMP))
        aInsert = AddFlag(aInsert, "isstop", IIf(loctable!ISSTOP, 1, 0))
        If cStr1 = "" Then
            cStr1 = " DELETE FROM FILE6_25 WHERE CODE =  " & MyParn(loctable!CODE)
            cStr2 = addInsert(aInsert, cFile)
        Else
            cStr1 = cStr1 & " ;  DELETE FROM FILE6_25 WHERE CODE =  " & MyParn(loctable!CODE)
            cStr2 = cStr2 & " ; " & addInsert(aInsert, cFile)
        End If
    End If
    loctable.MoveNext
    SendMAN = SendMAN + nAffect
Loop
    conShop2.Execute cStr1
    conShop2.Execute cStr2
    Inform " »«∆⁄Ì‰ " & nRecord
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
SendMAN = -1
Err.Clear
GoTo lastsub
End Function
Private Function NewPrice() As Long
On Error GoTo myerror
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
If cStr1 <> "" Then
    conShop2.Execute cStr2
    conShop2.Execute cStr1
    con.Execute cStr3
End If

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
Private Function SendDiscount() As Long
On Error GoTo myerror
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
Sub Fix_Ifs()
    Dim nBaR1 As Double, nBar2 As Double
    Dim IFFTABLE As New ADODB.Recordset
    IFFTABLE.Open "SELECT * FROM item_066 ", con, adOpenStatic, adLockReadOnly, adCmdText
    Inform "2"
    With IFFTABLE
        Do While Not .EOF
            nBaR1 = !ITEM_1
            nBar2 = !ITEM_2
            Inform !ITEM_1
            conShop2.Execute " UPDATE FILE6_20 SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBar2)
'            conShop2.Execute " UPDATE FILE7_10 SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBaR2)
'            conShop2.Execute " UPDATE FILE7_20 SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBaR2)
'            conShop2.Execute " UPDATE FILE1_60 SET ITEM = " & addvalue(nBaR1) & " WHERE ITEM = " & addvalue(nBaR2)
            conShop2.Execute " DELETE FROM FILE1_10 WHERE ITEM = " & nBar2, nRec
            If nRec > 0 Then Inform " DELETE  " & nBar2
            .MoveNext
        Loop
    End With
End Sub
Private Function getCharge(sBranch As String) As Long
cFile = "FILE8_60"
Dim loctable   As New ADODB.Recordset
Dim loctable2  As New ADODB.Recordset
'If XALL.Value = 1 Then
'    cString = "Select * from FILE8_60H "
'Else

'If IsDate(xdatesales.Text) Then
'    conShop2.Execute " UPDATE FILE8_60H SET ISNEW = 1 WHERE DATE >= " & DateSq(xdatesales.Text)
'    Exit Function
'End If

cString = "Select * from FILE8_60H WHERE ISNEW = 1 "


loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long
nRecordCount = loctable.RecordCount
    
'If XALL.Value = 1 Then
'    con.Execute " delete from file8_60 where DOC_NO IN ( SELECT DOC_NO FROM FILE8_60H WHERE BRANCH = " & MyParn(sBranch) & " ) "
'    con.Execute " delete from file8_60h where BRANCH = " & MyParn(sBranch)
'End If
Dim aInsert As Variant
Do Until loctable.EOF
    
'    If XALL.Value <> 1 Then
        con.Execute " delete from file8_60 where DOC_NO = " & MyParn(loctable!doc_no) & " AND DOC_NO IN ( SELECT DOC_NO FROM FILE8_60H WHERE BRANCH = " & MyParn(sBranch) & " ) "
        con.Execute " delete from file8_60h where DOC_NO = " & MyParn(loctable!doc_no) & " AND BRANCH = " & MyParn(sBranch)
'    End If
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DATE]", addDate(Format(loctable!Date, "dd-mm-yyyy")))
    aInsert = AddFlag(aInsert, "[USERNAME]", addstring(loctable!UserName))
    aInsert = AddFlag(aInsert, "[BRANCH]", addstring(sBranch))
    aInsert = AddFlag(aInsert, "ISCLOSED", 1)
    con.Execute addInsert(aInsert, cFile & "H"), nAffect

'''''''''''''
    Dim aInsert2 As Variant
    cString = "Select * FROM FILE8_60 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText
    Do Until loctable2.EOF
        aInsert2 = AddFlag(Empty, "DOC_NO", addstring(loctable2!doc_no))
        aInsert2 = AddFlag(aInsert2, "[BOX]", addstring(loctable2!BOX))
        aInsert2 = AddFlag(aInsert2, "[CHARGE]", addstring(loctable2!CHARGE))
        aInsert2 = AddFlag(aInsert2, "[DESCA]", addstring(loctable2!DESCA))
        aInsert2 = AddFlag(aInsert2, "[BRANCH]", addstring(loctable2!branch))
        aInsert2 = AddFlag(aInsert2, "[VALUE]", Val(loctable2!Value & ""))
        aInsert2 = AddFlag(aInsert2, "[ROW]", Val(loctable2!Row & ""))
        con.Execute addInsert(aInsert2, "FILE8_60"), nAffect
        loctable2.MoveNext
    Loop
'''''''''''''
    conShop2.Execute " UPDATE FILE8_60H SET ISNEW = 0 WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
getCharge = nAffectTotal
lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
getCharge = -1
GoTo lastsub
End Function
Function LoadConStringshop2(pServerIp, pServerData)
    Dim cServerName As String, cUserId As String, cPassword As String
    cUserId = "DATA_PRO" ' decrypt(RetSetting("userId", App.Path & "\conf.txt"), "dr")
    cPassword = "2015"
'    cPassword = decrypt(RetSetting("Password", App.Path & "\conf.txt"), "dr")
    LoadConStringshop2 = "provider=SQLOLEDB;data source=" & pServerIp & ";initial " _
                & "catalog=" & pServerData & ";user id = " & cUserId & ";" & "password = " & cPassword & ";Timeout=10"
End Function
Sub SubFixData()
    On Error Resume Next
    Dim cString As String

'    cString = " ALTER TABLE [dbo].[FILE0_90] ADD [DISC2] [int] NOT NULL CONSTRAINT [DF_FILE0_90_DISC2]  DEFAULT ((0) )"
'    createCommand cString, conShop2

'
'    cString = "ALTER TABLE [dbo].[FILE1_10]  ADD  [ISNOITEM2] [bit] NOT NULL CONSTRAINT [DF_FILE1_10_ISNOITEM2] DEFAULT ((0)) "
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE1_10SC] ADD [ISDISC] [bit] NULL"
'    createCommand cString, conShop2
''
'    cString = "ALTER TABLE [dbo].[DISCOUNT] ADD [ISDISC2] [bit] NULL"
'    createCommand cString, conShop2
''
''    cString = "ALTER TABLE [dbo].[FILE6_20H] ADD [NOTE_DISC] [nvarchar](50) NULL"
''    createCommand cString, conShop2
''
''    cString = "ALTER TABLE [dbo].[FILE6_20] ADD [ISDISC2_1] [bit]  NOT NULL CONSTRAINT [DF_FILE6_2_ISDISC2_1]  DEFAULT ((0)) "
''    createCommand cString, conShop2
''
''    cString = " create VIEW [dbo].[QFILE6_20] AS SELECT        FILE6_20H.DOC_NO, FILE6_20H.store, FILE6_20H.date, FILE6_20H.CASH, FILE6_20H.VISA, FILE6_20H.BOX, FILE6_20H.PAY, FILE6_20H.branch,                           FILE6_20H.username_disc, FILE6_20.ITEM, FILE6_20.QUANT, FILE6_20.PRICE, FILE6_20.TOTAL, FILE6_20H.discount, FILE6_20H.rate, FILE6_20.cost,                          FILE6_20.PRICE_C, FILE6_20.MAN, FILE6_20.ISDISC2_1, FILE6_20.S_OKAZ FROM            FILE6_20H INNER JOIN                          FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO    "
''    createCommand cString, conShop2
''
''    cString = " create VIEW [dbo].[FILE1_11_ALL] AS SELECT  * FROM FILE1_11"
''    createCommand cString, conShop2
''
''    createCommand cString, conShop2
''    cString = " create VIEW [dbo].[FILE1_11_ALL] AS SELECT FILE1_11.* FROM FILE1_11 inner join file0_40 on file0_40.code = file1_11.store where file0_40.ISSTOP = 0   UNION ALL SELECT * FROM FR1_11"
''
''    cString = "ALTER TABLE [dbo].[FILE4_10] ADD [disc2] [decimal](18, 2) NULL"
''    createCommand cString, conShop2
''
'    cString = "ALTER TABLE [dbo].[FILE1_10] ADD    [MODELSEX] [varchar](3) NULL , [MODELAGE] [varchar](3) NULL"
'    createCommand cString, conShop2
''
''    cString = "ALTER VIEW [dbo].[FILE1_10H] AS SELECT        MODEL, MAX(desca) AS DESCA, MAX([GROUP]) AS [GROUP], MAX(RATE) AS RATE, MAX(OKAZ) AS OKAZ, MAX(FACT) AS FACT, MAX(MOSM) AS MOSM,  MAX(modelno) AS MODELNO, MAX(MODELFACT) AS MODELFACT, MAX(code) AS code, MAX(SUPP) AS SUPP, MAX(MODELFACT0) AS MODELFACT0, REDEM,  FIXPRICE, SHOWSALES, ISOKAZITEM, SECTION, ISSTOP, OKAZ_2, MOSM2, ISNODEM, ISNOITEM, MODELSEX, MODELAGE FROM            dbo.FILE1_10 GROUP BY MODEL, REDEM, FIXPRICE, SHOWSALES, ISOKAZITEM, SECTION, ISSTOP, OKAZ_2, MOSM2, ISNODEM, ISNOITEM, MODELSEX, MODELAGE"
''    createCommand cString, conShop2
''
''    cString = "CREATE TABLE [dbo].[FILE0_81](  [CODE] [int] NOT NULL, [DESCA] [nvarchar](100) NULL, [ISDISC] [bit] NULL, CONSTRAINT [PK_FILE0_81] PRIMARY KEY CLUSTERED  ( [CODE] ASC )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ) ON [PRIMARY]"
''    createCommand cString, conShop2
''
''    cString = "CREATE TABLE [dbo].[FILE0_82](  [CODE] [int] NOT NULL, [DESCA] [nvarchar](100) NULL, [ISDISC] [bit] NULL, CONSTRAINT [PK_FILE0_82] PRIMARY KEY CLUSTERED  ( [CODE] ASC )WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY] ) ON [PRIMARY]"
''    createCommand cString, conShop2
''
''    cString = " ALTER VIEW [dbo].[QBRANCH_ALL] AS SELECT        code, desca, STORE FROM            dbo.branch    "
''    createCommand cString, conShop2
''
''    cString = " CREATE VIEW [dbo].[Q_ITEMCOST] AS SELECT        FILE7_20.ITEM, FILE7_20H.DATE, FILE7_20.TOTAL / FILE7_20.QUANT AS cost FROM            FILE7_20 INNER JOIN                          FILE7_20H ON FILE7_20.DOC_NO = FILE7_20H.DOC_NO"
''    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE1_10]  ADD  [ISBANKAHLY] [bit] NOT NULL CONSTRAINT [DF_FILE6_20H_ISBANKAHLY] DEFAULT ((0))  "
'    createCommand cString, conShop2
'
'    cString = "ALTER  TABLE [dbo].[FILE1_10]  ADD [BARCODE13] [nvarchar](50) NULL"
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[FILE6_20H] ADD [VISADOC] [nvarchar](20) NULL"
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
'    cString = " ALTER TABLE [dbo].[FILE1_10] ADD [barcode_gs1] [nvarchar](50) NULL  "
'    createCommand cString, conShop2
'
'    cString = " alter TABLE [dbo].[VISA] add    [ISSTOP] [bit] NOT NULL"
'    createCommand cString, conShop2
'    cString = " ALTER TABLE [dbo].[VISA] ADD  CONSTRAINT [DF_VISA_ISSTOP]  DEFAULT ((0)) FOR [ISSTOP]"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[users] ADD [isbranch] [bit] NOT NULL CONSTRAINT [DF_users_isbranch]  DEFAULT ((0))"
'    createCommand cString, conShop2
'
'    conShop2.Execute " UPDATE FILE6_20H SET SHIP = '0' + FILE6_20H.SHIP  WHERE LEN(SHIP) = 1 ", nRec
'
'
'    If Val(GetDesca("SELECT COUNT(*) FROM FILE0_90", conShop2) & "") = 0 Then
'        cString = " CREATE TABLE [dbo].[FILE0_90](  [DOC_NO] [varchar](6) NULL,     [MODELNO] [nvarchar](50) NULL,  [id] [int] IDENTITY(1,1) NOT NULL,      [ISSELECT] [bit] NOT NULL ) ON [PRIMARY]"
'        createCommand cString, conShop2
'    End If
'
'
'    cString = " CREATE TABLE [dbo].[FILE0_90H]( [DOC_NO] [varchar](6) NOT NULL, [DATE] [datetime] NULL, [DESCA] [nvarchar](250) NULL,   [QTY1] [int] NULL,  [QTY2] [int] NULL,  [DISC] [numeric](18, 2) NULL,   [ISSTOP] [bit] NOT NULL) ON [PRIMARY]"
'    createCommand cString, conShop2
'
'    cString = ReadFile(App.Path & "\TXT\Q_DATA_SEND.TXT")
'    createCommand cString, conShop2
'
'    cString = " CREATE TABLE [dbo].[FILE6_80]( [DOC_NO] [nvarchar](6) NULL, [DATE1] [smalldatetime] NULL, [DATE2] [smalldatetime] NULL, [BRANCH] [nvarchar](3) NULL, [VALUE_TARGET] [numeric](18, 2) NULL, [ID] [int] IDENTITY(1,1) NOT NULL, CONSTRAINT [PK_FILE6_80] PRIMARY KEY CLUSTERED  ( [ID] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
'    createCommand cString, conShop2
    
'    If Val(GetDesca("SELECT COUNT(*) FROM FILE0_90", conShop2) & "") = 0 Then
'        conShop2.Execute " DROP TABLE   FILE0_90 "
'        cString = " CREATE TABLE [dbo].[FILE0_90](  [DOC_NO] [varchar](6) NULL, [MODELNO] [nvarchar](50) NULL,  [id] [int] IDENTITY(1,1) NOT NULL) ON [PRIMARY]"
'        createCommand cString, conShop2
'    End If
    
'
'    cString = "ALTER TABLE [dbo].[FILE1_10] ADD [PRICE_P] [decimal](18, 2) NULL"
'    createCommand cString, conShop2
    
'    cString = ReadFile(App.Path & "\TXT\Q_FILE1_11.TXT")
'    createCommand cString, conShop2
        
'    cString = " ALTER TABLE [dbo].[FILE0_90] ADD    [DISC2] [int] NOT NULL CONSTRAINT [DF_FILE0_90_DISC2]  DEFAULT ((0)) "
'    createCommand cString, conShop2
    
'    cString = "     ALTER VIEW [dbo].[FILE1_11_ALL] AS SELECT        dbo.FILE1_11.TYPE, dbo.FILE1_11.DESCA, dbo.FILE1_11.doc_id, dbo.FILE1_11.codesup, dbo.FILE1_11.DATE, dbo.FILE1_11.STORE, dbo.FILE1_11.ITEM,                          dbo.FILE1_11.[IN], dbo.FILE1_11.OUT, dbo.FILE1_11.PRICE, dbo.FILE1_11.TOTAL, dbo.FILE1_11.DISCOUNT, dbo.FILE1_11.cost, dbo.FILE1_11.man " & _
'                " FROM  dbo.FILE1_11 INNER JOIN dbo.FILE1_10 ON dbo.FILE1_11.ITEM = dbo.FILE1_10.ITEM WHERE        (dbo.FILE1_10.ISNOITEM = 0) "
'    createCommand cString, conShop2
    
'    cString = ReadFile(App.Path & "\TXT\Q_FILE1_11.TXT")
'    createCommand cString, conShop2
'
'    cString = " ALTER VIEW [dbo].[FILE1_11_ALL] AS SELECT  * FROM FILE1_11    "
'    createCommand cString, conShop2
'
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
'    cString = " ALTER TABLE [dbo].[SUB_USER] ADD    [PASSWORD2] [nvarchar](50) NULL,    [STOP1] [bit]  NULL,     [STOP2] [bit]  NULL"
'    createCommand cString, conShop2
'
'    cString = " ALTER VIEW [dbo].[Q_DOC_SALES_DISC] AS SELECT        dbo.FILE6_20.ITEM, dbo.FILE6_20.QUANT, dbo.FILE6_20.PRICE, dbo.FILE0_90H.DOC_NO, dbo.FILE0_90H.QTY1, dbo.FILE0_90H.DISC, dbo.FILE6_20.DOC_NO AS DOC_SALES " & _
'                " FROM            dbo.FILE0_90H INNER JOIN dbo.FILE0_90 ON dbo.FILE0_90H.DOC_NO = dbo.FILE0_90.DOC_NO INNER JOIN dbo.FILE6_20 INNER JOIN dbo.FILE1_10 ON dbo.FILE6_20.ITEM = dbo.FILE1_10.ITEM ON dbo.FILE0_90.MODELNO = dbo.FILE1_10.modelno " & _
'                " WHERE        (dbo.FILE0_90H.ISSTOP = 0) AND (dbo.FILE0_90H.DATE1 <= CAST(GETDATE() AS DATE)) AND (dbo.FILE0_90H.DATE2 >= CAST(GETDATE() AS DATE)) "
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE6_20H] ADD   [DOC_AGE]  AS (datediff(day,[DATE],getdate())) "
'    createCommand cString, conShop2
'
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [note_ret] [nvarchar](100) NULL     "
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
'    cString = " alter TABLE [dbo].[FILE6_20H] add   [ONLINE_DOC] [nvarchar](10) NULL     "
'    createCommand cString, conShop2
    
    
'    cString = "  CREATE TABLE [dbo].[FILE6_90H]( [DOC_NO] [nvarchar](10) NOT NULL,   [E_MAIL] [nvarchar](50) NULL,   [SUBTOTAL] [numeric](18, 2) NULL,   [SHIPPING] [numeric](18, 2) NULL,   [TOTAL] [numeric](18, 2) NULL,  [DISCOUNT_CODE] [nvarchar](50) NULL,    [DISCOUNT] [numeric](18, 2) NULL, " & _
'            " hipping_Method] [nvarchar](50) NULL,  [DATE] [datetime] NULL, [NAME] [nvarchar](200) NULL,    [STREET] [nvarchar](1000) NULL, [ADDRESS1] [nvarchar](1000) NULL,   [ADDRESS2] [nvarchar](1000) NULL,   [CITY] [nvarchar](50) NULL, [PHONE] [nvarchar](50) NULL,    [Payment_Method] [nvarchar](50) NULL, " & _
'            " Shipping_City] [nvarchar](50) NULL, CONSTRAINT [PK_FILE6_90H] PRIMARY KEY CLUSTERED (  [DOC_NO] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY] "
'    createCommand cString, conShop2
'
'    cString = " CREATE TABLE [dbo].[FILE6_90](  [DOC_NO] [nvarchar](10) NULL,   [SKU] [nvarchar](50) NULL,  [ITEM] [int] NULL,  [ITEM_NAME] [nvarchar](200) NULL,   [PRICE] [numeric](18, 2) NULL,  [QUANT] [int] NULL, [TOTAL]  AS ([PRICE]*[QUANT]),  [ID] [int] IDENTITY(1,1) NOT NULL, CONSTRAINT [PK_FILE6_90] PRIMARY KEY CLUSTERED ( [ID] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]) ON [PRIMARY]"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE6_90]  WITH CHECK ADD  CONSTRAINT [FK_FILE6_90_FILE6_90H] FOREIGN KEY([DOC_NO]) REFERENCES [dbo].[FILE6_90H] ([DOC_NO]) ON UPDATE CASCADE"
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE6_90] CHECK CONSTRAINT [FK_FILE6_90_FILE6_90H]"
'    createCommand cString, conShop2
    
'    cString = " ALTER TABLE [dbo].[file6_20h] ADD [NORET_CHARGE1] [bit] NOT NULL CONSTRAINT [DF_file6_20H_NORET_CHARGE1] DEFAULT ((0))"
'    createCommand cString, conShop2
'
'    cString = ReadFile(App.Path & "\TXT\QBOXMOVE.TXT")
'    createCommand cString, conShop2
'
'    cString = "ALTER TABLE [dbo].[SUB_USER]       ADD  [ISRETURN] [bit] NOT NULL CONSTRAINT [DF_SUB_USER_ISRETURN] DEFAULT ((0)) "
'    createCommand cString, conShop2
'
'    cString = " ALTER TABLE [dbo].[FILE0_40] ADD   [ISNOGRDISC] [bit] NOT NULL CONSTRAINT [DF_FILE0_40_ISNOGRDISC]  DEFAULT ((0))     "
'    createCommand cString, conShop2
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
 
     cString = "ALTER VIEW [dbo].[QFILE6_20] AS SELECT        FILE6_20H.DOC_NO, FILE6_20H.store, FILE6_20H.date, FILE6_20H.CASH, FILE6_20H.VISA, FILE6_20H.BOX, FILE6_20H.PAY, FILE6_20H.branch, FILE6_20H.username_disc, FILE6_20.ITEM, FILE6_20.QUANT,(FILE6_20.PRICE ) AS PRICE, (FILE6_20.TOTAL ) AS TOTAL, FILE6_20H.discount, FILE6_20H.rate, FILE6_20.cost, FILE6_20.PRICE_C, FILE6_20.MAN, FILE6_20.ISDISC2_1,FILE6_20.S_OKAZ, FILE1_10.MODEL, FILE1_10.[GROUP], FILE1_10.code, FILE1_10.MOSM, FILE1_10.SECTION, 1 AS iSBRANCH, FILE6_20H.SALES_RET, FILE6_20H.username_ret, FILE6_20H.PHONE, FILE1_10.Outlet,FILE6_20H.DOC_DISC, FILE6_20H.CODE AS CUST, FILE6_20H.ISRET, FILE6_20H.ONLINE , SHIP , (FILE6_20H.cash + FILE6_20H.visa + FILE6_20H.visa3 + FILE6_20H.point + FILE6_20H.late )  as T_BOON   , FILE6_20.MAN2 " & _
            " FROM    FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM"
    createCommand cString, conShop2

 
    Err.Clear
End Sub
Private Function ToBranch_COST() As Boolean
On Error GoTo myerror
Dim lMyPrice As Boolean, lMyPriceItem As Boolean
Inform " aa " & grid1.TextMatrix(grid1.Row, 1)
Inform "Send All Price & BARCODE13 "

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
On Error GoTo myerror
Dim cFile As String, cString As String, cStrall As String, nCount As Double, cStrDel1 As String
cFile = "FILE1_10"
Inform "Select Item "
If Check6.Value <> 0 Then
'    cString = "SELECT * FROM FILE1_10 WHERE   ISNOITEM = 1 OR ISNOITEM2 = 1  "
    cString = "SELECT * FROM FILE1_10 WHERE   ISSTOP = 1  "
Else
    If XBAR13.Value <> 0 Then
        cString = "SELECT * FROM FILE1_10 WHERE  ( BARCODE13 IS NOT NULL OR BARCODE_GS1 IS NOT NULL ) AND ITEM IN (SELECT ITEM FROM FILE1_11 INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE WHERE  FILE0_40.BRANCH = " & MyParn(pBranch) & ")  "
    Else
        If Check2.Value <> 0 Then
            cString = "SELECT * FROM FILE1_10 WHERE  ITEM IN (SELECT ITEM FROM FILE1_11 INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE WHERE  FILE0_40.BRANCH = " & MyParn(pBranch) & ")  "
        Else
            cString = "SELECT * FROM FILE1_10 WHERE  ITEM IN (SELECT ITEM FROM FILE1_11 INNER JOIN FILE0_40 ON FILE1_11.STORE = FILE0_40.CODE WHERE  FILE0_40.BRANCH = " & MyParn(pBranch) & ")   AND ITEM IN ( SELECT ITEM FROM NEW_ITEM WHERE BRANCH = " & MyParn(pBranch) & " ) "
        End If
    End If
End If
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
'Inform cString
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
sCaption = Me.Caption
If Check6.Value <> 0 Then
    conShop2.Execute " UPDATE FILE1_10 SET ISSTOP = 0 WHERE ISSTOP = 1 "
End If
Do Until loctable.EOF
    nCount = nCount + 1
    nRecord = nRecord + 1
'   Me.Caption = loctable!Item
    prog1.Visible = True
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    If Check6.Value = 0 Then
        cStrall = cStrall & " ; " & " UPDATE FILE1_10 SET COSTITEM = " & Val(loctable!costITEM & "") & " , cost = " & Val(loctable!cost & "") & " , cost2 = " & Val(loctable!cost2 & "") & " , price = " & Val(loctable!price & "") & " , price2 = " & Val(loctable!PRICE2 & "") & " , price_2 = " & Val(loctable!PRICE_2 & "") & " , OKAZ = " & Val(loctable!okaz & "") & " , OKAZ_2 = " & Val(loctable!okaz_2 & "") & " , barcode13 = " & addstring(loctable!BARCODE13) & " , barcode_GS1 = " & addstring(loctable!BARCODE_GS1) & " , ISNOITEM = " & IIf(loctable!ISNOITEM, 1, 0) & " , ISNOITEM2 = " & IIf(loctable!ISNOITEM2, 1, 0) & " WHERE ITEM = " & Val(loctable!Item & "")
    Else
        cStrall = cStrall & " ; " & " UPDATE FILE1_10 SET ISSTOP = 1  WHERE ITEM = " & Val(loctable!Item & "")
    End If
    If Check2.Value = 0 And Check6.Value = 0 Then
        cStrDel1 = cStrDel1 & " ; " & " DELETE FROM NEW_ITEM WHERE ITEM = " & loctable!Item & " AND BRANCH = " & MyParn(pBranch)
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


Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendItems_COST = -1
GoTo lastsub
End Function
Private Function sendTransHeader2() As Long

If lAddAll Then
    cString = "SELECT file1_61h.* FROM  file1_61h INNER JOIN   FILE0_40 ON file1_61h.store1 = FILE0_40.CODE INNER JOIN FILE0_40 AS FILE0_40_1 ON file1_61h.store2 = FILE0_40_1.CODE WHERE ( FILE0_40_1.BRANCH  = " & MyParn(pBranch) & " OR FILE0_40.BRANCH  = " & MyParn(pBranch) & " ) "
Else
    cString = "SELECT file1_61h.* FROM  file1_61h INNER JOIN   FILE0_40 ON file1_61h.store1 = FILE0_40.CODE INNER JOIN FILE0_40 AS FILE0_40_1 ON file1_61h.store2 = FILE0_40_1.CODE WHERE ISNEW = 1 AND ( FILE0_40_1.BRANCH  = " & MyParn(pBranch) & " OR FILE0_40.BRANCH  = " & MyParn(pBranch) & " ) "
End If

Dim loctable As New ADODB.Recordset
Dim cStr1 As String, cStr2 As String
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount

FixSubShop1

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[TYPEDOC]", addstring(loctable!typedoc))
    aInsert = AddFlag(aInsert, "[DOC_TRANS]", addstring(loctable!DOC_TRANS))
    aInsert = AddFlag(aInsert, "[NOTES2]", addstring(loctable!NOTES2))
    cStr1 = cStr1 & addUpdate(aInsert, "FILE1_61H", " DOC_NO = " & MyParn(loctable!doc_no))
    cStr2 = cStr2 & "UPDATE FILE1_61H SET ISNEW = 0 " & " ; "
    loctable.MoveNext
    sendTransHeader2 = sendTransHeader2 + nAffect
Loop
If cStr1 <> "" Then conShop2.Execute cStr1
If cStr2 <> "" Then con.Execute cStr2
Inform "  ÕÊÌ·«  Ê«—œ… " & sendTransHeader2
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

cString = " alter table [dbo].[DISCOUNT] add   [MAXDISC] [decimal](18, 2) NULL"
createCommand cString, conShop2

Err.Clear
End Sub
Private Function SendSection() As Long
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
On Error GoTo myerror
cString = "SELECT * FROM BARCODE_EDIT "
Me.Caption = " ⁄œÌ· »«—ﬂÊœ"
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount


''''''''''''''''''''''''

''''    conShop2.Execute " UPDATE FILE6_20  SET ITEM = 94685 WHERE PRICE < 100 AND ITEM =94684 "
''''    conShop2.Execute " UPDATE FILE7_10  SET ITEM = 94685 WHERE PRICE < 100 AND ITEM =94684 "
''''    conShop2.Execute " UPDATE FILE7_20  SET ITEM = 94685 WHERE PRICE < 100 AND ITEM =94684 "
''''    conShop2.Execute " UPDATE FILE1_60  SET ITEM = 94685 WHERE COST < 100 AND ITEM =94684 "
''''    conShop2.Execute " UPDATE FILE1_61  SET ITEM = 94685 WHERE COST < 100 AND ITEM =94684 "
''''
''''    conShop2.Execute " UPDATE FILE0_10  SET ITEM = 94685 WHERE COST < 100 AND ITEM =94684 "
''''    conShop2.Execute " UPDATE FILE0_100 SET ITEM = 94685 WHERE COST < 100 AND ITEM =94684 "
''''
''''    If pBranch = "04" Or pBranch = "03" Then
''''        conShop2.Execute " UPDATE FILE1_11_ALL   SET ITEM = 94685 WHERE COST  < 100 AND ITEM = 94684 "
''''    End If
''''
''''''''''''''''''''''''

Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    If loctable!ITEMOLD = 59759 Then
        A = A
    End If
    Me.Caption = loctable!ITEMOLD
    
    If GetDesca("select item from file1_11 where item   = " & loctable!ITEMOLD, conShop2) <> "" Then
        conShop2.Execute " UPDATE FILE6_20  SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        conShop2.Execute " UPDATE FILE7_10  SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        conShop2.Execute " UPDATE FILE7_20  SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        conShop2.Execute " UPDATE FILE1_60  SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        conShop2.Execute " UPDATE FILE1_61  SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        
        conShop2.Execute " UPDATE FILE0_10  SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        conShop2.Execute " UPDATE FILE0_100 SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
    
        If pBranch = "04" Or pBranch = "03" Then
            conShop2.Execute " UPDATE FILE1_11_ALL   SET ITEM = " & addvalue(loctable!ITEMNEW) & " WHERE ITEM = " & (loctable!ITEMOLD)
        End If
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
Private Function FixBarcode_Barch() As Boolean
On Error GoTo myerror
Dim lMyPrice As Boolean, lMyPriceItem As Boolean
Inform " aa " & grid1.TextMatrix(grid1.Row, 1)
Inform "Fix Barcode "

nRecordCount = SubFixBarcode
If nRecordCount >= 0 Then Inform " „  ⁄œÌ· «”⁄«— " & nRecordCount & " ”Ã· „‰ »Ì«‰«  «·«’‰«›", "»‰Ã«Õ" Else GoTo myerror

FixBarcode_Barch = True
Inform " „ ‰ﬁ· «·»Ì«‰« "
FixBarcode_Barch = True
Exit Function
FixBarcode_Barch = False
myerror:
MsgBox "·„ Ì „ ‰ﬁ· ﬂ«›… «·»Ì«‰«  »‰Ã«Õ"
Err.Clear
End Function
Private Function SubFixBarcode() As Long
On Error GoTo myerror
Dim cFile As String, cString As String, cStrall As String, nCount As Double, cStrDel1 As String
cFile = "FILE1_10"
Inform " Fix Barcode "
cString = "SELECT * FROM FILE1_10 "
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText
'Inform cString
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
sCaption = Me.Caption
Do Until loctable.EOF
    If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & loctable!Item, con) = "" Then
        If GetDesca("SELECT ITEM FROM FILE6_20 WHERE ITEM = " & loctable!Item, conShop2) = "" Then
            conShop2.Execute " DELETE FROM FILE1_10 WHERE ITEM =  " & loctable!Item
            nCount = nCount + 1
        Else
            XCOST.Caption = XCOST.Caption & Chr(13) & pBranch & " - " & loctable!Item
        End If
    End If
    nRecord = nRecord + 1
    prog1.Visible = True
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    Me.Caption = loctable!Item
    loctable.MoveNext
Loop
lastsub:
Me.Caption = sCaption
prog1.Visible = False
SubFixBarcode = nCount
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SubFixBarcode = -1
GoTo lastsub
End Function


Private Sub CMD_FIXPHONE_Click()
    Me.Caption = "„—«Ã⁄… «·»«—ﬂÊœ ›Ï «·›—Ê⁄ "
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
                    If Not FixPhoneSales() Then
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
Private Function FixPhoneSales() As Boolean
On Error GoTo myerror
Dim lMyPrice As Boolean, lMyPriceItem As Boolean
Inform " aa " & grid1.TextMatrix(grid1.Row, 1)
Inform "Fix Barcode "

nRecordCount = SubFixPhone
If nRecordCount >= 0 Then Inform " „  ⁄œÌ· «—ﬁ«„  ·Ì›Ê‰«  " & nRecordCount, "»‰Ã«Õ" Else GoTo myerror

FixPhoneSales = True
Inform " „  ÕœÌÀ «—ﬁ«„ «· ·Ì›Ê‰«  "
Exit Function
FixPhoneSales = False
myerror:
MsgBox "·„ Ì „ ‰ﬁ· ﬂ«›… «·»Ì«‰«  »‰Ã«Õ"
Err.Clear
End Function


Private Function SubFixPhone() As Long
On Error GoTo myerror
Dim cFile As String, cString As String, cStrall As String, nCount As Double, cStrDel1 As String
Inform " Fix Barcode "
cString = "SELECT * FROM FILE6_20H WHERE PHONE IS NOT NULL "
Dim loctable As New ADODB.Recordset
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText

'Inform cString
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long, sCaption As String
nRecordCount = loctable.RecordCount
sCaption = Me.Caption
cStr1 = ""
nCount = 0
Do Until loctable.EOF
    cStr1 = cStr1 & " UPDATE FILE6_20H SET PHONE = " & addstring(loctable!phone) & " WHERE DOC_NO = " & MyParn(loctable!doc_no) & " ; "
    nCount = nCount + 1
    nRecord = nRecord + 1
    prog1.Visible = True
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    Me.Caption = loctable!doc_no
    If nCount = 200 Then
        If cStr1 <> "" Then con.Execute cStr1
        cStr1 = ""
        nCount = 0
    End If
    loctable.MoveNext
Loop
If cStr1 <> "" Then con.Execute cStr1
lastsub:
Me.Caption = sCaption
prog1.Visible = False
SubFixPhone = nRecord
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
SubFixPhone = -1
GoTo lastsub
End Function
Private Function senSubDem() As Long
On Error GoTo myerror
Dim loctable As New ADODB.Recordset, cStr1 As String
loctable.Open "SELECT * FROM MODEL_STOP WHERE BRANCH = " & MyParn(sStore), con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
cStr1 = ""
cStr2 = ""
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "MODEL", addstring(loctable!MODEL))
    aInsert = AddFlag(aInsert, "BRANCH", addstring(loctable!branch))
   cStr1 = cStr1 & " DELETE FROM MODEL_STOP WHERE MODEL = " & MyParn(loctable!MODEL) & " ; "
    cStr2 = cStr2 & addInsert(aInsert, "MODEL_STOP") & " ; "
    loctable.MoveNext
    senSubDem = senSubDem + nAffect
Loop
If cStr1 <> "" Then conShop2.Execute cStr1
If cStr2 <> "" Then conShop2.Execute cStr2
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
senSubDem = -1
GoTo lastsub
End Function

Private Function sendSUB_USER()
On Error GoTo myerror
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

Private Function sendBANK()
On Error GoTo myerror
cString = "SELECT * FROM FILE5_10 WHERE ISSHIP =1 "
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop2.Execute " delete from FILE5_10 "
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    aInsert = AddFlag(Empty, "CODE", addstring(loctable!CODE))
    aInsert = AddFlag(aInsert, "DESCA", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "F_DATE", addDate(loctable!F_DATE))
    conShop2.Execute addInsert(aInsert, "FILE5_10"), nAffect
    loctable.MoveNext
    sendBANK = sendBANK + nAffect
Loop
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendBANK = -1
GoTo lastsub
End Function
Private Function sendf_item_2016() As Long
cString = "SELECT * FROM f_item_2016 where store = '003' "
Dim loctable As New ADODB.Recordset, cStr1 As String
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
conShop2.Execute " delete from f_item_2016 "
Dim aInsert As Variant
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "item", addstring(loctable!Item))
    aInsert = AddFlag(aInsert, "store", addstring(loctable!STORE))
    aInsert = AddFlag(aInsert, "bal_all", Val(loctable!bal_all))
    cStr1 = cStr1 & addInsert(aInsert, "f_item_2016") & " ; "
    loctable.MoveNext
    sendf_item_2016 = sendf_item_2016 + nAffect
Loop
conShop2.Execute cStr1
lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendf_item_2016 = -1
GoTo lastsub
End Function


Private Function ToMain_cash() As Boolean
On Error GoTo myerror
With grid1
For nRow = 1 To .Rows - 1
    If .TextMatrix(nRow, 2) <> "" And TurnValue(.TextMatrix(nRow, 4), "True", True) Then
        IpShop = .TextMatrix(nRow, 2)
        cDataShop = .TextMatrix(nRow, 3)
        strConShop2 = LoadConStringshop2(IpShop, cDataShop)
        openConShop conShop2, strConShop2
        .Cell(flexcpBackColor, nRow, 1, nRow, .Cols - 1) = &HE0E0E0
        XCOST.Caption = ""
        lAddAll = TurnValue(.TextMatrix(nRow, 5), "True", True)
        If Not getData_cash(.TextMatrix(nRow, 0)) Then
            MsgBox "„‘ﬂ·… «À‰«¡ ‰ﬁ· „·›«  ›—⁄ : " & .TextMatrix(nRow, 1)
            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbRed
        Else
            Inform " „ ‰”Œ „·›«  ›—⁄ : " & .TextMatrix(nRow, 1), " »‰Ã«Õ", 1200
            .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = vbGreen
        End If
    End If
Next
End With
ToMain_cash = True
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
End Function



Private Function getData_cash(sBranch As String) As Boolean
Dim nRecordCount As Long

Me.Caption = "„»Ì⁄« "
nRecordCount = GetSales_cash(sBranch)
If nRecordCount >= 0 Then
    Inform " „ ”Õ» " & nRecordCount & " ”Ã· „‰ Õ—ﬂ«  «·»Ì⁄", "»‰Ã«Õ"
Else
    GoTo myerror
End If
getData_cash = True
Exit Function
myerror:
End Function


Private Function GetSales_cash(sBranch As String) As Long
Dim nCount As Double, nYear As Double, nMonth As Double
Set loctable = New ADODB.Recordset
nYear = Val(InputBox("", "") & "")
nMonth = Val(InputBox("", "") & "")
If nMonth = 0 Then
    cString = "Select * from file6_20h   where year(date) = " & nYear
Else
    cString = "Select * from file6_20h   where year(date) = " & nYear & " AND MONTH(date) = " & nMonth
End If
loctable.Open cString, conShop2, adOpenStatic, adLockReadOnly, adCmdText

prog1.Value = 0
prog1.Visible = True

Dim nRecordCount As Long, nRecord As Long, nAffect As Long, nAffectTotal As Long, lSkip As Boolean
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Dim aInsert2 As Variant
cStr1 = ""

Do Until loctable.EOF
    Me.Caption = " „»Ì⁄«  " & loctable!doc_no & "  " & loctable!Date
    nCount = nCount + 1
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    aInsert = AddFlag(Empty, "DOC_NO", addstring(loctable!doc_no))
    aInsert = AddFlag(aInsert, "[DISCOUNT]", Val(loctable!discount & ""))
    aInsert = AddFlag(aInsert, "[CASH]", Val(loctable!CASH & ""))
    aInsert = AddFlag(aInsert, "[VISA]", Val(loctable!Visa & ""))
    aInsert = AddFlag(aInsert, "VISA2", addvalue(loctable!VISA2))
    aInsert = AddFlag(aInsert, "point", Val(loctable!Point & ""))
    aInsert = AddFlag(aInsert, "TYPEVISA", addvalue(loctable!TYPEVISA))
    aInsert = AddFlag(aInsert, "[DATE_PAY]", addDate(loctable!DATE_PAY))
    cStr1 = cStr1 & addUpdate(aInsert, "file6_20h", " DOC_NO = " & MyParn(loctable!doc_no)) & " ; "
'    con.Execute addUpdate(aInsert, "file6_20h", " DOC_NO = " & MyParn(loctable!doc_no))
    If nCount = 500 Then
        con.Execute cStr1
        nCount = 0
        cStr1 = ""
    End If
    loctable.MoveNext
    nAffectTotal = nAffectTotal + nAffect
Loop
If cStr1 <> "" Then con.Execute cStr1

lastsub:
loctable.Close
Set loctable = Nothing
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
GetSales_cash = -1
GoTo lastsub
End Function
Private Function sendUSER()
On Error GoTo myerror
cString = "SELECT * FROM USERS WHERE ISBRANCH = 1 "
Dim loctable As New ADODB.Recordset, cCodeUser As String
Dim loctable2 As New ADODB.Recordset
Dim loctable3 As New ADODB.Recordset

loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
loctable3.Open "menu", con, adOpenStatic, adLockReadOnly, adCmdTable
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Dim aInsert As Variant
Dim aInsert2 As Variant

conShop2.Execute "DELETE FROM USERS WHERE isbranch = 1 "
conShop2.Execute " DELETE FROM menuSetting WHERE CODE in ( SELECT CODE FROM USERS WHERE isbranch = 1  ) "
'conShop2.Execute " DELETE FROM menu"

Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    
    cCodeUser = Val(loctable!CODE) + 100
    aInsert = AddFlag(Empty, "[password]", addstring(loctable!PassWord))
    aInsert = AddFlag(aInsert, "[CODE]", addstring(cCodeUser))
    aInsert = AddFlag(aInsert, "[desca]", addstring(loctable!DESCA))
    aInsert = AddFlag(aInsert, "[option1]", IIf(loctable!Option1, 1, 0))
    aInsert = AddFlag(aInsert, "[option2]", IIf(loctable!Option2, 1, 0))
    aInsert = AddFlag(aInsert, "[option3]", IIf(loctable!Option3, 1, 0))
    aInsert = AddFlag(aInsert, "[option4]", IIf(loctable!Option4, 1, 0))
    aInsert = AddFlag(aInsert, "[option5]", IIf(loctable!Option5, 1, 0))
    aInsert = AddFlag(aInsert, "[option6]", IIf(loctable!Option6, 1, 0))
    aInsert = AddFlag(aInsert, "[option7]", IIf(loctable!Option7, 1, 0))
    aInsert = AddFlag(aInsert, "[option8]", IIf(loctable!Option8, 1, 0))
    aInsert = AddFlag(aInsert, "[option9]", IIf(loctable!Option9, 1, 0))
    aInsert = AddFlag(aInsert, "[option10]", IIf(loctable!Option10, 1, 0))
    
    aInsert = AddFlag(aInsert, "SUPPER", IIf(loctable!SUPPER, 1, 0))
    aInsert = AddFlag(aInsert, "isbranch", 1)
    
    conShop2.Execute "DELETE FROM USERS WHERE code = " & addstring(cCodeUser)
    conShop2.Execute " DELETE FROM menuSetting WHERE code = " & addstring(cCodeUser)
    
    conShop2.Execute addInsert(aInsert, "USERS"), nAffect
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open "SELECT * FROM menuSetting WHERE CODE = " & MyParn(loctable!CODE), con, adOpenStatic, adLockReadOnly, adCmdText
    Do While Not loctable2.EOF
        aInsert2 = AddFlag(Empty, "[code]", addstring(cCodeUser))
        aInsert2 = AddFlag(aInsert2, "[CONTROL]", addstring(loctable2!Control))
        aInsert2 = AddFlag(aInsert2, "[VISIBLE]", IIf(loctable2!Visible, 1, 0))
        aInsert2 = AddFlag(aInsert2, "[EDITABLE]", IIf(loctable2!Editable, 1, 0))
        conShop2.Execute addInsert(aInsert2, "menuSetting")
        loctable2.MoveNext
    Loop
    loctable.MoveNext
    sendUSER = sendUSER + nAffect
Loop

'Do Until loctable3.EOF
'    con.Execute "insert into menu(control,Desca,menuNo,[order],mainmenu)" & _
'                   " values(" & _
'                   addstring(loctable3!Control) & "," & _
'                   addstring(loctable3!Desca) & "," & _
'                   addvalue(loctable3!MenuNo) & "," & _
'                   addvalue(loctable3![Order]) & "," & _
'                   addstring(loctable3!mainmenu) & _
'                    ")"
'    loctable3.MoveNext
'Loop

lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
sendUSER = -1
GoTo lastsub
End Function
Private Function SendGrDiscout() As Long
On Error GoTo myerror
Dim cFile As String
cString = "SELECT * FROM FILE0_90H "
Dim loctable As New ADODB.Recordset, cStr1 As String
Dim loctable2 As New ADODB.Recordset
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
    
    
    nCountR = 0
    cStr2 = ""
    Dim aInsert2 As Variant
    If loctable2.State = adStateOpen Then loctable2.Close
    loctable2.Open "SELECT * FROM FILE0_90 WHERE DOC_NO = " & MyParn(loctable!doc_no), con, adOpenStatic, adLockReadOnly, adCmdText
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
On Error GoTo myerror
Dim loctable As New ADODB.Recordset, cStr1 As String, cSendStore As String, nCountR As Double
cSendStore = GetDesca("SELECT STORE FROM BRANCH WHERE CODE = " & MyParn(pBranch), con)
loctable.Open "SELECT * FROM Q_FILE6_80  WHERE BRANCH = " & MyParn(cSendStore), con, adOpenStatic, adLockReadOnly, adCmdText

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

Private Sub CMD_Q_DATA_Click()
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
                    SubFixData
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
Private Function FixRecDoc() As Long
On Error GoTo myerror
Dim loctable As New ADODB.Recordset
loctable.Open "SELECT * FROM FILE7_20H WHERE STORE = " & MyParn(cStoreSend), con, adOpenStatic, adLockReadOnly, adCmdText
Dim aInsert As Variant
prog1.Value = 0
prog1.Visible = True
Dim nRecordCount As Long, nRecord As Long, nAffect As Long
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    conShop2.Execute " UPDATE FILE7_20H SET Received  = " & IIf(loctable!Received, 1, 0) & " , USERACCEPT = " & addstring(loctable!UserAccept) & " , TIMEACCEPT = " & addDate(loctable!TimeAccept) & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
Loop


loctable.Close
loctable.Open "SELECT * FROM FILE7_10H WHERE STORE = " & MyParn(cStoreSend), con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
prog1.Visible = True
 nRecord = 0
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    conShop2.Execute " UPDATE FILE7_10H SET Received  = " & IIf(loctable!Received, 1, 0) & " , USERACCEPT = " & addstring(loctable!UserAccept) & " , TIMEACCEPT = " & addDate(loctable!TimeAccept) & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
Loop


loctable.Close
loctable.Open "SELECT * FROM FILE1_60H WHERE STORE1 = " & MyParn(cStoreSend) & " OR STORE2 = " & MyParn(cStoreSend), con, adOpenStatic, adLockReadOnly, adCmdText
prog1.Value = 0
 nRecord = 0
prog1.Visible = True
nRecordCount = loctable.RecordCount
Do Until loctable.EOF
    nRecord = nRecord + 1
    prog1.Value = Round(nRecord / nRecordCount, 2) * 100
    conShop2.Execute " UPDATE FILE1_60H SET Received  = " & IIf(loctable!Received, 1, 0) & " , USERACCEPT = " & addstring(loctable!UserAccept) & " , TIMEACCEPT = " & addDate(loctable!TimeAccept) & " WHERE DOC_NO = " & MyParn(loctable!doc_no)
    loctable.MoveNext
Loop

lastsub:
prog1.Visible = False
Exit Function
myerror:
MsgBox Err.Description
Err.Clear
FixRecDoc = -1
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

