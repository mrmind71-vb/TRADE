VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Morshed_Bar96 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10350
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15240
   BeginProperty Font 
      Name            =   "Simplified Arabic"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10350
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   825
      Left            =   2340
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   0
      Visible         =   0   'False
      Width           =   2850
      Begin VB.CommandButton Command5 
         Caption         =   "≈÷«›… —’Ìœ „ÊœÌ·«  „Ê—œ"
         Height          =   435
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   270
         Visible         =   0   'False
         Width           =   2685
      End
   End
   Begin VB.Frame Frame5 
      Height          =   825
      Left            =   12780
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   2355
      Begin Threed.SSCommand CMD_PrintBar 
         Height          =   465
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   820
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "Morshed_Bar96.frx":0000
         Caption         =   " ÿ»«⁄… »«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame4 
      Height          =   825
      Left            =   5175
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   3570
      Begin VB.Label xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï ⁄œœ «·ﬁÿ⁄ ··ÿ»«⁄…"
         Height          =   390
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   315
         Width           =   1995
      End
   End
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   8775
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   3990
      Begin VB.TextBox xCol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   315
         Width           =   915
      End
      Begin VB.TextBox xRow 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   315
         Width           =   915
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "«·⁄„Êœ :"
         Height          =   345
         Left            =   1125
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   315
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "«·’›:"
         Height          =   300
         Left            =   3150
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   315
         Width           =   615
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -135
      Top             =   270
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8205
      Left            =   135
      TabIndex        =   4
      Top             =   855
      Width           =   15000
      _cx             =   26458
      _cy             =   14473
      _ConvInfo       =   1
      Appearance      =   1
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
      Height          =   870
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   8955
      Width           =   3315
      Begin VB.TextBox xDelModel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   315
         Width           =   1740
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " ÕœÌœ „ÊœÌ· :"
         Height          =   390
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      Height          =   915
      Left            =   7875
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   9000
      Width           =   7305
      Begin VB.CommandButton CmdExit 
         Height          =   555
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Morshed_Bar96.frx":2982
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   555
         Left            =   3105
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Morshed_Bar96.frx":2ACC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton CmdUndo 
         Height          =   555
         Left            =   4590
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Morshed_Bar96.frx":5366
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton CmdSave 
         Height          =   555
         Left            =   5895
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Morshed_Bar96.frx":78DF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   270
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Õ–› «·„ÿ»Ê⁄"
         Height          =   540
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   270
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Morshed_Bar96"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim tAddPrint As New ADODB.Recordset
Dim ItemTable As New ADODB.Recordset
Dim NCOLS As Double
Dim NROWS As Double
Private Sub CMD_PrintBar96_Click()
    If Val(xRow.Text) > 16 Then
        MsgBox "«·’› «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·’›Ê› "
        Exit Sub
    End If
    If Val(xCol.Text) > 6 Then
        MsgBox "«·⁄„Êœ «·„ÿ·Ê» «·ÿ»«⁄… „‰ ⁄‰œÂ «ﬂ»— „‰ ⁄œœ «·√⁄„œ… "
        Exit Sub
    End If
    myreplace
    DoprintStr
    Set Myform = Me
    CardPrintNew.Show 1
    If MsgBox("Õ–› „«  „  ÿ»«⁄ Â „‰ «·„” ‰œ", vbYesNo + vbDefaultButton2) = vbYes Then
        delcheck
    End If
End Sub
Private Sub CMD_PrintBar_Click()
Dim tPrint As New ADODB.Recordset
Dim tCard As New ADODB.Recordset


Dim nPrintPrice As Double
Dim nCountRec As Double

Dim nTCount As Double
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(0)
SettingArray(cLeftMargin) = MyMeasure(0.1)
SettingArray(cCardWidth) = MyMeasure(4)
SettingArray(cCardHeight) = MyMeasure(2.5)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 1
SettingArray(cCols) = 1
SettingArray(cPageWidth) = MyMeasure(21)

con.Execute "delete * From Card"
tCard.Open "card", con, adOpenStatic, adLockPessimistic, adCmdTable
Set tPrint = mydb.OpenRecordset("Select file1_10.model,file1_10.ModelFact,FILE1_10.SUPP ,FILE1_10.MOSM,addPrint.Item,addPrint.DescA,addPrint.Scal,file1_10.price,file1_10.price2 AS PRICE2 , addPrint.Quant,AddPrint.price2 From addPrint Inner join File1_10 on AddPrint.Item = File1_10.item Where addPrint.isPrint  ORDER BY FILE1_10.FACT , FILE1_10.MODEL , FILE1_10.C_COLOR , FILE1_10.C_SCAL ")


With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
NCOLS = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * NCOLS) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = NCOLS, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡
If tPrint.RecordCount = 0 Then Exit Sub
Do
'************
    For i = 1 To tPrint.Quant
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = NCOLS)
        
        nHeight = 0
        nLast = MyMeasure(0)
        nLastCol = MyMeasure(0.1)
        For nCount = 1 To Val(xCount.Text)
' «·„Õ·
            nCountRec = nCountRec + 1
            tCard.AddNew
            tCard!Top = MyMeasure(0.2) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(1.6) - IIf(blastcol, nLastCol, 0)
            tCard!FontSize = 12
            tCard!Width = MyMeasure(2)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!TextAlign = taRightTop
            tCard!Text = Secondtitle
            tCard!isBarcode = False
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

' ITEM
            tCard.AddNew
            tCard!Top = MyMeasure(0.2) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(0)
            tCard!TextAlign = taLeftTop
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!FontUnderline = False
            tCard!FontSize = 12
            tCard!Text = tPrint!Item
            tCard!CardNo = nCard
            tCard.Update

' PAGE
            tCard.AddNew
            tCard!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.7) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(0.4)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontSize = 6
            tCard!TextAlign = taCenterTop
            tCard!Text = Format(nCountRec, "#0")
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

            tCard.AddNew
            tCard!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.85) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(0.4)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontSize = 6
            tCard!TextAlign = taCenterTop
            tCard!Text = "----"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

            tCard.AddNew
            tCard!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(0.4)
            tCard!Height = MyMeasure(0)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontSize = 6
            tCard!TextAlign = taCenterTop
            tCard!Text = Format(Val(xTotal.Caption), "#0")
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update


' BARCODE
            tCard.AddNew
            tCard!Left = MyMeasure(0.6) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.7) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(2.4)
            tCard!Height = MyMeasure(0.7)
            tCard!FontName = "arial"
            tCard!FontBold = False
            tCard!FontSize = 8
            tCard!TextAlign = taCenterTop
            tCard!Text = tPrint!Item
            tCard!isBarcode = True
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update


'  «··Ê‰
            tCard.AddNew
            tCard!Left = MyMeasure(2.6) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(1)
            tCard!Top = MyMeasure(0.5) + nHeight
            tCard!Height = MyMeasure(0)
            tCard!TextAlign = taRightTop
            tCard!FontName = "Tahoma (Arabic)"
            tCard!FontBold = True
            tCard!FontSize = 7
            tCard!Text = retitem(tPrint.Item, "Color")
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update

            tCard.AddNew
            tCard!Left = MyMeasure(2.6) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(1)
            tCard!Top = MyMeasure(0.8) + nHeight
            tCard!Height = MyMeasure(0)
            tCard!TextAlign = taRightTop
            tCard!FontName = "Tahoma (Arabic)"
            tCard!FontBold = True
            tCard!FontSize = 7
            tCard!Text = retitem(tPrint.Item, "CODE")
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
        
            tCard.AddNew
            tCard!Left = MyMeasure(2.6) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(1)
            tCard!Top = MyMeasure(1.1) + nHeight
            tCard!Height = MyMeasure(0)
            tCard!TextAlign = taRightTop
            tCard!FontName = "Tahoma (Arabic)"
            tCard!FontBold = True
            tCard!FontSize = 7
            tCard!Text = retitem(tPrint.Item, "MOSM")
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
        
            tCard.AddNew
            tCard!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(3.6)
            tCard!Top = MyMeasure(1.4) + nHeight
            tCard!Height = MyMeasure(0)
            tCard!TextAlign = taCenterTop
            tCard!FontName = "Arial (Arabic)"
            tCard!FontBold = True
            tCard!FontSize = 8
            tCard!Text = retitem(tPrint!Item, "DESCA") & " " & SayCode(FlagTable, 3, retitem(tPrint!Item, "FACT"))
            tCard!CardNo = nCard
            tCard.Update
    
    
            tCard.AddNew
            tCard!Left = MyMeasure(0.2) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(2) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = 0
            tCard!Height = 0
            tCard!TextAlign = taLeftTop
            tCard!FontName = "arial"
            tCard!FontBold = True
            tCard!FontUnderline = False
            tCard!FontSize = 10
            If XpRICE.Value <> 0 Then
                tCard!Text = Format(tPrint!price, "Fixed") & " LE"
            Else
                tCard!Text = Format("23" & tPrint!price & "32", "#0")
            End If
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(1.7) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(2) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(1)
            tCard!Height = 0
            tCard!FontName = "Traditional Arabic"
            tCard!FontBold = True
            tCard!TextAlign = taCenterTop
            tCard!FontUnderline = False
            tCard!FontSize = 10
            tCard!ForeColor = vbBlack
            tCard!Text = retitem(tPrint!Item, "MODELFACT0")
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(2.6) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.85) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(1)
            tCard!Height = 0
            tCard!FontName = "Tahoma (Arabic)"
            tCard!FontBold = True
            tCard!TextAlign = taRightTop
            tCard!FontUnderline = False
            tCard!FontSize = 14
            tCard!ForeColor = vbBlack
            tCard!Text = retitem(tPrint!Item, "scal")
            tCard!CardNo = nCard
            tCard.Update
            
            nHeight = SettingArray(cCardHeight) / Val(xCount.Text)
        Next
' ----------------
    Next
    tPrint.MoveNext
Loop Until tPrint.EOF
End With

End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
myreplace
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
myload
End Sub
Private Sub cmduno_Click()
myload
End Sub
Private Sub Command1_Click()
If MsgBox("Õ–› „«  „  ÿ»«⁄ Â „‰ «·„” ‰œ", vbYesNo + vbDefaultButton2) = vbYes Then
    delcheck
End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
Private Sub CmdDelAll_Click()
If MsgBox("Õ–› ﬂ· «·”Ã·« ", vbYesNo + vbDefaultButton1) = vbYes Then
    grid1.Rows = 1
    grid1.AddItem ""
    con.Execute "delete  from addprint "
End If
End Sub
Private Sub Command5_Click()
    BalBar.Show 1
    tAddPrint.Requery
    myload
End Sub
Private Sub Form_Load()
    
    con.CursorLocation = adUseClient
    con.Open strCon
    ItemTable.Open "select * from file1_10", con, adOpenKeyset, adLockReadOnly, adCmdText

    With grid1
        .Cols = 10
        .Rows = 1
    End With
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    myload
    
End Sub
Sub myload()
    On Error GoTo myerror
    cString = "SELECT AddPrint.item , file1_10.mosm , fact.desca , file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  AddPrint.quant , AddPrint.isprint  , AddPrint.DOC_NO  , AddPrint.row  from (AddPrint inner join file1_10 on file1_10.item = AddPrint.item ) inner join fact on fact.code = file1_10.fact ORDER BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_SCAL "
   
    data1.RecordSource = cString
    data1.Refresh
    grid1.AddItem "'"
    FixGrid
    CalcTot
    Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub FixGrid()
With grid1
    .Cols = 12
    .TextMatrix(0, 0) = "»«—ﬂÊœ"
    .TextMatrix(0, 1) = "„Ê”„"
    .TextMatrix(0, 2) = "„’‰⁄"
    .TextMatrix(0, 3) = "„ﬂ »"
    .TextMatrix(0, 4) = "—ﬁ„ „ÊœÌ·"
    .TextMatrix(0, 5) = "«·’‰›"
    .TextMatrix(0, 6) = "·Ê‰"
    .TextMatrix(0, 7) = "„ﬁ«”"
    .TextMatrix(0, 8) = "⁄œœ"
    .TextMatrix(0, 9) = "ÿ»«⁄…"
    .TextMatrix(0, 10) = ""
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 800
    .ColWidth(2) = 3000
    .ColWidth(3) = 1000
    .ColWidth(4) = 800
    .ColWidth(5) = 4000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .ColWidth(10) = 0
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .ColHidden(10) = True
    .ColHidden(11) = True
    .ColDataType(9) = flexDTBoolean
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
tAddPrint.Close
ItemTable.Close
Set tAddPrint = Nothing
Set ItemTable = Nothing
closeCon con
Unload Search3
Unload Me
Set Dream_Bar = Nothing
End Sub
Private Sub grid1_EnterCell()
If grid1.Col = 0 Or grid1.Col = 8 Or grid1.Col = 9 Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End Sub
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·”Ã· „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        con.BeginTrans
        con.Execute " delete from addprint where row = " & Val(grid1.TextMatrix(grid1.Row, 11))
        con.CommitTrans
        grid1.RemoveItem grid1.Row
        CalcTot
    End If
End If
If KeyCode = 112 And grid1.Col = 0 Then
     ItemsLookupAll Me, Search3
End If
End Sub
Private Sub myreplace()
con.Execute "Delete  From addprint "
With grid1
For i = 1 To .Rows - 2
    If Not grid1.RowHidden(i) Then
        If .TextMatrix(i, 0) <> "" Then
            cString = "insert into ADDPRINT( QUANT,ISPRINT,item,DOC_NO)" & _
            " Values(" & _
            Val(.TextMatrix(i, 8)) & "," & _
            IIf(Val(.TextMatrix(i, 9)) = 0, 0, 1) & "," & _
            addvalue(.TextMatrix(i, 0)) & "," & _
            addstring(.TextMatrix(i, 10)) & _
            ")"
            con.Execute cString
        End If
    End If
Next
End With
End Sub
Private Function MYVALID() As Boolean
'With grid1
'For I = 1 To grid1.Rows - 2
'    If Val(.TextMatrix(I, 8)) = 0 Then
'        MsgBox "«·ﬂ„Ì… €Ì— „”Ã·…"
'        Exit Function
'    End If
'Next
MYVALID = True
'End With
End Function
Private Sub delcheck()
For i = 1 To grid1.Rows - 2
   If Val(grid1.TextMatrix(i, 9)) <> 0 Then
        grid1.RowHidden(i) = True
   End If
Next
myreplace
End Sub
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(Search3.grid1.TextMatrix(Search3.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        MsgBox ("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound)
        Exit Sub
    End If
        
    grid1.TextMatrix(grid1.Row, 0) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 7) = 1
    grid1.TextMatrix(grid1.Row, 9) = True
    
    grid1.TextMatrix(grid1.Row, 2) = "1"
    grid1.TextMatrix(grid1.Row, 4) = "-1"
    GrdDesc grid1.TextMatrix(grid1.Row, 0)
    
    If Not validRow(grid1.Row) Then Exit Sub
    
    If grid1.Row = grid1.Rows - 1 Then
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 0
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.Select grid1.Rows - 1, 0
    End If
End If
End Sub
Private Sub GrdDesc(nItem)
With grid1
    ItemTable.Find "item = " & nItem, , adSearchForward, adBookmarkFirst
    If Not ItemTable.EOF Then
        grid1.TextMatrix(grid1.Row, 1) = ItemTable!MOSM
        grid1.TextMatrix(grid1.Row, 2) = GetDesca("select desca from fact where code = " & MyParn(ItemTable!FACT))
        grid1.TextMatrix(grid1.Row, 3) = ItemTable!SUPP & ""
        grid1.TextMatrix(grid1.Row, 4) = ItemTable!MODELFACT0
        grid1.TextMatrix(grid1.Row, 5) = ItemTable!DESCA
        grid1.TextMatrix(grid1.Row, 6) = ItemTable!Color
        grid1.TextMatrix(grid1.Row, 7) = ItemTable!SCAL
    End If
End With
End Sub
Private Sub DoprintStr()

Dim tCard As New ADODB.Recordset
Dim tPrint As New ADODB.Recordset
nLastMargin = 0
SettingArray(cUpMargin) = MyMeasure(-0.15)
SettingArray(cLeftMargin) = MyMeasure(0.4)
SettingArray(cCardWidth) = MyMeasure(3.5)
SettingArray(cCardHeight) = MyMeasure(1.852)
SettingArray(cBeginRow) = 1
SettingArray(cBeginCol) = 1
SettingArray(cRows) = 16
SettingArray(cCols) = 6
SettingArray(cPageWidth) = MyMeasure(21)

contemp.Execute "delete * From Card"
tCard.Open "card", contemp, adOpenKeyset, adLockOptimistic, adCmdTable
tPrint.Open "SELECT AddPrint.item , file1_10.mosm , file1_10.cost , file1_10.price , file1_10.price2 , fact.desca , file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  AddPrint.quant , AddPrint.isprint  , AddPrint.DOC_NO  from (AddPrint inner join file1_10 on file1_10.item = AddPrint.item ) inner join fact on fact.code = file1_10.fact   WHERE     (ADDPRINT.ISPRINT = 1) ORDER BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_SCAL", con, adOpenKeyset, adLockReadOnly, adCmdText

With tCard
nCard = 0
nRow = 0
nCard = 0
nCol = 0
NCOLS = SettingArray(cCols)
NROWS = SettingArray(cRows)

' ·«Œ Ì«— «·’› Ê«·⁄„Êœ

nBegin = ((IIf(Val(xRow.Text) <= 0, 1, Val(xRow.Text)) - 1) * NCOLS) + IIf(Val(xCol.Text) <= 0, 1, Val(xCol.Text))
For i = 1 To nBegin - 1
    nCard = nCard + 1
    nCol = IIf(nCol = NCOLS, 1, nCol + 1)
    nRow = IIf(nCol = 1, nRow + 1, nRow)
    nRow = IIf(nRow > NROWS, 1, nRow)
    blastrow = (nRow = NROWS)
    tCard.AddNew
    tCard!CardNo = nCard
    tCard.Update
Next
'«‰ Â«¡
If tPrint.RecordCount = 0 Then Exit Sub
nadd = MyMeasure(0.2)
Do
'************
    For i = 1 To tPrint!Quant
        nCard = nCard + 1
        nCol = IIf(nCol = NCOLS, 1, nCol + 1)
        nRow = IIf(nCol = 1, nRow + 1, nRow)
        nRow = IIf(nRow > NROWS, 1, nRow)
        blastrow = (nRow = NROWS)
        blastcol = (nCol = NCOLS)
        
        nHeight = 0
        nLast = MyMeasure(0.2)
        nLastCol = MyMeasure(0.2)
        For nCount = 1 To 1

            tCard.AddNew
            tCard!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.9) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(1.8)
            tCard!Height = MyMeasure(0.3) + nadd
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!FontSize = 8
            tCard!TextAlign = taRightTop
            tCard!Text = tPrint!Item
            tCard!isBarcode = 1
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update


            tCard.AddNew
            tCard!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = 0
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!FontBold = 0
            tCard!FontUnderline = 1
            tCard!FontSize = 7
            tCard!TextAlign = taLeftTop
            tCard!Text = tPrint!MODELFACT0
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(0.3) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.8) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = MyMeasure(2)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!TextAngle = 900
            tCard!FontBold = 1
            tCard!FontUnderline = 0
            tCard!FontSize = 7
            tCard!Text = "«”Êﬁ «·„’ÿ›Ï"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(2)
            tCard!TextAlign = taRightTop
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = 1
            tCard!FontSize = 8
            tCard!Text = tPrint!Item
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(2) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = 0
            tCard!TextAlign = taRightTop
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!FontSize = 5
            tCard!Text = Format(Int(tPrint!PRICE2) + 5, "#")
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Width = MyMeasure(2)
            tCard!TextAlign = taRightTop
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = 1
            tCard!FontSize = 8
            tCard!Text = tPrint!Item
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(2.8) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(0.8) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = 0
            tCard!Height = MyMeasure(0.35)
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!FontSize = 7
            tCard!Text = tPrint!SUPP
            tCard!TextAlign = taRightTop
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
        
        
            tCard.AddNew
            tCard!Left = MyMeasure(2.8) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = 0
            tCard!TextAlign = taRightTop
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!FontSize = 5
            tCard!Text = Format(Int(tPrint!COST) + 20, "#")
            tCard!CardNo = nCard
            tCard.Update

'            tCard.AddNew
'            tCard!Left = MyMeasure(2.8) - IIf(blastcol, nLastCol, 0)
'            tCard!Top = MyMeasure(1.1) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
'            tCard!Width = 0
'            tCard!Height = MyMeasure(0.35)
'            tCard!FontName = "arial"
'            tCard!FontBold = 0
'            tCard!FontSize = 6
'            tCard!Text = Format(Int(tPrint!price) + 10, "#")
'            tCard!TextAlign = taRightTop
'            tCard!ForeColor = vbBlack
'            tCard!CardNo = nCard
'            tCard.Update
        
            tCard.AddNew
            tCard!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            tCard!Width = MyMeasure(1)
            tCard!Top = MyMeasure(0.45) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1)
            tCard!Height = 0
            tCard!TextAlign = taCenterTop
            tCard!FontName = "Simplified arabic"
            tCard!FontBold = 0
            tCard!FontSize = 7
            If Not (tPrint!Color = "„ÊÕœ" Or tPrint!Color = "√·Ê«‰" Or tPrint!Color = "---") Then
                tCard!Text = tPrint!Color
            End If
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
    
            tCard.AddNew
            tCard!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.27) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = MyMeasure(2.5)
            tCard!Height = 0
            tCard!FontName = "Arial"
            tCard!TextAlign = taCenterTop
            tCard!FontBold = 1
            tCard!ForeColor = vbBlack
            tCard!FontSize = 7
            tCard!Text = tPrint!DESCA
            tCard!CardNo = nCard
            tCard.Update
    
            tCard.AddNew
            tCard!Left = MyMeasure(0.7) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = 0
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!TextAlign = taCenterTop
            tCard!FontBold = 1
            tCard!FontUnderline = 0
            tCard!FontSize = 8
            tCard!Text = Format(tPrint!price, "Fixed") & "Ã‰Ì…"
            tCard!ForeColor = vbBlack
            tCard!CardNo = nCard
            tCard.Update
            
            tCard.AddNew
            tCard!Left = MyMeasure(1) - IIf(blastcol, nLastCol, 0)
            tCard!Top = MyMeasure(1.6) + nHeight - IIf(blastrow And nCount = 2, nLast, 0) - MyMeasure(0.1) + nadd
            tCard!Width = MyMeasure(2)
            tCard!Height = 0
            tCard!FontName = "arial"
            tCard!FontBold = 0
            tCard!TextAlign = taRightTop
            tCard!FontUnderline = 1
            tCard!FontSize = 7
            tCard!ForeColor = vbBlack
            tCard!Text = tPrint!SCAL
            tCard!CardNo = nCard
            tCard.Update
            
            nHeight = SettingArray(cCardHeight) / 2
        Next
            
' ----------------
    Next
    tPrint.MoveNext
Loop Until tPrint.EOF
End With
End Sub
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For i = 1 To grid1.Rows - 2
    If i <> nRow Then
        If Trim(grid1.TextMatrix(i, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = i
            Exit Function
        End If
    End If
Next
End Function
Private Sub foundOther()
For i = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(i, 0)
    If nRow <> -1 Then
        grid1.Select nRow, 0, nRow, grid1.Cols - 1
        grid1.ShowCell nRow, 0
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 0) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & nRow
                
        Exit Sub
    End If
Next
End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 And Trim(grid1.EditText) = "" Then
    Cancel = True
End If
End Sub

Private Sub xDelModel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
With grid1
If xDelModel.Text = "" Then
    For i = 1 To .Rows - 1
        .TextMatrix(i, 9) = False
    Next i
Else
    For i = 1 To .Rows - 1
        If xDelModel.Text = .TextMatrix(i, 4) Then
            .TextMatrix(i, 9) = True
        End If
    Next i
End If
End With
End If
End Sub
Function CalcTot()
    With grid1
        For i = 1 To .Rows - 1
            If TurnValue(.TextMatrix(i, 9), "", False) Then
                CalcTot = CalcTot & Val(.TextMatrix(i, 8))
            End If
        Next i
    End With
    CalcTot = Format(CalcTot, "#0")
End Function

Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
        CalcTot
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 Then
    .RemoveItem .Row
    CalcTot
End If
End With
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Not (IsNumeric(.TextMatrix(nRow, 0))) Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Col = 0 Then
    ItemTable.Find " item = " & MyParn(.TextMatrix(.Row, 0)), , adSearchForward, adBookmarkFirst
    If Not ItemTable.EOF Then
       GrdDesc .TextMatrix(.Row, 0)
        .TextMatrix(.Row, 8) = 1
        .TextMatrix(.Row, 9) = True
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    Else
        .TextMatrix(.Row, 0) = ""
    End If
    If .Col >= 8 Then CalcTot
End If
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = defBox
    CalcTot
End If
End With
End Sub

