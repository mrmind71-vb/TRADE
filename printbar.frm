VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{0BFA85A1-F9B8-11CF-8939-444553540000}#1.0#0"; "barcode.ocx"
Begin VB.Form CardPrintNew 
   Caption         =   "ÿ»«⁄… »«—ﬂÊœ"
   ClientHeight    =   6720
   ClientLeft      =   735
   ClientTop       =   3000
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   6720
   ScaleWidth      =   5775
   Begin VSPrinter7LibCtl.VSPrinter Vp 
      Height          =   5265
      Left            =   375
      TabIndex        =   0
      Top             =   150
      Width           =   4890
      _cx             =   8625
      _cy             =   9287
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      MousePointer    =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty HdrFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _ConvInfo       =   1
      AutoRTF         =   -1  'True
      Preview         =   -1  'True
      DefaultDevice   =   0   'False
      PhysicalPage    =   -1  'True
      AbortWindow     =   -1  'True
      AbortWindowPos  =   0
      AbortCaption    =   "Printing..."
      AbortTextButton =   "Cancel"
      AbortTextDevice =   "on the %s on %s"
      AbortTextPage   =   "Now printing Page %d of"
      FileName        =   ""
      MarginLeft      =   0
      MarginTop       =   0
      MarginRight     =   0
      MarginBottom    =   0
      MarginHeader    =   0
      MarginFooter    =   0
      IndentLeft      =   0
      IndentRight     =   0
      IndentFirst     =   0
      IndentTab       =   720
      SpaceBefore     =   0
      SpaceAfter      =   0
      LineSpacing     =   100
      Columns         =   1
      ColumnSpacing   =   180
      ShowGuides      =   2
      LargeChangeHorz =   300
      LargeChangeVert =   300
      SmallChangeHorz =   30
      SmallChangeVert =   30
      Track           =   0   'False
      ProportionalBars=   -1  'True
      Zoom            =   26.4470169189671
      ZoomMode        =   3
      ZoomMax         =   400
      ZoomMin         =   10
      ZoomStep        =   25
      EmptyColor      =   -2147483636
      TextColor       =   0
      HdrColor        =   0
      BrushColor      =   0
      BrushStyle      =   0
      PenColor        =   0
      PenStyle        =   0
      PenWidth        =   0
      PageBorder      =   0
      Header          =   ""
      Footer          =   ""
      TableSep        =   "|;"
      TableBorder     =   7
      TablePen        =   0
      TablePenLR      =   0
      TablePenTB      =   0
      NavBar          =   3
      NavBarColor     =   -2147483633
      ExportFormat    =   0
      URL             =   ""
      Navigation      =   3
      NavBarMenuText  =   "Whole &Page|Page &Width|&Two Pages|Thumb&nail"
   End
   Begin BARCODELib.Barcode Barcode1 
      Height          =   870
      Left            =   315
      TabIndex        =   1
      Top             =   5535
      Width           =   4950
      _Version        =   65536
      _ExtentX        =   8731
      _ExtentY        =   1535
      _StockProps     =   25
      Text            =   "1"
      Type            =   11
      TypeName        =   "Code 128A"
      Text            =   "1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowText        =   0   'False
   End
End
Attribute VB_Name = "CardPrintNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim nLeftMargin, nCardHeight, nPageWidth, nRightMargin
Dim i, i2, nRow, nUpMargin As Integer, nCardWidth As Integer
Dim nCol As Integer
Dim tCard As New ADODB.Recordset
Dim nTextWidth, nTextHeight
Dim cOldPrint As String
Private Sub Form_Activate()
PrintArray
End Sub
Private Sub Form_Load()
    On Error GoTo myerror
    If Not IsEmpty(aPrinterBar) Then
        Vp.Device = retFlag(aPrinterBar, "name")
        aPrinterBar = Empty
    End If

'    cOldPrint = Vp.Device
    openCon con
    'Vp.PaperSize = pprLegal
   ' Dim cPrintBarCode As String
 '   cPrintBarCode = GetDesca("SELECT PRINTBARCODE FROM [PRINT]")
  '  If cPrintBarCode <> "" Then Vp.Device = cPrintBarCode
    Exit Sub
myerror:
    MsgBox "·«Ì„ﬂ‰ ÿ»«⁄… «·»«—ﬂÊœ "
    MsgBox Err.Description
    Err.Clear
End Sub
Private Sub Form_Resize()
    Dim v!
    '----------------------------------------------------
    ' set height
    '----------------------------------------------------
    v = ScaleHeight - Vp.Top - 100
    If v > 0 Then Vp.Height = v
    
    '----------------------------------------------------
    ' set width
    '----------------------------------------------------
    v = ScaleWidth - Vp.Left
    If v > 0 Then Vp.Width = v
    
End Sub
Private Function myRight(nRight, nwidth)
myRight = nRight - nwidth
End Function
Private Sub PrintArray()
Dim bNewRow, bNewPage

'On Error GoTo PrintError
tCard.Open "Select * From Card order by CardNo", contemp, adOpenStatic, adLockReadOnly, adCmdText

nUpMargin = SettingArray(cUpMargin)
nLeftMargin = SettingArray(cLeftMargin)
nCardWidth = SettingArray(cCardWidth)
nCardHeight = SettingArray(cCardHeight)
nPrintRow = SettingArray(cRows)

nRow = 0
nCol = 0

NCOLS = SettingArray(cCols)
nPageWidth = SettingArray(cPageWidth)

'SetOriginalSettings
Vp.ZoomMode = zmWholePage
Vp.StartDoc

With Vp
    Do Until tCard.EOF
       If nCardNo <> tCard!CardNo Then
            nCol = IIf(nCol = NCOLS, 1, nCol + 1)
            nRow = IIf(nCol = 1, nRow + 1, nRow)
            nCardNo = tCard!CardNo
        End If
        
       If nRow > SettingArray(cRows) Then
           .NewPage
           nRow = 1
       End If
        
        If Not tCard!isPhoto Or Not tCard!isBarcode Then
            If Not IsNull(tCard!text) Then
                .FontName = tCard!FontName
                .FontBold = tCard!FontBold
                .FontSize = tCard!FontSize
                .FontUnderline = tCard!FontUnderline
                .FontItalic = tCard!FontItalic
                .FontStrikethru = tCard!Strikethrough
                If Not IsNull(tCard!PenColor) Then
                    .PenColor = vbBlack
                End If
                 
                If Not IsNull(tCard!PenWidth) Then
                    .PenWidth = tCard!PenWidth
                End If
                 
                .TextColor = TurnValue(tCard!ForeColor, Null, vbBlack)
                nFieldWidth = IIf(tCard!Width = 0, .TextWidth(tCard!text), tCard!Width)
                nFieldHeight = IIf(tCard!Height = 0, .TextHeight(tCard!text), tCard!Height)
                Vp.TextAlign = IIf(IsNull(tCard!TextAlign), taLeftTop, tCard!TextAlign)
                If Not IsNull(tCard!TextAngle) Then
                    .TextAngle = tCard!TextAngle
                Else
                    .TextAngle = 0
                End If
                .TextBox tCard!text, Calcx, CalcY, nFieldWidth, nFieldHeight
            End If
         End If
         If tCard!isBarcode Then
             nFieldWidth = tCard!Width
             nFieldHeight = tCard!Height
             Barcode1.text = tCard!text
             Barcode1.CreatePictureBySize tCard!Width, tCard!Height
            .DrawPicture Barcode1.Picture, Calcx, CalcY, tCard!Width, tCard!Height
         End If
        tCard.MoveNext
    Loop
End With
Vp.EndDoc
Exit Sub
PrintError:
MsgBox "Œÿ√ „« ﬁœ ÕœÀ «À‰«¡ «·ÿ»«⁄… " & Err.Description
Vp.EndDoc
End Sub
Private Function CalcY()
CalcY = ((nRow - 1) * nCardHeight) + nUpMargin + tCard!Top
End Function
Private Function Calcx()
Calcx = nLeftMargin + tCard!Left + ((nCol - 1) * nCardWidth)
End Function
Private Sub Form_Unload(Cancel As Integer)
'Vp.Device = cOldPrint
tCard.Close
Set tCard = Nothing
End Sub
