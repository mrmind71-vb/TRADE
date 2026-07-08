VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form printGrdTable 
   Caption         =   "ÿ»«⁄…"
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   5775
   WindowState     =   2  'Maximized
   Begin VSPrinter7LibCtl.VSPrinter Vp 
      Height          =   4515
      Left            =   180
      TabIndex        =   1
      Top             =   75
      Width           =   4665
      _cx             =   8229
      _cy             =   7964
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
      Zoom            =   23.3901515151515
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
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   6525
      Visible         =   0   'False
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "Positioning"
      Height          =   255
      Left            =   3600
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   3795
   End
End
Attribute VB_Name = "printGrdTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sHead1 As String, sHead2 As String, sHead3 As String
Public bLeft As Boolean, bLand As Boolean, pFontSize As Integer, pFontName As String
Public aColSpan As Variant, nBeginTotalSpan As Long, nValue As Long, aRowSpan As Variant, nGroup1 As Long
Public grid1 As Variant, nRate As Double, ptotal As Integer
Dim cHeader As String, cBody As String, cFormat As String
Dim nFontSize As Integer, nHeaderFontSize As Integer
Dim aHRow, ahCol
Private Sub Form_Activate()
DoEvents
doprint grid1, nRate, ptotal, sHead1, sHead2, sHead3, bLeft, bLand, pFontSize, pFontName, aColSpan, nBeginTotalSpan, nValue, nGroup1
End Sub
Private Sub Form_Load()
nHeaderFontSize = 14
AddDefault
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
Sub doprint(grid1, Optional nRate As Double = 1, Optional ptotal As Integer = -1, Optional pString1 As String = "", Optional pString2 As String = "", Optional pString3 As String = "", Optional bLeft As Boolean = False, Optional bLand As Boolean = False, Optional pFontSize As Integer = 11, Optional pFontName As String = "Simplified Arabic", Optional aColSpan As Variant, Optional nBeginTotalSpan As Long = -1, Optional nValue As Long, Optional ByVal nGroup1)
'Sub doprint(grid1, Optional nRate As Double = 1, Optional ptotal As Integer = -1, Optional pString1 As String = "", Optional pString2 As String = "", Optional pString3 As String = "", Optional bLeft As Boolean = False, Optional bLand As Boolean = False, Optional pFontSize As Integer = 11, Optional pFontName As String = "Simplified Arabic", Optional ByVal aSpan, Optional nRowSpan1 As Long = 0, Optional aColSpan As Variant)
nFontSize = pFontSize
'If Not isEmpty(phRow) Then aHRow = phRow
'If Not isEmpty(phCol) Then ahCol = phCol

sHead1 = ArbString(pString1)
sHead2 = ArbString(pString2)
sHead3 = ArbString(pString3)

nBegin = IIf(bLeft, 0, grid1.Cols - 1)
nEnd = IIf(bLeft, grid1.Cols - 1, 0)
nStep = IIf(bLeft, 1, -1)
Vp.Orientation = IIf(bLand, orLandscape, orPortrait)
Vp = " "
With Vp

Vp.FontSize = pFontSize
Vp.FontName = pFontName
Vp.MarginLeft = 300
Vp.MarginRight = 300
Vp.TextAlign = taCenterTop
Vp.MarginTop = 300
Vp.MarginBottom = 300
Vp.Paragraph = "test"

Dim nSpan As Long, nSpan2 As Long
Dim i As Long

For nCol = nBegin To nEnd Step nStep
    If Not grid1.ColHidden(nCol) Then
        cFormat = cFormat & turn(cFormat, "|") & "+>" & (nRate * grid1.ColWidth(nCol))
    End If
Next

If Not IsEmpty(aColSpan) Then
    For i = 0 To UBound(aColSpan)
        aColSpan(i) = Printcol(aColSpan(i), grid1)
    Next
End If
For nRow = 0 To grid1.FixedRows - 1
    If Not grid1.RowHidden(nRow) Then
        For nCol = nBegin To nEnd Step nStep
            cHeader = cHeader & turn(cHeader, "|") & grid1.TextMatrix(nRow, nCol)
        Next
    End If
Next



Dim nBeginRow As Long, bExit As Boolean
.StartDoc
nBeginRow = grid1.FixedRows
Do Until nRow >= grid1.Rows - 1
    cBody = ""
    
    For nRow = nBeginRow To grid1.Rows - 1
        If Not grid1.RowHidden(nRow) Then
            For nCol = nBegin To nEnd Step nStep
                If Not grid1.ColHidden(nCol) Then
                    If grid1.ColFormat(nCol) = "#.##%" Then
                        cBody = cBody & turn(grid1.TextMatrix(nRow, nCol), Round(Val(grid1.TextMatrix(nRow, nCol)) * 100, 2) & "%") & "|"
                    Else
                        cBody = cBody & grid1.TextMatrix(nRow, nCol) & "|"
                    End If
                End If
            Next
            cBody = cBody & turn(cBody, ";")
        End If
        If nRow > 0 Then
            If Left(Trim(grid1.TextMatrix(nRow, nGroup1)), 7) = "≈Ã„«·Ì " Or Trim(grid1.TextMatrix(nRow, nGroup1)) = "«·≈Ã„«·Ì" Then Exit For
        End If
    Next
    
    nBeginRow = nRow + 1
    cHeader = cHeader & turn(cHeader, ";")
        
    .StartTable
    
    If Trim(grid1.TextMatrix(nRow, nGroup1)) = "«·≈Ã„«·Ì" Then
        .AddTable cFormat, "", cBody, , , True
    Else
        .AddTable cFormat, cHeader, cBody, , , True
    End If
    
    .TableCell(tcFontBold, 0, 1, 0, Vp.TableCell(tcCols)) = taCenterMiddle
    .TableCell(tcAlign, 0, 1, 0, Vp.TableCell(tcCols)) = taCenterMiddle
    .TableCell(tcBackColor, 0, 1, 0, Vp.TableCell(tcCols)) = &H8000000F
'
    If ptotal = -1 Then
        .TableCell(tcFontBold, Vp.TableCell(tcRows), 0, Vp.TableCell(tcRows), Vp.TableCell(tcCols)) = True
        .TableCell(tcBackColor, Vp.TableCell(tcRows), 0, Vp.TableCell(tcRows), Vp.TableCell(tcCols)) = &HC0FFFF
    ElseIf ptotal = -2 Then
        .TableCell(tcFontBold, 1, 1, 1, Vp.TableCell(tcCols)) = True
        .TableCell(tcBackColor, 1, 1, 1, Vp.TableCell(tcCols)) = &HC0FFFF
    ElseIf ptotal = -3 And Not IsEmpty(aRowSpan) Then
        For i = 0 To UBound(aRowSpan)
            .TableCell(tcFontBold, aRowSpan(i), 1, aRowSpan(i), Vp.TableCell(tcCols)) = True
            .TableCell(tcBackColor, aRowSpan(i), 1, aRowSpan(i), Vp.TableCell(tcCols)) = &HC0FFFF
        Next
    End If
    
    If Not IsEmpty(aColSpan) Then MergeCols aColSpan
    If nBeginTotalSpan <> -1 Then MergeAuto grid1, nBeginTotalSpan, nValue
    
    .EndTable
    If nRow < grid1.Rows - 1 Then
        If Trim(grid1.TextMatrix(nRow + 1, nGroup1)) <> "«·≈Ã„«·Ì" Then
            .Paragraph = " "
        End If
    End If
Loop
.EndDoc

End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set printGrdTable = Nothing
End Sub
Private Sub vp_NewPage()
With Vp
.FontSize = 10
.FontBold = False
.TextAlign = taLeftTop
.TextBox "’›Õ… " & Vp.CurrentPage, 400, 20, 4000, 1000
.TextAlign = taRightTop
.TextBox "«· «—ÌŒ : " & Format(Date, "yyyy/mm/dd"), Vp.PageWidth - (4000 + 400), 20, 4000, 1000
.TextAlign = taCenterMiddle
'If .CurrentPage > 1 Then
    .FontSize = 14
    .FontBold = True
    .FontUnderline = True

     If sHead1 <> "" Then .Paragraph = sHead1
     If sHead2 <> "" Then .Paragraph = sHead2
     If sHead3 <> "" Then .Paragraph = sHead3
'End If
.FontBold = False
.FontUnderline = False
.FontSize = nFontSize
.TextAlign = taRightTop
End With
End Sub
Private Sub MergeRows(aRowSpan, nSpanBegin)
'For I = 0 To UBound(aRowSpan)
'   MergeRow aRowSpan(I), nSpanBegin
'Next
End Sub
Private Sub MergeCols(aColSpan)
Dim cString As String
For i = 0 To UBound(aColSpan)
   cString = cString & turn(cString, "-") & Abs(Abs(aColSpan(i)))
   MergeCol cString
Next
End Sub
Private Sub MergeRow(nRow, nSpanBegin)
Dim nValue As Integer, cString As String
cValue = "dummy"
NCOLS = Vp.TableCell(tcCols)
For i = nSpanBegin To NCOLS
     If Trim(Vp.TableCell(tcText, nRow, i)) <> Trim(cValue) Then
        If nValue > 1 Then
            Vp.TableCell(tcColSpan, nRow, i - (nValue)) = nValue
        End If
        cValue = Vp.TableCell(tcText, nRow, i)
        nValue = 1
    Else
        nValue = nValue + 1
    End If
Next
If nValue > 1 Then
    Vp.TableCell(tcColSpan, nRow, i - (nValue)) = nValue
End If
End Sub
Private Sub MergeCol(pString)
Dim nValue As Integer, aLocal As Variant, cString As String
cValue = "Dummy"
NROWS = Vp.TableCell(tcRows)
aLocal = Split(pString, "-")

nCol = aLocal(UBound(aLocal))

For i = 1 To NROWS
    cString = ""
    For i2 = 0 To UBound(aLocal)
        cString = cString & Vp.TableCell(tcText, i, aLocal(i2))
    Next
    If Trim(cString) <> Trim(cValue) Then
        If nValue > 1 Then
            Vp.TableCell(tcRowSpan, i - (nValue), nCol) = nValue
        End If
        cValue = cString
        nValue = 1
    Else
        nValue = nValue + 1
    End If
Next
If nValue > 1 Then
    Vp.TableCell(tcRowSpan, i - (nValue), nCol) = nValue
End If
End Sub
Private Function Printcol(nCol, pGrid) As Long
With pGrid
For i = .Cols - 1 To nCol Step -1
    If Not pGrid.ColHidden(i) Then
         Printcol = Printcol + 1
    End If
Next
End With
End Function
Private Sub MergeAuto(grid1, nBeginTotalSpan, nValue)
Dim nCol As Long
With grid1
nCol = Printcol(nBeginTotalSpan, grid1)
For nRow = 1 To Vp.TableCell(tcRows)
    If Left(Vp.TableCell(tcText, nRow, nCol, nRow, nCol), 7) = "≈Ã„«·Ì " Or Trim(Vp.TableCell(tcText, nRow, nCol, nRow, nCol)) = "«·≈Ã„«·Ì" Then
        'Vp.TableCell(tcColAlign, nRow, (nCol + 1) - nValue, nRow, (nCol + 1) - nValue) = taLeftTop
        Vp.TableCell(tcColSpan, nRow, (nCol + 1) - nValue, nRow) = nValue
        Vp.TableCell(tcText, nRow, (nCol + 1) - nValue, nRow, (nCol + 1) - nValue) = Replace(Vp.TableCell(tcText, nRow, nCol, nRow, nCol), Chr(254), "")
        Vp.TableCell(tcBackColor, nRow) = IIf(Left(Vp.TableCell(tcText, nRow, nCol, nRow, nCol), 7) = "≈Ã„«·Ì ", &HC0FFFF, &HFFFFC0)
        Vp.TableCell(tcFontBold, nRow) = True
    'Else
    '    Vp.TableCell(tcColAlign, nRow, (nCol + 1) - nValue, nRow, (nCol + 1) - nValue) = taRightTop
    End If
Next
End With
End Sub
Private Sub AddDefault()
If nRate = 0 Then nRate = 1
If pFontSize = 0 Then pFontSize = 11
If pFontName = "" Then pFontName = "Simplified Arabic"
End Sub

