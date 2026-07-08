VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form printGrd3 
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
      Height          =   6225
      Left            =   90
      TabIndex        =   1
      Top             =   75
      Width           =   5565
      _cx             =   9816
      _cy             =   10980
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
      Zoom            =   34.1856060606061
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
Attribute VB_Name = "printGrd3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sHead1 As String, sHead2 As String, sHead3 As String
Public bLeft As Boolean, bLand As Boolean, pFontSize As Integer, pFontName As String
Public aColSpan As Variant, nBeginTotalSpan As Long, nValue As Long, aRowSpan As Variant, nSpanBegin As Long
Public grid1 As Variant, nRate As Double, ptotal As Integer
Dim cHeader As String, cBody As String, cFormat As String, cFormatHeader As String
Dim nBegin As Long, nEnd As Long, nStep As Long
Dim nFontSize As Integer, nHeaderFontSize As Integer
Dim aFormat(10), aBody(10)
Dim cGroup1 As String, cGroup2 As String, nGroup1 As Long, nGroup2 As Long, bNewGroup As Boolean, bNewGroupSub As Boolean, nRowGroup As Long, nGroupSubPage As Long
Public myForm As Form
Dim aHRow, ahCol
Private Sub Form_Activate()
DoEvents
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
    v = ScaleHeight - Vp.Top - 50
    If v > 0 Then Vp.Height = v
    
    '----------------------------------------------------
    ' set width
    '----------------------------------------------------
    v = ScaleWidth - Vp.Left
    If v > 0 Then Vp.Width = v
End Sub
Sub doprint(Optional nRate As Double = 1, Optional ptotal As Integer = -1, Optional pString1 As String = "", Optional pString2 As String = "", Optional pString3 As String = "", Optional bLeft As Boolean = False, Optional bLand As Boolean = False, Optional pFontSize As Integer = 11, Optional pFontName As String = "Simplified Arabic", Optional pGroup1 As Long = -1, Optional pGroup2 As Long = -1, Optional bNewPage As Boolean = False)
nFontSize = pFontSize
sHead1 = ArbString(pString1)
sHead2 = ArbString(pString2)
sHead3 = ArbString(pString3)
nGroup1 = pGroup1
nGroup2 = pGroup2

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

Dim nSpan As Long, nSpan2 As Long
Dim i As Long


For nCol = nBegin To nEnd Step nStep
    If Not grid1.ColHidden(nCol) Then
        cFormat = cFormat & turn(cFormat, "|") & "+>" & (nRate * grid1.ColWidth(nCol))
    End If
Next

.StartDoc
Dim cField As String
cField = "Dummy123"
For nRow = 0 To grid1.FixedRows - 1
    aFormat(nRow) = ""
    aBody(nRow) = ""
    If Not grid1.RowHidden(nRow) Then
        For nCol = nBegin To nEnd Step nStep
            If (Not grid1.ColHidden(nCol)) And (nCol <> nGroup1) And (nCol <> nGroup2) Then
                If grid1.TextMatrix(nRow, nCol) <> cField Or IsNumeric(grid1.TextMatrix(nRow, nCol)) Then
                    nwidth = (nRate * grid1.ColWidth(nCol))
                    cField = grid1.TextMatrix(nRow, nCol)
                    aFormat(nRow) = aFormat(nRow) & turn(aFormat(nRow), "|") & "+^" & (nwidth)
                    aBody(nRow) = aBody(nRow) & grid1.TextMatrix(nRow, nCol) & "|"
                Else
                    nwidth = nwidth + (nRate * grid1.ColWidth(nCol))
                    nFind = findLast(aFormat(nRow), "|")
                    If nFind > 0 Then aFormat(nRow) = Mid(aFormat(nRow), 1, nFind - 1) Else aFormat(nRow) = ""
                    aFormat(nRow) = aFormat(nRow) & turn(aFormat(nRow), "|") & "+^" & nwidth
                End If
            End If
        Next
        aBody(nRow) = aBody(nRow) & turn(aBody(nRow), ";")
    End If
Next

If nGroup1 < 0 Then
    For i = 0 To UBound(aFormat)
        If Format(i) <> "" Then
            .AddTable aFormat(i), "", aBody(i), , &H8000000F
        End If
    Next
End If
Dim sValue As String, bTotal As Boolean, nGroupCount As Long, bEndTotal As Boolean, bTotalRow As Boolean
cField = "Dummy123"
cGroup1 = "Dummy123"
cGroup2 = "Dummy123"
''myform.prog1.Visible = True
'myform.prog1.Value = 0
Dim nPage1 As Long, nPage2 As Long
nPage1 = 1
nPage2 = 1
For nRow = grid1.FixedRows To grid1.Rows - 1
    If Not grid1.RowHidden(nRow) Then
        If nGroup1 >= 0 Then
            If cGroup1 <> grid1.TextMatrix(nRow, nGroup1) And Left(grid1.TextMatrix(nRow, 0), 7) <> "«Ã„«·Ì " Then
                cGroup1 = grid1.TextMatrix(nRow, nGroup1)
                If nPage1 <> Vp.CurrentPage Then
                    nPage1 = Vp.CurrentPage
                    If bNewPage Then Vp.NewPage
                Else
                    Vp.Paragraph = ""
                End If
                
                Vp.FontSize = Vp.FontSize + 3
                Vp.FontBold = True
                If nPage1 <> Vp.CurrentPage And nPage > 1 Then
                    nPage1 = Vp.CurrentPage
                Else
                    Vp.Paragraph = cGroup1
                End If
                Vp.FontSize = Vp.FontSize - 3
                Vp.FontBold = True
                
                If nPage1 <> Vp.CurrentPage And nPage > 1 Then
                    nPage1 = Vp.CurrentPage
                Else
                    For i = 0 To UBound(aFormat)
                        If Format(i) <> "" Then
                            .AddTable aFormat(i), "", aBody(i), , &H8000000F
                        End If
                    Next
                End If
                Vp.FontBold = False
            End If
        End If
        
        If nGroup2 >= 0 Then
            If cGroup2 <> grid1.TextMatrix(nRow, nGroup2) Or bNewGroup Then
                cGroup2 = grid1.TextMatrix(nRow, nGroup2)
                Vp.FontSize = Vp.FontSize + 1
                Vp.FontBold = True
                If nPage2 <> Vp.CurrentPage Then
                    nPage2 = Vp.CurrentPage
                Else
                    Vp.Paragraph = cGroup2
                End If
                Vp.FontSize = Vp.FontSize - 1
                Vp.FontBold = False
            End If
        End If
            
        cBody = ""
        cFormat = ""
        sValue = ""
        cField = ""
        bEndTotal = False
        For nCol = nBegin To nEnd Step nStep
            If (Not grid1.ColHidden(nCol)) And nCol <> nGroup1 And nCol <> nGroup2 Then
                If grid1.ColFormat(nCol) = "#.##%" Then sValue = Round(Val(grid1.TextMatrix(nRow, nCol)) * 100, 2) & "%" Else sValue = grid1.TextMatrix(nRow, nCol)
                bTotalRow = Left(grid1.TextMatrix(nRow, 0), 7) = "«Ã„«·Ì " Or Left(grid1.TextMatrix(nRow, 0), 9) = "«·«Ã„«·Ì "
                If (Not IsNumeric(sValue)) And bTotalRow Then sValue = grid1.TextMatrix(nRow, 0)
                If IsNumeric(sValue) Then bEndTotal = True
                If sValue <> cField Or IsNumeric(sValue) Or ptotal = 0 Or Not bTotalRow Then
                    nwidth = (nRate * grid1.ColWidth(nCol))
                    If (Not IsNumeric(sValue)) And bTotalRow Then
                        cField = grid1.TextMatrix(nRow, 0)
                    Else
                        cField = grid1.TextMatrix(nRow, nCol)
                    End If
                    cFormat = cFormat & turn(cFormat, "|") & "+>" & (nwidth)
                    cBody = cBody & IIf(Not bEndTotal, grid1.TextMatrix(nRow, nCol), sValue) & "|"
                    bTotal = False
                Else
                    nwidth = nwidth + (nRate * grid1.ColWidth(nCol))
                    nFind = findLast(cFormat, "|")
                    If nFind > 0 Then cFormat = Mid(cFormat, 1, nFind - 1) Else cFormat = ""
                    cFormat = cFormat & turn(cFormat, "|") & "+>" & nwidth
                    bTotal = True
                End If
            End If
        Next
        cBody = cBody & turn(cBody, ";")
        .FontBold = bTotal
        .AddTable cFormat, "", cBody, , IIf(bTotal, vbYellow, vbWhite)
'        myform.prog1.Value = Round(nRow / (GRID1.Rows - 1), 2) * 100
    End If
Next
.EndDoc
'myform.prog1.Visible = False
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set printGrd3 = Nothing
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
.FontSize = 14
.FontBold = True
.FontUnderline = True

 If sHead1 <> "" Then .Paragraph = sHead1
 If sHead2 <> "" Then .Paragraph = sHead2
 If sHead3 <> "" Then .Paragraph = sHead3

.FontUnderline = False
.FontSize = nFontSize

'cBodyHeader = ""
'With grid1
'    For nRow = 0 To grid1.FixedRows - 1
'        If Not grid1.RowHidden(nRow) Then
'            For nCol = nBegin To nEnd Step nStep
'                If Not grid1.ColHidden(nCol) Then
'                    If grid1.ColFormat(nCol) = "#.##%" Then
'                        cBodyHeader = cBodyHeader & turn(grid1.TextMatrix(nRow, nCol), Round(Val(grid1.TextMatrix(nRow, nCol)) * 100, 2) & "%") & "|"
'                    Else
'                        cBodyHeader = cBodyHeader & grid1.TextMatrix(nRow, nCol) & "|"
'                    End If
'                End If
'            Next
'            cBodyHeader = cBodyHeader & turn(cBodyHeader, ";")
'        End If
'        Vp.AddTable cFormatHeader, "", cBodyHeader, , &H8000000F
'    Next
'End With
If Vp.CurrentPage > 1 Then
    If nGroup1 >= 0 Then
        'Vp.TextAlign = taRightMiddle
        Vp.Paragraph = ""
        Vp.FontSize = Vp.FontSize + 3
        Vp.FontBold = True
        Vp.Paragraph = cGroup1
        Vp.FontSize = Vp.FontSize - 3
        Vp.FontBold = False
    End If
    
    For i = 0 To UBound(aFormat)
        If aFormat(i) <> "" Then
            Vp.AddTable aFormat(i), "", aBody(i), , &H8000000F
        End If
    Next
    
    'If (Not bNewGroupSub) And nGroup2 > 0 Then
    If nGroup2 >= 0 Then
        Vp.FontSize = Vp.FontSize + 2
        Vp.FontBold = True
        Vp.Paragraph = cGroup2
        Vp.FontSize = Vp.FontSize - 2
        Vp.FontBold = False
    End If
End If
End With
Vp.FontBold = False
End Sub
Private Sub MergeRows(aRowSpan, nSpanBegin)
For i = 0 To UBound(aRowSpan)
   MergeRow aRowSpan(i), nSpanBegin
Next
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
For nRow = 0 To grid1.Rows - 1
    If Not grid1.RowHidden(nRow) Then
        If Left(Trim(.TextMatrix(nRow, 0)), 7) = "≈Ã„«·Ì " Or Trim(.TextMatrix(nRow, 0)) = "«·≈Ã„«·Ì" Then
            Vp.TableCell(tcColSpan, nRow, (nCol + 1) - nValue) = nValue
            Vp.TableCell(tcText, nRow, (nCol + 1) - nValue, nRow, (nCol + 1) - nValue) = .TextMatrix(nRow, 0)
            Vp.TableCell(tcColAlign, nRow, (nCol + 1) - nValue, nRow, (nCol + 1) - nValue) = taLeftTop
            Vp.TableCell(tcBackColor, nRow) = IIf(Left(Trim(.TextMatrix(nRow, 0)), 7) = "≈Ã„«·Ì ", &HC0FFFF, &HFFFFC0)
            Vp.TableCell(tcFontBold, nRow) = True
        End If
    End If
Next
End With
End Sub
Private Sub AddDefault()
If nRate = 0 Then nRate = 1
If pFontSize = 0 Then pFontSize = 11
If pFontName = "" Then pFontName = "Simplified Arabic"
End Sub
Private Sub myMergeRow(nRow)
Dim nValue As Integer, cString As String
cValue = "dummy"
NCOLS = Vp.TableCell(tcCols)
For nCol = nBegin To nEnd Step nStep
     If Trim(Vp.TableCell(tcText, nRow, i)) <> Trim(cValue) Then
        If nValue > 1 Then Vp.TableCell(tcColSpan, nRow, i - (nValue)) = nValue
        cValue = Vp.TableCell(tcText, nRow, i)
        nValue = 1
    Else
        nValue = nValue + 1
    End If
Next
If nValue > 1 Then Vp.TableCell(tcColSpan, nRow, i - (nValue)) = nValue
End Sub
Private Function findLast(pString, cFind) As Long
For i = Len(pString) To 1 Step -1
    If Mid(pString, i, 1) = cFind Then
        findLast = i
        Exit Function
    End If
Next
End Function
