VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form PrintGrd2 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ØÈÇÚÉ"
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
   WindowState     =   2  'Maximized
   Begin VSPrinter7LibCtl.VSPrinter Vp 
      Height          =   4515
      Left            =   150
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
      DefaultDevice   =   -1  'True
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
      Zoom            =   21.9946571682992
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
Attribute VB_Name = "PrintGrd2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cHeader As String, cBody As String, cFormat As String, cPageHeader As String, cPageHeader2 As String, cPageHeader3 As String, cFormat2 As String
Dim nFontSize As Integer
Dim aHRow, ahCol
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
Sub doprint(grid1, Optional nRate As Double = 1, Optional ptotal As Integer = -1, Optional pString As String = "", Optional pString2 As String = "", Optional pString3 As String = "", Optional bLeft As Boolean = False, Optional bLand As Boolean = False, Optional pFontSize As Integer = 11, Optional pFontName As String = "Simplified Arabic", Optional ByVal phRow, Optional phCol)
nFontSize = pFontSize
If Not IsMissing(phRow) Then aHRow = phRow
If Not IsMissing(phCol) Then ahCol = phCol
cPageHeader = pString
cPageHeader2 = pString2
cPageHeader3 = pString3

nBegin = IIf(bLeft, 0, grid1.Cols - 1)
nEnd = IIf(bLeft, grid1.Cols - 1, 0)
nStep = IIf(bLeft, 1, -1)
Vp.Orientation = IIf(bLand, orLandscape, orPortrait)
Vp = " "
With Vp
.ExportFormat = vpxRTF
.ExportFile = "D:\Ali Mail" & "\RepDoc.RTF"

Vp.FontSize = pFontSize
Vp.FontName = pFontName
Vp.MarginLeft = 300
Vp.MarginRight = 300
Vp.TextAlign = taCenterTop
Vp.MarginTop = 300
Vp.MarginBottom = 300

For nCol = nBegin To nEnd Step nStep
    If Not grid1.ColHidden(nCol) Then cFormat = cFormat & "+^~" & nRate * grid1.ColWidth(nCol) & "|"
    If Not grid1.ColHidden(nCol) Then cFormat2 = cFormat2 & "+^" & nRate * grid1.ColWidth(nCol) & "|"
Next
cFormat = Mid(cFormat, 1, Len(cFormat) - 1) & ";"
cFormat2 = Mid(cFormat2, 1, Len(cFormat2) - 1) & ";"

For nRow = 0 To grid1.Rows - 1
    If Not grid1.RowHidden(nRow) Then
        For nCol = nBegin To nEnd Step nStep
            If Not grid1.ColHidden(nCol) Then
                If nRow < grid1.FixedRows Then
                    cHeader = cHeader & grid1.TextMatrix(nRow, nCol) & "|"
                Else
                    cBody = cBody & grid1.TextMatrix(nRow, nCol) & "|"
                End If
            End If
        Next
        If Right(cHeader, 1) = "|" Then cHeader = Mid(cHeader, 1, Len(cHeader) - 1) & ";"
        If Right(cBody, 1) = "|" Then cBody = Mid(cBody, 1, Len(cBody) - 1) & ";"
    End If
Next

.StartDoc
'For I = 0 To grid1.Cols - 1
'    If Not grid1.ColHidden(I) Then nCols = nCols + 1
'Next

'For I = 0 To grid1.Rows - 1
'    If Not grid1.RowHidden(I) Then nRows = nRows + 1
'Next
.StartTable
.FontSize = 12
.TextAlign = taRightMiddle
.Paragraph = "          " & Format(Date, "DD-MM-YYYY")
.TextAlign = taCenterMiddle

.FontSize = 14

.Paragraph = cPageHeader
.FontBold = True
.FontUnderline = True
.Paragraph = cPageHeader2
.FontSize = pFontSize
.FontBold = False
.FontUnderline = False

.AddTable cFormat2, "", cHeader
If Not IsEmpty(aHRow) Then MergeRows aHRow
If Not IsEmpty(ahCol) Then MergeCols ahCol
Vp.FontBold = False

.AddTable cFormat, "", cBody, , , True
If ptotal = -1 Then
    .TableCell(tcFontBold, Vp.TableCell(tcRows), 0, Vp.TableCell(tcRows), Vp.TableCell(tcCols)) = True
ElseIf ptotal = -2 Then
    .TableCell(tcFontBold, 2, 1, 2, Vp.TableCell(tcCols)) = True
End If
'If Not IsEmpty(aHRow) Then MergeRows aHRow
If Not IsEmpty(ahCol) Then MergeCols ahCol
.EndTable
.EndDoc
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set PrintGrd = Nothing
End Sub
Private Sub vp_NewPage()
With Vp
.TextAlign = taLeftMiddle
.Paragraph = "page No. " & Vp.CurrentPage & "              "
.TextAlign = taCenterMiddle
If .CurrentPage > 1 And cPageHeader <> "" Then
    .FontSize = 14
    .Paragraph = cPageHeader
    .FontBold = True
    .FontUnderline = True
    .Paragraph = cPageHeader2
    .FontSize = nFontSize
    .FontBold = False
    .FontUnderline = False
    .AddTable cFormat2, "", cHeader
    If Not IsEmpty(aHRow) Then MergeRows aHRow
    If Not IsEmpty(ahCol) Then MergeCols ahCol
    .FontBold = False
End If
End With
End Sub
Private Sub MergeRows(aHRow)
For i = 0 To UBound(aHRow)
   MergeRow aHRow(i)
Next
End Sub
Private Sub MergeCols(aHRow)
For i = 0 To UBound(aHRow)
   MergeCol ahCol(i)
Next
End Sub
Private Sub MergeRow(nRow)
Dim nValue As Integer
cValue = ""
nCols = Vp.TableCell(tcCols)
For i = 1 To nCols
    If Vp.TableCell(tcText, nRow, i) = cValue Then
        nValue = nValue + 1
    Else
        If nValue > 0 Then
            Vp.TableCell(tcColSpan, nRow, i - (nValue + 1)) = nValue + 1
        End If
       cValue = Vp.TableCell(tcText, nRow, i)
       nValue = 0
    End If
Next
If nValue > 0 Then
    Vp.TableCell(tcColSpan, nRow, i - (nValue + 1)) = nValue + 1
End If
End Sub
Private Sub MergeCol(nCol)
Dim nValue As Integer
cValue = ""
NROWS = Vp.TableCell(tcRows)
For i = 1 To NROWS
    If Vp.TableCell(tcText, i, nCol) = cValue Then
        nValue = nValue + 1
    Else
        If nValue > 0 Then
            Vp.TableCell(tcRowSpan, i - (nValue + 1), nCol) = nValue + 1
        End If
       cValue = Vp.TableCell(tcText, i, nCol)
       nValue = 0
    End If
Next
If nValue > 0 Then
    Vp.TableCell(tcRowSpan, i - (nValue + 1), nCol) = nValue + 1
End If
End Sub


