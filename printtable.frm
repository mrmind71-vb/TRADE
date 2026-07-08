VERSION 5.00
Object = "{A8561640-E93C-11D3-AC3B-CE6078F7B616}#1.0#0"; "VSPRINT7.ocx"
Begin VB.Form PrintTable 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ÿ»«⁄…"
   ClientHeight    =   6720
   ClientLeft      =   735
   ClientTop       =   3000
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Simplified Arabic"
      Size            =   8.25
      Charset         =   178
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
      Left            =   585
      TabIndex        =   0
      Top             =   405
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
End
Attribute VB_Name = "PrintTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cHeader As String, cBody As String, cFormat As String, cPageHeader As String, cPageHeader2 As String, cPageHeader3 As String, cFormat2 As String
Dim grouptable As New ADODB.Recordset
Public nDigitAdd As Integer
Public bTotal As Boolean
Dim nFontSize As Integer, cPar, cMainGroup As String, cHeader1, cHeader2, cheader3, bShowMainGroup As Boolean, bShowGroup As Boolean, bNewPage As Boolean
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
Sub doprint(Optional pHeader1 As String = "", Optional pHeader2 As String = "", Optional pHeader3 As String = "", Optional pShowGroup As Boolean = True, Optional pShowMainGroup As Boolean = True, Optional pNewPage As Boolean = False)
Dim locTable As New ADODB.Recordset, nGrandTotal As Double
cHeader1 = pHeader1: cHeader2 = pHeader2: cheader3 = pHeader3
bShowGroup = pShowGroup: bShowMainGroup = pShowMainGroup: bNewPage = pNewPage
With Vp
Vp = " "
Vp.Orientation = orLandscape
Vp.FontSize = 14
Vp.FontName = "Simplified Arabic"
Vp.MarginLeft = 400
Vp.MarginRight = 200
Vp.TextAlign = taCenterTop
Vp.MarginTop = 400
Vp.MarginBottom = 200
Vp.Header = "«· «—ÌŒ : " & Format(Date, "dd-mm-yyyy") & "||’›Õ… %d"
grouptable.Open " SELECT PRINT1.MAINGROUP,PRINT1.GROUP,PRINT1.GROUPDESCA FROM PRINT1 GROUP BY PRINT1.MAINGROUP, PRINT1.GROUP,PRINT1.GROUPDESCA order by PRINT1.MAINGROUP, PRINT1.GROUP", contemp, adOpenStatic, adLockReadOnly, adCmdText
.StartDoc
Do Until grouptable.EOF
    If grouptable!MainGroup <> cMainGroup Then
        If cMainGroup <> "" And bNewPage Then Vp.NewPage
        If bShowMainGroup Then
            .TextAlign = taRightTop
            .FontUnderline = True
            .FontBold = True
            .FontName = "Simplified Arabic"
            .FontSize = 14
            .Paragraph = "„Ã„Ê⁄… " & GetDesca("select desca from file1_50 where code = " & MyParn(grouptable!MainGroup))
        End If
        cMainGroup = grouptable!MainGroup
    End If
    PrintTable grouptable!Group
    grouptable.MoveNext
Loop
Dim TotalTable As New ADODB.Recordset
TotalTable.Open "select sum(quant) as sumofQuant from print1", contemp, adOpenStatic, adLockReadOnly, adCmdText
If Not TotalTable.EOF Then nGrandTotal = Val(TotalTable!sumOfQuant & "")
TotalTable.Close: Set TotalTable = Nothing
.TextAlign = taRightTop
.FontBold = True
If bTotal Then
.FontSize = 14
.Paragraph = "«·«Ã„«·Ì :  " & nGrandTotal
End If
.EndDoc
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
grouptable.Close
Set grouptable = Nothing
Set printTablefrm = Nothing
Set printTablefrm = Nothing
End Sub
Private Sub Vp_NewPage()
With Vp
'.FontBold = False: .TextAlign = tarightTop: .FontSize = 12: .FontUnderline = False
'.Text = "page no. " & Vp.CurrentPage
'.TextAlign = taRightTop
'.Text = "Date : " & Format(Date, "d-m-yyyy")
.DrawLine .MarginLeft, .CurrentY, .PageWidth - .MarginRight, .CurrentY
If cHeader1 & cHeader2 & cheader3 <> "" Then
    .FontBold = True
    '.FontUnderline = True
    .TextAlign = taCenterTop
    .FontSize = 14
    .Paragraph = ""
    If cHeader1 <> "" Then .Paragraph = cHeader1
    If cHeader2 <> "" Then .Paragraph = cHeader2
    If cheader3 <> "" Then .Paragraph = cheader3
   ' .Paragraph = ""
    .TextAlign = taRightTop
    .FontUnderline = False
End If

If .CurrentPage > 1 Then
    If Not grouptable.EOF Then
        If cMainGroup = grouptable!MainGroup And bShowMainGroup Then
            .TextAlign = taRightTop
            .FontBold = True
            .FontUnderline = True
            '.Paragraph = ""
            .FontSize = 14
            .Paragraph = "  «»⁄ „Ã„Ê⁄… " & GetDesca("select desca from file1_51 where code = " & MyParn(grouptable!MainGroup))
            .FontUnderline = False
        End If
     
        If cPar = grouptable!GROUPDESCA And bShowGroup Then
            .FontBold = True
            .FontUnderline = False
            .TextAlign = taRightTop
            .FontSize = 14
            .Paragraph = grouptable!GROUPDESCA & " <-"
            .FontBold = False
        End If
    End If
    '.Paragraph = GroupTable!GroupDesca
End If
End With
'With Vp
'If .CurrentPage = 1 And cPageHeader <> "" Then
'If cPageHeader <> "" Then
'    .Paragraph = FirstTitle & IIf(sBranch <> "", "-", "") & sBranch
'    .FontSize = Vp.FontSize + 2
'    .FontBold = True
'    .Paragraph = cPageHeader
'    If cPageHeader2 <> "" Then .Paragraph = cPageHeader2
'    If cPageHeader3 <> "" Then .Paragraph = cPageHeader3
''    .Paragraph = ""
'    .FontSize = Vp.FontSize - 2
'End If
'.AddTable cFormat2, "", cHeader
'.FontBold = False
'End With
End Sub
Private Sub PrintTable(pGroup)
Dim locTable As New ADODB.Recordset, nRow As Integer, nCol As Integer, f As String, h As String, cBody As String
Dim aCols() As String
' ⁄„· «·«” ⁄·«„
cString = "TRANSFORM Sum(PRINT1.QUANT) AS SumOfQUANT " & _
          " SELECT PRINT1.ROW & PRINT1.WIDTH1 , Sum(PRINT1.QUANT) AS [«·≈Ã„«·Ì] " & _
          " FROM PRINT1 " & _
          " WHERE PRINT1.GROUP = " & MyParn(pGroup) & _
          " GROUP BY PRINT1.ROW & PRINT1.WIDTH1 PIVOT PRINT1.COL & PRINT1.WIDTH2; "
locTable.Open cString, contemp, adOpenStatic, adLockReadOnly, adCmdText
locTable.MoveLast
nRecordcount = locTable.RecordCount + 1
locTable.MoveFirst

NCOLS = locTable.Fields.Count
With Vp
    .TextAlign = taRightTop
    .FontBold = True
    .FontUnderline = False
    
    If bShowGroup Then .Paragraph = grouptable!GROUPDESCA
     cPar = grouptable!GROUPDESCA
    .FontBold = False

    NCOLS = locTable.Fields.Count
    .GetMargins
    nColWidth = ((.PageWidth - 160) - (.MarginRight + .MarginLeft + ((NCOLS + 1) * 0))) / NCOLS
    
    If nColWidth > 1000 Then nColWidth = 1000
    ReDim aCols(NCOLS - 1)
    For nCol = 1 To NCOLS
        'f = f & IIf(f <> "", "|", "") & "^+" & nColWidth
        If nCol <> 2 Then f = f & IIf(f <> "", "|", "") & ">+" & nColWidth + IIf(nCol = 1, 500, 0)
        If nCol <> 2 Then h = h & IIf(h <> "", "|", "") & IIf(nCol = 1, " ", Chr(254) & Mid(Replace(locTable.Fields(nCol - 1).Name, "_", "."), 3))
    Next
    If bTotal Then
        f = f & IIf(f <> "", "|", "") & "<+" & (nColWidth + 80)
        h = h & IIf(h <> "", "|", "") & IIf(nCol = 1, " ", locTable.Fields(1).Name)
    End If
    
    nFontSize = retFontSize(nColWidth)
    Do Until locTable.EOF
        For nCol = 1 To NCOLS
            If nCol = 1 Then cRow = cRow & IIf(cRow <> "", "|", "") & Mid(locTable(nCol - 1), 4)
            If nCol <> 2 And nCol <> 1 Then cRow = cRow & "|" & myFormat(locTable(nCol - 1) & "", nFontSize)
            If nCol <> 1 And nCol <> 2 Then
                aCols(nCol - 2) = Val(aCols(nCol - 2)) + Val(locTable.Fields(nCol - 1) & "")
            End If
        Next
        If bTotal Then cRow = cRow & IIf(cRow <> "", "|", "") & myFormat(locTable(1) & "", nFontSize, True)
        cBody = cBody & IIf(cBody <> "", ";", "") & cRow
        cRow = ""
        locTable.MoveNext
    Loop
    If bTotal Then
        For I = 1 To NCOLS - 2
            aCols(NCOLS - 1) = Val(aCols(NCOLS - 1)) + Val(aCols(I))
        Next
        cRow = "«·«Ã„«·Ì"
        For I = 1 To NCOLS - 1
            cRow = cRow & IIf(cRow <> "", "|", "") & myFormat(aCols(I), nFontSize, I = NCOLS - 1)
        Next
        cBody = cBody & IIf(cBody <> "", ";", "") & cRow & ";"
    Else
        cBody = cBody & turnFound(cBody, ";")
    End If
    
    f = RevTable(f)
    h = RevTable(h)
    cBody = RevTable(cBody, True)
    
' ⁄„· «·ÃœÊ·
    .StartTable
    .AddTable f, h, cBody
    
    .TableCell(tcFontName, 0, 1, nRecordcount, NCOLS) = "Simplified Arabic"
    .TableCell(tcFontSize, 0, 1, nRecordcount, NCOLS) = nFontSize
    
    .TableCell(tcRowHeight, 0, 1, nRecordcount, NCOLS) = 350
    .TableCell(tcIndent) = 0
   
    .TableCell(tcFontBold, 0, 1, 0, NCOLS) = True
    .TableCell(tcFontBold, 1, NCOLS, nRecordcount, NCOLS) = True
    .TableCell(tcFontBold, 1, 1, nRecordcount, 1) = True
    .TableCell(tcFontBold, nRecordcount, 1, nRecordcount, NCOLS) = True
    .EndTable
    .Paragraph = ""
End With
locTable.Close
Set locTable = Nothing
End Sub
Private Function myFormat(cString As String, pFontSize, Optional pLastCol As Boolean = False) As String
nDec = 20
If IsNumeric(cString) Then
    If Val(cString) = 0 Then myFormat = "": Exit Function
    myFormat = Round(Val(cString), 2)
     'nDec = IIf(pFontSize > 7, 5, 4) + IIf(pLastCol, 1, 0) + nDigitAdd + IIf(pFontSize >= 9, 3, 0)
    'nLen = Len(myFormat)
'    If Len(myFormat) > nDec Then
'        myFormat = Mid(myFormat, 1, nDec)
'        If Not IsNumeric(Right(myFormat, 1)) And Len(Trim(myFormat)) > 2 Then
'            myFormat = Mid(myFormat, 1, Len(myFormat) - 1)
'        End If
'    End If
Else
   myFormat = cString
End If
'myFormat = TurnValue(Val(myFormat), 0, "")
'If Val(myFormat) < 100 Or Val(myFormat) >= 1000 Then
'    myFormat = Mid(myFormat, 1, 4)
'Else
'    myFormat = Mid(myFormat, 1, 3)
'End If
End Function
Private Function retFontSize(nColWidth)
Dim nFontSize, cFontName
nFontSize = Vp.FontSize: cFontSize = "Simplified Arabic"
Vp.FontSize = 11: retFontSize = 11
For I = 1 To 1
    If Vp.TextWidth("123457") < nColWidth Then Exit For
    retFontSize = retFontSize - 1
Next
Vp.FontSize = nFontSize: Vp.FontName = cFontName
End Function
