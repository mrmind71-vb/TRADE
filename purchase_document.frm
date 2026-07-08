VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form purchase_document 
   BackColor       =   &H00FFFFFF&
   Caption         =   "’Ê— «·’‰›"
   ClientHeight    =   10290
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20250
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   13770
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   8505
      Width           =   3795
      Begin VB.CheckBox chkFit 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fit Width"
         Height          =   195
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   315
         Width           =   1095
      End
      Begin VB.ComboBox cboZoomBy 
         Height          =   315
         ItemData        =   "purchase_document.frx":0000
         Left            =   1125
         List            =   "purchase_document.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Zoom By %"
         Height          =   255
         Left            =   135
         TabIndex        =   10
         Top             =   315
         Width           =   930
      End
   End
   Begin ImgXCtrl6.ImgXCtrl ImgX1 
      Height          =   10095
      Left            =   45
      TabIndex        =   6
      Top             =   90
      Width           =   13695
      _ExtentX        =   24156
      _ExtentY        =   17806
      BackColor       =   16777215
      BorderStyle     =   4
      SelectionLineType=   4
      Center          =   -1  'True
      ImageBorderThickness=   1
      ScrollSmallChange=   0.1
      AntialiasDisplay=   -1  'True
      WMFPlaceable    =   0   'False
      SelectionEnableMove=   -1  'True
      AutoZoomType    =   2
      DoubleBuffer    =   -1  'True
      LicenseUserName =   "mrmind"
      LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   17595
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   8505
      Width           =   2580
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
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
         Picture         =   "purchase_document.frx":0004
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1305
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   979
         _Version        =   196610
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
         Picture         =   "purchase_document.frx":2327
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "purchase_document.frx":469D
      End
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   -2175
      TabIndex        =   0
      Top             =   75
      Visible         =   0   'False
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog Common1 
      Left            =   4140
      Top             =   10530
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   16425
      Top             =   9630
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
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
   Begin ImgXCtrl6.ImgXCtrl imgx2 
      Height          =   8565
      Left            =   -855
      TabIndex        =   2
      Top             =   9765
      Visible         =   0   'False
      Width           =   9555
      _ExtentX        =   16854
      _ExtentY        =   15108
      BackColor       =   16777215
      BorderStyle     =   4
      AutoZoom        =   -1  'True
      SelectionLineType=   4
      Center          =   -1  'True
      ImageBorderThickness=   1
      DoubleBuffer    =   -1  'True
      LicenseUserName =   "mrmind"
      LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   8385
      Left            =   13770
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   90
      Width           =   6405
      _cx             =   11298
      _cy             =   14790
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
      BackColorSel    =   12648447
      ForeColorSel    =   -2147483630
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
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
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.Label xPhoto 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   9630
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   8730
      Visible         =   0   'False
      Width           =   2040
   End
End
Attribute VB_Name = "purchase_document"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents twain As ImgXTwain
Attribute twain.VB_VarHelpID = -1
Public bEdit As Boolean
Private WithEvents MyPrinter As imgxPrint
Attribute MyPrinter.VB_VarHelpID = -1
Dim bIg As Boolean
Dim con As New ADODB.Connection

Dim imageTable As ADODB.Recordset

Dim con_image As New ADODB.Connection
Public sDoc_no As String
Dim fs As New FileSystemObject
Dim aPhoto
Private Sub cboZoomBy_Click()
If cboZoomBy.text = "" Then
    If chkFit.Value = 1 Then
        chkFit_Click
    Else
        ImgX1.Zoom = 1
    End If
Else
    ImgX1.Zoom = Val(cboZoomBy.text) / 100
End If
End Sub

Private Sub Check1_Click()

End Sub

Private Sub chkFit_Click()
If Not bIg Then
    ImgX1.AutoZoom = chkFit.Value = 1
    If chkFit.Value = 0 Then ImgX1.Zoom = 1
    LoadPhoto
End If
'ImgX1.Update
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set ImgXTwain = Nothing
closeCon con

addText chkFit.Name, chkFit.Value
addText cboZoomBy.Name, cboZoomBy.ListIndex

Set documentsfrm = Nothing
Err.Clear
End Sub

Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
   myAddItem
End If
If Not myreplace(Row) Then
    myLoadGrd
ElseIf grid1.TextMatrix(Row, grid1.Cols - 1) = "" Then
    myLoadGrd
End If
End With
End Sub
Private Sub Form_Load()
openConDoc con_image

Set grid1.DataSource = DATA1

fixZoom

chkFit.Value = Val(RetTxt(chkFit.Name))
cboZoomBy.ListIndex = 2

myLoadGrd
CellPos 13, 0, grid1.Cols - 1
End Sub
Private Sub ScanImage()
On Error GoTo myerror
Set twain = New ImgXTwain
twain.OpenTwain Me.hwnd
If twain.QuerySupport(ixtcResolution) Then
     twain.Resolution = 150
End If
'twain.Acquire  Check1.Value = 1, Me.hWnd
twain.Acquire False, Me.hwnd

Exit Sub
myerror:
MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub
Private Sub ImgX1_MouseWheelScroll(ByVal ScrollValue As Long, ByVal Keys As Long)
If chkFit.Value = 1 Then chkFit.Value = 0
ImgX1.Zoom = ImgX1.Zoom + (ScrollValue / 100)
End Sub
Private Sub Twain_ImageAcquired(Image As ImgX_Image)
ImgX1.Images.Replace Image, , False
ImgX1.TIFCompression = ixtcJPEGCompression

myReplaceImage xPhoto.Caption

LoadPhoto
End Sub
Private Sub cmdPrint_Click()
Dim i As Long
Dim Index As Integer
Dim sList As String

Set MyPrinter = New imgxPrint
MyPrinter.PageFrom = 1
MyPrinter.PageTo = grid1.Rows - 2
MyPrinter.PageMax = grid1.Rows - 2
MyPrinter.MarginLeft = 0
MyPrinter.MarginRight = 0
MyPrinter.MarginTop = 0
MyPrinter.MarginBottom = 15
MyPrinter.PageMin = 1
MyPrinter.Antialias = True
'On Error GoTo myerror

If MyPrinter.ShowPrinter(Me.hwnd) Then
    If MyPrinter.Range = iprAllPages Then
    ElseIf MyPrinter.Range = ixprSelection And xPhoto.Caption <> "" Then
        sList = xPhoto.Caption
    Else
        For i = MyPrinter.PageFrom To MyPrinter.PageTo
            If grid1.TextMatrix(i, grid1.Cols - 1) <> "" Then
                sList = sList & Tr(sList, ",") & grid1.TextMatrix(i, grid1.Cols - 1)
            End If
        Next
    End If
End If

Dim loctable As New ADODB.Recordset
Dim cString As String

cString = "select * from PURCHASE_IMAGES " & _
          " WHERE DOC_NO = " & MyParn(sDoc_no)
If sList <> "" Then
    cString = cString & " AND ID_IMAGE IN (" & sList & ")"
End If

Set loctable = cmd(cString, con_image).Execute
Do Until loctable.EOF
    imgx2.Images.Clear
    imgx2.Import.FromMemoryFile (loctable!Image)
    MyPrinter.PrintImage "Print Document", imgx2.Images(0), False, True
    loctable.MoveNext
Loop
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub LoadPhoto()
On Error GoTo myerror
ImgX1.Images.Clear
If xPhoto.Caption <> "" Then
    ImgX1.Images.Clear
    
    Dim loctable As New ADODB.Recordset
    Set loctable = cmd("SELECT * FROM PURCHASE_IMAGES WHERE ID = " & xPhoto.Caption, con_image).Execute
    
    If Not loctable.EOF Then
        ImgX1.Import.FromMemoryFile (loctable!Image)
    End If
End If
Exit Sub
myerror:

'Err.Clear
End Sub
Private Sub fixGrd()
With grid1
.FormatString = "„”·”·|«·»Ì«‰|„”Õ|’Ê—…|Õ–›"
.ColWidth(0) = 600
.ColWidth(1) = 3800
.ColWidth(2) = 600
.ColWidth(3) = 600
.ColWidth(4) = 600
.ColComboList(2) = "..."
.ColComboList(3) = "..."
.ColComboList(4) = "..."
.ColHidden(.Cols - 1) = True
For i = 0 To grid1.Cols - 1
    grid1.ColAlignment(i) = flexAlignRightCenter
Next
MakeSerial
End With
End Sub
Private Sub myloadgrd2()
'Dim aPrm As Variant
'aPrm = AddFlag(aPrm, "ID_CASH", sId_cash)
End Sub
Private Sub Handlecontrols()
'Me.cmdDelAll.Enabled = bedit
'Me.cmdAddReplace.Enabled = bedit
'Me.cmdDelCur.Enabled = bedit
'Me.cmdAddPhoto.Enabled = bedit
End Sub
Private Function myCut(pString) As String
Dim aLocal As Variant
aLocal = Split(pString, "\")
For i = 0 To UBound(aLocal) - 2
    myCut = myCut & turn(myCut, "\") & aLocal(i)
Next
End Function
Private Function myreplace(Optional Row As Long = -1)
Dim aInsert As Variant
With grid1
    con_image.BeginTrans
    On Error GoTo myerror
    For i = IIf(Row = -1, 1, Row) To IIf(Row = -1, grid1.Rows - 2, Row)
        aInsert = AddFlag(Empty, "DESCA", addstring(.TextMatrix(i, 1)))
        aInsert = AddFlag(aInsert, "DOC_NO", addstring(sDoc_no))
        If grid1.TextMatrix(i, .Cols - 1) = "" Then
            con_image.Execute addInsert(aInsert, "PURCHASE_IMAGES")
        Else
            con_image.Execute addUpdate(aInsert, "PURCHASE_IMAGES", "ID = " & .TextMatrix(i, .Cols - 1))
        End If
    Next
    con_image.CommitTrans
End With
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con_image.RollbackTrans
Err.Clear
End Function
Private Sub myLoadGrd()
'Set ImageTable = New ADOdb.Recordset
'Set ImageTable = cmd("sp_get_account_image", con_image, adStoredProc, aPrm).Execute

Dim cString As String
cString = "SELECT DESCA as [«·«”„],NULL ,NULL,NULL,PURCHASE_IMAGES.ID" & _
          " FROM PURCHASE_IMAGES" & _
          " WHERE PURCHASE_IMAGES.DOC_NO = " & MyParn(sDoc_no)
Set DATA1.Recordset = myRecordSet(cString, con_image)
grid1_EnterCell
myAddItem
fixGrd
End Sub
Private Function myAddItem()
grid1.AddItem ""
End Function
Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
On Error GoTo myerror
Set fs = CreateObject("Scripting.FileSystemObject")
If Col = 2 And xPhoto.Caption <> "" Then
    nOption = 2
    ScanImage
'    camerafrm.sPicture = Document_Files(sId_cash, xPhoto.Caption)
'    camerafrm.Show 1
     LoadPhoto
ElseIf Col = 3 And xPhoto.Caption <> "" Then
    'MyCreateFolder doc_dir(sId_cash)
    Dim cFile As String, cNewFile As String
    Common1.FileName = ""
    Common1.InitDir = doc_dir()
    Common1.Filter = "Pictures (*.Jpg)|*.Jpg"
    Common1.ShowOpen
    If Common1.FileTitle <> "" Then
        cFile = Common1.FileName
        If cFile <> "" Then
            myReplaceImage xPhoto.Caption, cFile
        End If
        LoadPhoto
    End If
ElseIf Col = 4 Then
    If MsgBox("Õ–› «·’Ê—…", vbOKCancel + vbDefaultButton2) Then
        con_image.Execute "UPDATE PURCHASE_IMAGES SET IMAGE = NULL where ID = " & xPhoto.Caption
        LoadPhoto
    End If
End If
Exit Sub
myerror:
        MsgBox Err.Description
        Err.Clear
End Sub
Private Sub grid1_EnterCell()
If bEdit Then
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
If grid1.Row <= 0 Then
    xPhoto.Caption = ""
ElseIf grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    xPhoto.Caption = ""
Else
    xPhoto.Caption = grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
End If
End Sub

Private Sub grid1_KeyPress(KeyAscii As Integer)
If keyscii = 13 And grid1.Col <> 2 And grid1.Col <> 3 Then KeyAscii = 0
End Sub
Private Function validRow(Row As Long) As Boolean
With grid1
If Trim(.TextMatrix(Row, 1)) = "" Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Trim(grid1.EditText) = "" Then
    Cancel = True
End If
End Sub
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub xPhoto_Change()
LoadPhoto
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
ElseIf KeyCode = 46 And bEdit Then
    If Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) <> "" Then
        If MsgBox("«·€«¡ «·”Ã· ?? Â· «‰  „ √ﬂœ", vbYesNo + vbDefaultButton2) = vbYes Then
            con_image.BeginTrans
            On Error GoTo myerror
            con_image.Execute "delete from PURCHASE_IMAGES where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con_image.CommitTrans
            myreplace
            myLoadGrd
        End If
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
myLoadGrd
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then CellPos KeyCode, Row, Col
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 2 Then
    grid1.Select Row, Col + 1
ElseIf Row < grid1.Rows - 1 Then
    grid1.Select Row + 1, NextEmpty(grid1, Row + 1, 1, 1)
    grid1.ShowCell Row + 1, 1
Else
    grid1.Select Row, Col
End If
End Sub
Private Sub Grid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With grid1
If .Rows = 0 Then Exit Sub
If Trim(grid1.TextMatrix(grid1.Row, grid1.Cols - 1)) = "" Then Exit Sub
R = .Row
R = .DragRow(R)
If R <> grid1.Rows - 1 Then
    myreplace
End If
myLoadGrd
End With
End Sub
Private Sub Twain_CanCloseTwain()
    ' This event is called after you call Acquire.
    ' It let's you know when it's safe to call CloseTwain.
    twain.CloseTwain
    ' Steps menu
End Sub
Public Function Document_Files(pId As String) As String
'Document_Files = doc_dir & "\" & sDoc_no & "-" & sId_cash & "-" & pId & ".jpg"
End Function
Function doc_dir() As String
doc_dir = App.Path & "\Documents"
End Function
Private Sub fixZoom()
    LoadingCboZoomBy = True
    With cboZoomBy
        .Clear
        .AddItem ""
        .AddItem "80"
        .AddItem "90"
        'Some fiddly code to ensure we always have 100% in the list in its
        'correct position even if our loop doesn't hit I = 100.  Also
        'record its index for use in Zoom100Percent.
        Zoom100PctCboZoomIndex = -1
        For i = 100 To 150 Step 10
            If i > 100 And Zoom100PctCboZoomIndex < 0 Then
                .AddItem i
                Zoom100PctCboZoomIndex = .NewIndex
            End If
            .AddItem CStr(i)
            If i = 100 Then Zoom100PctCboZoomIndex = .NewIndex
        Next
    End With
End Sub
Private Sub addImage(Optional pFile As String = "", Optional Image As ImgX_Image)


        ' --- Clean Up ---
LoadPhoto

myInform "Done ..."
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear

End Sub
Private Function myReplaceImage(sid As String, Optional pFile As String = "") As Boolean
If sid <> "" Then
    myReplaceImage = SaveImageToDB(pFile, ImgX1, "PURCHASE_IMAGES", "IMAGE", "ID", sid, con_image)
Else
    myReplaceImage = SaveImageToDB(pFile, ImgX1, "PURCHASE_IMAGES", "IMAGE", "ID", sid, con_image)
End If
End Function
