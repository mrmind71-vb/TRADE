VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form search_Empty 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ÇÓÊÚáÇã"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   12360
   StartUpPosition =   2  'CenterScreen
   Tag             =   "Factory"
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7260
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   12165
      _cx             =   21458
      _cy             =   12806
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
      Rows            =   1
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
      TabBehavior     =   0
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
      AutoSizeMouse   =   0   'False
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   12300
      TabIndex        =   5
      Top             =   7425
      Width           =   12360
      Begin VB.CheckBox xEnter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   420
         Left            =   45
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CheckBox xBegin 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Begin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002F2F2F&
         Height          =   420
         Left            =   900
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   780
      End
      Begin Threed.SSCommand cmdFilter 
         Height          =   375
         Left            =   1890
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   45
         Visible         =   0   'False
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   661
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Filter"
         ButtonStyle     =   3
      End
      Begin VB.Label lblCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8820
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   45
         Width           =   2760
      End
   End
   Begin MSAdodcLib.Adodc Ado1 
      Height          =   510
      Left            =   300
      Top             =   1125
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   900
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
      Index           =   0
      Left            =   0
      Top             =   0
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   90
      TabIndex        =   3
      Top             =   7380
      Visible         =   0   'False
      Width           =   5145
      Begin VB.TextBox txtlookup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   135
         MaxLength       =   100
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Visible         =   0   'False
         Width           =   3300
      End
      Begin MSDataListLib.DataCombo cmbLookup 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   180
         Visible         =   0   'False
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   688
         _Version        =   393216
         MatchEntry      =   -1  'True
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   3555
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   225
         Visible         =   0   'False
         Width           =   705
      End
   End
End
Attribute VB_Name = "Search_empty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Generalarray, listarray, GrdArray
Public sCaption As String
Public sid As String, aValue As Variant, nFontSize As Integer
Public aFilter As Variant, aPar As Variant, bNoRef As Boolean, aAddRow As Variant
Public bNoSort As Boolean
Dim bAct As Boolean, cString As String
Dim bEnterwork As Boolean
Public sControl As String, bGo As Boolean, bEnter As Boolean
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdGo_Click()
myLoadGrd
Fixgrd
If grid1.Rows > 1 Then
    grid1.SetFocus
End If
End Sub
Private Sub cmbLookup_Change(Index As Integer)
If (cmbLookup(Index).MatchedWithList Or Trim(cmbLookup(Index).BoundText) = "") And Not bEnter Then
    myLoadGrd
End If
End Sub

Private Sub cmbLookup_Click(Index As Integer, Area As Integer)
'If Area = 2 Then myLoadGrd
'If (cmbLookup(I).MatchedWithList Or Trim(cmbLookup(I).BoundText) = "") Then
'    myLoadGrd
'End If
End Sub

Private Sub cmbLookup_GotFocus(Index As Integer)
myGotFocus cmbLookup(Index)
End Sub

Private Sub cmbLookup_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
bEnterwork = True
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If retFlag(aPar, "noClick") Then
        Generalarray(0).ProcNoClick
    Else
        'Unload Me
        Me.Hide
    End If
End If
End Sub
Private Sub Form_Activate()
If bAct And (Not bNoRef) Then
    LoadArrays
    myLoadGrd
End If
bAct = True
On Error Resume Next
bEnterwork = False
grid1.SetFocus
DoEvents
'If txtlookup(1).Visible Then
'    txtlookup(1).SetFocus
'ElseIf cmbLookup(1).Visible Then
'    cmbLookup(1).SetFocus
'End If
Err.Clear
End Sub
Private Sub Form_Load()
bAct = False
Me.Caption = sCaption

Set grid1.DataSource = Ado1
If Not bNoSort Then grid1.ExplorerBar = flexExSort
Ado1.CommandType = adCmdText

LoadArrays

Frame2.Width = Generalarray(3)
grid1.Cols = UBound(GrdArray) + 1
LoadControls

SetValue
'xEnter.Visible = Not bEnter
xEnter.Value = IIf(retFlag(aValue, "enter") Or bEnter, 1, 0)
xBegin.Value = IIf(retFlag(aValue, "begin"), 1, 0)


If nFontSize <> 0 Then grid1.Font.Size = nFontSize
If UBound(Generalarray) = 3 Then
    myLoadGrd
Else
    If UBound(Generalarray) >= 4 Then
        If Not Generalarray(4) Then myLoadGrd Else Fixgrd
    End If
End If
'cmdFilter.Visible = retFlag(aFilter, "filter")
Handlecontrols
End Sub
Private Sub Form_Unload(Cancel As Integer)
For I = 0 To UBound(listarray)
    If listarray(I, 4) = "" Then
        If I < txtlookup.UBound Then
            If mySplit(listarray(I, 3), 1, ":") = "last_clicked" Then
                addSetting "SEARCH_LIST_" & mySplit(listarray(I, 3), 2, ":"), txtlookup(I + 1).text, TempSave(Generalarray(0), sid)
            End If
        End If
    ElseIf I < cmbLookup.UBound Then
        If mySplit(listarray(I, 5), 1, ":") = "last_clicked" Then
            addSetting "SEARCH_LIST_" & mySplit(listarray(I, 5), 2, ":"), cmbLookup(I + 1).BoundText, TempSave(Generalarray(0), sid)
        End If
    End If
Next
Set Search = Nothing
End Sub
Private Sub grid1_AfterSort(ByVal Col As Long, Order As Integer)
If Not IsEmpty(aAddRow) Then
    grid1.AddItem "", 1
    grid1.TextMatrix(1, Val(retFlag(aAddRow, "col"))) = retFlag(aAddRow, "text") & ""
    grid1.Cell(flexcpBackColor, 1, 0, , grid1.Cols - 1) = &HE0E0E0
End If
End Sub
Private Sub grid1_BeforeSort(ByVal Col As Long, Order As Integer)
If Not IsEmpty(aAddRow) Then
    grid1.RemoveItem 1
End If
End Sub

Private Sub grid1_DblClick()
If grid1.Row > 0 Then
    If sControl = "" Then Generalarray(0).myProc Else Generalarray(0).myProc sControl
Else
    bEnterPress = False
End If
End Sub
Private Sub grid1_GotFocus()
For I = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(I) Then Exit For
Next
If grid1.Rows > 1 Then grid1.Select 1, I
End Sub
Sub myLoadGrd()
On Error GoTo myerror
cString = Generalarray(1)
For I = 0 To UBound(listarray)
    If listarray(I, 4) = "" Then
        If I < txtlookup.UBound Then
            If txtlookup(I + 1).text <> "" Then
                cCond = Replace(listarray(I, 1), "%cFilter%", FixString(txtlookup(I + 1).text))
                cCond = FixMulti(cCond, txtlookup(I + 1).text)
                cCond = FixValue(cCond, txtlookup(I + 1).text)
                cCond = FixZero(cCond, txtlookup(I + 1).text)
                If UBound(listarray, 2) <= 4 Then cCond = FixDate(cCond, txtlookup(I + 1).text, "=") Else cCond = FixDate(cCond, txtlookup(I + 1).text, IIf(IsEmpty(listarray(I, 5)), "=", listarray(I, 5)))
                cCond = Replace(cCond, "cFilter", txtlookup(I + 1).text)
                cString = cString & Space(1) & turn(cString) & Space(1) & cCond
            End If
        End If
    ElseIf I < cmbLookup.UBound Then
       If cmbLookup(I + 1).MatchedWithList Then
            cCond = Replace(listarray(I, 1), "cFilter", cmbLookup(I + 1).BoundText)
            cCond = FixMulti(cCond, cmbLookup(I + 1).BoundText)
            cCond = FixValue(cCond, cmbLookup(I + 1).BoundText)
            cCond = FixZero(cCond, cmbLookup(I + 1).BoundText)
            cString = cString & Space(1) & turn(cString) & Space(1) & cCond
        End If
    End If
Next
cString = cString & Space(1) & Generalarray(2)
Set Ado1.Recordset = myRecordSet(cString, GetCon)
If Not IsEmpty(aAddRow) Then
    grid1.AddItem "", 1
    grid1.TextMatrix(1, Val(retFlag(aAddRow, "col"))) = retFlag(aAddRow, "text") & ""
    grid1.Cell(flexcpBackColor, 1, 0, , grid1.Cols - 1) = &HE0E0E0
End If
Fixgrd
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Handlecontrols()
End Sub
Private Sub Fixgrd()
For I = 0 To grid1.Cols - 1
   grid1.TextMatrix(0, I) = GrdArray(I, 0)
   grid1.ColWidth(I) = GrdArray(I, 1)
   grid1.ColAlignment(I) = flexAlignRightCenter
   nwidth = nwidth + grid1.ColWidth(I)
   If UBound(GrdArray, 2) = 2 Then
        If GrdArray(I, 2) = "d" Then grid1.ColDataType(I) = flexDTDate
   End If
Next
grid1.Width = nwidth + 400
Me.Width = grid1.Width + 400
lblCount.Caption = IIf(grid1.Rows = 1, "áÇ ÊæÌÏ ÓÌáÇÊ", "ÚÏÏ ÇáÓÌáÇÊ : " & grid1.Rows - 1)
lblCount.Left = Me.Width - lblCount.Width - 300
End Sub
Private Sub LoadControls()
nVSpace = 420
nFrame = Frame2.Height
For I = 0 To UBound(listarray)
    nRow = nRow + 1
    Frame2.Height = nFrame + (nVSpace * (nRow - 1))
    If listarray(I, 4) = "" Then
        Load txtlookup(nRow)
        txtlookup(nRow).Visible = True
        txtlookup(nRow).Top = txtlookup(0).Top + (nVSpace * (nRow - 1))
        If mySplit(listarray(I, 3), 1, ":") = "last_clicked" Then
            txtlookup(nRow).text = RetSetting("SEARCH_LIST_" & mySplit(listarray(I, 3), 2, ":"), TempSave(Generalarray(0), sid))
        End If
    Else
        Load cmbLookup(nRow)
        cmbLookup(nRow).Visible = True
        cmbLookup(nRow).Top = cmbLookup(0).Top + (nVSpace * (nRow - 1))
        Load data2(nRow)
        data2(nRow).ConnectionString = strCon
        data2(nRow).RecordSource = listarray(I, 2)
        Set cmbLookup(nRow).RowSource = data2(nRow)
        cmbLookup(nRow).BoundColumn = listarray(I, 3)
        cmbLookup(nRow).ListField = listarray(I, 4)
        If mySplit(listarray(I, 5), 1, ":") = "last_clicked" Then
            cmbLookup(nRow).BoundText = RetSetting("SEARCH_LIST_" & mySplit(listarray(I, 5), 2, ":"), TempSave(Generalarray(0), sid))
        Else
            cmbLookup(nRow).BoundText = listarray(I, 5)
        End If
        If Not cmbLookup(nRow).MatchedWithList Then cmbLookup(nRow).BoundText = ""
    End If
    Load Label1(nRow)
    Label1(nRow).Top = Label1(0).Top + (nVSpace * (nRow - 1))
    Label1(nRow).Caption = listarray(I, 0) & " :"
    lblWidth = IIf(lblWidth < Label1(nRow).Width, Label1(nRow).Width, lblWidth)
Next
If nRow >= 2 Then Me.Height = Me.Height + (nVSpace * (nRow - 1))
For I = 1 To Label1.Count - 1
    If listarray(I - 1, 4) = "" Then
        txtlookup(I).Width = Frame2.Width - (lblWidth + 400)
        Label1(I).Left = txtlookup(I).Left + 100 + txtlookup(I).Width
        If listarray(I - 1, 2) <> "" Then txtlookup(I).text = listarray(I - 1, 2)
    Else
        cmbLookup(I).Width = Frame2.Width - (lblWidth + 400)
        Label1(I).Left = cmbLookup(I).Left + 100 + cmbLookup(I).Width
    End If
    Label1(I).Caption = ArbString(Label1(I).Caption)
    Label1(I).Left = Label1(I).Left
    Label1(I).Top = Label1(I).Top
    Label1(I).Visible = True
Next
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And bEnterwork Then
    KeyCode = 0
    grid1_DblClick
End If
End Sub

Private Sub txtlookup_Change(Index As Integer)
If xEnter.Value = 0 Then myLoadGrd
txtlookup(Index).Tag = txtlookup(Index).text
End Sub
Private Sub txtlookup_GotFocus(Index As Integer)
myGotFocus txtlookup(Index)
End Sub
Private Function FixString(pString)
aString = Split(Trim(pString), " ")
For I = 0 To UBound(aString)
    If Trim(aString(I)) <> "" Then FixString = FixString & " " & Trim(aString(I))
Next
FixString = "%" & Replace(Trim(FixString), " ", "%") & "%"
End Function
Private Sub txtlookup_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0
End Sub
Private Function FixMulti(ByVal cString, cSearch) As String
Dim nFound As Long, nFound2 As Long, aString As Variant, cField As String
FixMulti = cString
For I = 1 To Len(FixMulti)
    If xBegin.Value = 0 Then
        nFound = InStr(1, FixMulti, "%%")
        If nFound = 0 Then Exit Function
        nFound2 = InStr(nFound + 3, FixMulti, "%%")
        cField = Mid(FixMulti, nFound + 2, nFound2 - (nFound + 2))
        aString = Split(Trim(cSearch), " ")
        cString2 = ""
        For i2 = 0 To UBound(aString)
            If Trim(aString(i2)) <> "" Then cString2 = cString2 & IIf(cString2 = "", "", " and ") & cField & " Like " & "'%" & aString(i2) & "%'"
        Next
        FixMulti = Replace(FixMulti, "%%" & cField & "%%", "(" & cString2 & ")")
    Else
        nFound = InStr(1, FixMulti, "%%")
        If nFound = 0 Then Exit Function
        nFound2 = InStr(nFound + 3, FixMulti, "%%")
        cField = Mid(FixMulti, nFound + 2, nFound2 - (nFound + 2))
        FixMulti = Replace(FixMulti, "%%" & cField & "%%", "(" & cField & " Like " & MyParn(cSearch & "%") & ")")
    End If
Next
End Function
Private Function FixValue(ByVal cString, cSearch) As String
Dim cSign As String, nAfter As Integer
FixValue = cString
For I = 1 To Len(FixValue)
    nFound = InStr(1, FixValue, "**")
    
    If nFound = 0 Then Exit Function
    nFound2 = InStr(nFound + 3, FixValue, "**")
    nAfter = nFound2 + 1
    
    If nAfter < Len(cString) Then
        If Mid(cString, nAfter + 1, 1) = ">" Or Mid(cString, nAfter + 1, 1) = "<" Then cSign = Mid(cString, nAfter + 1, 1)
        If nAfter + 1 < Len(cString) Then If Mid(cString, nAfter + 2, 1) = "=" Then cSign = cSign & "="
    End If
    
    cField = Mid(FixValue, nFound + 2, nFound2 - (nFound + 2))

    If IsNumeric(cSearch) Then
        cReplace = cField & Space(1) & IIf(cSign = "", "=", cSign) & Space(1) & Val(cSearch)
    Else
         cReplace = "1 = 2"
    End If
    FixValue = Replace(FixValue, "**" & cField & "**" & cSign, "(" & cReplace & ")")
Next
End Function
Private Function FixZero(ByVal cString, cSearch) As String
Dim nAfter As Integer, nZero As Integer
FixZero = cString
For I = 1 To Len(FixZero)
    nFound = InStr(1, FixZero, "@@")

    If nFound = 0 Then Exit Function
    nFound2 = InStr(nFound + 3, FixZero, "@@")
    nAfter = nFound2 + 1

    If nAfter < Len(cString) Then
        nZero = Val(Mid(cString, nAfter + 1, 2))
    End If

    cField = Mid(FixZero, nFound + 2, nFound2 - (nFound + 2))

    cReplace = cField & Space(1) & " = " & Space(1) & MyParn(RetZero(cSearch, nZero))
    FixZero = Replace(FixZero, "@@" & cField & "@@" & nZero, "(" & cReplace & ")")
Next
End Function
Private Function FixDate(ByVal cString, cSearch, pSign) As String
FixDate = cString
For I = 1 To Len(FixDate)
    cString2 = ""
    nFound = InStr(1, FixDate, "##")
    If nFound = 0 Then Exit Function
    nFound2 = InStr(nFound + 3, FixDate, "##")
    cField = Mid(FixDate, nFound + 2, nFound2 - (nFound + 2))
    
    If IsDate(cSearch) Then
        If DateValue(cSearch) >= DateValue(myFormat("01-01-1753")) Then
            cString2 = cField & " " & pSign & " " & DateSq(cSearch)
        Else
            cString2 = "(1 = 3)"
        End If
    Else
        cString2 = "(1 = 3)"
    End If
    FixDate = Replace(FixDate, "##" & cField & "##", cString2)
Next
End Function
Private Function SetValue() As Variant
Dim cField As String
cField = "search_enter:" & Generalarray(0).Name
aValue = AddFlag(Empty, "enter", RetSetting(cField, TempSave(Generalarray(0), sid)) = "TRUE")

cField = "search_begin:" & Generalarray(0).Name
aValue = AddFlag(aValue, "begin", RetSetting(cField, TempSave(Generalarray(0), sid)) = "TRUE")
End Function
Private Sub txtlookup_LostFocus(Index As Integer)
myLostFocus txtlookup(Index)
End Sub

Private Sub xEnter_Click()
bEnter = xEnter.Value = 1
End Sub

Private Sub xEnter_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cField = "search_enter:" & Generalarray(0).Name
addSetting cField, IIf(xEnter.Value = 1, "TRUE", "FALSE"), TempSave(Generalarray(0), sid)
End Sub
Private Sub xBegin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
cField = "search_begin:" & Generalarray(0).Name
addSetting cField, IIf(xBegin.Value = 1, "TRUE", "FALSE"), TempSave(Generalarray(0), sid)
End Sub
Private Sub LoadArrays()
Generalarray = searchArray(0)
listarray = searchArray(1)
GrdArray = searchArray(2)
End Sub
Function TempSave(pform As Variant, Optional sFlag As String = "", Optional sExt As String = "txt") As String
TempSave = tempPath & turn(tempPath, "\") & pform.Name & turn(sFlag, "_" & sFlag) & "." & sExt
End Function

