VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Search5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈” ⁄·«„"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11640
   Tag             =   "Factory"
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   45
      TabIndex        =   2
      Top             =   7155
      Width           =   5115
      Begin VB.TextBox txtlookup 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   180
         Visible         =   0   'False
         Width           =   3165
      End
      Begin MSDataListLib.DataCombo cmbLookup 
         Height          =   390
         Index           =   0
         Left            =   135
         TabIndex        =   4
         Top             =   180
         Visible         =   0   'False
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   688
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   270
         Visible         =   0   'False
         Width           =   660
      End
   End
   Begin MSAdodcLib.Adodc Ado1 
      Height          =   330
      Left            =   300
      Top             =   1125
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
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   7845
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   714
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7080
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   11535
      _cx             =   20346
      _cy             =   12488
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arabic Transparent"
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
      AllowSelection  =   -1  'True
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
End
Attribute VB_Name = "Search5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim Generalarray, listarray, GrdArray
Dim cString As String
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdGo_Click()
Fillgrd
FIXGRID
If Grid1.Rows > 1 Then
    Grid1.SetFocus
   ' grid1.Row = 1
End If
End Sub
Private Sub cmbLookup_Click(Index As Integer, Area As Integer)
If Area = 2 Then Fillgrd
End Sub
Private Sub cmbLookup_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not cmbLookup(Index).MatchedWithList Then cmbLookup(Index).BoundText = ""
    Fillgrd
End If
End Sub
Private Sub cmbLookup_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Then cmbLookup(Index).BoundText = ""
End Sub
Private Sub cmbLookup_LostFocus(Index As Integer)
On Error Resume Next
If Index = UBound(listarray) + 1 Then
    Grid1.SetFocus
    If Grid1.Rows > 1 Then Grid1.Row = 1
End If
Err.Clear
Exit Sub
End Sub

Private Sub cmbLookup_Validate(Index As Integer, Cancel As Boolean)
If Not cmbLookup(Index).MatchedWithList Then
    cmbLookup(Index).BoundText = ""
    Fillgrd
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    If UBound(Generalarray) < 5 Then
        Unload Me
    Else
        If Generalarray(5) Then Me.Hide
        If Not Generalarray(5) Then Unload Me
    End If
End If
End Sub
Private Sub Form_Activate()
'If txtlookup(1).Visible Then txtlookup(1).SetFocus
End Sub
Private Sub Form_Load()
openCon con
Set Grid1.DataSource = Ado1
Grid1.ExplorerBar = flexExSort
StatusBar1.Panels(1).Width = 2500
Ado1.ConnectionString = strCon
Ado1.CommandType = adCmdText
Generalarray = searchArray(0)
listarray = searchArray(1)
GrdArray = searchArray(2)

Frame2.Width = Generalarray(3)
Grid1.Cols = UBound(GrdArray) + 1
LoadControls

If UBound(Generalarray) = 3 Then
    Fillgrd
Else
    If UBound(Generalarray) >= 4 Then If Not Generalarray(4) Then Fillgrd
End If

Handlecontrols
End Sub

Private Sub grid1_DblClick()
    If Grid1.Row > 0 Then Generalarray(0).myProc
End Sub
Private Sub Grid1_GotFocus()
    If Grid1.Row = 0 And Grid1.Rows > 1 Then Grid1.Select 1, 0
End Sub
Private Sub Grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grid1_DblClick
End If
End Sub
Sub Fillgrd()
'On Error GoTo myerror
cString = Generalarray(1)
For i = 0 To UBound(listarray)
   If listarray(i, 4) = "" Then
        If txtlookup(i + 1).text <> "" Then
            cCond = Replace(listarray(i, 1), "%cFilter%", FixString(txtlookup(i + 1).text))
            cCond = FixMulti(cCond, txtlookup(i + 1).text)
            cCond = FixValue(cCond, txtlookup(i + 1).text)
            cCond = FixZero(cCond, txtlookup(i + 1).text)
            If UBound(listarray, 2) <= 4 Then cCond = FixDate(cCond, txtlookup(i + 1).text, "=") Else cCond = FixDate(cCond, txtlookup(i + 1).text, IIf(IsEmpty(listarray(i, 5)), "=", listarray(i, 5)))
            cCond = Replace(cCond, "cFilter", txtlookup(i + 1).text)
            cString = cString & Space(1) & turnFound(cString) & Space(1) & cCond
        End If
    Else
         If cmbLookup(i + 1).BoundText <> "" Then
            cCond = Replace(listarray(i, 1), "cFilter", cmbLookup(i + 1).BoundText)
            cString = cString & Space(1) & turnFound(cString) & Space(1) & cCond
        End If
    End If
Next
cString = cString & Space(1) & Generalarray(2)
Ado1.RecordSource = cString
Ado1.Refresh
FIXGRID
Exit Sub
myerror:
MsgBox "«œŒ«· ‰’ €Ì— „‰«”»"
End Sub
Private Sub Handlecontrols()
End Sub
Private Sub FIXGRID()
For i = 0 To Grid1.Cols - 1
   Grid1.TextMatrix(0, i) = GrdArray(i, 0)
   Grid1.ColWidth(i) = GrdArray(i, 1)
   Grid1.ColAlignment(i) = 6
   nwidth = nwidth + Grid1.ColWidth(i) + 75
   If UBound(GrdArray, 2) = 2 Then
        If GrdArray(i, 2) = "d" Then Grid1.ColDataType(i) = flexDTDate
   End If
Next
Grid1.Width = nwidth + 200
Me.Width = Grid1.Width + 400
'Label2.Caption = IIf(Grid1.Rows = 1, "·«  ÊÃœ ”Ã·« ", "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & Grid1.Rows - 1)
StatusBar1.Panels(1).text = IIf(Grid1.Rows = 1, "·«  ÊÃœ ”Ã·« ", "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & Grid1.Rows - 1)
End Sub
Private Sub LoadControls()
nVSpace = 375
nFrame = Frame2.Height
For i = 0 To UBound(listarray)
    nRow = nRow + 1
    Frame2.Height = nFrame + (nVSpace * (nRow - 1))
    If listarray(i, 4) = "" Then
        Load txtlookup(nRow)
        txtlookup(nRow).Visible = True
        txtlookup(nRow).Top = txtlookup(0).Top + (nVSpace * (nRow - 1))
    Else
        Load cmbLookup(nRow)
        cmbLookup(nRow).Visible = True
        cmbLookup(nRow).Top = cmbLookup(0).Top + (nVSpace * (nRow - 1))
        Load data2(nRow)
        data2(nRow).ConnectionString = strCon
        data2(nRow).RecordSource = listarray(i, 2)
        Set cmbLookup(nRow).RowSource = data2(nRow)
        cmbLookup(nRow).BoundColumn = listarray(i, 3)
        cmbLookup(nRow).ListField = listarray(i, 4)
        cmbLookup(nRow).BoundText = listarray(i, 5)
    End If
    
    Load Label1(nRow)
    Label1(nRow).Visible = True
    Label1(nRow).Top = Label1(0).Top + (nVSpace * (nRow - 1))
    Label1(nRow).Caption = listarray(i, 0)
    lblWidth = IIf(lblWidth < Label1(nRow).Width, Label1(nRow).Width, lblWidth)
Next
If nRow >= 2 Then
    Me.Height = Me.Height + (nVSpace * (nRow - 1))
End If
For i = 1 To Label1.Count - 1
    If listarray(i - 1, 4) = "" Then
        txtlookup(i).Width = Frame2.Width - (lblWidth + 400)
        Label1(i).Left = txtlookup(i).Left + 100 + txtlookup(i).Width
    Else
        cmbLookup(i).Width = Frame2.Width - (lblWidth + 400)
        Label1(i).Left = cmbLookup(i).Left + 100 + cmbLookup(i).Width
    End If
Next
End Sub
Private Sub txtlookup_Change(Index As Integer)
Fillgrd
End Sub

Private Sub txtlookup_GotFocus(Index As Integer)
txtlookup(Index).SelStart = 0
txtlookup(Index).SelLength = Len(txtlookup(Index).text)
End Sub
Private Sub txtlookup_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then Grid1.SetFocus
End Sub
Private Function FixString(cString)
aString = Split(Trim(cString), " ")
For i = 0 To UBound(aString)
    If Trim(aString(i)) <> "" Then FixString = FixString & " " & Trim(aString(i))
Next
FixString = "%" & Replace(Trim(FixString), " ", "%") & "%"
End Function
Private Sub txtlookup_LostFocus(Index As Integer)
'On Error Resume Next
'If Index = UBound(listarray) + 1 Then
'    Grid1.SetFocus
'    Grid1.Row = 1
'End If
'Err.Clear
End Sub
Private Function FixMulti(ByVal cString, cSearch) As String
FixMulti = cString
For i = 1 To Len(FixMulti)
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
Next
End Function
Private Function FixValue(ByVal cString, cSearch) As String
Dim cSign As String, nAfter As Integer
FixValue = cString
For i = 1 To Len(FixValue)
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
For i = 1 To Len(FixZero)
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
For i = 1 To Len(FixDate)
    cString2 = ""
    nFound = InStr(1, FixDate, "##")
    If nFound = 0 Then Exit Function
    nFound2 = InStr(nFound + 3, FixDate, "##")
    cField = Mid(FixDate, nFound + 2, nFound2 - (nFound + 2))
    
    If IsDate(cSearch) Then
        If DateValue(cSearch) >= DateValue("01-01-1753") Then
            cString2 = cField & " " & pSign & " " & DateSq(Format(cSearch, "dd/mm/yyyy"))
        Else
            cString2 = "(1 = 3)"
        End If
    Else
        cString2 = "(1 = 3)"
    End If
    FixDate = Replace(FixDate, "##" & cField & "##", cString2)
Next
End Function

