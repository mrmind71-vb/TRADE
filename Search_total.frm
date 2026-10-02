VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Search_total 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "≈” ⁄·«„"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
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
   ScaleHeight     =   8325
   ScaleWidth      =   11790
   Tag             =   "Factory"
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      ScaleHeight     =   435
      ScaleWidth      =   11730
      TabIndex        =   0
      Top             =   7830
      Width           =   11790
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   7335
         ScaleHeight     =   465
         ScaleWidth      =   4380
         TabIndex        =   2
         Top             =   0
         Width           =   4380
         Begin Threed.SSCommand cmdPrint 
            Height          =   285
            Left            =   3510
            TabIndex        =   4
            Top             =   90
            Visible         =   0   'False
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   503
            _Version        =   196610
            ForeColor       =   4210752
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
            Caption         =   "Print"
            ButtonStyle     =   3
         End
         Begin Threed.SSCommand cmdExcel 
            Height          =   285
            Left            =   2745
            TabIndex        =   5
            Top             =   90
            Visible         =   0   'False
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   196610
            ForeColor       =   32768
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
            Caption         =   "Excel"
            ButtonStyle     =   3
         End
      End
      Begin VB.Label lblCount 
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   45
         Width           =   2760
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Bindings        =   "Search_total.frx":0000
      Height          =   7620
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   11625
      _cx             =   20505
      _cy             =   13441
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
      RowHeightMax    =   300
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
Attribute VB_Name = "Search_total"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Generalarray, listarray, GrdArray
Public sid As String, aFormat As Variant, bTotal As Boolean, bPrint As Boolean
Dim aValue As Variant
Public aFilter As Variant
Dim cString As String, bAtChange As Boolean
Public sFlag As String, bEnter As Boolean
Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdGo_Click()
myloadgrd
Fixgrd
If grid1.Rows > 1 Then
    grid1.SetFocus
End If
End Sub
Private Sub cmdExcel_Click()
ToFileExel grid1
End Sub

Private Sub cmdFilter_Click()
If grid1.Rows > 1001 Then
    MsgBox "⁄œœ «·”ÿÊ— «ﬂ»— „‰ 1000"
    Exit Sub
End If
Dim cString As String, nCol As Long, sType As String, cFilter As String
nCol = Val(retFlag(aFilter, "col"))
sType = IIf(retFlag(aField, "type") = "", "s", retFlag(aField, "type"))
For i = 1 To grid1.Rows - 1
    cFilter = cFilter & IIf(cFilter = "", "", ",") & IIf(sType = "s", MyParn(grid1.TextMatrix(i, nCol)), grid1.TextMatrix(i, nCol))
Next
If cFilter <> "" Then
    cFilter = retFlag(aFilter, "field") & " IN(" & cFilter & ")"
End If
Generalarray(0).myproc2 cFilter
End Sub
Private Sub cmdPrint_Click()
Generalarray(0).myPrint sFlag
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
'    If UBound(Generalarray) < 5 Then
        Unload Me
'    Else
'        If Generalarray(5) Then Me.Hide
'        If Not Generalarray(5) Then Unload Me
'    End If
End If
End Sub
Private Sub Form_Activate()
On Error Resume Next
'If Generalarray(4) Then myloadgrd
End Sub
Private Sub Form_Load()
cmdPrint.Visible = bPrint
grid1.ExplorerBar = flexExSort


Generalarray = searchArray(0)
listarray = searchArray(1)
GrdArray = searchArray(2)

grid1.Cols = UBound(GrdArray) + 1


If UBound(Generalarray) = 3 Then
    myloadgrd
Else
    If UBound(Generalarray) >= 4 Then
        If Not Generalarray(4) Then myloadgrd Else Fixgrd
    End If
End If

Picture2.Left = Me.Width - Picture2.Width - 100
End Sub

Private Sub Form_Resize()
'Picture2.Width = 4155
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set Search_total = Nothing
End Sub

Private Sub grid1_DblClick()
If grid1.Row > IIf(bTotal, 1, 0) Then
    If sFlag = "" Then Generalarray(0).myProc Else Generalarray(0).myProc sFlag
End If
End Sub
Private Sub grid1_GotFocus()
For i = 0 To grid1.Cols - 1
    If Not grid1.ColHidden(i) Then Exit For
Next
If grid1.Rows > 1 Then grid1.Select 1, i
End Sub
Sub myloadgrd()
On Error GoTo myError
Dim cString As String
cString = Generalarray(1)
cString = cString & Space(1) & Generalarray(2)
Dim db As New clsDb
Set grid1.DataSource = db.myRs(cString)
Set db = Nothing
Fixgrd
Exit Sub
myError:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Fixgrd()
For i = 0 To grid1.Cols - 1
    grid1.TextMatrix(0, i) = GrdArray(i, 0)
    grid1.ColWidth(i) = GrdArray(i, 1)
    If Not retFlag(aFormat, "left") Then grid1.ColAlignment(i) = flexAlignRightCenter
    nWidth = nWidth + grid1.ColWidth(i)
    If UBound(GrdArray, 2) = 2 Then
        If GrdArray(i, 2) = "d" Then grid1.ColDataType(i) = flexDTDate
    End If
    If GrdArray(i, 3) Then
        bTotal = True
        grid1.SubtotalPosition = flexSTAbove
        grid1.Subtotal flexSTSum, -1, i, "#0.00", &HC0FFC0, vbBlack, True, "≈Ã„«·Ï"
    End If
Next

For i = 0 To grid1.Cols - 1
    If GrdArray(i, 3) Then
        If grid1.Rows > 1 Then
            grid1.TextMatrix(1, i) = mRound(grid1.ValueMatrix(1, i))
        End If
    End If
Next

grid1.Width = nWidth + 400
Me.Width = grid1.Width + 400
If retFlag(aFormat, "left") = True Then
    grid1.RightToLeft = False
    grid1.FontSize = 10
End If
lblCount.Caption = IIf(grid1.Rows < 3, "·«  ÊÃœ ”Ã·« ", "⁄œœ «·”Ã·«  «·„ÿ«»ﬁ… : " & grid1.Rows - 2)
End Sub
Private Sub LoadControls()
nVSpace = 420
nFrame = Frame2.Height
If nRow >= 2 Then
    Me.Height = Me.Height
End If
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    grid1_DblClick
End If
End Sub
