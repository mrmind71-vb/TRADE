VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form InvoiceTransfrm 
   Caption         =   "»Ì«‰«  «· ÿ»Ìﬁ ⁄·Ì «·„Êﬁ⁄"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   4770
      TabIndex        =   0
      Top             =   6300
      Width           =   5820
      Begin VB.CommandButton CmdExit 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "InvoiceTransfrm.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1500
      End
      Begin Threed.SSCommand cmdItemsUpdate 
         Height          =   510
         Left            =   3735
         TabIndex        =   1
         Top             =   135
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "√—”«· «·›Ê« Ì—"
         ButtonStyle     =   3
      End
      Begin Threed.SSCommand cmdFixMobileSetting 
         Height          =   510
         Left            =   1620
         TabIndex        =   4
         Top             =   135
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
         Enabled         =   0   'False
         ActiveColors    =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "÷»ÿ »Ì«‰«  «·„Êﬁ⁄"
         ButtonStyle     =   3
      End
   End
   Begin MSComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   7110
      Visible         =   0   'False
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   6180
      Left            =   4770
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   90
      Width           =   5775
      _cx             =   10186
      _cy             =   10901
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
      Rows            =   0
      Cols            =   1
      FixedRows       =   0
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
   Begin VSFlex7Ctl.VSFlexGrid grdMissing 
      Height          =   6900
      Left            =   90
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   90
      Width           =   4650
      _cx             =   8202
      _cy             =   12171
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
      Cols            =   3
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
End
Attribute VB_Name = "InvoiceTransfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bSend As Boolean
Dim con As New ADODB.Connection
Dim bAct As Boolean
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdFixMobileSetting_Click()
invoice_send_settingfrm.Show 1
End Sub
Private Sub cmdItemsUpdate_Click()
sendinvoicesBranch
End Sub
Private Sub Form_Activate()
If Not bAct Then
    bAct = True
    If bSend Then
        If sendinvoicesBranch Then Unload Me
    End If
End If
End Sub
Private Sub Form_Load()
openCon con

Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

grid1.ColWidth(0) = grid1.Width - 400
grid1.ColAlignment(0) = flexAlignRightCenter

getMissingDays
FixGrdMissing
End Sub
Public Function AddNote(sNote As String, Optional color As Long = 0)
grid1.AddItem sNote
grid1.Select grid1.Rows - 1, 0
grid1.ShowCell grid1.Rows - 1, 0
If color > 0 Then
    grid1.Cell(flexcpBackColor, grid1.Rows - 1, 0, grid1.Rows - 1, grid1.Cols - 1) = color
End If
DoEvents
End Function
Private Sub FixGrdMissing()
With grdMissing
    .TextMatrix(0, 0) = "«·ÌÊ„"
    .TextMatrix(0, 1) = "«·ÌÊ„"
    .TextMatrix(0, 2) = "«·≈Ã„«·Ì"
    .ColWidth(0) = 1200
    .ColWidth(0) = 1300
    .ColWidth(0) = 1200
    
    .MergeCells = flexMergeFixedOnly
    .MergeRow(0) = True
    .ColAlignment(0) = flexAlignRightCenter
    .ColAlignment(1) = flexAlignRightCenter
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter

    .SubtotalPosition = flexSTAbove
    .Subtotal flexSTSum, -1, 2, "#0.00", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ì"

End With
End Sub
Private Function SendDay(pDate As String) As Boolean
Dim jArray As New ChilkatJsonArray
Dim sResponse As String
Me.MousePointer = vbHourglass

SendDay = SendInvoices(con, Me, myFormat(pDate), jArray, sResponse)
Me.MousePointer = vbNormal
End Function
Private Function getMissingDays() As Boolean
Me.MousePointer = vbHourglass

grid1.Rows = 0


Dim json As String
json = GetMissing(con, Me)

Dim jobject As New ChilkatJsonObject
success = jobject.Load(json)

If jobject.IsNullOf("count") Then Exit Function

nCount = Val(jobject.StringOf("count"))

Dim jArray As New ChilkatJsonArray
If Not jobject.IsNullOf("entityCollection") Then
   Set jArray = jobject.ArrayOf("entityCollection")
    For i = 0 To jArray.Size - 1
        grdMissing.AddItem ""
        grdMissing.TextMatrix(grdMissing.Rows - 1, 1) = myFormat_p(jArray.StringAt(i))
        grdMissing.TextMatrix(grdMissing.Rows - 1, 0) = arbDay(myFormat(grdMissing.TextMatrix(grdMissing.Rows - 1, 1)))
        grdMissing.TextMatrix(grdMissing.Rows - 1, 2) = MyFuncValue("[dbo].[fn_invoices_branch_total_day]", con, MyParn("019"), MyParn(myFormat(jArray.StringAt(i))))
    Next
End If

If grdMissing.Rows > 1 Then
    Me.AddNote " „ «÷«›… " & grdMissing.Rows - 1 & "ÌÊ„ ·„ Ì „ «—”«· „»Ì⁄« Â"
End If
Me.MousePointer = vbNormal
End Function
Private Function sendinvoicesBranch() As Boolean
Dim i As Long
If grid1.Rows = 1 Then
    Inform "·«  ÊÃœ ÌÊ„Ì… „»Ì⁄«  ·«—”«·Â«"
    sendinvoicesBranch = True
    Exit Function
End If
With grdMissing
For i = 2 To .Rows - 1
    .ShowCell i, 0
    .Select i, 0
    sendinvoicesBranch = SendDay(myFormat(.TextMatrix(i, 1)))
Next
End With
End Function
Private Sub Form_Unload(Cancel As Integer)
Set InvoiceTransfrm = Nothing
End Sub
