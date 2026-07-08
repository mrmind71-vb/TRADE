VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ModelMove 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Õ—ﬂ… „ÊœÌ·"
   ClientHeight    =   9675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15270
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
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   15270
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "ÿ»«⁄… Õ—ﬂ… «·„ÊœÌ·"
      Height          =   420
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   810
      Width           =   3210
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   3600
      Top             =   810
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   6030
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   45
      Width           =   9105
      Begin VB.TextBox xBarCode 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   630
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   990
         Width           =   1770
      End
      Begin VB.TextBox xdate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5580
         MaxLength       =   15
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   900
         Width           =   2310
      End
      Begin VB.TextBox xItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5580
         MaxLength       =   15
         TabIndex        =   0
         Top             =   180
         Width           =   2310
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   5580
         TabIndex        =   1
         Top             =   547
         Width           =   2310
         _ExtentX        =   4075
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xscal 
         Height          =   315
         Left            =   3330
         TabIndex        =   17
         Top             =   540
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xColor 
         Height          =   315
         Left            =   630
         TabIndex        =   19
         Top             =   540
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ »«—ﬂÊœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2475
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "·Ê‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   585
         Width           =   315
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„ﬁ«”"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4815
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   585
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "„‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   945
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   630
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8100
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   945
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   630
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   4695
      End
   End
   Begin VB.TextBox LastOne 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   -555
      MaxLength       =   2
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1920
      Width           =   405
   End
   Begin VSFlex7Ctl.VSFlexGrid invGrid 
      Height          =   7215
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1665
      Width           =   15045
      _cx             =   26538
      _cy             =   12726
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Simplified Arabic"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
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
      TabBehavior     =   1
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
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   90
      Width           =   3210
      Begin VB.CommandButton CmdExit 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Œ—ÊÃ "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   180
         Width           =   1365
      End
      Begin VB.CommandButton CmdGo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "≈ŸÂ«— «·Õ—ﬂ…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      Height          =   600
      Left            =   3330
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   90
      Width           =   2490
      Begin VB.Label xBal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "«·—’Ìœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   525
      End
   End
End
Attribute VB_Name = "ModelMove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ItemTable As New ADODB.Recordset
Dim con As New ADODB.Connection
Sub Fillgrd()
Dim loctable As New ADODB.Recordset
cString = "select file1_11.*,file1_12.desca ,file3_10.desca as custDesca,file4_10.desca as supDesca,file0_40.desca as storeDesca" & _
          " From ((((file1_11 inner join file1_10 on file1_11.item = file1_10.item) left join file3_10 on file1_11.codecust = file3_10.code) left join file4_10 on file1_11.codesup = file4_10.code) left join file1_12 on file1_11.type = file1_12.code) left join file0_40 on file1_11.code = file0_40.code"
cString = cString & turnFound2(cString) & " file1_11.item = " & MyParn(xItem.Text)

If IsDate(xdate.Text) Then
    cString = cString & turn(cString) & " file1_11.date >= " & DateSq(xdate.Text)
End If

If Trim(xStore.BoundText) <> "" Then
    cWhere = cWhere & turn(cString) & " file1_11.store = " & MyParn(xStore.BoundText)
End If

cString = cString & cWhere & " Order by Date,[out],FILE1_12.[ORDER]"

With invGrid
    .Rows = 1
    If IsDate(xdate.Text) Then
       cString2 = "Select sum([IN] - OUT ) as Balance from file1_11 where file1_11.item = " & MyParn(xItem.Text) & _
                  " and file1_11.date < " & DateSq(xdate.Text) & cWhere
       nPrevious = Val(GetDesca(cString2))
       If nPrevious <> 0 Then
            .AddItem ""
            .TextMatrix(.Rows - 1, 0) = "—’Ìœ ﬁ»· " & xdate.Text
            .TextMatrix(.Rows - 1, 3) = nPrevious
       End If
    End If

    loctable.Open cString, con, adOpenStatic, adLockReadOnly, adcdmtext

    Do Until loctable.EOF
         invGrid.AddItem ""
         nPrevious = nPrevious + Val(loctable!In & "") - Val(loctable!out & "")
        .TextMatrix(.Rows - 1, 0) = loctable!Desca & " " & loctable!StoreDesca
        .TextMatrix(.Rows - 1, 1) = Format(Val(loctable!out & ""), "#0.00")
        .TextMatrix(.Rows - 1, 2) = Format(Val(loctable!In & ""), "#0.00")
        .TextMatrix(.Rows - 1, 3) = Format(nPrevious, "#0.00")
        .TextMatrix(.Rows - 1, 4) = Format(loctable!Date, "yyyy/mm/dd")
        .TextMatrix(.Rows - 1, 5) = Format(Val(loctable!price & ""), "fixed")
        .TextMatrix(.Rows - 1, 6) = loctable!doc_ID & ""
        .TextMatrix(.Rows - 1, 7) = IIf(Trim(loctable!CustDesca) & "" = "", loctable!SUPDESCA & "", loctable!CustDesca)
        .TextMatrix(.Rows - 1, 8) = loctable!Type & ""
        If .TextMatrix(.Rows - 1, 8) = "8" Then
            .TextMatrix(.Rows - 1, 7) = GetDesca("Select Remark from file1_82H WHERE FILE1_82h.doc_no = " & MyParn(.TextMatrix(.Rows - 1, 6)))
        End If
        loctable.MoveNext
    Loop
End With
End Sub
Sub myProc()
    ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Search3.Hide
End Sub
Function MYVALID() As Boolean
If xItem.Text = "" Then
    MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
    Exit Function
End If
ItemTable.Find " item = " & MyParn(xItem.Text), , adSearchForward, adBookmarkFirst
If ItemTable.EOF Then
    MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
    Exit Function
End If
If (Not IsDate(xdate.Text)) And Trim(xdate.Text) <> "" Then
    MsgBox "«· «—ÌŒ €Ì— ’«·Õ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub cmdcorect_Click()

End Sub
Private Sub CmdGo_Click()
If Not MYVALID Then Exit Sub
Fillgrd
xBal.Caption = Format(Val(invGrid.TextMatrix(invGrid.Rows - 1, 3)), "#0.00")
End Sub
Private Sub cmdExit_Click()
Unload StoreMove
End Sub

Private Sub Command1_Click()
Load rpitem2
rpitem2.xItem.Text = xItem.Text
rpitem2.xdate1.Text = xdate.Text
rpitem2.xStore.BoundText = xStore.BoundText
rpitem2.Show 1
End Sub
Private Sub Form_Load()
openCon con
ItemTable.Open "file1_10", con, adOpenStatic, adLockReadOnly, adCmdTable
data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FILE0_40"
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

With invGrid
invGrid.Cols = 9
.TextMatrix(0, 0) = "»Ì«‰"
.TextMatrix(0, 1) = "’«œ—"
.TextMatrix(0, 2) = "Ê«—œ"
.TextMatrix(0, 3) = "—’Ìœ"
.TextMatrix(0, 4) = " «—ÌŒ"
.TextMatrix(0, 5) = "”⁄—"
.TextMatrix(0, 6) = "„” ‰œ"
.TextMatrix(0, 7) = "≈”„"


invGrid.ColWidth(0) = 3000
invGrid.ColWidth(1) = 1000
invGrid.ColWidth(2) = 1000
invGrid.ColWidth(3) = 1000
invGrid.ColWidth(4) = 1500
invGrid.ColWidth(5) = 1000
invGrid.ColWidth(6) = 1700
invGrid.ColWidth(7) = 3500
invGrid.ColHidden(8) = True
End With
For I = 0 To invGrid.Cols - 1
    invGrid.ColAlignment(I) = flexAlignRightCenter
Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Unload Search3
Err.Clear
End Sub

Private Sub invGrid_dblClick()
    'bEdit = RetSec("xPurchase")
    'If bEdit Then
    Select Case invGrid.TextMatrix(invGrid.Row, 8)
    Case "2", "7"
        Load Purchasefrm
        Purchasefrm.myPublic = IIf(invGrid.TextMatrix(invGrid.Row, 8) = "2", 0, 1)
        Purchasefrm.myproc2 invGrid.TextMatrix(invGrid.Row, 6)
        Purchasefrm.Show 1
    Case "3", "6"
        SalesFrm.myPublic = IIf(invGrid.TextMatrix(invGrid.Row, 8) = "6", 0, 1)
        SalesFrm.sDoc_No = invGrid.TextMatrix(invGrid.Row, 6)
        SalesFrm.Show 1
    End Select
    'Else
    '    MsgBox "·Ì” ·œÌﬂ «·’·«ÕÌ… ·œŒÊ· „·› «·„‘ —Ì« "
    '    Exit Sub
    'End If
End Sub

Private Sub xdate_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub
Private Sub xITEM_Change()
    invGrid.Rows = 1
    CmdGo.Enabled = Trim(xItem.Text) <> ""
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdGo_Click
End Sub
Private Sub xItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then SendKeys "{Tab}"
    If KeyCode = 112 Then ModelLookupAll Me, Search3
End Sub
Private Sub xitem_LostFocus()
If Not CmdGo.Enabled And xStore.BoundText <> "" Then CmdGo.Enabled = True
xDesca.Caption = ""
If Trim(xItem.Text) = "" Then Exit Sub
ItemTable.Find "item = " & MyParn(xItem.Text), , adSearchForward, adBookmarkFirst
If Not ItemTable.EOF Then xDesca.Caption = ItemTable!Desca & ""
End Sub
Private Sub xStore_Click(Area As Integer)
If Not CmdGo.Enabled Then CmdGo.Enabled = True
End Sub
Private Sub xStore_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then CmdGo_Click
End Sub
Private Sub xstore_LostFocus()
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub

