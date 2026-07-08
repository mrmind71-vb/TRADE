VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Trans_Balnce 
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Simplified Arabic"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CMD_SAVE 
      Caption         =   "Õ›Ÿ «· ÕÊÌ·"
      Height          =   495
      Left            =   2340
      RightToLeft     =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1215
      Width           =   2055
   End
   Begin VB.CommandButton CMD_E 
      Caption         =   " ÕÊÌ· ﬂ· «·—’Ìœ"
      Height          =   495
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1215
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "»ÕÀ ⁄‰ „ÊœÌ·« "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   8055
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   135
      Width           =   7125
      Begin VB.CommandButton CMD_OK 
         Caption         =   "⁄—÷ √—’œ… «·«’‰«›"
         Height          =   540
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   765
         Width           =   2055
      End
      Begin VB.TextBox xModelFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   330
         Left            =   3240
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   810
         Width           =   2715
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   360
         Left            =   3240
         TabIndex        =   2
         Top             =   315
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   360
         Left            =   180
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   315
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   855
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’‰⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   6030
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   555
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2475
      Top             =   0
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
      Caption         =   "DATA1"
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
      Height          =   7710
      Left            =   135
      TabIndex        =   0
      Top             =   1755
      Width           =   15045
      _cx             =   26538
      _cy             =   13600
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
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
      BackColorSel    =   16777152
      ForeColorSel    =   128
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   2475
      Top             =   720
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
      Caption         =   "DATA3"
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   2475
      Top             =   360
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
      Caption         =   "DATA1"
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
End
Attribute VB_Name = "Trans_Balnce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection

Private Sub CMD_E_Click()
With grid1
    For i = 1 To .Rows - 1
        .TextMatrix(i, 9) = .TextMatrix(i, 8)
    Next i
End With
End Sub
Private Sub CMD_OK_Click()
    myload
End Sub
Private Sub CMD_SAVE_Click()
    If myreplace Then
        Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
        
    End If
    On Error Resume Next
    xModelFact.SetFocus
    Err.Clear
End Sub
Private Sub Form_Load()
    con.CursorLocation = adUseClient
    con.Open strCon
    
    DATA3.ConnectionString = strCon
    DATA3.RecordSource = "SELECT CODE, DESCA FROM FACT ORDER BY DESCA "
    Set xFact.RowSource = DATA3
    xFact.ListField = "DESCA"
    xFact.BoundColumn = "CODE"
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "SELECT * FROM MOSM ORDER BY DATE DESC "
    Set xMosm.RowSource = DATA2
    xMosm.ListField = "DESCA"
    xMosm.BoundColumn = "MOSM"
    
    
    With grid1
        .Cols = 10
        .Rows = 1
        .RowHeight(0) = 1000
        .WordWrap = True
    End With
    Set grid1.DataSource = data1
    data1.ConnectionString = strCon
    grid1.Rows = 1
    FIXGRID
End Sub
Sub myload()
    On Error GoTo myerror
    cString = "SELECT FILE1_10.item , file1_10.mosm , fact.desca , file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  BAL_ITEM_S.balitem , ' ' AS Q FROM  (BAL_ITEM_S RIGHT OUTER JOIN FILE1_10 ON BAL_ITEM_S.ITEM = FILE1_10.ITEM ) INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT  WHERE STORE = " & MyParn(Transfrm.xStore1.BoundText)
    If xMosm.BoundText <> "" Then cString = cString & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xModelFact.text <> "" Then cString = cString & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xFact.BoundText <> "" Then cString = cString & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    
    
    cString = cString & " ORDER BY FILE1_10.MODEL, FILE1_10.COLOR, FILE1_10.C_SCAL "
    data1.RecordSource = cString
    data1.Refresh
    grid1.AddItem "'"
    FIXGRID
    Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Sub FIXGRID()
With grid1
    .Cols = 10
    .TextMatrix(0, 0) = "»«—ﬂÊœ"
    .TextMatrix(0, 1) = "„Ê”„"
    .TextMatrix(0, 2) = "„’‰⁄"
    .TextMatrix(0, 3) = "„ﬂ »"
    .TextMatrix(0, 4) = "—ﬁ„ „ÊœÌ·"
    .TextMatrix(0, 5) = "«·’‰›"
    .TextMatrix(0, 6) = "·Ê‰"
    .TextMatrix(0, 7) = "„ﬁ«”"
    .TextMatrix(0, 8) = "—’Ìœ " & Transfrm.xStore1.text
    .TextMatrix(0, 9) = "ﬂ„Ì… „ÕÊ·… ≈·Ï " & Transfrm.XSTORE2.text
    
    .ColWidth(0) = 1000
    .ColWidth(1) = 800
    .ColWidth(2) = 2500
    .ColWidth(3) = 1000
    .ColWidth(4) = 800
    .ColWidth(5) = 4000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    .ColWidth(8) = 1000
    .ColWidth(9) = 1000
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
    .ColDataType(8) = flexDTDouble
    .ColDataType(9) = flexDTDouble
    For i = 1 To .Rows - 1
        nFoundRow = Transfrm.grid1.FindRow(.TextMatrix(i, 0), , 0)
        If nFoundRow <> -1 Then
            grid1.TextMatrix(i, 9) = Transfrm.grid1.TextMatrix(nFoundRow, 8)
        End If
    Next i

End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
closeCon con
Set Trans_Balnce = Nothing
End Sub
Private Sub ModelFact_Lookup()

Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(2, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  MODELFACT0 , DESCA , SUPP FROM FILE1_10H WHERE MOSM = " & MyParn(xMosm.BoundText) & " AND FACT = " & MyParn(xFact.BoundText)
Generalarray(2) = "Order by MODELFACT0 "
Generalarray(3) = 6000
Generalarray(5) = True


listarray(0, 0) = "«·—ﬁ„-≈”„ "
listarray(0, 1) = "DESCA Like '%cFilter%' "


GrdArray(0, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·’‰›"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ê—œ"
GrdArray(2, 1) = 1000


searchArray = Array(Generalarray, listarray, GrdArray)
Load Search1
Search1.Caption = "«” ⁄·«„"
Search1.Show 1
End Sub
Sub myProc()
On Error GoTo myerror
ActiveControl.text = Search1.grid1.TextMatrix(Search1.grid1.Row, 0)
Unload Search1
Exit Sub
myerror:
End Sub
Private Sub xModelFact_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    If xFact.BoundText <> "" And xMosm.BoundText <> "" Then
        ModelFact_Lookup
    End If
End If
End Sub
Private Function myreplace() As Boolean
Dim cString As String, nRow As Integer, cFile As String, nItem As Integer
If Transfrm.xDoc_No.Tag = "1" Then
    myreplaceH
End If
'On Error GoTo myerror
con.BeginTrans

cDoc_No = Transfrm.xDoc_No.text

With grid1
    For nRow = 1 To .Rows - 1
        con.Execute "DELETE FROM FILE1_60 WHERE ITEM = " & Val(.TextMatrix(nRow, 0)) & " AND DOC_NO = " & MyParn(cDoc_No)
        If Val(grid1.TextMatrix(nRow, 9)) <> 0 Then
            nCost = GetDesca("SELECT COST FROM FILE1_10 WHERE ITEM = " & Val(.TextMatrix(nRow, 0)), con)
            cString = "Insert into FILE1_60 (doc_no,item,Quant,COST)" & _
                        "Values(" & _
                        addstring(cDoc_No) & "," & _
                        addvalue(.TextMatrix(nRow, 0)) & "," & _
                        Val(grid1.TextMatrix(nRow, 9)) & "," & _
                        Val(nCost) & _
                               ")"
            con.Execute cString
        End If
    Next
End With
con.CommitTrans
myreplace = True
Exit Function
myerror:
    MsgBox Err.Description
    con.RollbackTrans
    Err.Clear
End Function
Private Function myreplaceH()
On Error GoTo myerror
Dim aInsert(3, 1)

aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(Transfrm.xDoc_No.text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(Transfrm.xDate.text)

aInsert(2, 0) = "STORE1"
aInsert(2, 1) = addstring(Transfrm.xStore1.BoundText)

aInsert(3, 0) = "Store2"
aInsert(3, 1) = addstring(Transfrm.XSTORE2.BoundText)

con.BeginTrans
If Transfrm.xDoc_No.Tag = "1" Then
    Transfrm.xDoc_No.text = RetZero(Newflag("FILE1_60H", "doc_no", con))
    aInsert(0, 1) = addstring(Transfrm.xDoc_No.text)
    con.Execute CreateInsert(aInsert, "FILE1_60H")
End If
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function

