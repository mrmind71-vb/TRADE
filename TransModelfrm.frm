VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form TransModelfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ”ÃÌ·  ÕÊÌ· ·„ÊœÌ· "
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13650
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   13650
   WindowState     =   2  'Maximized
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5580
      Left            =   90
      TabIndex        =   1
      Top             =   2205
      Width           =   13470
      _cx             =   23760
      _cy             =   9842
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   65535
      ForeColorSel    =   64
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
      Rows            =   3
      Cols            =   2
      FixedRows       =   3
      FixedCols       =   2
      RowHeightMin    =   300
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
      TabBehavior     =   1
      OwnerDraw       =   0
      Editable        =   2
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
   Begin VB.Frame Frame3 
      Height          =   1725
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   2310
      Begin VB.CommandButton cmdSave 
         Caption         =   " ”ÃÌ·  ÕÊÌ· «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   2130
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Œ—ÊÃ"
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
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1170
         Width           =   2130
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   2430
      TabIndex        =   5
      Top             =   0
      Width           =   11130
      Begin VB.TextBox xModel 
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
         Left            =   7065
         MaxLength       =   20
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   2985
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
         Left            =   10125
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   270
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   300
      Top             =   -540
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
      Caption         =   "Adodc2"
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
      Height          =   1365
      Left            =   2430
      TabIndex        =   7
      Top             =   810
      Width           =   11130
      Begin VB.Label xBalNo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   8460
         TabIndex        =   16
         Top             =   945
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   990
         Width           =   510
      End
      Begin VB.Label xMosm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   3240
         TabIndex        =   14
         Top             =   540
         Width           =   600
      End
      Begin VB.Label xModelFact 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   6870
         TabIndex        =   13
         Top             =   540
         Width           =   2625
      End
      Begin VB.Label xFact 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   3885
         TabIndex        =   12
         Top             =   540
         Width           =   2940
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ „ÊœÌ·"
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
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   607
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈”„ «·„ÊœÌ·"
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
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   1005
      End
      Begin VB.Label xSupp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   9540
         TabIndex        =   9
         Top             =   540
         Width           =   510
      End
      Begin VB.Label xdesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   180
         Width           =   6810
      End
   End
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   3510
      _ExtentX        =   6191
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
End
Attribute VB_Name = "TransModelfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Search1 As New Search, Search2 As New Search
Dim ItemTable As New ADODB.Recordset
Dim ModelTable As New ADODB.Recordset
Dim con As New ADODB.Connection
Private Sub CmdExit_Click()
    Unload Me
    Set TransModelfrm = Nothing
End Sub
Private Function myreplace() As Boolean
Dim cString As String, nRow As Integer, cFile As String, nItem As Integer
On Error GoTo myerror
If Transfrm.xDoc_No.Tag = "1" Then
    myreplaceH
End If
cDoc_No = Transfrm.xDoc_No.Text
cString = " DELETE FROM FILE1_60 FROM   FILE1_10 INNER JOIN                       FILE1_60 ON FILE1_10.ITEM = FILE1_60.ITEM WHERE DOC_NO = " & MyParn(cDoc_No) & " and FILE1_10.MODEL = " & MyParn(xModel.Text)
con.BeginTrans
con.Execute cString

With grid1
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(grid1.TextMatrix(nRow, nCol)) <> 0 Then
                aret = aGetDesca("SELECT item , MODEL ,COST from file1_10 where MODEL = " & MyParn(xModel.Text) & " AND SCAL = " & MyParn(.TextMatrix(0, nCol)) & " AND C_COLOR = " & Val(.TextMatrix(nRow, 0)))
                If UBound(aret) > 0 Then
                    cString = "Insert into FILE1_60 (doc_no,item,Quant,COST)" & _
                               "Values(" & _
                               addstring(cDoc_No) & "," & _
                               addvalue(aret(1)) & "," & _
                               Val(grid1.TextMatrix(nRow, nCol)) & "," & _
                               addvalue(aret(2)) & _
                               ")"
                    con.Execute cString
                End If
            Else
                con.Execute "DELETE FROM FILE1_60 WHERE ITEM = " & aret(1) & " AND DOC_NO = " & MyParn(cDoc_No)
            End If
        Next
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
Private Sub CMD_EXIT_Click()
    Unload Me
End Sub
Private Sub cmdSave_Click()
    If myreplace Then
        Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
        myDefine
    End If
    On Error Resume Next
    xModel.SetFocus
    Err.Clear
End Sub
Private Sub Form_Load()
    con.CursorLocation = adUseClient
    con.Open strCon
    ItemTable.Open "FILE1_10", con, adOpenStatic, adLockReadOnly, adCmdTable
    ModelTable.Open "FILE1_10H", con, adOpenStatic, adLockReadOnly, adCmdTable
 
    myDefine
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then SendKeys "{tAB}"
End Sub
Public Sub myloadgrd()
    Dim aret As Variant, cFieldas As String, cField As String
    Dim cModelNo As String
    myDefineGrd
    aret = retFields
    cField = aret(0)
    cFieldas = aret(1)
    FillItem cFieldas, cField
    FixCost cFieldas, cField
    Fixgrd
    On Error Resume Next
    grid1.SetFocus
End Sub
Private Sub Fixgrd()
grid1.ColWidth(0) = 500
grid1.ColWidth(1) = 1200
nColWidth = (grid1.Width - 200 - grid1.ColWidth(0) - grid1.ColWidth(1)) / grid1.Cols
If nColWidth < 500 Then nColWidth = 500
If nColWidth > 1000 Then nColWidth = 1000
For nCol = 2 To grid1.Cols - 1
    grid1.ColWidth(nCol) = nColWidth
    grid1.ColAlignment(nCol) = flexAlignCenterCenter
Next
grid1.Select 3, 2
grid1.ShowCell 3, 2
End Sub
Private Sub myDefineGrd()

grid1.Rows = 3
grid1.Cols = 2
grid1.FixedCols = 2
grid1.FixedRows = 3
grid1.ColHidden(0) = True
grid1.MergeCells = flexMergeRestrictRows
grid1.TextMatrix(0, 1) = "«·„ﬁ«”"
grid1.TextMatrix(1, 1) = "”⁄— „’‰⁄"
grid1.TextMatrix(2, 1) = "”⁄— „” Â·ﬂ"

grid1.RowHidden(1) = True
End Sub
Private Sub FixCost(cFieldas, cField)
' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,cost from file1_10 WHERE MODEL = " & MyParn(xModel.Text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(cost)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grid1.Cols - 1
        grid1.TextMatrix(1, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If
' „·∆ «·ÃœÊ·
cString = "Select " & cFieldas & _
          " From " & _
          " (Select scal,cost from file1_10 WHERE MODEL = " & MyParn(xModel.Text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(cost)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  "

loctable.Close
loctable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If Not loctable.EOF Then
    For nCol = 2 To grid1.Cols - 1
        grid1.TextMatrix(2, nCol) = loctable.Fields(nCol - 2).Value & ""
    Next
End If
loctable.Close
Set loctable = Nothing
End Sub
Private Function retFields()
Dim aret(1) As String
Dim FieldTable As New ADODB.Recordset
'  ⁄—Ì› «·«⁄„œ…
FieldTable.Open "Select SCAL from file1_10 where model = " & MyParn(xModel.Text) & " group by SCAL,C_SCAL order by c_scal", con, adOpenStatic, adLockReadOnly
Do Until FieldTable.EOF
    If Not IsNull(FieldTable!scal) Then
        cFieldas = cFieldas & turn(cField, ",") & "[" & FieldTable!scal & "]" & " as " & "[" & FieldTable!scal & "]"
        cField = cField & turn(cField, ",") & "[" & FieldTable!scal & "]"
    End If
    FieldTable.MoveNext
Loop

aret(0) = cField
aret(1) = cFieldas
retFields = aret
' ⁄œ„ ÊÃÊœ «⁄„œ…
FieldTable.Close
Set FieldTable = Nothing
End Function
Private Sub FillItem(cFieldas, cField)
Dim GRDTABLE As New ADODB.Recordset

cString = "Select c_color as [—ﬁ„ «··Ê‰] ,color as [«··Ê‰] " & turn(cFieldas, ",") & cFieldas & _
          " From " & _
          " (Select c_color,Color,scal,item from file1_10 WHERE MODEL = " & MyParn(xModel.Text) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(item)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  " & _
          " order by pvt.C_color"


' „·∆ «·ÃœÊ·
'cString = "Select c_color as [—ﬁ„ «··Ê‰] ,color as [«··Ê‰] " & turn(cFieldas, ",") & cFieldas & _
'          " From " & _
'          " (Select c_color,Color,scal,item from file1_10 WHERE MODEL = " & MyParn(xModel.Caption) & " ) AS TABLE1" & _
'          " PIVOT " & _
'          " (max(item)" & _
'          " FOR SCAL IN " & _
'          "(" & cField & ")" & _
'          ") as pvt  " & _
'          " order by pvt.c_color"

GRDTABLE.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grid1.Cols = GRDTABLE.Fields.Count

For nCol = 2 To GRDTABLE.Fields.Count - 1
    grid1.TextMatrix(0, nCol) = GRDTABLE.Fields(nCol).Name
Next
Do Until GRDTABLE.EOF
    grid1.AddItem ""
    For nCol = 0 To GRDTABLE.Fields.Count - 1
        If nCol <= 1 Then
            grid1.TextMatrix(grid1.Rows - 1, nCol) = GRDTABLE.Fields(nCol).Value & ""
        Else
            grid1.TextMatrix(grid1.Rows - 1, nCol) = GetDesca("SELECT QUANT FROM FILE1_60 WHERE DOC_NO = " & MyParn(Transfrm.xDoc_No.Text) & " AND ITEM = " & GRDTABLE.Fields(nCol).Value) & ""
'            nFoundRow = Transfrm.grid1.FindRow(GrdTable.Fields(nCol).Value & "", , 0)
'            If nFoundRow <> -1 Then
'                 grid1.TextMatrix(grid1.Rows - 1, nCol) = Transfrm.grid1.TextMatrix(nFoundRow, 8)
'             Else
'                 grid1.TextMatrix(grid1.Rows - 1, nCol) = ""
'             End If
        End If
    Next
    GRDTABLE.MoveNext
Loop
GRDTABLE.Close
Set GRDTABLE = Nothing
End Sub
Private Sub myDefine()
    xModelFact.Caption = ""
    xdesca.Caption = ""
    xModel.Text = ""
    xSupp.Caption = ""
    xFact.Caption = ""
    xMosm.Caption = ""
    xBalNo.Caption = ""
    myDefineGrd
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    closeCon con
Err.Clear
End Sub
Private Sub Grid1_EnterCell()
    With grid1
        .Cell(flexcpBackColor, 3, 2, .Rows - 1, .Cols - 1) = vbWhite
        .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = vbYellow
''''
'       xBalNo.Caption = BALITEM
''''
    End With
End Sub
Private Function myreplaceH()
On Error GoTo myerror
Dim aInsert(3, 1)

aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(Transfrm.xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(Transfrm.xDate.Text)

aInsert(2, 0) = "STORE1"
aInsert(2, 1) = addstring(Transfrm.xstore1.BoundText)

aInsert(3, 0) = "Store2"
aInsert(3, 1) = addstring(Transfrm.xstore2.BoundText)

con.BeginTrans
If Transfrm.xDoc_No.Tag = "1" Then
    Transfrm.xDoc_No.Text = RetZero(Newflag("FILE1_60H", "doc_no"))
    aInsert(0, 1) = addstring(Transfrm.xDoc_No.Text)
    con.Execute CreateInsert(aInsert, "FILE1_60H")
End If
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myProc()
On Error GoTo myerror
ActiveControl.Text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
Unload Search3
XModel_LostFocus
Exit Sub
myerror:
End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
'    If Val(.EditText) > Val(xBalNo.Caption) Then
'        MsgBox "«·—’Ìœ ·« Ì”„Õ"
'    End If
End With
End Sub

Private Sub xModel_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
    ModelLookupAll Me, Search3
End If
End Sub
Private Sub XModel_LostFocus()
    If xModel.Text = "" Then Exit Sub
    If Len(xModel.Text) <= 7 Then
        ItemTable.Find " ITEM = " & Val(xModel.Text), , adSearchForward, adBookmarkFirst
        If Not ItemTable.EOF Then
            xModel.Text = ItemTable!Model
            xSupp.Caption = ItemTable!SUPP & ""
            xdesca.Caption = ItemTable!desca & ""
            xMosm.Caption = ItemTable!MOSM & ""
            xSupp.Caption = ItemTable!SUPP & ""
            xModelFact.Caption = ItemTable!MODELFACT0 & ""
            xFact.Caption = GetDesca("SELECT DESCA FROM FACT WHERE CODE = " & MyParn(ItemTable!Fact))
        Else
            myDefine
        End If
    Else
        ModelTable.Find " MODEL = " & Val(xModel.Text), , adSearchForward, adBookmarkFirst
        If Not ItemTable.EOF Then
            xSupp.Caption = ItemTable!SUPP & ""
            xdesca.Caption = ItemTable!desca & ""
            xMosm.Caption = ItemTable!MOSM & ""
            xSupp.Caption = ItemTable!SUPP & ""
            xModelFact.Caption = ItemTable!MODELFACT0 & ""
            xFact.Caption = GetDesca("SELECT DESCA FROM  FACT WHERE CODE = " & MyParn(ItemTable!Fact))
        End If
    End If
    myloadgrd
End Sub
