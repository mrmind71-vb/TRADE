VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form StockFrm_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ã—œ „Œ«“‰"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14310
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
   ScaleHeight     =   11040
   ScaleWidth      =   15270
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame11 
      Height          =   645
      Left            =   7170
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   0
      Width           =   1530
      Begin VB.CommandButton Command2 
         Caption         =   "ÿ»«⁄… «·Ã—œ"
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   180
         Width           =   1350
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ÿ»«⁄… ··Ã—œ"
         Height          =   375
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   180
         Visible         =   0   'False
         Width           =   1905
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   5460
      Begin VB.CommandButton cmdInform 
         Caption         =   "≈” ⁄·«„"
         Height          =   420
         Left            =   4095
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewinv 
         Caption         =   "„” ‰œ ÃœÌœ"
         Height          =   420
         Left            =   2745
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         Height          =   420
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Caption         =   "Õ–› «·„” ‰œ"
         Height          =   420
         Left            =   1395
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   7020
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   585
      Width           =   8160
      Begin VB.TextBox xDesca 
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
         Height          =   315
         Left            =   3870
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   585
         Width           =   3165
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   135
         TabIndex        =   3
         Top             =   585
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.TextBox xDoc_No 
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
         Height          =   315
         Left            =   5715
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   225
         Width           =   1335
      End
      Begin VB.TextBox xDate 
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
         Height          =   315
         Left            =   105
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   2100
      End
      Begin VB.Label Label4 
         Caption         =   "»Ì«‰ :"
         Height          =   240
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   630
         Width           =   705
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         Height          =   195
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   315
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         Height          =   240
         Left            =   7110
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Œ“‰ :"
         Height          =   195
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   675
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1065
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   540
      Width           =   1425
      Begin VB.CommandButton CmdUndo 
         Caption         =   " —«Ã⁄"
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   585
         Width           =   1230
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Õ›Ÿ "
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   135
         Width           =   1230
      End
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   5355
      Top             =   675
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
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
   Begin VB.Frame Frame5 
      Height          =   555
      Left            =   11295
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   9765
      Width           =   3885
      Begin VB.CommandButton cmdPost 
         BackColor       =   &H00E0E0E0&
         Caption         =   " —ÕÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2925
         RightToLeft     =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   135
         Width           =   915
      End
      Begin VB.CommandButton cmdunPost 
         BackColor       =   &H00E0E0E0&
         Caption         =   "≈·€«¡  —ÕÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   135
         Width           =   1185
      End
      Begin VB.CommandButton cmdFix 
         BackColor       =   &H00E0E0E0&
         Caption         =   "«⁄«œ… ÷»ÿ «·Ã—œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   135
         Width           =   1635
      End
   End
   Begin VB.Frame Frame7 
      Height          =   555
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   9675
      Width           =   1950
      Begin VB.CommandButton cmdfirst 
         Caption         =   ">|"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   135
         Width           =   465
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   975
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "”«»ﬁ"
         Top             =   135
         Width           =   465
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   510
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   " «·Ì"
         Top             =   135
         Width           =   465
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "|<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "«ŒÌ—"
         Top             =   135
         Width           =   465
      End
   End
   Begin VB.Frame Frame9 
      Height          =   555
      Left            =   3375
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   9675
      Width           =   2310
      Begin VB.Label xTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   180
         Width           =   1350
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«·≈Ã„«·Ì:"
         Height          =   240
         Left            =   1485
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   225
         Width           =   645
      End
   End
   Begin VB.Frame Frame10 
      Height          =   555
      Left            =   5715
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   9675
      Width           =   2490
      Begin VB.TextBox xfilter 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   135
         Width           =   2400
      End
   End
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   5355
      Top             =   270
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   330
      Left            =   5355
      Top             =   945
      Visible         =   0   'False
      Width           =   1665
      _ExtentX        =   2937
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
   Begin MSComctlLib.ProgressBar prog1 
      Height          =   375
      Left            =   45
      TabIndex        =   34
      Top             =   10305
      Visible         =   0   'False
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   7890
      Left            =   90
      TabIndex        =   35
      Top             =   1665
      Width           =   15090
      _cx             =   26617
      _cy             =   13917
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
End
Attribute VB_Name = "StockFrm_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CardTable As ADODB.Recordset, Search31 As New Search3, searchImp As New Search3
Dim ItemTable As New ADODB.Recordset
Dim clist1 As String, cList2 As String
Dim itemBalTable As ADODB.Recordset
Dim formMode, dDateLast As String
Dim con As New ADODB.Connection
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
Dim Generalarray(5)
Dim listarray(1, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select File1_10.item,File1_10.Desca,file1_50.desca,file1_10.cost From file1_10 left join file1_50 on file1_10.[GROUP] = file1_50.code"
Generalarray(2) = "Order by file1_10.Desca"
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(FILE1_10.ITEM LIKE 'cFilter%' or  %%FILE1_10.DESCA%%) "

listarray(1, 0) = "«·„Ã„Ê⁄…"
listarray(1, 1) = "(%%FILE1_50.DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·’‰›"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "≈”„ «·’‰›"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ã„Ê⁄…"
GrdArray(2, 1) = 2000

GrdArray(3, 0) = "«· ﬂ·›…"
GrdArray(3, 1) = 0

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„ «·«’‰«›"
Search3.Show 1
End Sub
Private Function MyReplace() As Boolean
Dim aInsert(3, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.Text)

aInsert(2, 0) = "store"
aInsert(2, 1) = addstring(xStore.BoundText)

aInsert(3, 0) = "Desca"
aInsert(3, 1) = addstring(xDesca.Text)

con.BeginTrans
If xDoc_No.Enabled Then
    xDoc_No.Text = RetZero(Val(Newflag("FILE0_10h", "doc_no", con)))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, "FILE0_10h")
Else
    con.Execute CreateUpdate(aInsert, "FILE0_10h", " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd
con.CommitTrans
MyReplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    If foundOther Then Exit Sub
    cItem = grid1.TextMatrix(grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 0) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    grid1.TextMatrix(grid1.Row, 2) = "1"
    If Trim(cItem) <> Trim(grid1.TextMatrix(grid1.Row, 0)) Then calcRow grid1.Row, 0
    If grid1.Row = grid1.Rows - 1 Then
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 0
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.Select grid1.Rows - 1, 0
    End If
ElseIf ActiveControl.Name = cmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(Search31.grid1.TextMatrix(Search31.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    myload
    Search31.Hide
ElseIf ActiveControl.Name = xdocImp.Name Then
    xdocImp.Text = searchImp.grid1.TextMatrix(searchImp.grid1.Row, 0)
    searchImp.Hide
End If
End Sub
Private Sub cmd_item_Click()
Load frmItem
frmItem.Show 1
End Sub
Private Sub cmdaddImp_Click()
If xStore.BoundText = "" Then
    MsgBox "≈Œ Ì«— «·„Œ“‰"
    Exit Sub
End If
Dim impTable As New ADODB.Recordset
cString = "select FILE7_60.*,file1_10.desca,file1_10.[unit] as unit2,file1_10.cost,file1_10.price,file1_10.price2,file1_10.package,file1_10.[group],file1_10.[section],file1_10.discount from file7_60 inner join file1_10 on file7_60.item = file1_10.item where doc_no = " & MyParn(xdocImp.Text) & " AND FILE7_60.STORE = " & MyParn(xStore.BoundText)
If XGROUP.BoundText <> "" Then cString = cString & " AND FILE1_10.[GROUP] = " & XGROUP.BoundText
impTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
With grid1
.Rows = 1
Do Until impTable.EOF
    .AddItem ""
    .TextMatrix(.Rows - 1, 0) = impTable!Item
    .TextMatrix(.Rows - 1, 1) = impTable![FILE1_10.DESCA] & ""
    .TextMatrix(.Rows - 1, 2) = impTable!unit2 & ""
    .TextMatrix(.Rows - 1, 5) = Val(.TextMatrix(.Rows - 1, 4)) - Val(.TextMatrix(.Rows - 1, 3))
    .TextMatrix(.Rows - 1, 7) = Val(.TextMatrix(.Rows - 1, 5)) * Val(.TextMatrix(.Rows - 1, 6))
    impTable.MoveNext
Loop
.AddItem ""
End With
CalcTotals
impTable.Close
Set impTable = Nothing
End Sub

Private Sub CmdAdditem_Click()
itemsgrdFrm.bEdit = True
itemsgrdFrm.Show 1
rdItem.Requery
End Sub

Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE0_10 where Doc_No = " & MyParn(xDoc_No.Text)
    con.Execute "Delete  From FILE0_10H where Doc_No = " & MyParn(xDoc_No.Text)
    con.CommitTrans
    CardTable.Requery
    If (CardTable.EOF And CardTable.BOF) Then
        myDefine
    Else
        CardTable.Find "doc_no < " & MyParn(xDoc_No.Text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
con.RollbackTrans
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdFix_Click()
If Not MYVALIDfix Then Exit Sub
With grid1
'Time = Now
prog1.Value = 0
prog1.Visible = True
For I = 1 To grid1.Rows - 2
     prog1.Value = Round(I / (grid1.Rows - 2), 2) * 100
    .TextMatrix(I, 3) = RetItemBalance(.TextMatrix(I, 0), xStore.BoundText, xDate.Text, con) & ""
    .TextMatrix(I, 5) = Val(.TextMatrix(I, 4)) - Val(.TextMatrix(I, 3))
    .TextMatrix(I, 7) = Val(.TextMatrix(I, 5)) * Val(.TextMatrix(I, 6))
Next
prog1.Visible = False
MyReplace
'MsgBox DateDiff("s", nTime, Now)
MsgBox " „ ÷»ÿ «·„” ‰œ »‰Ã«Õ"
End With
End Sub
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE,CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA " & _
                  " FROM FILE0_10H INNER JOIN FILE0_40 ON FILE0_10H.Store = FILE0_40.CODE "

Generalarray(2) = "Order by [Date]"
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ-«·„Œ“‰"
listarray(0, 1) = "@@Doc_No@@6 or  FILE0_40.DESCA Like '%cFilter%' OR " & _
                  "##[DATE]##"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·„Œ“‰"
GrdArray(3, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search31
Search31.Caption = "«” ⁄·«„"
Search31.Show 1
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CmdLast_Click()
CardTable.MoveLast
myload
End Sub
Private Sub CmdNext_Click()
CardTable.MoveNext
If CardTable.EOF Then
    CardTable.MovePrevious
Else
    myload
End If
End Sub
Private Sub CmdPrevious_Click()
CardTable.MovePrevious
If CardTable.BOF Then
    CardTable.MoveNext
Else
    myload
End If
End Sub
Private Sub cmdPost_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload

On Error GoTo myerror
con.BeginTrans
cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 1 WHERE FILE0_10H.DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute cString
con.CommitTrans

CardTable.Requery
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
Handlecontrols LoadMode
MsgBox " „  —ÕÌ· «·„” ‰œ »‰Ã«Õ"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub CmdNewInv_Click()
'CardTable.MoveLast
'xDoc_No.Text = RetZero(Val(CardTable!doc_no & ""))
myDefine
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not MyReplace Then Exit Sub
CardTable.Requery
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
End Sub
Private Sub CmdUndo_Click()
If CardTable.EOF And CardTable.BOF Then
    grid1.Rows = 1
    grid1.AddItem ""
    Exit Sub
End If
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then
    CardTable.MoveLast
    myload
Else
    myload
End If
End Sub
Private Sub cmdunPost_Click()
'On Error GoTo Myerror
con.BeginTrans
cString = " UPDATE FILE0_10H SET FILE0_10H.closed = 0 WHERE FILE0_10H.DOC_NO = " & MyParn(xDoc_No.Text)
con.Execute cString
con.CommitTrans
CardTable.Requery

CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
Handlecontrols LoadMode
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub

Private Sub Command3OLD_Click()
Load PrintGrd
grid1.TextMatrix(grid1.Rows - 1, grid1.Cols - 1) = xTotal.Caption
grid1.TextMatrix(grid1.Rows - 1, 0) = "«·«Ã„«·Ì"
PrintGrd.doprint grid1, 1, True
PrintGrd.Show 1
grid1.TextMatrix(grid1.Rows - 1, grid1.Cols - 1) = ""
grid1.TextMatrix(grid1.Rows - 1, 0) = ""
End Sub

Private Sub Command1_Click()
doprint "stock1.rpt"
End Sub

Private Sub Command2_Click()
If MsgBox("ÿ»«⁄… »Ì«‰ »œÊ‰ √—œ’… ··Ã—œ", vbYesNo) = vbYes Then
    doprint "stock3.rpt"
Else
    doprint "stock2.rpt"
End If
End Sub

Private Sub Command3_Click()
Dim listTable As ADODB.Recordset
Set listTable = New ADODB.Recordset
xBar.Value = 0
xBar.Min = 0
'cmdSave_Click
cString = "select FILE1_10.ITEM , FILE1_10.DESCA , SUM([In] - [OUT]) AS BALITEM FROM   FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM " & _
            " WHERE  FILE1_11.DATE < " & DateSq(xDate.Text) & " and file1_11.store = " & MyParn(xStore.BoundText) & " AND FILE1_10.ITEM NOT IN ( SELECT ITEM FROM FILE0_10 INNER JOIN FILE0_10H ON FILE0_10.DOC_NO = FILE0_10H.DOC_NO WHERE FILE0_10H.DATE = " & DateSq(xDate.Text) & " and FILE0_10H.store = " & MyParn(xStore.BoundText) & " )   GROUP BY FILE1_10.ITEM , FILE1_10.DESCA HAVING (((Sum([In]-[OUT]))<>0))"
 
listTable.Open cString, con, adOpenForwardOnly, adLockReadOnly, adCmdText
'XW.Caption = "Ì „ «·√‰ ≈÷«›… «·«’‰«›"
With grid1
Do Until listTable.EOF
   ' xBar.Value = grid1.Rows - 1
'    XW.Caption = listTable!Item
    .TextMatrix(.Rows - 1, 0) = listTable!Item
    .TextMatrix(.Rows - 1, 1) = listTable!desca & ""
    .TextMatrix(.Rows - 1, 3) = listTable!BalItem & ""
    .AddItem ""
    listTable.MoveNext
Loop
End With
'XW.Caption = " „ «·√÷«›… "

End Sub

Private Sub fix_Click()
    con.Execute " UPDATE FILE0_10 LEFT JOIN FILE1_10 ON FILE0_10.item = FILE1_10.ITEM SET FILE0_10.cost = [file1_10].[price] WHERE (((FILE0_10.cost) Is Null Or (FILE0_10.cost)=0)) "
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo) And (ActiveControl.Name <> xfilter.Name) Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
openCon con
clist1 = StrList("select * from file1_50 order by desca")
cList2 = StrList("select * from file1_10SC order by desca")
Set CardTable = New ADODB.Recordset

CardTable.Open "SELECT * FROM File0_10H  ORDER BY DOC_NO ", con, adOpenStatic, adLockReadOnly, adCmdText

DATA1.ConnectionString = strCon
DATA1.RecordSource = "FILE0_40"
Set xStore.RowSource = DATA1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"

DATA4.ConnectionString = strCon
DATA4.RecordSource = "FILE1_50"
Set XGROUP.RowSource = DATA4
XGROUP.ListField = "Desca"
XGROUP.BoundColumn = "Code"

Set grid1.DataSource = DATA3
DATA3.ConnectionString = strCon

'data2.ConnectionString = strCon
'data2.RecordSource = "SELECT Distinct Doc_No, DESCA FROM FILE0_30"

'Set xList.RowSource = data2
'xList.ListField = "Desca"
'xList.BoundColumn = "Doc_No"

'With grid1
'    .Cols = 7
'    .Rows = 1
'    .Editable = flexEDKbd
'    .FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "«·ÊÕœ…|" & "—’Ìœ ﬂÊ„»ÌÊ —|" & "—’Ìœ Ã—œ |" & "«·›—ﬁ |" & " ﬂ·›… |" & "«· ﬁÌ„"
'    .ColWidth(0) = 1200
'    .ColWidth(1) = 3200
'    .ColWidth(2) = 1300
'    .ColWidth(3) = 1300
'    .ColWidth(4) = 1300
'    .ColWidth(5) = 1300
'    .ColWidth(6) = 1300
'    .ColWidth(7) = 1300
'    .ColHidden(2) = True
'
'    .ColDataType(3) = flexDTDouble
'    .ColDataType(4) = flexDTDouble
'    .ColDataType(5) = flexDTDouble
'    .ColDataType(6) = flexDTDouble
'    .ColDataType(0) = flexDTString
'
'    .ColAlignment(0) = flexAlignRightCenter
'    .ColAlignment(1) = flexAlignRightCenter
'    .ColAlignment(2) = flexAlignRightCenter
'    .ColAlignment(3) = flexAlignRightCenter
'    .ColAlignment(4) = flexAlignRightCenter
'    .ColAlignment(5) = flexAlignRightCenter
'    .ColAlignment(6) = flexAlignRightCenter
'    .ColAlignment(7) = flexAlignRightCenter
'End With
'With Grid2
'    .Cols = 7
'    .Rows = 1
'    .Editable = flexEDKbd
'    .FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "—’Ìœ Ã—œ |" & "«·›—ﬁ |" & " ﬂ·›… |" & "«· ﬁÌ„"
'End With
grid1.ColHidden(2) = True
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    myDefine
End If
End Sub
Sub dispProc()
formMode = dispMode
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload Search3
Unload Search31
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set CardTable = Nothing
closeCon con
Err.Clear
End Sub
Private Sub Grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col = 0 Then
    calcRow Row, Col
End If
CalcTotals
End Sub

Private Sub grid1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
ItemsLookup
End Sub

Private Sub Grid1_EnterCell()
If (grid1.Col = 4 Or grid1.Col = 0) Or grid1.Col >= 8 Then
  grid1.Editable = flexEDKbd
    SetKbLayout Lang_EN
Else
   grid1.Editable = flexEDNone
End If
End Sub

Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.SetFocus
    grid1.Select 1, 2
End If
End Sub

Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 45 Then grid1.AddItem "", grid1.Row
If KeyCode = 112 And grid1.Col = 0 And grid1.Row <> 0 Then ItemsLookupAll Me, Search3
End Sub

Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If KeyAscii = 13 And grid1.Col = 0 Then
    If grid1.Row = grid1.Rows - 1 Then
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 1
    Else
        grid1.Select grid1.Row + 1, 1
    End If
End If
If KeyAscii = 13 Then
    
    Select Case Col
        Case 0
            grid1.Col = 4
            grid1.Row = Row
        Case 4
            grid1.Col = 0
            grid1.Row = Row + 1
    End Select
End If

End Sub

Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then ItemsLookup

If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And CmdSave.Enabled Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from file0_10 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
            con.CommitTrans
        End If
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case grid1.Col
    Case 0
        If KeyCode = 27 Then Exit Sub
        If KeyCode = 112 Then ItemsLookup
End Select
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then grid1.AddItem ""
End Sub
Private Function MYVALID() As Boolean
If foundOther Then Exit Function
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If


With grid1
For I = 1 To .Rows - 2
    If .TextMatrix(I, 0) = "" Then
        .Select I, 0, I, grid1.Cols - 1
        MsgBox "ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
        Exit Function
    Else
        If GetDesca("select item from file1_10 where file1_10.item = " & MyParn(.TextMatrix(I, 0))) = "" Then
            .Select I, 0, I, 2
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        End If
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore.BoundText = CardTable!store & ""
xDesca.Text = CardTable!desca & ""
cString = "SELECT FILE0_10.ITEM, FILE1_10.DESCA,FILE1_10.UNIT as unit2,FILE0_10.ComputerBal ,File0_10.RealBal,File0_10.differ,File0_10.cost,file0_10.Realbal * file0_10.Cost as Total,file1_10.price as [”⁄— «·Ã„·…],file1_10.price2 as [”⁄— «·ﬁÿ«⁄Ì],package as [«·⁄»Ê…],unit as [«·ÊÕœ…],[group] as [«·„Ã„Ê⁄…],[section] as [«·ﬁ”„],FILE1_10.DISCOUNT,ID  " & _
      " FROM FILE0_10 INNER JOIN FILE1_10 ON FILE0_10.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " order by FILE0_10.ROW"
DATA3.RecordSource = cString
DATA3.Refresh
grid1.AddItem ""
Handlecontrols LoadMode
CalcTotals
FixGrd
End Sub
Private Sub myDefine()
xDoc_No.Text = RetZero(Val(Newflag("FILE0_10h", "doc_no", con)))
xDate.Text = Format(Date, "dd-mm-yyyy")
xStore.BoundText = ""
xDesca.Text = ""
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
FixGrd
End Sub
Private Sub Handlecontrols(nMode)
Dim bClosed As Boolean
'cmdAddList.Enabled = (grid1.Rows <= 2)
cmdNewinv.Enabled = nMode = LoadMode And bEdit
cmdfirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
If Not (CardTable.EOF And CardTable.BOF) Then bClosed = CardTable!CLOSED

cmdPost.Enabled = (Not bClosed) And bEdit And nMode = LoadMode
cmdunPost.Enabled = bClosed And bEdit And nMode = LoadMode
cmdFix.Enabled = (Not bClosed) And bEdit And nMode = LoadMode
CmdSave.Enabled = (nMode = DefineMode Or Not bClosed) And bEdit
CmdDelInv.Enabled = (Not bClosed) And bEdit And nMode = LoadMod
End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 And Trim(grid1.EditText) <> "" Then
    nFound = FoundOtheritem(Row, Col, Trim(grid1.EditText))
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & grid1.TextMatrix(nFound, 1)
        Cancel = True
    End If
    If GetDesca("select item from file1_10 where file1_10.item = " & MyParn(grid1.EditText)) = "" Then
        MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        Cancel = True
    End If
End If

End Sub

Private Sub xDate_Change()
cmdAddList.Enabled = xStore.BoundText <> "" And IsDate(xDate.Text)
End Sub

Private Sub xDoc_No_LostFocus()
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If (CardTable.EOF And CardTable.BOF) Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function MYVALIDfix() As Boolean
If foundOther Then Exit Function
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If


With grid1
For I = 1 To .Rows - 2
    If .TextMatrix(I, 0) = "" Then
        .Select I, 0, I, grid1.Cols - 1
        MsgBox "ﬂÊœ «·’‰› €Ì— „ÊÃÊœ"
        Exit Function
    Else
        If GetDesca("SELECT ITEM FROM FILE1_10 WHERE ITEM = " & MyParn(.TextMatrix(I, 0))) = "" Then
            .Select I, 0, I, 2
            MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
            Exit Function
        End If
    End If
Next
End With
MYVALIDfix = True
End Function
Private Sub calcRow(nRow, nCol)
With grid1
.TextMatrix(nRow, 1) = ""
.TextMatrix(nRow, 2) = ""
.TextMatrix(nRow, 3) = ""
.TextMatrix(nRow, 5) = ""
If nCol = 0 Then .TextMatrix(nRow, 6) = ""
.TextMatrix(nRow, 7) = ""

If Trim(grid1.TextMatrix(nRow, 0)) = "" Then Exit Sub

Dim aret As Variant
aret = aGetDesca("Select Desca,Price,Price2,Package,Unit,[Group],[section],Discount from file1_10 where file1_10.item = " & MyParn(grid1.TextMatrix(nRow, 0)))
If UBound(aret) > 0 Then
    .TextMatrix(nRow, 1) = aret(1) & ""
    If IsDate(xDate.Text) And Trim(xStore.BoundText) <> "" Then
        grid1.TextMatrix(nRow, 3) = RetItemBalance(grid1.TextMatrix(nRow, 0), xStore.BoundText, xDate.Text, con) & ""
    End If
    .TextMatrix(nRow, 5) = Val(.TextMatrix(nRow, 4)) - Val(.TextMatrix(nRow, 3)) & ""
    If nCol = 0 Then .TextMatrix(nRow, 6) = itemCost(grid1.TextMatrix(nRow, 0), xDate.Text) & ""
    .TextMatrix(nRow, 7) = Val(.TextMatrix(nRow, 4)) * Val(.TextMatrix(nRow, 6)) & ""
    .TextMatrix(nRow, 8) = aret(2) & ""
    .TextMatrix(nRow, 9) = aret(3) & ""
    .TextMatrix(nRow, 10) = aret(4) & ""
    .TextMatrix(nRow, 11) = aret(5) & ""
    .TextMatrix(nRow, 12) = aret(6) & ""
    .TextMatrix(nRow, 13) = aret(7) & ""
    .TextMatrix(nRow, 14) = aret(8) & ""
End If
End With
End Sub
Private Function RetItemBalance2(cItem, cStore, dDate) As Double
If cItem = "" Then Exit Function
cString = "Select sum(val([IN] & '') - VAL([OUT] & '')) as Balance From file1_11 where item = " & MyParn(cItem) & _
          " and Store = " & MyParn(cStore) & " and Date < " & DateSq(dDate)
RetItemBalance2 = Val(GetDesca(cString) & "")
End Function

Private Sub xdocimp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then impcostLookup
End Sub

Private Sub xStore_Click(Area As Integer)
cmdAddList.Enabled = xStore.BoundText <> "" And IsDate(xDate.Text)
End Sub
Private Function RetTotalCost()
For I = 1 To grid1.Rows - 1
    RetTotalCost = RetTotalCost + Val(grid1.TextMatrix(I, 0))
Next
End Function
Private Sub CalcTotals()
Dim nTotal As Double
For I = 1 To grid1.Rows - 1
    nTotal = nTotal + (Val(grid1.TextMatrix(I, 4)) * Val(grid1.TextMatrix(I, 6)))
Next
xTotal.Caption = Format(nTotal, "#0.00")
End Sub
Private Sub FixGrd()
With grid1
.FormatString = "ﬂÊœ|" & "«·’‰‹‹‹‹‹‹›|" & "«·ÊÕœ…|" & "—’Ìœ ﬂÊ„»ÌÊ —|" & "—’Ìœ Ã—œ |" & "«·›—ﬁ |" & " ﬂ·›… |" & "«· ﬁÌ„|" & "”⁄— Ã„·…|" & "”⁄— ﬁÿ«⁄Ì|" & "«·⁄»Ê…|" & "«·ÊÕœ…|" & "«·„Ã„Ê⁄…|" & "«·ﬁ”„|" & "«·Œ’„|"
.ColWidth(0) = 2500
.ColWidth(1) = 5000
.ColWidth(2) = 1500
.ColWidth(3) = 1500
.ColWidth(4) = 1500
.ColWidth(5) = 1500
For I = 6 To .Cols - 1
'    .ColWidth(i) = 500
    .ColHidden(I) = True
Next I

'.ColHidden(8) = True
.ColComboList(12) = clist1
.ColComboList(13) = cList2
.ColDataType(3) = flexDTDouble
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(0) = flexDTString
For I = 0 To grid1.Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
.ColHidden(2) = True

'.ColComboList(0) = "..."
End With
End Sub
Private Sub doprint(cReport)
Dim loctable As New ADODB.Recordset
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset

loctable.Open "select file1_10.item,file1_10.[GROUP],file1_50.desca " & _
              " from file1_10 left join file1_50 on file1_10.[GROUP] = file1_50.code", con, adOpenStatic, adLockReadOnly, adCmdText
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

For I = 1 To grid1.Rows - 2
    temptable.AddNew
    aret = retFilter(loctable, "item = " & MyParn(grid1.TextMatrix(I, 0)))
    temptable!str3 = TurnValue(xStore.Text)
    temptable!date3 = DateFix(xDate.Text)
    temptable!str4 = TurnValue(grid1.TextMatrix(I, 0))
    temptable!str5 = TurnValue(grid1.TextMatrix(I, 1))
    temptable!str10 = aret(0)
    temptable!Str11 = aret(1)
    temptable!str12 = aret(2)
'    CGR = GetDesca("SELECT GROUP FROM FILE1_10 WHERE ITEM = " & MyParn(Grid1.TextMatrix(I, 0)))
'    temptable!Str12 = GetDesca("SELECT DESCA FROM FILE1_50 WHERE CODE = " & MyParn(CGR))
    temptable!val1 = TurnValue(Val(grid1.TextMatrix(I, 3)))
    temptable!val2 = TurnValue(Val(grid1.TextMatrix(I, 4)))
    temptable!val3 = TurnValue(Val(grid1.TextMatrix(I, 5)))
    temptable!val4 = TurnValue(Val(grid1.TextMatrix(I, 6)))
    temptable!val5 = TurnValue(Val(grid1.TextMatrix(I, 7)))
    temptable!str21 = "ÿ»«⁄… Ã—œ " & IIf(Trim(xDesca.Text) = "", "", " : " & Trim(xDesca.Text))
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\" & cReport
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub

Private Function foundOther() As Boolean
For I = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(I, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ " & I & " Ê " & nRow
        grid1.Select nRow, 0, nRow, grid1.Cols - 1
        If Not grid1.RowIsVisible(nRow) Then grid1.ShowCell nRow, 0
        foundOther = True
        Exit Function
    End If
Next
End Function
Private Function FoundOtherRow(nRow, nCol) As Integer
FoundOtherRow = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = Trim(grid1.TextMatrix(nRow, nCol)) Then
            FoundOtherRow = I
            Exit Function
        End If
    End If
Next
End Function
Private Function FoundOtheritem(nRow, nCol, nValue) As Integer
FoundOtheritem = -1
For I = 1 To grid1.Rows - 2
    If I <> nRow Then
        If Trim(grid1.TextMatrix(I, nCol)) = nValue Then
            FoundOtheritem = I
            Exit Function
        End If
    End If
Next
End Function
Private Sub impcostLookup()
Dim Generalarray(5)
Dim listarray(1, 5)
Dim GrdArray(7, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT FILE7_60H.DOC_NO,FILE7_60H.DATE,CONVERT(VARCHAR(10),FILE7_60H.[DATE],111),FILE4_10.Desca,FILE7_60H.Credit,FILE7_60H.Vessel, CASE WHEN FILE7_20H.DOCIMP IS NULL THEN '€Ì— „—Õ·' ELSE '„—Õ·' END" & _
                  " FROM ((FILE7_60H left JOIN FILE4_10 ON FILE7_60H.CODE = FILE4_10.CODE) LEFT JOIN FILE7_20H ON FILE7_60H.DOC_NO = FILE7_20H.DOCIMP) "

Generalarray(2) = "Order by FILE7_60H.Date"
Generalarray(3) = 6000
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-≈”„ «·„Ê—œ-«· «—ÌŒ-«”„ «·”›Ì‰…"
listarray(0, 1) = "@@FILE7_60H.Doc_No @@ or  %%FILE4_10.DESCA%% OR %%FILE7_60H.VESSEL%% OR " & _
                  "##FILE7_60H.Date##)"

listarray(1, 0) = "—ﬁ„ «·«⁄ „«œ"
listarray(1, 1) = "(FILE7_60H.Credit Like '%cFilter%')"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·≈”„"
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "—ﬁ„ «·«⁄ „«œ"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "«”„ «·„—ﬂ»"
GrdArray(5, 1) = 1500

GrdArray(6, 0) = "«· —ÕÌ·"
GrdArray(6, 1) = 1100


searchArray = Array(Generalarray, listarray, GrdArray)
searchImp.Caption = "«” ⁄·«„"
searchImp.Show 1
End Sub
Private Sub myreplaceGrd()
Dim nCost As Double
Dim aInsert(6, 1)
With grid1
    For I = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(I, 0))
        
        aInsert(2, 0) = "ComputerBal"
        aInsert(2, 1) = Val(.TextMatrix(I, 3))

        aInsert(3, 0) = "RealBal"
        aInsert(3, 1) = Val(.TextMatrix(I, 4))

        aInsert(4, 0) = "Differ"
        aInsert(4, 1) = Val(.TextMatrix(I, 5))

        aInsert(5, 0) = "row"
        aInsert(5, 1) = I
        
        If grid1.TextMatrix(I, .Cols - 1) = "" Then
            nCost = Val(GetDesca("select price from file1_10 where file1_10.item = " & MyParn(.TextMatrix(I, 0))) & "")
        End If
        aInsert(6, 0) = "cost"
        aInsert(6, 1) = nCost
        
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "FILE0_10")
        Else
            con.Execute CreateUpdate(aInsert, "FILE0_10", " where ID = " & grid1.TextMatrix(I, .Cols - 1), Array(0, 6))
        End If
    Next
End With

End Sub
