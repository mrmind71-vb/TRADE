VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Transfrm2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " ÕÊÌ·«  „Œ«“‰"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15300
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   9045
   ScaleWidth      =   15300
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   9135
      Width           =   2715
      Begin VB.CommandButton cmdNext 
         Height          =   375
         Left            =   1365
         Picture         =   "transnew.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         ToolTipText     =   "«· «·Ì"
         Top             =   180
         Width           =   600
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   375
         Left            =   660
         Picture         =   "transnew.frx":25C0
         Style           =   1  'Graphical
         TabIndex        =   32
         TabStop         =   0   'False
         ToolTipText     =   "«·”«»ﬁ"
         Top             =   180
         Width           =   600
      End
      Begin VB.CommandButton cmdLast 
         Height          =   375
         Left            =   2070
         Picture         =   "transnew.frx":4B93
         Style           =   1  'Graphical
         TabIndex        =   31
         TabStop         =   0   'False
         ToolTipText     =   "√ŒÌ—"
         Top             =   180
         Width           =   510
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   375
         Left            =   45
         Picture         =   "transnew.frx":726D
         Style           =   1  'Graphical
         TabIndex        =   30
         TabStop         =   0   'False
         ToolTipText     =   "√Ê·"
         Top             =   180
         Width           =   510
      End
   End
   Begin VB.Frame Frame4 
      Height          =   555
      Left            =   1485
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   1035
      Width           =   4200
      Begin VB.Label xTq 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   135
         Width           =   1770
      End
      Begin VB.Label xCount 
         Alignment       =   2  'Center
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   135
         Width           =   1770
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   9675
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   0
      Width           =   5460
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "transnew.frx":994C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         TabStop         =   0   'False
         ToolTipText     =   "Œ—ÊÃ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1410
         MaskColor       =   &H00FFFFFF&
         Picture         =   "transnew.frx":9A96
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Õ–›"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   2730
         MaskColor       =   &H00FFFFFF&
         Picture         =   "transnew.frx":C330
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   23
         TabStop         =   0   'False
         ToolTipText     =   "«÷«›…"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4050
         Picture         =   "transnew.frx":E8DC
         Style           =   1  'Graphical
         TabIndex        =   22
         TabStop         =   0   'False
         ToolTipText     =   "«” ⁄·«„"
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame9 
      Height          =   645
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   0
      Width           =   5640
      Begin Threed.SSCommand cmdPrint 
         Height          =   465
         Left            =   3915
         TabIndex        =   18
         Top             =   135
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   820
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "transnew.frx":110AF
         Caption         =   "  ÿ»«⁄… «·„” ‰œ"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand CMD_PrintBar 
         Height          =   465
         Left            =   45
         TabIndex        =   19
         Top             =   135
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   820
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "transnew.frx":134BE
         Caption         =   " ÿ»«⁄… »«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
      Begin Threed.SSCommand CMD_BAR 
         Height          =   465
         Left            =   1935
         TabIndex        =   20
         Top             =   135
         Width           =   1950
         _ExtentX        =   3440
         _ExtentY        =   820
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "transnew.frx":15E40
         Caption         =   "  ÕÊÌ· ··»«—ﬂÊœ"
         Alignment       =   1
         PictureAlignment=   3
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1050
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   540
      Width           =   1410
      Begin VB.CommandButton CmdUndo 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "transnew.frx":187D9
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   585
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton CmdSave 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "transnew.frx":1AD52
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame5 
      Height          =   645
      Left            =   5715
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   3930
      Begin VB.CommandButton CMD_BAL 
         Caption         =   " ÕÊÌ· √—’œ… ·„Œ“‰"
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   1860
      End
      Begin VB.CommandButton TRANS_MODEL 
         Caption         =   " ÕÊÌ· „ÊœÌ·"
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
         Left            =   2070
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   180
         Width           =   1770
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   8625
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   960
      Left            =   5700
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   630
      Width           =   9465
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
         Left            =   6570
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1470
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
         Left            =   1350
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo xStore1 
         Height          =   315
         Left            =   5310
         TabIndex        =   5
         Top             =   540
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xStore2 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   540
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   "DataCombo1"
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "≈·Ì „Œ“‰ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   630
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„‰ „Œ“‰ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   8145
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   255
         Width           =   930
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   0
      Top             =   -300
      Visible         =   0   'False
      Width           =   1890
      _ExtentX        =   3334
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
      Caption         =   "data1"
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
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   1620
      Top             =   495
      Visible         =   0   'False
      Width           =   1290
      _ExtentX        =   2275
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
      Caption         =   "data1"
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
      Height          =   7440
      Left            =   90
      TabIndex        =   10
      Top             =   1620
      Width           =   15090
      _cx             =   26617
      _cy             =   13123
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
Attribute VB_Name = "Transfrm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim Search31 As New Search3, search32 As New Search3
Dim CardTable As ADODB.Recordset
Dim tBalStore  As New ADODB.Recordset
Dim ItemTable As New ADODB.Recordset
Public myPublic As Integer

Dim formMode, dDateLast As String
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(3, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.Text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = DateSq(xDate.Text)

aInsert(2, 0) = "store1"
aInsert(2, 1) = addstring(xStore1.BoundText)

aInsert(3, 0) = "store2"
aInsert(3, 1) = addstring(xStore2.BoundText)

con.BeginTrans
If xDoc_No.Tag = DefineMode Then
    xDoc_No.Text = RetZero(Val(Newflag("FILE1_60h", "doc_no")))
    aInsert(0, 1) = addstring(xDoc_No.Text)
    con.Execute CreateInsert(aInsert, "FILE1_60h")
Else
    con.Execute CreateUpdate(aInsert, "FILE1_60h", " where doc_no = " & addstring(xDoc_No.Text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Sub myProc()
'On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    nFound = grid1.FindRow(Search3.grid1.TextMatrix(Search3.grid1.Row, 0), , 0)
    If nFound <> -1 Then
        If MsgBox("«·’‰› „ÊÃÊœ ›Ï ﬁ»· ›Ï «·”ÿ— " & nFound & " √÷«›… ‰⁄„ «„ ·« ", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If

    grid1.TextMatrix(grid1.Row, 0) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    GrdDesc grid1.Row
    If grid1.Row = grid1.Rows - 1 Then
        grid1.TextMatrix(grid1.Rows - 1, 8) = ""
        grid1.AddItem ""
        grid1.Select grid1.Rows - 1, 0
    ElseIf grid1.Row = grid1.Rows - 2 Then
        grid1.TextMatrix(grid1.Rows - 2, 8) = ""
        grid1.Select grid1.Rows - 1, 0
    End If
    
Else
    xDoc_No.Text = Search31.grid1.TextMatrix(Search31.grid1.Row, 0)
    CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    If Not CardTable.EOF Then myload
    Unload Search31
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
Unload Search
End Sub
Private Sub CMD_BAL_Click()
    If Not MYVALID2 Then Exit Sub
    Trans_Balnce.Show 1
    CardTable.Requery
    CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
End Sub

Private Sub CMD_BAR_Click()
On Error GoTo myerror
    With grid1
        con.BeginTrans
        For I = 1 To .Rows - 1
            If Val(.TextMatrix(I, 8)) > 0 Then
                con.Execute "Insert Into ADDPRINT(Item,Quant,isPrint) " & _
                    " Values(" & _
                    addstring(.TextMatrix(I, 0)) & "," & _
                    addvalue(.TextMatrix(I, 8)) & "," & _
                    "1" & _
                    ")"
            End If
        Next I
        con.CommitTrans
    End With
    Inform " „ «· —ÕÌ·"
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub CMD_PrintBar_Click()
    Morshed_Bar96.Show
End Sub
Private Sub cmdDelinv_Click()
    If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        con.Execute " Delete  From FILE1_60 where Doc_No = " & MyParn(xDoc_No.Text)
        con.Execute " Delete  From FILE1_60H where Doc_No = " & MyParn(xDoc_No.Text)
        con.CommitTrans
        CardTable.Requery
        
        CmdNewInv_Click
        Inform " „ Õ–› «·„” ‰œ »‰Ã«Õ"
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
Private Sub CmdInform_Click()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me
If myPublic = 0 Then
    Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA,FILE0_40_1.DESCA " & _
                      " FROM (FILE1_60H INNER JOIN FILE0_40 ON FILE1_60H.Store1 = FILE0_40.CODE) INNER JOIN FILE0_40 AS FILE0_40_1 ON FILE1_60H.STORE2 = FILE0_40_1.CODE "
Else
    Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),FILE0_40.DESCA,FILE0_40_1.DESCA " & _
                      " FROM (FILE1_61H INNER JOIN FILE0_40 ON FILE1_61H.Store1 = FILE0_40.CODE) INNER JOIN FILE0_40 AS FILE0_40_1 ON FILE1_61H.STORE2 = FILE0_40_1.CODE "
End If
Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False


listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "„‰ „Œ“‰"
GrdArray(3, 1) = 2000

GrdArray(4, 0) = "≈·Ì „Œ“‰"
GrdArray(4, 1) = 2000

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
Private Sub CmdNewInv_Click()
    mydefine
    On Error Resume Next
    xDoc_No.SetFocus
    Err.Clear
End Sub
Private Sub cmdPrint_Click()
    doprint
End Sub
Private Sub cmdSave_Click()
foundOther
If Not myvalid Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Requery
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub CmdUndo_Click()
If CardTable.BOF And CardTable.EOF Then
    mydefine
    Exit Sub
End If
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Shift = 2 And KeyCode = 83 Then cmdSave_Click
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
bEdit = True
con.CursorLocation = adUseClient
con.Open strCon
Set CardTable = New ADODB.Recordset
If myPublic = 0 Then
    CardTable.Open "SELECT * FROM FILE1_60H ORDER BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText
Else
    Frame5.Visible = False
    cmdNewInv.Visible = False
    CmdDelInv.Visible = False
    cmdSave.Visible = False
    CardTable.Open "SELECT * FROM FILE1_61H ORDER BY DOC_NO", con, adOpenStatic, adLockReadOnly, adCmdText
End If
ItemTable.Open "FILE1_10", con, adOpenStatic, adLockReadOnly, adCmdTable
data1.ConnectionString = strCon
data1.RecordSource = "FILE0_40"
Set xStore1.RowSource = data1
xStore1.ListField = "Desca"
xStore1.BoundColumn = "Code"

Set xStore2.RowSource = data1
xStore2.ListField = "Desca"
xStore2.BoundColumn = "Code"

xStore2.Enabled = False
xStore1.BoundText = cPStore
Set grid1.DataSource = DATA10
DATA10.ConnectionString = strCon
If myPublic = 0 Then
    CmdNewInv_Click
Else
    If Not (CardTable.EOF And CardTable.BOF) Then
        CardTable.MoveLast
        myload
    Else
        mydefine
        xDoc_No.Text = ""
    End If
End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Unload Search3
Unload Search31
If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
CardTable.Close
tBalStore.Close
Set CardTable = Nothing
Set tBalStore = Nothing
closeCon con
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If grid1.Col = 0 Then GrdDesc grid1.Row
CalcTotals

End Sub
Private Sub grid1_EnterCell()
With grid1
If myPublic = 1 Then
    grid1.Editable = flexEDNone
    Exit Sub
End If
If grid1.Col = 0 Or grid1.Col = 8 Then
    .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = vbWhite
    .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = vbYellow
    grid1.Editable = flexEDKbdMouse
Else
    grid1.Editable = flexEDNone
End If
End With
End Sub
Private Sub Grid1_GotFocus()
With grid1
    If grid1.Row <= 1 Then
    .Select 1, 0, 1, 0
    .ShowCell 1, 0
    End If
End With
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If myPublic = 1 Then Exit Sub
    If KeyCode = 45 And grid1.Row <> grid1.Rows - 1 Then grid1.AddItem "", grid1.Row
End Sub
Private Sub Grid1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
If myPublic = 1 Then Exit Sub

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
            grid1.Col = 8
            grid1.Row = Row
        Case 8
            If grid1.TextMatrix(grid1.Row + 1, 0) = "" Then
                grid1.Row = Row + 1
                grid1.Col = 0
            Else
                grid1.Row = Row + 1
                grid1.Col = 8
            End If
     End Select
End If
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If grid1.Row = grid1.Rows - 1 Then grid1.AddItem ""
If Col = 2 And Trim(grid1.TextMatrix(Row, 3)) = "" Then
nBalance = LastBalance(rid1.TextMatrix(Row, 0), xStore1.BoundText, con)
grid1.TextMatrix(Row, 9) = nBalance
End If

End Sub
Private Sub grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If Col = 0 And Trim(grid1.EditText) <> "" Then
    ItemTable.Find " ITEM = " & grid1.EditText, , adSearchForward, adBookmarkFirst
    If ItemTable.EOF Then
        MsgBox "»«—ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
        Cancel = True
    End If
    nFound = FoundOtheritem(Row, 0, Trim(grid1.EditText))
    If nFound <> -1 Then
        MsgBox "«·’‰› „ÊÃÊœ ›Ì «·”ÿ— —ﬁ„ " & nFound
        Cancel = True
    End If
End If
If Col = 8 Then
    If Val(grid1.EditText) >= Val(grid1.TextMatrix(Row, 9)) Then
        MsgBox "«·—’Ìœ ·« Ì”„Õ"
    End If
End If
End Sub

Private Sub TRANS_MODEL_Click()
    If Not MYVALID2 Then Exit Sub
    TransModelfrm.Show 1
    CardTable.Requery
    CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload

End Sub

Private Sub xDate_GotFocus()
xDate.SelStart = 0
xDate.SelLength = Len(xDate.Text)
End Sub
Private Sub xDoc_No_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CmdInform_Click
End Sub
Private Function myvalid() As Boolean
If xDoc_No.Text = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
If xStore1.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·«Ê·"
    Exit Function
End If

If xStore2.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·À«‰Ì"
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
        MsgBox "ﬂÊœ «·’‰› €Ì— „”Ã·"
        Exit Function
    Else
        cItem = GetDesca("select item from file1_10 where item = " & MyParn(.TextMatrix(I, 0))) & ""
        If cItem = "" Then
            MsgBox "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
            Exit Function
        End If
    End If
    If Val(.TextMatrix(I, 8)) = 0 Then
        .Select I, 0, I, grid1.Cols - 1
        MsgBox "ﬂ„Ì… «·’‰› €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
myvalid = True
End Function

Private Function MYVALID2() As Boolean
If Not IsDate(xDate.Text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If
If xStore1.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·«Ê·"
    Exit Function
End If

If xStore2.BoundText = "" Then
    MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ «·À«‰Ì"
    Exit Function
End If
MYVALID2 = True
End Function
Private Sub myload()
xDoc_No.Text = CardTable!doc_no
xDate.Text = Format(CardTable!Date, "dd-mm-yyyy")
xStore1.BoundText = CardTable!STORE1
xStore2.BoundText = CardTable!STORE2
xDoc_No.Enabled = False
myloadgrd
Handlecontrols LoadMode
CalcTotals
End Sub
Private Sub mydefine()
xDoc_No.Enabled = True
If myPublic = 0 Then
    xDoc_No.Text = RetZero(Val(Newflag("FILE1_60h", "doc_no")))
End If
xDate.Text = Format(Date, "dd-mm-yyyy")
xStore2.BoundText = ""
xStore1.Enabled = False
xStore2.Enabled = True
xStore1.BoundText = cPStore
StatusBar1.Panels(1).Text = ""
xCount.Caption = ""
xTq.Caption = ""
grid1.Rows = 1
grid1.AddItem ""
Handlecontrols DefineMode
Fixgrd
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode And bEdit
cmdSave.Enabled = (bEdit)
CmdDelInv.Enabled = nMode = LoadMode And bEdit
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
End Sub
Private Sub xDoc_No_LostFocus()
If xDoc_No.Text = "" Then Exit Sub
xDoc_No.Text = RetZero(xDoc_No.Text)
If CardTable.BOF And CardTable.BOF Then Exit Sub
CardTable.Find "doc_no = " & MyParn(xDoc_No.Text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Sub GRID1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 0 Then
    ItemsLookupAll Me, Search3
End If

If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.BeginTrans
            con.Execute "delete from file1_60 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
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
If myPublic = 1 Then Exit Sub
Select Case Col
    Case 0
        If KeyCode = 27 Then Exit Sub
        If KeyCode = 112 Then ItemsLookupAll Me, Search3
End Select
End Sub
Private Sub GrdDesc(Row)
Dim nBalance As Double
If grid1.TextMatrix(Row, 0) = "" Then Exit Sub
nBalance = LastBalance(grid1.TextMatrix(Row, 0), xStore1.BoundText, con)
grid1.TextMatrix(Row, 9) = nBalance

With grid1
    ItemTable.Find "item = " & grid1.TextMatrix(Row, 0), , adSearchForward, adBookmarkFirst
    If Not ItemTable.EOF Then
        grid1.TextMatrix(grid1.Row, 1) = ItemTable!MOSM
        grid1.TextMatrix(grid1.Row, 2) = GetDesca("select desca from fact where code = " & MyParn(ItemTable!Fact))
        grid1.TextMatrix(grid1.Row, 3) = ItemTable!SUPP & ""
        grid1.TextMatrix(grid1.Row, 4) = ItemTable!MODELFACT0
        grid1.TextMatrix(grid1.Row, 5) = ItemTable!DESCA
        grid1.TextMatrix(grid1.Row, 6) = ItemTable!Color
        grid1.TextMatrix(grid1.Row, 7) = ItemTable!scal
    End If
End With
If Val(grid1.TextMatrix(Row, 9)) = 0 Then MsgBox "·« ÌÊÃœ —’Ìœ ··’‰›"
End Sub
Private Function CalcTotals()
Dim nCount As Double
Dim nTotalQuant As Double, nTotalCost As Double
With grid1
For I = 1 To grid1.Rows - 2
    nCount = nCount + 1
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(I, 8))
Next
'e.StatusBar1.Panels(1).Text = IIf(nTotalQuant <> 0, "≈Ã„«·Ì «·ﬂ„Ì… : " & nTotalQuant, "")
xCount.Caption = "⁄œœ «·«’‰«› : " & Format(nCount, "#0")
xTq.Caption = "«·ﬂÌ„…  : " & Format(nTotalQuant, "#0")
End With
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
Private Sub foundOther()
For I = 1 To grid1.Rows - 2
    nRow = FoundOtherRow(I, 0)
    If nRow <> -1 Then
        MsgBox "«·’‰› " & grid1.TextMatrix(nRow, 1) & " „ﬂ—— " & "›Ï «·”ÿ— —ﬁ„ ====> " & nRow
        Exit Sub
    End If
Next
End Sub
Private Sub doprint()
Dim temptable As New ADODB.Recordset
Dim sourceTable As New ADODB.Recordset
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
For I = 1 To grid1.Rows - 2
    temptable.AddNew
    temptable!str6 = xDoc_No.Text
    temptable!Date1 = xDate.Text
    temptable!str14 = TurnValue(xStore1.Text)
    temptable!str15 = TurnValue(xStore2.Text)
    
    temptable!str4 = TurnValue(grid1.TextMatrix(I, 0))
    temptable!str11 = TurnValue(grid1.TextMatrix(I, 1))
    temptable!str12 = TurnValue(grid1.TextMatrix(I, 4))
    temptable!str10 = TurnValue(grid1.TextMatrix(I, 2))
    temptable!str2 = TurnValue(grid1.TextMatrix(I, 5))
    temptable!str5 = TurnValue(grid1.TextMatrix(I, 6))
    temptable!str7 = TurnValue(grid1.TextMatrix(I, 7))
    temptable!val1 = Val(grid1.TextMatrix(I, 8))
    temptable!val2 = Val(GetDesca("select price from file1_10 where item = " & grid1.TextMatrix(I, 0)) & "")
    temptable!str1 = TurnValue(grid1.TextMatrix(I, 2)) & TurnValue(grid1.TextMatrix(I, 1)) & RetZero(TurnValue(grid1.TextMatrix(I, 4)), 4)
    temptable!Val10 = I
    temptable.Update
Next
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
main.Report1.ReportFileName = App.Path & "\Reports\TRANS.rpt"
main.Report1.DataFiles(0) = tempFile
main.Report1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
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
Private Sub Fixgrd()
With grid1
'                   0            1      2               3        4           5       6           7          8
.FormatString = "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "«·ﬂ„Ì…|" & "«·—’Ìœ|"
.WordWrap = True

.RowHeight(0) = 1000

.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 2000
.ColWidth(3) = 1000
.ColWidth(4) = 1500
.ColWidth(5) = 4000
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 0
'.ColComboList(2) = StrList("SELECT CODE , DESCA FROM FACT ")
For I = 0 To .Cols - 1
    .ColAlignment(I) = flexAlignRightCenter
Next
End With
End Sub
Private Sub myreplaceGrd()
Dim aInsert(3, 1)
With grid1
    For I = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.Text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addvalue(grid1.TextMatrix(I, 0))
        
        aInsert(2, 0) = "quant"
        aInsert(2, 1) = Val(.TextMatrix(I, 8))
        
        aInsert(3, 0) = "COST"
        aInsert(3, 1) = Val(GetDesca("SELECT COST FROM FILE1_10 WHERE ITEM = " & Val(grid1.TextMatrix(I, 0))) & "")
        
        
        If grid1.TextMatrix(I, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "FILE1_60")
        Else
            con.Execute CreateUpdate(aInsert, "FILE1_60", " where ID = " & grid1.TextMatrix(I, .Cols - 1))
        End If
    Next
End With
End Sub
Private Sub myloadgrd()
'                    0               1               2               3           4                       5               6               7               8                   9
If myPublic = 0 Then
    cString = "SELECT FILE1_60.item , file1_10.mosm , fact.desca , file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  FILE1_60.quant , ' ' as bal ,FILE1_60.ID from (FILE1_60 inner join file1_10 on file1_10.item = FILE1_60.item ) inner join fact on fact.code = file1_10.fact WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " ORDER BY ID "
Else
    cString = "SELECT FILE1_61.item , file1_10.mosm , fact.desca , file1_10.supp, file1_10.modelfact0 ,file1_10.desca , file1_10.scal , file1_10.color ,  FILE1_61.quant , ' ' as bal ,FILE1_61.ID from (FILE1_61 inner join file1_10 on file1_10.item = FILE1_61.item ) inner join fact on fact.code = file1_10.fact WHERE DOC_NO = " & MyParn(xDoc_No.Text) & " ORDER BY ID "
End If
DATA10.RecordSource = cString
DATA10.Refresh
grid1.AddItem ""
Fixgrd
End Sub


