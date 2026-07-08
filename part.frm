VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "Mscomctl.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form Partfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "‰ﬁœÌ…"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13875
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
   ScaleHeight     =   8235
   ScaleWidth      =   13875
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.Frame FRM_CLOSED 
      Height          =   1815
      Left            =   45
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   0
      Width           =   2085
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·„” ‰œ „€·ﬁ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   135
         Visible         =   0   'False
         Width           =   1410
      End
      Begin Threed.SSCommand cmd_closed 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   90
         TabIndex        =   29
         Top             =   1125
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   1058
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "part.frx":0000
         Alignment       =   4
         PictureAlignment=   9
      End
      Begin Threed.SSCommand cmd_CLOSEDDATE 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   1035
         TabIndex        =   30
         Top             =   135
         Width           =   960
         _ExtentX        =   1693
         _ExtentY        =   1693
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "part.frx":25CC
         Caption         =   "≈€·«ﬁ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
      Begin Threed.SSCommand cmd_open 
         CausesValidation=   0   'False
         Height          =   960
         Left            =   90
         TabIndex        =   31
         Top             =   135
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   1693
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "part.frx":4C44
         Caption         =   "› Õ › —…"
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame Frame9 
      Height          =   870
      Left            =   2115
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   0
      Width           =   4155
      Begin VB.Label XTIME2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   26
         Top             =   495
         Width           =   2220
      End
      Begin VB.Label xUserName2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   495
         Width           =   1815
      End
      Begin VB.Label XTIME1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         TabIndex        =   24
         Top             =   135
         Width           =   2220
      End
      Begin VB.Label xUserName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   2295
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   135
         Width           =   1815
      End
      Begin VB.Label xusercode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   -270
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Frame Frame1 
      Height          =   645
      Left            =   8280
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   5505
      Begin VB.CommandButton CmdExit 
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "part.frx":7285
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdDelInv 
         Height          =   420
         Left            =   1440
         MaskColor       =   &H00FFFFFF&
         Picture         =   "part.frx":73CF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton cmdNewInv 
         Height          =   420
         Left            =   2790
         MaskColor       =   &H00FFFFFF&
         Picture         =   "part.frx":9C69
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdInform 
         Height          =   420
         Left            =   4095
         Picture         =   "part.frx":C215
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   135
         Width           =   1320
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1005
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   630
      Width           =   1500
      Begin VB.CommandButton CmdUndo 
         Height          =   375
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "part.frx":E9E8
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   540
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
      Begin VB.CommandButton CmdSave 
         Height          =   375
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         Picture         =   "part.frx":10F61
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1320
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   13875
      TabIndex        =   6
      Top             =   7380
      Width           =   13875
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   0
         Width           =   3975
         Begin VB.CommandButton cmdNext 
            Height          =   375
            Left            =   2985
            Picture         =   "part.frx":113A3
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            ToolTipText     =   "«· «·Ì"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdPrevious 
            Height          =   375
            Left            =   2070
            Picture         =   "part.frx":13963
            Style           =   1  'Graphical
            TabIndex        =   19
            TabStop         =   0   'False
            ToolTipText     =   "«·”«»ﬁ"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdLast 
            Height          =   375
            Left            =   915
            Picture         =   "part.frx":15F36
            Style           =   1  'Graphical
            TabIndex        =   18
            TabStop         =   0   'False
            ToolTipText     =   "√ŒÌ—"
            Top             =   135
            Width           =   915
         End
         Begin VB.CommandButton cmdFirst 
            Height          =   375
            Left            =   45
            Picture         =   "part.frx":18610
            Style           =   1  'Graphical
            TabIndex        =   17
            TabStop         =   0   'False
            ToolTipText     =   "√Ê·"
            Top             =   135
            Width           =   870
         End
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   5
      Top             =   7935
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "11:30 ’"
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   1050
      Left            =   9090
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   585
      Width           =   4740
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
         Left            =   2340
         MaxLength       =   8
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1290
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
         Left            =   2340
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   540
         Width           =   1290
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   555
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   210
         Width           =   930
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   5415
      Left            =   45
      TabIndex        =   7
      Top             =   1845
      Width           =   13740
      _cx             =   24236
      _cy             =   9551
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
   Begin MSAdodcLib.Adodc DATA1 
      Height          =   375
      Left            =   4950
      Top             =   675
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
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
   Begin VB.Label xbranch 
      Alignment       =   1  'Right Justify
      Caption         =   "Label2"
      Height          =   285
      Left            =   7605
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   180
      Visible         =   0   'False
      Width           =   510
   End
End
Attribute VB_Name = "Partfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myPublic As Byte, cStrBox As String
Public con As New ADODB.Connection
Public bedit As Boolean
Dim CardTable As ADODB.Recordset
Dim dLastdate As String, defBox As String
Dim DocField As String, dDateLast As String
Dim formMode, cFileHeader As String
Dim lCellButton As Boolean
Const LoadMode = 0, DefineMode = 1
Private Function myreplace() As Boolean
Dim aInsert(5, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.text)

aInsert(1, 0) = "[Date]"
aInsert(1, 1) = addDate(xDate.text)

aInsert(2, 0) = "branch"
aInsert(2, 1) = addstring(xbranch.Caption)

On Error GoTo myerror
con.BeginTrans
If xDoc_No.Enabled Then
    
    aInsert(3, 0) = "username"
    aInsert(3, 1) = addstring(cusername)
    aInsert(4, 0) = "time"
    aInsert(4, 1) = "getdate()"
    
    If lIsBranchStore Or cBranch = "00" Then
        xDoc_No.text = RetZero(Val(NewflagBranch("FILE8_70H", "doc_no", cBranch, con)))
    Else
        xDoc_No.text = RetZero(Val(Newflag("FILE8_70H", "doc_no", con)))
    End If
    aInsert(0, 1) = addstring(xDoc_No.text)
    con.Execute CreateInsert(aInsert, "FILE8_70H")
Else
    aInsert(3, 0) = "username2"
    aInsert(3, 1) = addstring(cusername)
    aInsert(4, 0) = "time2"
    aInsert(4, 1) = "getdate()"
    
    con.Execute CreateUpdate(aInsert, "FILE8_70H", " where doc_no = " & addstring(xDoc_No.text))
End If
myreplaceGrd
con.CommitTrans
myreplace = True
Exit Function
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Function
Private Sub myreplaceGrd()
Dim aInsert(6, 1)
With grid1
    For i = 1 To .Rows - 2
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.text)
               
        aInsert(1, 0) = "Box"
        aInsert(1, 1) = addstring(.TextMatrix(i, 0))
        
        aInsert(2, 0) = "Code"
        aInsert(2, 1) = addstring(grid1.TextMatrix(i, 1))
        
        aInsert(3, 0) = "Desca"
        aInsert(3, 1) = addstring(grid1.TextMatrix(i, 3))
        
        aInsert(4, 0) = "[value_M]"
        aInsert(4, 1) = Val(grid1.TextMatrix(i, 4))

        aInsert(5, 0) = "[value_P]"
        aInsert(5, 1) = Val(grid1.TextMatrix(i, 5))

        aInsert(6, 0) = "[Row]"
        aInsert(6, 1) = i

        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "FILE8_70")
        Else
            con.Execute CreateUpdate(aInsert, "FILE8_70", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    Next
End With
End Sub
Sub myProc()
If ActiveControl.Name = grid1.Name Then
    If grid1.Col = 1 Then
        grid1.TextMatrix(grid1.Row, 1) = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
        grid1.TextMatrix(grid1.Row, 2) = Search3.grid1.TextMatrix(Search3.grid1.Row, 1)
        If grid1.Row = grid1.Rows - 1 Then
            grid1.AddItem ""
            grid1.TextMatrix(grid1.Rows - 1, 0) = defBox
        End If
        Unload Search3
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "doc_No = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
    Unload Search3
End If
End Sub
Private Sub cmdDelinv_Click()
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    On Error GoTo myerror
    con.BeginTrans
    con.Execute "Delete  From FILE8_70 where Doc_No = " & MyParn(xDoc_No.text)
    con.Execute "Delete  From FILE8_70H where Doc_No = " & MyParn(xDoc_No.text)
    con.CommitTrans
    CardTable.Requery
    If CardTable.EOF And CardTable.EOF Then
        mydefine
    Else
        CardTable.Find "Doc_No < " & MyParn(xDoc_No.text), , adSearchBackward, adBookmarkLast
        If CardTable.BOF Then CardTable.MoveFirst
        myload
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub CmdFirst_Click()
CardTable.MoveFirst
myload
End Sub
Private Sub CardLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
cString = "SELECT FILE8_70H.Doc_No, Convert(varChar(10),FILE8_70H.Date,111),Min(FILE8_71.Desca) " & _
          " FROM (FILE8_70H inner join FILE8_70 on FILE8_70H.doc_no = FILE8_70.Doc_NO) Inner Join FILE8_71 on FILE8_70.CODE = FILE8_71.Code "
If lIsBranchStore Then cString = cString & " where file8_70h.branch = " & MyParn(cBranch)
Generalarray(1) = cString
Generalarray(2) = " group by FILE8_70H.Doc_No,FILE8_70H.Date order by FILE8_70H.Doc_No,FILE8_70H.Date"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·«”„- «—ÌŒ «·„” ‰œ"
listarray(0, 1) = "(%%FILE8_71.Desca%% or " & _
                  " ##FILE8_70H.Date##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = " «—ÌŒ «·„” ‰œ"
GrdArray(1, 1) = 1500

GrdArray(2, 0) = "«·≈”„"
GrdArray(2, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
Search3.Caption = "«” ⁄·«„ Ã«—Ì «·‘—ﬂ«¡"
Search3.Show 1
End Sub
Private Sub CmdInform_Click()
    CardLookup
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
xDoc_No.SetFocus
End Sub
Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
myload
End Sub
Private Sub CmdUndo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    mydefine
Else
    If xDoc_No.Enabled Then CardTable.MoveLast Else CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
    myload
End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DBCombo Then SendKeys "{TAB}"
End If
End Sub
Private Sub Form_Load()
Me.Caption = DocTitle
openCon con
cStrBox = StrBox
FRM_CLOSED.Visible = Not lIsBranchStore
If cBranch = "00" Then
    con.Execute "   update file8_70h set branch = '00' where branch is null "
    con.Execute "   update file8_70h set doc_no = '00' + file8_70h.doc_no where len(doc_no) = 6"
    con.Execute "   update file8_70  set doc_no = '00' + file8_70.doc_no where len(doc_no) = 6"
End If
FRM_CLOSED.Visible = lSupperVisor Or (bopt2 And (cBranch = "00" Or cBranch > "60"))
'CmdDelInv.Visible = (cBranch = "00" Or cBranch > "60")

Set CardTable = New ADODB.Recordset
If lIsBranchStore Then
    CardTable.Open "SELECT * FROM FILE8_70H where branch = " & MyParn(cBranch) & " ORDER BY DOC_NO", con, adOpenStatic, adLockOptimistic, adCmdText
Else
    CardTable.Open "SELECT * FROM FILE8_70H ORDER BY DOC_NO", con, adOpenStatic, adLockOptimistic, adCmdText
End If
cFileHeader = "FILE8_70H"
Set grid1.DataSource = DATA1
DATA1.ConnectionString = strCon

StatusBar1.Panels.Add 1
StatusBar1.Panels(1).Width = 2000
'defBox = RetDefBox
If Not (CardTable.EOF And CardTable.BOF) Then
    CardTable.MoveLast
    myload
Else
    Fixgrd
    mydefine
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
CardTable.Close
Set CardTable = Nothing
closeCon con
End Sub
Private Sub grid1_EnterCell()
'If grid1.Col = 2 Or grid1.Col = 7 Or ((Val(grid1.TextMatrix(grid1.Row, 4)) <> 0 And grid1.Col = 5) Or (Val(grid1.TextMatrix(grid1.Row, 5)) <> 0 And grid1.Col = 4)) Then
If grid1.Col = 2 Or grid1.Col = 7 Then
    grid1.Editable = flexEDNone
Else
    grid1.Editable = flexEDKbdMouse
End If
    
End Sub

Private Sub Grid1_GotFocus()
If grid1.Row = 0 Then
    grid1.SetFocus
    grid1.Select 1, 0
End If
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 And grid1.Col = 1 Then grdLookup
If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And grid1.Rows > 3 And bedit Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        On Error GoTo myerror
        con.BeginTrans
        If grid1.TextMatrix(grid1.Row, grid1.Cols - 1) <> "" Then
            con.Execute "Delete from FILE8_70 where ID = " & grid1.TextMatrix(grid1.Row, grid1.Cols - 1)
        End If
        con.CommitTrans
        grid1.RemoveItem grid1.Row
    End If
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub
Private Sub grid1_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 112 And Col = 2 Then grdLookup
If KeyCode = 46 And Row <> grid1.Rows - 1 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        grid1.RemoveItem Row
    End If
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
Select Case grid1.Col
    Case 0
        If KeyCode = 112 Then grdLookup
End Select
End Sub
Private Function MYVALID() As Boolean
If Trim(xDoc_No.text) = "" Then
    MsgBox "—ﬁ„ «·„” ‰œ ·„ Ì”Ã·"
    Exit Function
End If

If cBranch <> "00" And cBranch < "60" Then
    If DateValue(xDate.text) < dMaxDate And Not lSupperVisor Then
        MsgBox " «—ÌŒ €Ì— ’«·Õ "
        Exit Function
    End If
End If

If Not IsDate(xDate.text) Then
    MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If grid1.Rows < 3 Then
    MsgBox "·«  ÊÃœ «’‰«›  „  ”ÃÌ·Â«"
    Exit Function
End If

With grid1
For i = 1 To .Rows - 2
    If .TextMatrix(i, 1) = "" Then
        .Select i, 0, i, grid1.Cols - 1
        MsgBox "ﬂÊœ " & sName & "  €Ì— „ÊÃÊœ"
        Exit Function
    End If
    If Val(.TextMatrix(i, 4)) = 0 And Val(.TextMatrix(i, 5)) = 0 Then
        MsgBox "«·ﬁÌ„… €Ì— „”Ã·…"
        Exit Function
    End If
Next
End With
MYVALID = True
End Function
Private Sub myload()

    If CardTable!ISCLOSED Then
        cmd_closed.BackColor = &H8080FF
        cmd_closed.Caption = "› Õ „” ‰œ"
        xClosed.Value = ssCBChecked
        xClosed.Visible = True
    Else
        cmd_closed.BackColor = vbGreen
        cmd_closed.Caption = "≈€·«ﬁ „” ‰œ"
        xClosed.Value = ssCBUnchecked
        xClosed.Visible = False
    End If
    
    xbranch.Caption = CardTable!branch & ""
    xUserName.Caption = CardTable!UserName & ""
    xUserName2.Caption = CardTable!UserName2 & ""
    XTIME1.Caption = CardTable!Time & ""
    XTIME2.Caption = CardTable!TIME2 & ""

xDoc_No.text = CardTable!doc_no
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")
Handlecontrols LoadMode
myloadgrd
End Sub
Private Sub myloadgrd()
With grid1
    cString = "SELECT FILE8_70.[BOX], FILE8_70.CODE, FILE8_71.DESCA, FILE8_70.desca, case when FILE8_70.VALUE_M = 0 THEN  NULL ELSE FILE8_70.VALUE_M END , CASE WHEN FILE8_70.VALUE_P = 0 THEN NULL ELSE FILE8_70.VALUE_P END ," & _
              " FILE8_70.ID " & _
              " FROM FILE8_70 LEFT JOIN FILE8_71 ON FILE8_70.CODE = FILE8_71.CODE" & _
              " where FILE8_70.doc_no = " & MyParn(xDoc_No.text) & " Order by Row"
    DATA1.RecordSource = cString
    DATA1.Refresh
    grid1.AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = defBox
End With
CalcTotals
Fixgrd
End Sub
Private Sub mydefine()
    
    xUserName.Caption = ""
    xUserName2.Caption = ""
    XTIME1.Caption = ""
    XTIME2.Caption = ""
    xbranch.Caption = cBranch
    cmd_closed.BackColor = &H8080FF
    cmd_closed.Caption = "› Õ „” ‰œ"
    xClosed.Visible = False
    xClosed.Value = ssCBUnchecked

    If lIsBranchStore Or cBranch = "00" Then
        xDoc_No.text = NewflagBranch("FILE8_70H", "doc_no", cBranch, con)
    Else
        xDoc_No.text = RetZero(Val(Newflag("FILE8_70H", "doc_no", con)))
    End If
    If cBranch = "00" Then
        xDate.text = ""
    Else
        xDate.text = Format(dSalesDate, "DD-MM-YYYY")
    End If
grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = defBox
Handlecontrols DefineMode
CalcTotals
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = (nMode = LoadMode)
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
CmdDelInv.Enabled = (nMode = LoadMode) And (xClosed.Value = 0)
cmdPrevious.Enabled = (nMode = LoadMode)
xDoc_No.Enabled = (nMode = DefineMode)
CmdSave.Enabled = (xClosed.Value = 0)
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
If Col = 1 Then
    If Trim(.EditText) <> "" Then .EditText = RetZero(.EditText, 2)
ElseIf Col = 4 And Val(.TextMatrix(Row, 5)) <> 0 And Val(.EditText) <> 0 Then
    .TextMatrix(Row, 5) = ""
ElseIf Col = 5 And Val(.TextMatrix(Row, 4)) <> 0 And Val(.EditText) <> 0 Then
    .TextMatrix(Row, 4) = ""
End If
End With
End Sub
Private Sub xDoc_No_LostFocus()
If Trim(xDoc_No.text) = "" Then Exit Sub
xDoc_No.text = RetZero(xDoc_No.text)
CardTable.Find "Doc_no = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload
End Sub
Private Function StrBox()
Dim boxtable As ADODB.Recordset
Set boxtable = New ADODB.Recordset
If lIsBranchStore Then
    boxtable.Open "SELECT * FROM file0_50 WHERE branch = " & MyParn(cBranch) & "   ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
Else
    boxtable.Open "SELECT * FROM file0_50 WHERE ISSTOP = 0  ORDER BY CODE ", con, adOpenStatic, adLockReadOnly, adCmdText
End If
If Not (boxtable.EOF And boxtable.BOF) Then
    boxtable.MoveFirst
    StrBox = "#  " & ";       "
    Do Until boxtable.EOF
        StrBox = StrBox & "|#" & boxtable!CODE & ";" & boxtable!DESCA
        boxtable.MoveNext
    Loop
End If
End Function
Private Sub grdLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select code ,DescA From FILE8_71"
Generalarray(2) = "Order by code"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·≈”„"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·≈”„"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1
End Sub
Private Function CalcTotals()
Dim nTotal As Double, nTotal2 As Double
With grid1
For i = 1 To grid1.Rows - 2
    nTotal = nTotal + Round(Val(grid1.TextMatrix(i, 4)), 2)
    nTotal2 = nTotal2 + Round(Val(grid1.TextMatrix(i, 5)), 2)
Next
StatusBar1.Panels(1).text = "≈Ã„«·Ì „”ÕÊ»«  : " & Format(nTotal, "Fixed")
StatusBar1.Panels(2).text = "≈Ã„«·Ì «Ìœ«⁄«  : " & Format(nTotal2, "Fixed")
End With
End Function
Private Sub GrdDesc(nRow)
    grid1.TextMatrix(nRow, 2) = GetDesca("Select Desca From FILE8_71 Where code = " & MyParn(grid1.TextMatrix(nRow, 1)), con) & ""
End Sub
Private Function RetDefBox() As String
Dim loctable As New ADODB.Recordset
loctable.Open "file0_50", con, adOpenStatic, adLockReadOnly, adCmdTable
If loctable.EOF And loctable.BOF Then Exit Function
loctable.MoveLast
If loctable.RecordCount = 1 Then
    loctable.MoveFirst
    RetDefBox = Trim(loctable!CODE & "")
End If
End Function
Private Sub xDoc_No_Validate(Cancel As Boolean)
If xDoc_No.text = "" Then Cancel = True
End Sub
Private Sub Fixgrd()
With grid1
    .Cols = 7
    .FormatString = "Œ“‰…|" & "ﬂÊœ «·‘—Ìﬂ|" & "≈”„ «·‘—Ìﬂ|" & "«·»Ì«‰|" & "„”ÕÊ»« |" & "«Ìœ«⁄« |"
    .ColWidth(0) = 2000
    .ColWidth(1) = 1000
    .ColWidth(2) = 3500
    .ColWidth(3) = 4000
    .ColWidth(4) = 1500
    .ColWidth(5) = 1500
    .ColHidden(.Cols - 1) = True
    
    For i = 1 To grid1.Cols - 1
        .ColAlignment(i) = flexAlignRightCenter
    Next
    .ColComboList(0) = cStrBox
End With
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
If OldRow <> NewRow And OldRow <> .Rows - 1 And OldRow <> 0 Then
    If Not validRow(OldRow) Then
        .RemoveItem OldRow
        CalcTotals
    End If
End If
End With
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
With grid1
If Not validRow(.Row) And .Row <> .Rows - 1 And .Row <> 0 Then
    .RemoveItem .Row
    CalcTotals
End If
End With
End Sub
Private Function validRow(nRow) As Boolean
With grid1
If Not (IsNumeric(.TextMatrix(nRow, 1)) And (Val(.TextMatrix(nRow, 4)) <> 0 Or Val(.TextMatrix(nRow, 5)) <> 0)) Then Exit Function
End With
validRow = True
End Function
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With grid1
If Col = 1 Then GrdDesc Row
If Not validRow(Row) Then Exit Sub
If Row = .Rows - 1 Then
    .AddItem ""
    grid1.TextMatrix(grid1.Rows - 1, 0) = defBox
    CalcTotals
End If
End With
End Sub
Private Sub cmd_CLOSEDDATE_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sFieldClose = "isclosed"
oClosefrm.sCaption = Me.Caption
oClosefrm.nMode = 0
oClosefrm.Show 1

CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload
End Sub
Private Sub cmd_open_Click()
Dim oClosefrm As New closefrm
oClosefrm.sFile = cFileHeader
oClosefrm.sFieldClose = "isclosed"
oClosefrm.sCaption = Me.Caption
oClosefrm.nMode = 1
oClosefrm.Show 1

CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload

End Sub
Private Sub cmd_closed_Click()
If CardTable!ISCLOSED Then
    con.Execute " update " & cFileHeader & " set isclosed = 0 where doc_no = " & MyParn(xDoc_No.text)
Else
    con.Execute " update " & cFileHeader & " set isclosed = 1 where doc_no = " & MyParn(xDoc_No.text)
End If
CardTable.Requery
CardTable.Find "Doc_No = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If CardTable.EOF Then CardTable.MoveLast
myload

End Sub

