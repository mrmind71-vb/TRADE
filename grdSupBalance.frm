VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdSupBalance 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ì «—’œ… „Ê—œÌ‰"
   ClientHeight    =   10290
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10290
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   270
      Width           =   6900
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdSupBalance.frx":0000
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1170
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
         _Version        =   196610
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdSupBalance.frx":2323
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdSupBalance.frx":4699
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   5760
         TabIndex        =   4
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdSupBalance.frx":681C
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2295
         TabIndex        =   6
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdSupBalance.frx":96CF
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPdf 
         Height          =   555
         Left            =   3420
         TabIndex        =   5
         Top             =   135
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdSupBalance.frx":BBFA
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdSql 
         Height          =   555
         Left            =   4545
         TabIndex        =   22
         Top             =   135
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   979
         _Version        =   196610
         ForeColor       =   0
         BackColor       =   16777215
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "grdSupBalance.frx":E0A7
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1005
      Left            =   9450
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   0
      Width           =   10815
      Begin VB.CheckBox chkMoveOnly 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "„Ê—œÌ‰ ·œÌÂ„ Õ—ﬂ«  ›ﬁÿ"
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
         Height          =   240
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   630
         Width           =   2175
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8055
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6210
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xSup 
         Height          =   330
         Left            =   6210
         TabIndex        =   2
         Top             =   555
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   330
         Left            =   135
         TabIndex        =   3
         Top             =   135
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         BackColor       =   -2147483643
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·› —…"
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
         Height          =   270
         Left            =   10035
         TabIndex        =   14
         Top             =   225
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "„Ã„Ê⁄… „Ê—œÌ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   1680
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "«·„Ê—œ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9975
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   540
         Width           =   645
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -270
      Top             =   -225
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Left            =   45
      Top             =   -90
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Height          =   7800
      Left            =   90
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1035
      Width           =   20220
      _cx             =   35666
      _cy             =   13758
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
      GridColor       =   12632256
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   16
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
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   405
      Top             =   225
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   15
      Top             =   9825
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   714
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   1
         Left            =   4095
         TabIndex        =   17
         Top             =   45
         Width           =   4005
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   2
         Left            =   8100
         TabIndex        =   18
         Top             =   45
         Width           =   4000
         _ExtentX        =   7064
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   3
         Left            =   12150
         TabIndex        =   19
         Top             =   45
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
      Begin Threed.SSPanel panel1 
         Height          =   330
         Index           =   4
         Left            =   16155
         TabIndex        =   20
         Top             =   45
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelOuter      =   1
         RoundedCorners  =   0   'False
         FloodShowPct    =   -1  'True
      End
   End
   Begin ComctlLib.ProgressBar prog1 
      Align           =   2  'Align Bottom
      Height          =   195
      Left            =   0
      TabIndex        =   21
      Top             =   9630
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "grdSupBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim FieldTable As ADODB.Recordset
Dim cString As String, aHeader()
Private Sub CMD_EXEL_Click()
End Sub
Private Sub CMD_PRINT_Click()
End Sub

Private Sub Check1_Click()

End Sub

Private Sub cmdExcel_Click()
Me.MousePointer = 11
Fixgrd False
Dim aRow As Variant
aSub = AddFlag(Empty, "row", 0)
aSub = AddFlag(aSub, "bold", True)
aSub = AddFlag(aSub, "word_wrap", True)
aRow = AddFlag(aRow, aSub)
If grid1.Rows > 1 Then
    aSub = AddFlag(Empty, "row", 1)
    aSub = AddFlag(aSub, "bold", True)
    aSub = AddFlag(aSub, "word_wrap", False)
    aSub = AddFlag(aSub, "back_color", 40)
    aRow = AddFlag(aRow, aSub)
End If
ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption, retHeader(aHeader, 0, 4), retHeader(aHeader, 4, 4), retHeader(aHeader, 8, 4))
Fixgrd True
Me.MousePointer = 0
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
If Not MYVALID Then Exit Sub
Me.MousePointer = vbHourglass
myload
Me.MousePointer = vbNormal
End Sub
Private Function MYVALID() As Boolean
If Not IsDate(xDate1.text) Then
    MsgBox " «—ÌŒ «·„œ… «·«Ê· €Ì— „Õœœ"
    Exit Function
End If
If Not IsDate(xDate2.text) Then
    MsgBox " «—ÌŒ «·„œ… «·À«‰Ì €Ì— „Õœœ"
    Exit Function
End If
MYVALID = True
End Function
Private Sub cmdPdf_Click()
myPrint "1", True
End Sub
Private Sub cmdPrint_Click()
myPrint
End Sub

Private Sub cmdSql_Click()
myload True
End Sub

Private Sub Form_Load()
openCon con

Set data1.Recordset = myCmd("Select Code,DescA From File4_50 ORDER BY code ", con)
Set xGroup.RowSource = data1
xGroup.ListField = "Desca"
xGroup.BoundColumn = "Code"


Set DATA2.Recordset = myCmd("Select Code,DescA From File4_10 order by Desca", con)
Set xSup.RowSource = DATA2
xSup.ListField = "Desca"
xSup.BoundColumn = "Code"

Set grid1.DataSource = DATA11

Fixgrd

xDate1.text = myFormat_p("1-1-" & Year(Date))
xDate2.text = myFormat_p(Date)

End Sub
Private Sub myload(Optional bString As Boolean)
With grid1
ReDim aHeader(13)

Dim aPrm As Variant
If IsDate(xDate1.text) Then
    aPrm = AddFlag(aPrm, "DATE1", addDate(xDate1.text))
    aHeader(0) = BetweenString(xDate1.text, xDate2.text)
End If

If IsDate(xDate2.text) Then
    aPrm = AddFlag(aPrm, "DATE2", addDate(xDate2.text))
    aHeader(0) = BetweenString(xDate1.text, xDate2.text)
End If

If xSup.MatchedWithList Then
    aPrm = AddFlag(aPrm, "SUP", addstring(xSup.BoundText))
    aHeader(2) = "«·„Ê—œ : " & xSupp.text
End If

If xGroup.MatchedWithList Then
    aPrm = AddFlag(aPrm, "GROUP", addstring(xGroup.BoundText))
    aHeader(2) = "„Ã„Ê⁄… „Ê—œÌ‰ : " & xGroup.text
End If

If chkMoveOnly.Value = 1 Then
    aPrm = AddFlag(aPrm, "MOVE_ONLY", "1")
    aHeader(3) = "⁄„·«¡ ·œÌÂ„ —’Ìœ «Ê Õ—ﬂ… : "
End If

Dim cString As String
cString = myPrcString("[dbo].[sp_sup_balance]", aPrm)

If bString Then
    Clipboard.Clear
    Clipboard.SetText cString
    Exit Sub
End If

Set DATA11.Recordset = myCmd(cString, con, , , 600)
End With
Fixgrd
panel1(0).Caption = "⁄œœ «·”Ã·«  : " & DATA11.Recordset.RecordCount
End Sub
Sub Fixgrd(Optional bGroup As Boolean = True)
With grid1
.RowHeight(0) = 1000
.WordWrap = True
    .FrozenCols = 2
    
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    
    .TextMatrix(0, 2) = "«·—’Ìœ " & xDate1.text
    .TextMatrix(0, 3) = "Ã. „‘ —Ì« "
    .TextMatrix(0, 4) = "Ã. „— Ã⁄«  "
    .TextMatrix(0, 5) = "’«›Ï ﬁÌ„… „‘ —Ì« "
    
    .TextMatrix(0, 6) = "”œ«œ ‰ﬁœÏ"
    .TextMatrix(0, 7) = "”œ«œ √.œ"
    .TextMatrix(0, 8) = " ”ÊÌ« "
    .TextMatrix(0, 9) = "„’«—Ì›"
    .TextMatrix(0, 10) = "Õ”«»« "
    .TextMatrix(0, 11) = "—’Ìœ " & xDate2.text
    .TextMatrix(0, 12) = "√.œ  Õ  «· Õ’Ì·"
    
    .TextMatrix(0, 13) = "—’Ìœ »÷«⁄… «„«‰« "
    .TextMatrix(0, 14) = "«·„” Õﬁ"
    .TextMatrix(0, 15) = " ﬂ·›… „»Ì⁄«  «„«‰« "
    
    .ColWidth(0) = 900
    .ColWidth(1) = 2000
    .ColWidth(2) = 1400
    .ColWidth(3) = 1400
    .ColWidth(4) = 1400
    
    .ColWidth(5) = 1300
    .ColWidth(6) = 1300
    .ColWidth(7) = 1200
    .ColWidth(8) = 1100
    .ColWidth(9) = 1100
    .ColWidth(10) = 1200
    .ColWidth(11) = 1500
    .ColWidth(12) = 1200
    .ColWidth(13) = 1300
    .ColWidth(14) = 1300
    .ColWidth(15) = 1300
    
    For I = 2 To .Cols - 1
        .SubtotalPosition = flexSTAbove
        .Subtotal flexSTSum, -1, I, "##,##", vbYellow, vbRed, True, "«·≈Ã„«·Ì"
    Next

'    .ColFormat(16) = "##,##"
'    .ColFormat(17) = "##,##.##"
'    .ColFormat(18) = "##,##.##"
'    .ColFormat(19) = "##,##.##"
'
.ExplorerBar = flexExSort
.Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
'
'
'.SubtotalPosition = flexSTAbove
'If bGroup Then
'    For Col = 7 To 19
'        .Subtotal flexSTSum, 0, Col, "##,##", &HC0FFC0, vbBlack, True, "≈Ã„«·Ì %s"
'    Next
'End If

'.SubtotalPosition = flexSTAbove
'For Col = 7 To 19
'    .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
'Next
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
SaveText Me
closeCon con
Set grdSupBalance = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
myload
End Sub
Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
If grid1.Rows < 3 Then Exit Sub
Dim aRow As Variant, aSub As Variant

For I = 1 To grid1.Rows - 2
    If grid1.TextMatrix(I, 2) = "" Then
        aSub = AddFlag(Empty, "row", I)
        aSub = AddFlag(aSub, "col", 1)
        aSub = AddFlag(aSub, "cols", 2)
        aSub = AddFlag(aSub, "text", grid1.TextMatrix(I, 1))
        aRow = AddFlag(aRow, aSub)
    End If
Next

aSub = AddFlag(Empty, "row", grid1.Rows - 1)
aSub = AddFlag(aSub, "col", 1)
aSub = AddFlag(aSub, "cols", 2)
aSub = AddFlag(aSub, "text", "≈Ã„«·Ì")
aRow = AddFlag(aRow, aSub)

'PrintGrdNew.sDeviceType = pDevice
'PrintGrdNew.Vp.Preview = Not bIgPreview

'PrintGrdNew.bIgPreview = bIgPreview
PrintGrdNew.DOPRINT Me.grid1, 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption, retHeader(aHeader, 0, 2), , False, True, 10, , aRow, Array(1)
If Not bIgPreview Then

    PrintGrdNew.Show 1
Else
    Unload PrintGrdNew
    'PrintGrdNew.Show 1
    'Unload PrintGrdNew
End If
End Sub


Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xdate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xdate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub XMOSM2_GotFocus()
myGotFocus XMOSM2
End Sub
Private Sub XMOSM2_LostFocus()
myLostFocus XMOSM2
End Sub
Private Sub xModelFACT_GotFocus()
myGotFocus xModelFact
End Sub
Private Sub XModelFACT_LostFocus()
myLostFocus xModelFact
End Sub
Private Sub xFact_GotFocus()
myGotFocus xFact
End Sub
Private Sub xFACT_LostFocus()
myLostFocus xFact
If Not xFact.MatchedWithList Then xFact.BoundText = ""
End Sub
Private Sub xSupp_GotFocus()
myGotFocus xSupp
End Sub
Private Sub xSupp_LostFocus()
myLostFocus xSupp
If Not xSupp.MatchedWithList Then xSupp.BoundText = ""
End Sub
Private Sub xage_GotFocus()
myGotFocus xage
End Sub
Private Sub xage_LostFocus()
myLostFocus xage
If Not xage.MatchedWithList Then xage.BoundText = ""
End Sub
Private Sub xsex_GotFocus()
myGotFocus xsex
End Sub
Private Sub xsex_LostFocus()
myLostFocus xsex
If Not xsex.MatchedWithList Then xsex.BoundText = ""
End Sub
Private Sub xMosm_GotFocus()
myGotFocus xMosm
End Sub
Private Sub xMosm_LostFocus()
myLostFocus xMosm
If Not xMosm.MatchedWithList Then xMosm.BoundText = ""
End Sub
Private Sub XSTORE_GotFocus()
myGotFocus xStore
End Sub
Private Sub XSTORE_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub xmosm20_GotFocus()
myGotFocus xmosm20
End Sub
Private Sub xmosm20_LostFocus()
myLostFocus xmosm20
If Not xmosm20.MatchedWithList Then xmosm20.BoundText = ""
End Sub
