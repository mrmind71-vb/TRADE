VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Begin VB.Form grdInvCollect 
   BackColor       =   &H00FFFFFF&
   Caption         =   "≈Ã„«·Ì  Õ’Ì· Œ·«· › —…"
   ClientHeight    =   10365
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   20370
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   20370
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   7980
      Left            =   45
      TabIndex        =   23
      Top             =   1080
      Width           =   20175
      _ExtentX        =   35586
      _ExtentY        =   14076
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "≈Ã„«·Ì „‰œÊ»Ì‰"
      TabPicture(0)   =   "grdInvCollect.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "grid2"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "≈Ã„«·Ì ⁄„·«¡"
      TabPicture(1)   =   "grdInvCollect.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7Ctl.VSFlexGrid grid1 
         Height          =   7530
         Left            =   90
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   360
         Width           =   19995
         _cx             =   35269
         _cy             =   13282
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
         Cols            =   4
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   7530
         Left            =   -74910
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   360
         Width           =   19995
         _cx             =   35269
         _cy             =   13282
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
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1755
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   315
      Width           =   5145
      Begin Threed.SSCommand cmdExit 
         Height          =   555
         Left            =   45
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
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
         Picture         =   "grdInvCollect.frx":0038
         Alignment       =   8
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdPrint 
         Height          =   555
         Left            =   1305
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
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
         Picture         =   "grdInvCollect.frx":235B
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         PictureDisabledFrames=   1
         PictureDisabled =   "grdInvCollect.frx":46D1
      End
      Begin Threed.SSCommand cmdGo 
         Height          =   555
         Left            =   3870
         TabIndex        =   6
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
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
         Picture         =   "grdInvCollect.frx":6854
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExcel 
         Height          =   555
         Left            =   2565
         TabIndex        =   15
         Top             =   135
         Width           =   1275
         _ExtentX        =   2249
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
         Picture         =   "grdInvCollect.frx":9707
         ButtonStyle     =   3
         PictureAlignment=   11
         BevelWidth      =   0
         ShapeSize       =   1
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   6975
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   90
      Width           =   13200
      Begin VB.TextBox Xcode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10710
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10095
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1815
      End
      Begin VB.TextBox xDate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   8235
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1815
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   90
         TabIndex        =   5
         Top             =   540
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xman 
         CausesValidation=   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„‰œÊ»"
         Height          =   270
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   225
         Width           =   600
      End
      Begin VB.Label xCustName 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   8235
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   540
         Width           =   2445
      End
      Begin VB.Label LLL 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ ⁄„Ì·"
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   1
         Left            =   12015
         TabIndex        =   10
         Top             =   585
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "„Ã„Ê⁄… ⁄„·«¡"
         Height          =   240
         Index           =   0
         Left            =   3465
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   630
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   12030
         TabIndex        =   8
         Top             =   225
         Width           =   660
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2475
      Top             =   75
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   150
      Top             =   75
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
   Begin Threed.SSPanel SSPanel1 
      Align           =   2  'Align Bottom
      Height          =   465
      Left            =   0
      TabIndex        =   16
      Top             =   9900
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   820
      _Version        =   196610
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RoundedCorners  =   0   'False
      FloodShowPct    =   -1  'True
      Begin Threed.SSPanel panel1 
         Height          =   405
         Index           =   0
         Left            =   0
         TabIndex        =   17
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
         TabIndex        =   18
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
         TabIndex        =   19
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
         TabIndex        =   20
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
         TabIndex        =   21
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
      TabIndex        =   11
      Top             =   9705
      Visible         =   0   'False
      Width           =   20370
      _ExtentX        =   35930
      _ExtentY        =   344
      _Version        =   327682
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "grdInvCollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aHeader(5)
Dim cWhere As String
Dim oSearch As New Search_abd
Dim oSearchtotal As New Search_total
Dim oSearchMan As New Search_total
Private Sub cmdExcel_Click()
Me.MousePointer = 11
If SSTab1.Tab = 1 Then
    If grid1.Rows > 1 Then
        aSub = AddFlag(Empty, "row", 1)
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "word_wrap", False)
        aSub = AddFlag(aSub, "back_color", 19)
        aRow = AddFlag(aRow, aSub)
    End If
    ToFileExelNew grid1, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption & "-" & SSTab1.TabCaption(1), retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2))
Else
    If grid2.Rows > 1 Then
        aSub = AddFlag(Empty, "row", 1)
        aSub = AddFlag(aSub, "bold", True)
        aSub = AddFlag(aSub, "word_wrap", False)
        aSub = AddFlag(aSub, "back_color", 19)
        aRow = AddFlag(aRow, aSub)
    End If
    ToFileExelNew grid2, , , aRow, Array(1), 0.9, , , , , , Me, Array(Me.Caption & "-" & SSTab1.TabCaption(0), retHeader(aHeader, 0, 2), retHeader(aHeader, 2, 2))
End If
Me.MousePointer = 0

End Sub

Private Sub CmdExit_Click()
Unload Me

End Sub
Private Sub CmdUndo_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
    myload2

End Sub

Private Sub cmdPdf_Click()

End Sub

Private Sub cmdPrint_Click()
myPrint
End Sub

Private Sub Form_Load()
    xDate1.text = myFormat_p("1-1-" & Year(Date))
    xDate2.text = myFormat_p(Date)
    
    Dim db As New clsDb
        Set xGroup.RowSource = db.myRs("SELECT * FROM FILE3_50")
        xGroup.ListField = "Desca"
        xGroup.BoundColumn = "Code"
                
        Set xman.RowSource = db.myRs("SELECT * FROM FILE6_25  WHERE isstop = 0  ORDER BY FILE6_25.DESCA")
        xman.ListField = "Desca"
        xman.BoundColumn = "Code"
                
        fixGrd
    Set db = Nothing
End Sub
Private Sub myload()
Dim strSql As String
strSql = "SELECT  FILE3_10.CODE," & _
         "FILE3_10.[DESCA]," & _
         "SUM(V.PAID)" & _
         " FROM  vw_inv_collect v" & _
         " INNER JOIN FILE3_10 ON v.code = FILE3_10.CODE"

cWhere = ""
If IsDate(xDate1.text) Then
    cWhere = "v.DATE >= " & DateSq(xDate1.text)
    aHeader(0) = BetweenString(xDate1.text, xDate2.text)
End If

If IsDate(xDate2.text) Then
    cWhere = cWhere & Tr(cWhere) & " v.DATE <= " & DateSq(xDate2.text)
    aHeader(0) = BetweenString(xDate1.text, xDate2.text)
End If

If Trim(Xcode.text) <> "" Then
    cWhere = cWhere & Tr(cWhere) & " FILE6_20H.CODE = " & MyParn(Xcode.text)
    aHeader(1) = "«·⁄„Ì· : " & xCustName.Caption
End If

If xGroup.MatchedWithList Then
    cWhere = cWhere & Tr(cWhere) & " FILE3_10.[GROUP] = " & MyParn(xGroup.BoundText)
    aHeader(2) = "«·„Ã„Ê⁄… : " & xGroup.text
End If

If xman.MatchedWithList Then
    cWhere = cWhere & Tr(cWhere) & " v.[MAN] = " & MyParn(xman.BoundText)
    aHeader(3) = "«·„‰œÊ» : " & xman.text
End If

If cWhere <> "" Then
    strSql = strSql & " WHERE " & cWhere
End If

strSql = strSql & " GROUP BY  FILE3_10.CODE," & _
         "FILE3_10.[DESCA]"
Dim db As New clsDb
Set grid1.DataSource = db.myRs(strSql)
Set db = Nothing
fixGrd
End Sub
Sub fixGrd()
With grid1
    .RowHeight(0) = 400
    .WordWrap = True
    .ColHidden(0) = True
    .TextMatrix(0, 0) = "„”·”·"
    .TextMatrix(0, 1) = "ﬂÊœ «·⁄„Ì·"
    .TextMatrix(0, 2) = "«·«”„"
    .TextMatrix(0, 3) = "≈Ã„«·Ì «· Õ’Ì·"
        
    .ColWidth(0) = 800
    .ColWidth(1) = 1000
    .ColWidth(2) = 8000
    .ColWidth(3) = 2000
    
    For Row = 1 To .Rows - 1
        .TextMatrix(Row, 0) = Row
    Next
        
    For Col = 3 To .Cols - 1
        .ColDataType(Col) = flexDTDouble
    Next
    
    .SubtotalPosition = flexSTAbove
    For Col = 3 To .Cols - 1
        .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    Next
        
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
        
    
    panel1(0).Caption = IIf(.Rows = 1, "·«  ÊÃœ ”Ã·«  ", "⁄œœ «·”Ã·«  : " & .Rows - 2)
    End With
End Sub
Private Sub myload2()
Dim strSql As String
strSql = "SELECT  v.man," & _
         "FILE6_25.[DESCA]," & _
         "SUM(V.PAID)" & _
         " FROM  vw_inv_collect v" & _
         " LEFT JOIN FILE6_25 ON v.MAN = FILE6_25.CODE"

cWhere = ""
If IsDate(xDate1.text) Then
    cWhere = "v.DATE >= " & DateSq(xDate1.text)
    aHeader(0) = BetweenString(xDate1.text, xDate2.text)
End If

If IsDate(xDate2.text) Then
    cWhere = cWhere & Tr(cWhere) & " v.DATE <= " & DateSq(xDate2.text)
    aHeader(0) = BetweenString(xDate1.text, xDate2.text)
End If

If Trim(Xcode.text) <> "" Then
    cWhere = cWhere & Tr(cWhere) & " FILE6_20H.CODE = " & MyParn(Xcode.text)
    aHeader(1) = "«·⁄„Ì· : " & xCustName.Caption
End If

If xGroup.MatchedWithList Then
    cWhere = cWhere & Tr(cWhere) & " FILE3_10.[GROUP] = " & MyParn(xGroup.BoundText)
    aHeader(2) = "«·„Ã„Ê⁄… : " & xGroup.text
End If

If xman.MatchedWithList Then
    cWhere = cWhere & Tr(cWhere) & " v.[MAN] = " & MyParn(xman.BoundText)
    aHeader(3) = "«·„‰œÊ» : " & xman.text
End If

If cWhere <> "" Then
    strSql = strSql & " WHERE " & cWhere
End If

strSql = strSql & " GROUP BY  v.MAN," & _
         "FILE6_25.[DESCA]"
Dim db As New clsDb
Set grid2.DataSource = db.myRs(strSql)
Set db = Nothing
Fixgrd2
End Sub
Sub Fixgrd2()
With grid2
    .RowHeight(0) = 400
    .WordWrap = True
    .ColHidden(0) = True
    
    .TextMatrix(0, 0) = "„”·”·"
    .TextMatrix(0, 1) = "ﬂÊœ «·„‰œÊ»"
    .TextMatrix(0, 2) = "«·«”„"
    .TextMatrix(0, 3) = "≈Ã„«·Ì «· Õ’Ì·"
    .ColWidth(0) = 800
    .ColWidth(1) = 1000
    .ColWidth(2) = 8000
    .ColWidth(3) = 2000
    
    For Row = 1 To .Rows - 1
        .TextMatrix(Row, 0) = Row
    Next
        
    For Col = 3 To .Cols - 1
        .ColDataType(Col) = flexDTDouble
    Next
    
    .SubtotalPosition = flexSTAbove
    For Col = 3 To .Cols - 1
        .Subtotal flexSTSum, -1, Col, "##,##", &HC0FFC0, vbBlack, True, "«·≈Ã„«·Ï"
    Next
        
    .ExplorerBar = flexExSort
    .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 4
        
    panel1(1).Caption = IIf(grid2.Rows = 1, "·«  ÊÃœ ”Ã·«  ", "⁄œœ «·”Ã·«  : " & grid2.Rows - 2)
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
Set grdInvCollect = Nothing
End Sub
Private Sub grid1_DblClick()
If grid1.Row > 1 Then
    PayLookup
End If
End Sub

Private Sub grid2_DblClick()
If grid2.Row > 1 Then
    PayManLookup
End If
End Sub

Private Sub xCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then CLIENTLOOKUP Me, oSearch
End Sub
Private Sub xcode_LostFocus()
myLostFocus Xcode
xCustName.Caption = ""
If Xcode.text = "" Then Exit Sub
Xcode.text = RetZero(Xcode.text, 4)
Dim db As New clsDb
xCustName.Caption = db.rsField("select desca from FILE3_10 where code = " & MyParn(Xcode.text)) & ""
Set db = Nothing
End Sub
Public Sub myProc(sFlag As String)
If ActiveControl.Name = Xcode.Name Then
    Xcode.text = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    Me.xCustName.Caption = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 1)
    Unload oSearch
'ElseIf sFlag = "PAYMENT" Then
'    'Unload oSearchTotal
End If
End Sub

Private Sub xcode_GotFocus()
myGotFocus Xcode
End Sub
Private Sub xDate1_GotFocus()
myGotFocus xDate1
End Sub
Private Sub xDate1_LostFocus()
myLostFocus xDate1
myValidDate xDate1
End Sub
Private Sub xDate2_GotFocus()
myGotFocus xDate2
End Sub
Private Sub xDate2_LostFocus()
myLostFocus xDate2
myValidDate xDate2
End Sub
Private Sub xDate_due_GotFocus()
myGotFocus xdate_Due
End Sub
Private Sub xDate_due_LostFocus()
myLostFocus xdate_Due
myValidDate xdate_Due
End Sub
Private Sub xGroup_GotFocus()
myGotFocus xGroup
End Sub
Private Sub xgroup_LostFocus()
myLostFocus xGroup
If Not xGroup.MatchedWithList Then xGroup.BoundText = ""
End Sub
Private Sub myPrint(Optional pDevice As String = "", Optional bIgPreview As Boolean = False)
If grid1.Rows < 3 Then Exit Sub
Dim aRow As Variant, aSub As Variant
aSub = AddFlag(Empty, "row", 1)
aSub = AddFlag(aSub, "col", 1)
aSub = AddFlag(aSub, "cols", 2)
aSub = AddFlag(aSub, "text", "≈Ã„«·Ì")
aRow = AddFlag(aRow, aSub)
If SSTab1.Tab = 1 Then
    PrintGrdNew.doprint Me.grid1, 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption & "-" & SSTab1.TabCaption(1), retHeader(aHeader, 0, 2), , False, True, 10, , aRow, Array(1)
Else
    PrintGrdNew.doprint Me.grid2, 0.95, 0, "ÃÊ‰ÌÊ—", Me.Caption & "-" & SSTab1.TabCaption(0), retHeader(aHeader, 0, 2), , False, True, 10, , aRow, Array(1)
End If
If Not bIgPreview Then
    PrintGrdNew.Show 1
Else
    Unload PrintGrdNew
End If
'PrintGrdNew.Show 1
'Unload PrintGrdNew
End Sub
Private Sub PayLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 3)
Set Generalarray(0) = Me

Dim sb As New ChilkatStringBuilder

sb.Append "SELECT v.DESCA ," & _
          " FORMAT(DATE,'yyyy/M/d') AS DATE_PAID," & _
          " FILE6_25.DESCA," & _
          " PAID" & _
          " FROM vw_inv_collect v " & _
          " LEFT JOIN FILE6_25 On v.MAN = FILE6_25.CODE" & _
          " WHERE v.CODE = " & MyParn(grid1.TextMatrix(grid1.Row, 1))

If cWhere <> "" Then
    sb.Append " AND " & cWhere
End If

Generalarray(1) = sb.GetAsString
Generalarray(2) = " ORDER BY DATE,DOC_NO,DOC_PAID"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = " «—ÌŒ «·”œ«œ"
listarray(0, 1) = "(%%DATE_PAID%%)"

GrdArray(0, 0) = "‰Ê⁄ «·”œ«œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1300

GrdArray(2, 0) = "«·„‰œÊ»"
GrdArray(2, 1) = 5000

GrdArray(3, 0) = "«·„»·€"
GrdArray(3, 1) = 1000
GrdArray(3, 3) = True

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchtotal.sFlag = "PAYMENT"
oSearchtotal.Caption = "«” ⁄·«„  ›’Ì·Ì «· Õ’Ì·"
oSearchtotal.Show 1
End Sub
Private Sub PayManLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(3, 3)
Set Generalarray(0) = Me

Dim sb As New ChilkatStringBuilder

sb.Append "SELECT v.DESCA ," & _
          " FORMAT(DATE,'yyyy/M/d') AS DATE_PAID," & _
          " FILE3_10.DESCA," & _
          " PAID" & _
          " FROM vw_inv_collect v " & _
          " INNER JOIN FILE3_10 On v.CODE = FILE3_10.CODE"
If grid2.TextMatrix(grid2.Row, 1) <> "" Then
    sb.Append " WHERE v.MAN = " & MyParn(grid2.TextMatrix(grid2.Row, 1))
Else
    sb.Append " WHERE v.MAN IS NULL"
End If

Generalarray(1) = sb.GetAsString
Generalarray(2) = " ORDER BY DATE,DOC_NO,DOC_PAID"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = " «—ÌŒ «·”œ«œ"
listarray(0, 1) = "(%%DATE_PAID%%)"

GrdArray(0, 0) = "‰Ê⁄ «·”œ«œ"
GrdArray(0, 1) = 2000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 1300

GrdArray(2, 0) = "«·⁄„Ì·"
GrdArray(2, 1) = 5000

GrdArray(3, 0) = "«·„»·€"
GrdArray(3, 1) = 1000
GrdArray(3, 3) = True

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchMan.sFlag = "PAYMENT2"
oSearchMan.Caption = "«” ⁄·«„  ›’Ì·Ì «· Õ’Ì· ··⁄„·«¡"
oSearchMan.Show 1
End Sub


