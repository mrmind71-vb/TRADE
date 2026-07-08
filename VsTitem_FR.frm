VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form VsTitem_FR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "„ «»⁄… «·„ÊœÌ·« "
   ClientHeight    =   11040
   ClientLeft      =   75
   ClientTop       =   450
   ClientWidth     =   15090
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
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   RightToLeft     =   -1  'True
   ScaleHeight     =   11040
   ScaleWidth      =   15090
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame3 
      Caption         =   "’Ê—… «·„ÊœÌ· "
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   7110
      RightToLeft     =   -1  'True
      TabIndex        =   40
      Top             =   45
      Width           =   5460
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   3930
         Left            =   90
         TabIndex        =   41
         Top             =   225
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   6932
         BackColor       =   16777215
         BorderStyle     =   4
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         AutoZoomType    =   3
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "amr123"
         LicenseRegCode  =   "íß“µª∫≥ºΩ∫æ“±æß´¥Ω∏•OANH-FAZOHIZB-EFTP6gI"
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1365
      Left            =   225
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   2880
      Width           =   3885
      Begin VB.CommandButton cmdBarCode 
         Height          =   465
         Left            =   2520
         Picture         =   "VsTitem_FR.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   810
         Width           =   1275
      End
      Begin VB.CommandButton cmd_dem 
         Caption         =   "„ «»⁄… ÿ·»Ì« "
         Height          =   465
         Left            =   1305
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1395
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmd_ret 
         Caption         =   "„ «»⁄… ‘ﬂ«ÊÏ"
         Height          =   465
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1395
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Height          =   465
         Left            =   1305
         Picture         =   "VsTitem_FR.frx":2529
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   315
         Width           =   1185
      End
      Begin VB.CommandButton cmdExit 
         Height          =   465
         Left            =   90
         Picture         =   "VsTitem_FR.frx":4953
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   315
         Width           =   1185
      End
      Begin VB.CommandButton cmdGo 
         Height          =   465
         Left            =   2520
         Picture         =   "VsTitem_FR.frx":6DBF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "⁄—÷"
         Top             =   315
         Width           =   1275
      End
      Begin Threed.SSCommand cmd_excel 
         Height          =   465
         Left            =   45
         TabIndex        =   38
         Top             =   810
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   820
         _Version        =   196610
         BackColor       =   13822956
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "VsTitem_FR.frx":92B1
         Caption         =   " ÕÊÌ· «ﬂ”Ì· "
         Alignment       =   4
         PictureAlignment=   1
      End
      Begin ComctlLib.ProgressBar prog1 
         Height          =   195
         Left            =   90
         TabIndex        =   39
         Top             =   135
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Õ”» »Ì«‰«  «·„ÊœÌ· "
      Height          =   2850
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   45
      Width           =   7170
      Begin VSFlex7Ctl.VSFlexGrid VsTot 
         Height          =   2535
         Left            =   90
         TabIndex        =   36
         Top             =   225
         Width           =   6945
         _cx             =   12250
         _cy             =   4471
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   12648384
         ForeColorSel    =   64
         BackColorBkg    =   12648447
         BackColorAlternate=   14737632
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
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   400
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
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
   Begin VB.Frame Frame1 
      Height          =   2850
      Left            =   12600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   7485
      Begin VB.ComboBox xmosm2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4770
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   2340
         Width           =   1050
      End
      Begin VB.TextBox XOKAZ 
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
         Left            =   135
         MaxLength       =   15
         TabIndex        =   29
         Top             =   855
         Width           =   1545
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   135
         Width           =   2805
         Begin Threed.SSCommand CMD_RATE 
            Height          =   375
            Left            =   1530
            TabIndex        =   27
            Top             =   135
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   661
            _Version        =   196610
            ForeColor       =   64
            Caption         =   "‰”»… „»Ì⁄« "
            ButtonStyle     =   2
         End
         Begin VB.TextBox X2 
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
            Left            =   45
            MaxLength       =   15
            TabIndex        =   26
            Top             =   180
            Width           =   600
         End
         Begin VB.TextBox X1 
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
            Left            =   765
            MaxLength       =   15
            TabIndex        =   25
            Top             =   180
            Width           =   600
         End
      End
      Begin VB.TextBox xBarCode 
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
         Left            =   4770
         MaxLength       =   15
         TabIndex        =   14
         Top             =   1620
         Width           =   1635
      End
      Begin VB.TextBox xModelFact 
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
         Left            =   135
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1575
         Width           =   1545
      End
      Begin VB.TextBox xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1935
         Width           =   3915
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   2970
         TabIndex        =   2
         Top             =   1260
         Width           =   3435
         _ExtentX        =   6059
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
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   2970
         TabIndex        =   3
         Top             =   900
         Width           =   3435
         _ExtentX        =   6059
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
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   4770
         TabIndex        =   4
         Top             =   1980
         Width           =   1635
         _ExtentX        =   2884
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
      Begin MSDataListLib.DataCombo xSupp 
         Height          =   315
         Left            =   2970
         TabIndex        =   12
         Top             =   540
         Width           =   3435
         _ExtentX        =   6059
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
      Begin Threed.SSCheck XRET 
         Height          =   330
         Left            =   2970
         TabIndex        =   28
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   582
         _Version        =   196610
         BackColor       =   12648447
         Caption         =   "√Õ›«¡ «·„— Ã⁄ »«·ﬂ«„·"
      End
      Begin MSDataListLib.DataCombo xsection 
         Height          =   315
         Left            =   135
         TabIndex        =   34
         Top             =   1215
         Width           =   1545
         _ExtentX        =   2725
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
      Begin MSComCtl2.DTPicker XDATE1 
         Height          =   375
         Left            =   4815
         TabIndex        =   37
         Top             =   180
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   177930241
         CurrentDate     =   43098.0832986111
      End
      Begin VB.Label Label2 
         Caption         =   "„Ê”„ œ«Œ·Ï"
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
         Index           =   6
         Left            =   5895
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   2385
         Width           =   1140
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "«·ﬁ”„ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1245
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "‰”»… «Êﬂ«“ÊÌ‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   900
         Width           =   1185
      End
      Begin VB.Label xModelNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4140
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2610
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "»«—ﬂÊœ :"
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
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1725
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   630
         Width           =   600
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "—ﬁ„ „ÊœÌ· :"
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
         Left            =   1755
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1620
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«·»Ì«‰ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Index           =   4
         Left            =   4095
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2025
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "„Ã„Ê⁄… :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1395
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "«·„’‰⁄ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   975
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "«·„Ê”„ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   2025
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Õ Ï  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   870
      End
   End
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   5985
      Top             =   45
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
   Begin MSAdodcLib.Adodc data3 
      Height          =   330
      Left            =   5685
      Top             =   45
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
      Left            =   5385
      Top             =   45
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   4800
      Top             =   45
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   4500
      Top             =   45
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
   Begin MSAdodcLib.Adodc data6 
      Height          =   330
      Left            =   5100
      Top             =   45
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
   Begin VSFlex7Ctl.VSFlexGrid grid2 
      Height          =   1275
      Left            =   12645
      TabIndex        =   17
      Top             =   2970
      Visible         =   0   'False
      Width           =   7440
      _cx             =   13123
      _cy             =   2249
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      AllowSelection  =   -1  'True
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
   Begin VSFlex7Ctl.VSFlexGrid GRID1 
      Height          =   5820
      Left            =   225
      TabIndex        =   23
      Top             =   4275
      Width           =   19860
      _cx             =   35031
      _cy             =   10266
      _ConvInfo       =   1
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      BackColorSel    =   16777152
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   18
      FixedRows       =   2
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
   Begin MSAdodcLib.Adodc data20 
      Height          =   330
      Left            =   4500
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
   Begin MSAdodcLib.Adodc data12 
      Height          =   330
      Left            =   -675
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
   Begin MSAdodcLib.Adodc data8 
      Height          =   330
      Left            =   0
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
Attribute VB_Name = "VsTitem_FR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oSearch As New Search3
Public cMyModel As String
Dim con As New ADODB.Connection
Dim StoreTable As New ADODB.Recordset, clist1 As String
Private Sub CMD_PRINT_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    cHead1 = "»Ì«‰ ≈Ã„«·Ï  ⁄«„·«  «·„ÊœÌ·«  "
    cHead2 = " Õ Ï   «—ÌŒ " & Format(xDate1.Value, "DD-MM-YYYY")
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 0.75, -2, cHead1, cHead2, , False, True, 6
    PrintGrd.Show 1
End Sub

Private Sub cmd_dem_Click()
grditemfrm1.Show
End Sub

Private Sub cmd_excel_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
End Sub

Private Sub cmd_nopict_Click()
Dim fs As New FileSystemObject
With grid1
    For nRow = 2 To .Rows - 1
        If fs.FileExists(retFile(.TextMatrix(nRow, 0))) Then .RowHidden(nRow) = True
    Next nRow
End With
End Sub

Private Sub CMD_RATE_Click()
With grid1
    For nRow = 2 To .Rows - 1
        .RowHidden(nRow) = False
        If X1.text <> "" Then
            If Val(.TextMatrix(nRow, 10)) < Val(X1.text) Then .RowHidden(nRow) = True
        End If
    
        If X2.text <> "" Then
            If Val(.TextMatrix(nRow, 10)) > Val(X2.text) Then .RowHidden(nRow) = True
        End If
    Next nRow
End With
End Sub


Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdGo_Click()
    myload
    MyLOadTot
End Sub
Private Sub cmdBarCode_Click()
con.BeginTrans
On Error GoTo myerror
If xModelNo.Caption = "" Then Exit Sub
cString = "INSERT INTO ADDPRINT(ITEM,BRANCH,QUANT,ISPRINT)"
cString = cString & _
          " SELECT FILE1_10.ITEM," & addstring(cBranch) & ",SUM([IN]-[OUT]),1 FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE FILE1_10.MODELNO = " & MyParn(xModelNo.Caption)
cString = cString & " GROUP BY FILE1_10.ITEM "
con.Execute cString
con.CommitTrans
Inform " „ «· ÕÊÌ· »‰Ã«Õ"
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub cmdPrint_Click()
    Dim cHead1 As String
    Dim cHead2 As String
    grid1.ColHidden(12) = True
    grid1.ColHidden(13) = True
    grid1.ColHidden(15) = True
    cHead1 = "»Ì«‰ ≈Ã„«·Ï  ⁄«„·«  «·„ÊœÌ·«  " & Format(xDate1.Value, "DD-MM-YYYY")
    cHead2 = xSupp.text & "  " & xMosm.text
    
    Load PrintGrd
    PrintGrd.doprint Me.grid1, 1.25, 2, cHead1, cHead2, , False, True, 11
    PrintGrd.Show 1
    grid1.ColHidden(12) = False
    grid1.ColHidden(13) = False
    grid1.ColHidden(15) = False

End Sub
Private Sub Form_Load()
    openCon con
    StoreTable.Open "FILE0_40", con, adOpenStatic, adLockReadOnly, adCmdTable
'    XDATE1.VALUE = Format(Date, "dd-mm-yyyy")
    
    XMOSM2.AddItem " "
    XMOSM2.AddItem "S"
    XMOSM2.AddItem "W"
    XMOSM2.AddItem "M"
    
    clist1 = StrList("SELECT CODE , DESCA FROM FILE1_10SC ", con)
    Dim ImgXTmp As ImgX
    Set ImgXTmp = New ImgX
'   ImgXTmp.Import.FromMemoryFile LoadResData("BACKGROUND", "CUSTOM")
'   Set ImgX1.BackgroundImage = ImgXTmp.image
    Set ImgXTmp = Nothing
    
    DATA12.ConnectionString = strCon
    DATA12.RecordSource = "SELECT * FROM file1_10sc order by desca "
    Set xSection.RowSource = DATA12
    xSection.ListField = "DESCA"
    xSection.BoundColumn = "code"
    
    DATA2.ConnectionString = strCon
    DATA2.RecordSource = "Select Code,DescA From File4_10 order by Desca"
    Set xSupp.RowSource = DATA2
    xSupp.ListField = "Desca"
    xSupp.BoundColumn = "Code"
    
    data3.ConnectionString = strCon
    data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    
    data4.ConnectionString = strCon
    data4.RecordSource = "Select mosm ,descA From mosm ORDER BY date DESC "
    Set xMosm.RowSource = data4
    xMosm.ListField = "Desca"
    xMosm.BoundColumn = "MOSM"
    xMosm.BoundText = cPMosm
    
    
    data5.ConnectionString = strCon
    data5.RecordSource = "Select code,desca From fact ORDER BY code "
    Set xFact.RowSource = data5
    xFact.ListField = "Desca"
    xFact.BoundColumn = "Code"
    
    Set grid1.DataSource = DATA6
    DATA6.ConnectionString = strCon
    
    
    Set GRID2.DataSource = data20
    data20.ConnectionString = strCon
    
    Frame2.Visible = bOpt10
    
    grid1.Rows = 2
    
    GRID2.Rows = 1
    FixTot
    Fixgrd
    Fixgrd2
    Frame2.Visible = bopt2
End Sub
Private Sub myload()
    Dim cF11 As String
    cF11 = "(SELECT COUNT(DISTINCT DOC_ID) FROM Items_Purch INNER JOIN FILE1_10 AS FILE1_10_2 ON Items_Purch.ITEM = FILE1_10_2.ITEM WHERE FILE1_10_2.MODELNO = FILE1_10.MODELNO )"
    Dim cWhere As String

    cField = myiif("type = 'z' ", "([out] - [in]) ") & " AS t_comp"
    
    cField = cField & "," & _
            myiif("type = '2' ", "([IN] - [OUT]) ") & " AS t_purch "
    
    cField = cField & "," & _
             myiif("type = '7' ", "([out] - [in]) ") & " AS t_ret"

    cField = cField & "," & _
             myiif("type = '6' ", "([out] - [in]) ") & " AS t_sal"



    Dim aRet As Variant
    aRet = GetRows("Select code,Desca from file0_40", con)
    For i = 0 To UBound(aRet)
        cField3 = cField3 & turn(cField3, ",") & _
                   myiif("FILE1_11.STORE = " & MyParn(retFlag(aRet(i), "CODE")), "([IN]-[OUT]) ") & " AS [" & retFlag(aRet(i), "DESCA") & "]"
    Next
    
    cField2 = "DATEDIFF(dd , " & myiif2("type = '2' ", "FILE1_11.date", "MIN") & " , GETDATE()) AS DATEDIFF"
    
    '                   0                   1               2           3                   4
    cString = " SELECT file1_10.ModelNo, fact.desca , file1_10.mosm, file1_10.MODELFACT0, file1_10.desca," & _
            cField & _
            ", sum([in] - out) as [«·—’Ìœ] ,  " & _
            " ' ' as rate , min(file1_10.price) , min(file1_10.costitem )  , sum((file1_11.[in] - file1_11.out) * file1_10.costitem  ) as costbal ," & cField2 & _
            ", MAX(file1_10.cost ) , max(OKAZ)  , MIN(PRICE2), " & cF11 & _
            " FROM FILE1_10 INNER JOIN FACT ON FILE1_10.FACT = FACT.CODE INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM   INNER JOIN MOSM ON FILE1_10.MOSM = MOSM.MOSM WHERE file1_10.ISNOITEM  = 0 "
    If xMosm.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If XMOSM2.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MOSM2 = " & MyParn(XMOSM2.text)
    If xModelFact.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.MODELFACT0 = " & MyParn(xModelFact.text)
    If xbarcode.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.item = " & Val(xbarcode.text)
    If xSupp.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xGroup.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If xFact.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xSection.BoundText <> "" Then cStr2 = cStr2 & " AND FILE1_10.[section] = " & MyParn(xSection.BoundText)
    If XOKAZ.text <> "" Then cStr2 = cStr2 & " AND FILE1_10.[OKAZ] = " & Val(XOKAZ.text)
    If Trim(xDesca.text) <> "" Then cString = cString & turn(cString) & MyParnAnd(xDesca.text, "file1_10.desca")
    
    If IsDate(xDate1.Value) Then cStr2 = cStr2 & " AND FILE1_11.[DATE] <= " & DateSq(xDate1.Value)
    cString = cString & cStr2 & " GROUP BY file1_10.ModelNo, FACT.desca , file1_10.mosm, file1_10.MODELFACT0, file1_10.desca,FILE1_10.FACT   "

    DATA6.RecordSource = cString
    DATA6.Refresh
    Fixgrd
    If xMosm.BoundText <> "" Or xSupp.BoundText <> "" Or xFact.BoundText <> "" Or xGroup.BoundText <> "" Then
        CalcTot
    End If

End Sub
Sub Fixgrd()
    With grid1
    .ExplorerBar = flexExSortShow
    .FixedRows = 2
    .RowHidden(1) = True
    .RowHeight(0) = 700
    .WordWrap = True
    .ColHidden(0) = True
    .RowHidden(1) = True
    .FrozenCols = 4
    .Cols = 19
    .TextMatrix(0, 0) = "„ÊœÌ·"
    .TextMatrix(0, 1) = "«·„’‰⁄"
    .TextMatrix(0, 2) = "„Ê”„"
    .TextMatrix(0, 3) = "„ÊœÌ· "
    .TextMatrix(0, 4) = "«·’‰›"

    .TextMatrix(0, 5) = "Ã—œ"
    .TextMatrix(0, 6) = "„‘ —Ì« "
    .TextMatrix(0, 7) = "„— Ã⁄« "
    .TextMatrix(0, 8) = "„»Ì⁄« "
    .TextMatrix(0, 9) = "—’Ìœ"

    .TextMatrix(0, 10) = "‰”»… »Ì⁄"
    .TextMatrix(0, 11) = " ”⁄— »Ì⁄"
    .TextMatrix(0, 12) = "”⁄—  ﬂ·›…"
    .TextMatrix(0, 13) = " ﬁÌ„…  ﬂ·›… «·—’Ìœ"
'
    .TextMatrix(0, 14) = "⁄„— „ÊœÌ·"
    .TextMatrix(0, 15) = "”⁄— «·„’‰⁄"
    .TextMatrix(0, 16) = "‰”»… —»Õ"
'    .TextMatrix(0, 17) = "”⁄— √Êﬂ«“ÊÌ‰"

    .TextMatrix(0, 18) = " ﬂ—«—"

    .ColHidden(12) = Not bOpt10
    .ColHidden(13) = Not bOpt10
    .ColHidden(15) = Not bOpt10
    .ColHidden(16) = Not bOpt10
    
    .ColWidth(0) = 0
    .ColWidth(1) = 1200
    .ColWidth(2) = 500
    .ColWidth(3) = 1500
    .ColWidth(4) = 2000
   
    .ColWidth(5) = 700
    .ColWidth(6) = 700
    .ColWidth(7) = 700
    .ColWidth(8) = 700
    .ColWidth(9) = 700
    
    .ColWidth(10) = 800
    .ColWidth(11) = 800
    .ColWidth(12) = 800
    .ColWidth(13) = 1000
    
    .ColWidth(14) = 850
    .ColWidth(15) = 800
    .ColWidth(16) = 800
    .ColWidth(17) = 0
    .ColWidth(18) = 500
    .Cell(flexcpAlignment, 0, 0, 1, .Cols - 1) = 4
    
    .ColDataType(0) = flexDTString
    .ColDataType(1) = flexDTString
    .ColDataType(2) = flexDTString
    .ColDataType(3) = flexDTString
    .ColDataType(4) = flexDTString
    .ColDataType(15) = flexDTString
    
    
'    .ColHidden(17) = Not bopt2
    
    For nCol = 5 To .Cols - 2
        .ColDataType(nCol) = flexDTDouble
    Next nCol
    .ColDataType(18) = flexDTString
    .Editable = flexEDNone
    .ColDataType(13) = flexDTDouble
    .ColFormat(13) = "#0.00"
    .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = 4
    .Subtotal flexSTClear
    .WordWrap = True
   
    If .Rows > 2 Then
        For i = 2 To .Rows - 1
            If (Val(.TextMatrix(i, 6)) = Val(.TextMatrix(i, 7))) Then
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = &HC0C0FF
            End If
            nRat = 0
            If (Val(.TextMatrix(i, 6)) - Val(.TextMatrix(i, 7))) > 0 Then nRat = Val(.TextMatrix(i, 8)) / (Val(.TextMatrix(i, 6)) - Val(.TextMatrix(i, 7))) * 100
            .TextMatrix(i, 10) = Format(nRat, "##0.00")
        
            nRate = 0
            If Val(.TextMatrix(i, 12)) <> 0 Then nRate = Round((Val(.TextMatrix(i, 11)) - Val(.TextMatrix(i, 12))) / Val(.TextMatrix(i, 12)) * 100, 2)
            .TextMatrix(i, 16) = nRate
        Next i
    End If
    
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 5, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 6, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 7, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 8, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 9, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    .Subtotal flexSTSum, -1, 13, "##0", , RGB(255, 0, 0), True, "≈Ã„«·Ï"
    
    If .Rows > 2 Then
        nRat = 0
        If (Val(.TextMatrix(.Rows - 1, 6)) - Val(.TextMatrix(.Rows - 1, 7))) > 0 Then nRat = Val(.TextMatrix(.Rows - 1, 8)) / (Val(.TextMatrix(.Rows - 1, 6)) - Val(.TextMatrix(.Rows - 1, 7))) * 100
        .TextMatrix(.Rows - 1, 10) = Format(nRat, "##0.00")
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = 7
    End If
End With
End Sub
Sub Fixgrd2()
    With GRID2
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .Cols = 8
    .TextMatrix(0, 0) = "ﬂÊœ"
    .TextMatrix(0, 1) = "«·„Ê—œ"
    .TextMatrix(0, 2) = "„‘ —Ì« "
    .TextMatrix(0, 3) = "„— Ã⁄« "
    .TextMatrix(0, 4) = "’«›Ï"
    .TextMatrix(0, 5) = "„»Ì⁄« "
    .TextMatrix(0, 6) = "—’Ìœ"
    .TextMatrix(0, 7) = "‰”»… »Ì⁄"
    .ColWidth(0) = 1000
    .ColWidth(1) = 3000
    .ColWidth(2) = 1000
    .ColWidth(3) = 1000
    .ColWidth(4) = 1000
    .ColWidth(5) = 1000
    .ColWidth(6) = 1000
    .ColWidth(7) = 1000
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 4)) <> 0 Then .TextMatrix(nRow, 7) = Format(Val(.TextMatrix(nRow, 5)) / Val(.TextMatrix(nRow, 4)), "#0.00%")
        If Val(.TextMatrix(nRow, 4)) = Val(.TextMatrix(nRow, 5)) Then
            If Val(.TextMatrix(nRow, 6)) = 0 Then .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = &HC0C0FF
        End If
    Next nRow
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub


Private Sub grid1_DblClick()
On Error GoTo myerror
    If grid1.Row <= 1 Then Exit Sub
    Dim cString As String
    If grid1.Col = 4 And bopt2 Then
        ITEMS.sModel = GetDesca("Select Model FROM FILE1_10H WHERE CODE = " & MyParn(GRID2.TextMatrix(GRID2.Row, 0)) & " AND MODELNO = " & MyParn(xModelNo.Caption), con)
        ITEMS.Show 1
    End If
    If grid1.Col <= 3 And bopt2 Then
        Load ModelMove
        ModelMove.XMODEL.text = xModelNo.Caption
        ModelMove.Show
    ElseIf grid1.Col = 6 Or grid1.Col = 7 Then
        InvModel.pModelNo = xModelNo.Caption
        InvModel.Show
    ElseIf grid1.Col = 8 And Val(grid1.TextMatrix(grid1.Row, 8)) <> 0 Then
        cString = "SELECT SUM(QUANT) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL" & _
                  " FROM FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM" & _
                  " INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO"
        If IsDate(xDate1.Value) Then cString = cString & turn(cString) & "FILE6_20H.DATE <= " & DateSq(xDate1.Value)
        cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
        cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL"
        grdModelfrm.pString = cString
        grdModelfrm.pFieldValue = "sumofQuant"
        grdModelfrm.pModelNo = grid1.TextMatrix(grid1.Row, 0)
        grdModelfrm.Show 1
    ElseIf grid1.Col = 9 Then
        cString = "SELECT SUM([IN]-[OUT]) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL" & _
                  " FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM"
        If IsDate(xDate1.Value) Then cString = cString & turn(cString) & "FILE1_11.DATE <= " & DateSq(xDate1.Value)
        cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
        cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL"
        grdModelfrm.pString = cString
        grdModelfrm.pFieldValue = "sumofQuant"
        grdModelfrm.pModelNo = grid1.TextMatrix(grid1.Row, 0)
        grdModelfrm.Show 1
'    ElseIf Grid1.Col = 9 Then
'        VsModelNoLook.sModelNo = Grid1.TextMatrix(Grid1.Row, 0)
'        VsModelNoLook.SFact = Grid1.TextMatrix(Grid1.Row, 1)
'        VsModelNoLook.SModelFact = Grid1.TextMatrix(Grid1.Row, 3)
'        VsModelNoLook.SMosm = Grid1.TextMatrix(Grid1.Row, 2)
'        VsModelNoLook.sdesca = Grid1.TextMatrix(Grid1.Row, 4)
'        VsModelNoLook.Show 1
    ElseIf grid1.Col = 11 Or grid1.Col = 12 Then
        Ordersfrm.sModelNo = xModelNo.Caption
        Ordersfrm.Show
    End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub grid1_EnterCell()
'    If grid1.Rows > 2 Then
'        If grid1.TextMatrix(2, 16) <> "" Then Exit Sub
'    End If
    xModelNo.Caption = grid1.TextMatrix(grid1.Row, 0)
'    If xModelNo.Caption <> "" Then myload2
    LoadPhoto grid1.TextMatrix(grid1.Row, 0)
    grid1.Editable = flexEDNone
    grid1.ToolTipText = GetDesca("SELECT DESCA FROM FACT WHERE CODE = " & MyParn(grid1.TextMatrix(grid1.Row, 1)), con)
    If grid1.Col = 16 And bopt2 Then
        grid1.Editable = flexEDKbdMouse
    End If

End Sub
Sub CalcTot()
    Dim cWher As String
    If xMosm.BoundText <> "" Then cWher = cWher & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cWher = cWher & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xFact.BoundText <> "" Then cWher = cWher & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If xSection.BoundText <> "" Then cWher = cWher & " AND FILE1_10.[section] = " & MyParn(xSection.BoundText)
    If xGroup.BoundText <> "" Then cWher = cWher & " AND FILE1_10.[GROUP] = " & MyParn(xGroup.BoundText)
    If IsDate(xDate1.Value) Then cWher = cWher & " AND [DATE] <= " & DateSq(xDate1.Value)
    Dim cF1 As String
    cWher22 = ""
    If xMosm.BoundText <> "" Then cWher22 = cWher22 & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cWher22 = cWher22 & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If IsDate(xDate1.Value) Then cWher22 = cWher22 & " AND [DATE] <= " & DateSq(xDate1.Value)
    
    
    cWher2 = ""
    If xMosm.BoundText <> "" Then cWher2 = cWher2 & " AND MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cWher2 = cWher2 & " AND CODE = " & MyParn(xSupp.BoundText)
    If IsDate(xDate1.Value) Then cWher2 = cWher2 & " AND [DATE] <= " & DateSq(xDate1.Value)
    cString = " SELECT COUNT(DISTINCT modelno) " & _
            " , ( SELECT SUM(([IN]-[OUT]) * FILE1_11.COST * ((100-FILE1_11.DISCOUNT)/100) ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM   WHERE (TYPE = '2' OR TYPE = '7' ) " & cWher22 & " )  " & _
            " , ( SELECT SUM(FILE6_20.TOTAL ) FROM FILE6_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_20.ITEM INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20.DOC_NO IS NOT NULL " & cWher & " )  " & _
            " , ( SELECT SUM(FILE6_20.QUANT * FILE6_20.COST ) FROM FILE6_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_20.ITEM INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO  WHERE FILE6_20.DOC_NO IS NOT NULL " & cWher & " ) " & _
            " , ( SELECT SUM((FILE1_11.[IN] -FILE1_11.OUT) * FILE1_10.costitem ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE FILE1_10.ITEM IS NOT NULL " & cWher & " ) " & _
            " , ( SELECT SUM(FILE7_20h.discount ) FROM FILE7_20h WHERE FILE7_20h.discount <> 0 " & cWher2 & " ) " & _
            " , ( SELECT SUM(FILE7_10h.discount ) FROM FILE7_10h WHERE FILE7_10h.discount <> 0 " & cWher2 & " ) "
    cWher = ""
    If xMosm.BoundText <> "" Then cWher = cWher & " AND MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cWher = cWher & " AND CODE = " & MyParn(xSupp.BoundText)
    If xFact.BoundText <> "" Then cWher = cWher & " AND FACT = " & MyParn(xFact.BoundText)
    If xSection.BoundText <> "" Then cWher = cWher & " AND [sec] = " & MyParn(xSection.BoundText)
    If xGroup.BoundText <> "" Then cWher = cWher & " AND [GROUP] = " & MyParn(xGroup.BoundText)
            
'    cString = cString & " FROM  FILE1_10  AS FILE1_10_0 WHERE FILE1_10_0.ITEM IS NOT NULL " & cWher
'    aret = aGetDesca(cString,CON)
'    If UBound(aret) > 0 Then
'        XCOUNT.Caption = Format(aret(1), "#0")
'        xTPurch.Caption = Format(aret(2), "#0")
'        xTSalCOST.Caption = Format(aret(4), "#0")
'        XTSAL.Caption = Format(aret(3), "#0")
'        xTBal.Caption = Format(aret(5), "#0")
'        xtdisc.Caption = Format(Val(aret(6) & "") - Val(aret(7) & ""), "#0")
'        If Val(xTPurch.Caption) <> 0 Then xRate.Caption = Format(Val(xTSalCOST.Caption) / Val(xTPurch.Caption) * 100, "#0.00")
'    End If
    If xSupp.BoundText <> "" Then
'        xbalcode.Caption = Round(BAL_SUPPL(xSupp.BoundText, XMOSM.BoundText), "#0.00")
'        If IsDate(XDATE1.VALUE) Then
'            xbalcode.Caption = Round(Val(GetDesca("SELECT SUM( SAL - PAY ) FROM FILE4_11 WHERE CODE = " & MyParn(xSupp.BoundText) & " AND MOSM = " & MyParn(xMosm.BoundText) & " AND DATE <= " & DateSq(XDATE1.VALUE)) & ""), 2)
'        Else
'            xbalcode.Caption = Round(Val(GetDesca("SELECT SUM( SAL - PAY ) FROM FILE4_11 WHERE CODE = " & MyParn(xSupp.BoundText) & " AND MOSM = " & MyParn(xMosm.BoundText)) & ""), 2)
'        End If
    End If
End Sub

Private Sub Grid2_dblClick()
    If bopt2 And GRID2.Col < 6 Then
        ITEMS.sModel = GetDesca("Select Model FROM FILE1_10H WHERE CODE = " & MyParn(GRID2.TextMatrix(GRID2.Row, 0)) & " AND MODELNO = " & MyParn(xModelNo.Caption), con)
        ITEMS.Show 1
    End If
    If GRID2.Col = 6 Then
        cString = "SELECT SUM([IN]-[OUT]) AS sumOfQuant,FILE1_10.COLOR,FILE1_10.SCAL" & _
                  " FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM where file1_10.code = " & MyParn(GRID2.TextMatrix(GRID2.Row, 0))
        If IsDate(xDate1.Value) Then cString = cString & turn(cString) & "FILE1_11.DATE <= " & DateSq(xDate1.Value)
        cString = cString & turn(cString) & " FILE1_10.modelno = " & MyParn(grid1.TextMatrix(grid1.Row, 0))
        cString = cString & " GROUP BY FILE1_10.COLOR,FILE1_10.SCAL"
        grdModelfrm.pString = cString
        grdModelfrm.pFieldValue = "sumofQuant"
        grdModelfrm.pModelNo = grid1.TextMatrix(grid1.Row, 0)
        grdModelfrm.Show 1
    End If
End Sub

Private Sub ImgX1_Click()
If validPhoto(retFile(grid1.TextMatrix(grid1.Row, 0))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(grid1.TextMatrix(grid1.Row, 0))
    oMaxPhoto.cDesca = grid1.TextMatrix(grid1.Row, 4)
    oMaxPhoto.Show 1
End If
End Sub
Sub MYLOAD2()
    cField2 = myiif("type = '2' ", "[IN]  ") & " AS t_purch, "
    cField3 = myiif("type = '7' ", "[OUT] ") & " AS t_ret  , "
    cField4 = myiif("type = '2' OR type = '7'", "([IN] - [OUT]) ") & " AS t_NET   , "
    cField5 = myiif("type = '6' ", "[out]  ") & " AS t_sal  , "
    
    '                   0                   1               2           3                   4
    cString = " SELECT file4_10.CODE , FILE4_10.DESCA , " & _
            cField2 & cField3 & cField4 & cField5 & _
            " sum([in] - out) as bal   " & _
            " FROM   FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE WHERE FILE1_10.MODELNO = " & MyParn(xModelNo.Caption)
    cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA  "

    data20.RecordSource = cString
    data20.Refresh
    Fixgrd2
End Sub
Sub myload20()
    Dim cField1 As String, cField2 As String, cField3 As String, cField4 As String, cField As String
    cField1 = myiif("type = '2' ", "[IN]  ") & " AS t_purch"
    cField2 = myiif("type = '7' ", "[OUT] ") & " AS t_ret"
    cField3 = myiif("type = '2' OR type = '7'", "([IN] - [OUT]) ") & " AS t_NET"
    cField4 = myiif("type = '6' ", "[out]  ") & " AS t_sal "
    cField = cField1 & "," & cField2 & "," & cField3 & "," & cField4
    '                   0                   1               2           3                   4
    cString = " SELECT file4_10.CODE , FILE4_10.DESCA" & turn(cField, ",") & cField & _
            " FROM  (FILE1_11 INNER  JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM) INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE "
    cString = cString & turn(cString) & " FILE1_10.MODELNO = " & MyParn(xModelNo.Caption)
    cString = cString & " GROUP BY FILE4_10.CODE, FILE4_10.DESCA  "

    data20.RecordSource = cString
    data20.Refresh
    Fixgrd2
End Sub
Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
ImgX1.Images.Clear
ImgX1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub

Private Sub VsTot_DBLClick()
    With VsTot
        supMovefrm.XCODE.text = xSupp.BoundText
        supMovefrm.xMosm.Caption = xMosm.BoundText
        supMovefrm.Show
    End With
End Sub

Private Sub xFact_Validate(Cancel As Boolean)
If Not xFact.MatchedWithList Then
    cFact = xFact.BoundText
    xFact.BoundText = cFact
    If xFact.MatchedWithList Then Exit Sub
End If

End Sub

Private Sub XRET_Click(Value As Integer)
With grid1
    If XRET.Value Then
        For nRow = 2 To .Rows - 1
            If Val(.TextMatrix(nRow, 6)) = Val(.TextMatrix(nRow, 7)) Then .RowHidden(nRow) = True
        Next nRow
    Else
        For nRow = 2 To .Rows - 1
            .RowHidden(nRow) = False
        Next nRow
    End If
End With
End Sub

Private Sub xSection_LostFocus()
    data3.ConnectionString = strCon
    If xSection.BoundText = "" Then
        data3.RecordSource = "Select Code,DescA From File1_50 ORDER BY DESCA"
    Else
        data3.RecordSource = "Select Code,DescA From File1_50 where [group] = " & Val(xSection.BoundText) & " ORDER BY DESCA"
    End If
    Set xGroup.RowSource = data3
    xGroup.ListField = "Desca"
    xGroup.BoundColumn = "Code"
    data3.Refresh

End Sub

Private Sub xSupp_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then SuppLookupAll Me, oSearch
End Sub
Sub myProc()
If ActiveControl.Name = xSupp.Name Then
    xSupp.BoundText = oSearch.grid1.TextMatrix(oSearch.grid1.Row, 0)
    oSearch.Hide
End If
If ActiveControl.Name = xFact.Name Then
    xFact.BoundText = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
If ActiveControl.Name = xdoc_no.Name Then
    xdoc_no.text = Search3.grid1.TextMatrix(Search3.grid1.Row, 0)
    Unload Search3
End If
End Sub


Private Sub xFACT_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)
Set Generalarray(0) = Me
Generalarray(1) = "Select code ,DescA From FACT"
Generalarray(2) = "Order by DESCA"
Generalarray(3) = 5000
Generalarray(5) = False

listarray(0, 0) = "«·Ê’›"
listarray(0, 1) = "(%%DESCA%%)"

GrdArray(0, 0) = "«·ﬂÊœ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«·Ê’›"
GrdArray(1, 1) = 6000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "≈” ⁄·«„ "
Search3.Show 1

End If
End Sub
Private Sub xsupp_Validate(Cancel As Boolean)
If xSupp.BoundText = "" Then Exit Sub
If Not xSupp.MatchedWithList Then
    cFact = GetDesca("SELECT CODE FROM FILE4_10 WHERE SUBCODE = " & MyParn(xSupp.BoundText), con)
    xSupp.BoundText = cFact
    If xSupp.MatchedWithList Then Exit Sub
End If
End Sub
Sub DoFixPrice2(cModel, nOkaz, nRow)
    Dim SubItemTable As New ADODB.Recordset
    Dim nRem As Double, nPrice As Double
    Dim nFPrice As Double
    cStr1 = " SELECT * FROM FILE1_10  WHERE MODELNO = " & MyParn(cModel)
    SubItemTable.Open cStr1, con, adOpenForwardOnly, adLockPessimistic, adCmdText
'   SubItemTable.Open cStr1, con, adOpenForwardOnly, adLockReadOnly, adCmdText
    
    SubItemTable.MoveFirst
        Do While Not SubItemTable.EOF
            If nOkaz > 0 Then
                nPrice = Val(SubItemTable!price & "") * (100 - nOkaz) / 100
                nRem = nPrice - Int(nPrice)
                Select Case nRem
                    Case Is = 0
                        nPrice = Int(nPrice)
                    Case Is <= 0.5
                        nPrice = Int(nPrice) + 0.5
                    Case Is < 1
                        nPrice = Int(nPrice) + 1
                End Select
                If nPrice > SubItemTable!price Then nPrice = SubItemTable!price
            Else
                nPrice = SubItemTable!price
            End If
            If nFPrice = 0 Then nFPrice = nPrice
            SubItemTable!PRICE2 = nPrice
            SubItemTable!okaz = nOkaz
            SubItemTable.Update
'            con.Execute " UPDATE FILE1_10 SET PRICE2 = " & Val(nPrice) & " WHERE ITEM = " & MyParn(SubItemTable!Item)
            SubItemTable.MoveNext
        Loop
'       con.Execute " UPDATE FILE1_10  SET OKAZ          = " & Val(nOkaz) & " WHERE MODELNO = " & MyParn(cModel)
'        grid1.TextMatrix(nRow, 17) = Format(nFPrice, "#0.00")
'        GRID1.TextMatrix(NROW, 16) = Format(nOkaz, "#0")
End Sub
Sub FixTot()
    With VsTot
    .Rows = 5
    .Cols = 6
    .ColWidth(0) = 1300
    .ColWidth(1) = 1000
    .ColWidth(2) = 1300
    .ColWidth(3) = 1000
    .ColWidth(4) = 1300
    .ColWidth(5) = 1000
    
    .FixedCols = 0
    .FixedRows = 0
    
    .TextMatrix(0, 0) = "⁄œœ „ÊœÌ·« "
    .TextMatrix(0, 2) = "‰”»… „»Ì⁄« "
    
    .TextMatrix(1, 0) = "ﬁÌ„… „‘ —Ì« "
    .TextMatrix(1, 2) = "ﬁÌ„… „— Ã⁄« "
    .TextMatrix(1, 4) = "’«›Ï „‘ —Ì« "
    
    .TextMatrix(2, 0) = "ﬁÌ„… „»Ì⁄« "
    .TextMatrix(2, 2) = " ﬂ·›… „»Ì⁄« "
    .TextMatrix(2, 4) = "—»Õ   „»Ì⁄« "
    
    .TextMatrix(3, 0) = "”œ«œ „Ê—œÌ‰ "
    .TextMatrix(3, 2) = "Œ’„  „Ê—œÌ‰ "
    
    .TextMatrix(4, 0) = "—’Ìœ «·Õ”«»"
    .TextMatrix(4, 2) = " ﬁÌ„ «·»÷«⁄…"
    .TextMatrix(4, 4) = "œ›⁄… „” Õﬁ…"
    
    End With
End Sub
Sub MyLOadTot()
    On Error Resume Next
    Dim nBalSupl  As Double
    With VsTot
    For nRow = 0 To .Rows - 1
        .TextMatrix(nRow, 1) = ""
        .TextMatrix(nRow, 3) = ""
        .TextMatrix(nRow, 5) = ""
    Next nRow
    
    Dim cWher As String
    Dim cF1 As String
    If xMosm.BoundText <> "" Then cWher = cWher & " AND FILE1_10.MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cWher = cWher & " AND FILE1_10.CODE = " & MyParn(xSupp.BoundText)
    If xFact.BoundText <> "" Then cWher = cWher & " AND FILE1_10.FACT = " & MyParn(xFact.BoundText)
    If IsDate(xDate1.Value) Then cWher = cWher & " AND [DATE] <= " & DateSq(xDate1.Value)
    
    cWher22 = ""
    If xMosm.BoundText <> "" Then cWher22 = cWher22 & " AND MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cWher22 = cWher22 & " AND CODE = " & MyParn(xSupp.BoundText)
    If xFact.BoundText <> "" Then cWher22 = cWher22 & " AND FACT = " & MyParn(xFact.BoundText)
    If IsDate(xDate1.Value) Then cWher22 = cWher22 & " AND [DATE] <= " & DateSq(xDate1.Value)
    
    
    cWher2 = ""
    If xMosm.BoundText <> "" Then cWher2 = cWher2 & " AND MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cWher2 = cWher2 & " AND CODE = " & MyParn(xSupp.BoundText)
'   If xFact.BoundText <> "" Then cWher2 = cWher2 & " AND FACT = " & MyParn(xFact.BoundText)
    If IsDate(xDate1.Value) Then cWher2 = cWher2 & " AND [DATE] <= " & DateSq(xDate1.Value)

'    cString = " SELECT COUNT(DISTINCT modelno) " & _
'            " , ( SELECT SUM(SAL)               FROM FILE4_11 WHERE TYPE = '4'  " & cWher22 & " )  " & _
'            " , ( SELECT SUM(PAY)               FROM FILE4_11 WHERE TYPE = '5' " & cWher22 & " )  " & _
'            " , ( SELECT SUM(SAL - PAY)         FROM FILE4_11 WHERE (TYPE = '5' OR TYPE = '4' )" & cWher22 & " )  " & _
'            " , ( SELECT SUM(FILE6_20.TOTAL * ( ( 100 - file6_20h.rate ) / 100 ))   FROM FILE6_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_20.ITEM INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20.DOC_NO IS NOT NULL " & cWher & " )  " & _
'            " , ( SELECT SUM(FILE6_20.QUANT * FILE6_20.COST ) FROM FILE6_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_20.ITEM INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO  WHERE FILE6_20.DOC_NO IS NOT NULL " & cWher & " ) " & _
'            " , ( SELECT SUM((FILE1_11.[IN] -FILE1_11.OUT) * FILE1_10.costitem ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE FILE1_10.ITEM IS NOT NULL " & cWher & " ) " & _
'            " , ( SELECT SUM(PAY - SAL )        FROM FILE4_11 WHERE (TYPE <> '5' AND TYPE <> '4' AND TYPE <> '9' )" & cWher22 & " )  " & _
'            " , ( SELECT SUM(PAY)               FROM FILE4_11 WHERE (TYPE = '9' )" & cWher22 & " )  " & _
'            " , ( SELECT SUM(SAL - PAY)         FROM FILE4_11 WHERE CODE IS NOT NULL " & cWher22 & " )  "
    
    
    cString = " SELECT COUNT(DISTINCT modelno) " & _
            " , ( SELECT SUM(TOTAL)             FROM FILE7_20 INNER JOIN FILE1_10 AS FILE1_10_2 ON FILE1_10_2.ITEM = FILE7_20.ITEM  WHERE FILE1_10_2.ITEM IS NOT NULL " & cWher22 & " )  " & _
            " , ( SELECT SUM(TOTAL)             FROM FILE7_10 INNER JOIN FILE1_10 AS FILE1_10_2 ON FILE1_10_2.ITEM = FILE7_10.ITEM  WHERE FILE1_10_2.ITEM IS NOT NULL " & cWher22 & " )  " & _
            " , ' ' AS N2  " & _
            " , ( SELECT SUM(FILE6_20.TOTAL * ( ( 100 - file6_20h.rate ) / 100 ))   FROM FILE6_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_20.ITEM INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20.DOC_NO IS NOT NULL " & cWher & " )  " & _
            " , ( SELECT SUM(FILE6_20.QUANT * FILE6_20.COST ) FROM FILE6_20 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE6_20.ITEM INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO  WHERE FILE6_20.DOC_NO IS NOT NULL " & cWher & " ) " & _
            " , ( SELECT SUM((FILE1_11.[IN] -FILE1_11.OUT) * FILE1_10.costitem ) FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_10.ITEM = FILE1_11.ITEM WHERE FILE1_10.ITEM IS NOT NULL " & cWher & " ) " & _
            " , ( SELECT SUM(PAY - SAL )        FROM FILE4_11 WHERE (TYPE <> '5' AND TYPE <> '4' AND TYPE <> '9' )" & cWher2 & " )  " & _
            " , ( SELECT SUM(PAY)               FROM FILE4_11 WHERE (TYPE = '9' )" & cWher2 & " )  " & _
            " , ( SELECT SUM(SAL - PAY)         FROM FILE4_11 WHERE CODE IS NOT NULL " & cWher2 & " )  "
    cWher = ""
    If xMosm.BoundText <> "" Then cWher = cWher & " AND MOSM = " & MyParn(xMosm.BoundText)
    If xSupp.BoundText <> "" Then cWher = cWher & " AND CODE = " & MyParn(xSupp.BoundText)
    If xFact.BoundText <> "" Then cWher = cWher & " AND FACT = " & MyParn(xFact.BoundText)
    cString = cString & " FROM  FILE1_10  AS FILE1_10_0 WHERE FILE1_10_0.ITEM IS NOT NULL " & cWher
    aRet = aGetDesca(cString, con)
    If UBound(aRet) > 0 Then
        .TextMatrix(0, 1) = Format(Val(aRet(1) & ""), "#0")
        
        .TextMatrix(1, 1) = Format(Val(aRet(2) & ""), "#0")
        .TextMatrix(1, 3) = Format(Val(aRet(3) & ""), "#0")
        .TextMatrix(1, 5) = Format(Val(aRet(2) & "") - Val(aRet(3) & ""), "#0")
        
        nTot = Val(.TextMatrix(1, 5))
        If nTot <> 0 Then nRate = Round(Val(aRet(6) & "") / nTot * 100, 2)
        .TextMatrix(0, 3) = Format(Val(nRate & ""), "#0.00")
        
        .TextMatrix(2, 1) = Format(Val(aRet(5) & ""), "#0")
        .TextMatrix(2, 3) = Format(Val(aRet(6) & ""), "#0")
        .TextMatrix(2, 5) = Format(Val(aRet(5) & "") - Val(aRet(6) & ""), "#0")
        
        If xSupp.BoundText <> "" Then
            .TextMatrix(3, 1) = Format(Val(aRet(8) & ""), "#0")
            .TextMatrix(3, 3) = Format(Val(aRet(9) & ""), "#0")
            
            .TextMatrix(4, 1) = Format(Val(aRet(10) & ""), "#0")
            .TextMatrix(4, 5) = Format(Val(aRet(10) & "") - Val(aRet(7) & ""), "#0")
        End If
        .TextMatrix(4, 3) = Format(Val(aRet(7) & ""), "#0")
    End If
    End With
    Err.Clear
End Sub
Private Sub XDOC_NO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        PurchLook
    End If
End Sub
Sub PurchLook()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(5, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT  DOC_NO,[DATE] AS DATE1,DATE, FILE4_10.Desca , inv_no , mosm " & _
                  " FROM  FILE7_20H left JOIN  FILE4_10 ON FILE7_20H.CODE   = FILE4_10.CODE "

Generalarray(2) = "Order by Date"
Generalarray(3) = 6000
Generalarray(5) = True


listarray(0, 0) = "«·—ﬁ„-≈”„ " & cCodeDesca & "-«· «—ÌŒ"
listarray(0, 1) = "(inv_no Like '%cFilter%' or Doc_No Like '%cFilter%' or  FILE4_10.DESCA LIKE '%cFilter%' OR " & _
                  "##date##)"


GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "≈”„ "
GrdArray(3, 1) = 3000

GrdArray(4, 0) = "—ﬁ„ ›« Ê—…"
GrdArray(4, 1) = 1000

GrdArray(5, 0) = "„Ê”„"
GrdArray(5, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
Load Search3
Search3.Caption = "«” ⁄·«„"
Search3.Show 1
End Sub


Private Sub myLoadPicture(pModel As String)
On Error Resume Next
    If Not lServerPict Then Exit Sub
    ImgX1.Images.Clear
    If pModel = "" Then Exit Sub
    If conPict Is Nothing Then Exit Sub
    Dim loctable As Recordset
    Dim bytes() As Byte, nSize As Long
    ImgX1.Images.Clear
    Set loctable = myRecordSet("select PICT,size from PICT where MODELNO = " & MyParn(pModel), conPict)
    If loctable.EOF Then Exit Sub
    nSize = CLng(loctable("size").Value)
    bytes = loctable("PICT").GetChunk(nSize)
    ImgX1.Import.FromMemoryFile bytes, ixmfJPG
    Err.Clear
End Sub


Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
With grid1
    If OldRow <> NewRow Then
        myLoadPicture .TextMatrix(NewRow, 0)
    End If
End With

End Sub

