VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "vsflex7L.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form ITEMS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "»Ì«‰«  «·„ÊœÌ·« "
   ClientHeight    =   7530
   ClientLeft      =   405
   ClientTop       =   1455
   ClientWidth     =   15240
   FillColor       =   &H00808080&
   FillStyle       =   0  'Solid
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
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   RightToLeft     =   -1  'True
   ScaleHeight     =   7530
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   4815
      RightToLeft     =   -1  'True
      TabIndex        =   55
      Top             =   4905
      Width           =   3840
      Begin Threed.SSCommand Calc2 
         Height          =   345
         Left            =   1035
         TabIndex        =   56
         Top             =   180
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1/10"
      End
      Begin Threed.SSCommand Calc 
         Height          =   345
         Left            =   90
         TabIndex        =   57
         Top             =   180
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1/12"
      End
      Begin Threed.SSCommand Calc3 
         Height          =   345
         Left            =   1980
         TabIndex        =   58
         Top             =   180
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1◊ 12"
      End
      Begin Threed.SSCommand Calc4 
         Height          =   345
         Left            =   2925
         TabIndex        =   59
         Top             =   180
         Visible         =   0   'False
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   609
         _Version        =   196610
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "1◊ 10"
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "’Ê—… «·„ÊœÌ· "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4875
      Left            =   90
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   675
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdScan 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   90
         Picture         =   "Items4.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   225
         Width           =   645
      End
      Begin VB.CommandButton cmdRotate 
         Caption         =   "Rotate"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1485
         Width           =   645
      End
      Begin VB.CommandButton BrtPlus 
         Caption         =   "B+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         TabIndex        =   44
         Top             =   3150
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton BrtMinus 
         Caption         =   "B-"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         TabIndex        =   43
         Top             =   2700
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton cmdDelPhoto 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–›"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   45
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   3600
         Width           =   690
      End
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   4500
         Left            =   765
         TabIndex        =   47
         Top             =   225
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   7938
         BackColor       =   16777215
         BorderStyle     =   4
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "amr01"
         LicenseRegCode  =   "íß“∂∫©æºµ≠¶“™Ωπß≥´¨≠MOCG-EOUEXNJI-EQTP6gI"
      End
   End
   Begin VB.Frame Frame8 
      Height          =   645
      Left            =   180
      TabIndex        =   38
      Top             =   9675
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   15240
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   0
      Width           =   15240
      Begin VB.Frame Frame7 
         Height          =   780
         Left            =   6435
         TabIndex        =   50
         Top             =   -90
         Width           =   4875
         Begin VB.CommandButton cmdSave 
            Caption         =   "Õ›Ÿ"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   3240
            TabIndex        =   53
            Top             =   225
            Width           =   1545
         End
         Begin VB.CommandButton cmdDel 
            BackColor       =   &H000000FF&
            Caption         =   "Õ–›"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   1665
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   225
            Width           =   1545
         End
         Begin VB.CommandButton Cmd_Undo 
            Caption         =   " —«Ã⁄"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   90
            TabIndex        =   51
            Top             =   225
            Width           =   1545
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "FIX MODEL"
         CausesValidation=   0   'False
         Height          =   465
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   45
         Width           =   2130
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "Œ—ÊÃ"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         MaskColor       =   &H00FFFFFF&
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   90
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.CommandButton cmd_show 
         Caption         =   "—’Ìœ «·„ÊœÌ·"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1275
         TabIndex        =   35
         Top             =   90
         Width           =   1635
      End
      Begin VB.CommandButton CmdInform 
         BackColor       =   &H00C0FFFF&
         Caption         =   "«” ⁄·«„"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   13635
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   90
         Width           =   1545
      End
      Begin VB.TextBox xItem 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
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
         Height          =   420
         Left            =   11340
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   90
         Width           =   1365
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»«—ﬂÊœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   12825
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   180
         Width           =   525
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid VsColor 
      Height          =   1815
      Left            =   180
      TabIndex        =   27
      Top             =   7785
      Width           =   14955
      _cx             =   26379
      _cy             =   3201
      _ConvInfo       =   1
      Appearance      =   1
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
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   12648384
      ForeColorFixed  =   8388608
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777215
      GridColor       =   -2147483627
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Items4.frx":273E
      ScrollTrack     =   -1  'True
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
      OutlineCol      =   1
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   4
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   4815
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   675
      Width           =   10320
      Begin VB.TextBox xModelNo 
         Alignment       =   2  'Center
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
         Left            =   90
         MaxLength       =   50
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   585
         Visible         =   0   'False
         Width           =   3120
      End
      Begin VB.TextBox xModel 
         Alignment       =   2  'Center
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
         Left            =   6390
         MaxLength       =   50
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   585
         Width           =   3120
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
         Left            =   1665
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   1545
      End
      Begin VB.TextBox xSupp 
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
         Left            =   90
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   180
         Width           =   555
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   315
         Left            =   7155
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   180
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo XMOSM 
         Height          =   315
         Left            =   4365
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   180
         Width           =   2085
         _ExtentX        =   3678
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ÊœÌ· :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   3285
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   195
         Width           =   1050
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ﬂ » Ã„·…"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   675
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   180
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’‰⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   9540
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   195
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1680
      Left            =   4815
      TabIndex        =   12
      Top             =   1845
      Width           =   10320
      Begin VB.TextBox xOkaz 
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
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1260
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox xRedem 
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
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox xDescA 
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
         Left            =   3420
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   540
         Width           =   4785
      End
      Begin VB.TextBox xCode 
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
         Left            =   6615
         Locked          =   -1  'True
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   900
         Width           =   1590
      End
      Begin VB.TextBox xRate 
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
         Left            =   6615
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1260
         Width           =   1590
      End
      Begin VB.TextBox xFixPrice 
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
         Left            =   3420
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1260
         Visible         =   0   'False
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo xGroup 
         Height          =   315
         Left            =   5400
         TabIndex        =   0
         Top             =   180
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin MSDataListLib.DataCombo xSection 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   180
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·ﬁ”„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„Ã„Ê⁄… «·„ÊœÌ·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   8325
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   195
         Width           =   1425
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… ”⁄— √Êﬂ«“ÌÊ‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1350
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Õœ ≈⁄«œ… «·ÿ·»"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1620
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   945
         Width           =   1185
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂÊœ «·„Ê—œ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   945
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "≈”„ «·„ÊœÌ·  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   585
         Width           =   1140
      End
      Begin VB.Label xCodeDesc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   900
         Width           =   3075
      End
      Begin VB.Label EQWEW 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·„” Â·ﬂ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1350
         Width           =   1245
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "”⁄— »Ì⁄ „ÊÕœ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1350
         Visible         =   0   'False
         Width           =   1170
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   9180
      Top             =   3735
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
      Left            =   9720
      Top             =   3780
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
   Begin VSFlex7LCtl.VSFlexGrid vsScal 
      Height          =   2130
      Left            =   180
      TabIndex        =   4
      Top             =   5580
      Width           =   14955
      _cx             =   26379
      _cy             =   3757
      _ConvInfo       =   1
      Appearance      =   1
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
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   12648384
      ForeColorFixed  =   8388608
      BackColorSel    =   -2147483643
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777215
      GridColor       =   -2147483627
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "lt"
      ScrollTrack     =   -1  'True
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
      OutlineCol      =   1
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   4
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
   End
   Begin VB.Frame Frame9 
      Height          =   645
      Left            =   2205
      TabIndex        =   24
      Top             =   9675
      Width           =   8070
      Begin VB.CommandButton xFIX 
         Caption         =   "÷»ÿ  ”⁄Ì— √Êﬂ«“ÌÊ‰"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   90
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   135
         Width           =   1950
      End
      Begin VB.CommandButton Cmd_Okaz 
         Caption         =   "÷»ÿ ”⁄— «·√Êﬂ«“ÌÊ‰"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2070
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   135
         Width           =   1950
      End
      Begin VB.CommandButton xCmd_Fix 
         Caption         =   "Õ›Ÿ «·”⁄— «·„ÊÕœ"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4050
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   135
         Width           =   1950
      End
   End
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   9180
      Top             =   3645
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
   Begin MSAdodcLib.Adodc DATA4 
      Height          =   330
      Left            =   9000
      Top             =   3735
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
   Begin VSFlex7LCtl.VSFlexGrid VSMYCOLOR 
      Height          =   915
      Left            =   4815
      TabIndex        =   60
      Top             =   3915
      Width           =   10320
      _cx             =   18203
      _cy             =   1614
      _ConvInfo       =   1
      Appearance      =   1
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
      MousePointer    =   1
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   12648384
      ForeColorFixed  =   8388608
      BackColorSel    =   -2147483643
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   16777215
      GridColor       =   -2147483627
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   "lt"
      ScrollTrack     =   -1  'True
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
      OutlineCol      =   1
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   1
      OwnerDraw       =   4
      Editable        =   2
      ShowComboButton =   -1  'True
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   4
   End
End
Attribute VB_Name = "ITEMS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection, bEdit As Boolean
Public CardTable As New ADODB.Recordset
Dim ColorTable As New ADODB.Recordset
Public sItem As String
Public strX1 As String, strX2 As String, strX3 As String, strX4 As String, strDesca As String, strModel As String, StrCode As String, strCodeDesca As String
Dim nRecordNumber As Integer
Sub mydefine()
ImgX1.Images.Clear

xmodel.Text = ""
xFact.BoundText = ""
xmodelfact.Text = ""
xMosm.BoundText = ""
xGroup.BoundText = ""
xSection.BoundText = ""
xSupp.Text = ""
xRate.Text = ""
xCode.Text = ""
xCodeDesc.Caption = ""
xFixPrice.Text = ""
xRedem.Text = ""
xOkaz.Text = ""
xDescA.Text = ""
vsScal.Rows = 1
VsColor.Rows = 1
xItem.Text = ""
myDefineGrd
Handlecontrols DefineMode
End Sub
Sub myProc()
If ActiveControl.Name = CmdInform.Name Then
    CardTable.Find "Model = " & MyParn(Search3.grid1.TextMatrix(Search3.grid1.Row, 0)), , adSearchForward, adBookmarkFirst
    If Not CardTable.EOF Then myload
    Unload Search3
End If
End Sub
Public Sub myload()
xmodel.Text = CardTable!Model
xmodelfact.Text = DelZero(CardTable!modelfact) & ""
xSupp.Text = CardTable!SUPP & ""
xMosm.BoundText = CardTable!MOSM & ""
xFact.BoundText = CardTable!Fact & ""
xOkaz.Text = CardTable!OKAZ & ""
xRate.Text = Format(CardTable!Rate, "#0.00")
xRedem.Text = Format(CardTable!REDEM, "#0.00")
xGroup.BoundText = CardTable!Group & ""
xSection.BoundText = CardTable!Section & ""
xFixPrice.Text = Format(CardTable!FixPrice, "#0.00")
xCode.Text = CardTable!CODE & ""
xCodeDesc.Caption = CardTable!SUPDESCA & ""
xDescA.Text = CardTable!DESCA & ""
xmodelno.Text = xFact.Text & xMosm.BoundText & DelZero(xmodelfact.Text)
LoadPhoto
myloadgrd
myloadgrd2
xRecordNumber = "”Ã· " & CardTable.AbsolutePosition + 1 & " „‰ " & nRecordNumber
Handlecontrols LoadMode
End Sub
Private Sub BrtMinus_Click()
On Error GoTo myerror
If validPhoto(retFile(xmodel.Text)) Then
    ImgX1.Filters.Brightness (-5)
    ImgX1.Export.ToFile retFile(xmodel.Text), ixfsJPG
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub BrtPlus_Click()
On Error GoTo myerror
If validPhoto(retFile(xmodel.Text)) Then
    ImgX1.Filters.Brightness (5)
    ImgX1.Export.ToFile retFile(xmodel.Text), ixfsJPG
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub Calc_Click()
With vsScal
    If Val(.TextMatrix(1, .Col)) > 0 Then .TextMatrix(1, .Col) = Format(Val(.TextMatrix(1, .Col)) / 12, "#0.00")
End With
End Sub
Private Sub Calc2_Click()
With vsScal
    If Val(.TextMatrix(1, .Col)) > 0 Then .TextMatrix(1, .Col) = Format(Val(.TextMatrix(1, .Col)) / 10, "#0.00")
End With
End Sub
Private Sub Calc3_Click()
With vsScal
    If Val(.TextMatrix(1, .Col)) > 0 Then .TextMatrix(1, .Col) = Format(Val(.TextMatrix(1, .Col)) * 12, "#0.00")
End With
End Sub
Private Sub Calc4_Click()
With vsScal
    If Val(.TextMatrix(1, .Col)) > 0 Then .TextMatrix(1, .Col) = Format(Val(.TextMatrix(1, .Col)) * 10, "#0.00")
End With
End Sub

Private Sub Cmd_Undo_Click()
CardTable.Requery
If CardTable.EOF And CardTable.BOF Then
    mydefine
Else
    CardTable.Find "MODEL = " & MyParn(xmodel.Text), , adSearchForward, adBookmarkFirst
    If CardTable.EOF Then CardTable.MoveLast
    myload
End If
End Sub
Private Sub CmdDel_Click()
If GetDesca("SELECT FILE1_10.ITEM FROM FILE1_10 INNER JOIN FILE1_11 ON FILE1_11.ITEM = FILE1_10.ITEM WHERE FILE1_10.MODEL = " & MyParn(xmodel.Text)) <> "" Then
    MsgBox "ÌÊÃœ Õ—ﬂ… ⁄·Ï «·„ÊœÌ·"
    Exit Sub
End If
On Error GoTo myerror
con.BeginTrans
If MsgBox("”Ì „ «·€«¡ «·„ÊœÌ· »«·ﬂ«„· ", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
    cString = " DELETE  FROM FILE1_10 WHERE MODEL = " & MyParn(xmodel.Text)
    con.Execute cString
End If
con.CommitTrans
mydefine
CardTable.Requery
MsgBox "  „ Õ–› «·„ÊœÌ· "
Exit Sub
myerror:
con.RollbackTrans
MsgBox "·„ Ì „ Õ–› «·„ÊœÌ· - «·„Õ«Ê·… „—… «Œ—Ï  "
Err.Clear

End Sub

Private Sub cmdDelPhoto_Click()
On Error GoTo myerror
If validPhoto(retFile(xmodel.Text)) Then fs.DeleteFile retFile(xmodel.Text)
LoadPhoto
Handlecontrols LoadMode
Exit Sub
myerror:
    MsgBox Err.Description
    Err.Clear
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdInform_Click()
    ModelLookupAll Me, Search3
End Sub

Private Sub cmdRotate_Click()
On Error GoTo myerror
If validPhoto(retFile(xmodel.Text)) Then
    ImgX1.Effects.Rotate 90
    ImgX1.Export.ToFile retFile(xmodel.Text), ixfsJPG
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub cmdSave_Click()
If Not MYVALID Then Exit Sub
If Not myreplace Then Exit Sub
Inform " „ Õ›Ÿ «·»Ì«‰«  »‰Ã«Õ"
If lExirSave Then Unload Me
End Sub

Private Sub cmdScan_Click()
If xFact.Text <> "" And xMosm.BoundText <> "" And xFact.BoundText <> "" Then
    xmodelno.Text = xFact.Text & xMosm.BoundText & DelZero(xmodelfact.Text)
    MyCreateFolder App.Path & "\photos"
    ScanImage
End If
End Sub

Private Sub Command1_Click()
Dim cMod As String
Dim cMod0 As String
Dim cModel As String
Dim cModelNo  As String
With CardTable
    Do While Not .EOF
        Me.Caption = !Model
        cMod = RetZero(!modelfact, 10)
        cMod = Replace(cMod, "'", "-")
        cMod0 = DelZero(cMod)
        cModelNo = !Fact & !MOSM & cMod
        cModel = !Fact & !MOSM & cMod & TurnValue(!SUPP)
        con.Execute " UPDATE FILE1_10 SET MODELFACT = " & MyParn(cMod) & " , MODELFACT0 = " & MyParn(cMod0) & " , MODEL = " & MyParn(cModel) & " , MODELNO = " & MyParn(cModelNo) & " WHERE MODEL = " & MyParn(!Model)
        .MoveNext
    Loop
End With
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And (TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo) Then SendKeys "{tAB}"
End Sub
Private Sub Form_Load()
Dim ImgXTmp As ImgX
Set ImgXTmp = New ImgX
ImgXTmp.Import.FromMemoryFile LoadResData("BACKGROUND", "CUSTOM")
Set ImgX1.BackgroundImage = ImgXTmp.Image
Set ImgXTmp = Nothing

ImgX1.AutoZoom = True
ImgX1.AutoZoomType = 1

con.CursorLocation = adUseClient
con.Open strCon
CardTable.Open "select file1_10h.*,file4_10.desca as supDesca from file1_10h left join file4_10 on file1_10h.code = file4_10.code   order by model", con, adOpenStatic, adLockReadOnly, acmdtext

ColorTable.Open "COLOR", con, adOpenStatic, adLockReadOnly, adCmdTable
With VSMYCOLOR
    .Editable = flexEDNone
    .Rows = 1
    .Cols = 0
    .FixedCols = 0
    .FixedRows = 0
    .RowHeight(0) = 600
    Do While Not ColorTable.EOF
        .Cols = .Cols + 1
        .ColWidth(.Cols - 1) = 800
        .TextMatrix(0, .Cols - 1) = ColorTable!Color
        ColorTable.MoveNext
    Loop
End With


data1.ConnectionString = strCon
data1.RecordSource = "SELECT * FROM FACT ORDER BY DESCA "
Set xFact.RowSource = data1
xFact.ListField = "DESCA"
xFact.BoundColumn = "CODE"

data2.ConnectionString = strCon
data2.RecordSource = "SELECT * FROM FILE1_50 ORDER BY DESCA"
Set xGroup.RowSource = data2
xGroup.ListField = "DESCA"
xGroup.BoundColumn = "CODE"

DATA3.ConnectionString = strCon
DATA3.RecordSource = "SELECT * FROM FILE1_10SC ORDER BY DESCA"
Set xSection.RowSource = DATA3
xSection.ListField = "DESCA"
xSection.BoundColumn = "CODE"

data4.ConnectionString = strCon
data4.RecordSource = "SELECT * FROM MOSM ORDER BY DATE DESC "
Set xMosm.RowSource = data4
xMosm.ListField = "DESCA"
xMosm.BoundColumn = "MOSM"

myDefineGrd
vsScal.RowHidden(vsScal.Rows - 1) = True
VsColor.RowHidden(VsColor.Rows - 1) = True
mydefine

End Sub
Private Sub Form_Unload(Cancel As Integer)
    closeCon con
End Sub
Private Sub VsColor_AfterEdit(ByVal Row As Long, ByVal Col As Long)
With VsColor
    If .TextMatrix(0, .Col) = "" And .TextMatrix(1, .Col) <> "" Then
        .TextMatrix(0, .Col) = Val(.TextMatrix(0, .Col - 1)) + 1
    End If
End With
End Sub
Private Sub VsColor_EnterCell()
With VsColor
    If .Row = 1 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub VsColor_KeyDown(KeyCode As Integer, Shift As Integer)
With VsColor
If KeyCode = 46 And .Col <> .Cols - 1 And .Col <> 0 And .Row = 0 Then
    If .TextMatrix(.Rows - 1, .Col) <> "" Then
        If GetDesca("SELECT ITEM FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM WHERE FILE1_10.MODEL = " & MyParn(xmodel.Text) & " AND C_COLOR = " & MyParn(.TextMatrix(.Rows - 1, .Col))) <> "" Then
            MsgBox "ÌÊÃœ Õ—ﬂ… ⁄·Ï «··Ê‰ "
            Exit Sub
        End If
    End If
    
    If MsgBox("Õ–› «··Ê‰ ø", vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
    If Trim(.TextMatrix(.Rows - 1, .Col)) = "" Then
        DelCol VsColor, .Col
    Else
        con.BeginTrans
        cString = "DELETE  FROM FILE1_10 WHERE MODEL = " & MyParn(xmodel.Text) & " AND c_color = " & MyParn(.TextMatrix(.Rows - 1, .Col))
        con.Execute cString, nAffect
        con.CommitTrans
        nCol = .Col
        myloadgrd2
        .Col = nCol
    End If
End If
End With
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub VsColor_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VsColor
If Col = .Cols - 1 Then
    .Cols = .Cols + 1
    FixBackColor2
End If
End With
End Sub
Private Sub VsColor_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With VsColor
If Trim(.EditText) = "" Then
    MsgBox "·Ê‰ €Ì— „”Ã·"
    Cancel = True
Else
    .EditText = UCase(.EditText)
    For i = 1 To .Cols - 2
        If i <> Col Then
            If Trim(.EditText) = .TextMatrix(0, i) Then
                MsgBox " „ «œŒ«· ‰›” «··Ê‰ „‰ ﬁ»·"
                Cancel = True
            End If
        End If
    Next
End If
End With
End Sub

Private Sub VSMYCOLOR_DBLClick()
With VSMYCOLOR
    VsColor.TextMatrix(0, VsColor.Cols - 1) = Val(VsColor.TextMatrix(0, VsColor.Cols - 2)) + 1
    VsColor.TextMatrix(1, VsColor.Cols - 1) = .TextMatrix(.Row, .Col)
    VsColor.Cols = VsColor.Cols + 1
End With
End Sub
Private Sub VsScal_EnterCell()
With vsScal
    If .Row = 0 And Trim(.TextMatrix(.Rows - 1, .Col)) <> "" Then
        .Editable = flexEDNone
    Else
        If .Row = 3 And Not lSupperVisor Then
            .Editable = False
        Else
            .Editable = flexEDKbdMouse
        End If
    End If
        
    .Cell(flexcpBackColor, 0, 0, .Rows - 1, .Cols - 1) = vbWhite
    .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = vbYellow
    If .Row = 1 And .Col > 1 And Val(.TextMatrix(1, .Col)) = 0 Then
         .TextMatrix(1, .Col) = Val(.TextMatrix(1, .Col - 1))
    End If
    
    If .Row = 2 And .Col > 1 And Val(.TextMatrix(2, .Col)) = 0 Then
         .TextMatrix(2, .Col) = Val(.TextMatrix(2, .Col - 1))
    End If
    If .Row = 0 And .Col > 2 And .TextMatrix(0, .Col) = "" Then
        If IsNumeric(.TextMatrix(0, .Col - 2)) Then
            .TextMatrix(0, .Col) = Val(.TextMatrix(0, .Col - 1)) + Val(.TextMatrix(0, .Col - 1)) - Val(.TextMatrix(0, .Col - 2))
            vsScal_StartEdit vsScal.Row, vsScal.Col, False
        Else
            Select Case .TextMatrix(O, .Col - 1)
                Case "XS"
                    .TextMatrix(0, .Col) = "S"
                    .Cols = .Cols + 1
                Case "S"
                    .TextMatrix(0, .Col) = "M"
                    .Cols = .Cols + 1
                Case "M"
                    .TextMatrix(0, .Col) = "L"
                    .Cols = .Cols + 1
                Case "L"
                    .TextMatrix(0, .Col) = "XL"
                    .Cols = .Cols + 1
                Case "XL"
                    .TextMatrix(0, .Col) = "XXL"
                    .Cols = .Cols + 1
                Case "XXL"
                    .TextMatrix(0, .Col) = "3XL"
                    .Cols = .Cols + 1
                Case "3XL"
                    .TextMatrix(0, .Col) = "4XL"
                    .Cols = .Cols + 1
                Case "4XL"
                    .TextMatrix(0, .Col) = "5XL"
                    .Cols = .Cols + 1
            End Select
        End If
    End If
End With
End Sub
Private Sub vsScal_KeyUp(KeyCode As Integer, Shift As Integer)
'On Error GoTo myerror
With vsScal
    If KeyCode = 45 Then
        InsertCol vsScal, .Col
    End If
    If KeyCode = 46 And .Col <> .Cols - 1 And .Col <> 0 And .Row = 0 Then
        If GetDesca("SELECT file1_10.ITEM FROM FILE1_11 INNER JOIN FILE1_10 ON FILE1_11.ITEM = FILE1_10.ITEM WHERE FILE1_10.MODEL = " & MyParn(xmodel.Text) & " AND C_SCAL = " & MyParn(.TextMatrix(.Rows - 1, .Col))) <> "" Then
            MsgBox "ÌÊÃœ Õ—ﬂ… ⁄·Ï «··Ê‰ "
            Exit Sub
        End If
        
        If MsgBox("Õ–› «·„ﬁ«” ø", vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
        If Trim(.TextMatrix(.Rows - 1, .Col)) = "" Then
            DelCol vsScal, .Col
        Else
'           con.BeginTrans
            cString = "DELETE  FROM FILE1_10 WHERE file1_10.MODEL = " & MyParn(xmodel.Text) & " AND SCAL = " & MyParn(.TextMatrix(.Rows - 1, .Col))
            con.Execute cString
'           con.CommitTrans
            nCol = .Col
            myloadgrd
            .Col = nCol
        End If
    End If
End With
Exit Sub
myerror:
'con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub

Private Sub VsScal_LeaveCell()
With vsScal
If .Row = 2 And Trim(.TextMatrix(3, .Col)) = "" Then
    If Val(xOkaz.Text) > 0 Then
        .TextMatrix(3, .Col) = Format(Val(.TextMatrix(2, .Col)) * Val(xOkaz.Text) / 100, "#0.00")
    End If
End If


Dim nPrice As Double
Dim nRem As Double
With vsScal
    If .Row = 1 Then
    If xRate.Text <> "" And Val(.TextMatrix(.Row, .Col)) > 0 And Val(.TextMatrix(3, .Col)) = 0 Then
        nPrice = Val(.TextMatrix(.Row, .Col)) * (100 + Val(xRate.Text)) / 100
        nRem = nPrice - Int(nPrice)
        Select Case nRem
            Case Is = 0
                nPrice = Int(nPrice)
            Case Is > 0
                nPrice = Int(nPrice) + 1
'            Case Is < 0.5
'                nPrice = Int(nPrice) + 0.5
        End Select
        .TextMatrix(2, .Col) = Format(nPrice, "#0.00")
    End If
    End If
End With

End With
End Sub
Private Sub vsScal_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With vsScal
If Col = .Cols - 1 Then
    .Cols = .Cols + 1
    FixBackColor
End If
End With
End Sub
Private Sub VsScal_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With vsScal
If Row = 0 Then
    If Trim(.EditText) = "" Then
        MsgBox "„ﬁ«” €Ì— „”Ã·"
        Cancel = True
    Else
        .EditText = UCase(.EditText)
        For i = 1 To .Cols - 2
            If i <> Col Then
                If Trim(.EditText) = .TextMatrix(0, i) Then
                    MsgBox " „ «œŒ«· ‰›” «·„ﬁ«” „‰ ﬁ»·"
                    Cancel = True
                End If
            End If
        Next
    End If
End If
End With
End Sub
Private Sub xCmd_Fix_Click()
If MsgBox("”Ì „  ⁄œÌ· ”⁄— «·√Êﬂ«“ÌÊ‰ ·ÌﬂÊ‰ ”⁄— „ÊÕœ »ﬁÌ„… " & Format(xFixPrice.Text, "##0.00"), vbOKCancel) = vbOK Then
    With vsScal
    For i = 1 To .Cols - 1
        .TextMatrix(4, i) = Val(xFixPrice.Text)
    Next i
    End With

    cStr1 = " update FILE1_10 SET file1_10.price2 = " & Val(xFixPrice.Text) & _
            " where item is not null "
    If xFact.Text <> "" Then cStr1 = cStr1 & " and fact = " & MyParn(xFact.Text)
    If xmodelfact.Text <> "" Then cStr1 = cStr1 & " and modelfact = " & MyParn(xmodelfact.Text)
    If xMosm.BoundText <> "" Then cStr1 = cStr1 & " and mosm = " & MyParn(xMosm.BoundText)
    mydb.Execute cStr1
End If
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
Private Sub xDescA_GotFocus()
    If Trim(xDescA.Text) = "" Then xDescA.Text = xGroup.Text & "  "
    xDescA.SelStart = Len(xDescA.Text) + 1
    xDescA.SelLength = 1
    xDescA.BackColor = vbYellow
End Sub

Private Sub xDescA_LostFocus()
xDescA.BackColor = vbWhite
End Sub

Private Sub xgroup_GotFocus()
    xGroup.BackColor = vbYellow
End Sub
Private Sub xGroup_LostFocus()
    xGroup.BackColor = vbWhite
    If xDescA.Text = "" Then xDescA.Text = xGroup.Text & "  "
End Sub
Private Sub myloadgrd()
Dim loctable As New ADODB.Recordset, i As Integer
loctable.CursorLocation = adUseClient
loctable.Open "select scal,Cost,Price ,PRICE2,Cost2  , c_scal from file1_10 where model = " & MyParn(xmodel.Text) & " group by file1_10.c_scal,scal,Cost,Price ,Cost2,price2 order by file1_10.c_scal", con, adOpenForwardOnly, adLockReadOnly, adCmdText
With vsScal
.Cols = 1
Do Until loctable.EOF
    .Cols = .Cols + 1
    For i = 0 To 5
        .TextMatrix(i, .Cols - 1) = loctable.Fields(i) & ""
    Next
    loctable.MoveNext
Loop
.Cols = .Cols + 1
Fixgrd
loctable.Close
Set loctable = Nothing
End With
End Sub
Private Sub myloadgrd2()
Dim loctable As New ADODB.Recordset
loctable.CursorLocation = adUseServer
loctable.Open "select file1_10.C_Color,Color from file1_10 where model = " & MyParn(xmodel.Text) & " group by file1_10.C_color,COLOR order by file1_10.C_color", con, adOpenStatic, adLockReadOnly, adCmdText
With VsColor
.Cols = 1
Do Until loctable.EOF
    .Cols = .Cols + 1
    .TextMatrix(0, .Cols - 1) = loctable!c_Color & ""
    .TextMatrix(1, .Cols - 1) = loctable!Color & ""
    loctable.MoveNext
Loop
.Cols = .Cols + 1
Fixgrd2
End With
loctable.Close
Set loctable = Nothing
End Sub
Private Sub Fixgrd()
With vsScal
For i = 1 To .Cols - 1
    .ColWidth(i) = (vsScal.Width / 15) - 200
Next
End With
FixBackColor
End Sub
Private Sub Fixgrd2()
With VsColor
For i = 1 To .Cols - 1
    .ColWidth(i) = (vsScal.Width / 15) - 200
Next
End With
FixBackColor2
End Sub
Private Sub FixBackColor()
With Me.vsScal
.Cell(flexcpBackColor, 0, 1, 0, .Cols - 1) = &H8000000F
For i = 1 To .Cols - 1
    .Cell(flexcpBackColor, 0, i, 0, i) = IIf(Trim(.TextMatrix(.Rows - 1, i)) = "", &HFFFFFF, &H8000000F)
Next
End With
End Sub
Private Sub FixBackColor2()
With Me.VsColor
'.Cell(flexcpBackColor, 0, 1, 0, .Cols - 1) = &H8000000F
For i = 1 To .Cols - 1
    .Cell(flexcpBackColor, 0, i, 0, i) = IIf(Trim(.TextMatrix(.Rows - 1, i)) = "", &HFFFFFF, &H8000000F)
Next
End With
End Sub
Private Sub InsertCol(pGrid As Variant, pos%)
With pGrid
.Cols = .Cols + 1
.ColPosition(.Cols - 1) = pos
End With
End Sub
Private Sub DelCol(pGrid As Variant, pos%)
With pGrid
.Cols = .Cols + 1
.ColPosition(pos) = .Cols - 1
.Cols = .Cols - 2

End With
End Sub
Private Function myreplace() As Boolean
On Error GoTo myerror
Dim nItem As Integer
Dim aInsert(23, 1)
Dim nColor As Integer, nScal As Integer
Dim lNew As Boolean

If xSupp.Text <> "" Then
    cModel = xFact.BoundText & xMosm.BoundText & RetZero(xmodelfact.Text, 10) & xSupp.Text
Else
    cModel = xFact.BoundText & xMosm.BoundText & RetZero(xmodelfact.Text, 10)
End If

cModelNo = xFact.BoundText & xMosm.BoundText & RetZero(xmodelfact.Text, 10)
aInsert(0, 0) = "Model"
aInsert(0, 1) = addstring(cModel)

aInsert(1, 0) = "ModelNO"
aInsert(1, 1) = addstring(cModelNo)

aInsert(2, 0) = "Desca"
aInsert(2, 1) = addstring(xDescA.Text)

aInsert(3, 0) = "FACT"
aInsert(3, 1) = addstring(xFact.BoundText)

aInsert(4, 0) = "MOSM"
aInsert(4, 1) = addstring(xMosm.BoundText)

aInsert(5, 0) = "MODELFACT"
aInsert(5, 1) = addstring(RetZero(xmodelfact.Text, 10))

aInsert(6, 0) = "MODELFACT0"
aInsert(6, 1) = addstring(DelZero(xmodelfact.Text))

aInsert(7, 0) = "SUPP"
aInsert(7, 1) = addstring(xSupp.Text)

aInsert(8, 0) = "FIXPRICE"
aInsert(8, 1) = Val(xFixPrice.Text)

aInsert(9, 0) = "OKAZ"
aInsert(9, 1) = Val(xOkaz.Text)

aInsert(10, 0) = "RATE"
aInsert(10, 1) = Val(xRate.Text)

aInsert(11, 0) = "REDEM"
aInsert(11, 1) = Val(xRedem.Text)

aInsert(12, 0) = "Code"
aInsert(12, 1) = addstring(xCode.Text)

aInsert(13, 0) = "[group]"
aInsert(13, 1) = addstring(xGroup.BoundText)

aInsert(14, 0) = "[section]"
aInsert(14, 1) = addvalue(xSection.BoundText)

For nScal = 1 To vsScal.Cols - 1
    For nColor = 1 To VsColor.Cols - 1
        lNew = False
        If Val(VsColor.TextMatrix(0, nColor)) <> 0 And vsScal.TextMatrix(0, nScal) <> "" Then
            nItem = Val(GetDesca("SELECT ITEM FROM FILE1_10 WHERE MODEL = " & MyParn(xmodel.Text) & " AND C_COLOR = " & Val(VsColor.TextMatrix(0, nColor)) & " AND SCAL = " & MyParn(vsScal.TextMatrix(0, nScal))) & "")
            If nItem = 0 Then
                nItem = Val(GetDesca("SELECT MAX(ITEM) FROM FILE1_10") & "") + 1
                If nItem < 100 Then nItem = 101
                lNew = True
            End If
            aInsert(15, 0) = "item"
            aInsert(15, 1) = addvalue(nItem)
            
            aInsert(16, 0) = "Cost"
            aInsert(16, 1) = addvalue(vsScal.TextMatrix(1, nScal))
            
            aInsert(17, 0) = "Price"
            aInsert(17, 1) = addvalue(vsScal.TextMatrix(2, nScal))
            
            aInsert(18, 0) = "Price2"
            aInsert(18, 1) = addvalue(vsScal.TextMatrix(3, nScal))
            
            aInsert(19, 0) = "Cost2"
            aInsert(19, 1) = addvalue(vsScal.TextMatrix(4, nScal))
        
            aInsert(20, 0) = "Scal"
            aInsert(20, 1) = addstring(vsScal.TextMatrix(0, nScal))
        
            aInsert(21, 0) = "C_Scal"
            aInsert(21, 1) = nScal
            
            aInsert(22, 0) = "Color"
            aInsert(22, 1) = addstring(VsColor.TextMatrix(1, nColor))
    
            aInsert(23, 0) = "C_Color"
            aInsert(23, 1) = addstring(VsColor.TextMatrix(0, nColor))
    
'            con.BeginTrans
            If lNew Then
                con.Execute CreateInsert(aInsert, "File1_10")
            Else
                con.Execute CreateUpdate(aInsert, "File1_10", " where item = " & nItem)
            End If
'            con.CommitTrans
        End If
    Next nColor
Next nScal
CardTable.Requery
myreplace = True
Exit Function
myerror:
'con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Private Function MYVALID() As Boolean
If xFact.BoundText = "" Then
    MsgBox "«·„’‰⁄"
    Exit Function
End If
If xmodelfact.Text = "" Then
    MsgBox "—ﬁ„ «·„ÊœÌ·"
    Exit Function
End If
If xMosm.BoundText = "" Then
    MsgBox "«·„Ê”„ €Ì— „”Ã·"
    Exit Function
End If

If xDescA.Text = "" Then
    MsgBox "≈”„ «·’‰› €Ì— „”Ã·"
    Exit Function
End If

If xGroup.BoundText = "" Then
    MsgBox "«·„Ã„Ê⁄… €Ì— „”Ã·…"
    Exit Function
End If

With vsScal
    For nCol = 1 To .Cols - 3
        If .TextMatrix(0, nCol) = "" Then
            MsgBox "„ﬁ«” €Ì— „”Ã·"
            Exit Function
        End If
        For K = nCol + 1 To .Cols - 2
            If Trim(.TextMatrix(0, nCol)) = Trim(.TextMatrix(0, K)) Then
                MsgBox "«·„ﬁ«” „ﬂ——"
                Exit Function
            End If
        Next
    Next
End With
      
With VsColor
    For nCol = 1 To .Cols - 3
        If .TextMatrix(0, nCol) = "" Then
            MsgBox "·Ê‰ €Ì— „”Ã·"
            Exit Function
        End If
        For K = nCol + 1 To .Cols - 2
            If Trim(.TextMatrix(0, nCol)) = Trim(.TextMatrix(0, K)) Then
                MsgBox "«··Ê‰ „ﬂ——"
                Exit Function
            End If
        Next
    Next
End With
MYVALID = True
End Function
Private Sub Handlecontrols(nMode)

CmdDel.Enabled = (nMode = LoadMode)
CmdInform.Enabled = (nMode = LoadMode)
xmodel.Enabled = Not (nMode = LoadMode)
End Sub
Private Sub myDefineGrd()
With vsScal
    .Cols = 1
    .Rows = 6
    .TextMatrix(0, 0) = "«·„ﬁ«”"
    .TextMatrix(1, 0) = "”⁄— «·„’‰⁄"
    .TextMatrix(2, 0) = "”⁄— «·„” Â·ﬂ"
    .TextMatrix(3, 0) = "”⁄— √Êﬂ«“ÌÊ‰"
    .TextMatrix(4, 0) = "”⁄— Ã„·…"
    .TextMatrix(5, 0) = ""
    .ColWidth(0) = 2000
    .FixedCols = 1
    .Cols = .Cols + 1
    .TextMatrix(0, 1) = "---"
    .TextMatrix(5, 1) = 1
    .RowHidden(5) = True
End With
With VsColor
    .Cols = 1
    .Rows = 3
    .TextMatrix(0, 0) = "«·ﬂÊœ"
    .TextMatrix(1, 0) = "«··Ê‰"
    .ColWidth(0) = 1500
    .Cols = .Cols + 1
    .TextMatrix(1, 1) = "---"
    .TextMatrix(0, 1) = 1
End With
End Sub
Private Sub xGroup_Validate(Cancel As Boolean)
'On Error GoTo myerror
Dim cGr As String
If xGroup.Text = "" Then
    Cancel = True
Else
    cGr = GetDesca("SELECT CODE FROM FILE1_50 WHERE CODE = " & MyParn(xGroup.BoundText))
    If cGr = "" Then
        If MsgBox("„Ã„Ê⁄… ÃœÌœ… Â·  —Ìœ  ”ÃÌ·Â« ", vbYesNo) = vbYes Then
            cGr = IncRec(GetDesca("SELECT MAX(CODE)  FROM FILE1_50 "))
            If cGr = "" Then cGr = "001"
            con.BeginTrans
            con.Execute " INSERT INTO FILE1_50 ( CODE , DESCA ) VALUES( " & addstring(cGr) & " , " & addstring(xGroup.Text) & " ) "
            con.CommitTrans
            data2.Refresh
            xGroup.BoundText = cGr
        
        End If
    End If
End If
Exit Sub
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Sub
Private Sub xITEM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And IsNumeric(xItem) Then
        With CardTable
            xmodel.Text = GetDesca("SELECT MODEL FROM FILE1_10 WHERE ITEM = " & Val(xItem.Text))
            If xmodel.Text <> "" Then
                CardTable.Find "Model = " & MyParn(xmodel.Text), , adSearchForward, adBookmarkFirst
                myload
            Else
                mydefine
            End If
        End With
    End If
End Sub
Private Sub xRate_GotFocus()
    xRate.BackColor = vbYellow
End Sub
Private Sub xRate_LostFocus()
    xRate.BackColor = vbWhite
End Sub
Private Sub LoadPhoto()
On Error GoTo myerror
ImgX1.Images.Clear
ImgX1.Import.FromFile retFile(xmodelno.Text)
Exit Sub
myerror:
Err.Clear
End Sub

Private Sub ScanImage()
'On Error GoTo myerror
Set twain = New ImgXTwain
twain.OpenTwain Me.hWnd
If twain.QuerySupport(ixtcResolution) Then
     twain.Resolution = 150
End If
twain.Acquire False, Me.hWnd
Exit Sub
myerror:
MsgBox Err.Number & vbCrLf & Err.Description
Err.Clear
End Sub
Private Sub Twain_ImageAcquired(Image As ImgX_Image)
    addImage Image
End Sub
Private Sub Twain_TwainError(ByVal erNum As Long, ByVal erSource As String, ByVal Description As String)
MsgBox "Error Number:  " & erNum & vbCrLf & Description, vbInformation, erSource
End Sub
Private Sub Twain_CanCloseTwain()
    ' This event is called after you call Acquire.
    ' It let's you know when it's safe to call CloseTwain.
    twain.CloseTwain
    ' Steps menu
End Sub
Private Sub addImage(Image As ImgX_Image)
Dim csource As String
csource = retFile(xmodelno.Text)
ImgX1.Images.Replace Image, , False
ImgX1.TIFCompression = ixtcJPEGCompression
ImgX1.Export.ToFile csource, ixfsTIF
End Sub
