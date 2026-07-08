VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form salestAX 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   10095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17760
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
   ScaleHeight     =   10095
   ScaleWidth      =   17760
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.TextBox XINV_TAX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5580
      MaxLength       =   100
      RightToLeft     =   -1  'True
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   1935
      Width           =   3660
   End
   Begin VB.Frame Frame10 
      Height          =   4245
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   110
      Top             =   2340
      Width           =   17520
      Begin TabDlg.SSTab SSTab1 
         Height          =   4020
         Left            =   90
         TabIndex        =   111
         Top             =   180
         Width           =   17385
         _ExtentX        =   30665
         _ExtentY        =   7091
         _Version        =   393216
         MousePointer    =   1
         Tabs            =   2
         Tab             =   1
         TabsPerRow      =   2
         TabHeight       =   529
         TabCaption(0)   =   "«·›« Ê—… «·≈·ﬂ —ÊÌ‰…"
         TabPicture(0)   =   "salesTax.frx":0000
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "SSCommand1"
         Tab(0).Control(1)=   "GRID2"
         Tab(0).ControlCount=   2
         TabCaption(1)   =   " ›’Ì·Ï «·›« Ê—…"
         TabPicture(1)   =   "salesTax.frx":001C
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "grid1"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         Begin VSFlex7Ctl.VSFlexGrid grid1 
            Height          =   3570
            Left            =   90
            TabIndex        =   112
            Top             =   360
            Width           =   17205
            _cx             =   30348
            _cy             =   6297
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Simplified Arabic"
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
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   3
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
            AutoSizeMouse   =   0   'False
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
         Begin VSFlex7Ctl.VSFlexGrid GRID2 
            Height          =   3570
            Left            =   -74910
            TabIndex        =   113
            Top             =   405
            Width           =   17205
            _cx             =   30348
            _cy             =   6297
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Simplified Arabic"
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
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   3
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
            AutoSizeMouse   =   0   'False
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
         End
         Begin Threed.SSCommand SSCommand1 
            Height          =   375
            Left            =   -74865
            TabIndex        =   114
            Top             =   0
            Width           =   2415
            _ExtentX        =   4260
            _ExtentY        =   661
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
            Picture         =   "salesTax.frx":0038
            Caption         =   " ÕÊÌ· «ﬂ”Ì· "
            Alignment       =   4
            PictureAlignment=   1
         End
      End
   End
   Begin Threed.SSCommand CMD_ADD_RECEVD 
      Height          =   525
      Left            =   7965
      TabIndex        =   108
      Top             =   630
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   926
      _Version        =   196610
      Font3D          =   3
      CaptionStyle    =   1
      ForeColor       =   128
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "≈÷«›… „‰ «” ·«„ ÿ·»Ì…"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand CMD_SHOINV 
      Height          =   555
      Left            =   7965
      TabIndex        =   107
      Top             =   45
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   979
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "⁄—÷ «·›« Ê—… «·„Ã„⁄…"
      PictureAlignment=   3
   End
   Begin VB.Frame Frame13 
      Caption         =   "»ÕÀ ›Ï «·„” ‰œ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   5130
      RightToLeft     =   -1  'True
      TabIndex        =   98
      Top             =   8955
      Width           =   6090
      Begin VB.TextBox xfilter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3510
         RightToLeft     =   -1  'True
         TabIndex        =   102
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   180
         Width           =   1410
      End
      Begin VB.TextBox xfilter2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   101
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   180
         Width           =   1770
      End
      Begin VB.TextBox xfilter3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   45
         RightToLeft     =   -1  'True
         TabIndex        =   100
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   585
         Width           =   1770
      End
      Begin VB.TextBox xfilter4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         ToolTipText     =   "»ÕÀ"
         Top             =   585
         Width           =   2355
      End
      Begin VB.Label Label4 
         Caption         =   "»«—ﬂÊœ"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5040
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Top             =   270
         Width           =   885
      End
      Begin VB.Label Label31 
         Caption         =   "—ﬁ„ „ÊœÌ· "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1890
         RightToLeft     =   -1  'True
         TabIndex        =   105
         Top             =   180
         Width           =   885
      End
      Begin VB.Label Label32 
         Caption         =   "„’‰⁄"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1980
         RightToLeft     =   -1  'True
         TabIndex        =   104
         Top             =   630
         Width           =   885
      End
      Begin VB.Label Label33 
         Caption         =   "«”„ «·’‰› "
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   103
         Top             =   630
         Width           =   885
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "›« Ê—… „»Ì⁄«  „Ã„⁄… ·› —…"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   4545
      RightToLeft     =   -1  'True
      TabIndex        =   90
      Top             =   -45
      Width           =   3030
      Begin VB.CommandButton CMD_TOINV 
         Caption         =   " —ÕÌ· ›Ê« Ì— «·› —… ·›« Ê—… „»Ì⁄« "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   675
         Width           =   2805
      End
      Begin VB.TextBox XDATESALES1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1575
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   270
         Width           =   1320
      End
      Begin VB.TextBox XDATESALES2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   270
         Width           =   1320
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "FIX REPRICE"
      Height          =   375
      Left            =   4500
      RightToLeft     =   -1  'True
      TabIndex        =   88
      Top             =   1935
      Width           =   1050
   End
   Begin VB.Frame Frame5 
      Height          =   870
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   1035
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
         TabIndex        =   82
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
         TabIndex        =   81
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
         TabIndex        =   80
         Top             =   135
         Width           =   2220
      End
      Begin VB.Label Label27 
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
         TabIndex        =   79
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
         TabIndex        =   78
         Top             =   -270
         Visible         =   0   'False
         Width           =   105
      End
   End
   Begin VB.Frame Frame8 
      Height          =   1005
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   74
      Top             =   8415
      Visible         =   0   'False
      Width           =   3885
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·»Ì⁄ »”⁄— «·„’‰⁄"
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
         Height          =   510
         Left            =   2250
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   270
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin Threed.SSCommand cmdAddItems 
         Height          =   735
         Left            =   90
         TabIndex        =   76
         Top             =   180
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   1296
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "salesTax.frx":2833
         Caption         =   " ”ÃÌ· „ÊœÌ·«  "
         Alignment       =   8
         PictureAlignment=   6
      End
   End
   Begin VB.Frame Frame7 
      Height          =   2310
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   45
      Visible         =   0   'False
      Width           =   2265
      Begin VB.CheckBox xClosed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "«·„” ‰œ „€·ﬁ"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1890
         Width           =   1320
      End
      Begin VB.Frame FRM_CLOSED 
         Height          =   1725
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   135
         Visible         =   0   'False
         Width           =   2040
         Begin Threed.SSCommand cmd_closed 
            CausesValidation=   0   'False
            Height          =   420
            Left            =   90
            TabIndex        =   58
            Top             =   1215
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   741
            _Version        =   196610
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "≈€·«ﬁ «·„” ‰œ"
         End
         Begin Threed.SSCommand cmd_CLOSEDDATE 
            CausesValidation=   0   'False
            Height          =   1005
            Left            =   1035
            TabIndex        =   59
            Top             =   180
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   1773
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
            Picture         =   "salesTax.frx":503B
            Caption         =   "≈€·«ﬁ › —…"
            Alignment       =   8
            PictureAlignment=   6
         End
         Begin Threed.SSCommand cmd_open 
            CausesValidation=   0   'False
            Height          =   1005
            Left            =   90
            TabIndex        =   60
            Top             =   180
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   1773
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
            Picture         =   "salesTax.frx":7B10
            Caption         =   "› Õ › —…"
            Alignment       =   8
            PictureAlignment=   6
         End
      End
      Begin MSComDlg.CommonDialog Common1 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label xF_date 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   1890
         Width           =   1995
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   4140
      Picture         =   "salesTax.frx":A642
      RightToLeft     =   -1  'True
      ScaleHeight     =   2355
      ScaleWidth      =   3210
      TabIndex        =   53
      Top             =   6615
      Width           =   3210
      Begin ImgXCtrl6.ImgXCtrl imgx1 
         Height          =   2175
         Left            =   45
         TabIndex        =   54
         Top             =   45
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   3836
         BackColor       =   16777215
         BorderStyle     =   0
         AutoZoom        =   -1  'True
         SelectionLineType=   4
         Center          =   -1  'True
         ImageBorderThickness=   1
         DoubleBuffer    =   -1  'True
         LicenseUserName =   "mrmind"
         LicenseRegCode  =   "íß“ªª•≤≥Ω≠∞“±≤ß´¥©ÆØOOHH-FAOOYNJB-EQCF6gI"
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   3255
      Left            =   11340
      Picture         =   "salesTax.frx":DD19
      RightToLeft     =   -1  'True
      ScaleHeight     =   3255
      ScaleWidth      =   3885
      TabIndex        =   43
      Top             =   6615
      Width           =   3885
      Begin VB.TextBox xRate 
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
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   1800
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1755
         Width           =   690
      End
      Begin VB.TextBox xDiscount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Left            =   270
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1755
         Width           =   1275
      End
      Begin VB.TextBox xTotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00292952&
         Height          =   600
         Left            =   180
         Locked          =   -1  'True
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2340
         Width           =   2715
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï «·›« Ê—…"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   2385
         Width           =   660
         WordWrap        =   -1  'True
      End
      Begin VB.Label xTotalQuant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   1710
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   90
         Width           =   780
      End
      Begin VB.Label xTotalQuant2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   90
         Width           =   780
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "„— Ã⁄ :"
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   135
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœ ﬁÿ⁄"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   180
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï «·ﬁÌ„…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   585
         Width           =   1005
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ √’‰«›"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   990
         Width           =   960
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï ›« Ê—…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1395
         Width           =   1035
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Œ’„ ›« Ê—…"
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
         Left            =   2655
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1845
         Width           =   900
      End
      Begin VB.Label xTotalItemNoDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   540
         Width           =   2220
      End
      Begin VB.Label xTotalItemDiscount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   945
         Width           =   2220
      End
      Begin VB.Label xTotalItem 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1350
         Width           =   2220
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1575
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1800
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Height          =   780
      Left            =   13230
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   0
      Width           =   4425
      Begin VB.CommandButton cmdExit 
         Height          =   600
         Left            =   45
         Picture         =   "salesTax.frx":11A17
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   135
         Width           =   1095
      End
      Begin VB.CommandButton CmdDel 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   1125
         MaskColor       =   &H00FFFFFF&
         Picture         =   "salesTax.frx":13E83
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdNewInv 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   2205
         MaskColor       =   &H00FFFFFF&
         Picture         =   "salesTax.frx":1671D
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton CmdInform 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   3285
         Picture         =   "salesTax.frx":18CC9
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   135
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "«·—’Ìœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   6615
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   1080
      Width           =   960
      Begin VB.Label xBalance 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   405
         Width           =   780
      End
   End
   Begin VB.Frame Frame3 
      Height          =   690
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   0
      Width           =   2760
      Begin VB.CommandButton cmdSave 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   45
         MaskColor       =   &H00FFFFFF&
         Picture         =   "salesTax.frx":1B49C
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Õ›Ÿ"
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton CmdUndo 
         CausesValidation=   0   'False
         Height          =   465
         Left            =   1485
         MaskColor       =   &H00FFFFFF&
         Picture         =   "salesTax.frx":1D7FF
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   135
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   7380
      Picture         =   "salesTax.frx":1FD78
      RightToLeft     =   -1  'True
      ScaleHeight     =   2355
      ScaleWidth      =   3885
      TabIndex        =   25
      Top             =   6615
      Width           =   3885
      Begin VB.TextBox xRTax 
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
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   450
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   84
         Top             =   855
         Width           =   690
      End
      Begin VB.TextBox xTax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   390
         Left            =   1170
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   855
         Width           =   1320
      End
      Begin VB.Label xRTax2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   450
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   1305
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "≈Ã„«·Ï »⁄œ Œ’„"
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
         Left            =   2565
         RightToLeft     =   -1  'True
         TabIndex        =   87
         Top             =   315
         Width           =   1215
      End
      Begin VB.Label xTotDisc 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   465
         Left            =   270
         RightToLeft     =   -1  'True
         TabIndex        =   86
         Top             =   225
         Width           =   2220
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   85
         Top             =   900
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "«·÷—Ì»… : "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2745
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   945
         Width           =   840
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1680
      Left            =   9990
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   720
      Width           =   7665
      Begin VB.CommandButton cmd_edit 
         Caption         =   "..."
         Height          =   285
         Left            =   4995
         RightToLeft     =   -1  'True
         TabIndex        =   115
         Top             =   180
         Width           =   330
      End
      Begin VB.TextBox XDISC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   360
         Left            =   90
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   94
         Top             =   875
         Width           =   960
      End
      Begin VB.TextBox xpo_no 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2970
         MaxLength       =   50
         RightToLeft     =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   135
         Width           =   1275
      End
      Begin VB.CommandButton cmdClient2 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   330
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   -315
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.TextBox xDoc_no2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5355
         Locked          =   -1  'True
         MaxLength       =   6
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   180
         Width           =   1185
      End
      Begin VB.TextBox xCode3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5355
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1800
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox xCode2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         IMEMode         =   3  'DISABLE
         Left            =   5355
         MaxLength       =   15
         PasswordChar    =   "*"
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   -330
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.TextBox xNotes 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2970
         MaxLength       =   75
         RightToLeft     =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1260
         Width           =   3570
      End
      Begin VB.TextBox xCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5355
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   530
         Width           =   1185
      End
      Begin VB.TextBox xDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   900
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   180
         Width           =   1230
      End
      Begin MSDataListLib.DataCombo xStore 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Top             =   535
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   2070
         Visible         =   0   'False
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox xDoc_No 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7425
         MaxLength       =   12
         RightToLeft     =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   765
         Visible         =   0   'False
         Width           =   195
      End
      Begin MSDataListLib.DataCombo xMan 
         Height          =   330
         Left            =   2970
         TabIndex        =   4
         Top             =   900
         Width           =   3570
         _ExtentX        =   6297
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
      Begin VB.Label Label7 
         Caption         =   "—ﬁ„ ÿ·»Ì… "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4230
         RightToLeft     =   -1  'True
         TabIndex        =   109
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "«·»«∆⁄ :"
         BeginProperty Font 
            Name            =   "Simplified Arabic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   810
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "—ﬁ„ „” ‰œ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   225
         Width           =   885
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê”„ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   2160
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "«·„Ê—œ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1845
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label xcodeDesca3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1800
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "⁄„·«¡ ﬂ—Ê  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   -270
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label xCodeDesca2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   3330
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   -315
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label xtime 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   780
      End
      Begin VB.Label xusername 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   315
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1260
         Width           =   2025
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„” Œœ„ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1305
         Width           =   705
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "„·«ÕŸ«  :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1305
         Width           =   660
      End
      Begin VB.Label xCodeDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   2970
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   540
         Width           =   2355
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«· «—ÌŒ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "«·„Œ“‰ :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2220
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   630
         Width           =   540
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         Caption         =   "«·⁄„Ì· :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   585
         Width           =   510
      End
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   -3240
      Top             =   6660
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   -855
      Top             =   6885
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
   Begin MSAdodcLib.Adodc data2 
      Height          =   330
      Left            =   -4770
      Top             =   8145
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
   Begin MSAdodcLib.Adodc data4 
      Height          =   330
      Left            =   -3060
      Top             =   6975
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
   Begin MSAdodcLib.Adodc data5 
      Height          =   330
      Left            =   -1755
      Top             =   6885
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
      Align           =   2  'Align Bottom
      Height          =   150
      Left            =   0
      TabIndex        =   21
      Top             =   9945
      Visible         =   0   'False
      Width           =   17760
      _ExtentX        =   31327
      _ExtentY        =   265
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   2400
      Left            =   630
      Picture         =   "salesTax.frx":22F9E
      RightToLeft     =   -1  'True
      ScaleHeight     =   2400
      ScaleWidth      =   3165
      TabIndex        =   27
      Top             =   3150
      Visible         =   0   'False
      Width           =   3165
   End
   Begin MSAdodcLib.Adodc DATA11 
      Height          =   330
      Left            =   7560
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
   Begin VB.Frame Frame6 
      Height          =   645
      Left            =   810
      RightToLeft     =   -1  'True
      TabIndex        =   63
      Top             =   6570
      Width           =   3300
      Begin Threed.SSCommand cmdLast 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   90
         TabIndex        =   64
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "salesTax.frx":26675
         Caption         =   "«ŒÌ—"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "salesTax.frx":28845
      End
      Begin Threed.SSCommand cmdNext 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   870
         TabIndex        =   65
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "salesTax.frx":2A98D
         Caption         =   "·«Õﬁ "
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "salesTax.frx":2CB55
      End
      Begin Threed.SSCommand cmdPrevious 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   1620
         TabIndex        =   66
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "salesTax.frx":2ECA4
         Caption         =   "”«»ﬁ"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "salesTax.frx":30E84
      End
      Begin Threed.SSCommand cmdFirst 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   2430
         TabIndex        =   67
         Top             =   135
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   741
         _Version        =   196610
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
         Picture         =   "salesTax.frx":32FDF
         Caption         =   "√Ê·"
         Alignment       =   4
         PictureAlignment=   9
         PictureDisabledFrames=   1
         PictureDisabled =   "salesTax.frx":3519B
      End
   End
   Begin Threed.SSCommand CMD_PRINT 
      Height          =   465
      Left            =   2475
      TabIndex        =   68
      Top             =   45
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   820
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "salesTax.frx":372EA
      Caption         =   "ÿ»«⁄… «·›« Ê—…  "
      Alignment       =   1
      PictureAlignment=   3
   End
   Begin MSAdodcLib.Adodc data12 
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
   Begin Threed.SSCommand cmd_excel 
      Height          =   420
      Left            =   2475
      TabIndex        =   89
      Top             =   1935
      Width           =   1965
      _ExtentX        =   3466
      _ExtentY        =   741
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
      Picture         =   "salesTax.frx":396F9
      Caption         =   " ÕÊÌ· «ﬂ”Ì· "
      Alignment       =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmd_addexel 
      Height          =   435
      Left            =   7965
      TabIndex        =   96
      Top             =   1170
      Width           =   1875
      _ExtentX        =   3307
      _ExtentY        =   767
      _Version        =   196610
      Font3D          =   3
      CaptionStyle    =   1
      ForeColor       =   128
      BackColor       =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "«÷«›… „‰ «ﬂ”Ì·"
      ButtonStyle     =   2
   End
   Begin Threed.SSCommand CMD_PRINTREC 
      Height          =   465
      Left            =   2475
      TabIndex        =   97
      Top             =   540
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   820
      _Version        =   196610
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ÿ»«⁄… «–‰  ”·Ì„ „Œ“‰"
      Alignment       =   1
      PictureAlignment=   3
   End
   Begin MSAdodcLib.Adodc DATA22 
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
   Begin Threed.SSCommand cmdSign 
      Height          =   510
      Left            =   855
      TabIndex        =   118
      Top             =   7290
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   900
      _Version        =   196610
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
      Picture         =   "salesTax.frx":3BEF4
      Caption         =   "«—”«· ›« Ê—… «·Ìﬂ —Ê‰Ì…  ›’Ì·Ì…"
      PictureAlignment=   9
      PictureDisabledFrames=   1
      PictureDisabled =   "salesTax.frx":3E435
   End
   Begin Threed.SSCommand cmdSignGroup 
      Height          =   510
      Left            =   855
      TabIndex        =   119
      Top             =   7785
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   900
      _Version        =   196610
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
      Picture         =   "salesTax.frx":40866
      Caption         =   "«—”«· ›« Ê—… «·Ìﬂ —Ê‰Ì… «Ã„«·Ì…"
      PictureAlignment=   9
      PictureDisabledFrames=   1
      PictureDisabled =   "salesTax.frx":42DA7
   End
   Begin VB.Label Label9 
      Caption         =   "›« Ê—… ≈·ﬂ —Ê‰Ì…"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9315
      RightToLeft     =   -1  'True
      TabIndex        =   117
      Top             =   1800
      Width           =   660
   End
End
Attribute VB_Name = "salestAX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim temptable As New ADODB.Recordset
Dim lManSales As Boolean
Public sDoc_no As String, lSave As Boolean
Dim cFilter As String
Dim tCard As ADODB.Recordset
Dim nLeftMargin, nCardHeight, nPageWidth, nRightMargin
Dim nRound As Long, clist1 As String, cList2 As String
Dim CardTable As ADODB.Recordset
Public bRetvalue As Boolean
Dim cDefBox As String, cDefClient As String, cDefClientDesca As String, cDefCasher As String, cDefStore As String, cdefman As String
Dim oSearchItem As New Search3, oSearchDoc   As New Search3, oSearchCl1 As New Search3, bMarket As Boolean
Dim oSearchCl2 As New Search3, oSearchCl3 As New Search3
Dim bedit As Boolean
Dim cFile As String, cCodeDesca As String
Dim con As New ADODB.Connection
Dim formMode, dDateLast As String
Public myPublic As Integer, lIsRet As Boolean
Const LoadMode = 0, DefineMode = 1
Sub ItemsLookup()
ItemsLookupAll Me, oSearchItem
End Sub
Private Function myreplace(Optional nRow As Long = -1, Optional bModel As Boolean) As Boolean
Dim aInsert(21, 1)
aInsert(0, 0) = "Doc_No"
aInsert(0, 1) = addstring(xDoc_No.text)

aInsert(1, 0) = "code"
aInsert(1, 1) = addstring(xCode.text)

aInsert(2, 0) = "[Date]"
aInsert(2, 1) = addDate(xDate.text)

aInsert(3, 0) = "store"
aInsert(3, 1) = addstring(xStore.BoundText)

aInsert(4, 0) = "[Notes]"
aInsert(4, 1) = addstring(xNotes.text)

aInsert(5, 0) = "Rate"
aInsert(5, 1) = Val(xRate.text)

aInsert(6, 0) = "Discount"
aInsert(6, 1) = Val(xDiscount.text)

aInsert(7, 0) = "Cash"
aInsert(7, 1) = 0

aInsert(8, 0) = "late"
aInsert(8, 1) = Val(xTotal.text)

aInsert(9, 0) = "MAN"
If xMan.BoundText = "" Then xMan.BoundText = "001"
aInsert(9, 1) = addstring(xMan.BoundText)

aInsert(10, 0) = "ISINVOICE"
aInsert(10, 1) = 1

aInsert(11, 0) = "TAX"
aInsert(11, 1) = Val(xTax.text)

aInsert(12, 0) = "RTAX"
aInsert(12, 1) = Val(xRTax.text)

aInsert(16, 0) = "USER_IP"
aInsert(16, 1) = addstring(cIpName)

aInsert(17, 0) = "DATESALES1"
aInsert(17, 1) = addDate(XDATESALES1.text)

aInsert(18, 0) = "DATESALES2"
aInsert(18, 1) = addDate(XDATESALES2.text)

aInsert(19, 0) = "ISRET"
aInsert(19, 1) = IIf(lIsRet, 1, 0)

aInsert(20, 0) = "PO_NO"
aInsert(20, 1) = addstring(xpo_no.text)

aInsert(21, 0) = "INV_TAX"
aInsert(21, 1) = addstring(XINV_TAX.text)

con.BeginTrans
''On Error GoTo myerror
If xDoc_No.Tag = DefineMode Then
    aInsert(13, 0) = "userName"
    aInsert(13, 1) = addstring(xusername.Caption)

    aInsert(14, 0) = "TIME1"
    aInsert(14, 1) = addDate("GETDATE()")

    aInsert(0, 1) = addstring(xDoc_No.text)
    con.Execute CreateInsert(aInsert, "FILE6_20H")
Else
    aInsert(13, 0) = "userName2"
    aInsert(13, 1) = addstring(xUserName2.Caption)

    aInsert(14, 0) = "TIME2"
    aInsert(14, 1) = addDate("GETDATE()")
    
    con.Execute CreateUpdate(aInsert, "FILE6_20H", " where doc_no = " & addstring(xDoc_No.text))
End If


    If bModel Then
        myreplaceGrdModel
    Else
        myreplaceGrd nRow
    End If

'con.Execute " UPDATE FILE6_20 SET MAN = " & MyParn(xMan.BoundText) & " WHERE DOC_NO = " & MyParn(xDoc_No.Text)
con.CommitTrans
myreplace = True

Exit Function
myerror:
prog1.Visible = False
MsgBox Err.Description
con.RollbackTrans
Err.Clear
xDoc_No.Tag = 1
End Function
Sub myProc()
''On Error GoTo myerror
If ActiveControl.Name = grid1.Name Then
    Dim bNew As Boolean
'    bNew = grid1.Row = grid1.Rows - 1
    If ActiveControl.Name = grid1.Name Then
        grid1.TextMatrix(grid1.Row, 1) = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    ElseIf ActiveControl.Name = LOOKPH.Name Then
        grid1.TextMatrix(grid1.Row, 1) = ItemLookPh.grid1.TextMatrix(ItemLookPh.grid1.Row, 0)
    ElseIf ActiveControl.Name = LOOKPH2.Name Then
        grid1.TextMatrix(grid1.Row, 1) = ItemLookPh2.grid1.TextMatrix(ItemLookPh2.grid1.Row, 0)
    End If
    GrdDesc grid1.TextMatrix(grid1.Row, 1), grid1.Row
    grid1_AfterEdit grid1.Row, grid1.Col
    If Not bNew Then
        If ActiveControl.Name = grid1.Name Then Unload oSearchItem Else Unload ItemLookPh
        CellPos 13, grid1.Row, grid1.Col
    End If
ElseIf ActiveControl.Name = CmdInform.Name Then
    xDoc_No.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    xDate.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 2)
    Unload oSearchDoc
    myUndo
ElseIf ActiveControl.Name = xCode.Name Then
    ActiveControl.text = oSearchCl1.grid1.TextMatrix(oSearchCl1.grid1.Row, 0)
    xCode_LostFocus
    Unload oSearchCl1
ElseIf ActiveControl.Name = xMan.Name Then
    xMan.BoundText = oSearchCl2.grid1.TextMatrix(oSearchCl1.grid1.Row, 0)
    Unload oSearchCl1
ElseIf ActiveControl.Name = xpo_no.Name Then
    xpo_no.text = oSearchDoc.grid1.TextMatrix(oSearchDoc.grid1.Row, 0)
    Unload oSearchDoc
End If
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub c_Click()
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me

End Sub

Private Sub CMD_ADD_RECEVD_Click()
If xClosed.Value = O And xpo_no.text <> "" Then
    If Val(GetDesca(" SELECT COUNT(*) FROM FILE6_20 WHERE DOC_NO =   " & MyParn(xDoc_No.text), con) & "") = 0 Then
        If Not MYVALID Then Exit Sub
        If MsgBox(" ≈÷«›… «’‰«› ≈” ·«„ „‰ «·ÿ·»Ì… ··›« Ê—… ", vbYesNo + vbDefaultButton2) = vbYes Then
            cmdSave_Click
            con.Execute " INSERT INTO FILE6_20 (DOC_NO , MAN , ITEM, QUANT, PRICE ) SELECT  " & addstring(xDoc_No) & " , " & addstring(xMan.BoundText) & " ,  FILE6_52.ITEM, SUM(FILE6_52.QUANT) AS QUANT , (SELECT TOP 1 PRICE FROM FILE6_51 WHERE FILE6_51.DOC_NO = " & MyParn(xpo_no.text) & " AND FILE6_52.ITEM = FILE6_51.ITEM ) FROM FILE6_52 INNER JOIN FILE6_52H ON FILE6_52.DOC_NO = FILE6_52H.DOC_NO WHERE FILE6_52H.PO_NO = " & MyParn(xpo_no.text) & " GROUP BY FILE6_52.ITEM", nRec
            con.Execute " UPDATE       FILE6_20 SET  PRICE_C = FILE1_10.PRICE, PRICE_C2 = FILE6_20.PRICE FROM            FILE1_10 INNER JOIN FILE6_20 ON FILE1_10.ITEM = FILE6_20.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.text)
            MsgBox "  „ «÷«›… ⁄œœ ”Ã·  " & nRec
            myUndo
        End If
    End If
End If
End Sub
Private Sub cmd_addexel_Click()
    If xStore.BoundText = "" Then
        MsgBox " ≈Œ Ì«— «·„Œ“‰ "
        Exit Sub
    End If
    If xCode.text = "" Then
        MsgBox " ≈Œ Ì«— «·⁄„Ì· "
        Exit Sub
    End If
    If Not myreplace Then Exit Sub
    AddFromExel
    Inform "  „ «÷«›… «·«’‰«› "
    myUndo
End Sub

Private Sub cmd_closed_Click()
    If CardTable!ISCLOSED Then
        con.Execute " update file6_20h set isclosed = 0 where doc_no = " & MyParn(xDoc_No.text)
    Else
        con.Execute " update file6_20h set isclosed = 1 where doc_no = " & MyParn(xDoc_No.text)
    End If
    myUndo
End Sub
Private Sub cmd_CLOSEDDATE_Click()
Dim DDate1 As Date, DDate2 As Date
DDate1 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  „‰  «—ÌŒ", "1-1"), "", "1-1-2000")
DDate2 = TurnValue(InputBox("", "≈€·«ﬁ „” ‰œ«  Õ Ï  «—ÌŒ", Format(Date, "DD-MM-YYYY")), "", "1-1-2000")
If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
If IsDate(DDate1) And IsDate(DDate2) Then
    If MsgBox("”Ê› Ì „ ≈€·«ﬁ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
'        con.ConnectionTimeout = 600
        con.Execute " update file6_20h set   isclosed = 1 where ISINVOICE = 1 AND DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
        MsgBox " „ ≈€·«ﬁ „” ‰œ«  «·› —…" & nRec
        myUndo
    End If
End If
End Sub

Private Sub cmd_edit_Click()
Dim cNewDoc As String, cNewDoc2 As String
If MsgBox("  ⁄œÌ· —ﬁ„ «·›« Ê—… «·÷—Ì»Ì»…  ", vbYesNo + vbDefaultButton2) = vbYes Then
    cNewDoc = InputBox("", "«·—ﬁ„ «·ÃœÌœ ")
    cNewDoc2 = RetZero("1" & RetZero(cNewDoc, 11))
    If GetDesca("SELECT DOC_NO FROM FILE6_20H WHERE DOC_NO = " & MyParn(cNewDoc2), con) <> "" Then
        MsgBox " —ﬁ„ „”Ã· „‰ ﬁ»· "
    Else
        If MsgBox(" ”Ê› Ì „  ⁄œÌ· ««·› «Ê—… «·Õ«·Ì… ··—ﬁ„ «·ÃœÌœ " & cNewDoc, vbYesNo + vbDefaultButton2) = vbYes Then
            con.Execute " UPDATE FILE6_20H SET DOC_NO = " & addstring(cNewDoc2) & " WHERE DOC_NO = " & addstring(xDoc_No.text)
              
            xDoc_No.text = cNewDoc2
            myUndo
        
        End If
        
    End If
End If
End Sub

Private Sub cmd_excel_Click()
    
    grid1.ColHidden(12) = False
    grid1.ColHidden(13) = False
    grid1.ColHidden(15) = False
    grid1.ColHidden(16) = False
    
    grid1.ColHidden(grid1.Cols - 2) = False
    grid1.ColHidden(grid1.Cols - 3) = False
    grid1.ColHidden(grid1.Cols - 1) = True
    ToFileExel2 grid1, , , , , 1.1, , , , , , Me
    grid1.ColHidden(grid1.Cols - 2) = True
    grid1.ColHidden(grid1.Cols - 3) = True

    grid1.ColHidden(12) = True
    grid1.ColHidden(13) = True
    grid1.ColHidden(15) = True
    grid1.ColHidden(16) = True

End Sub

Private Sub cmd_open_Click()
    Dim DDate1 As Date, DDate2 As Date
    DDate1 = TurnValue(InputBox("", "› Õ  „” ‰œ«  „‰  «—ÌŒ", ""), "", "1-1-2000")
    DDate2 = TurnValue(InputBox("", "› Õ „” ‰œ«  Õ Ï  «—ÌŒ", ""), "", "1-1-2000")
    If Not IsDate(DDate2) And IsDate(DDate1) Then DDate2 = dDate
    If IsDate(DDate1) And IsDate(DDate2) Then
        If MsgBox("”Ê› Ì „ › Õ „” ‰œ«  «·› —…", vbOKCancel) = vbOK Then
            con.Execute " update file6_20h set ISINVOICE = 1 AND  isclosed = 0 where DATE > = " & DateSq(DDate1) & " AND DATE <= " & DateSq(DDate2), nRec
            MsgBox " „ › Õ „” ‰œ«  «·› —…" & nRec
            myUndo
        End If
    End If
End Sub
Private Sub CMD_PRINT_Click()
    If IsDate(XDATESALES1.text) And IsDate(XDATESALES2.text) Then
        If MsgBox("  Ã„Ì⁄ ›Ê« Ì— «·› —… ··⁄„Ì· ·›« Ê—… „»Ì⁄«  ÷—Ì»Ì»… ", vbYesNo + vbDefaultButton2) = vbYes Then
            con.Execute " UPDATE FILE6_20H SET INV_NO = NULL WHERE INV_NO =  " & MyParn(xDoc_No.text)
            con.Execute " UPDATE FILE6_20H SET INV_NO = " & addstring(xDoc_No.text) & " WHERE ISINVOICE = 0 AND CODE = " & MyParn(xCode.text) & " AND DATE >= " & DateSq(XDATESALES1.text) & " AND DATE <= " & DateSq(XDATESALES2.text)
        End If
        doprint_SUBINV
    Else
        doprint_doc
    End If
End Sub
Private Sub cmd_tsal_Click()
    TDaySal.Show
End Sub

Private Sub CMD_PRINTREC_Click()
    doprint_DOC_REC
End Sub

Private Sub CMD_SHOINV_Click()
    If IsDate(XDATESALES1.text) And IsDate(XDATESALES2.text) Then
        xClosed.Value = 1
        cmd_closed.BackColor = &H8080FF
        cmd_closed.Caption = "› Õ „” ‰œ"
        xClosed.Value = ssCBChecked
        xClosed.Visible = True
        cmdSave.Enabled = False
        CmdDel.Enabled = False
        ShowInv_Grid
        ShowInv_Grid2
    End If
End Sub

Private Sub CMD_TOINV_Click()
If MsgBox("  Ã„Ì⁄ ›Ê« Ì— «·› —… ··⁄„Ì· ·›« Ê—… „»Ì⁄«  ÷—Ì»Ì»… ", vbYesNo + vbDefaultButton2) = vbYes Then
    con.Execute " UPDATE FILE6_20H SET INV_NO = NULL WHERE INV_NO =  " & MyParn(xDoc_No.text)
    con.Execute " UPDATE FILE6_20H SET INV_NO = " & addstring(xDoc_No.text) & " WHERE ISINVOICE = 0 AND CODE = " & MyParn(xCode.text) & " AND DATE >= " & DateSq(XDATESALES1.text) & " AND DATE <= " & DateSq(XDATESALES2.text)
    AddLod_Data cusername, 0, "  —ÕÌ· ›Ê« Ì— „Ã„⁄… " & Me.Caption, con, , XDATESALES1.text, , XDATESALES2.text

End If
End Sub
Private Sub cmdAddItems_Click()
    If grid1.Row = 0 Then grid1.Row = grid1.Rows - 1
    Set salesModel.myForm = Me
    salesModel.nColItem = 0
    salesModel.nColQuant = 10
    salesModel.strItem = grid1.TextMatrix(grid1.Row, 0)
    salesModel.strStore = xStore.BoundText
    salesModel.bshowBal = True
    salesModel.Show 1
End Sub
Private Sub CmdDel_Click()
On Error GoTo myerror
If MsgBox("Õ–› «·„” ‰œ »«·ﬂ«„·  ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
    
    con.BeginTrans
    If IsDate(XDATESALES1.text) Then con.Execute "UPDATE FILE6_20H SET INV_NO = NULL WHERE INV_NO = " & MyParn(xDoc_No.text)
    
    cStr1 = " INSERT INTO COPY6_20 (DOC_NO, ITEM, QUANT,ID2, PRICE, TOTAL, DISCOUNT,  cost, PRICE_C, row, MAN, S_OKAZ, ISDISC2_1, DATE                          , DELUSER                       , DEL_IP) " & _
             " SELECT               DOC_NO, ITEM, QUANT,ID, PRICE, TOTAL, DISCOUNT,  cost, PRICE_C, row, MAN, S_OKAZ, ISDISC2_1, " & addDate(xDate.text) & "   , " & addstring(cusername) & "   , " & addstring(cIpName) & _
            "   From FILE6_20 WHERE DOC_NO = " & addstring(xDoc_No.text)
    con.Execute cStr1
    
    con.Execute "Delete  From FILE6_20 where Doc_No = " & MyParn(xDoc_No.text)
    con.Execute "Delete  From FILE6_20H where Doc_No = " & MyParn(xDoc_No.text)
    con.CommitTrans
    
    AddLod_Data cusername, 2, " Õ–› " & Me.Caption, con, xDoc_No.text, xDate.text, , xCodeDesca.Caption
    
    openCardTable
    
    
    CmdNewInv_Click
End If
Exit Sub
myerror:
MsgBox Err.Description
con.RollbackTrans
Err.Clear
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub CmdInform_Click()
    CardLookup
End Sub
Private Sub CmdNewInv_Click()
myDefine
On Error Resume Next
xDoc_No.SetFocus
Err.Clear
'grid1.SetFocus
End Sub
Private Sub cmdSave_Click()
    If Not MYVALID Then Exit Sub
    mysave
    AddLod_Data cusername, 1, " Õ›Ÿ " & Me.Caption, con, xDoc_No.text, xDate.text, , xCodeDesca.Caption

End Sub
Private Sub cmdSign_Click()
AddLod_Data cusername, 0, " «—”«· ›« Ê—…  ›’Ì·Ì… " & Me.Caption, con, xDoc_No.text, xDate.text, , xCodeDesca.Caption

Me.MousePointer = vbHourglass
cmdSign.Enabled = False
cmdSignGroup.Enabled = False
SendInvoice
Handlecontrols xDoc_No.Tag
Me.MousePointer = vbNormal
End Sub
Private Sub cmdSignGroup_Click()
AddLod_Data cusername, 0, " «—”«· ›« Ê—… ≈Ã„«·Ì… " & Me.Caption, con, xDoc_No.text, xDate.text, , xCodeDesca.Caption

Me.MousePointer = vbHourglass
cmdSign.Enabled = False
cmdSignGroup.Enabled = False
SendInvoice True
Me.MousePointer = vbNormal
End Sub
Private Function SendInvoice(Optional bGroup As Boolean = False) As Boolean
Dim cError As String
bsuccess = SubmitDocument(xDoc_No.text, con, cError, Me, IIf(lIsRet, "C", "I"), bGroup)
If bsuccess Then
    eInform " „ «—”«· «·›« Ê—… «·«·Ìﬂ —Ê‰Ì… »‰Ã«Õ"
    openCardTable
    myUndo
ElseIf cError <> "" Then
    MsgBox cError
Else
    eInform "Œÿ« «À‰«¡ «—”«· «·›« Ê—…"
End If
End Function

Private Sub CmdUndo_Click()
    myUndo
    grid1.ShowCell grid1.Rows - 1, 1
    grid1.Select grid1.Rows - 1, 1
    On Error Resume Next
    xMan.SetFocus
    'grid1.SetFocus
    Err.Clear
End Sub
Private Sub Command1_Click()
If InputBox("", "654321") = "654321" Then
    If MsgBox("  €Ì— «”⁄«— «·»Ì⁄ Õ”» ‰”»… Œ’„ «·⁄„Ì· ⁄·Ï ”⁄— «·„” Â·ﬂ ﬁ»· «·÷—Ì»… ", vbYesNo + vbDefaultButton2) = vbYes Then
        With grid1
            For nRow = 1 To .Rows - 2
                Dim loctable As ADODB.Recordset
                Set loctable = ItemFind(.TextMatrix(nRow, 1), con)
                If Not (loctable.EOF And loctable.BOF) Then
                    nPrice = Round((loctable!price / 1.14) * ((100 - Val(XDISC.text)) / 100), 2)
                    con.Execute " UPDATE FILE6_20 SET PRICE =  " & nPrice & " WHERE DOC_NO = " & MyParn(xDoc_No.text) & " AND ITEM = " & Val(.TextMatrix(nRow, 1))
                End If
                Me.Caption = nRow
            Next nRow
            Inform "  „  ÕœÌÀ «·«”⁄«— "
            myloadgrd
        End With
    End If
End If
End Sub

Private Sub Command2_Click()
Dim pkcs11 As New ChilkatPkcs11

' Set your shared lib path -- either a full path to the DLL, .so, or .dylib,
' or just the filename if the driver is in a location that can be automatically found. (Such as in C:\Windows\System32)
pkcs11.SharedLibPath = "C:\Windows\System32\eps2003csp11.dll"

Dim success As Long
success = pkcs11.Initialize()
If (success = 0) Then
    Debug.Print pkcs11.LastErrorText
    Exit Sub
End If

' Pass -1 for the slotID to open a session on the first non-empty slot.
Dim slotID As Long
slotID = -1

' Open a session.
Dim readWrite As Long
readWrite = 1
success = pkcs11.OpenSession(slotID, readWrite)
If (success = 0) Then
    Debug.Print pkcs11.LastErrorText
    Exit Sub
End If

' Make it an authenticated session by calling Login.
'
' If we don't authenticate, then we won't be able to see the private keys, and thus
' we won't know which certificates have an associated private key stored on the smart card.

' The smart card PIN is passed to the Login method.
' userType 1 indicates a "Normal User".
Dim userType As Long
userType = 1
Dim pin As String
pin = "44027605"
success = pkcs11.login(userType, pin)
If (success = 0) Then
    Debug.Print pkcs11.LastErrorText
    success = pkcs11.CloseSession()
    Exit Sub
End If

' --------------------------------------------------------------------------
' The FindCert method can find a particular certificate in a number of different ways.
' I'll demonstrate some common ways..

' 1) In many cases you'll be working with a smart card that contains one certificate that is to be used for
' signing, and it is the certificate that is associated with the private key also stored on the smart card (or USB token).
' There may be other certificates on the card, but these are the issuer certificates in the chain of authentication.
' You're just interested in getting the certificate with the private key.
' You can do it like this:
Dim cert As New ChilkatCert
' Pass the keyword "privateKey" in the 1st argument, and an empty string in the 2nd arg.
' This returns the 1st certificate that has a private key.
'success = pkcs11.FindCert("privateKey", "", cert)
'If (success = 1) Then
'    Debug.Print "Cert with private key: " & cert.SubjectCN
'Else
'    Debug.Print "No certificates having a private key were found."
'End If

' 2) Find a certificate by the Subject Common Name
success = pkcs11.FindCert("subjectCN", "Egypt Trust)", cert)
If (success = 1) Then
    Debug.Print "Found: " & cert.SubjectCN
Else
    Debug.Print "Not found: " & cert.SubjectCN
End If

' 3) Find a certificate by hex serial number
success = pkcs11.FindCert("serial", "2B9F271600580037", cert)
If (success = 1) Then
    Debug.Print "Found: " & cert.SerialNumber
Else
    Debug.Print "Not found: " & cert.SerialNumber
End If

' 4) Find a certificate by hex serial number and Issuer Common Name
'success = pkcs11.FindCert("serial:issuerCN", "5087bf1feda006af54a02f23a851104948acc26f:Ibanity Production Third Party Application CA", cert)
success = pkcs11.FindCert("issuerCN", "Egypt Trust", cert)
If (success = 1) Then
    Debug.Print "Found: " & cert.SubjectCN
Else
    Debug.Print "Not found: " & cert.SubjectCN
End If

MsgBox cert.IssuerDN

' --------------------------------------------------------------------------

' Revert to an unauthenticated session by calling Logout.
success = pkcs11.Logout()
If (success = 0) Then
    Debug.Print pkcs11.LastErrorText
    success = pkcs11.CloseSession()
    Exit Sub
End If

' When finished, close the session.
' It is important to close the session (memory leaks will occur if the session is not properly closed).
success = pkcs11.CloseSession()
If (success = 0) Then
    Debug.Print pkcs11.LastErrorText
    Exit Sub
End If

Debug.Print "Success."
End Sub
Private Sub Command3_Click()

Dim pkcs11 As New ChilkatPkcs11
Dim crypt As New ChilkatCrypt2
Dim cert As New ChilkatCert
Dim success As Long

Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

' Set your shared lib path -- either a full path to the DLL, .so, or .dylib,
' or just the filename if the driver is in a location that can be automatically found. (Such as in C:\Windows\System32)
pkcs11.SharedLibPath = "C:\Windows\System32\eps2003csp11.dll"

success = pkcs11.Initialize()
If (success = 0) Then
    Debug.Print pkcs11.LastErrorText
    Exit Sub
End If

' Pass -1 for the slotID to open a session on the first non-empty slot.
Dim slotID As Long
slotID = -1

' Open a session.
Dim readWrite As Long
readWrite = 1
success = pkcs11.OpenSession(slotID, readWrite)
If (success = 0) Then
    Debug.Print pkcs11.LastErrorText
    Exit Sub
End If

' Make it an authenticated session by calling Login.
'
' If we don't authenticate, then we won't be able to see the private keys, and thus
' we won't know which certificates have an associated private key stored on the smart card.

' The smart card PIN is passed to the Login method.
' userType 1 indicates a "Normal User".
Dim userType As Long
userType = 1
Dim pin As String
pin = "44027605"
success = pkcs11.login(userType, pin)
If (success = 0) Then
    Debug.Print pkcs11.LastErrorText
    success = pkcs11.CloseSession()
    Exit Sub
End If

' Call FindAllCerts to find all certificates on the smart card or USB token.
success = pkcs11.FindAllCerts()
If (success = 0) Then
    Debug.Print pkcs11.LastErrorText
    success = pkcs11.CloseSession()
    Exit Sub
End If

' The NumCerts property is set by FindAllCerts.
Dim numCerts As Long
numCerts = pkcs11.numCerts
Debug.Print "Number of certs: " & numCerts

Dim privateKeyNote As String
Dim i As Long
i = 0
Do While i < numCerts
    success = pkcs11.GetCert(i, cert)
    privateKeyNote = ""
    If (cert.HasPrivateKey() = 1) Then
        privateKeyNote = "(has private key)"
    End If

    Debug.Print i & ": " & privateKeyNote & " " & cert.SubjectDN
    Debug.Print i & ": issuer: " & cert.IssuerDN
    Debug.Print "----"
    i = i + 1
Loop

Dim cmsOptions As New ChilkatJsonObject
' Setting "DigestData" causes OID 1.2.840.113549.1.7.5 (digestData) to be used.
success = cmsOptions.UpdateBool("DigestData", 1)
success = cmsOptions.UpdateBool("OmitAlgorithmIdNull", 1)
crypt.cmsOptions = cmsOptions.Emit()

' The CadesEnabled property applies to all methods that create CMS/PKCS7 signatures.
' To create a CAdES-BES signature, set this property equal to true.
crypt.CadesEnabled = 1

crypt.HashAlgorithm = "sha256"

Dim jsonSigningAttrs As New ChilkatJsonObject
success = jsonSigningAttrs.UpdateInt("contentType", 1)
success = jsonSigningAttrs.UpdateInt("signingTime", 1)
success = jsonSigningAttrs.UpdateInt("messageDigest", 1)
success = jsonSigningAttrs.UpdateInt("signingCertificateV2", 1)
crypt.SigningAttributes = jsonSigningAttrs.Emit()

' By default, all the certs in the chain of authentication are included in the signature.
' If desired, we can choose to only include the signing certificate:
crypt.includeCertChain = 0

' Make sure we sign the utf-8 byte representation of the JSON string
crypt.Charset = "utf-8"

' Create the CAdES-BES signature.
Dim textToSign As String
textToSign = Serialize(Clipboard.GetText)

crypt.EncodingMode = "base64"
Dim sigBase64 As String
sigBase64 = crypt.SignStringENC(textToSign)
If (crypt.LastMethodSuccess = 0) Then
    MsgBox crypt.LastErrorText
    Exit Sub
End If

Debug.Print "Base64 signature:"
MsgBox sigBase64


Debug.Print "Success"
End Sub
Private Sub Form_Activate()
On Error Resume Next
xDoc_No.SetFocus
Err.Clear
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If xClosed.Value = 1 Then Exit Sub
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is DataCombo Then SendKeys "{TAB}"
End If
If KeyAscii = 27 Then
    On Error Resume Next
    xTotal.SetFocus
    Err.Clear
End If
End Sub
Private Sub Form_Load()
Dim obj As New ChilkatGlobal
success = obj.UnlockBundle("MABFTH.CB4082022_DqFFZRYK0Rmf")

Dim ImgXTmp As ImgX
Set ImgXTmp = New ImgX

cmd_edit.Visible = bSupermode And lIsRet = 0
bedit = True
openCon con
Set ImgXTmp = Nothing
Me.Caption = dSalesDate

clist1 = StrList("SELECT CODE , DESCA FROM FACT", con)
cList2 = StrList("SELECT CODE , DESCA FROM file6_25 ", con)
nRound = 2
Set data12.Recordset = myRecordSet("SELECT * FROM FILE6_25", con)
Set xMan.RowSource = data12
xMan.ListField = "Desca"
xMan.BoundColumn = "Code"


FRM_CLOSED.Visible = bopt2
Frame7.Visible = bopt2
Dim aRet As Variant
lManSales = False

cMoveName = "„»Ì⁄« "
Me.Caption = IIf(lIsRet, "›Ê« Ì— „— Ã⁄« ", " ›Ê« Ì— „»Ì⁄«  ")

Set data1.Recordset = myRecordSet("SELECT * FROM FILE0_40", con)
Set xStore.RowSource = data1
xStore.ListField = "Desca"
xStore.BoundColumn = "Code"
xStore.BoundText = cBranchStore

Set grid1.DataSource = DATA11
Set GRID2.DataSource = DATA22

openCardTable
If sDoc_no = "" Then
    CmdNewInv_Click
Else
    myUndo
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

CardTable.Close
Set CardTable = Nothing

closeCon con
Unload Search3
Unload oSearchDoc
Unload search32
Set SalesFrm = Nothing
Err.Clear
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
''On Error GoTo myerror
Dim bNew As Boolean

With grid1
If .Col = 9 Then
    If Not lIsRet Then
''''''        xBalance.Caption = LastBalance(grid1.TextMatrix(Row, 1), XSTORE.BoundText, con) + Val(GetDesca("select quant from file6_20 where id = " & Val(.TextMatrix(Row, .Cols - 1))))
''''''        If Val(xBalance.Caption) < Val(.TextMatrix(Row, 10)) Then
''''''            MsgBox "—’Ìœ «·’‰› «ﬁ· „‰ «·„»Ì⁄«  "
''''''        End If
    End If
End If
If .Col = 13 Then
    nPrice = Round((.ValueMatrix(Row, 12) / 1.14) * ((100 - .ValueMatrix(Row, 13)) / 100), 2)
    .TextMatrix(Row, 10) = nPrice
    con.Execute " UPDATE FILE6_20 SET PRICE =  " & nPrice & " WHERE DOC_NO = " & MyParn(xDoc_No.text) & " AND ID = " & Val(.TextMatrix(Row, .Cols - 1))
End If
If Not validRow(Row) Then Exit Sub
CalcTotals

If myreplace(Row) Then
    HandleCntEdit
    bNew = grid1.TextMatrix(Row, .Cols - 1) = ""
End If

myloadgrd


'MsgBox Row & "   " & Col
'If bNew Then
'    grid1.Row = Row
'    grid1.Col = 9
'    grid1.ShowCell Row, 9
''Else
'    grid1.Row = grid1.Rows - 1
'    grid1.Col = 1
'    grid1.ShowCell grid1.Rows - 1, 1
'End If

End With
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myEdit(Row As Long, Col As Long)
End Sub
Private Sub grid1_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If OldRow <> NewRow And OldRow <> grid1.Rows - 1 And OldRow <> 0 And grid1.TextMatrix(OldRow, grid1.Cols - 1) = "" Then
    If Not validRow(OldRow) Then
        grid1.RemoveItem OldRow
        CalcTotals
    End If
End If
End Sub
Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
imgx1.Images.Clear
imgx1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub
Private Sub grid1_EnterCell()
If grid1.TextMatrix(grid1.Row, 12) <> "" Then
    LoadPhoto grid1.TextMatrix(grid1.Row, 12)
ElseIf grid1.Row = grid1.Rows - 1 And Trim(grid1.TextMatrix(grid1.Row, 1)) = "" And grid1.Row > 1 Then
    LoadPhoto grid1.TextMatrix(grid1.Row - 1, 12)
End If

If (xClosed.Value <> 0) Then
    grid1.Editable = flexEDNone
    Exit Sub
End If

With grid1
    If .Col = 1 Or .Col = 9 Or .Col = 10 Or .Col = 13 Then
        .Editable = flexEDKbdMouse
    Else
        .Editable = flexEDNone
    End If
End With
End Sub
Private Sub Grid1_GotFocus()
If grid1.Rows < 2 Then Exit Sub
If grid1.Row = 0 Then
    grid1.Row = 1
    grid1.Col = 1
End If
grid1_EnterCell
End Sub
Private Sub Grid1_KeyDown(KeyCode As Integer, Shift As Integer)


If KeyCode = 112 And xClosed.Value = 0 Then
    ItemsLookupAll Me, oSearchItem
End If
End Sub

Private Sub Grid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    With grid1
    If grid1.Row = grid1.Rows - 1 And grid1.Row > 1 And Trim(grid1.TextMatrix(grid1.Row, 1)) = "" Then
        KeyAscii = 0
        If xClosed.Value = 0 Then cmdSave_Click
        Err.Clear
    End If
'    If .Col = 1 Then
'        .Col = 9
'        .Row = 1
'        .Select .Row, 9
'        .ShowCell .Row, 9
'    End If
    End With
End If
End Sub
Private Sub grid1_LostFocus()
    imgx1.Images.Clear
End Sub
Private Sub Grid1_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 10 Then
        xBalance.Caption = Val(LastBalance(grid1.TextMatrix(Row, 1), xStore.BoundText, con)) + Val(GetDesca("select quant from file6_20 where id = " & Val(grid1.TextMatrix(Row, grid1.Cols - 1)), con))
    End If
End Sub
Private Sub Grid1_Validate(Cancel As Boolean)
If (Not validRow(grid1.Row)) And grid1.Row <> grid1.Rows - 1 And grid1.Row <> 0 And grid1.TextMatrix(grid1.Row, grid1.Cols - 1) = "" Then
    grid1.RemoveItem grid1.Row
End If
End Sub
Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
With grid1
    If Col = 1 Then
        
'''        If .EditText <> "" And Len(.EditText) > 2 Then
'''            If UCase(Mid(.EditText, 1, 1)) = "J" Then
'''                cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(.EditText))
'''            ElseIf UCase(Mid(.EditText, 1, 1)) = "S" Then
'''                cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE2 = " & MyParn(.EditText))
'''            ElseIf Len(.EditText) >= 11 Then
'''                cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE13 = " & MyParn(.EditText))
'''                If cMyItem = "" Then cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE = " & MyParn(.EditText))
'''            End If
'''            If cMyItem <> "" Then
'''                .EditText = Val(cMyItem)
'''            End If
'''        End If
'''        If .EditText <> "" Then .EditText = DelZero(.EditText)
'''        If Len(.EditText) > 1 And Mid(.EditText, 1, 1) = "+" Then
'''            .TextMatrix(.Row - 1, 10) = Val(Mid(.EditText, 2))
'''            .TextMatrix(.Row - 1, 13) = Val(.TextMatrix(.Row - 1, 10)) * Val(.TextMatrix(.Row - 1, 11))
'''            .EditText = ""
'''            CalcTotals
'''            grid1_AfterEdit .Row - 1, 10
'''            Exit Sub
'''        End If
'''
'''        If Len(.EditText) > 1 And Mid(.EditText, 1, 1) = "-" Then
'''            .TextMatrix(.Row - 1, 10) = Val(.EditText)
'''            .TextMatrix(.Row - 1, 13) = Val(.TextMatrix(.Row - 1, 10)) * Val(.TextMatrix(.Row - 1, 11))
'''            .EditText = ""
'''            CalcTotals
'''            grid1_AfterEdit .Row - 1, 10
'''            Exit Sub
'''        End If
'''
'''        If .EditText = "-" And Row > 1 Then
'''            .TextMatrix(.Row - 1, 10) = -1
'''            .TextMatrix(.Row - 1, 13) = Val(.TextMatrix(.Row - 1, 10)) * Val(.TextMatrix(.Row - 1, 11))
'''            .EditText = ""
'''            CalcTotals
'''            grid1_AfterEdit .Row - 1, 10
'''            Exit Sub
'''        End If
'''        If IsNumeric(.EditText) And Len(.EditText) = 1 Then .EditText = RetZero(.EditText, 3)
'''        If .EditText <> "" And Len(.EditText) > 2 Then
'''            cMyItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE BARCODE = " & MyParn(.EditText))
'''            If cMyItem <> "" Then
'''                .EditText = Val(cMyItem)
'''            End If
'''        End If
        
        If GrdDesc(.EditText, Row) Then
            .EditText = .TextMatrix(.Row, 1)
        Else
            Cancel = True
        End If
    End If
End With
End Sub

Private Sub SSCommand1_Click()
    
    ToFileExel2 GRID2, , , , , 1.1, , , , , , Me
End Sub

Private Sub xcode_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 And xClosed.Value = 0 Then CLIENTLOOKUP
End Sub
Private Sub xCode_LostFocus()
    myLostFocus xCode
    xCodeDesca.Caption = ""
    xCodeDesca2.Caption = ""
    If xCode.text = "" Then Exit Sub
    xCode.text = RetZero(xCode.text, 3)
    Dim aRet
    aRet = aGetDesca("select desca , DISC , Manager from FILE3_10 where code = " & MyParn(xCode.text), con)
    If UBound(aRet) > 0 Then
        xCodeDesca.Caption = aRet(1) & ""
        xCodeDesca2.Caption = aRet(3) & ""
        XDISC.text = aRet(2) & ""
    Else
        xCodeDesca.Caption = ""
        xCodeDesca2.Caption = ""
        xCode.text = ""
        XDISC.text = ""
    End If
End Sub
Private Sub xDate_Validate(Cancel As Boolean)
If Not IsDate(xDate.text) Then Cancel = True
End Sub
Private Sub xDiscount_LostFocus()
    myLostFocus xDiscount
    CalcTotals
End Sub
Private Function MYVALID(Optional bIgMsg As Boolean = False) As Boolean
If Not IsDate(xDate.text) Then
    If Not bIgMsg Then MsgBox "«· «—ÌŒ €Ì— ”·Ì„"
    Exit Function
End If

If xStore.BoundText = "" Then
    If Not bIgMsg Then MsgBox "·„ Ì „ «œŒ«· «·„Œ“‰ "
    Exit Function
End If
If xCodeDesca.Caption = "" Then
    If Not bIgMsg Then MsgBox "·« ÌÊÃœ ﬂÊœ ⁄„Ì·"
    Exit Function
End If
If xMan.BoundText = "" Then
    If Not bIgMsg Then MsgBox "·« ÌÊÃœ »«∆⁄ "
    Exit Function
End If
MYVALID = True
End Function
Private Sub myload(Optional bLeaveBal As Boolean = False)
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
xDoc_no2.Locked = True
xDoc_No.text = CardTable!doc_no
xpo_no.text = CardTable!po_no & ""

xDoc_no2.text = CardTable!Doc_no2
xDate.text = Format(CardTable!Date, "dd-mm-yyyy")


XDATESALES1.text = Format(CardTable!DATESALES1, "dd-mm-yyyy")
XDATESALES2.text = Format(CardTable!DATESALES2, "dd-mm-yyyy")

xStore.BoundText = CardTable!STORE & ""
xMan.BoundText = CardTable!MAN & ""
xNotes.text = CardTable!NOTES & ""
XINV_TAX.text = CardTable!inv_tax & ""
xpo_no.text = CardTable!po_no & ""
xCode.text = CardTable!CODE & ""
xCodeDesca.Caption = CardTable!ClientDesca & ""
xCodeDesca2.Caption = CardTable!Manager & ""
XDISC.text = CardTable!disc & ""
xDiscount.text = TurnValue(Val(CardTable!discount & ""), 0, "")
xRate.text = TurnValue(Val(CardTable!Rate & ""), 0, "")
xTax.text = TurnValue(Val(CardTable!tax & ""), 0, "")
xRTax.text = TurnValue(Val(CardTable!RTAX & ""), 0, "")

'xTotal.Enabled = IIf(xPrinted.Value = 0, True, False)

xtime.Caption = Format(CardTable!Time, "hh:nn")
myloadgrd

Handlecontrols LoadMode
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub
Private Sub myDefine()
xDoc_no2.Locked = False
cmd_closed.BackColor = &H8080FF
cmd_closed.Caption = "› Õ „” ‰œ"
xClosed.Visible = False
xClosed.Value = ssCBUnchecked
If Not lIsRet Then
    xDoc_No.text = RetZero(IncRec(GetDesca("select MAX(doc_no) from file6_20h where ISRET = 0 AND isinvoice = 1", con)), 12)
Else
    xDoc_No.text = RetZero(IncRec(GetDesca("select MAX(doc_no) from file6_20h where ISRET = 1 AND isinvoice = 1", con)), 12)
    If xDoc_No.text = "000000000000" Then xDoc_No.text = "200000000001"
End If
xTotDisc.Caption = ""
xDoc_no2.text = Mid(xDoc_No.text, 7, 6)
XDATESALES1.text = ""
XDATESALES2.text = ""
XINV_TAX.text = ""
xRTax2.Caption = ""
xCode.Enabled = True
xDiscount.Locked = False
xRate.Locked = False
xTax.text = ""
xRTax.text = "14"
xMan.BoundText = "0001"
xDoc_No.Tag = DefineMode
xusername.Caption = ""
xUserName2.Caption = ""
XTIME1.Caption = ""
XTIME2.Caption = ""
xDate.text = Format(Date, "DD-MM-YYYY")
xBalance.Caption = ""
xCode.text = ""
xCodeDesca.Caption = ""
xCodeDesca2.Caption = ""


xDiscount.text = ""
xTotalQuant.Caption = ""
xTotalItemNoDiscount.Caption = ""
xTotalItem.Caption = ""
xTotalItemDiscount.Caption = ""
xRate.text = ""
xTotal.text = ""
xDiscount.text = ""
xNotes.text = ""
xtime.Caption = Format(Time, "hh:nn")


grid1.Rows = 1
grid1.AddItem ""
grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
xDoc_No.Tag = 0
xTotal.text = ""
imgx1.Images.Clear
Fixgrd
Fixgrd2
If lManSales Then
    grid1.TextMatrix(1, 1) = cdefman
    GrdDesc grid1.TextMatrix(1, 1), 1
    grid1.TextMatrix(1, 1) = ""
End If
Handlecontrols DefineMode
cmdSave.Enabled = True
End Sub
Private Sub Handlecontrols(nMode)
cmdNewInv.Enabled = nMode = LoadMode
cmdSave.Enabled = xClosed.Value = 0
CmdDel.Enabled = xClosed.Value = 0
cmdFirst.Enabled = (nMode = LoadMode)
cmdLast.Enabled = (nMode = LoadMode)
cmdNext.Enabled = (nMode = LoadMode)
cmdPrevious.Enabled = (nMode = LoadMode)

xDoc_No.Enabled = (nMode = DefineMode)
xDoc_No.Tag = nMode
xCode.Enabled = (xClosed.Value = 0)
xDiscount.Locked = (xClosed.Value <> 0)

cmdSign.Enabled = Trim(XINV_TAX.text = "") And nMode = LoadMode
cmdSignGroup.Enabled = Trim(XINV_TAX.text = "") And nMode = LoadMode
End Sub
Private Function retBool(cFieldName) As Boolean
If Not (CardTable.EOF Or CardTable.BOF) Then
    retBool = CardTable(cFieldName)
End If
End Function
Private Sub xDoc_No_LostFocus()
myLostFocus xDoc_No
If CardTable.EOF And CardTable.BOF Then Exit Sub
CardTable.Find "Doc_no = " & MyParn(xDoc_No.text), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then myload True
End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If xClosed.Value = 1 Then Exit Sub
If KeyCode = 46 And grid1.Row = grid1.Rows - 1 And xClosed.Value = 0 And grid1.Rows > 3 Then
    grid1.Select grid1.Rows - 2, 0
End If

If KeyCode = 116 Then
    cmdSave_Click
End If

If KeyCode = 46 And grid1.Row <> grid1.Rows - 1 And CmdDel.Enabled And grid1.Rows > 3 Then
    If MsgBox("Õ–› «·’‰› „‰ «·„” ‰œ ?, Â· «‰  „Ê«›ﬁ ø", 1 + 256) = vbOK Then
        RemoveItem (grid1.Row)
        grid1.Select grid1.Rows - 1, 1
        grid1.ShowCell grid1.Rows - 1, 1
        
        CalcTotals
        'UpdateHeader
        MakeSerial grid1.Row
    End If
End If
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Function GrdDesc(sitem As String, Row As Long) As Boolean
Dim lMosmDisc As Boolean, cBarCode As String, nItem   As Double
Dim cMyItem As String
With grid1

    If Len(sitem) > 7 Or Not IsNumeric(sitem) Then
        cBarCode = sitem
        nItem = 0
    Else
        cBarCode = ""
        nItem = Int(sitem & "")
    End If
    If nItem <> 0 Then
        Set loctable = ItemFind(sitem, con)
    Else
        Set loctable = ItemFind_BARCODE(cBarCode, con)
    End If
    If Not (loctable.EOF And loctable.BOF) Then
        grid1.TextMatrix(Row, 1) = loctable!Item
        grid1.TextMatrix(Row, 2) = loctable!MOSM
        grid1.TextMatrix(Row, 3) = loctable!Fact
        grid1.TextMatrix(Row, 4) = loctable!SUPP & ""
        grid1.TextMatrix(Row, 5) = loctable!modelfact0 & ""
        grid1.TextMatrix(Row, 6) = loctable!DESCA
        grid1.TextMatrix(Row, 7) = loctable!scal
        grid1.TextMatrix(Row, 8) = loctable!color
        nPrice = Round((loctable!price / 1.14) * ((100 - Val(XDISC.text)) / 100), 2)
        grid1.TextMatrix(Row, 10) = Format(nPrice, "#0.00")
        grid1.TextMatrix(Row, 11) = Format(nPrice, "#0.00")
        grid1.TextMatrix(Row, 12) = Format(loctable!price, "#0.00")
        grid1.TextMatrix(Row, 13) = Format(Val(XDISC.text), "#0.00")
        grid1.TextMatrix(Row, 14) = loctable!MODELNO & ""
        If Not lIsRet Then
            xBalance.Caption = LastBalance(loctable!Item, xStore.BoundText, con)
            If Val(xBalance.Caption) < 1 Then
                MsgBox "—’Ìœ «·’‰› ’›— «Ê √ﬁ·"
            End If
        End If
        GrdDesc = True
        LoadPhoto loctable!MODELNO & ""
    Else
        Inform "ﬂÊœ «·’‰› €Ì— ’ÕÌÕ"
    End If
End With
End Function
Private Function CalcTotals(Optional nMode As Integer = 0)
Dim nTotalQuant As Integer, nTotalItem As Single, nTotalItemNoDiscount As Single, nTotalDiscount As Single, nTotalMin As Single
With grid1
For i = 1 To grid1.Rows - 2
    nTotalQuant = nTotalQuant + Val(grid1.TextMatrix(i, 9))
    nTotalItem = nTotalItem + Val(.TextMatrix(i, 11))
    nTotalItemNoDiscount = nTotalItem
    nTotalItemDiscount = 0
Next
xTotalQuant.Caption = Myvalue(nTotalQuant, "#0")
xTotalQuant2.Caption = Myvalue(nTotalQuant2, "#0")
xTotalItemNoDiscount.Caption = Myvalue(nTotalItemNoDiscount, "Fixed")
xTotalItemDiscount.Caption = Myvalue(nTotalItemDiscount, "Fixed")
xTotalItem.Caption = Myvalue(nTotalItem, "Fixed")
If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Then
        xRate.text = Myvalue(Round((Val(xDiscount.text) / Val(xTotalItem.Caption)) * 100, nRound))
    End If
Else
    xRate.text = ""
End If
xTotal.text = Round(nTotalItem - Val(xDiscount.text) + Val(xTax.text), 2)
xTotDisc.Caption = Round(nTotalItem - Val(xDiscount.text), 2)
xRTax2.Caption = ""
If Val(xTotalItem.Caption) <> 0 Then xRTax2.Caption = Round((Val(xTax.text) / Val(xTotalItem.Caption)) * 100, 2)
End With
End Function
Private Sub CardLookup(Optional pFilter As String = "")
Dim Generalarray(5)
Dim listarray()
Dim GrdArray(9, 1)
Set Generalarray(0) = Me
Generalarray(1) = "SELECT   FILE6_20H.DOC_NO, FILE6_20H.DOC_NO2, FILE6_20H.INV_NO, CONVERT(VARCHAR(10), FILE6_20H.date, 111) AS Expr1, FILE0_50.DESCA, CONVERT(VARCHAR(10), FILE6_20H.time, 108) AS Expr2, FILE3_10.DESCA  , SUM(FILE6_20.TOTAL) AS total , DATESALES1 , DATESALES2 " & _
                  " FROM    FILE6_20H LEFT JOIN  FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO INNER JOIN FILE3_10 ON FILE6_20H.CODE = FILE3_10.CODE LEFT OUTER JOIN    FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE WHERE FILE6_20H.ISINVOICE = 1 and isret = " & IIf(lIsRet, 1, 0)
Generalarray(2) = "GROUP BY FILE6_20H.DOC_NO, FILE3_10.DESCA , FILE6_20H.DOC_NO2, FILE6_20H.INV_NO, CONVERT(VARCHAR(10), FILE6_20H.date, 111), FILE0_50.DESCA, CONVERT(VARCHAR(10),  FILE6_20H.time, 108) , DATESALES1 , DATESALES2  Order by doc_no  DESC "
Generalarray(3) = 6000
Generalarray(5) = False
If bopt3 Then
   ReDim listarray(2, 5)
   listarray(0, 0) = "«·—ﬁ„-«·Œ“‰…"
   listarray(0, 1) = "(@@Doc_No2@@4 or %%file0_50.DESCA%% or %%file3_10.DESCA%%) "

   listarray(1, 0) = "«· «—ÌŒ"
   listarray(1, 1) = "(##DATE##)"

   listarray(2, 0) = "—ﬁ„ «·«Ì’«·"
   listarray(2, 1) = "(%%INV_NO%%)"
Else
   ReDim listarray(1, 5)

    listarray(0, 0) = "«·—ﬁ„"
    listarray(0, 1) = "(@@Doc_No2@@4 or %%file3_10.DESCA%%)"
   
   listarray(1, 0) = "—ﬁ„ «·«Ì’«·"
   listarray(1, 1) = "(%%INV_NO%%)"
End If

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(1, 1) = 1200

GrdArray(2, 0) = "—ﬁ„ «·«Ì’«·"
GrdArray(2, 1) = 0

GrdArray(3, 0) = "«· «—ÌŒ"
GrdArray(3, 1) = 1500

GrdArray(4, 0) = "«·›—⁄"
GrdArray(4, 1) = 1500

GrdArray(5, 0) = "«·Êﬁ "
GrdArray(5, 1) = 1300

GrdArray(6, 0) = "«·⁄„Ì·"
GrdArray(6, 1) = 2500

GrdArray(7, 0) = "«·ﬁÌ„…"
GrdArray(7, 1) = 1300

GrdArray(8, 0) = "„‰  «—ÌŒ"
GrdArray(8, 1) = 1300

GrdArray(9, 0) = "≈·Ï  «—ÌŒ "
GrdArray(9, 1) = 1300

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchDoc.sId = "sales" & IIf(bopt3, "1", "0")
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub xDoc_No_Validate(Cancel As Boolean)

If xDoc_No.text = "" Then Cancel = True
End Sub
Private Sub xDoc_no2_LostFocus()
    If lIsRet Then
        If xDoc_no2.text <> "" Then xDoc_No.text = "2" & RetZero(xDoc_no2.text, 11)
        If xDoc_no2.text <> "" Then xDoc_no2.text = RetZero(xDoc_no2.text, 5)
    Else
        If xDoc_no2.text <> "" Then xDoc_No.text = "1" & RetZero(xDoc_no2.text, 11)
        If xDoc_no2.text <> "" Then xDoc_no2.text = RetZero(xDoc_no2.text, 5)
    End If
End Sub
Private Sub xMan_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then MANLOOKUP
End Sub
Private Sub xRate_LostFocus()
myLostFocus xRate
If Val(xTotalItem.Caption) <> 0 Then
    If Round(Val(xRate.text), nRound) <> Round(Val(xDiscount.text) / Val(xTotalItem.Caption) * 100, nRound) Or xDiscount.Locked Then
        xDiscount.text = Round((Val(xRate.text) * Val(xTotalItem.Caption)) / 100, nRound)
    End If
Else
    xDiscount.text = ""
End If
CalcTotals
End Sub
Private Function RetItemBalance(cItem, cStore, dDate) As Double
If cItem = "" Then Exit Function
movetable.Seek Array(cItem, cStore), adSeekFirstEQ
Do Until movetable.EOF
    If IsNull(movetable!Date) Then Exit Do
    If Trim(movetable!Item) <> cItem Or cStore <> movetable!STORE Or DateValue(movetable!Date) > DateValue(Format(dDate, "dd-mm-yyyy")) Then Exit Do
    RetItemBalance = RetItemBalance + TurnValue(movetable!In, Null, 0) - TurnValue(movetable!out, Null, 0)
    movetable.MoveNext
Loop
End Function
Private Sub MakeSerial(Optional nBeginRow As Integer = 1)
For i = 1 To grid1.Rows - 1
    grid1.TextMatrix(i, 0) = i
Next
End Sub
Private Sub Fixgrd()
With grid1

'                0        1             2       3          4           5           6           7         8          9           10          11           12                 13              14              15          16
.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "⁄œœ|" & "”⁄— «·»Ì⁄|" & "«·≈Ã„«·Ï|" & "”⁄— „” Â·ﬂ|" & "‰”»… Œ’„|" & "—ﬁ„ «·„ÊœÌ·|" & "BARCODE GS1|" & "AISN|"
.RowHeight(0) = 800
.WordWrap = True
.ColHidden(.Cols - 1) = True
.ColWidth(0) = 600
.ColWidth(1) = 1300
.ColWidth(2) = 800
.ColWidth(3) = 2000
.ColWidth(4) = 0
.ColWidth(5) = 1500
.ColWidth(6) = 4000
.ColWidth(7) = 1000
.ColWidth(8) = 1000
.ColWidth(9) = 800
.ColWidth(10) = 1000
.ColWidth(11) = 1000

.ColWidth(12) = 1000
.ColWidth(13) = 1000

.ColWidth(15) = 2000
.ColWidth(16) = 1500

.ColHidden(14) = True
.ColHidden(15) = True
.ColHidden(16) = True

.ColComboList(3) = clist1
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
.ShowCell .Rows - 1, 0
For nRow = 1 To .Rows - 1
    nRate = 0
    If Val(.TextMatrix(nRow, 10)) > 0 Then
        nPrice = Round(Val(.TextMatrix(nRow, 12)) / 1.14)
        If nPrice <> 0 Then nRate = Round((nPrice - Val(.TextMatrix(nRow, 10))) / nPrice * 100, 2)
    End If
    .TextMatrix(nRow, 13) = nRate
Next nRow
End With
End Sub
Private Sub CLIENTLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From file3_10"
Generalarray(2) = "Order by file3_10.Desca"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ «·⁄„Ì·"
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ «·⁄„Ì·"
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCl1.Caption = "«” ⁄·«„"
oSearchCl1.Show 1
End Sub



Private Sub xRTax_DblClick()
    xTax.text = Round(Val(xTotDisc.Caption) * ((Val(xRTax.text)) / 100), 2)
End Sub
Private Sub xStore_GotFocus()
myGotFocus xStore
End Sub
Private Sub xstore_LostFocus()
myLostFocus xStore
If Not xStore.MatchedWithList Then xStore.BoundText = ""
End Sub
Private Sub xStore_Validate(Cancel As Boolean)
    If Trim(xStore.BoundText) = "" Then Cancel = True
End Sub
Private Sub xTax_LostFocus()
    xTax.BackColor = &H80000005
    CalcTotals
End Sub
Private Function RemoveItem(nRow) As Boolean
'On Error GoTo myerror
con.BeginTrans
If grid1.TextMatrix(nRow, grid1.Cols - 1) <> "" Then
    con.Execute "Delete  From FILE6_20 where id = " & grid1.TextMatrix(nRow, grid1.Cols - 1)
End If
grid1.RemoveItem nRow
con.CommitTrans
Exit Function
myerror:
con.RollbackTrans
MsgBox Err.Description
Err.Clear
End Function
Sub myproc2(nDoc_no)
CardTable.Find "Doc_no = " & MyParn(nDoc_no), , adSearchForward, adBookmarkFirst
If Not CardTable.EOF Then
    myload
Else
    MsgBox "—ﬁ„ «·›« Ê—… €Ì— ’ÕÌÕ"
    Unload Me
End If
End Sub
Function itemPrice(cItem) As Single
    itemPrice = GetDesca("select PRICE from file1_10 where item = " & MyParn(cItem), con)
End Function
Private Function myreplaceGrd(nRow) As Boolean
Dim aInsert(8, 1)
With grid1
    For i = IIf(nRow = -1, 1, nRow) To IIf(nRow = -1, grid1.Rows - 2, nRow)
        aInsert(0, 0) = "doc_no"
        aInsert(0, 1) = addstring(xDoc_No.text)
        
        aInsert(1, 0) = "item"
        aInsert(1, 1) = addstring(grid1.TextMatrix(i, 1))
        
        aInsert(2, 0) = "quant"
        If lIsRet Then
            aInsert(2, 1) = Abs(Val(.TextMatrix(i, 9))) * -1
        Else
            aInsert(2, 1) = Val(.TextMatrix(i, 9))
        End If

        aInsert(3, 0) = "Price"
        aInsert(3, 1) = Val(.TextMatrix(i, 10))

        aInsert(4, 0) = "Price_C"
        aInsert(4, 1) = Val(.TextMatrix(i, 12))

        aInsert(5, 0) = "row"
        aInsert(5, 1) = i
        
        aInsert(6, 0) = "COST"
        aInsert(6, 1) = Item_cost(.TextMatrix(i, 1), con)

        aInsert(7, 0) = "MAN"
        aInsert(7, 1) = addstring(xMan.BoundText)

        aInsert(8, 0) = "USER_IP"
        aInsert(8, 1) = addstring(cIpName)


        If grid1.TextMatrix(i, grid1.Cols - 1) = "" Then
            con.Execute CreateInsert(aInsert, "FILE6_20")
        Else
            con.Execute CreateUpdate(aInsert, "FILE6_20", " where ID = " & grid1.TextMatrix(i, .Cols - 1))
        End If
    
'        If nRow <> 0 Then grid1.Row = Row
    Next
End With
myreplaceGrd = True
End Function
Private Sub myloadgrd()
With grid1
    Dim cString As String
'                           1          2                3                4                   5               6             7                8                9                 10               11           12   13        14              15                                                                                  16          17
    cString = "SELECT FILE6_20.ITEM, FILE1_10.MOSM, FILE1_10.FACT, FILE1_10.SUPP, FILE1_10.MODELFACT0, FILE1_10.DESCA, FILE1_10.SCAL , FILE1_10.COLOR , file6_20.Quant, FILE6_20.PRICE , FILE6_20.TOTAL ,PRICE_C, 0 ,FILE1_10.MODELNO  , case when file1_10.barcode13 is null then barcode else file1_10.barcode13 end , BARCODE_GS1, ENA , ID " & _
          " FROM (FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM) LEFT JOIN FILE6_25 ON FILE6_20.MAN  = FILE6_25.CODE"
    cString = cString & turn(cString) & " DOC_NO = " & MyParn(xDoc_No.text)
    cString = cString & " ORDER BY FILE6_20.ID"
    Set DATA11.Recordset = myRecordSet(cString, con)
    .AddItem ""
End With
CalcTotals
Fixgrd
myloadgrd2
'grid1.Row = grid1.Rows - 1
grid1.ShowCell grid1.Rows - 1, 1
End Sub
Private Sub xNotes_GotFocus()
myGotFocus xNotes
End Sub
Private Sub xCode_GotFocus()
myGotFocus xCode
End Sub
Private Sub xDoc_No_GotFocus()
myGotFocus xDoc_No
End Sub
Private Sub xdate_GotFocus()
myGotFocus xDate
End Sub
Private Sub xRate_GotFocus()
myGotFocus xRate
End Sub
Private Sub xRateDis_GotFocus()
xRateDis.SelStart = 0
xRateDis.SelLength = Len(xRateDis.text)
End Sub
Private Sub xDiscount_GotFocus()
myGotFocus xDiscount
End Sub
Private Sub xTax_GotFocus()
xTax.SelStart = 0
xTax.SelLength = Len(xTax.text)
End Sub
Private Function mysave(Optional bEnd As Boolean = True, Optional bPrint As Boolean = True) As Boolean
If Not MYVALID Then Exit Function
CalcTotals
If Not myreplace Then Exit Function
If bEnd And lSave Then
    Inform " „ Õ›Ÿ «·„” ‰œ »‰Ã«Õ"
    myDefine
Else
    myUndo
End If
End Function
Private Sub old_openCardTable()
Set CardTable = Nothing
Set CardTable = New ADODB.Recordset
cFilter = ""
cString = "SELECT FILE6_20H.*,FILE3_10.DESCA AS CLIENTDESCA , FILE3_10.DISC , file3_10.Manager FROM FILE6_20H INNER JOIN FILE3_10 ON FILE6_20H.Code = FILE3_10.CODE WHERE ISINVOICE = 1 AND ISRET = " & IIf(lIsRet, 1, 0)
If sDoc_no <> "" Then
    cFilter = cFilter & turn(cFilter, " AND ") & " DOC_NO = " & MyParn(sDoc_no)
End If
If cFilter <> "" Then cString = cString & turn(cString) & cFilter
cString = cString & " ORDER BY DOC_NO "
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Sub
Private Function openCardTable(Optional pDoc_No As String = "", Optional pSign As String = "=")
Dim cString As String, cWhere As String
Set CardTable = New ADODB.Recordset
cFilter = ""
cFilterLook = ""
cString = "SELECT TOP 1 FILE6_20H.* , FILE3_10.DESCA AS CLIENTDESCA , FILE3_10.DISC , file3_10.Manager from FILE6_20H inner join file3_10 on file3_10.code = FILE6_20H.code INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE WHERE ISINVOICE = 1 AND ISRET = " & IIf(lIsRet, 1, 0)
If sDoc_no <> "" Then
    cFilter = cFilter & turn(cFilter, " AND ") & " DOC_NO = " & MyParn(sDoc_no)
End If
If pSign = "=" Then
    If pDoc_No <> "" Then cWhere = " AND DOC_NO  " & pSign & addstring(pDoc_No)
Else
    If pDoc_No <> "" Then cWhere = " AND DOC_NO  " & pSign & addstring(pDoc_No)
End If
If cFilter <> "" Then cWhere = cWhere & cFilter
If cWhere <> "" Then cString = cString & cWhere
If pSign = "<" Or pSign = "<=" Then
    cString = cString & " order by doc_no desc"
ElseIf pSign = ">=" Or pSign = ">" Then
    cString = cString & " order by doc_no ASC"
End If
Set CardTable = New ADODB.Recordset
CardTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
End Function

Private Function validRow(Row As Long, Optional bIgMsg As Boolean = False, Optional bIgMsgsub As Boolean = True) As Boolean
With grid1
If Not MYVALID(bIgMsg) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 1)) Then Exit Function
'If Not IsNumeric(.TextMatrix(Row, 9)) Then Exit Function
If Not IsNumeric(.TextMatrix(Row, 10)) Then Exit Function
End With
validRow = True
End Function
Private Sub HandleCntEdit()
xDoc_No.Tag = LoadMode
xDoc_No.Enabled = False
CmdDel.Enabled = True
cmdSave.Enabled = (xClosed.Value = 0)

End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col = 1 Then
    grid1.Col = 9
    grid1.Row = Row
    grid1.ShowCell Row, 9
    grid1.Select Row, 9
    Exit Sub
ElseIf Col >= 9 Then
    grid1.Row = Row + 1
    grid1.Col = 1
    grid1.ShowCell grid1.Row, 1
    Exit Sub
End If
End Sub

Private Sub xNotes_LostFocus()
myLostFocus xNotes
End Sub
Private Sub xDate_LostFocus()
myLostFocus xDate
End Sub
Private Sub doprint_doc()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double
Dim cAddres As String, cTaxNo As String

cStr1 = "SELECT FILE6_20.ITEM , FILE1_10.DESCA, FILE1_10.DESCA2 , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.COLOR , FILE6_20.PRICE , SUM(FILE6_20.QUANT ) AS QUANT , FILE1_10.MODELFACT0 , SUM(FILE6_20.TOTAL) AS TOTAL  , FILE1_10.BARCODE_GS1 FROM FILE1_10 INNER JOIN FILE6_20 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.text) & " GROUP BY FILE6_20.ITEM , FILE1_10.DESCA, FILE1_10.DESCA2 , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.COLOR , FILE6_20.PRICE , FILE1_10.MODELFACT0, FILE1_10.BARCODE_GS1"
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
cAddres = GetDesca("SELECT ADDRESS FROM FILE3_10 WHERE CODE = " & MyParn(xCode.text), con)
cTaxNo = GetDesca("SELECT EMAIL FROM FILE3_10 WHERE CODE = " & MyParn(xCode.text), con)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With sourcetable
Do While Not .EOF
    temptable.AddNew
    temptable!str1 = xDoc_no2.text
    temptable!str2 = IIf(xCodeDesca2.Caption = "", xCodeDesca.Caption, xCodeDesca2.Caption)
    temptable!Date1 = xDate.text
    temptable!str6 = TurnValue(cAddres)
    temptable!str15 = TurnValue(cTaxNo)
    temptable!str14 = !modelfact0
    If lIsRet Then
        temptable!val3 = !TOTAL * -1
        temptable!VAL1 = !Quant * -1
        temptable!VAL7 = Val(xTotalItem.Caption) * -1
        temptable!val9 = (Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption)) * -1
        temptable!val18 = Val(xTotalItem.Caption) * -1
        temptable!val17 = 0
        temptable!STR19 = 0
        temptable!val15 = Val(xTax.text) * -1
        temptable!val16 = Val(xTotal.text) * -1
        temptable!str16 = MyOnly(Val(xTotal.text) * -1)
        temptable!VAL7 = Val(xDiscount.text) * -1
        temptable!Val8 = Val(xTotDisc.Caption) * -1
        temptable!str3 = " ›« Ê—… „— Ã⁄ "
    Else
        temptable!str3 = " ›« Ê—… „»Ì⁄«  "
        temptable!val3 = !TOTAL
        temptable!VAL1 = !Quant
        temptable!VAL7 = Val(xTotalItem.Caption)
        temptable!val9 = Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption)
        temptable!val18 = Val(xTotalItem.Caption)
        temptable!val17 = 0
        temptable!STR19 = 0
        temptable!val15 = Val(xTax.text)
        temptable!val16 = Val(xTotal.text)
        temptable!str16 = MyOnly(Val(xTotal.text))
        temptable!VAL7 = Val(xDiscount.text)
        temptable!Val8 = Val(xTotDisc.Caption)
    End If
    temptable!val2 = !price
    temptable!str8 = !DESCA
    temptable!str9 = !scal
    temptable!VAL19 = !C_SCAL
    temptable!STR7 = !color
    
    temptable!str18 = "Œ’„ «·›« Ê—…"
    
    temptable!str17 = TurnValue(!BARCODE_GS1)
    temptable!str12 = TurnValue(!Desca2)
    temptable!str13 = "PO NO   : " & xpo_no.text
    
    temptable.Update
    .MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
If xCode.text = "2019" Then
    Main.REPORT1.ReportFileName = App.Path & "\Reports\PRINT_INVOICE_amazon.rpt"
Else
    Main.REPORT1.ReportFileName = App.Path & "\Reports\PRINT_INVOICE.rpt"
End If
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub xMAN_GotFocus()
    xMan.BackColor = &HC0FFFF
End Sub
Private Sub xMan_LostFocus()
xMan.BackColor = &H80000005
End Sub
Private Sub xMAN_Validate(Cancel As Boolean)
If IsNumeric(xMan.text) Then
    xMan.text = RetZero(xMan.text, 3)
    xMan.BoundText = xMan.text
End If
xMan.BoundText = RetZero(xMan.BoundText, 3)
If Not xMan.MatchedWithList Then xMan.BoundText = ""
'If Trim(xMan.BoundText) = "" Then
'    Cancel = True
'    MsgBox " ”ÃÌ· «·»«∆⁄"
'End If


End Sub
Private Sub MANLOOKUP()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me
Generalarray(1) = "Select Code, DescA From file6_25 WHERE ISSTOP = 0 "
Generalarray(2) = "Order by Desca"
Generalarray(3) = 4000
Generalarray(5) = False

listarray(0, 0) = "«·ﬂÊœ √Ê «·«”„"
listarray(0, 1) = "(%%DESCA%%) "

GrdArray(0, 0) = "ﬂÊœ "
GrdArray(0, 1) = 1000

GrdArray(1, 0) = "≈”„ "
GrdArray(1, 1) = 3000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchCl1.Caption = "«” ⁄·«„"
oSearchCl1.Show 1
End Sub
Private Function myreplaceGrdModel() As Boolean
Dim cString As String, aitem As Variant
Dim cFPrice As String
If Check1.Value = 0 Then
    cFPrice = "PRICE"
Else
    cFPrice = "COST"
End If
cString = " DELETE FROM FILE6_20 FROM FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM"
cString = cString & turn(cString) & "DOC_NO = " & MyParn(xDoc_No.text)
cString = cString & turn(cString) & "FILE1_10.MODEL = " & MyParn(salesModel.XMODEL.text)
con.Execute cString, nDel
With salesModel.grid1
    For nRow = 3 To .Rows - 1
        For nCol = 2 To .Cols - 1
            If Val(.TextMatrix(nRow, nCol)) <> 0 Then
                aitem = ItemFields(salesModel.GRID2.TextMatrix(nRow, nCol), con)
                If Not IsEmpty(aitem) Then
                    cString = "Insert into FILE6_20 (doc_no,item,price,price_c,cost,price_c2,man,Quant)" & _
                               "Values(" & _
                                addstring(xDoc_No.text) & "," & _
                                addvalue(retFlag(aitem, "item")) & "," & _
                                Val(retFlag(aitem, cFPrice) & "") & "," & _
                                Val(retFlag(aitem, cFPrice) & "") & "," & _
                                Val(retFlag(aitem, "COST") & "") & "," & _
                                Val(retFlag(aitem, cFPrice) & "") & "," & _
                                addstring(SalesFrm.xMan.BoundText) & "," & _
                                Val(.TextMatrix(nRow, nCol)) & _
                               ")"
                    con.Execute cString
                End If
            End If
        Next
    Next
End With


End Function



Sub myProcAddModel()
If Not MYVALID(False) Then Exit Sub
If myreplace(, True) Then
    Inform " „  «÷«›… «·„ÊœÌ·«  »‰Ã«Õ"
End If
myloadgrd
myUndo

grid1.ShowCell grid1.Rows - 1, 2 + 1
grid1.Select grid1.Rows - 1, 2 + 1

End Sub
Private Sub doprint_SUBINV()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double, nDisc As Double, nTotalItem As Double, nRate As Double, nTotalQuant As Double
Dim cAddres As String, cTaxNo As String

nDisc = Val(GetDesca("SELECT SUM(DISCOUNT) FROM FILE6_20H  WHERE ISINVOICE = 0 AND CODE = " & MyParn(xCode.text) & " AND DATE >= " & DateSq(XDATESALES1.text) & " AND DATE <= " & DateSq(XDATESALES2.text), con) & "") + Val(xDiscount.text)
nTotalItem = Val(GetDesca("SELECT SUM(FILE6_20.TOTAL) FROM FILE6_20H  INNER JOIN FILE6_20 ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text), con) & "")
nRate = 1 - (Round(nDisc / nTotalItem, 10))
nTotalQuant = Val(GetDesca("SELECT SUM(QUANT) FROM FILE6_20H  INNER JOIN FILE6_20 ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text), con) & "")

cStr1 = "SELECT FILE6_20.ITEM , FILE1_10.DESCA , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.COLOR , SUM(FILE6_20.QUANT ) AS QUANT , FILE1_10.MODELFACT0 , SUM(FILE6_20.TOTAL) AS TOTAL  FROM FILE1_10 INNER JOIN FILE6_20 ON FILE6_20.ITEM = FILE1_10.ITEM  INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO  WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text) & " GROUP BY FILE6_20.ITEM , FILE1_10.DESCA , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.COLOR , FILE1_10.MODELFACT0"
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
cAddres = GetDesca("SELECT ADDRESS FROM FILE3_10 WHERE CODE = " & MyParn(xCode.text), con)
cTaxNo = GetDesca("SELECT EMAIL FROM FILE3_10 WHERE CODE = " & MyParn(xCode.text), con)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With sourcetable
Do While Not .EOF
    If !Quant <> 0 Then
        temptable.AddNew
        temptable!str1 = xDoc_no2.text
        temptable!str2 = IIf(xCodeDesca2.Caption = "", xCodeDesca.Caption, xCodeDesca2.Caption)
        temptable!Date1 = xDate.text
        temptable!str6 = TurnValue(cAddres)
        temptable!str15 = TurnValue(cTaxNo)
        temptable!str14 = !modelfact0
        temptable!val2 = Round((((!TOTAL / !Quant) * nRate) / 1.14), 6)
        temptable!val3 = Round((((!TOTAL / !Quant) * nRate) / 1.14), 6) * !Quant
        temptable!VAL1 = !Quant
        temptable!str8 = !DESCA
        temptable!str9 = !scal
        temptable!VAL19 = !C_SCAL
        temptable!STR7 = !color
        temptable!str3 = " ›« Ê—… „»Ì⁄«  "
        
    '   temptable!VAL7 = Val(xTotalItem.Caption)
        temptable!val9 = nTotalQuant
        
    '   temptable!str18 = "≈Ã„«·Ï «’‰«›"
    '   temptable!val18 = Val(xTotalItem.Caption)
    '   temptable!val17 = 0
    '   temptable!Val8 = Val(xTotDisc.Caption)
        
        temptable!val15 = (nTotalItem * nRate) - Round((nTotalItem * nRate) / 1.14, 2)
        temptable!val16 = nTotalItem * nRate
    '   temptable!str16 = MyOnly(Val(xTotal.Text))
    '   temptable!str18 = "Œ’„ «·›« Ê—…"
    '   temptable!VAL7 = Val(xDiscount.Text)
        
        temptable.Update
    End If
    .MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\PRINT_INVOICE.rpt"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Sub AddFromExel()
    Dim xl As New Excel.Application
    Dim xlsheet As Excel.Worksheet
    Dim xlwbook As Excel.Workbook, cItem As String
    Dim cFileName As String, cBarCode As String
    Dim nCount As Double, nItem As Double, cModel As String, nquant As Double, nPrice As Double
    If MsgBox("≈÷«›… »Ì«‰«  „‰ „·› Excel", vbYesNo) = vbYes Then
        Common1.InitDir = ""
        Common1.FileName = ""
        Common1.Filter = "Excel (*.XLS*)|*.XLS*"
        Common1.ShowOpen
        cFileName = Common1.FileName
        If cFileName <> "" Then
            Set xlwbook = xl.Workbooks.Open(cFileName)
            Set xlsheet = xlwbook.Sheets.Item(1)
            With grid1
            i = 1
            Do While True
                i = i + 1
                Me.Caption = i
                If (xlsheet.Cells(i, 1)) <> "" Then
                    nCount = i
                Else
                    Exit Do
                End If
            Loop
            If nCount < 2 Then Exit Sub
            prog1.Visible = True
            prog1.Value = 0
            prog1.Max = nCount
            prog1.Min = 0
            .Rows = 1
            For nRow = 2 To nCount
                If xlsheet.Cells(nRow, 2) <> "" Then
                    cBarCode = (xlsheet.Cells(nRow, 1))
                    cItem = ""
                    If IsNumeric(cBarCode) And Len(cBarCode) <= 7 Then
                        cItem = cBarCode
                    End If
                    Dim loctable As ADODB.Recordset
                    If cItem <> "" Then
                        Set loctable = ItemFind(cItem, con)
                    Else
                        Set loctable = ItemFind_BARCODE(cBarCode, con)
                    End If
                    cItem = ""
                    If Not (loctable.EOF And loctable.BOF) Then
    '                    cItem = GetDesca("SELECT ITEM FROM FILE1_10 WHERE  ITEM = " & Val(cBarCode) & " OR BARCODE2 = " & MyParn(cBarCode) & " OR BARCODE13 = " & MyParn(cBarCode) & " OR BARCODE = " & MyParn(cBarCode), con)
    '                    nPriceC = Val(GetDesca("SELECT PRICE FROM FILE1_10 WHERE  ITEM = " & Val(cBarCode) & " OR BARCODE2 = " & MyParn(cBarCode) & " OR BARCODE13 = " & MyParn(cBarCode) & " OR BARCODE = " & MyParn(cBarCode), con) & "")
                        cItem = loctable!Item
                        nPriceC = loctable!price
                        If cItem <> 0 Then
                            grid1.AddItem ""
                            nquant = Val(xlsheet.Cells(nRow, 2))
                            nPrice = Abs(Val(xlsheet.Cells(nRow, 3)))
                            grid1.TextMatrix(grid1.Rows - 1, 0) = grid1.Rows - 1
                            grid1.TextMatrix(grid1.Rows - 1, 1) = cItem
                            If lIsRet Then
                                grid1.TextMatrix(grid1.Rows - 1, 9) = Abs(nquant) * -1
                            Else
                                grid1.TextMatrix(grid1.Rows - 1, 9) = nquant
                            End If
                            GrdDesc grid1.TextMatrix(grid1.Rows - 1, 1), grid1.Rows - 1
                            If nPrice <> 0 Then grid1.TextMatrix(grid1.Rows - 1, 10) = nPrice
                            myreplace (grid1.Rows - 1)
                        Else
                            MsgBox " »«—ﬂÊœ €Ì— ’ÕÌÕ " & cBarCode
                        End If
                    End If
                End If
            Next nRow
            MsgBox "«„ «÷«›… «·«’‰«›  "
            End With
        End If
    End If
End Sub


Private Sub doprint_DOC_REC()
Dim aHeader(2)
If Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double
Dim cAddres As String, cTaxNo As String

cStr1 = "SELECT FILE1_10.BARCODE_GS1 , FILE6_20.ITEM , FILE1_10.DESCA , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.COLOR , FILE6_20.PRICE , SUM(FILE6_20.QUANT ) AS QUANT , FILE1_10.MODELFACT0 , SUM(FILE6_20.TOTAL) AS TOTAL  FROM FILE1_10 INNER JOIN FILE6_20 ON FILE6_20.ITEM = FILE1_10.ITEM WHERE DOC_NO = " & MyParn(xDoc_No.text) & " GROUP BY FILE6_20.ITEM , FILE1_10.DESCA , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.COLOR , FILE6_20.PRICE , FILE1_10.BARCODE_GS1 ,FILE1_10.MODELFACT0"
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText
cAddres = GetDesca("SELECT ADDRESS FROM FILE3_10 WHERE CODE = " & MyParn(xCode.text), con)
cTaxNo = GetDesca("SELECT EMAIL FROM FILE3_10 WHERE CODE = " & MyParn(xCode.text), con)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable
With sourcetable
Do While Not .EOF
    temptable.AddNew
    temptable!str1 = xDoc_no2.text
    temptable!str2 = IIf(xCodeDesca2.Caption = "", xCodeDesca.Caption, xCodeDesca2.Caption)
    temptable!Date1 = xDate.text
    temptable!str6 = TurnValue(cAddres)
    temptable!str15 = TurnValue(cTaxNo)
    temptable!str14 = !modelfact0
    If lIsRet Then
        temptable!val3 = !TOTAL * -1
        temptable!VAL1 = !Quant * -1
        temptable!VAL7 = Val(xTotalItem.Caption) * -1
        temptable!val9 = (Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption)) * -1
        temptable!val18 = Val(xTotalItem.Caption) * -1
        temptable!val17 = 0
        temptable!STR19 = 0
        temptable!val15 = Val(xTax.text) * -1
        temptable!val16 = Val(xTotal.text) * -1
        temptable!str16 = MyOnly(Val(xTotal.text) * -1)
        temptable!VAL7 = Val(xDiscount.text) * -1
        temptable!Val8 = Val(xTotDisc.Caption) * -1
        temptable!str3 = " ›« Ê—… „— Ã⁄ "
    Else
        temptable!str3 = " ›« Ê—… „»Ì⁄«  "
        temptable!val3 = !TOTAL
        temptable!VAL1 = !Quant
        temptable!VAL7 = Val(xTotalItem.Caption)
        temptable!val9 = Val(xTotalQuant.Caption) - Val(xTotalQuant2.Caption)
        temptable!val18 = Val(xTotalItem.Caption)
        temptable!val17 = 0
        temptable!STR19 = 0
        temptable!val15 = Val(xTax.text)
        temptable!val16 = Val(xTotal.text)
        temptable!str16 = MyOnly(Val(xTotal.text))
        temptable!VAL7 = Val(xDiscount.text)
        temptable!Val8 = Val(xTotDisc.Caption)
    End If
    temptable!val2 = !price
    temptable!str8 = !DESCA
    temptable!str9 = !scal
    temptable!VAL19 = !C_SCAL
    temptable!STR7 = !color
    
    temptable!str10 = !BARCODE_GS1
    temptable!str18 = "Œ’„ «·›« Ê—…"
    
    temptable.Update
    .MoveNext
Loop
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\PRINT_REC.rpt"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub xfilter2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter2.text, 5
    End If
End Sub
Private Sub xfilter3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter3.text, 3
    End If
End Sub
Private Sub xfilter4_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter4.text, 6
    End If
End Sub
Private Sub xfilter_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        FilterGrd grid1, xfilter.text, 1
    End If
End Sub
Private Sub ShowInv_Grid()
Dim sourcetable As New ADODB.Recordset
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double, nDisc As Double, nTotalItem As Double, nRate As Double, nTotalQuant As Double
Dim cAddres As String, cTaxNo As String

nDisc = Val(GetDesca("SELECT SUM(DISCOUNT) FROM FILE6_20H  WHERE ISINVOICE = 0 AND CODE = " & MyParn(xCode.text) & " AND DATE >= " & DateSq(XDATESALES1.text) & " AND DATE <= " & DateSq(XDATESALES2.text), con) & "") + Val(xDiscount.text)
nTotalItem = Val(GetDesca("SELECT SUM(FILE6_20.TOTAL) FROM FILE6_20H  INNER JOIN FILE6_20 ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text), con) & "")
nRate = 1 - (Round(nDisc / nTotalItem, 10))
nTotalQuant = Val(GetDesca("SELECT SUM(QUANT) FROM FILE6_20H  INNER JOIN FILE6_20 ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text), con) & "")
cStr1 = "SELECT FILE1_10.MOSM , FACT.DESCA AS FACTNAME , FILE6_20.ITEM , FILE1_10.DESCA , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.COLOR , SUM(FILE6_20.QUANT ) AS QUANT , FILE1_10.MODELFACT0 , SUM(FILE6_20.TOTAL) AS TOTAL  , FILE1_10.BARCODE_GS1 FROM FILE1_10 INNER JOIN FILE6_20 ON FILE6_20.ITEM = FILE1_10.ITEM  INNER JOIN FACT ON FACT.CODE = FILE1_10.FACT INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO  WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text) & " GROUP BY FILE6_20.ITEM , FILE1_10.DESCA , FILE1_10.SCAL , FILE1_10.C_SCAL , FILE1_10.COLOR , FILE1_10.MODELFACT0 , FILE1_10.BARCODE_GS1 , FILE1_10.MOSM , FACT.DESCA "
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText

'               0           1           2       3           4           5           6           7           8        9      10                  11          12              13              14          15
'.FormatString = "„.|" & "»«—ﬂÊœ|" & "„Ê”„|" & "„’‰⁄|" & "„ﬂ »|" & "—ﬁ„ „ÊœÌ·|" & "«·’‰›|" & "„ﬁ«”|" & "«··Ê‰|" & "⁄œœ|" & "”⁄— «·»Ì⁄|" & "«·≈Ã„«·Ï|" & "”⁄— „” Â·ﬂ|" & "‰”»… Œ’„|" & "—ﬁ„ «·„ÊœÌ·|" & "BARCODE GS1|"
With sourcetable
i = 0
Do While Not .EOF
    If !Quant <> 0 Then
        i = i + 1
        grid1.TextMatrix(i, 0) = i
        grid1.TextMatrix(i, 1) = !Item
        grid1.TextMatrix(i, 2) = !MOSM
        grid1.TextMatrix(i, 3) = !FACTNAME
        grid1.TextMatrix(i, 5) = !modelfact0
        grid1.TextMatrix(i, 6) = !DESCA
        grid1.TextMatrix(i, 7) = !scal
        grid1.TextMatrix(i, 8) = !color
        grid1.TextMatrix(i, 9) = !Quant
        grid1.TextMatrix(i, 10) = Round((((!TOTAL / !Quant) * nRate) / 1.14), 2)
        grid1.TextMatrix(i, 11) = Round((((!TOTAL / !Quant) * nRate) / 1.14), 2) * !Quant
        grid1.TextMatrix(i, 15) = !BARCODE_GS1 & ""
        grid1.AddItem ""
    End If
    .MoveNext
Loop
xTax.text = Round((nTotalItem * nRate), 2) - Round((nTotalItem * nRate) / 1.14, 2)
xTotalQuant.Caption = nTotalQuant
xTotal.text = Round(nTotalItem, 2) - Val(xDiscount.text)
xTotDisc.Caption = Round(nTotalItem, 2) - Val(xDiscount.text) - Val(xTax.text)
'grid1.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = 7
End With
End Sub

Private Sub xPO_NO_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then OrderNo_LooK
End Sub
Sub OrderNo_LooK()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(3, 1)

Set Generalarray(0) = Me
Generalarray(1) = "SELECT DOC_NO,DATE, CONVERT(VARCHAR(10),[DATE],111),NOTES " & _
                  " FROM FILE6_51H WHERE doc_sales IS NULL AND ISPOSTED = 0 "
If cFilter <> "" Then
    Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & cFilter
End If

Generalarray(2) = "Order by Date , DOC_NO "
Generalarray(3) = 4200
Generalarray(5) = False

listarray(0, 0) = "«·—ﬁ„-«· «—ÌŒ"
listarray(0, 1) = "(@@Doc_No@@6 OR " & _
                  " ##[DATE]##)"

GrdArray(0, 0) = "—ﬁ„ «·„” ‰œ"
GrdArray(0, 1) = 1200

GrdArray(1, 0) = "«· «—ÌŒ"
GrdArray(1, 1) = 0

GrdArray(2, 0) = "«· «—ÌŒ"
GrdArray(2, 1) = 1500

GrdArray(3, 0) = "«·»Ì«‰"
GrdArray(3, 1) = 2000


searchArray = Array(Generalarray, listarray, GrdArray)
Load oSearchDoc
oSearchDoc.Caption = "«” ⁄·«„"
oSearchDoc.Show 1
End Sub
Private Sub myUndo()
If IsNumeric(xDoc_No.text) Then
    openCardTable xDoc_No.text
    If Not CardTable.EOF Then
        myload
        Exit Sub
    End If
End If
openCardTable , "<"
If CardTable.EOF Then myDefine Else myload
Exit Sub
myerror:
MsgBox Err.Description
Err.Clear
End Sub

Private Sub CmdNext_Click()
openCardTable xDoc_No.text, ">"
If CardTable.EOF Then openCardTable xDoc_No.text
myload
End Sub
Private Sub CmdPrevious_Click()
openCardTable xDoc_No.text, "<"
If CardTable.EOF Then openCardTable xDoc_No.text
myload
End Sub
Private Sub CmdFirst_Click()
openCardTable , ">"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub
Private Sub CmdLast_Click()
openCardTable , "<"
If Not CardTable.EOF Then
    myload
Else
    myDefine
End If
End Sub
Private Sub ShowInv_Grid2()
Dim sourcetable As New ADODB.Recordset
Dim lPrice As Boolean
Dim nFBal As Double
Dim nPay As Double
Dim nBal As Double, nDisc As Double, nTotalItem As Double, nRate As Double, nTotalQuant As Double
Dim cAddres As String, cTaxNo As String

nDisc = Val(GetDesca("SELECT SUM(DISCOUNT) FROM FILE6_20H  WHERE ISINVOICE = 0 AND CODE = " & MyParn(xCode.text) & " AND DATE >= " & DateSq(XDATESALES1.text) & " AND DATE <= " & DateSq(XDATESALES2.text), con) & "") + Val(xDiscount.text)
nTotalItem = Val(GetDesca("SELECT SUM(FILE6_20.TOTAL) FROM FILE6_20H  INNER JOIN FILE6_20 ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text), con) & "")
nRate = 1 - (Round(nDisc / nTotalItem, 10))
nTotalQuant = Val(GetDesca("SELECT SUM(QUANT) FROM FILE6_20H  INNER JOIN FILE6_20 ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO WHERE FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text), con) & "")
cStr1 = "SELECT FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10SC.BARCODE_GS1_SEC , SUM(FILE6_20.QUANT ) AS QUANT , ROUND(SUM(FILE6_20.TOTAL)/SUM(FILE6_20.QUANT),2)  , SUM(FILE6_20.TOTAL) AS TOTAL  FROM FILE1_10 INNER JOIN FILE6_20 ON FILE6_20.ITEM = FILE1_10.ITEM  INNER JOIN FILE1_10SC ON FILE1_10SC.CODE = FILE1_10.[SECTION] INNER JOIN FILE6_20H ON FILE6_20.DOC_NO = FILE6_20H.DOC_NO  WHERE QUANT <> 0 AND FILE6_20H.CODE = " & MyParn(xCode.text) & " AND FILE6_20H.DATE >= " & DateSq(XDATESALES1.text) & " AND FILE6_20H.DATE <= " & DateSq(XDATESALES2.text) & " GROUP BY FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10SC.BARCODE_GS1_SEC HAVING SUM(FILE6_20.QUANT) <> 0 "
sourcetable.Open cStr1, con, adOpenStatic, adLockReadOnly, adCmdText


With sourcetable
i = 0
    GRID2.Rows = 2
Do While Not .EOF
    If !Quant <> 0 Then
        i = i + 1
        GRID2.TextMatrix(i, 0) = !CODE
        GRID2.TextMatrix(i, 1) = !DESCA
        GRID2.TextMatrix(i, 2) = !BARCODE_GS1_SEC & ""
        GRID2.TextMatrix(i, 3) = !Quant
        GRID2.TextMatrix(i, 4) = Round((((!TOTAL / !Quant) * nRate) / 1.14), 4)
        GRID2.TextMatrix(i, 5) = Round((((!TOTAL / !Quant) * nRate) / 1.14), 4) * !Quant
        GRID2.AddItem ""
    End If
    .MoveNext
Loop
GRID2.Cell(flexcpAlignment, 1, 0, GRID2.Rows - 1, GRID2.Cols - 1) = 7
End With
Fixgrd2
End Sub
Private Sub Fixgrd2()
With GRID2
.FixedCols = 0
.Cols = 6
'                0        1         2           3               4           5       6
.FormatString = "ﬂÊœ|" & "«·ﬁ”„|" & "BARCODE GS1|" & "«·⁄œœ|" & "«·”⁄—|" & "«·≈Ã„«·Ï"
.RowHeight(0) = 800
.WordWrap = True
.ColWidth(0) = 1000
.ColWidth(1) = 3000
.ColWidth(2) = 2000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1500
.Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'.Cell(flexcpAlignment, 1, 0, .Rows - 1, .Cols - 1) = flexAlignRightCenter
.ShowCell .Rows - 1, 0
End With
End Sub
Private Sub myloadgrd2()
    Dim cString As String
'                           0               1                2                              3                                   4                                   5
    cString = "SELECT FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10SC.BARCODE_GS1_SEC , SUM(FILE6_20.QUANT) , ROUND(SUM(FILE6_20.TOTAL)/SUM(FILE6_20.QUANT) ,2), SUM(FILE6_20.TOTAL ) " & _
          " FROM (FILE6_20 INNER JOIN FILE1_10 ON FILE6_20.ITEM = FILE1_10.ITEM) INNER JOIN FILE1_10SC ON FILE1_10.[SECTION] = FILE1_10SC.CODE"
    cString = cString & " where FILE6_20.QUANT <> 0 and DOC_NO = " & MyParn(xDoc_No.text)
    cString = cString & " GROUP BY FILE1_10SC.CODE , FILE1_10SC.DESCA , FILE1_10SC.BARCODE_GS1_SEC  ORDER BY FILE1_10SC.CODE "
    Set DATA22.Recordset = myRecordSet(cString, con)
    Fixgrd2
End Sub
