VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form TDaySal 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   10230
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   10230
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "»Ê‰« "
      Height          =   240
      Left            =   2430
      RightToLeft     =   -1  'True
      TabIndex        =   32
      Top             =   270
      Width           =   825
   End
   Begin VB.TextBox XDEL_DOC 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Left            =   3330
      PasswordChar    =   "*"
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   180
      Width           =   1545
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈⁄«œ…  ÕœÌÀ «·ÌÊ„ ··›—⁄"
      Height          =   465
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   450
      Width           =   2265
   End
   Begin MSComctlLib.ProgressBar PROG1 
      Height          =   345
      Left            =   45
      TabIndex        =   25
      Top             =   945
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FFFF&
      Caption         =   "„ﬁ«—‰… »Ê‰«  «·»Ì⁄ ··›—⁄"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   0
      Width           =   2265
   End
   Begin VB.Frame Frame4 
      Height          =   690
      Left            =   4950
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   90
      Width           =   2535
      Begin VB.CommandButton cmdExit 
         Height          =   510
         Left            =   90
         Picture         =   "TDaySal.frx":0000
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   135
         Width           =   1140
      End
      Begin Threed.SSCommand cmdPrint 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   1260
         TabIndex        =   14
         Top             =   135
         Visible         =   0   'False
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   900
         _Version        =   196610
         CaptionStyle    =   1
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
         Picture         =   "TDaySal.frx":246C
         Caption         =   "ÿ»«⁄… «·ÌÊ„ "
         Alignment       =   4
         PictureAlignment=   10
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   735
      Left            =   7560
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   45
      Width           =   7485
      Begin VB.CommandButton cmdGo 
         Height          =   450
         Left            =   90
         Picture         =   "TDaySal.frx":48A6
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox xdate2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   1770
      End
      Begin VB.TextBox xDate1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   4590
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   1815
      End
      Begin VB.Label Label4 
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
         Left            =   3375
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "„‰  «—ÌŒ :"
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
         Left            =   6525
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   315
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3300
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   810
      Width           =   11535
      Begin VSFlex7Ctl.VSFlexGrid GridBox 
         Bindings        =   "TDaySal.frx":6D98
         Height          =   2985
         Left            =   90
         TabIndex        =   2
         Top             =   135
         Width           =   11355
         _cx             =   20029
         _cy             =   5265
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         TabBehavior     =   0
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
   Begin MSComCtl2.MonthView xDate 
      Height          =   2610
      Left            =   0
      TabIndex        =   0
      Top             =   1350
      Width           =   3510
      _ExtentX        =   6191
      _ExtentY        =   4604
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   12632256
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      StartOfWeek     =   144506887
      TitleBackColor  =   12632256
      TrailingForeColor=   12632256
      CurrentDate     =   40954
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   2340
      Top             =   630
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
      Left            =   1710
      Top             =   -630
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   1710
      Top             =   -1035
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
   Begin TabDlg.SSTab SSTAB 
      Height          =   5745
      Left            =   45
      TabIndex        =   3
      Top             =   4095
      Width           =   14970
      _ExtentX        =   26405
      _ExtentY        =   10134
      _Version        =   393216
      Tabs            =   7
      Tab             =   1
      TabsPerRow      =   7
      TabHeight       =   520
      TabMaxWidth     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "≈Ã„«·Ì »Ê‰«  «·»Ì⁄ "
      TabPicture(0)   =   "TDaySal.frx":6DAB
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "PRINT1"
      Tab(0).Control(1)=   "Grid1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "≈Ã„«·Ì „»Ì⁄«  «·«’‰«›"
      TabPicture(1)   =   "TDaySal.frx":6DC7
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Picture1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "SSCommand1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "grid2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      TabCaption(2)   =   " ›’Ì·Ì »Ê‰«  «·»Ì⁄"
      TabPicture(2)   =   "TDaySal.frx":6DE3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture5"
      Tab(2).Control(1)=   "PRINT4"
      Tab(2).Control(2)=   "GRID3"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "„»Ì⁄«  »«·”«⁄…"
      TabPicture(3)   =   "TDaySal.frx":6DFF
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GRIDHOUR"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "«·„” Âœ› «·‘Â—Ï"
      TabPicture(4)   =   "TDaySal.frx":6E1B
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame3"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "≈Ã„«·Ï „»Ì⁄«  ÌÊ„Ï - ›—Ê⁄"
      TabPicture(5)   =   "TDaySal.frx":6E37
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "GRID5"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "≈Ã„«·Ï ›Ì“«"
      TabPicture(6)   =   "TDaySal.frx":6E53
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "GRID6"
      Tab(6).ControlCount=   1
      Begin VB.Frame Frame3 
         Caption         =   "«·„” Âœ› «·‘Â—Ï"
         Height          =   1995
         Left            =   -74955
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   330
         Width           =   14595
         Begin VSFlex7Ctl.VSFlexGrid GridTarget 
            Height          =   1410
            Left            =   45
            TabIndex        =   28
            Top             =   360
            Width           =   14370
            _cx             =   25347
            _cy             =   2487
            _ConvInfo       =   1
            Appearance      =   0
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   14.25
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
      Begin VSFlex7Ctl.VSFlexGrid GRID3 
         Height          =   5145
         Left            =   -74775
         TabIndex        =   4
         Top             =   420
         Width           =   14640
         _cx             =   25823
         _cy             =   9075
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
      Begin VSFlex7Ctl.VSFlexGrid grid2 
         Height          =   5145
         Left            =   135
         TabIndex        =   5
         Top             =   420
         Width           =   14685
         _cx             =   25903
         _cy             =   9075
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
      Begin VSFlex7Ctl.VSFlexGrid Grid1 
         Height          =   4875
         Left            =   -74910
         TabIndex        =   6
         Top             =   465
         Width           =   14775
         _cx             =   26061
         _cy             =   8599
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
      Begin Threed.SSCommand PRINT1 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   -74910
         TabIndex        =   16
         Top             =   5865
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   900
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TDaySal.frx":6E6F
         Alignment       =   4
         PictureAlignment=   10
      End
      Begin Threed.SSCommand PRINT4 
         CausesValidation=   0   'False
         Height          =   420
         Left            =   -74910
         TabIndex        =   17
         Top             =   4965
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   741
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TDaySal.frx":92A9
         Alignment       =   4
         PictureAlignment=   10
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   2355
         Left            =   -74955
         Picture         =   "TDaySal.frx":B6E3
         RightToLeft     =   -1  'True
         ScaleHeight     =   2355
         ScaleWidth      =   3210
         TabIndex        =   18
         Top             =   3255
         Width           =   3210
         Begin ImgXCtrl6.ImgXCtrl imgx3 
            Height          =   2175
            Left            =   45
            TabIndex        =   19
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
         End
      End
      Begin Threed.SSCommand SSCommand1 
         CausesValidation=   0   'False
         Height          =   510
         Left            =   90
         TabIndex        =   20
         Top             =   5595
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   900
         _Version        =   196610
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "TDaySal.frx":EDBA
         Alignment       =   4
         PictureAlignment=   10
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   2355
         Left            =   45
         Picture         =   "TDaySal.frx":111F4
         RightToLeft     =   -1  'True
         ScaleHeight     =   2355
         ScaleWidth      =   3210
         TabIndex        =   21
         Top             =   3210
         Width           =   3210
         Begin ImgXCtrl6.ImgXCtrl Imgx2 
            Height          =   2175
            Left            =   45
            TabIndex        =   22
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
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid GRIDHOUR 
         Height          =   5145
         Left            =   -71445
         TabIndex        =   23
         Top             =   420
         Width           =   11310
         _cx             =   19950
         _cy             =   9075
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
      Begin VSFlex7Ctl.VSFlexGrid GRID5 
         Height          =   5190
         Left            =   -74910
         TabIndex        =   29
         Top             =   405
         Width           =   14775
         _cx             =   26061
         _cy             =   9155
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
      Begin VSFlex7Ctl.VSFlexGrid GRID6 
         Height          =   5055
         Left            =   -74910
         TabIndex        =   30
         Top             =   405
         Width           =   14775
         _cx             =   26061
         _cy             =   8916
         _ConvInfo       =   1
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
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
   Begin MSAdodcLib.Adodc DATABOX 
      Height          =   330
      Left            =   1665
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
   Begin MSAdodcLib.Adodc DATAHOUR 
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
   Begin MSAdodcLib.Adodc DATA5 
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
   Begin MSAdodcLib.Adodc DATA6 
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
Attribute VB_Name = "TDaySal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim conShop As New ADODB.Connection
Dim con As New ADODB.Connection

Private Sub CmdExit_Click()
    Unload Me
End Sub
Private Sub cmdPrint_Click()
    doprint1
End Sub
Private Sub cmdGo_Click()
        
    myloadBox
    MYLOAD1 cBranchBox
    If Check2.Value = 0 Then
        MYLOAD2 cBranchBox
        MYLOAD4 cBranchBox
        myload5 cBranchBox
        myload6 cBranchBox
        myloadHour cBranchBox
    End If
End Sub
Private Sub Command1_Click()
    If cBranch <> "00" Then
        con.Execute " update file6_20h set isnew = 1 where date = " & DateSq(xDate.Value), nRec
        Inform " „»Ì⁄«  ⁄œœ " & nRec
    End If
End Sub
Private Sub Form_Load()
    Dim nTarget As Double, nSal1 As Double
    SSTAB.Tab = 2
    openCon con
    xDate.Value = dSalesDate
    
    XDEL_DOC.Visible = Not lServerOnLine And (cBranch = "00")
        
    If cUserBox = "" Then cBranchBox = ""
    GridTarget.Visible = (cBranch <> "00")
    Set grid1.DataSource = data1
    Set GRID2.DataSource = DATA2
    Set GridBox.DataSource = DATABOX
    Set GRIDHOUR.DataSource = DATAHOUR
    
    Set GRID5.DataSource = data5
    Set grid6.DataSource = DATA6
    
    xDate.Visible = bopt3
    Frame1.Visible = bopt3
    
    Command1.Visible = lSupperVisor
    
    grid1.Rows = 1
    GRID2.Rows = 1
    GRID3.Rows = 1
    GridBox.Rows = 1
    GRID5.Rows = 1
    grid6.Rows = 1
    
    Fixgrd1
    Fixgrd2
    Fixgrd4
    fixGrdBox
    FIXGRID5
    FIXGRID6

    xDate1.text = xDate.Value
    xdate2.text = xDate.Value
    myloadBox
    Check1.Visible = (cBranch = "00")
    Frame3.Visible = (cBranch <> "00")
    If cBranch <> "00" Then
    With GridTarget
        .Rows = 2
        .Cols = 7
        .WordWrap = True
        .RowHeight(0) = 900
        .FixedRows = 1
        .FixedCols = 0
        .Editable = flexEDNone
        .TextMatrix(0, 0) = "‰”»… «·ÌÊ„"
        .TextMatrix(0, 1) = "«·„” Âœ› «·‘Â—Ï"
        .TextMatrix(0, 2) = "«·„»Ì⁄« "
        .TextMatrix(0, 3) = "«·›—ﬁ"
        .TextMatrix(0, 4) = "‰”»… "
        .TextMatrix(0, 5) = "›—ﬁ «·‰”»… "
        .TextMatrix(0, 6) = "„” Âœ› ÌÊ„ÌÏ"
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1500
        .ColWidth(6) = 1500
        
        If lIsBranchStore Then
            nTarget = Val(GetDesca("SELECT VALUE_TARGET FROM Q_FILE6_80 WHERE BRANCH = " & MyParn(cBranchStore) & " AND MONTH(DATE1) = " & Month(dSalesDate) & " AND YEAR(DATE1) = " & Year(dSalesDate), con) & "")
            nSal1 = Round(Val(GetDesca("SELECT  SUM(T_TOTAL - DISCOUNT)   FROM T_SALESDOC WHERE T_SALESDOC.STORE =  " & MyParn(cBranchStore) & " AND  MONTH(DATE) = " & Month(dSalesDate) & " AND YEAR(DATE) = " & Year(dSalesDate), con) & ""), 2)
        Else
            nTarget = Val(GetDesca("SELECT VALUE_TARGET FROM FILE6_80 WHERE BRANCH = " & MyParn(cBranchStore) & " AND MONTH(DATE1) = " & Month(dSalesDate) & " AND YEAR(DATE1) = " & Year(dSalesDate), con) & "")
            nSal1 = Round(Val(GetDesca("SELECT  SUM(T_TOTAL - DISCOUNT)   FROM T_SALESDOC WHERE T_SALESDOC.STORE =  " & MyParn(cBranchStore) & " AND  MONTH(DATE) = " & Month(dSalesDate) & " AND YEAR(DATE) = " & Year(dSalesDate), con) & ""), 2)
        End If

        nMaxDay = LastDayOfMonth(DateValue(dSalesDate))
        nRateDay = Round((Day(Date) / nMaxDay) * 100, 2)
        .TextMatrix(1, 0) = nRateDay
        .TextMatrix(1, 1) = nTarget
        .TextMatrix(1, 2) = nSal1

        If nTarget > 0 Then
            .TextMatrix(1, 3) = nTarget - nSal1
            .TextMatrix(1, 4) = Round((nSal1 / nTarget) * 100, 2)
            .TextMatrix(1, 5) = Round(Val(.TextMatrix(1, 4)) - nRateDay, 2)
            If (nMaxDay - Day(Date)) > 0 Then
                .TextMatrix(1, 6) = Round((nTarget - nSal1) / ((nMaxDay - Day(Date))))
            End If
        End If
        .ColFormat(1) = "#,###,##0"
        .ColFormat(2) = "#,###,##0"
        .ColFormat(3) = "#,###,##0"
    End With
    End If
End Sub
Private Sub Fixgrd2()
With GRID2
.Cols = 8
.TextMatrix(0, 0) = "„’‰⁄"
.TextMatrix(0, 1) = "„Ê—œ"
.TextMatrix(0, 2) = "„Ê”„"
.TextMatrix(0, 3) = "—ﬁ„ „ÊœÌ·"
.TextMatrix(0, 4) = "«·’‰›"
.TextMatrix(0, 5) = "⁄œœ „»Ì⁄« "
.TextMatrix(0, 6) = "ﬁÌ„… „»Ì⁄« "
.FixedCols = 0
.FixedRows = 1
.ColHidden(.Cols - 1) = True

For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 3) = DelZero(.TextMatrix(nRow, 3))
Next nRow

.ColWidth(0) = 2000
.ColWidth(1) = 2000
.ColWidth(2) = 700
.ColWidth(3) = 2000
.ColWidth(4) = 5000
.ColWidth(5) = 800
.ColWidth(6) = 1500

.MergeCells = flexMergeFree
.MergeCol(0) = True
.MergeCol(1) = True
.MergeCol(2) = True

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTBelow
For nCol = 5 To .Cols - 1
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
    .Subtotal flexSTSum, 0, nCol, "#0.00", vbYellow, , True, " "
Next nCol
End With
End Sub
Private Sub Fixgrd1()
With grid1
.Cols = 20
.TextMatrix(0, 0) = "—ﬁ„ «·„” ‰œ"
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 2) = "«·Êﬁ "
.TextMatrix(0, 3) = "«·Œ“‰…"
.TextMatrix(0, 4) = "⁄œœ «·ﬁÿ⁄"
.TextMatrix(0, 5) = "≈Ã„«·Ï «·ﬁÌ„…"
.TextMatrix(0, 6) = "«·Œ’„"
.TextMatrix(0, 7) = "«·’«›Ï"
.TextMatrix(0, 8) = "”œ«œ ‰ﬁœÏ"
.TextMatrix(0, 9) = "”œ«œ ›Ì“«"
.TextMatrix(0, 10) = "point"
.TextMatrix(0, 11) = "›—ﬁ"
.TextMatrix(0, 12) = "⁄„Ê·… ›Ì“«"
.TextMatrix(0, 13) = "’·«ÕÌ… Œ’„"
.TextMatrix(0, 14) = "’·«ÕÌ… „— Ã⁄"
.TextMatrix(0, 15) = "„” ‰œ »Ì⁄"
.TextMatrix(0, 16) = "≈Ã„«·Ï ›—⁄"
.TextMatrix(0, 18) = " ·Ì›Ê‰"

.TextMatrix(0, 19) = "≈Ã„«·Ï ›—⁄"

.ColFormat(1) = "DD-MM-YYYY"
.ColFormat(2) = "SHORT TIME"
.ColFormat(4) = "#0"
.ColFormat(5) = "#0.00"
.ColFormat(6) = "#0.00"
.ColFormat(7) = "#0.00"
.ColFormat(8) = "#0.00"
.ColFormat(9) = "#0.00"
.ColFormat(10) = "#0.00"
.ColDataType(1) = flexDTDate
.ColDataType(2) = flexDTShort
.ColDataType(4) = flexDTDouble
.ColDataType(5) = flexDTDouble
.ColDataType(6) = flexDTDouble
.ColDataType(7) = flexDTDouble
.ColDataType(8) = flexDTDouble
.ColDataType(9) = flexDTDouble
.ColDataType(10) = flexDTDouble
.ColDataType(14) = flexDTDouble

.ColHidden(16) = True
.ColHidden(17) = True

.ColWidth(0) = 1200
.ColWidth(1) = 0
.ColWidth(2) = 1000
.ColWidth(3) = 1200
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 1000
.ColWidth(11) = 1000
.ColWidth(12) = 1000

.ColWidth(13) = 1500
.ColWidth(14) = 1500
.ColWidth(15) = 2200

.ColWidth(16) = 1800
.ColWidth(18) = 1500

.ColWidth(19) = 1100

.FixedCols = 0
.FixedRows = 1

For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
For nRow = 1 To .Rows - 1
   .TextMatrix(nRow, 11) = Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)) - Val(.TextMatrix(nRow, 9)) - Val(.TextMatrix(nRow, 10))
    .TextMatrix(nRow, 1) = DateValue(.TextMatrix(nRow, 1))
    .TextMatrix(nRow, 2) = TimeValue(.TextMatrix(nRow, 2))
Next nRow

.ExplorerBar = flexExSortShow
.Subtotal flexSTSum, -1, 4, "#0", vbYellow, , True, " "
For nCol = 5 To 11
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
Next nCol
End With
End Sub
Private Sub fixGrdBox()
With GridBox
.FixedCols = 0
.FixedRows = 1
.Cols = 13
.RowHeight(0) = 1000
.WordWrap = True
If cBranch = "00" Then
    .TextMatrix(0, 0) = "«·›—⁄"
    .TextMatrix(0, 1) = "«· «—ÌŒ"
    .TextMatrix(0, 2) = "«·›—⁄"
Else
    .TextMatrix(0, 0) = "«·Œ“‰…"
    .TextMatrix(0, 1) = "«· «—ÌŒ"
    .TextMatrix(0, 2) = "«·›—⁄"
End If
.TextMatrix(0, 3) = "⁄œœ »Ê‰« "
.TextMatrix(0, 4) = "⁄œœ ﬁÿ⁄ „»Ì⁄« "
.TextMatrix(0, 5) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 6) = "«·Œ’„"
.TextMatrix(0, 7) = "«·’«›Ï"
.TextMatrix(0, 8) = "‰ﬁœÏ"
.TextMatrix(0, 9) = "›Ì“«"
.TextMatrix(0, 10) = "point"
.TextMatrix(0, 11) = "«·›—ﬁ"
.TextMatrix(0, 12) = "⁄„Ê·… ›Ì“«"
'.RowHeight(0) = 800
'.WordWrap = True
.ColWidth(1) = 1100
.ColWidth(2) = 1200
.ColWidth(3) = 600
.ColWidth(4) = 600
.ColWidth(5) = 1000
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1000
.ColWidth(9) = 1000
.ColWidth(10) = 900
.ColWidth(11) = 800
.ColWidth(12) = 800

.ColHidden(0) = True

For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 11) = Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)) - Val(.TextMatrix(nRow, 9)) - Val(.TextMatrix(nRow, 10))
    For nCol = 5 To .Cols - 1
        .TextMatrix(nRow, nCol) = Format(Val(.TextMatrix(nRow, nCol)), "#0.00")
    Next nCol
Next nRow
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTAbove
For nCol = 3 To .Cols - 1
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
Next nCol
.Subtotal flexSTSum, -1, 3, "#0", vbYellow, , True, " "
.Subtotal flexSTSum, -1, 4, "#0", vbYellow, , True, " "
End With

End Sub
Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
With grid1
    If KeyCode = 46 And cBranch = "00" And Check1.Value <> 0 And XDEL_DOC.text <> "654321" And lServerOnLine Then
        If MsgBox(" Õ–› »Ê‰ «·„»Ì⁄«   ", vbYesNo + vbDefaultButton2) = vbYes Then
            con.Execute " delete from file6_20  where doc_no =  " & MyParn(.TextMatrix(.Row, 17))
            con.Execute " delete from file6_20h  where doc_no =  " & MyParn(.TextMatrix(.Row, 17))
            conShop.Execute " update file6_20h  set isnew = 1 where doc_no =  " & MyParn(.TextMatrix(.Row, 17))
            Inform "  „ Õ–› «·„” ‰œ "
        End If
    End If

    If KeyCode = 46 And cBranch = "00" And XDEL_DOC.text = "654321" And Not lServerOnLine Then
        con.Execute " delete from file6_20  where doc_no =  " & MyParn(.TextMatrix(.Row, 17))
        con.Execute " delete from file6_20h  where doc_no =  " & MyParn(.TextMatrix(.Row, 17))
        .RemoveItem .Row
'        MYLOAD1
'        Grid1.Select 0, 7, 0, 7
'        Grid1.Sort = flexSortGenericDescending
'        Grid1.Select 2, 7, 2, 7
     End If
End With

End Sub
Private Sub grid2_EnterCell()
    MyLoadPhoto Mid(GRID2.TextMatrix(GRID2.Row, GRID2.Cols - 1), 1, 10), Imgx2
End Sub
Private Sub GRID3_EnterCell()
    MyLoadPhoto Mid(GRID3.TextMatrix(GRID3.Row, 3), 1, 10), imgx3
End Sub
Private Sub gridbox_DblClick()
If grid1.Rows > 1 Then
    If cBranch = "00" Then
        MYLOAD1 "", GridBox.TextMatrix(GridBox.Row, 0)
        If Check2.Value = 0 Then
        MYLOAD2 "", GridBox.TextMatrix(GridBox.Row, 0)
        MYLOAD4 "", GridBox.TextMatrix(GridBox.Row, 0)
        myload5 "", GridBox.TextMatrix(GridBox.Row, 0)
        myload6 "", GridBox.TextMatrix(GridBox.Row, 0)
        myloadHour GridBox.TextMatrix(GridBox.Row, 0)
        End If
    Else
        MYLOAD1 GridBox.TextMatrix(GridBox.Row, 0)
        If Check2.Value = 0 Then
        MYLOAD2 GridBox.TextMatrix(GridBox.Row, 0)
        MYLOAD4 GridBox.TextMatrix(GridBox.Row, 0)
        myload5 GridBox.TextMatrix(GridBox.Row, 0)
        myload6 GridBox.TextMatrix(GridBox.Row, 0)
        myloadHour GridBox.TextMatrix(GridBox.Row, 0)
        End If
    End If
End If
End Sub

Private Sub Imgx2_DblClick()
If validPhoto(retFile(Mid(GRID2.TextMatrix(GRID2.Row, GRID2.Cols - 1), 1, 10))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(Mid(GRID2.TextMatrix(GRID2.Row, GRID2.Cols - 1), 1, 10))
    oMaxPhoto.cDesca = GRID2.TextMatrix(GRID2.Row, 4)
    oMaxPhoto.Show 1
End If

End Sub

Private Sub imgx3_DblClick()
If validPhoto(retFile(Mid(GRID3.TextMatrix(GRID3.Row, 3), 1, 10))) Then
    Dim oMaxPhoto As New maxPhoto
    oMaxPhoto.cFile = retFile(Mid(GRID3.TextMatrix(GRID3.Row, 3), 1, 10))
    oMaxPhoto.cDesca = GRID3.TextMatrix(GRID3.Row, 6)
    oMaxPhoto.Show 1
End If
End Sub

Private Sub PRINT1_Click()
Load PrintGrd
PrintGrd.doprint grid1, , , "≈Ã„«·Ï »Ê‰«  «·»Ì⁄ ··› —… „‰ " & xDate1.text & " Õ Ï " & xdate2.text, , , , True
PrintGrd.Show 1
End Sub

Private Sub print2_Click()
Load PrintGrd
PrintGrd.doprint GRID2, , , "≈Ã„«·Ï „»Ì⁄«  «·„ÊœÌ·«  ··› —… „‰ " & xDate1.text & " Õ Ï " & xdate2.text, , , , True
PrintGrd.Show 1
End Sub
Private Sub PRINT4_Click()
    PrintGrd.doprint GRID3, , , " ›’Ì·Ì »Ê‰«  «·»Ì⁄ " & xDate1.text & " Õ Ï " & xdate2.text, , , , True
    PrintGrd.Show 1
End Sub
Private Sub SSCommand2_Click()
    doprint_day
End Sub
Private Sub xdate_DateClick(ByVal DateClicked As Date)
    xDate1.text = Format(xDate.Value, "DD-MM-YYYY")
    xdate2.text = Format(xDate.Value, "DD-MM-YYYY")
    myloadBox
End Sub
Sub myloadBox()
Dim cStr1 As String
If cBranch = "00" Then
    cStr1 = "SELECT t_Salesdoc.STORE , date , STORENAME , COUNT(DOC_NO) , SUM(T_QUANT) , SUM(T_TOTAL) , SUM(discount) , SUM(T_TOTAL-discount), SUM(CASH) , SUM(VISA+VISA3) , SUM(point) ,0, SUM(visa2+visa3_2)  FROM  T_SALESDOC"
Else
    cStr1 = "SELECT t_Salesdoc.box,date,boxname, COUNT(DOC_NO) , SUM(T_QUANT) , SUM(T_TOTAL) , SUM(discount) , SUM(T_TOTAL-discount), SUM(CASH) , SUM(VISA+VISA3), SUM(point),0 ,SUM(VISA2+visa3_2) FROM  T_SALESDOC"
    If lIsBranchStore Then cStr1 = cStr1 & turn(cStr1) & "BRANCH = " & MyParn(cBranch)
End If
cStr1 = cStr1 & turn(cStr1) & "DATE >= " & DateSq(xDate1.text)
cStr1 = cStr1 & turn(cStr1) & "DATE <= " & DateSq(xdate2.text)
If Not bOpt5 Then cStr1 = cStr1 & " AND t_Salesdoc.[store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "

If cBranch <> "00" Then
    If cBranchBox <> "" Then cStr1 = cStr1 & turn(cStr1) & "box = " & MyParn(cBranchBox)
End If
If cBranch <> "00" Then
    cStr1 = cStr1 & " GROUP BY date,boxname,box ORDER BY BOX "
Else
    cStr1 = cStr1 & " GROUP BY date,STORENAME , STORE ORDER BY STORE "
End If
Set DATABOX.Recordset = myRecordSet(cStr1, con)
fixGrdBox

MYLOAD1 cBranchBox
If Check2.Value = 0 Then
    MYLOAD2 cBranchBox
    MYLOAD4 cBranchBox
    myload5 cBranchBox
    myload6 cBranchBox
    myloadHour cBranchBox
End If
End Sub
Sub MYLOAD2(Optional cBox As String, Optional cStore As String)
    Dim cString As String
    cString = "SELECT  FACTDESCA, SUPPDESCA,FILE1_10.MOSM , FILE1_10.MODELFACT, FILE1_10.desca, SUM(QUANT) AS T_QTY, SUM(TOTAL) AS T_TOTAL,FILE1_10.MODELNO FROM SALES_MODEL INNER JOIN FILE1_10 ON SALES_MODEL.ITEM = FILE1_10.ITEM"
    If lIsBranchStore Then cString = cString & turn(cString) & "BRANCH = " & MyParn(cBranch)
    If IsDate(xDate1.text) Then cString = cString & turn(cString) & "DATE >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cString = cString & turn(cString) & "DATE <= " & DateSq(xdate2.text)
    If cBox <> "" Then cString = cString & turn(cString) & "SALES_MODEL.BOX = " & MyParn(cBox)
    If cStore <> "" Then cString = cString & turn(cString) & "SALES_MODEL.STORE = " & MyParn(cStore)
    If Not bOpt5 Then cString = cString & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    cString = cString & " GROUP BY FILE1_10.MOSM, SUPPDESCA, FACTDESCA, FILE1_10.MODELFACT, FILE1_10.desca,MODELNO order by SUPPDESCA, FACTDESCA , FILE1_10.MOSM,FILE1_10.MODELFACT"
    Set DATA2.Recordset = myRecordSet(cString, con)
    Fixgrd2
End Sub
Sub MYLOAD1(Optional cBox As String, Optional cStore As String)
    Dim cString As String
'                           0                   1               2               3                   4                   5                   6                   7                                           8           9               10            11    12                     13              14               15                16            17
    cString = "SELECT FILE6_20H.DOC_NO2 ,FILE6_20H.date ,FILE6_20H.TIME, FILE0_50.DESCA , SUM(FILE6_20.QUANT) , SUM(FILE6_20.TOTAL) ,FILE6_20H.discount , SUM(FILE6_20.TOTAL) - FILE6_20H.discount , FILE6_20H.CASH ,FILE6_20H.VISA+VISA3 ,FILE6_20H.point ,0,FILE6_20H.visa2+FILE6_20H.visa3_2, FILE6_20H.username_disc , username_ret , SALES_RET , ' ' as n14 , FILE6_20H.DOC_NO, FILE6_20H.phone" & _
    " FROM (FILE6_20H INNER JOIN FILE6_20 ON FILE6_20H.DOC_NO = FILE6_20.DOC_NO) LEFT JOIN FILE0_50 ON FILE6_20H.BOX = FILE0_50.CODE "
    If lIsBranchStore Then cString = cString & turn(cString) & "FILE6_20H.BRANCH = " & MyParn(cBranch)
    If IsDate(xDate1.text) Then cString = cString & turn(cString) & "DATE >= " & DateSq(xDate1.text)
    If IsDate(xdate2.text) Then cString = cString & turn(cString) & "DATE <= " & DateSq(xdate2.text)
    If cBox <> "" Then cString = cString & turn(cString) & "FILE6_20H.BOX = " & MyParn(cBox)
    If cStore <> "" Then cString = cString & turn(cString) & "FILE6_20H.STORE = " & MyParn(cStore)
    If Not bOpt5 Then cString = cString & " AND FILE6_20H.[store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
    cString = cString & " GROUP BY " & "FILE6_20H.DOC_NO,FILE6_20H.DOC_NO2,FILE6_20H.date,FILE6_20H.TIME,FILE0_50.DESCA,FILE6_20H.DISCOUNT,FILE6_20H.CASH,FILE6_20H.VISA+VISA3,FILE6_20H.VISA2+FILE6_20H.VISA3_2  ,FILE6_20H.point , FILE6_20H.username_disc , username_ret , SALES_RET , FILE6_20H.STORE , FILE6_20H.phone"
    cString = cString & " ORDER BY DATE , TIME,FILE6_20H.DOC_NO2"
    Set data1.Recordset = myRecordSet(cString, con)
    
    Fixgrd1
    With grid1
        If cStore <> "" And Check1.Value <> 0 Then
            pServerIp = GetDesca("select ipsql from branch where store = " & MyParn(cStore), con)
            pServerData = GetDesca("select DATA from branch where store = " & MyParn(cStore), con)
            
            strConShop = LoadConStringshop
            If openConShop(conShop, strConShop) = "ok" Then
                prog1.Visible = True
                prog1.Min = 0
                prog1.Max = .Rows - 1
                prog1.Value = 0
                .ColHidden(14) = False
                For nRow = 2 To .Rows - 1
                    prog1.Value = nRow
                    .TextMatrix(nRow, 19) = Val(GetDesca("SELECT SUM(T_TOTAL-DISCOUNT) FROM T_SALESDOC WHERE DOC_NO = " & MyParn(.TextMatrix(nRow, 15)), conShop) & "")
                    If Val(.TextMatrix(nRow, 14)) <> Val(.TextMatrix(nRow, 7)) Then
                        .Cell(flexcpBackColor, nRow, 0, nRow, .Cols - 1) = &HFF&
                    End If
                Next nRow
            Else
                .ColHidden(14) = True
                Inform "·« ÌÊÃœ « ’«· »«·›—⁄"
            End If
        End If
    End With
End Sub
Private Sub doprint1()
Dim temptable As New ADODB.Recordset
Dim sourcetable As New ADODB.Recordset
ReDim aHeader(1)
contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

cString = "SELECT  SUPPDESCA, FACTDESCA, SUM(QUANT) AS TQUANT, SUM(TOTAL) AS TTOTAL, BOX FROM         SALES_MODEL "
If IsDate(xDate1.text) Then
    cString = cString & turn(cString) & " date >= " & DateSq(xDate1.text)
    aHeader(0) = "[" & BetweenString(Format(xDate1.text, "d-m-yyyy"), Format(xdate2.text, "d-m-yyyy")) & "]"
End If

If IsDate(xdate2.text) Then
    cString = cString & turn(cString) & " date <= " & DateSq(xdate2.text)
     aHeader(0) = "[" & BetweenString(Format(xDate1.text, "d-m-yyyy"), Format(xdate2.text, "d-m-yyyy")) & "]"
End If

If cBox <> "" Then
    cString = cString & turnFound(cString) & " AND BOX = " & MyParn(cBox)
    aHeader(1) = "[" & SalesFrm.xBox.text & "]"
End If
    
cString = cString & " GROUP BY SUPPDESCA, FACTDESCA, BOX "
sourcetable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
If sourcetable.EOF And sourcetable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
With sourcetable
Do Until sourcetable.EOF
    temptable.AddNew
    temptable!str3 = " ÌÊ„Ì… " & xDate.Value
    temptable!str1 = !BOX
    temptable!str2 = GetDesca("SELECT DESCA FROM FILE0_50 WHERE CODE = " & MyParn(!BOX), con)
    temptable!str5 = !SUPPDESCA
    If !SUPPDESCA <> !FACTDESCA Then temptable!str6 = !FACTDESCA
    temptable!VAL1 = !TQUANT
    temptable!val3 = !TTOTAL
    temptable.Update
    sourcetable.MoveNext
Loop
End With
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\TDAY.RPT"
Main.REPORT1.DataFiles(0) = "c:\tempmrshd\temp.mdb"
Main.REPORT1.Action = 1
End Sub
Private Sub MYLOAD4(Optional cBox As String, Optional cStore As String)
Dim cWhere As String, cString As String
If IsDate(xDate1.text) Then cWhere = " AND DATE >= " & DateSq(xDate1.text)
If IsDate(xdate2.text) Then cWhere = cWhere & " AND DATE <= " & DateSq(xdate2.text)
If lIsBranchStore Then cWhere = cWhere & " AND STORE = " & MyParn(cBranchStore)
If cBox <> "" Then cWhere = cWhere & " AND BOX = " & MyParn(cBox)
If cStore <> "" Then cWhere = cWhere & " AND STORE = " & MyParn(cStore)
If Not bOpt5 Then cWhere = cWhere & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "

cString = "SELECT SALESDTL.*  " & _
          " FROM SALESDTL WHERE DOC_NO IS NOT NULL "
If cWhere <> "" Then cString = cString & cWhere
cString = cString & " ORDER BY DOC_NO,FLAG"
Dim loctable As New ADODB.Recordset
loctable.Open cString, con, adOpenKeyset, adLockReadOnly, adCmdText
With GRID3
GRID3.Rows = 1
Do Until loctable.EOF
    .AddItem ""
    If loctable!Flag = 0 Then
        .TextMatrix(.Rows - 1, 0) = Mid(loctable!doc_no, 7, 6)
        .TextMatrix(.Rows - 1, 1) = Format(loctable!Time, "HH:NN")
        .TextMatrix(.Rows - 1, 2) = loctable!MANDESCA & ""
        .TextMatrix(.Rows - 1, 3) = Format(loctable!modelfact0 & "")
        .TextMatrix(.Rows - 1, 4) = loctable!fact_Desca & ""
        .TextMatrix(.Rows - 1, 5) = loctable!Item & ""
        .TextMatrix(.Rows - 1, 6) = loctable!DESCA & ""
        If loctable!Quant > 0 Then
            .TextMatrix(.Rows - 1, 7) = loctable!Quant
        Else
            .TextMatrix(.Rows - 1, 8) = loctable!Quant * -1
        End If
        .TextMatrix(.Rows - 1, 9) = loctable!price & ""
        .TextMatrix(.Rows - 1, 10) = Format(Val(loctable!TOTAL & ""), "Fixed")
        .TextMatrix(.Rows - 1, 11) = loctable!PRICE_C & ""
        
        
        
        If loctable!price <> loctable!PRICE_C Then
            .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbRed
            .Cell(flexcpBackColor, .Rows - 1, 11, .Rows - 1, 11) = vbYellow
        End If
        nTotal = nTotal + Val(loctable!TOTAL & "")
    ElseIf loctable!Flag = 1 Then
        .TextMatrix(.Rows - 1, 0) = Mid(loctable!doc_no, 7, 6)
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = loctable!DESCA
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HD3BD78
        
'        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
'        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalDiscount = nTotalDiscount + Val(.TextMatrix(.Rows - 1, 8 + 1))
        If loctable!Flag = 0 Then nTotalSalesDis = nTotalSalesDis + Val(.TextMatrix(.Rows - 1, 8 + 1))
    ElseIf loctable!Flag = 3 Then
        .TextMatrix(.Rows - 1, 0) = DelZero(Mid(loctable!doc_no, 7, 6))
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = "«·«Ã„«·Ì"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = loctable!TOTAL
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
    ElseIf loctable!Flag = 4 Then
        .TextMatrix(.Rows - 1, 0) = DelZero(Mid(loctable!doc_no, 7, 6))
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = "‰ﬁœÌ…"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = Format(loctable!TOTAL, "#0.00")
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalCash = nTotalCash + Val(loctable!TOTAL & "")
    ElseIf loctable!Flag = 5 Then
        '.TextMatrix(.Rows - 1, 0) = LOCTABLE!doc_no
        For i = 0 To 8
            .TextMatrix(.Rows - 1, i) = "›Ì“«"
        Next
        .MergeRow(.Rows - 1) = True
        .TextMatrix(.Rows - 1, 8 + 1) = Format(loctable!TOTAL, "#0.00")
        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = &HC0FFFF
        .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
        nTotalVisa = nTotalVisa + Val(loctable!TOTAL & "")
    End If
    .TextMatrix(.Rows - 1, 12) = loctable!username_disc & ""
    .TextMatrix(.Rows - 1, 13) = loctable!username_ret & ""
    .TextMatrix(.Rows - 1, 14) = loctable!SALES_RET & ""
    loctable.MoveNext
Loop

If nTotal <> 0 Then
    .AddItem ""
    For i = 0 To 8
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·ÌÊ„"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = nTotal - nTotalDiscount
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0E0FF
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If nTotalCash <> 0 Then
    .AddItem ""
    For i = 0 To 8
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·‰ﬁœÌ…"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = nTotalCash
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If nTotalVisa <> 0 Then
    .AddItem ""
    For i = 0 To 7 + 1
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·›Ì“«"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = nTotalVisa
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If

If Round((nTotal - nTotalDiscount) - (nTotalCash + nTotalVisa), 0) <> 0 Then
    .AddItem ""
    For i = 0 To 8
        .TextMatrix(.Rows - 1, i) = "≈Ã„«·Ì «·¬Ã·"
    Next
    .MergeRow(.Rows - 1) = True
    .TextMatrix(.Rows - 1, 8 + 1) = Round((nTotal - nTotalDiscount) - (nTotalCash + nTotalVisa), 2)
    .Cell(flexcpBackColor, .Rows - 1, 1, .Rows - 1, .Cols - 1) = &HC0FFC0
    .Cell(flexcpForeColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = vbBlue
End If
End With
Fixgrd4
End Sub
Private Sub Fixgrd4()
With GRID3
.Cols = 15
.TextMatrix(0, 0) = "—ﬁ„ „” ‰œ"
.TextMatrix(0, 1) = "«·Êﬁ "
.TextMatrix(0, 2) = "«·»«∆⁄"
.TextMatrix(0, 3) = "«·„ÊœÌ·"
.TextMatrix(0, 4) = "«·„’‰⁄"
.TextMatrix(0, 5) = "»«—ﬂÊœ"
.TextMatrix(0, 6) = "«·’‰›"
.TextMatrix(0, 7) = "„»Ì⁄« "
.TextMatrix(0, 8) = "„— Ã⁄« "
.TextMatrix(0, 9) = "«·”⁄—"
.TextMatrix(0, 10) = "«·≈Ã„«·Ï"
.TextMatrix(0, 11) = "”⁄— „” Â·ﬂ"

.TextMatrix(0, 12) = "’·«ÕÌ… Œ’„"
.TextMatrix(0, 13) = "’·«ÕÌ… „— Ã⁄"
.TextMatrix(0, 14) = "„” ‰œ »Ì⁄"


.MergeCells = flexMergeFree
.MergeCol(0) = True

.ColWidth(0) = 1000
.ColWidth(1) = 1000
.ColWidth(2) = 1200
.ColWidth(2) = 800
.ColWidth(3) = 1800
.ColWidth(4) = 1000
.ColWidth(5) = 1200
.ColWidth(6) = 3000
.ColWidth(7) = 800
.ColWidth(8) = 1100
.ColWidth(9) = 1100
.ColWidth(10) = 1100
.ColWidth(11) = 1100

.ColWidth(12) = 1300
.ColWidth(13) = 1300
.ColWidth(14) = 1800

.RowHeight(0) = 700
.WordWrap = True
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
End With
End Sub
Sub myloadHour(Optional cBox As String, Optional cStore As String)
Dim cStr1 As String
cStr1 = "SELECT { fn HOUR(TIME) } , COUNT(DOC_NO) , SUM(T_QUANT) , SUM(T_TOTAL) , SUM(discount) , SUM(T_TOTAL-discount), SUM(CASH) , SUM(VISA+VISA3) FROM  T_SALESDOC"
If lIsBranchStore Then cStr1 = cStr1 & turn(cStr1) & "BRANCH = " & MyParn(cBranch)
cStr1 = cStr1 & turn(cStr1) & "DATE >= " & DateSq(xDate1.text)
cStr1 = cStr1 & turn(cStr1) & "DATE <= " & DateSq(xdate2.text)
If cBox <> "" Then cStr1 = cStr1 & turn(cStr1) & "STORE = " & MyParn(cBox)
If Not bOpt5 Then cStr1 = cStr1 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "

cStr1 = cStr1 & " GROUP BY { fn HOUR(TIME) } ORDER BY { fn HOUR(TIME) }  "
Set DATAHOUR.Recordset = myRecordSet(cStr1, con)
fixGrdHour
End Sub
Private Sub fixGrdHour()
With GRIDHOUR
.FixedCols = 0
.FixedRows = 1
.Cols = 9
.TextMatrix(0, 0) = "«·”«⁄…"
.TextMatrix(0, 1) = "⁄œœ »Ê‰« "
.TextMatrix(0, 2) = "⁄œœ ﬁÿ⁄ „»Ì⁄« "
.TextMatrix(0, 3) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 4) = "«·Œ’„"
.TextMatrix(0, 5) = "«·’«›Ï"
.TextMatrix(0, 6) = "‰ﬁœÏ"
.TextMatrix(0, 7) = "›Ì“«"
.TextMatrix(0, 8) = "«·›—ﬁ"
.RowHeight(0) = 800
.WordWrap = True
.ColWidth(1) = 1000
.ColWidth(2) = 1000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1000
.ColWidth(6) = 1000
.ColWidth(7) = 1000
.ColWidth(8) = 1000

For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 8) = Val(.TextMatrix(nRow, 5)) - Val(.TextMatrix(nRow, 6)) - Val(.TextMatrix(nRow, 7))
Next nRow
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignCenterCenter
Next
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTAbove
For nCol = 1 To .Cols - 1
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
Next nCol
End With
End Sub
Private Sub doprint_day()
Dim aHeader(3)
'f Not MYVALID Then Exit Sub
Dim temptable As New ADODB.Recordset

contemp.Execute "DELETE * FROM TEMP"
temptable.Open "temp", contemp, adOpenStatic, adLockOptimistic, adCmdTable

With GRID3
For nRow = 1 To .Rows - 1
    If IsNumeric(.TextMatrix(nRow, 5)) Then
        temptable.AddNew
        aRet = aGetDesca("select mosm , scal from file1_10 where item = " & Val(.TextMatrix(nRow, 5)), con)
        temptable!str1 = .TextMatrix(nRow, 5)
        temptable!STR7 = .TextMatrix(nRow, 1)
        temptable!str11 = aRet(1)
        temptable!str12 = .TextMatrix(nRow, 4)
        temptable!str13 = .TextMatrix(nRow, 3)
        
        temptable!str2 = .TextMatrix(nRow, 3)
        
        temptable!str6 = .TextMatrix(nRow, 2)
        temptable!str3 = .TextMatrix(nRow, 6)
        temptable!str4 = .TextMatrix(nRow, 0)
        temptable!str5 = aRet(2)
        temptable!Date1 = xDate.Value
'       temptable!str7 = " „ «»⁄… ÌÊ„Ì… „»Ì⁄«  "
        temptable!VAL1 = Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8))
        temptable!val3 = Val(.TextMatrix(nRow, 9))
        temptable!VAL4 = Val(.TextMatrix(nRow, 10))
        temptable.Update
    End If
    If .TextMatrix(nRow, 5) = "«·Œ’„" Then
        temptable.AddNew
        temptable!Date1 = xDate.Value
        temptable!VAL4 = Val(.TextMatrix(nRow, 9)) * -1
        temptable!str4 = .TextMatrix(nRow - 1, 0)
        temptable!STR7 = "9999"
        temptable!str3 = .TextMatrix(nRow, 5)
        temptable.Update
    End If
Next nRow
End With
If temptable.EOF And temptable.BOF Then
    MsgBox "·«  ÊÃœ »Ì«‰«  »«· ﬁ—Ì—"
    Exit Sub
End If
contemp.BeginTrans
contemp.CommitTrans
Main.REPORT1.ReportFileName = App.Path & "\Reports\rep_day.rpt"
Main.REPORT1.DataFiles(0) = tempFile
Main.REPORT1.Action = 1
temptable.Close
Set temptable = Nothing
End Sub
Private Sub FIXGRID5()
With GRID5
.FixedCols = 0
.FixedRows = 1
.Cols = 13
.RowHeight(0) = 1000
.WordWrap = True
.TextMatrix(0, 1) = "«· «—ÌŒ"
.TextMatrix(0, 1) = "ﬂÊœ"
.TextMatrix(0, 2) = "«·›—⁄"
.TextMatrix(0, 3) = "⁄œœ »Ê‰« "
.TextMatrix(0, 4) = "⁄œœ ﬁÿ⁄ „»Ì⁄« "
.TextMatrix(0, 5) = "ﬁÌ„… „»Ì⁄« "
.TextMatrix(0, 6) = "«·Œ’„"
.TextMatrix(0, 7) = "«·’«›Ï"
.TextMatrix(0, 8) = "‰ﬁœÏ"
.TextMatrix(0, 9) = "›Ì“«"
.TextMatrix(0, 10) = "point"
.TextMatrix(0, 11) = "«·›—ﬁ"
.TextMatrix(0, 12) = "⁄„Ê·… ›Ì“«"

.ColWidth(1) = 1500
.ColWidth(2) = 2000
.ColWidth(3) = 1000
.ColWidth(4) = 1000
.ColWidth(5) = 1200
.ColWidth(6) = 1200
.ColWidth(7) = 1200
.ColWidth(8) = 1200
.ColWidth(9) = 1200
.ColWidth(10) = 1200
.ColWidth(11) = 1200
.ColWidth(12) = 1200

.ColHidden(1) = True
.SubtotalPosition = flexSTBelow
For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 11) = Val(.TextMatrix(nRow, 7)) - Val(.TextMatrix(nRow, 8)) - Val(.TextMatrix(nRow, 9)) - Val(.TextMatrix(nRow, 10))
    For nCol = 5 To .Cols - 1
        .TextMatrix(nRow, nCol) = Format(Val(.TextMatrix(nRow, nCol)), "#0.00")
    Next nCol
Next nRow
For i = 0 To .Cols - 1
    .ColAlignment(i) = flexAlignRightCenter
Next
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTBelow
For nCol = 3 To .Cols - 1
    .Subtotal flexSTSum, -1, nCol, "#0.00", vbYellow, , True, " "
    If IsDate(xDate1.text) Then
        If DateValue(xDate1.text) <> DateValue(xDate1.text) Then .Subtotal flexSTSum, 0, nCol, "#0.00", vbYellow, , True, " "
    End If
Next nCol
If IsDate(xDate1.text) Then
    If DateValue(xDate1.text) <> DateValue(xDate1.text) Then .Subtotal flexSTSum, 0, 3, "#0", vbYellow, , True, " "
    If DateValue(xDate1.text) <> DateValue(xDate1.text) Then .Subtotal flexSTSum, 0, 4, "#0", vbYellow, , True, " "
End If
.Subtotal flexSTSum, -1, 3, "#0", vbYellow, , True, " "
.Subtotal flexSTSum, -1, 4, "#0", vbYellow, , True, " "
End With
End Sub
Sub myload5(Optional cBox As String, Optional cStore As String)
Dim cStr1 As String

cStr1 = "SELECT date ,t_Salesdoc.STORE ,  STORENAME , COUNT(DOC_NO) , SUM(T_QUANT) , SUM(T_TOTAL) , SUM(discount) , SUM(T_TOTAL-discount), SUM(CASH) , SUM(VISA+VISA3) , SUM(point) ,0, SUM(visa2+visa3_2)  FROM  T_SALESDOC"
cStr1 = cStr1 & turn(cStr1) & "DATE >= " & DateSq(xDate1.text)
cStr1 = cStr1 & turn(cStr1) & "DATE <= " & DateSq(xdate2.text)
If cBox <> "" Then cString = cString & turn(cString) & "SALES_MODEL.BOX   = " & MyParn(cBox)
If cStore <> "" Then cString = cString & turn(cString) & "SALES_MODEL.STORE = " & MyParn(cStore)
If Not bOpt5 Then cStr1 = cStr1 & " AND t_Salesdoc.[store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "

If cBranch <> "00" Then
    If cBranchBox <> "" Then cStr1 = cStr1 & turn(cStr1) & "box = " & MyParn(cBranchBox)
    If lIsBranchStore Then cStr1 = cStr1 & turn(cStr1) & "branch = " & MyParn(cBranch)
End If
cStr1 = cStr1 & " GROUP BY date,STORENAME , STORE ORDER BY STORE "
Set data5.Recordset = myRecordSet(cStr1, con)
FIXGRID5
End Sub


Private Sub FIXGRID6()
With grid6
.FixedCols = 0
.FixedRows = 1
.Cols = 3
.RowHeight(0) = 1000
.WordWrap = True
.TextMatrix(0, 0) = "«·›Ì“«"
.TextMatrix(0, 1) = "«·›—⁄ "
.TextMatrix(0, 2) = "«·ﬁÌ„… "

.ColWidth(0) = 2000
.ColWidth(1) = 2000
.ColWidth(2) = 1500

.SubtotalPosition = flexSTBelow
For nRow = 1 To .Rows - 1
    .TextMatrix(nRow, 2) = Format(Round(Val(.TextMatrix(nRow, 2)), 2), "#0.00")
Next nRow
.ExplorerBar = flexExSortShow
.SubtotalPosition = flexSTBelow
.Subtotal flexSTSum, -1, 2, "#0.00", vbYellow, , True, " "
If cBranch = "00" Then
    .Subtotal flexSTSum, 0, 2, "#0.00", vbYellow, , True, " "
End If
.ColFormat(2) = "#,###,##0.00"
End With
End Sub
Sub myload6(Optional cBox As String, Optional cStore As String)
Dim cStr1 As String
cStr1 = "SELECT VISA.DESCA , FILE0_40.DESCA , SUM(FILE6_20H.VISA)   FROM  FILE6_20H LEFT JOIN VISA ON VISA.CODE = FILE6_20H.CODEVISA INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE WHERE VISA <> 0 "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xDate1.text)
cStr1 = cStr1 & " AND DATE <= " & DateSq(xdate2.text)
If Not bOpt5 Then cStr1 = cStr1 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
If cBranch <> "00" Then
    If cBranchBox <> "" Then cStr1 = cStr1 & " AND FILE6_20H.box = " & MyParn(cBranchBox)
    If lIsBranchStore Then cStr1 = cStr1 & " AND FILE6_20H.branch = " & MyParn(cBranch)
End If
If cStore <> "" And cBranch = "00" Then cStr1 = cStr1 & " AND FILE6_20H.STORE = " & MyParn(cStore)
cStr1 = cStr1 & " GROUP BY VISA.DESCA , FILE0_40.DESCA , FILE6_20H.STORE , VISA.CODE "
cStr1 = cStr1 & " UNION ALL "
cStr1 = cStr1 & "SELECT VISA.DESCA , FILE0_40.DESCA , SUM(FILE6_20H.VISA3)   FROM  FILE6_20H LEFT JOIN VISA ON VISA.CODE = FILE6_20H.CODEVISA3 INNER JOIN FILE0_40 ON FILE0_40.CODE = FILE6_20H.STORE WHERE VISA3 <> 0 "
cStr1 = cStr1 & " AND DATE >= " & DateSq(xDate1.text)
cStr1 = cStr1 & " AND DATE <= " & DateSq(xdate2.text)
If Not bOpt5 Then cStr1 = cStr1 & " AND [store] IN (SELECT STORE FROM USERSHOP WHERE CODE = " & nusercode & " ) "
If cBranch <> "00" Then
    If cBranchBox <> "" Then cStr1 = cStr1 & " AND FILE6_20H.box = " & MyParn(cBranchBox)
    If lIsBranchStore Then cStr1 = cStr1 & " AND FILE6_20H.branch = " & MyParn(cBranch)
End If
If cStore <> "" And cBranch = "00" Then cStr1 = cStr1 & " AND FILE6_20H.STORE = " & MyParn(cStore)

cStr1 = cStr1 & " GROUP BY VISA.DESCA , FILE0_40.DESCA , FILE6_20H.STORE , VISA.CODE "

Set DATA6.Recordset = myRecordSet(cStr1, con)
FIXGRID6
End Sub
Private Sub xDate1_LostFocus()
    If IsDate(xDate1.text) Then xDate.Value = xDate1.text

End Sub
