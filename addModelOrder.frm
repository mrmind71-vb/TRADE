VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{BF5DA8BB-099C-41DC-88F2-87E2D46819E4}#3.3#0"; "ImgX61.ocx"
Begin VB.Form addModelOrder 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14850
   BeginProperty Font 
      Name            =   "Arabic Transparent"
      Size            =   11.25
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "’Ê—… «·„ÊœÌ· "
      Height          =   3390
      Left            =   225
      TabIndex        =   19
      Top             =   5625
      Width           =   4290
      Begin ImgXCtrl6.ImgXCtrl ImgX1 
         Height          =   3030
         Left            =   45
         TabIndex        =   20
         Top             =   270
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   5345
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
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   4860
      TabIndex        =   12
      Top             =   1440
      Width           =   9915
      Begin VB.TextBox xbarcode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000012&
         Height          =   390
         Left            =   45
         MaxLength       =   40
         RightToLeft     =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "»«—ﬂÊœ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1305
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   180
         Width           =   645
      End
      Begin VB.Label xModel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5940
         TabIndex        =   15
         Top             =   180
         Width           =   2625
      End
      Begin VB.Label xDesca 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2070
         TabIndex        =   14
         Top             =   180
         Width           =   3795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "—ﬁ„ «·„ÊœÌ· :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8685
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   180
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ﬂ„Ì… «·’‰›"
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   225
      TabIndex        =   10
      Top             =   1260
      Width           =   1995
      Begin VB.Label xTotal 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   135
         TabIndex        =   11
         Top             =   315
         Width           =   1725
      End
   End
   Begin VB.Frame Frame9 
      Height          =   1995
      Left            =   2250
      TabIndex        =   7
      Top             =   90
      Width           =   2580
      Begin Threed.SSCommand cmdSave 
         Height          =   600
         Left            =   45
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   180
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   1058
         _Version        =   196610
         ForeColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "addModelOrder.frx":0000
         Caption         =   " ”ÃÌ· ﬂ„Ì«  «·„ÊœÌ· "
         Alignment       =   4
         ButtonStyle     =   1
         PictureAlignment=   1
         BevelWidth      =   10
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdAddItem 
         Height          =   600
         Left            =   45
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   765
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   1058
         _Version        =   196610
         ForeColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "addModelOrder.frx":2536
         Caption         =   " ⁄œÌ· »Ì«‰«  «·„ÊœÌ· "
         Alignment       =   4
         ButtonStyle     =   1
         PictureAlignment=   1
         BevelWidth      =   10
         ShapeSize       =   1
      End
      Begin Threed.SSCommand cmdExit 
         CausesValidation=   0   'False
         Height          =   600
         Left            =   45
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1350
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   1058
         _Version        =   196610
         ForeColor       =   0
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "addModelOrder.frx":5811
         Caption         =   "Œ—ÊÃ "
         Alignment       =   4
         ButtonStyle     =   1
         PictureAlignment=   1
         BevelWidth      =   10
         ShapeSize       =   1
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid grid1 
      Height          =   3495
      Left            =   225
      TabIndex        =   4
      Top             =   2115
      Width           =   14550
      _cx             =   25665
      _cy             =   6165
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
      FixedRows       =   3
      FixedCols       =   3
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
   Begin MSAdodcLib.Adodc DATA2 
      Height          =   330
      Left            =   135
      Top             =   405
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
      Left            =   -1170
      Top             =   90
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
      Height          =   1455
      Left            =   4860
      TabIndex        =   5
      Top             =   0
      Width           =   9915
      Begin VB.CheckBox xedit 
         Caption         =   " ⁄œÌ·"
         Enabled         =   0   'False
         Height          =   360
         Left            =   2250
         TabIndex        =   34
         Top             =   225
         Width           =   1140
      End
      Begin VB.CommandButton Cmd_Model 
         BackColor       =   &H00C0FFC0&
         Caption         =   "≈” ⁄·«„ „ÊœÌ·"
         CausesValidation=   0   'False
         Height          =   420
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   135
         Width           =   1905
      End
      Begin VB.TextBox xdiscount 
         Alignment       =   2  'Center
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   135
         MaxLength       =   10
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   990
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "„Ê”„ „” „—"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   4365
         TabIndex        =   28
         Top             =   225
         Width           =   1275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         Height          =   330
         Left            =   3780
         TabIndex        =   26
         Top             =   990
         Width           =   375
      End
      Begin VB.CheckBox xTrust_model 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000C&
         Caption         =   "»÷«⁄… √„«‰…"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   135
         TabIndex        =   25
         Top             =   1800
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox xModelFact 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   390
         Left            =   135
         MaxLength       =   10
         TabIndex        =   2
         Top             =   585
         Width           =   1905
      End
      Begin MSDataListLib.DataCombo xMosm 
         Height          =   345
         Left            =   5760
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   180
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   635
         _Version        =   393216
         Enabled         =   0   'False
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xFact 
         Height          =   345
         Left            =   5760
         TabIndex        =   1
         Top             =   585
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo xRemark 
         Height          =   345
         Left            =   4185
         TabIndex        =   27
         Top             =   990
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
         RightToLeft     =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arabic Transparent"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "‰”»… «·Œ’„"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1035
         Width           =   945
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„·«ÕŸ«  :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1035
         Width           =   840
      End
      Begin VB.Label xCodeSub 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   5175
         TabIndex        =   21
         Top             =   585
         Width           =   555
      End
      Begin VB.Label xSupp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   3780
         TabIndex        =   18
         Top             =   180
         Width           =   555
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "„ÊœÌ· «·„’‰⁄ :"
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2115
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   675
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„Ê”„ :"
         Height          =   270
         Left            =   8685
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "«·„’‰⁄ :"
         Height          =   270
         Left            =   8685
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   675
         Width           =   690
      End
   End
   Begin MSAdodcLib.Adodc DATA10 
      Height          =   330
      Left            =   1170
      Top             =   630
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
      Left            =   -2925
      Top             =   -450
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
   Begin MSAdodcLib.Adodc DATA20 
      Height          =   330
      Left            =   135
      Top             =   225
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
   Begin VSFlex7Ctl.VSFlexGrid Grid2 
      Height          =   315
      Left            =   4500
      TabIndex        =   22
      Top             =   5715
      Visible         =   0   'False
      Width           =   3750
      _cx             =   6615
      _cy             =   556
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
      Rows            =   1
      Cols            =   1
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
   Begin VSFlex7Ctl.VSFlexGrid GRID20 
      Height          =   3345
      Left            =   4590
      TabIndex        =   24
      Top             =   5805
      Visible         =   0   'False
      Width           =   10185
      _cx             =   17965
      _cy             =   5900
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
      SelectionMode   =   0
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
      ScrollBars      =   2
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
   Begin MSAdodcLib.Adodc DATA3 
      Height          =   330
      Left            =   135
      Top             =   855
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
Attribute VB_Name = "addModelOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public cDoc_No As String, myForm As Form, nColItem As Integer, nColQuant As Long, sStore As String
Public sModel As String
Public sSection As String
Public bshowBal As Boolean, bshowcost As Boolean, bAddItem As Boolean
Public SMosm As String, sSupp As String, SFact As String, SModelFact As String, Sdiscount As String
Public bCost1 As Boolean
Dim con As New ADODB.Connection
Public strModel As String
Dim oSearch As New Search3, oSearchMF As New Search3, oSearchFact As New Search3, oSearchRemark As New Search3
Dim bUpDate As Boolean, oSearchItem As New Search3
Const LoadMode = 0, DefineMode = 1
Private Sub Check1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 1 Then
    xMosm.BoundText = "000"
Else
    xMosm.BoundText = SMosm
End If
End Sub
Private Sub cmdAddItem_Click()
If Not myValidAdd Then Exit Sub
ITEMS.bAddModel = True
ITEMS.bedit = True
ITEMS.sModel = XMODEL.Caption
ITEMS.sSupp = xSupp.Caption
ITEMS.SFact = xFact.BoundText
ITEMS.SMosm = xMosm.BoundText
ITEMS.SModelFact = xModelFact.text
ITEMS.sCode = Ordere_Purch.xCode.text
ITEMS.Show 1
If Not myload Then
    On Error Resume Next
    xModelFact.SetFocus
    Err.Clear
Else
    On Error Resume Next
    grid1.SetFocus
    grid1.Select 3, 3
    Err.Clear
End If
End Sub
Private Sub cmdExit_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
If grid1.Cols <= 2 Then
    MsgBox "·«  ÊÃœ „ÊœÌ·«  ·«÷«› Â«"
    Exit Sub
End If

myForm.myProcAddModel
xModelFact.text = ""
mydefine
On Error Resume Next
xModelFact.SetFocus
myGotFocus xModelFact
Err.Clear
End Sub
Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 13 Then KeyCode = 0
End Sub
Private Sub Command1_Click()
ReDim aPublic(5)
Dim sCode As String
sCode = xRemark.text
aPublic(0) = "REMARK"
aPublic(1) = "Code"
aPublic(2) = "Desca"
aPublic(3) = "ﬂÊœ"
aPublic(4) = "«·»Ì«‰"
aPublic(5) = "«·„Õ·ÊŸ« "
FlagFrm.bedit = True
FlagFrm.aPublic = aPublic
FlagFrm.Show 1
data3.Refresh
xRemark.BoundText = sCode
End Sub

Private Sub Form_Activate()
On Error Resume Next
If sModel <> "" Then
    grid1.SetFocus
    sModel = ""
End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 27 Then Unload Me
End Sub
Private Sub Form_Load()
Me.Top = 0: Me.Left = 0
openCon con
Set GRID20.DataSource = data20
data20.ConnectionString = strCon

Set data1.Recordset = myRecordSet("SELECT CODE, DESCA FROM FACT ORDER BY DESCA ", con)
Set xFact.RowSource = data1
xFact.ListField = "DESCA"
xFact.BoundColumn = "CODE"

Set DATA2.Recordset = myRecordSet("SELECT * FROM MOSM ORDER BY DATE DESC ", con)
Set xMosm.RowSource = DATA2
xMosm.ListField = "MOSM"
xMosm.BoundColumn = "MOSM"
If xMosm.BoundText = "" Then xMosm.BoundText = Ordere_Purch.xMosm.BoundText

Set data3.Recordset = myRecordSet("SELECT * FROM REMARK ORDER BY DESCA", con)
Set xRemark.RowSource = data3
xRemark.ListField = "DESCA"
xRemark.BoundColumn = "DESCA"
xRemark.BoundText = Ordere_Purch.xMosm.BoundText

myDefineGrd
Fixgrd

'Fixgrd2

GRID20.Visible = (Ordere_Purch.myPublic = 0 Or Ordere_Purch.myPublic = 1)

XMODEL.Caption = sModel
xMosm.BoundText = SMosm
xSupp.Caption = sSupp
xFact.BoundText = SFact
xCodeSub.Caption = SFact
xModelFact.text = SModelFact
xDiscount.text = Sdiscount
Handlecontrols
myload
End Sub
Private Sub VsModel_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim nTot As Double
    With VsModel
        For r = 4 To .Rows - 1
            For c = 2 To .Cols - 1
                nTot = nTot + Val(.TextMatrix(r, c))
            Next c
        Next r
    End With
End Sub
Private Sub VsModel_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
KeyAscii = RetNumber(KeyAscii, False)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 19 Then
    cmdSave_Click
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
closeCon con
Set addModelfrm = Nothing
End Sub
Private Function myLoadModelgrd() As Boolean
Dim aRet As Variant, cFieldas As String, cField As String
myDefineGrd
aRet = retFields
If aRet(0) = "" Then Exit Function

cField = aRet(0)
cFieldas = aRet(1)

FillItem cFieldas, cField
FixCost cFieldas, cField
Fixgrd
CalcTotals
myLoadModelgrd = True
End Function
Private Sub Fixgrd()
With grid1
.ColHidden(0) = True
'.ColWidth(0) = 500
.ColWidth(1) = 1300
'.RowHidden(1) = (Not bopt1) And (Not bshowcost)
nColWidth = (.Width - 200 - .ColWidth(0) - .ColWidth(1)) / .Cols
If nColWidth < 500 Then nColWidth = 500
If nColWidth > 1200 Then nColWidth = 1200
For nCol = 2 To .Cols - 1
    .ColWidth(nCol) = nColWidth
    .ColAlignment(nCol) = flexAlignCenterCenter
Next
End With
End Sub
Private Sub myDefineGrd()
GRID20.Rows = 1
grid2.Rows = 3
grid2.Cols = 2

grid1.Rows = 3
grid1.Cols = 2

grid1.MergeCells = flexMergeRestrictRows
grid1.TextMatrix(0, 1) = "«·„ﬁ«”"
grid1.TextMatrix(1, 1) = "”⁄— „’‰⁄"
grid1.TextMatrix(2, 1) = "”⁄— „” Â·ﬂ"
xTotal.Caption = ""
'grid1.RowHidden(1) = Not bCost1
'grid1.RowHidden(2) = Not bCost1
End Sub
Private Sub FixCost(cFieldas, cField)
' „·∆ «·ÃœÊ·

If Ordere_Purch.myPublic = 0 Or Ordere_Purch.myPublic = 1 Then
    cString = "Select " & cFieldas & _
              " From " & _
              " (Select scal,cost from file1_10 WHERE MODEL = " & MyParn(XMODEL.Caption) & " ) AS TABLE1" & _
              " PIVOT " & _
              " (max(cost)" & _
              " FOR SCAL IN " & _
              "(" & cField & ")" & _
              ") as pvt  "
Else
    cString = "Select " & cFieldas & _
              " From " & _
              " (Select scal,cost2 as cost from file1_10 WHERE MODEL = " & MyParn(XMODEL.Caption) & " ) AS TABLE1" & _
              " PIVOT " & _
              " (max(cost)" & _
              " FOR SCAL IN " & _
              "(" & cField & ")" & _
              ") as pvt  "
End If
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
          " (Select scal,price from file1_10 WHERE MODEL = " & MyParn(XMODEL.Caption) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(price)" & _
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
Dim aRet(1) As String
Dim FieldTable As New ADODB.Recordset
'  ⁄—Ì› «·«⁄„œ…
Dim cString As String, cWhere As String
FieldTable.Open "Select SCAL from file1_10 where model = " & MyParn(XMODEL.Caption) & " GROUP BY SCAL,C_SCAL order by c_scal", con, adOpenStatic, adLockReadOnly
Do Until FieldTable.EOF
    If Not IsNull(FieldTable!scal) Then
        cFieldas = cFieldas & turn(cField, ",") & "[" & FieldTable!scal & "]" & " as " & "[" & FieldTable!scal & "]"
        cField = cField & turn(cField, ",") & "[" & FieldTable!scal & "]"
    End If
    FieldTable.MoveNext
Loop

aRet(0) = cField
aRet(1) = cFieldas
retFields = aRet
' ⁄œ„ ÊÃÊœ «⁄„œ…
FieldTable.Close
Set FieldTable = Nothing
End Function
Private Sub FillItem(cFieldas, cField)
Dim GrdTable As New ADODB.Recordset, cString As String
' „·∆ «·ÃœÊ·
cString = "Select c_color as [—ﬁ„ «··Ê‰] ,color as [«··Ê‰] " & turn(cFieldas, ",") & cFieldas & _
          " From " & _
          " (Select c_color,Color,scal,item from file1_10 WHERE MODEL = " & MyParn(XMODEL.Caption) & " ) AS TABLE1" & _
          " PIVOT " & _
          " (max(item)" & _
          " FOR SCAL IN " & _
          "(" & cField & ")" & _
          ") as pvt  " & _
          " order by pvt.C_color"

GrdTable.Open cString, con, adOpenStatic, adLockReadOnly, adCmdText
grid1.Cols = GrdTable.Fields.Count: grid2.Cols = GrdTable.Fields.Count

For nCol = 2 To GrdTable.Fields.Count - 1
    grid1.TextMatrix(0, nCol) = GrdTable.Fields(nCol).Name
Next
xedit.Value = 0
Do Until GrdTable.EOF
    grid2.AddItem ""
    grid1.AddItem ""
    For nCol = 0 To GrdTable.Fields.Count - 1
        If nCol <= 1 Then
            grid1.TextMatrix(grid2.Rows - 1, nCol) = GrdTable.Fields(nCol).Value & ""
        Else
            grid2.TextMatrix(grid2.Rows - 1, nCol) = GrdTable.Fields(nCol).Value & ""
            nFoundRow = myForm.grid1.FindRow(GrdTable.Fields(nCol).Value & "", , nColItem)
            If nFoundRow <> -1 Then
                xedit.Value = 1
                 grid1.TextMatrix(grid1.Rows - 1, nCol) = Myvalue(myForm.grid1.TextMatrix(nFoundRow, nColQuant))
                 If Val(myForm.grid1.TextMatrix(nFoundRow, myForm.grid1.Cols - 1)) Then xTrust_model.Value = 1
             Else
                 grid1.TextMatrix(grid1.Rows - 1, nCol) = ""
             End If
        End If
    Next
    GrdTable.MoveNext
Loop
GrdTable.Close
Set GrdTable = Nothing
End Sub
Private Sub mydefine()
xDesca.Caption = ""
XMODEL.Caption = ""
xDesca.Caption = ""
xTrust_model.Value = 0
xDiscount.text = ""
myLoadPicture ""
myDefineGrd
Handlecontrols
End Sub
Private Sub grid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If Col >= 2 And Row = 3 Then
    If emptyRow(grid1, Row, Col + 1) Then
        If xedit.Value = 0 Then
            For i = Col To grid1.Cols - 1
                grid1.TextMatrix(Row, i) = grid1.TextMatrix(Row, Col)
            Next
        End If
    End If
End If
lastsub:
CalcTotals
End Sub
Private Sub grid1_EnterCell()
If bshowBal And grid1.Cols > 2 And grid1.Rows > 3 And strStore <> "" Then
    'xBalance_Item.Caption = Balance_sales(Grid2.TextMatrix(grid1.Row, grid1.Col), strStore, con)
End If
End Sub
Private Sub Grid1_GotFocus()
With grid1
    If .Cols > 2 And .Rows > 3 Then
        .Select 3, 2
    End If
End With
End Sub
Private Sub Grid1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, grid1.Row, grid1.Col
End If
End Sub
Private Sub CellPos(ByRef KeyCode, ByVal Row As Long, ByVal Col As Long)
KeyCode = 0
If Col < grid1.Cols - 1 Then
    Col = Col + 1
    If grid1.TextMatrix(Row, Col) = "" And Val(grid1.TextMatrix(Row, Col - 1)) > 0 Then
         grid1.TextMatrix(Row, Col) = Val(grid1.TextMatrix(Row, Col - 1))
    End If
    grid1.Col = Col
Else
    cmdSave_Click
End If
End Sub
Private Sub grid1_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
If KeyCode = 13 Then
    CellPos KeyCode, Row, Col
End If
End Sub

Private Sub Grid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
If (Not IsNumeric(grid1.EditText)) Then
    Cancel = True
    Exit Sub
End If
If bshowBal Then
'    If Grid2.TextMatrix(Row, Col) <> "" Then
'        Dim nQuant As Double
'       nFoundRow = myForm.grid1.FindRow(Grid2.TextMatrix(Row, Col) & "", , nColItem)
'       If nFoundRow <> -1 Then
'            nQuant = Val(myForm.grid1.TextMatrix(nFoundRow, nColQuant))
'       End If
'        If Val(grid1.EditText) > Val(xBalance_item.Caption) + Val(nQuant) Then
'            MsgBox "«·—’Ìœ " & nBalance & " Ê·« Ì”„Õ "
'            Cancel = True
'        End If
'    End If
End If
End Sub
Private Sub xModel_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 112 Then ModelLookupAll Me, oSearch
End Sub
Private Sub xModel_GotFocus()
myGotFocus XMODEL
End Sub
Private Sub XModel_LostFocus()
myLostFocus XMODEL
End Sub
Sub myProc()
On Error GoTo myerror
If ActiveControl.Name = Cmd_Model.Name Then
    XMODEL.Caption = oSearchItem.grid1.TextMatrix(oSearchItem.grid1.Row, 0)
    Unload oSearchItem
    myloadModel
End If
If ActiveControl.Name = xModelFact.Name Then
    ActiveControl.text = oSearchMF.grid1.TextMatrix(oSearchMF.grid1.Row, 0)
    Unload oSearchMF
    xModelFACT_GotFocus
ElseIf ActiveControl.Name = xFact.Name Then
    xFact.BoundText = oSearchFact.grid1.TextMatrix(oSearchFact.grid1.Row, 0)
    Unload oSearchFact
Else
    xRemark.text = oSearchFact.grid1.TextMatrix(oSearchFact.grid1.Row, 1)
    Unload oSearchRemark
End If
Exit Sub
myerror:
End Sub
Private Function myloadModel() As Boolean
myDefineGrd
grid2.Rows = 1
If Trim(XMODEL.Caption) = "" Then Exit Function
Dim loctable As New ADODB.Recordset, cString As String
Set loctable = ModelFind(Mid(XMODEL.Caption, 1, 16), con)
If Not (loctable.EOF And loctable.BOF) Then
    XMODEL.Caption = loctable!MODEL
    xDesca.Caption = loctable!DESCA & ""
    xFact.BoundText = loctable!Fact
    xCodeSub.Caption = xFact.BoundText
    xMosm.BoundText = loctable!MOSM
    xModelFact.text = loctable!modelfact
    xRemark.text = GetDesca("SELECT REMARK FROM " & Ordere_Purch.cFile & " WHERE DOC_NO = " & MyParn(Ordere_Purch.xdoc_no.text) & " AND MODEL = " & MyParn(XMODEL.Caption), con)
    myLoadModelgrd
'    If Ordere_Purch.myPublic = 0 Or Ordere_Purch.myPublic = 1 Then
'        myloadgrd2
'    End If
    myloadModel = True
End If
End Function
Private Sub CalcTotals()
Dim nTotal As Double
For nRow = 3 To grid1.Rows - 1
    For nCol = 2 To grid1.Cols - 1
        nTotal = nTotal + Val(grid1.TextMatrix(nRow, nCol))
    Next
Next
xTotal.Caption = Myvalue(nTotal)
End Sub

Private Sub xbarcode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And xbarcode.text <> "" Then
    aRet = aGetDesca("select model , modelfact , fact , mosm from file1_10 where  item = " & Val(xbarcode.text), con)
    If UBound(aRet) > 0 Then
        xModelFact.text = aRet(2)
        xFact.BoundText = aRet(3)
        XMODEL.Caption = aRet(1)
        myloadModel
    End If
End If
End Sub

Private Sub xFact_Change()
If xFact.MatchedWithList Then xCodeSub.Caption = xFact.BoundText
End Sub
Private Sub xFact_GotFocus()
    myGotFocus xFact
End Sub
Private Sub xFACT_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    KeyCode = 0
    SendKeys "{tAB}"
ElseIf KeyCode = 112 Then
    FactLookupAll Me, oSearchFact
End If
End Sub

Private Sub xFACT_LostFocus()
myLostFocus xFact
End Sub
Private Sub xFact_Validate(Cancel As Boolean)
If (Not xFact.MatchedWithList) And Trim(xFact.BoundText) = "" Then
    Cancel = True
    Exit Sub
End If
If Not xFact.MatchedWithList Then
    cFact = xFact.BoundText
    xFact.BoundText = cFact
    If xFact.MatchedWithList Then Exit Sub
    cCode = IncRec(GetDesca("SELECT MAX(CODE) FROM FACT WHERE CODE > '300' ", con))
    If cCode = "" Then cCode = "301"
    addFactfrm.sdesca = cFact
    addFactfrm.sCode = cCode
    addFactfrm.Show 1
    data1.Refresh
    xFact.text = cFact
    If Not xFact.MatchedWithList Then xFact.BoundText = ""
End If
xCodeSub.Caption = xFact.BoundText
End Sub
Private Sub xModelFACT_GotFocus()
    myGotFocus xModelFact
End Sub
Private Sub xModelFact_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Trim(DelZero(xModelFact.text)) <> "" Then
    KeyCode = 0
    If xModelFact.text <> "" Then
        xModelFact.text = RetZero(xModelFact.text, 10)
        If Not myload Then cmdAddItem_Click
        If grid1.Cols > 2 Then grid1.SetFocus
    End If
ElseIf KeyCode = 112 And xMosm.MatchedWithList And xFact.MatchedWithList Then
    ModelFactLookup
End If
End Sub
Private Sub XModelFACT_LostFocus()
    xModelFact.text = RetZero(xModelFact.text, 10)
    myLostFocus xModelFact
End Sub
Private Sub xModelFact_Validate(Cancel As Boolean)
    If Trim(xModelFact.text) <> "" Then xModelFact.text = RetZero(xModelFact.text, 10)
    myload
End Sub
Private Function myload() As Boolean
Dim sSubMosm As String
If xMosm.MatchedWithList And xFact.MatchedWithList And Trim(xModelFact.text) <> "" Then
    If XMODEL.Caption = "" Then XMODEL.Caption = xFact.BoundText & xMosm.text & RetZero(xModelFact.text, 10) & retCode(Ordere_Purch.xCode.text, con)
    If GetDesca("select model from file1_10 where model = " & MyParn(XMODEL.Caption), con) = "" Then
        sSubMosm = GetDesca("select mosm from file1_10 where fact = " & MyParn(xFact.BoundText) & " and modelfact = " & MyParn(xModelFact.text), con)
        If sSubMosm <> "" Then
            If sSubMosm <> "" And sSubMosm <> xMosm.BoundText Then
                If MsgBox(" «·„ÊœÌ· „”Ã· ⁄·Ï „Ê”„ „Õ ·› -  ”ÃÌ· «·„ÊœÌ· " & xMosm.BoundText, vbYesNo) = vbYes Then
                    xMosm.BoundText = sSubMosm
                    XMODEL.Caption = xFact.BoundText & xMosm.text & RetZero(xModelFact.text, 10)
                End If
            End If
        End If
    End If
    myLoadPicture xFact.BoundText & xMosm.text & RetZero(xModelFact.text, 10)
Else
    XMODEL.Caption = ""
    myLoadPicture ""
End If
myload = myloadModel
End Function
Private Sub Handlecontrols()
'xModelFact.Enabled = sModelFact = ""
xFact.Enabled = SFact = ""
End Sub
Private Sub xMosm_Validate(Cancel As Boolean)
myload
End Sub
Private Sub ModelFactLookup()
Dim Generalarray(5)
Dim listarray(0, 4)
Dim GrdArray(4, 1)

Set Generalarray(0) = Me

Generalarray(1) = "SELECT MODELFACT0,DESCA, SUPP , MIN(COST) AS COST , MIN(PRICE ) AS PRICE FROM FILE1_10 WHERE ISSTOP = 0 "
If xMosm.MatchedWithList Then Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & "MOSM = " & MyParn(xMosm.BoundText)
If xFact.MatchedWithList Then Generalarray(1) = Generalarray(1) & turn(Generalarray(1)) & "Fact = " & MyParn(xFact.BoundText)
Generalarray(2) = "GROUP BY MODELFACT0, DESCA , SUPP Order by MODELFACT0 "
Generalarray(3) = 6000
Generalarray(5) = True

listarray(0, 0) = "«·—ﬁ„-≈”„ "
listarray(0, 1) = "%%desca+modelfact0%%"

GrdArray(0, 0) = "—ﬁ„ „ÊœÌ·"
GrdArray(0, 1) = 1500

GrdArray(1, 0) = "«·’‰›"
GrdArray(1, 1) = 4000

GrdArray(2, 0) = "«·„Ê—œ"
GrdArray(2, 1) = 1000

GrdArray(3, 0) = "”⁄— „’‰⁄"
GrdArray(3, 1) = 1000

GrdArray(4, 0) = "”⁄— „” Â·ﬂ"
GrdArray(4, 1) = 1000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchMF.Caption = "«” ⁄·«„"
oSearchMF.Show 1
End Sub
Private Function myValidAdd() As Boolean
If Trim(xModelFact.text) = "" Then
    MsgBox "„ÊœÌ· «·„’‰⁄ „ÿ·Ê»"
    Exit Function
End If
If Not xFact.MatchedWithList Then
    MsgBox "«·„’‰⁄ €Ì— „”Ã·"
    Exit Function
End If
myValidAdd = True
End Function
Private Sub LoadPhoto(sModelNo As String)
On Error GoTo myerror
imgx1.Images.Clear
If Trim(sModelNo) <> "" Then imgx1.Import.FromFile retFile(sModelNo)
Exit Sub
myerror:
Err.Clear
End Sub
Sub myloadgrd2()
    Dim cModelNo As String, cString As String
    cModelNo = Mid(XMODEL.Caption, 1, 16)
    cField2 = " ,(SELECT    SUM(OUT) FROM FILE1_10 AS FILE1_10_1 INNER JOIN FILE1_11 AS FILE1_11_1 ON FILE1_10_1.ITEM = FILE1_11_1.ITEM WHERE FILE1_11_1.[TYPE] = '6' AND FILE1_10_1.MODELNO = " & MyParn(cModelNo) & " ) "
    cField3 = " ,(SELECT    SUM([IN]-[OUT]) FROM FILE1_10 AS FILE1_10_1 INNER JOIN FILE1_11 AS FILE1_11_1 ON FILE1_10_1.ITEM = FILE1_11_1.ITEM WHERE FILE1_10_1.MODELNO = " & MyParn(cModelNo) & " ) "
     '                   0                   1               2               3                   4              5                           6                           7
    cString = " SELECT  FILE1_10.MODEL , FILE4_10.CODE, FILE4_10.DESCA, FILE1_11.DATE, FILE1_11.DOC_ID   , SUM( FILE1_11.[IN] ) , SUM( FILE1_11.OUT ) , SUM( FILE1_11.[IN] -  FILE1_11.OUT ) " & _
            cField2 & cField3 & _
            " FROM    FILE1_10 INNER JOIN FILE1_11 ON FILE1_10.ITEM = FILE1_11.ITEM  INNER JOIN FILE4_10 ON FILE1_10.code = FILE4_10.CODE  WHERE FILE1_10.MODELNO = " & MyParn(cModelNo) & _
            " AND (FILE1_11.TYPE = '2' OR FILE1_11.TYPE = '7')"
    cString = cString & " GROUP BY  FILE1_10.MODEL , FILE4_10.CODE, FILE4_10.DESCA, FILE1_11.DATE, FILE1_11.DOC_ID  ORDER BY FILE1_11.DATE , FILE4_10.CODE "
    Set data20.Recordset = myRecordSet(cString, con)
    Fixgrd2
End Sub
Sub Fixgrd2()
    With GRID20
    .ExplorerBar = flexExSortShow
    .FixedRows = 1
    .Cols = 11
    
    .TextMatrix(0, 1) = "ﬂÊœ"
    .TextMatrix(0, 2) = "«·„Ê—œ"
    .TextMatrix(0, 3) = " «—ÌŒ"
    .TextMatrix(0, 4) = "—ﬁ„ „” ‰œ"
    .TextMatrix(0, 5) = "„‘ —Ì« "
    .TextMatrix(0, 6) = "„— Ã⁄« "
    .TextMatrix(0, 7) = "’«›Ï"
    
    .TextMatrix(0, 8) = "„»Ì⁄« "
    .TextMatrix(0, 9) = "—’Ìœ"
    .TextMatrix(0, 10) = "‰”»… »Ì⁄"
    
    .ColWidth(0) = 0
    .ColWidth(1) = 800
    .ColWidth(2) = 2500
    .ColWidth(3) = 1400
    .ColWidth(4) = 1000
    .ColWidth(5) = 700
    .ColWidth(6) = 700
    .ColWidth(7) = 700
    .ColWidth(8) = 700
    .ColWidth(9) = 700
    .ColWidth(10) = 700
    .MergeCells = flexMergeFree
    
    .MergeCol(8) = True
    .MergeCol(9) = True
    .MergeCol(10) = True
    
    For nRow = 1 To .Rows - 1
        If Val(.TextMatrix(nRow, 6)) > 0 Then .Cell(flexcpForeColor, nRow, 0, nRow, .Cols - 1) = vbRed
    Next nRow
    .SubtotalPosition = flexSTBelow
    .Subtotal flexSTSum, -1, 5, "#0", vbYellow, vbRed, True, ""
    .Subtotal flexSTSum, -1, 6, "#0", vbYellow, vbRed, True, ""
    .Subtotal flexSTSum, -1, 7, "#0", vbYellow, vbRed, True, ""
    
    .Subtotal flexSTMax, -1, 8, "#0", vbYellow, vbRed, True, ""
    .Subtotal flexSTMax, -1, 9, "#0", vbYellow, vbRed, True, ""
    If .Rows > 1 And Val(.TextMatrix(.Rows - 1, 7)) <> 0 Then nRate = Round(Val(.TextMatrix(.Rows - 1, 8)) / Val(.TextMatrix(.Rows - 1, 7)) * 100, 2)
    
    For nRow = 1 To .Rows - 1
        .TextMatrix(nRow, 10) = nRate
    Next nRow
    
    End With
End Sub
Private Sub XREMARK_GotFocus()
myGotFocus xRemark
End Sub
Private Sub xRemark_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then RemarkLookup
End Sub

Private Sub XREMARK_LostFocus()
myLostFocus xRemark
End Sub
Private Sub RemarkLookup()
Dim Generalarray(5)
Dim listarray(0, 5)
Dim GrdArray(1, 1)

Set Generalarray(0) = Me

Generalarray(1) = "Select CODE , DESCA FROM REMARK "
Generalarray(2) = "order by DESCA"
Generalarray(3) = 8000
Generalarray(5) = False

listarray(0, 0) = "«·„Õ·ÊŸ… "
listarray(0, 1) = "(%%REMARK.DESCA%%)"


GrdArray(0, 0) = "ﬂÊœ"
GrdArray(0, 1) = 0

GrdArray(1, 0) = "«·„Õ·ÊŸ…"
GrdArray(1, 1) = 10000

searchArray = Array(Generalarray, listarray, GrdArray)
oSearchRemark.Caption = "«·„Õ·ÊŸ« "
oSearchRemark.Show 1
End Sub
Private Sub Cmd_Model_Click()
    ModelLookupAll Me, oSearchItem
End Sub
Private Sub xModel_Change()
            myLoadPicture Mid(XMODEL.Caption, 1, 16)
End Sub
Private Sub myLoadPicture(pModel As String)
On Error Resume Next
    If Not lServerPict Then Exit Sub
    imgx1.Images.Clear
    If pModel = "" Then Exit Sub
    If conPict Is Nothing Then Exit Sub
    Dim loctable As Recordset
    Dim bytes() As Byte, nSize As Long
    imgx1.Images.Clear
    Set loctable = myRecordSet("select PICT,size from PICT where MODELNO = " & MyParn(pModel), conPict)
    If loctable.EOF Then Exit Sub
    nSize = CLng(loctable("size").Value)
    bytes = loctable("PICT").GetChunk(nSize)
    imgx1.Import.FromMemoryFile bytes, ixmfJPG
    Err.Clear
End Sub
Sub AddModelNorDEm(pDoc, pModel, pstore)
        
End Sub
